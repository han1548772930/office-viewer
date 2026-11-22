import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, cellKey, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { applySheetBorders } from '../src/features/sheet-workbench/sheet-toolbar-format.ts'
import { createSheetBorderPlan, normalizeSheetBorderColor, normalizeSheetBorderPreset, normalizeSheetBorderStyle,
  sheetBorderLabels, sheetBorderStyleLabels, sheetMergedBorderSegments, SheetBorderStyleTypes, SheetBorderTypes } from '../src/features/sheet-workbench/sheet-cell-border.ts'
import { applySheetBorderFormatCommand, applySheetViewBorderFormatCommand, borderFormatActionForCommand, createSheetBorderFormatState,
  SHEET_BORDER_FORMAT_COMMANDS as commands, SHEET_BORDER_FORMAT_COMMAND_IDS, sheetBorderFormatCommandInput, sheetBorderFormatInputRanges } from '../src/features/sheet-workbench/sheet-border-format-command.ts'
import { createSheetContextMenuItems } from '../src/components/spreadsheet/sheet-context-menu.ts'
import { sheetToolbarBorderApiStyleLocales } from '../src/i18n/sheets/locales/toolbar-colors.ts'
import { supportedOfficeLocales } from '../src/i18n/regional-settings.ts'

const range = (startRow, endRow = startRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const pen = { style: 'thin', color: '#aabbcc' }
const action = (type, style = pen.style, color = pen.color) => ({ state: { type, style, color, active: true }, apply: true })
const sourceFor = sheet => ({ readBorder: (row, col) => sheet.cells[row + ':' + col]?.style.border,
  existingCells: () => Object.values(sheet.cells).map(({ row, col }) => ({ row, col })), mergedRanges: sheet.mergedRanges })
const border = (sheet, row, col = 0) => sheet.cells[row + ':' + col]?.style.border ?? {}
const edges = (sheet, merge, visible) => sheetMergedBorderSegments(merge, sourceFor(sheet).readBorder, visible)
const keys = sheet => Object.fromEntries(Object.entries(sheet.cells).map(([key, cell]) => [key, Object.keys(cell.style.border ?? {}).sort()]).filter(([, sides]) => sides.length))
const beforeWithoutBorders = sheet => Object.fromEntries(Object.entries(sheet.cells).map(([key, cell]) => {
  const copy = structuredClone(cell); delete copy.style.border; return [key, copy]
}))
const lock = target => ({ id: 'locked-border', permissionId: 'locked-border-permission', ranges: [target], ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '' })
function sample() {
  const book = createBlankWorkbook(), sheet = book.sheets[0]
  book.id = 'border-contract'; book.calculationMode = 'manual'
  sheet.id = 'border-contract-sheet'; book.activeSheetId = sheet.id
  sheet.rowCount = 100; sheet.columnCount = 26
  sheet.cells['0:0'] = { ...createCell(0, 0, '00123'), dataType: 'text', metadata: { source: 'sample' }, richText: [{ text: '00123' }], hyperlink: 'https://example.com' }
  sheet.cells['1:0'] = { ...createCell(1, 0, .25), formula: '=1/4', display: '0.25', dataType: 'formula', cachedDataType: 'number' }
  sheet.cells['2:0'] = createCell(2, 0, 7)
  sheet.cells['0:1'] = createCell(0, 1, 'neighbor')
  sheet.cells['4:4'] = createCell(4, 4, 'outside')
  return book
}

test('fixed position aliases and fourteen numeric pens retain fifteen presets and twelve menu styles', () => {
  assert.equal(Object.keys(sheetBorderLabels).length, 15)
  assert.equal(Object.keys(sheetBorderStyleLabels).length, 12)
  assert.deepEqual(Object.values(SheetBorderStyleTypes), Array.from({ length: 14 }, (_, index) => index))
  assert.deepEqual(Object.values(SheetBorderStyleTypes).map(normalizeSheetBorderStyle), [
    'none', 'thin', 'hair', 'dotted', 'dashed', 'dash-dot', 'dash-dot-dot', 'double', 'medium',
    'medium-dashed', 'medium-dash-dot', 'medium-dash-dot-dot', 'slant-dash-dot', 'thick',
  ])
  assert.deepEqual(Object.values(SheetBorderTypes).map(normalizeSheetBorderPreset), Object.keys(sheetBorderLabels))
  assert.equal(SheetBorderStyleTypes.SLANT_DASH_DOT, 12)
  assert.equal(SheetBorderStyleTypes.THICK, 13)
})

test('pen validation normalizes HEX while preserving nonempty imported CSS colors', () => {
  assert.equal(normalizeSheetBorderColor(' #AbC '), '#aabbcc')
  assert.equal(normalizeSheetBorderColor('rgb(10, 20, 30)'), 'rgb(10, 20, 30)')
  for (const invalid of [null, 12, '', '   ']) assert.throws(() => normalizeSheetBorderColor(invalid), TypeError)
  for (const invalid of [-1, 14, 1.5, null, '', 'invalid']) assert.throws(() => normalizeSheetBorderStyle(invalid), TypeError)
  for (const invalid of [null, 'diagonal', 1, '']) assert.throws(() => normalizeSheetBorderPreset(invalid), TypeError)
})

test('both coordinate formats normalize and explicit null or invalid bounds are rejected', () => {
  assert.equal(sheetBorderFormatInputRanges(), undefined)
  assert.deepEqual(sheetBorderFormatInputRanges({ ranges: [] }), [])
  assert.deepEqual(sheetBorderFormatInputRanges({ range: { startRow: 1, endRow: 2, startColumn: 3, endColumn: 4 } }), [range(1, 2, 3, 4)])
  assert.deepEqual(sheetBorderFormatInputRanges({ ranges: [range(0), range(3, 4, 2, 3)] }), [range(0), range(3, 4, 2, 3)])
  assert.deepEqual(sheetBorderFormatCommandInput('outside'), { value: 'outside' })
  assert.deepEqual(sheetBorderFormatCommandInput(), {})
  for (const input of [{ ranges: null }, { range: null }, { ranges: 'A1' }, { ranges: [null] }]) assert.throws(() => sheetBorderFormatInputRanges(input), TypeError)
  for (const target of [range(-1), range(2, 1), range(0, 0, 3, 2), range(.5), range(0, 1_048_576), range(0, 0, 0, 16_384)]) {
    assert.throws(() => sheetBorderFormatInputRanges({ ranges: [target] }), RangeError)
  }
})

test('state staging preserves the previous pen across invalid requests', () => {
  const initial = createSheetBorderFormatState()
  assert.deepEqual(initial, { type: 'all', style: 'thin', color: '#000000', active: false })
  const next = borderFormatActionForCommand(commands.basic, { value: { type: 'tlbr', style: 13, color: '#AbC' } }, initial)
  assert.deepEqual(next, { state: { type: 'diagonal-down', style: 'thick', color: '#aabbcc', active: true }, apply: true })
  for (const [id, params] of [[commands.basic, { value: { type: 'left', style: 'bad' } }], [commands.color, ''], [commands.style, 14], [commands.position, 'invalid'], [commands.color, { value: '#fff', activate: 'yes' }]]) {
    assert.throws(() => borderFormatActionForCommand(id, params, initial), TypeError)
  }
  assert.deepEqual(initial, createSheetBorderFormatState())
  assert.equal(borderFormatActionForCommand('sheet.command.copy', {}, initial), undefined)
})

test('inactive API color and style configure the pen while an explicit UI activation applies it', () => {
  const colored = borderFormatActionForCommand(commands.color, '#AbC')
  assert.deepEqual(colored, { state: { type: 'all', style: 'thin', color: '#aabbcc', active: false }, apply: false })
  const styled = borderFormatActionForCommand(commands.style, 7, colored.state)
  assert.equal(styled.apply, false)
  assert.equal(styled.state.style, 'double')
  const positioned = borderFormatActionForCommand(commands.position, 'outside', styled.state)
  assert.deepEqual(positioned, { state: { type: 'outside', style: 'double', color: '#aabbcc', active: true }, apply: true })
  const ui = borderFormatActionForCommand(commands.color, { value: '#ff0000', activate: true })
  assert.equal(ui.apply, true)
  assert.equal(ui.state.type, 'all')
})

test('cardinal presets follow fixed two-by-two edge masks', () => {
  const masks = {
    top: { '0:0': ['top'], '0:1': ['top'] }, bottom: { '1:0': ['bottom'], '1:1': ['bottom'] },
    left: { '0:0': ['left'], '1:0': ['left'] }, right: { '0:1': ['right'], '1:1': ['right'] },
    outside: { '0:0': ['left', 'top'], '0:1': ['right', 'top'], '1:0': ['bottom', 'left'], '1:1': ['bottom', 'right'] },
    inside: { '0:0': ['bottom', 'right'], '0:1': ['bottom', 'left'], '1:0': ['right', 'top'], '1:1': ['left', 'top'] },
    horizontal: { '0:0': ['bottom'], '0:1': ['bottom'], '1:0': ['top'], '1:1': ['top'] },
    vertical: { '0:0': ['right'], '0:1': ['left'], '1:0': ['right'], '1:1': ['left'] },
    all: Object.fromEntries(['0:0', '0:1', '1:0', '1:1'].map(key => [key, ['bottom', 'left', 'right', 'top']])),
  }
  for (const [preset, expected] of Object.entries(masks)) {
    const sheet = createBlankSheet()
    assert.equal(applySheetBorderFormatCommand(sheet, [range(0, 1, 0, 1)], action(preset)), true, preset)
    assert.deepEqual(keys(sheet), expected, preset)
    for (const cell of Object.values(sheet.cells)) for (const edge of Object.values(cell.style.border)) assert.deepEqual(edge, pen)
  }
})

test('five diagonal presets retain the distinct public line contracts', () => {
  const masks = {
    'diagonal-down': ['diagonalDown'], 'diagonal-down-pair': ['topLeftBottomCenter', 'topLeftMiddleRight'],
    'diagonal-down-triple': ['diagonalDown', 'topLeftBottomCenter', 'topLeftMiddleRight'],
    'diagonal-up': ['diagonalUp'], 'diagonal-up-pair': ['bottomCenterTopRight', 'middleLeftTopRight'],
  }
  for (const [preset, expected] of Object.entries(masks)) {
    const sheet = createBlankSheet()
    applySheetBorderFormatCommand(sheet, [range(0, 1, 0, 1)], action(preset, 'slant-dash-dot'))
    assert.deepEqual(keys(sheet), Object.fromEntries(['0:0', '0:1', '1:0', '1:1'].map(key => [key, [...expected].sort()])), preset)
    assert.ok(Object.values(sheet.cells).every(cell => Object.values(cell.style.border).every(edge => edge.style === 'slant-dash-dot')))
  }
})

test('No Border clears only known border keys over a sparse full-sheet range and retains metadata', () => {
  const sheet = createBlankSheet()
  for (const row of [0, 500_000, 1_048_575]) {
    const cell = createCell(row, row === 1_048_575 ? 16_383 : 0, 'kept')
    cell.style.border = { top: pen, bottom: pen, left: pen, right: pen, diagonal: pen, diagonalDown: pen, diagonalUp: pen,
      topLeftBottomCenter: pen, topLeftMiddleRight: pen, middleLeftTopRight: pen, bottomCenterTopRight: pen, vendor: { source: row } }
    sheet.cells[cell.id] = cell
  }
  const original = beforeWithoutBorders(sheet)
  assert.equal(applySheetBorderFormatCommand(sheet, [range(0, 1_048_575, 0, 16_383)], action('none')), true)
  assert.equal(Object.keys(sheet.cells).length, 3)
  for (const cell of Object.values(sheet.cells)) assert.deepEqual(cell.style.border, { vendor: { source: cell.row } })
  assert.deepEqual(beforeWithoutBorders(sheet), original)
  assert.equal(applySheetBorderFormatCommand(sheet, [range(0, 1_048_575, 0, 16_383)], action('none')), false)
})

test('irregular selection unions retain concave outside edges and omit their bounding gaps', () => {
  const ranges = [range(0, 0, 0, 1), range(0, 1, 0, 0)]
  const outside = createBlankSheet(), inside = createBlankSheet()
  applySheetBorderFormatCommand(outside, ranges, action('outside'))
  assert.deepEqual(keys(outside), { '0:0': ['left', 'top'], '0:1': ['bottom', 'right', 'top'], '1:0': ['bottom', 'left', 'right'] })
  applySheetBorderFormatCommand(inside, ranges, action('inside'))
  assert.deepEqual(keys(inside), { '0:0': ['bottom', 'right'], '0:1': ['left'], '1:0': ['top'] })
  assert.equal(outside.cells['1:1'], undefined)
  assert.equal(inside.cells['1:1'], undefined)
})

test('adjacent and disjoint ranges use union topology independent of duplicates and order', () => {
  const left = createBlankSheet(), right = createBlankSheet(), disconnected = createBlankSheet()
  applySheetBorderFormatCommand(left, [range(0), range(0, 0, 1, 1)], action('outside'))
  applySheetBorderFormatCommand(right, [range(0, 0, 1, 1), range(0), range(0)], action('outside'))
  assert.deepEqual(left.cells, right.cells)
  assert.deepEqual(keys(left), { '0:0': ['bottom', 'left', 'top'], '0:1': ['bottom', 'right', 'top'] })
  applySheetBorderFormatCommand(disconnected, [range(0), range(0, 0, 2, 2)], action('outside'))
  assert.equal(disconnected.cells['0:1'], undefined)
  assert.deepEqual(keys(disconnected), { '0:0': ['bottom', 'left', 'right', 'top'], '0:2': ['bottom', 'left', 'right', 'top'] })
})

test('planning never mutates the source and includes changed reverse neighbor edges', () => {
  const sheet = sample().sheets[0]
  sheet.cells['0:1'].style.border = { left: pen, top: { style: 'double', color: '#123456' }, vendor: { id: 'kept' } }
  const original = structuredClone(sheet)
  const plan = createSheetBorderPlan(sourceFor(sheet), [range(0)], 'right', { style: 'thick', color: '#ff0000' })
  assert.deepEqual(sheet, original)
  assert.equal(plan.patches.length, 2)
  assert.deepEqual(plan.patches.find(patch => patch.col === 1).border, { top: { style: 'double', color: '#123456' }, vendor: { id: 'kept' } })
  applySheetBorderFormatCommand(sheet, [range(0)], action('right', 'thick', '#ff0000'))
  assert.deepEqual(border(sheet, 0).right, { style: 'thick', color: '#ff0000' })
  assert.equal(Object.hasOwn(border(sheet, 0, 1), 'left'), false)
  assert.deepEqual(border(sheet, 0, 1).vendor, { id: 'kept' })
})

test('a protected changed neighbor rejects the complete model and view plan atomically', () => {
  const book = sample(), sheet = book.sheets[0]
  sheet.cells['0:1'].style.border = { left: pen }
  sheet.rangeProtectionRules = [lock(range(0, 0, 1, 1))]
  const view = toExcelWorkbookView(book).sheets[0], oldModel = structuredClone(sheet), oldView = structuredClone(view)
  assert.throws(() => applySheetBorderFormatCommand(sheet, [range(0)], action('right'), 'guest'), /相邻|保护/)
  assert.throws(() => applySheetViewBorderFormatCommand(view, [range(0)], action('right'), 'guest'), /相邻|保护/)
  assert.deepEqual(sheet, oldModel)
  assert.deepEqual(view, oldView)
})

test('an unchanged protected neighbor does not reject an otherwise editable border', () => {
  const book = sample(), sheet = book.sheets[0]
  sheet.cells['0:1'].style.border = { top: pen, vendor: 'keep' }
  sheet.rangeProtectionRules = [lock(range(0, 0, 1, 1))]
  const view = toExcelWorkbookView(book).sheets[0]
  assert.equal(applySheetBorderFormatCommand(sheet, [range(0)], action('right'), 'guest'), true)
  assert.equal(applySheetViewBorderFormatCommand(view, [range(0)], action('right'), 'guest'), true)
  assert.deepEqual(border(sheet, 0, 1), { top: pen, vendor: 'keep' })
})

test('partial merged selections expand to the logical cell with segmented perimeter and no inner lines', () => {
  const sheet = createBlankSheet(), merge = range(0, 2, 1, 2)
  sheet.mergedRanges = [merge]
  const plan = createSheetBorderPlan(sourceFor(sheet), [range(1, 1, 2, 2)], 'all', pen)
  assert.deepEqual(plan.ranges, [range(1, 1, 2, 2), merge])
  applySheetBorderFormatCommand(sheet, [range(1, 1, 2, 2)], action('all'))
  assert.deepEqual(border(sheet, 0, 1), { top: pen, bottom: pen, left: pen, right: pen })
  assert.deepEqual(border(sheet, 1, 1), { left: pen })
  assert.deepEqual(border(sheet, 1, 2), { right: pen })
  assert.equal(sheet.cells['0:0'], undefined)
  assert.equal(edges(sheet, merge).length, 10)
  for (const segment of edges(sheet, merge)) assert.deepEqual(segment.edge, pen)
})

test('merged diagonal groups write only the original main cell and retain its content', () => {
  const sheet = createBlankSheet(), merge = range(2, 4, 1, 3)
  sheet.mergedRanges = [merge]
  sheet.cells['2:1'] = createCell(2, 1, 'merged text')
  applySheetBorderFormatCommand(sheet, [range(3, 3, 2, 2)], action('diagonal-down-triple'))
  assert.equal(Object.keys(sheet.cells).length, 1)
  assert.deepEqual(keys(sheet), { '2:1': ['diagonalDown', 'topLeftBottomCenter', 'topLeftMiddleRight'] })
  assert.equal(sheet.cells['2:1'].value, 'merged text')
})

test('clearing A2 beside B1:B3 suppresses only the B2 left segment', () => {
  const sheet = createBlankSheet(), merge = range(0, 2, 1, 1)
  sheet.mergedRanges = [merge]
  applySheetBorderFormatCommand(sheet, [merge], action('all'))
  applySheetBorderFormatCommand(sheet, [range(1)], action('none'))
  assert.deepEqual(edges(sheet, merge).filter(segment => segment.side === 'left').map(segment => segment.edge), [pen, null, pen])
  assert.deepEqual(border(sheet, 0, 1).top, pen)
  assert.deepEqual(border(sheet, 2, 1).bottom, pen)
  assert.equal(sheet.cells['1:0'], undefined)
})

test('clearing an imported main-cell segment materializes the other inherited segments first', () => {
  const sheet = createBlankSheet(), merge = range(0, 2, 1, 1)
  sheet.mergedRanges = [merge]
  sheet.cells['0:1'] = createCell(0, 1, 'imported')
  sheet.cells['0:1'].style.border = { left: pen, right: { style: 'double', color: '#123456' }, vendor: { preserved: true } }
  applySheetBorderFormatCommand(sheet, [range(0)], action('none'))
  assert.deepEqual(edges(sheet, merge).filter(segment => segment.side === 'left').map(segment => segment.edge), [null, pen, pen])
  assert.ok(edges(sheet, merge).filter(segment => segment.side === 'right').every(segment => segment.edge.style === 'double'))
  assert.deepEqual(border(sheet, 0, 1).vendor, { preserved: true })
  assert.equal(sheet.cells['0:1'].value, 'imported')
})

test('disjoint clears beside an imported merge leave gaps inheriting their original edge', () => {
  const sheet = createBlankSheet(), merge = range(0, 4, 1, 1)
  sheet.mergedRanges = [merge]
  sheet.cells['0:1'] = createCell(0, 1); sheet.cells['0:1'].style.border = { left: pen }
  applySheetBorderFormatCommand(sheet, [range(1), range(3), range(1)], action('none'))
  assert.deepEqual(edges(sheet, merge).filter(segment => segment.side === 'left').map(segment => segment.edge), [pen, null, pen, null, pen])
  assert.equal(Object.keys(sheet.cells).length, 3)
  assert.equal(sheet.cells['2:1'], undefined)
})

test('changing a primary fallback also checks permissions for all materialized inherited segments', () => {
  const book = sample(), sheet = book.sheets[0], merge = range(0, 2, 1, 1)
  sheet.mergedRanges = [merge]
  sheet.cells['0:1'].style.border = { left: pen }
  sheet.rangeProtectionRules = [lock(range(2, 2, 1, 1))]
  const view = toExcelWorkbookView(book).sheets[0], oldModel = structuredClone(sheet), oldView = structuredClone(view)
  assert.throws(() => applySheetBorderFormatCommand(sheet, [range(0)], action('none'), 'guest'), /相邻|保护/)
  assert.throws(() => applySheetViewBorderFormatCommand(view, [range(0)], action('none'), 'guest'), /相邻|保护/)
  assert.deepEqual(sheet, oldModel)
  assert.deepEqual(view, oldView)
})

test('a partial merged range cannot bypass protection on the rest of its logical cell', () => {
  const sheet = createBlankSheet(), merge = range(0, 2, 1, 1)
  sheet.mergedRanges = [merge]; sheet.rangeProtectionRules = [lock(range(2, 2, 1, 1))]
  const original = structuredClone(sheet)
  assert.throws(() => applySheetBorderFormatCommand(sheet, [range(1, 1, 1, 1)], action('left'), 'guest'), /保护/)
  assert.deepEqual(sheet, original)
})

test('an explicit null segment blocks imported fallback and repeated neighbor clearing is a no-op', () => {
  const sheet = createBlankSheet(), merge = range(0, 2, 1, 1)
  sheet.mergedRanges = [merge]
  sheet.cells['0:1'] = createCell(0, 1); sheet.cells['0:1'].style.border = { left: pen }
  sheet.cells['1:1'] = createCell(1, 1); sheet.cells['1:1'].style.border = { left: null }
  assert.equal(applySheetBorderFormatCommand(sheet, [range(1)], action('none')), false)
  assert.deepEqual(edges(sheet, merge).filter(segment => segment.side === 'left').map(segment => segment.edge), [pen, null, pen])
})

test('the NONE pen clears a complete merged perimeter without reviving its main-cell summary', () => {
  const sheet = createBlankSheet(), merge = range(0, 2, 1, 2)
  sheet.mergedRanges = [merge]
  applySheetBorderFormatCommand(sheet, [merge], action('outside'))
  assert.equal(applySheetBorderFormatCommand(sheet, [merge], action('outside', 'none')), true)
  assert.ok(edges(sheet, merge).every(segment => segment.edge === null || segment.edge === undefined))
  for (const side of ['top', 'bottom', 'left', 'right']) assert.equal(border(sheet, 0, 1)[side], undefined)
  assert.equal(applySheetBorderFormatCommand(sheet, [merge], action('outside', 'none')), false)
})

test('the NONE diagonal pen clears the selected line group while preserving other diagonal types', () => {
  const sheet = createBlankSheet()
  sheet.cells['0:0'] = createCell(0, 0)
  sheet.cells['0:0'].style.border = { diagonalDown: pen, topLeftBottomCenter: pen, topLeftMiddleRight: pen, diagonalUp: pen, vendor: 'kept' }
  applySheetBorderFormatCommand(sheet, [range(0)], action('diagonal-down-triple', 'none'))
  assert.deepEqual(border(sheet, 0), { diagonalUp: pen, vendor: 'kept' })
})

test('sheet corner borders never read or create neighbors beyond Office row and column limits', () => {
  const positions = [], source = { readBorder: (row, col) => {
    assert.ok(row >= 0 && row <= 1_048_575 && col >= 0 && col <= 16_383)
    positions.push([row, col]); return undefined
  }, existingCells: () => [] }
  const plan = createSheetBorderPlan(source, [range(0), range(1_048_575, 1_048_575, 16_383, 16_383)], 'all', pen)
  assert.equal(plan.patches.length, 2)
  assert.deepEqual(plan.patches.map(({ row, col }) => [row, col]), [[0, 0], [1_048_575, 16_383]])
  assert.ok(positions.length > 0)
})

test('borders include filtered and manually hidden rows in both model and view adapters', () => {
  const book = sample(), sheet = book.sheets[0]
  sheet.autoFilter = { range: range(0, 2), columns: {}, filteredRows: [1] }; sheet.hiddenRows = [2]
  const view = toExcelWorkbookView(book).sheets[0]; view.filteredRows = [1]
  applySheetBorderFormatCommand(sheet, [range(0, 2)], action('all'))
  applySheetViewBorderFormatCommand(view, [range(0, 2)], action('all'))
  for (const row of [0, 1, 2]) {
    assert.deepEqual(border(sheet, row), { top: pen, bottom: pen, left: pen, right: pen })
    assert.deepEqual(view.cells['A' + (row + 1)].border, border(sheet, row))
  }
})

test('model, view and legacy four-argument paths preserve values, caches and unmodified metadata', () => {
  const book = sample(), sheet = book.sheets[0]
  sheet.cells['0:0'].style.border = { vendor: { styleId: 7 } }; sheet.cells['0:0'].style.metadata = { parsed: true }
  sheet.cells['0:1'].style.border = { left: pen, bottom: { style: 'double', color: '#123456' } }
  const view = structuredClone(toExcelWorkbookView(book).sheets[0]), legacy = structuredClone(view.cells), original = beforeWithoutBorders(sheet)
  const ranges = [range(0, 2)], gridRanges = ranges.map(({ startRow, endRow, startCol, endCol }) => ({ startRow, endRow, startColumn: startCol, endColumn: endCol }))
  applySheetBorderFormatCommand(sheet, ranges, action('right'))
  applySheetViewBorderFormatCommand(view, ranges, action('right'))
  applySheetBorders(legacy, gridRanges, 'right', pen)
  assert.deepEqual(beforeWithoutBorders(sheet), original)
  for (const cell of Object.values(sheet.cells)) {
    assert.deepEqual(view.cells[cellKey(cell.row, cell.col)]?.border, cell.style.border)
    assert.deepEqual(legacy[cellKey(cell.row, cell.col)]?.border, cell.style.border)
  }
  const reopened = applyExcelWorkbookView(book, { ...toExcelWorkbookView(book), sheets: [view] }).sheets[0]
  assert.equal(reopened.cells['0:0'].value, '00123')
  assert.equal(reopened.cells['1:0'].formula, '=1/4')
  assert.equal(reopened.cells['1:0'].value, .25)
  assert.deepEqual(reopened.cells['0:0'].style.border.vendor, { styleId: 7 })
  const defaultLegacy = {}
  applySheetBorders(defaultLegacy, [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }], 'left')
  assert.equal(defaultLegacy.A1.border.left.color, '#1f2937')
})

test('equivalent edge objects with different key order do not create content changes or empty cells', () => {
  const sheet = createBlankSheet()
  sheet.cells['0:0'] = createCell(0, 0); sheet.cells['0:0'].style.border = { top: { color: '#aabbcc', style: 'thin' } }
  assert.equal(applySheetBorderFormatCommand(sheet, [range(0)], action('top')), false)
  assert.equal(applySheetBorderFormatCommand(sheet, [range(2)], action('inside')), false)
  assert.equal(applySheetBorderFormatCommand(sheet, [range(3)], action('none')), false)
  assert.equal(Object.keys(sheet.cells).length, 1)
})

test('visible merged segments clip million-row perimeter reads while retaining the original main-cell fallback', () => {
  const merge = range(0, 1_048_575, 1, 1), visible = range(500_000, 500_004, 1, 1), reads = []
  const read = (row, col) => {
    reads.push([row, col])
    return row === 0 ? { left: pen, right: pen } : row === 500_002 ? { left: null } : undefined
  }
  const segments = sheetMergedBorderSegments(merge, read, visible)
  assert.equal(segments.length, 10)
  assert.equal(reads.length, 11)
  assert.deepEqual(reads[0], [0, 1])
  assert.ok(segments.every(segment => segment.row >= visible.startRow && segment.row <= visible.endRow && ['left', 'right'].includes(segment.side)))
  assert.deepEqual(segments.filter(segment => segment.side === 'left').map(segment => segment.edge), [pen, pen, null, pen, pen])
  assert.deepEqual(sheetMergedBorderSegments(range(0, 1, 1, 2), () => ({ top: pen })).length, 8)
  assert.deepEqual(sheetMergedBorderSegments(merge, read, range(5, 8, 2, 3)), [])
})

test('sparse No Border beside a huge imported merge touches only the adjacent selected segments', () => {
  const sheet = createBlankSheet(), merge = range(0, 1_048_575, 1, 1)
  sheet.mergedRanges = [merge]; sheet.cells['0:1'] = createCell(0, 1); sheet.cells['0:1'].style.border = { left: pen }
  const source = sourceFor(sheet); let reads = 0
  const plan = createSheetBorderPlan({ ...source, readBorder: (row, col) => { reads += 1; return source.readBorder(row, col) } }, [range(700_000)], 'none', pen)
  assert.deepEqual(plan.patches, [{ row: 700_000, col: 1, border: { left: null } }])
  assert.ok(reads < 20)
  assert.equal(Object.keys(sheet.cells).length, 1)
})

test('API-only style labels are present for every supported language without widening the menu', () => {
  assert.equal(Object.keys(sheetBorderStyleLabels).length, 12)
  for (const locale of supportedOfficeLocales) {
    const labels = sheetToolbarBorderApiStyleLocales[locale]
    assert.deepEqual(Object.keys(labels).sort(), ['sheet.border.style.none', 'sheet.border.style.slant-dash-dot'])
    assert.ok(Object.values(labels).every(label => typeof label === 'string' && label.length > 0), locale)
  }
})

let vite, OfficeStudio, FOffice, createSheetRibbon
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
  ;({ createSheetRibbon } = await vite.ssrLoadModule('/packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function withStudio(run, initial = sample()) {
  const studio = new OfficeStudio({ workbook: initial, currentUserId: 'local-user' })
  try { return run(studio, FOffice.newAPI(studio)) } finally { studio.dispose() }
}
const historySize = (studio, id) => studio.editor.history.units[id]?.undo.length ?? 0

test('Ribbon and cell, row, column and all-cell context entries use registered border commands', () => {
  const ribbon = createSheetRibbon({ hasSheet: true, fontEditable: true, borderPreset: 'outside', borderOptions: pen })[0].groups.find(group => group.id === 'font').items.find(item => item.id === 'sheet.menu.borders')
  assert.equal(ribbon.commandId, commands.legacy)
  assert.equal(ribbon.params, 'outside')
  assert.equal(ribbon.children.filter(item => item.id.startsWith('sheet.menu.border.')).length, 15)
  assert.equal(ribbon.children.find(item => item.id === 'sheet.menu.border-style').children.length, 12)
  const context = { sheetId: 'sheet', canDeleteRow: true, canDeleteColumn: true, frozen: false, hasHiddenRows: false, hasHiddenColumns: false, canFormat: true }
  for (const hit of [{ kind: 'cell', row: 0, column: 0, key: 'A1' }, { kind: 'row', row: 0 }, { kind: 'column', column: 0 }, { kind: 'all' }]) {
    const item = createSheetContextMenuItems(hit, context).find(item => item.id === 'borders')
    assert.equal(item.children.length, 15)
    assert.deepEqual(item.children.map(child => child.params), Object.keys(sheetBorderLabels))
    assert.ok(item.children.every(child => child.commandId === commands.position && child.disabled === false && child.translationKey === 'sheet.items.sheet.menu.border.' + child.params))
    const denied = createSheetContextMenuItems(hit, { ...context, canFormat: false }).find(item => item.id === 'borders')
    assert.equal(denied.disabled, true)
    assert.ok(denied.children.every(child => child.disabled))
  }
})

test('Facade enums and whole-range setBorder share one history and snapshot contract', () => withStudio((studio, api) => {
  assert.deepEqual(api.Enum.BorderType, SheetBorderTypes)
  assert.deepEqual(api.Enum.BorderStyleTypes, SheetBorderStyleTypes)
  const workbook = api.getActiveWorkbook(), sheet = workbook.getActiveSheet(), target = sheet.getRange('A1:B3'), events = []
  const before = structuredClone(studio.getWorkbookModel(workbook.getId()))
  studio.events.addEvent('CommandExecuted', event => events.push(event))
  assert.equal(target.setBorder(api.Enum.BorderType.OUTSIDE, api.Enum.BorderStyleTypes.DOUBLE, '#AbC'), target)
  assert.equal(historySize(studio, workbook.getId()), 1)
  assert.equal(events.at(-1).id, commands.basic)
  assert.equal(events.at(-1).params.ranges.length, 1)
  assert.deepEqual(events.at(-1).params.ranges[0], range(0, 2, 0, 1))
  const updated = structuredClone(studio.getWorkbookModel(workbook.getId()))
  assert.deepEqual(border(updated.sheets[0], 0).top, { style: 'double', color: '#aabbcc' })
  workbook.undo(); assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets, before.sheets)
  workbook.redo(); assert.deepEqual(studio.getWorkbookModel(workbook.getId()), updated)
  target.setBorder(api.Enum.BorderType.LEFT, api.Enum.BorderStyleTypes.THICK)
  assert.equal(border(studio.getWorkbookModel(workbook.getId()).sheets[0], 0).left.color, '#aabbcc')
  const saved = workbook.save(), reopened = api.createWorkbook({ ...saved, id: 'border-reopened' })
  assert.deepEqual(border(studio.getWorkbookModel(reopened.getId()).sheets[0], 0).top, { style: 'double', color: '#aabbcc' })
  assert.deepEqual(studio.getSheetBorderFormatState(reopened.getId(), sheet.getSheetId()), createSheetBorderFormatState())
}))

test('all six command IDs support current selection and explicit coordinate aliases', () => withStudio((studio, api) => {
  const workbook = api.getActiveWorkbook(), sheet = workbook.getActiveSheet(), events = []
  sheet.getRange('A1:A2').activate()
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  studio.executeCommandSync(commands.basic, { value: { type: 'top', style: 1 } })
  studio.executeCommandSync(commands.position, { value: 'left' })
  studio.executeCommandSync(commands.style, 7)
  studio.executeCommandSync(commands.color, { value: '#F05252' })
  studio.executeCommandSync(commands.set, { unitId: workbook.getId(), subUnitId: sheet.getSheetId(), range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 } })
  studio.executeCommandSync(commands.legacy, { unitId: workbook.getId(), sheetId: sheet.getSheetId(), ranges: [range(3)], value: 'outside' })
  assert.deepEqual([...events].sort(), [...SHEET_BORDER_FORMAT_COMMAND_IDS].sort())
  assert.deepEqual(border(studio.getWorkbookModel(workbook.getId()).sheets[0], 2).left, { style: 'double', color: '#f05252' })
  assert.equal(historySize(studio, workbook.getId()), 6)
}))

test('inactive API pen changes emit state and command events without content history or revision changes', () => withStudio((studio, api) => {
  const workbook = api.getActiveWorkbook(), sheet = workbook.getActiveSheet(), states = [], events = []
  sheet.getRange('A1').activate()
  const original = structuredClone(studio.getWorkbookModel(workbook.getId())), revision = studio.getWorkbookRevision(workbook.getId())
  studio.events.addEvent('SheetBorderFormatStateChanged', event => states.push(event))
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  assert.equal(studio.executeCommandSync(commands.color, { value: '#AbC' }), false)
  assert.equal(studio.executeCommandSync(commands.style, 12), false)
  assert.deepEqual(studio.getWorkbookModel(workbook.getId()), original)
  assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
  assert.equal(historySize(studio, workbook.getId()), 0)
  assert.deepEqual(events, [commands.color, commands.style])
  assert.equal(states.length, 2)
  assert.deepEqual(states[1].state, { type: 'all', style: 'slant-dash-dot', color: '#aabbcc', active: false })
  states[1].state.color = 'changed-by-listener'
  assert.equal(studio.getSheetBorderFormatState(workbook.getId(), sheet.getSheetId()).color, '#aabbcc')
  assert.equal(studio.executeCommandSync(commands.position, { value: 'right' }), true)
  assert.deepEqual(border(studio.getWorkbookModel(workbook.getId()).sheets[0], 0).right, { style: 'slant-dash-dot', color: '#aabbcc' })
  assert.equal(historySize(studio, workbook.getId()), 1)
}))

test('canceled, invalid and protected API requests preserve pen, model, events and history', () => withStudio((studio, api) => {
  const workbook = api.getActiveWorkbook(), sheet = workbook.getActiveSheet(), events = [], states = []
  const initial = structuredClone(studio.getWorkbookModel(workbook.getId()))
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  studio.events.addEvent('SheetBorderFormatStateChanged', event => states.push(event))
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === commands.basic) event.cancel = true })
  assert.throws(() => sheet.getRange('A1').setBorder('right', 'thick', '#ff0000'), /canceled/)
  cancel.dispose()
  assert.throws(() => sheet.getRange('A1').setBorder('right', 'thick', ''), TypeError)
  assert.throws(() => studio.executeCommandSync(commands.basic, { ranges: null, value: { type: 'right', style: 1 } }), TypeError)
  assert.deepEqual(studio.getWorkbookModel(workbook.getId()), initial)
  const protectedModel = structuredClone(initial)
  protectedModel.sheets[0].cells['0:1'].style.border = { left: pen }
  protectedModel.sheets[0].rangeProtectionRules = [lock(range(0, 0, 1, 1))]
  studio.editor.workbook.replace(protectedModel)
  assert.throws(() => sheet.getRange('A1').setBorder('right', 'thick', '#ff0000'), /相邻|保护|permission/i)
  assert.deepEqual(studio.getWorkbookModel(workbook.getId()), protectedModel)
  assert.deepEqual(studio.getSheetBorderFormatState(workbook.getId(), sheet.getSheetId()), createSheetBorderFormatState())
  assert.equal(historySize(studio, workbook.getId()), 0)
  assert.deepEqual(events, [])
  assert.deepEqual(states, [])
}))

test('pen state and history stay isolated across worksheets and workbooks and disposal clears state', () => {
  const initial = sample(), extra = createBlankSheet('Second')
  extra.id = 'border-second-sheet'; initial.sheets.push(extra)
  withStudio((studio, api) => {
    const first = api.getActiveWorkbook(), firstSheet = first.getActiveSheet()
    firstSheet.getRange('A1').setBorder('left', 'double', '#ff0000')
    assert.equal(studio.executeCommandSync(commands.color, { unitId: first.getId(), sheetId: extra.id, ranges: [range(0)], value: '#00ff00' }), false)
    assert.deepEqual(studio.getSheetBorderFormatState(first.getId(), firstSheet.getSheetId()), { type: 'left', style: 'double', color: '#ff0000', active: true })
    assert.deepEqual(studio.getSheetBorderFormatState(first.getId(), extra.id), { type: 'all', style: 'thin', color: '#00ff00', active: false })
    const second = api.createWorkbook({ ...sample(), id: 'border-second-book' })
    second.getActiveSheet().getRange('A1').setBorder('bottom', 'thin', '#0000ff')
    first.undo()
    assert.equal(border(studio.getWorkbookModel(first.getId()).sheets[0], 0).left, undefined)
    assert.equal(border(studio.getWorkbookModel(second.getId()).sheets[0], 0).bottom.color, '#0000ff')
    assert.equal(studio.activeWorkbookId, second.getId())
    first.redo()
    assert.equal(border(studio.getWorkbookModel(first.getId()).sheets[0], 0).left.color, '#ff0000')
    api.disposeUnit(first.getId())
    const recreated = api.createWorkbook({ ...sample(), id: first.getId() })
    assert.deepEqual(studio.getSheetBorderFormatState(recreated.getId(), firstSheet.getSheetId()), createSheetBorderFormatState())
    assert.equal(historySize(studio, recreated.getId()), 0)
  }, initial)
})

test('explicit UI activation applies a configured pen and accepted repeated borders do not add history', () => withStudio((studio, api) => {
  const workbook = api.getActiveWorkbook(), sheet = workbook.getActiveSheet(), states = []
  sheet.getRange('A1').activate()
  studio.events.addEvent('SheetBorderFormatStateChanged', event => states.push(event))
  assert.equal(studio.executeCommandSync(commands.color, { value: '#AbC', activate: true }), true)
  const revision = studio.getWorkbookRevision(workbook.getId())
  assert.equal(studio.executeCommandSync(commands.set, undefined), true)
  assert.equal(historySize(studio, workbook.getId()), 1)
  assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
  assert.equal(states.length, 1)
  const stateCopy = studio.getSheetBorderFormatState(workbook.getId(), sheet.getSheetId()); stateCopy.type = 'none'
  assert.equal(studio.getSheetBorderFormatState(workbook.getId(), sheet.getSheetId()).type, 'all')
}))

test('merge segment nulls survive history and saved snapshot reopening', () => {
  const initial = sample(), merge = range(0, 2, 1, 1)
  initial.sheets[0].mergedRanges = [merge]
  withStudio((studio, api) => {
    const workbook = api.getActiveWorkbook(), sheet = workbook.getActiveSheet()
    sheet.getRange('B1:B3').setBorder('all', 'thin', '#aabbcc')
    sheet.getRange('A2').setBorder('none', 'thin')
    const leftEdges = id => edges(studio.getWorkbookModel(id).sheets[0], merge).filter(segment => segment.side === 'left').map(segment => segment.edge)
    assert.deepEqual(leftEdges(workbook.getId()), [pen, null, pen])
    workbook.undo(); assert.deepEqual(leftEdges(workbook.getId()), [pen, pen, pen])
    workbook.redo(); assert.deepEqual(leftEdges(workbook.getId()), [pen, null, pen])
    const reopened = api.createWorkbook({ ...workbook.save(), id: 'border-null-reopened' })
    assert.deepEqual(leftEdges(reopened.getId()), [pen, null, pen])
  }, initial)
})
