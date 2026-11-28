import assert from 'node:assert/strict'
import { test } from 'vitest'
import { shallowReactive } from 'vue'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { serializeValidationListItems } from '../src/features/sheet-validation/validation-engine.ts'
import { createSheetValidationTagOptionsResolver, sheetValidationCellTagOptions } from '../src/features/sheet-validation/validation-cell-options.ts'
import { layoutSheetValidationTagCell, sheetValidationCellArrowSpace, sheetValidationTagLabels,
  sheetValidationSingleOptionColor } from '../src/features/sheet-validation/validation-tag-layout.ts'
import { useSheetValidationFeedback } from '../src/components/spreadsheet/use-sheet-validation-feedback.ts'
import { calculateSheetAxisAutoFit } from '../src/features/sheet-workbench/sheet-auto-fit.ts'
import { createSheetValidationTagCommands } from '../src/rendering/spreadsheet/validation-cell-tags.ts'
import { renderRectPoints } from '../src/rendering/affine-command-geometry.ts'

// Implementation-phase drafts: no test runner, browser, build or Office roundtrip has been executed.
const area = (row, column, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
const point = (row, column) => ({ row, column })
const scalar = value => ({ value, rawValue: value, dataType: 'text', fontSize: 11 })
const measure = text => [...text].length * 8
const measurePaint = (_font, text) => measure(text)
const theme = { fill: '#eeeeee', text: '#222222' }
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8)
function list(values, extra = {}) {
  return { id: 'list', type: 'listMultiple', ranges: [area(0, 0)], formula1: serializeValidationListItems(values), ...extra }
}
function model() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'tag-candidates'; workbook.activeSheetId = sheet.id = 'data'; workbook.calculationMode = 'manual'
  sheet.rowCount = 80; sheet.columnCount = 12
  return { workbook, sheet }
}

test('CSV tokens preserve punctuation and invalid choices while canonical candidate spelling and colors stay authoritative', () => {
  const rule = list(['Red', 'Green, dark', 'It"s blue'], { formula2: '#ff0000,#00ff00,#0000ff' })
  const cell = scalar(' red ,"Green, dark",It""s blue,missing'), before = JSON.stringify({ rule, cell })
  assert.deepEqual(sheetValidationTagLabels(rule, cell), [
    { text: 'Red', color: '#ff0000' }, { text: 'Green, dark', color: '#00ff00' },
    { text: 'It"s blue', color: '#0000ff' }, { text: 'missing', color: undefined },
  ])
  assert.equal(JSON.stringify({ rule, cell }), before)
  const candidates = [{ label: 'RED', color: '#abcdef' }, { label: 'red', color: '#123456' }]
  assert.deepEqual(sheetValidationTagLabels(rule, scalar('red'), candidates), [{ text: 'RED', color: '#abcdef' }])
  assert.deepEqual(sheetValidationTagLabels(rule, scalar('red'), []), [{ text: 'red', color: undefined }])
})

test.each(['arrow', 'text'])('single %s mode keeps ordinary text and derives only the selected cell background', renderMode => {
  const rule = list(['Red', 'Blue'], { type: 'list', renderMode, formula2: '#ff0000,#0000ff' }), cell = scalar('red')
  assert.equal(layoutSheetValidationTagCell(rule, cell, { width: 150, measure }), undefined)
  assert.equal(sheetValidationSingleOptionColor(rule, cell), '#ff0000')
  assert.equal(sheetValidationSingleOptionColor(rule, scalar('missing')), undefined)
  assert.equal(sheetValidationSingleOptionColor(rule, cell, [{ label: 'Red', color: '#008800' }]), '#008800')
  assert.equal(sheetValidationSingleOptionColor(rule, cell, []), undefined)
  assert.equal(sheetValidationCellArrowSpace(rule), renderMode === 'arrow' ? 22 : 0)
})

test('formatted range candidates and reference colors are read once per rule in one sparse batch', () => {
  const h = model()
  h.sheet.cells['0:2'] = createCell(0, 2, .25); h.sheet.cells['0:2'].style.numberFormat = '0%'
  h.sheet.cells['1:2'] = createCell(1, 2, 1234.5); h.sheet.cells['1:2'].style.numberFormat = '0.00'
  h.sheet.dataValidations = [list([], { ranges: [area(0, 0, 29, 0)], formula1: '=$C$1:$C$2', formula2: '#ff0000,#008000' })]
  let firstReads = 0, secondReads = 0
  Object.defineProperty(h.sheet.cells['0:2'], 'value', { enumerable: true, configurable: true, get: () => { firstReads++; return .25 } })
  Object.defineProperty(h.sheet.cells['1:2'], 'value', { enumerable: true, configurable: true, get: () => { secondReads++; return 1234.5 } })
  const candidates = sheetValidationCellTagOptions(h.workbook, h.sheet.id, Array.from({ length: 30 }, (_, row) => point(row, 0)), 'local-user', 'de-DE')
  assert.equal(firstReads, 1); assert.equal(secondReads, 1); assert.equal(candidates.size, 30)
  assert.deepEqual(candidates.get('A1'), [{ label: '25%', color: '#ff0000' }, { label: '1234,50', color: '#008000' }])
  assert.equal(candidates.get('A1'), candidates.get('A30'))
  for (const item of candidates.get('A1')) assert.deepEqual(Object.keys(item).sort(), ['color', 'label'])
})

test('candidate requests normalize to merged main cells and enforce View on the complete merge and every list source', () => {
  const h = model(); h.sheet.mergedRanges = [area(0, 0, 1, 1)]
  h.sheet.dataValidations = [list(['Red', 'Blue'])]
  assert.deepEqual([...sheetValidationCellTagOptions(h.workbook, h.sheet.id, [point(1, 1), point(0, 0), point(-1, 0), point(80, 0)], 'local-user').keys()], ['A1'])
  h.sheet.rangeProtectionRules = [{ id: 'merge-lock', ranges: [area(1, 1)], ownerUserId: 'peer', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'noOneElseCanView' }]
  assert.equal(sheetValidationCellTagOptions(h.workbook, h.sheet.id, [point(0, 0)], 'local-user').size, 0)
  h.sheet.mergedRanges = []; h.sheet.cells['0:2'] = createCell(0, 2, 'Red')
  h.sheet.dataValidations[0].formula1 = '=$C$1'
  h.sheet.rangeProtectionRules[0].ranges = [area(0, 2)]
  assert.deepEqual(sheetValidationCellTagOptions(h.workbook, h.sheet.id, [point(0, 0)], 'local-user').get('A1'), [])
  assert.equal(createSheetValidationTagOptionsResolver(h.workbook, 'absent', 'local-user')(0, 0), undefined)
})

test('paint-only candidate reads remain stable under a popup blocker and reject inactive hosts and replaced view owners', () => {
  const h = model(); h.sheet.dataValidations = [list(['Red'], { formula2: '#ff0000' })]
  const state = shallowReactive({ model: toExcelWorkbookView(h.workbook), sheetId: h.sheet.id, userId: 'local-user',
    region: 'en-US', active: true, blocked: false })
  const reader = useSheetValidationFeedback(() => state)
  assert.deepEqual(reader.readTagOptions([point(0, 0)]).get('A1'), [{ label: 'Red', color: '#ff0000' }])
  state.blocked = true
  assert.equal(reader.read([point(0, 0)]).size, 0)
  assert.deepEqual(reader.readTagOptions([point(0, 0)]).get('A1'), [{ label: 'Red', color: '#ff0000' }])
  state.active = false; assert.equal(reader.readTagOptions([point(0, 0)]).size, 0)
  state.active = true; state.model = { ...toExcelWorkbookView(h.workbook), sheets: [] }
  assert.equal(reader.readTagOptions([point(0, 0)]).size, 0)
})

test.each(['revision', 'generation', 'user', 'region', 'owner', 'disposed'])('candidate derivation discards the entire result after %s changes during a source read', change => {
  const h = model(); h.sheet.cells['0:2'] = createCell(0, 2, 'Red')
  h.sheet.dataValidations = [list([], { formula1: '=$C$1', formula2: '#ff0000' })]
  let revision = 0, generation = 1
  const studio = { activeWorkbookId: h.workbook.id, isDisposed: false, getWorkbookModel: () => h.workbook, getWorkbookRevision: () => revision,
    editor: { state: { activeKind: 'workbook' }, history: { captureSavepoint: () => ({ generation }) } } }
  const state = shallowReactive({ studio, model: toExcelWorkbookView(h.workbook), sheetId: h.sheet.id,
    userId: 'local-user', region: 'en-US', active: true, blocked: false })
  Object.defineProperty(h.sheet.cells['0:2'], 'value', { enumerable: true, configurable: true, get: () => {
    if (change === 'revision') revision++
    if (change === 'generation') generation++
    if (change === 'user') state.userId = 'peer'
    if (change === 'region') state.region = 'de-DE'
    if (change === 'owner') studio.activeWorkbookId = 'replaced'
    if (change === 'disposed') studio.isDisposed = true
    return 'Red'
  } })
  assert.equal(useSheetValidationFeedback(() => state).readTagOptions([point(0, 0)]).size, 0)
})

test('multiple selected tags wrap as items even when cell text uses overflow or clip', () => {
  const rule = list(['AAAA', 'BBBB']), cell = scalar('AAAA,BBBB')
  for (const wrapStrategy of ['overflow', 'clip']) {
    const layout = layoutSheetValidationTagCell(rule, { ...cell, wrapStrategy }, { width: 70, fit: 'row', measure })
    assert.equal(layout.lines.length, 2); assert.deepEqual(layout.lines.map(line => line.items.map(item => item.text)), [['AAAA'], ['BBBB']])
    near(layout.width, 44); near(layout.height, (11 * 4 / 3 + 6) * 2 + 4)
  }
})

test('long labels wrap internally, retain explicit breaks and center smaller tags within the tallest line', () => {
  const rule = list(['abcdefghij', 'a', 'b\nc'])
  const wrapped = layoutSheetValidationTagCell(rule, scalar('abcdefghij'), { width: 90, fit: 'row', measure })
  assert.equal(wrapped.lines[0].items[0].fragments.length, 1)
  const wrapping = layoutSheetValidationTagCell(rule, { ...scalar('abcdefghij'), wrapStrategy: 'wrap' }, { width: 90, fit: 'row', measure })
  assert.deepEqual(wrapping.lines[0].items[0].fragments.map(fragment => fragment.text), ['abcdef', 'ghij'])
  const mixed = layoutSheetValidationTagCell(rule, scalar('a,"b\nc"'), { width: 200, fit: 'row', measure })
  assert.equal(mixed.lines.length, 1); near(mixed.lines[0].height, 11 * 8 / 3 + 6)
  near(mixed.lines[0].items[0].y, 11 * 2 / 3); near(mixed.lines[0].items[1].y, 0)
})

test('stacked labels keep complete CJK, emoji and combining graphemes in their authoring order', () => {
  const rule = list(['汉👩‍💻é']), cell = { ...scalar('汉👩‍💻é'), rotationVertical: true }
  const layout = layoutSheetValidationTagCell(rule, cell, { width: 100, fit: 'row', measure })
  assert.deepEqual(layout.lines[0].items[0].fragments.map(fragment => fragment.text), ['汉', '👩‍💻', 'é'])
  near(layout.height, 3 * 11 * 4 / 3 + 6); assert.equal(layout.rotation, 0)
  const rect = { x: 0, y: 0, width: 100, height: 70 }
  const commands = createSheetValidationTagCommands(rule, cell, rect, rect, 1, theme, measurePaint)
  assert.deepEqual(commands.filter(command => command.kind === 'text').map(command => command.text), ['汉', '👩‍💻', 'é'])
})

test.each([-90, -45, 45, 90])('rotation %s transforms a tag background, label and decorations as one block', rotation => {
  const rule = list(['abcd']), cell = { ...scalar('abcd'), rotation, underline: true, strike: true, vertical: 'top' }
  const rect = { x: 20, y: 30, width: 160, height: 160 }
  const commands = createSheetValidationTagCommands(rule, cell, rect, rect, 1, theme, measurePaint)
  assert.deepEqual(commands.map(command => command.kind), ['polygon', 'text', 'line', 'line'])
  const matrix = commands[0].matrix, angle = rotation * Math.PI / 180
  near(matrix.a, Math.cos(angle)); near(matrix.b, Math.sin(angle))
  assert.ok(commands.every(command => command.matrix === matrix))
  const points = renderRectPoints({ x: 0, y: 0, width: 44, height: 11 * 4 / 3 + 6 }, matrix)
  near(Math.min(...points.map(point => point.x)), rect.x + 4)
  near(Math.min(...points.map(point => point.y)), rect.y + 4)
  for (const command of commands.slice(1)) assert.ok(command.localClip.width > 0 && command.localClip.height > 0)
})

test.each(['left', 'right', 'center'])('%s alignment shares indentation and hidden-arrow geometry with automatic dimensions', align => {
  const rule = list(['abcd']), cell = { ...scalar('abcd'), indent: 2, align }
  const layout = layoutSheetValidationTagCell(rule, cell, { width: 150, fit: 'row', measure })
  near(layout.leftInset, 4 + (align === 'left' ? 2 * 11 * 4 / 3 : 0))
  near(layout.rightInset, 26 + (align === 'right' ? 2 * 11 * 4 / 3 : 0))
  const hidden = layoutSheetValidationTagCell({ ...rule, showDropDown: false }, cell, { width: 150, fit: 'row', measure })
  near(layout.rightInset - hidden.rightInset, 22)
  assert.equal(sheetValidationCellArrowSpace({ ...rule, showDropDown: false }), 0)
})

test('shrink-to-fit scales geometry uniformly and automatic columns keep the full label width', () => {
  const rule = list(['abcdefghijk']), cell = { ...scalar('abcdefghijk'), shrinkToFit: true, underline: true, strike: true }
  const rect = { x: 0, y: 0, width: 80, height: 38 }
  const row = layoutSheetValidationTagCell(rule, cell, { width: rect.width, fit: 'row', measure })
  near(row.scale, .5); near(row.size.width, 50)
  const column = layoutSheetValidationTagCell(rule, cell, { width: rect.width, height: rect.height, fit: 'column', measure })
  near(column.scale, 1); near(column.size.width, 100)
  const commands = createSheetValidationTagCommands(rule, cell, rect, rect, 1, theme, measurePaint)
  assert.ok(commands.every(command => command.matrix === commands[0].matrix)); near(commands[0].matrix.a, .5)
})

test('row height and column width agree with the independently measured two-tag box after automatic sizing', () => {
  const rule = list(['A', 'B']), cell = scalar('A,B')
  const sheet = { id: 'data', name: 'Data', rowCount: 20, columnCount: 10, defaultRowHeight: 26, defaultColumnWidth: 160,
    cells: { A1: cell }, dataValidations: [rule], rowHeights: {}, columnWidths: {} }
  const height = calculateSheetAxisAutoFit(sheet, 'row', 0, { measure }), width = calculateSheetAxisAutoFit(sheet, 'column', 0, { measure })
  assert.equal(height, 29); assert.equal(width, 74)
  const rect = { x: 0, y: 0, width, height }
  const painted = createSheetValidationTagCommands(rule, cell, rect, rect, 1, theme, measurePaint)
  assert.deepEqual(painted.filter(command => command.kind === 'text').map(command => command.text), ['A', 'B'])
  for (const command of painted.filter(command => command.kind === 'polygon')) {
    const points = command.points.map(point => ({ x: command.matrix.a * point.x + command.matrix.c * point.y + command.matrix.e,
      y: command.matrix.b * point.x + command.matrix.d * point.y + command.matrix.f }))
    assert.ok(points.every(point => point.x >= 4 && point.x <= width - 26 + 1e-8 && point.y >= 4 && point.y <= height - 4))
  }
})

test('painting a manual-height multiline tag emits only visible text fragments while preserving the full natural layout', () => {
  const rule = list(['first\nsecond']), cell = { ...scalar('first\nsecond'), vertical: 'top', underline: true }
  const rect = { x: 0, y: 0, width: 160, height: 24 }
  const commands = createSheetValidationTagCommands(rule, cell, rect, rect, 1, theme, measurePaint)
  assert.deepEqual(commands.filter(command => command.kind === 'text').map(command => command.text), ['first'])
  assert.equal(commands.filter(command => command.kind === 'line').length, 1)
  const natural = layoutSheetValidationTagCell(rule, cell, { width: rect.width, fit: 'row', measure })
  assert.deepEqual(natural.lines[0].items[0].fragments.map(fragment => fragment.text), ['first', 'second'])
  assert.ok(natural.size.height + natural.verticalInset * 2 > rect.height)
})
