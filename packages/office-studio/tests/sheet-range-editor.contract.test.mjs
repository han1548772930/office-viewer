import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell, defaultCellStyle } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { applySheetClearCommand, applySheetViewClearCommand, clearModeForCommand, clearSpreadsheetCellContent,
  SHEET_CLEAR_COMMANDS, SHEET_CLEAR_COMMAND_ALIASES, sheetClearModeForOptions } from '../src/features/sheet-workbench/sheet-clear-command.ts'
import { applySheetMergeCommand, applySheetViewMergeCommand, createSheetMergePlan, mergeActionForCommand,
  previewSheetMergeCommand, SHEET_MERGE_COMMANDS, SHEET_MERGE_COMMAND_ALIASES } from '../src/features/sheet-workbench/sheet-merge-command.ts'
import { normalizeSheetCommandRange, partitionSheetCommandCells, sheetCommandRangesWithoutFilteredRows,
  sheetRangeCommandInputRanges, subtractSheetCommandRange } from '../src/features/sheet-workbench/sheet-range-command.ts'
import { validateSheetDataValidationChanges } from '../src/features/sheet-validation/validation-engine.ts'

const range = (startRow, startCol, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
const coordinates = value => ({ startRow: value.startRow, endRow: value.endRow, startColumn: value.startCol, endColumn: value.endCol })
const copy = value => structuredClone(value)
const merge = (mode = 'all', options = {}) => ({ kind: 'merge', mode, defaultMerge: true, isForceMerge: false, ...options })
function points(ranges) {
  const result = new Set()
  for (const item of ranges) for (let row = item.startRow; row <= item.endRow; row++) for (let col = item.startCol; col <= item.endCol; col++) result.add(`${row}:${col}`)
  return [...result].sort()
}
function protect(ranges) {
  return [{ id: 'protected', permissionId: 'protected-p', ownerUserId: 'owner', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'othersCanView', ranges }]
}
function validation(ranges, extra = {}) {
  return { id: 'validation', type: 'custom', formula1: '=A1>0', allowBlank: false,
    showErrorMessage: true, errorStyle: 'stop', error: '数值必须大于零。', ranges, ...extra }
}
function conditional(ranges) {
  return { id: 'conditional', type: 'expression', operator: 'equal', formula1: '=A1>0', style: { fill: '#123456' }, priority: 1, stopIfTrue: false, ranges, metadata: { preserved: true } }
}

test('clear defaults to all; the public flags, compatibility flags and fixed-version aliases remain distinct', () => {
  for (const options of [{}, { contentsOnly: false }, { formatOnly: false }, { contentsOnly: true, formatOnly: true }]) assert.equal(sheetClearModeForOptions(options), 'all')
  assert.equal(sheetClearModeForOptions({ contentsOnly: true }), 'content')
  assert.equal(sheetClearModeForOptions({ formatOnly: true }), 'format')
  assert.equal(sheetClearModeForOptions({ contents: false, format: false }), 'none')
  assert.equal(sheetClearModeForOptions({ contents: true, format: false }), 'content')
  assert.equal(sheetClearModeForOptions({ contents: false, format: true }), 'format')
  for (const name of ['content', 'format', 'all']) assert.equal(clearModeForCommand(SHEET_CLEAR_COMMAND_ALIASES[name]), name)
  for (const options of [null, false, [], { contentsOnly: 'yes' }, { format: null }]) assert.throws(() => sheetClearModeForOptions(options), TypeError)
  assert.equal(clearModeForCommand('unrelated'), undefined)
})

test('range inputs reject malformed data and retain the exact target rather than treating it as the active selection', () => {
  const target = range(2, 3, 4, 5)
  assert.deepEqual(normalizeSheetCommandRange(coordinates(target)), target)
  assert.deepEqual(sheetRangeCommandInputRanges({ unitId: 'book', subUnitId: 'sheet', selections: [coordinates(target)] }), [target])
  assert.deepEqual(sheetRangeCommandInputRanges({ range: target, ranges: [] }), [])
  for (const item of [null, false, {}, range(-1, 0), range(0, 0, 1_048_576, 0), range(0, 16_384), range(0, 0, 1.5, 0)]) assert.throws(() => normalizeSheetCommandRange(item), RangeError)
  for (const input of [null, [], { ranges: null }, { selections: {} }, { unitId: 2 }, { sheetId: '' }]) assert.throws(() => sheetRangeCommandInputRanges(input), TypeError)
})

test('rectangle subtraction preserves every cell outside the cut without overlaps, including edge and point intersections', () => {
  const source = range(1, 1, 6, 6)
  for (const cut of [range(3, 3), range(1, 1, 6, 1), range(0, 0, 3, 3), range(4, 4, 8, 8), range(0, 0, 9, 9), range(8, 8)]) {
    const result = subtractSheetCommandRange(source, cut), cutPoints = new Set(points([cut]))
    assert.deepEqual(points(result), points([source]).filter(point => !cutPoints.has(point)))
    assert.equal(result.reduce((count, item) => count + (item.endRow - item.startRow + 1) * (item.endCol - item.startCol + 1), 0), points(result).length)
  }
  assert.deepEqual(sheetCommandRangesWithoutFilteredRows([range(0, 0, 5, 2)], [4, 1, 1]), [range(0, 0, 0, 2), range(2, 0, 3, 2), range(5, 0, 5, 2)])
})

test('sparse cell partitioning handles tall merges, row merges and cells outside every scope', () => {
  const scopes = [range(0, 0, 50, 0), range(0, 2, 0, 6), range(20, 2, 20, 6), range(40, 2, 45, 6)]
  const cells = [{ row: 45, col: 6, id: 'bottom' }, { row: 50, col: 0, id: 'tall' }, { row: 20, col: 3, id: 'middle' },
    { row: 0, col: 6, id: 'top' }, { row: 19, col: 3, id: 'outside' }, { row: 99, col: 0, id: 'past-end' }]
  assert.deepEqual(partitionSheetCommandCells(scopes, cells).map(group => group.map(cell => cell.id)), [['tall'], ['top'], ['middle'], ['bottom']])
  assert.deepEqual(cells.map(cell => cell.id), ['bottom', 'tall', 'middle', 'top', 'outside', 'past-end'])
})

test('merge direction uses fixed Dimension values, deduplicates identical selections and rejects crossing rectangles', () => {
  assert.equal(mergeActionForCommand(SHEET_MERGE_COMMAND_ALIASES.merge, { value: 0 }).mode, 'vertical')
  assert.equal(mergeActionForCommand(SHEET_MERGE_COMMAND_ALIASES.merge, { value: 1 }).mode, 'horizontal')
  assert.deepEqual(createSheetMergePlan([], [range(0, 0, 1, 2), range(0, 0, 1, 2)], merge('horizontal')).added, [range(0, 0, 0, 2), range(1, 0, 1, 2)])
  assert.deepEqual(createSheetMergePlan([], [range(0, 0, 1, 2)], merge('vertical')).added, [range(0, 0, 1, 0), range(0, 1, 1, 1), range(0, 2, 1, 2)])
  assert.throws(() => createSheetMergePlan([], [range(0, 0, 1, 1), range(1, 1, 2, 2)], merge()), /交叉/)
  assert.throws(() => createSheetMergePlan([range(0, 0, 1, 1)], [range(0, 0, 2, 2)], merge()), /重叠/)
  assert.deepEqual(createSheetMergePlan([range(0, 0, 1, 1)], [range(0, 0, 1, 1)], merge('all', { isForceMerge: true })).added, [])
})

for (const [name, donor] of [['zero', createCell(0, 1, 0)], ['false', createCell(0, 1, false)],
  ['formula', { ...createCell(0, 1, null), formula: '=2+3', display: '5', dataType: 'formula', cachedDataType: 'number' }],
  ['rich text', { ...createCell(0, 1, null), richText: [{ text: '保留格式', style: { bold: true } }] }]]) {
  test(`merge preserves the first non-empty ${name}, including its identity, style and cell extensions`, () => {
    const sheet = createBlankSheet(), source = copy(donor)
    source.style.font.bold = true; source.style.metadata = { vendorStyle: 'preserved' }; source.metadata = { vendorCell: 'preserved' }
    sheet.cells['0:0'] = createCell(0, 0, null); sheet.cells['0:1'] = source; sheet.cells['1:0'] = createCell(1, 0, 'later')
    const result = applySheetMergeCommand(sheet, [range(0, 0, 1, 1)], merge())
    assert.equal(result.discardedContentCount, 1); assert.equal(result.changed, true)
    assert.deepEqual(sheet.cells['0:0'], { ...source, id: '0:0', row: 0, col: 0 })
    assert.equal(sheet.cells['0:1'].value, null); assert.equal(sheet.cells['0:1'].formula, null)
    assert.deepEqual(sheet.cells['0:1'].style, source.style)
    assert.deepEqual(sheet.cells['0:1'].metadata, source.metadata)
  })
}

test('confirmation preview leaves the sheet unchanged, counts only discarded entries and does not prompt for a single donor', () => {
  const sheet = createBlankSheet(); sheet.cells['1:1'] = createCell(1, 1, 'single donor')
  const action = merge('all', { defaultMerge: false }), initial = copy(sheet)
  assert.equal(previewSheetMergeCommand(sheet, [range(0, 0, 1, 1)], action).requiresConfirmation, false)
  assert.deepEqual(sheet, initial)
  sheet.cells['0:0'] = createCell(0, 0, 'master')
  const before = copy(sheet), result = applySheetMergeCommand(sheet, [range(0, 0, 1, 1)], action)
  assert.equal(result.requiresConfirmation, true); assert.equal(result.discardedContentCount, 1); assert.equal(result.changed, false)
  assert.deepEqual(sheet, before)
})

test('content clearing invalidates formula and rich caches while retaining styles and unknown cell data', () => {
  const cell = { ...createCell(0, 1, false), formula: '=1/0', display: '#DIV/0!', dataType: 'formula', cachedDataType: 'error',
    richText: [{ text: 'obsolete' }], hyperlink: 'https://example.test', sparkline: { sourceRange: 'A1:A3', type: 'line' }, metadata: { formulaError: '#DIV/0!', extension: { kept: true } } }
  const cleared = clearSpreadsheetCellContent(cell)
  for (const field of ['value', 'formula', 'cachedDataType', 'hyperlink']) assert.equal(cleared[field], null)
  assert.equal(cleared.display, ''); assert.equal(cleared.dataType, 'blank'); assert.equal(cleared.richText, undefined); assert.equal(cleared.sparkline, undefined)
  assert.deepEqual(cleared.style, cell.style); assert.deepEqual(cleared.metadata, { extension: { kept: true } })
  assert.equal(cell.metadata.formulaError, '#DIV/0!')
})

test('forced merges and intersecting unmerges check the complete affected merge permission before writing', () => {
  const sheet = createBlankSheet(); sheet.cells['0:0'] = createCell(0, 0, 'protected master'); sheet.mergedRanges = [range(0, 0, 2, 2)]
  sheet.rangeProtectionRules = protect([range(2, 2)])
  const initial = copy(sheet)
  for (const action of [merge('all', { isForceMerge: true }), { kind: 'unmerge' }]) {
    assert.throws(() => applySheetMergeCommand(sheet, [range(0, 0, 0, 1)], action, 'guest'), /保护/)
    assert.deepEqual(sheet, initial)
  }
})

test('unmerge covers all intersecting selections, restores master formatting and keeps actual border sides and explicit nulls', () => {
  const sheet = createBlankSheet(), master = createCell(0, 0, 'kept')
  master.style.font.bold = true; master.style.fill = '#abcdef'; master.style.border = {
    top: { style: 'thin', color: '#100000' }, bottom: { style: 'thin', color: '#200000' },
    left: { style: 'thin', color: '#300000' }, right: { style: 'thin', color: '#400000' },
  }
  sheet.cells['0:0'] = master; sheet.cells['0:1'] = createCell(0, 1, null)
  sheet.cells['0:1'].style.border = { left: null, right: { style: 'dashed', color: '#500000' } }
  sheet.cells['4:4'] = createCell(4, 4, 'other'); sheet.mergedRanges = [range(0, 0, 1, 1), range(4, 4, 4, 5)]
  const result = applySheetMergeCommand(sheet, [range(1, 1), range(4, 5)], { kind: 'unmerge' })
  assert.equal(result.changed, true); assert.deepEqual(sheet.mergedRanges, [])
  assert.equal(sheet.cells['1:1'].style.font.bold, true); assert.equal(sheet.cells['1:1'].style.fill, '#abcdef')
  assert.equal(sheet.cells['0:1'].style.border.left, null)
  assert.deepEqual(sheet.cells['0:1'].style.border.right, { style: 'dashed', color: '#500000' })
  assert.equal(sheet.cells['0:0'].style.border.right, undefined); assert.equal(sheet.cells['0:0'].style.border.bottom, undefined)
  assert.equal(sheet.cells['1:1'].value, null); assert.equal(sheet.cells['4:4'].value, 'other')
})

test('standalone merge and clear markers survive applying the view to an older imported model', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, null)
  sheet.cells['0:1'] = { ...createCell(0, 1, 'donor'), richText: [{ text: 'donor', style: { italic: true } }], metadata: { vendor: 'donor' } }
  sheet.cells['0:1'].style.metadata = { vendor: 'donor-style' }
  const view = toExcelWorkbookView(workbook), viewSheet = view.sheets[0]
  applySheetViewMergeCommand(viewSheet, [range(0, 0, 0, 1)], merge())
  let result = applyExcelWorkbookView(workbook, view)
  assert.deepEqual(result.sheets[0].cells['0:0'].metadata, { vendor: 'donor' })
  assert.deepEqual(result.sheets[0].cells['0:0'].richText, sheet.cells['0:1'].richText)
  assert.equal(result.sheets[0].cells['0:1'].richText, undefined)
  applySheetViewClearCommand(viewSheet, [range(0, 0)], 'content')
  result = applyExcelWorkbookView(workbook, view)
  assert.equal(result.sheets[0].cells['0:0'].richText, undefined)
  applySheetViewClearCommand(viewSheet, [range(0, 1)], 'format')
  result = applyExcelWorkbookView(workbook, view)
  assert.equal(result.sheets[0].cells['0:1'].style.metadata, undefined)
})

test('a default-style standalone donor replaces old master extensions, links and borders instead of inheriting them', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  const master = createCell(0, 0, null)
  master.hyperlink = 'https://old-master.example.test'
  master.metadata = { previousExtension: true, formulaError: 'obsolete' }
  master.style.metadata = { previousStyleExtension: true }
  master.style.font.bold = true; master.style.fill = '#123456'
  master.style.border = { right: { style: 'thin', color: '#abcdef' } }
  master.style.rotation = 30; master.style.shrinkToFit = true
  sheet.cells['0:0'] = master; sheet.cells['0:1'] = createCell(0, 1, false)
  const view = toExcelWorkbookView(workbook)
  applySheetViewMergeCommand(view.sheets[0], [range(0, 0, 0, 1)], merge())
  const restored = applyExcelWorkbookView(workbook, view).sheets[0].cells['0:0']
  assert.equal(restored.value, false); assert.equal(restored.hyperlink, null)
  assert.equal(restored.metadata, undefined); assert.equal(restored.style.metadata, undefined)
  assert.equal(restored.style.font.bold, false); assert.equal(restored.style.fill, null)
  assert.deepEqual(restored.style.border, {}); assert.equal(restored.style.rotation, 0)
  assert.equal(restored.style.shrinkToFit, false); assert.equal(restored.cachedDataType, null)
  assert.equal(sheet.cells['0:0'].hyperlink, 'https://old-master.example.test')
})

test('whole-sheet clear remains sparse, skips filtered rows and still processes manually hidden rows', () => {
  const sheet = createBlankSheet(); sheet.rowCount = 1_048_576; sheet.columnCount = 16_384
  for (const row of [0, 1, 2, 1_048_575]) sheet.cells[`${row}:0`] = createCell(row, 0, `value ${row}`)
  sheet.hiddenRows = [2]; sheet.mergedRanges = [range(3, 0, 4, 1)]
  const filtered = copy(sheet.cells['1:0'])
  assert.equal(applySheetClearCommand(sheet, [range(0, 0, 1_048_575, 16_383)], 'content', 'local-user', [1]), true)
  assert.equal(Object.keys(sheet.cells).length, 4); assert.deepEqual(sheet.cells['1:0'], filtered)
  assert.equal(sheet.cells['2:0'].value, null); assert.deepEqual(sheet.hiddenRows, [2]); assert.deepEqual(sheet.mergedRanges, [range(3, 0, 4, 1)])
  assert.equal(applySheetClearCommand(sheet, [range(0, 0, 1_048_575, 16_383)], 'all', 'local-user', [1]), true)
  assert.deepEqual(Object.keys(sheet.cells), ['1:0']); assert.deepEqual(sheet.mergedRanges, [])
})

test('clear format and clear all remove every intersecting merge and reject inaccessible parts outside the explicit request', () => {
  for (const mode of ['format', 'all']) {
    const sheet = createBlankSheet(); sheet.cells['0:0'] = createCell(0, 0, 'kept'); sheet.mergedRanges = [range(0, 0, 2, 2)]
    sheet.rangeProtectionRules = protect([range(2, 2)])
    const initial = copy(sheet)
    assert.throws(() => applySheetClearCommand(sheet, [range(0, 0)], mode, 'guest'), /保护/)
    assert.deepEqual(sheet, initial)
    sheet.rangeProtectionRules = []
    assert.equal(applySheetClearCommand(sheet, [range(0, 0)], mode), true)
    assert.deepEqual(sheet.mergedRanges, [])
    if (mode === 'format') { assert.equal(sheet.cells['0:0'].value, 'kept'); assert.deepEqual(sheet.cells['0:0'].style, defaultCellStyle()) }
  }
})

test('clear all clips validation and conditional resources to the exact unfiltered remainder', () => {
  const sheet = createBlankSheet(), source = range(0, 0, 4, 4), cut = range(1, 1, 3, 3)
  sheet.dataValidations = [validation([source])]; sheet.conditionalFormats = [conditional([source])]
  applySheetClearCommand(sheet, [cut], 'all', 'local-user', [2])
  const removed = new Set(points([range(1, 1, 1, 3), range(3, 1, 3, 3)])), expected = points([source]).filter(point => !removed.has(point))
  assert.deepEqual(points(sheet.dataValidations[0].ranges), expected)
  assert.deepEqual(points(sheet.conditionalFormats[0].ranges), expected)
  assert.equal(sheet.dataValidations[0].id, 'validation'); assert.deepEqual(sheet.conditionalFormats[0].metadata, { preserved: true })
})

test('clipping the initial rule anchor rebases relative formulas without moving absolute references or list sources', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  for (let row = 0; row < 4; row++) sheet.cells[`${row}:0`] = createCell(row, 0, row + 1)
  sheet.cells['0:3'] = createCell(0, 3, 0)
  sheet.dataValidations = [validation([range(0, 0, 3, 0)], { formula1: '=A1>$D$1' }),
    validation([range(0, 1, 3, 1)], { id: 'list', type: 'list', formula1: '=D1:D3' })]
  sheet.conditionalFormats = [conditional([range(0, 0, 3, 0)])]
  const before = validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 2, column: 0 }])
  applySheetClearCommand(sheet, [range(0, 0, 0, 1)], 'all')
  assert.equal(sheet.dataValidations[0].formula1, '=A2>$D$1')
  assert.equal(sheet.dataValidations[1].formula1, '=D1:D3')
  assert.equal(sheet.conditionalFormats[0].formula1, '=A2>0')
  assert.deepEqual(validateSheetDataValidationChanges(workbook, sheet.id, [{ row: 2, column: 0 }]), before)
})

test('conditional expressions and visual thresholds retain relative meaning when clipping reorders the first visible range', () => {
  const sheet = createBlankSheet(), scopes = [range(2, 2, 3, 3), range(0, 0, 2, 2)]
  sheet.conditionalFormats = [conditional(scopes), { ...conditional(scopes), id: 'visual', type: 'colorScale', config: {
    min: { type: 'formula', value: '=A1' }, max: { type: 'formula', value: '=$D$4' },
    thresholds: [{ value: { type: 'formula', value: '=A1+1' }, color: '#ff0000' }],
    icons: [{ value: { type: 'formula', value: '=A1+2' }, iconType: '3Arrows', iconId: '0' }],
  } }]
  applySheetClearCommand(sheet, [range(0, 0)], 'all')
  assert.equal(sheet.conditionalFormats[0].formula1, '=B1>0')
  const config = sheet.conditionalFormats[1].config
  assert.equal(config.min.value, '=B1'); assert.equal(config.max.value, '=$D$4')
  assert.equal(config.thresholds[0].value.value, '=B1+1'); assert.equal(config.icons[0].value.value, '=B1+2')
  assert.deepEqual(sheet.conditionalFormats[0].ranges[0], range(2, 2, 3, 3))
})

test('clear format removes hyperlinks and stale display formats while retaining formula, rich text and source extensions', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = { ...createCell(0, 0, 0.25), display: '25%', hyperlink: 'https://example.test', metadata: { source: 'kept' }, richText: [{ text: 'rich' }] }
  sheet.cells['0:0'].style.numberFormat = '0%'; sheet.cells['0:0'].style.font.bold = true; sheet.cells['0:0'].style.metadata = { sourceStyle: 'cleared' }
  const view = toExcelWorkbookView(workbook)
  applySheetViewClearCommand(view.sheets[0], [range(0, 0)], 'format')
  const result = applyExcelWorkbookView(workbook, view).sheets[0].cells['0:0']
  assert.equal(result.display, '0.25'); assert.equal(result.value, 0.25); assert.equal(result.hyperlink, null)
  assert.deepEqual(result.richText, sheet.cells['0:0'].richText); assert.deepEqual(result.metadata, { source: 'kept' }); assert.equal(result.style.metadata, undefined)
})

let vite, sdk
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  sdk = await vite.ssrLoadModule('/packages/office-studio/src/index.ts')
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
async function withStudio(workbook, run) {
  workbook.id = 'range-commands'; workbook.calculationMode ??= 'manual'
  const studio = new sdk.OfficeStudio({ workbook })
  try { await run(studio, sdk.FOffice.newAPI(studio)) } finally { studio.dispose() }
}
const undoCount = studio => studio.editor.history.units[studio.activeWorkbookId]?.undo.length ?? 0

test('Facade clear defaults to all and explicit upstream targets do not edit an unrelated active sheet', async () => {
  const workbook = createBlankWorkbook(), first = workbook.sheets[0], second = createBlankSheet('Second')
  first.cells['0:0'] = createCell(0, 0, 'active'); second.cells['0:0'] = createCell(0, 0, 'target'); second.cells['0:0'].style.font.bold = true
  workbook.sheets.push(second)
  await withStudio(workbook, (studio, api) => {
    const events = []; studio.events.addEvent('SheetValueChanged', event => events.push(copy(event)))
    studio.setSelection(workbook.id, first.id, coordinates(range(0, 0)))
    studio.executeCommandSync(SHEET_CLEAR_COMMAND_ALIASES.content, { unitId: workbook.id, subUnitId: second.id, ranges: [coordinates(range(0, 0))] })
    assert.equal(studio.getWorkbookModel(workbook.id).sheets[0].cells['0:0'].value, 'active')
    assert.equal(studio.getWorkbookModel(workbook.id).sheets[1].cells['0:0'].style.font.bold, true)
    api.getActiveWorkbook().getSheetBySheetId(second.id).getRange('A1').clear()
    assert.equal(studio.getWorkbookModel(workbook.id).sheets[1].cells['0:0'], undefined)
    assert.equal(undoCount(studio), 2); assert.equal(events[0].sheetId, second.id)
    assert.deepEqual(events[0].ranges, [{ ...coordinates(range(0, 0)), unitId: workbook.id, sheetId: second.id }])
  })
})

test('merge confirmation is a versioned request with no history until accepted; stale or canceled requests are atomic', async () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 'first'); sheet.cells['0:1'] = createCell(0, 1, 'second')
  await withStudio(workbook, (studio, api) => {
    const requests = []; studio.events.addEvent('SheetMergeConfirmationRequested', event => requests.push(copy(event)))
    studio.setSelections(workbook.id, sheet.id, [coordinates(range(0, 0, 0, 1))], { primary: { row: 0, column: 1 } })
    const initialSelection = copy(studio.getSelectionState())
    const input = { unitId: workbook.id, sheetId: sheet.id, ranges: [range(0, 0, 0, 1)] }
    assert.equal(studio.executeCommandSync(SHEET_MERGE_COMMANDS.merge, input), false)
    assert.equal(undoCount(studio), 0); assert.equal(requests.length, 1)
    assert.deepEqual(studio.getSelectionState(), initialSelection)
    assert.deepEqual(requests[0].input.ranges, input.ranges); assert.equal(requests[0].discardedContentCount, 1)
    const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === SHEET_MERGE_COMMANDS.merge) event.cancel = true })
    const before = copy(studio.getWorkbookModel(workbook.id))
    assert.throws(() => studio.executeCommandSync(SHEET_MERGE_COMMANDS.merge, { ...input, defaultMerge: true }), /canceled/)
    assert.deepEqual(studio.getWorkbookModel(workbook.id), before); assert.equal(undoCount(studio), 0)
    assert.deepEqual(studio.getSelectionState(), initialSelection); cancel.dispose()
    api.getActiveWorkbook().getActiveSheet().getRange('D1').setValue('new revision')
    const changed = copy(studio.getWorkbookModel(workbook.id))
    assert.throws(() => studio.executeCommandSync(requests[0].commandId, { ...requests[0].input, defaultMerge: true, expectedRevision: requests[0].baseRevision }), /已更改/)
    assert.deepEqual(studio.getWorkbookModel(workbook.id), changed); assert.equal(undoCount(studio), 1)
    assert.deepEqual(studio.getSelectionState(), initialSelection)
    assert.equal(studio.executeCommandSync(SHEET_MERGE_COMMANDS.merge, input), false)
    const request = requests[1]
    studio.setSelections(workbook.id, sheet.id, [coordinates(range(0, 3, 0, 4))], { primary: { row: 0, column: 4 } })
    const acceptedSelection = copy(studio.getSelectionState()), book = api.getActiveWorkbook()
    assert.equal(studio.executeCommandSync(request.commandId, { ...request.input, defaultMerge: true, expectedRevision: request.baseRevision }), true)
    assert.deepEqual(studio.getSelectionState(), acceptedSelection)
    assert.deepEqual(studio.getWorkbookModel(workbook.id).sheets[0].mergedRanges, input.ranges)
    assert.equal(undoCount(studio), 2)
    book.undo()
    assert.deepEqual(studio.getSelectionState(), acceptedSelection)
    assert.equal(studio.getWorkbookModel(workbook.id).sheets[0].cells['0:1'].value, 'second')
    book.redo()
    assert.deepEqual(studio.getSelectionState(), acceptedSelection)
    assert.deepEqual(studio.getWorkbookModel(workbook.id).sheets[0].mergedRanges, input.ranges)
  })
})

test('all merge modes retain actual primary coordinates and derive its merged shape through one content history', async () => {
  for (const [mode, mergedPrimary] of [
    ['all', range(1, 1, 3, 3)], ['horizontal', range(2, 1, 2, 3)], ['vertical', range(1, 2, 3, 2)],
  ]) {
    const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], target = range(1, 1, 3, 3)
    await withStudio(workbook, (studio, api) => {
      const book = api.getActiveWorkbook(), facade = book.getActiveSheet()
      studio.setSelections(workbook.id, sheet.id, [coordinates(range(8, 8)), coordinates(target)], { primary: { row: 2, column: 2 } })
      const before = copy(studio.getSelectionState())
      assert.equal(studio.executeCommandSync(SHEET_MERGE_COMMANDS.merge, { unitId: workbook.id, sheetId: sheet.id, ranges: [target], mode, defaultMerge: true }), true)
      assert.deepEqual(studio.getSelectionState(), before)
      const current = { ...coordinates(mergedPrimary), actualRow: 2, actualColumn: 2, rangeType: 0, isMerged: true, isMergedMainCell: false }
      assert.deepEqual(facade.getSelection().getCurrentCell(), current)
      assert.equal(undoCount(studio), 1)
      studio.setSelection(workbook.id, sheet.id, coordinates(range(10, 10)))
      book.undo()
      assert.deepEqual(studio.getSelectionState(), before)
      assert.equal(facade.getSelection().getCurrentCell().isMerged, false)
      book.redo()
      assert.deepEqual(studio.getSelectionState(), before)
      assert.deepEqual(facade.getSelection().getCurrentCell(), current)
    })
  }
})

test('unmerge relocates only the active primary and restores complete multi-selections through undo and redo', async () => {
  for (const activeRangeIndex of [0, 1]) {
    const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
    sheet.mergedRanges = [range(0, 0, 1, 1), range(3, 3, 4, 4)]
    await withStudio(workbook, (studio, api) => {
      const book = api.getActiveWorkbook(), facade = book.getActiveSheet(), targets = [range(0, 0, 2, 2), range(3, 3, 5, 5)]
      const primary = activeRangeIndex === 0 ? { row: 1, column: 1 } : { row: 4, column: 4 }
      studio.setSelections(workbook.id, sheet.id, targets.map(coordinates), { activeRangeIndex, primary })
      const before = copy(studio.getSelectionState()), notifications = []
      studio.events.addEvent('SelectionsChanged', event => notifications.push(copy(event.selection)))
      assert.equal(studio.executeCommandSync(SHEET_MERGE_COMMAND_ALIASES.unmerge, { unitId: workbook.id, subUnitId: sheet.id, ranges: targets }), true)
      const active = targets[activeRangeIndex], after = { ...before, primary: { row: active.startRow, column: active.startCol } }
      assert.deepEqual(studio.getSelectionState(), after)
      assert.deepEqual(notifications.at(-1), after)
      assert.equal(facade.getSelection().getCurrentCell().isMerged, false)
      assert.equal(undoCount(studio), 1)
      studio.setSelection(workbook.id, sheet.id, coordinates(range(10, 10)))
      book.undo()
      assert.deepEqual(studio.getSelectionState(), before)
      assert.equal(facade.getSelection().getCurrentCell().isMerged, true)
      studio.setSelection(workbook.id, sheet.id, coordinates(range(11, 11)))
      book.redo()
      assert.deepEqual(studio.getSelectionState(), after)
      assert.deepEqual(notifications.at(-1), after)
      assert.deepEqual(studio.getWorkbookModel(workbook.id).sheets[0].mergedRanges, [])
    })
  }
})

test('explicit unmerge outside saved selections leaves the active and inactive sheet selections unchanged', async () => {
  const workbook = createBlankWorkbook(), first = workbook.sheets[0], second = createBlankSheet('Second')
  second.mergedRanges = [range(0, 0, 1, 1)]; workbook.sheets.push(second)
  await withStudio(workbook, (studio, api) => {
    const book = api.getActiveWorkbook()
    studio.setSelections(workbook.id, second.id, [coordinates(range(4, 4, 5, 5))], { primary: { row: 5, column: 5 } })
    const inactiveSelection = copy(studio.getSelectionState(workbook.id, second.id))
    studio.setSelections(workbook.id, first.id, [coordinates(range(7, 7, 8, 8))], { primary: { row: 8, column: 8 } })
    const activeSelection = copy(studio.getSelectionState()), input = { unitId: workbook.id, subUnitId: second.id, ranges: [range(0, 0, 1, 1)] }
    assert.equal(studio.executeCommandSync(SHEET_MERGE_COMMAND_ALIASES.unmerge, input), true)
    assert.deepEqual(studio.getSelectionState(), activeSelection)
    assert.deepEqual(studio.getSelectionState(workbook.id, second.id), inactiveSelection)
    for (const transition of [() => book.undo(), () => book.redo()]) {
      transition()
      assert.deepEqual(studio.getSelectionState(), activeSelection)
      assert.deepEqual(studio.getSelectionState(workbook.id, second.id), inactiveSelection)
      assert.equal(studio.getWorkbookModel(workbook.id).activeSheetId, first.id)
    }
    assert.equal(studio.executeCommandSync(SHEET_MERGE_COMMAND_ALIASES.unmerge, input), false)
    assert.equal(undoCount(studio), 1)
    assert.deepEqual(studio.getSelectionState(), activeSelection)
  })
})

test('merge and clear enforce data validation in the candidate and never leave partial topology or content edits', async () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 1); sheet.cells['0:1'] = createCell(0, 1, 2); sheet.dataValidations = [validation([range(0, 1)], { formula1: '=B1>0' })]
  await withStudio(workbook, (studio, api) => {
    const book = api.getActiveWorkbook(), target = book.getActiveSheet().getRange('A1:B1'), initial = copy(studio.getWorkbookModel(workbook.id))
    assert.throws(() => target.merge(), /大于零/); assert.deepEqual(studio.getWorkbookModel(workbook.id), initial)
    assert.throws(() => target.clearContent(), /大于零/); assert.deepEqual(studio.getWorkbookModel(workbook.id), initial)
    assert.equal(undoCount(studio), 0)
    target.clear(); assert.equal(studio.getWorkbookModel(workbook.id).sheets[0].dataValidations.length, 0)
    assert.equal(undoCount(studio), 1)
  })
})

test('merge, breakApart, undo/redo and internal save/reopen preserve content and keep document histories isolated', async () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:1'] = { ...createCell(0, 1, false), metadata: { extension: 'preserved' }, richText: [{ text: 'false rich' }] }
  sheet.cells['0:1'].style.font.bold = true
  await withStudio(workbook, (studio, api) => {
    const book = api.getActiveWorkbook(), target = book.getActiveSheet().getRange('A1:B1')
    target.merge(); assert.equal(target.isMerged(), true); assert.equal(target.isPartOfMerge(), true)
    assert.equal(studio.getWorkbookModel(workbook.id).sheets[0].cells['0:0'].value, false)
    assert.throws(() => target.merge(), /重叠/)
    const count = undoCount(studio)
    studio.executeCommandSync(SHEET_MERGE_COMMANDS.merge, { unitId: workbook.id, sheetId: sheet.id, ranges: [range(0, 0, 0, 1)], defaultMerge: true, isForceMerge: true })
    assert.equal(undoCount(studio), count)
    const saved = book.save(), reopened = api.createWorkbook({ ...saved, id: 'range-reopened' })
    assert.equal(reopened.getActiveSheet().getRange('A1').getValue(), false)
    target.breakApart(); assert.equal(target.isPartOfMerge(), false)
    book.undo(); assert.equal(target.isMerged(), true)
    book.redo(); assert.equal(target.isPartOfMerge(), false)
    assert.equal(reopened.getActiveSheet().getRange('A1:B1').isMerged(), true)
    assert.deepEqual(studio.getWorkbookModel(reopened.getId()).sheets[0].cells['0:0'].metadata, { extension: 'preserved' })
  })
})

test('a no-op clear on an automatic workbook does not recalculate unrelated formulas or create a history entry', async () => {
  const workbook = createBlankWorkbook(); workbook.calculationMode = 'automatic'
  workbook.sheets[0].cells['0:0'] = { ...createCell(0, 0, null), formula: '=1+1', display: 'uncomputed', dataType: 'formula', cachedDataType: null }
  await withStudio(workbook, (studio, api) => {
    const book = api.getActiveWorkbook(); book.getActiveSheet().getRange('Z99').clearContent()
    assert.equal(undoCount(studio), 0); assert.equal(studio.getWorkbookModel(workbook.id).sheets[0].cells['0:0'].value, null)
    book.recalculate(); assert.equal(studio.getWorkbookModel(workbook.id).sheets[0].cells['0:0'].value, 2)
    assert.equal(undoCount(studio), 1)
  })
})
