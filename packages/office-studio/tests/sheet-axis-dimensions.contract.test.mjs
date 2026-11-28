import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { assignSheetAxisSpan, canonicalSheetAxisSpans, effectiveSheetRowSizes, sheetColumnWidth,
  sheetRowHeight, sheetRowIsAutoHeight, SparseAxisIndex } from '../src/domain/sheet-axis.ts'
import { transformSheetAxisDimensions, transformSheetAxisSpans } from '../src/domain/sheet-axis-transform.ts'
import { normalizeWorkbookAxisDimensions } from '../src/parser/workbookAxisDimensions.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { calculateSheetAxisAutoFit, SHEET_AUTO_FIT_MAX_ROW_HEIGHT, SHEET_AUTO_FIT_MAX_COLUMN_WIDTH } from '../src/features/sheet-workbench/sheet-auto-fit.ts'
import { applySheetAxisDimensionCommand, applySheetViewAxisDimensionCommand, refreshSheetViewAutoRowHeights, refreshWorkbookAutoRowHeights,
  SHEET_AXIS_DIMENSION_COMMANDS, sheetAxisDimensionActionForCommand, sheetAxisDimensionCommandInput,
  sheetAxisDimensionTargetRanges } from '../src/features/sheet-workbench/sheet-axis-dimensions.ts'
import { findVisibleSheetRow, findVisibleSheetColumn, isSheetRowVisible, isSheetColumnVisible } from '../src/features/sheet-workbench/sheet-row-visibility.ts'

const range = (startRow, startCol, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
const copy = value => structuredClone(value)
const options = { measure: text => [...text].length * 8, region: 'zh-CN' }
const manualRow = { axis: 'row', kind: 'set' }, manualColumn = { axis: 'column', kind: 'set' }
const autoRow = { axis: 'row', kind: 'auto' }, autoColumn = { axis: 'column', kind: 'auto' }
function view(extra = {}) {
  return { id: 'axes', name: 'Axes', axisUnit: 'cssPixel', rowCount: 100, columnCount: 26, cells: {},
    defaultRowHeight: 26, defaultColumnWidth: 112, rowHeights: {}, columnWidths: {}, ...extra }
}

test('axis spans retain later-declaration priority, zero sizes and gaps without mutating inputs', () => {
  const input = [{ start: 2, end: 8, value: 30 }, { start: 4, end: 5, value: 0 }, { start: 7, end: 10, value: 40 }]
  const initial = copy(input), output = canonicalSheetAxisSpans(input)
  assert.deepEqual(output, [{ start: 2, end: 3, value: 30 }, { start: 4, end: 5, value: 0 }, { start: 6, end: 6, value: 30 }, { start: 7, end: 10, value: 40 }])
  assert.deepEqual(assignSheetAxisSpan(output, 4, 6, 30), [{ start: 2, end: 6, value: 30 }, { start: 7, end: 10, value: 40 }])
  assert.deepEqual(assignSheetAxisSpan(output, 4, 6), [{ start: 2, end: 3, value: 30 }, { start: 7, end: 10, value: 40 }])
  assert.deepEqual(input, initial)
  for (const bad of [[-1, 0], [2, 1], [0, Number.MAX_SAFE_INTEGER]]) assert.throws(() => assignSheetAxisSpan([], ...bad, 20), RangeError)
})

test('prefix offsets, hit coordinates and visibility agree with independently expanded small axes', () => {
  const sizes = { 0: 15.5, 5: 7.25, 9: 0 }, hidden = [2, 11]
  const spans = [{ start: 1, end: 4, value: 0 }, { start: 5, end: 8, value: 20.125 }]
  const values = [15.5, 0, 0, 0, 0, 7.25, 20.125, 20.125, 20.125, 0, 10, 0, 10, 10]
  const axis = new SparseAxisIndex(values.length, 10, sizes, hidden, spans)
  let offset = 0
  for (let index = 0; index < values.length; index++) {
    assert.equal(axis.offset(index), offset); assert.equal(axis.size(index), values[index])
    if (values[index] > 0) assert.equal(axis.locate(offset + values[index] / 2), index)
    offset += values[index]
  }
  assert.equal(axis.offset(values.length), offset)
  assert.equal(axis.nextVisible(1, 1), 5); assert.equal(axis.nextVisible(4, -1), 0)
  assert.deepEqual(axis.visible(0, offset), [0, 5, 6, 7, 8, 10, 12, 13])
})

test('a million zero-size rows with sparse exceptions remain navigable without materializing a hidden-row set', () => {
  const sheet = view({ rowCount: 1_048_576, rowHeightSpans: [{ start: 0, end: 1_048_575, value: 0 }],
    rowHeights: { 200: 31.25, 900000: 44 }, rowAutoHeightSpans: [{ start: 400, end: 400, value: true }], rowAutoHeights: { 400: 50 }, hiddenRows: [900000] })
  assert.equal(isSheetRowVisible(sheet, 199), false); assert.equal(isSheetRowVisible(sheet, 200), true)
  assert.equal(findVisibleSheetRow(sheet, 0, 1, false), 200)
  assert.equal(findVisibleSheetRow(sheet, 201, 1, false), 400)
  assert.equal(findVisibleSheetRow(sheet, 899999, -1, false), 400)
  assert.equal(findVisibleSheetRow(sheet, 401, 1, false), undefined)
  const effective = effectiveSheetRowSizes(sheet), axis = new SparseAxisIndex(sheet.rowCount, 26, effective.sizes, [900000], effective.spans)
  assert.equal(axis.offset(sheet.rowCount), 31.25 + 50)
  assert.equal(Object.keys(sheet.rowHeights).length, 2); assert.equal(sheet.rowHeightSpans.length, 1)
})

test('zero-width spans, explicit overrides and first-column fallback use the same column navigation', () => {
  const sheet = view({ defaultColumnWidth: undefined, columnWidthSpans: [{ start: 0, end: 8, value: 0 }], columnWidths: { 3: 42 }, hiddenColumns: [10] })
  assert.equal(sheetColumnWidth(sheet, 0), 0); assert.equal(isSheetColumnVisible(sheet, 0), false)
  assert.equal(findVisibleSheetColumn(sheet, 0, 1, false), 3)
  assert.equal(findVisibleSheetColumn(sheet, 4, 1, false), 9)
  assert.equal(findVisibleSheetColumn(sheet, 10, -1, false), 9)
  assert.equal(findVisibleSheetColumn(view({ defaultColumnWidth: undefined }), 0, 1, false), 0)
})

test('parser geometry preserves marked pixels, logical counts, automatic caches and legacy unmarked values', () => {
  const normalized = normalizeWorkbookAxisDimensions({ axisUnit: 'cssPixel', rowCount: 400, columnCount: 40,
    rowHeights: [{ index: 3, size: 0 }, { index: 4, size: 31.123456789 }], columnWidths: { 1: 7 },
    rowAutoHeightSpans: [{ start: 7, end: 9, value: true }], rowAutoHeights: [{ index: 7, size: 42.875 }], hiddenRows: [9, 9, 2], hiddenColumns: [3] })
  assert.equal(normalized.rowHeights[4], 31.123456789); assert.equal(normalized.rowHeights[3], 0)
  assert.equal(normalized.columnWidths[1], 7); assert.equal(normalized.rowAutoHeights[7], 42.875)
  assert.equal(normalized.rowCount, 400); assert.equal(normalized.columnCount, 40)
  assert.deepEqual(normalized.hiddenRows, [2, 9])
  const legacy = normalizeWorkbookAxisDimensions({ columnWidths: [{ index: 1, size: 8.43 }] })
  assert.equal(legacy.columnWidths[1], 8.43); assert.equal(legacy.compatibility, 'legacyUnspecified')
  const exactDefaults = normalizeWorkbookAxisDimensions({ axisUnit: 'cssPixel', sourceAxisDimensions: { projectAxisExtension: { restored: true } } })
  assert.equal(exactDefaults.defaultColumnWidth, undefined); assert.equal(exactDefaults.defaultRowHeight, undefined)
  for (const input of [{ axisUnit: 'point' }, { rowCount: 0 }, { columnCount: 16_385 }, { rowHeights: 12 }, { rowHeights: [{ index: 2, size: 1 }, { index: 2, size: 3 }] },
    { rowHeights: { bad: 30 } }, { columnWidths: { 0: -1 } }, { hiddenRows: [1.5] }, { rowAutoHeightSpans: [{ start: 0, end: 0, value: 1 }] }]) assert.throws(() => normalizeWorkbookAxisDimensions(input))
})

test('header delta targets selected whole axes only when the anchor belongs to them', () => {
  const sheet = view(), rows = [range(2, 0, 3, 25), range(8, 0, 9, 25)], cells = range(2, 1, 5, 4)
  assert.deepEqual(sheetAxisDimensionTargetRanges(sheet, 'row', rows, 3), rows)
  assert.deepEqual(sheetAxisDimensionTargetRanges(sheet, 'row', [...rows, cells], 6), [range(6, 0, 6, 25)])
  assert.deepEqual(sheetAxisDimensionTargetRanges(sheet, 'column', [range(0, 2, 99, 3)], 2), [range(0, 2, 99, 3)])
  assert.throws(() => sheetAxisDimensionTargetRanges(sheet, 'row', rows, 100), RangeError)
  assert.deepEqual(sheetAxisDimensionCommandInput({ subUnitId: 'axes', ranges: [{ startRow: 2, endRow: 3, startColumn: 0, endColumn: 25 }], value: 0 }).ranges, [range(2, 0, 3, 25)])
  assert.deepEqual(sheetAxisDimensionActionForCommand('sheet.command.set-col-width'), manualColumn)
})

test('manual size changes switch row modes atomically, preserve zero and skip identical imported sparse overrides', () => {
  const sheet = view({ rowHeights: { 4: 40 }, rowAutoHeightSpans: [{ start: 5, end: 5, value: true }], rowAutoHeights: { 5: 80 } })
  assert.equal(applySheetViewAxisDimensionCommand(sheet, [range(4, 0, 4, 25)], manualRow, { value: 40 }, 'local-user', options), false)
  assert.deepEqual(sheet.rowHeights, { 4: 40 }); assert.equal(sheet.rowHeightSpans, undefined)
  assert.equal(applySheetViewAxisDimensionCommand(sheet, [range(5, 0, 5, 25)], manualRow, { value: 0 }, 'local-user', options), true)
  assert.equal(sheetRowHeight(sheet, 5), 0); assert.equal(sheetRowIsAutoHeight(sheet, 5), false)
  assert.equal(sheet.rowAutoHeights[5], undefined)
  const before = copy(sheet)
  assert.throws(() => applySheetViewAxisDimensionCommand(sheet, [range(5, 0)], manualRow, { value: -1 }, 'local-user', options), RangeError)
  assert.deepEqual(sheet, before)
  const model = createBlankSheet(); model.rowHeights = { 2: 32.5 }
  assert.equal(applySheetAxisDimensionCommand(model, [range(2, 0, 2, model.columnCount - 1)], manualRow, { value: 32.5 }, 'local-user', options), false)
  assert.deepEqual(model.rowHeights, { 2: 32.5 })
})

test('whole-axis manual sizes use compact spans; multi-axis dragging applies the anchor destination to every selected axis', () => {
  const sheet = view({ rowCount: 1_048_576 })
  applySheetViewAxisDimensionCommand(sheet, [range(0, 0, 1_048_575, 25)], manualRow, { value: 32.125 }, 'local-user', options)
  assert.deepEqual(sheet.rowHeightSpans, [{ start: 0, end: 1_048_575, value: 32.125 }]); assert.equal(Object.keys(sheet.rowHeights).length, 0)
  applySheetViewAxisDimensionCommand(sheet, [range(2, 0, 4, 25), range(8, 0, 8, 25)], { axis: 'row', kind: 'delta' }, { anchorRow: 3, deltaY: 10 }, 'local-user', options)
  for (const row of [2, 3, 4, 8]) assert.equal(sheetRowHeight(sheet, row), 42.125)
  assert.equal(sheetRowHeight(sheet, 5), 32.125)
})

test('auto sizing uses displayed text, explicit line breaks, zero-width targets and fixed-version limits', () => {
  const sheet = view({ columnWidths: { 1: 0, 3: 77 }, cells: {
    B1: { value: 'ABCDEFGHIJ', fontSize: 11 }, C2: { value: Array(400).fill('一').join('\n'), fontSize: 11, wrap: true },
  } })
  assert.equal(calculateSheetAxisAutoFit(sheet, 'column', 1, options), 92)
  assert.equal(calculateSheetAxisAutoFit(sheet, 'column', 3, options), 77)
  assert.equal(calculateSheetAxisAutoFit(sheet, 'row', 1, options), 2000)
  assert.equal(SHEET_AUTO_FIT_MAX_ROW_HEIGHT, 2000); assert.equal(SHEET_AUTO_FIT_MAX_COLUMN_WIDTH, 2000)
  applySheetViewAxisDimensionCommand(sheet, [range(0, 1, 99, 1)], autoColumn, {}, 'local-user', options)
  assert.equal(sheetColumnWidth(sheet, 1), 92)
  applySheetViewAxisDimensionCommand(sheet, [range(1, 0, 1, 25)], autoRow, {}, 'local-user', options)
  assert.equal(sheetRowIsAutoHeight(sheet, 1), true); assert.equal(sheetRowHeight(sheet, 1), 2000)
})

test('automatic row heights track edits and width changes while forced heights stay manual', () => {
  const sheet = view({ columnWidths: { 1: 44 }, cells: { B2: { value: 'abcdefghij', fontSize: 11, wrap: true } } })
  refreshSheetViewAutoRowHeights(sheet, undefined, options)
  const tall = sheetRowHeight(sheet, 1); assert.ok(tall > 26)
  const previous = copy(sheet); sheet.columnWidths[1] = 200
  assert.equal(refreshSheetViewAutoRowHeights(sheet, previous, options), true)
  assert.ok(sheetRowHeight(sheet, 1) < tall)
  applySheetViewAxisDimensionCommand(sheet, [range(1, 0)], manualRow, { value: 30.5 }, 'local-user', options)
  const manual = copy(sheet); sheet.cells.B2.value = Array(100).fill('line').join('\n')
  refreshSheetViewAutoRowHeights(sheet, manual, options)
  assert.equal(sheetRowHeight(sheet, 1), 30.5); assert.equal(sheetRowIsAutoHeight(sheet, 1), false)
})

test('insert/delete transforms keep dimension exceptions and automatic caches attached to their axes', () => {
  const sheet = view({ rowHeights: { 1: 40, 7: 0 }, rowHeightSpans: [{ start: 2, end: 5, value: 30 }],
    rowAutoHeightSpans: [{ start: 4, end: 6, value: true }], rowAutoHeights: { 4: 50, 6: 70 } })
  const original = copy(sheet), inserted = transformSheetAxisDimensions(sheet, { axis: 'row', kind: 'insert', start: 3, count: 2, inheritedSize: 30 })
  assert.deepEqual(inserted.rowHeights, { 1: 40, 9: 0 })
  assert.deepEqual(inserted.rowHeightSpans, [{ start: 2, end: 7, value: 30 }])
  assert.deepEqual(inserted.rowAutoHeights, { 6: 50, 8: 70 })
  assert.deepEqual(inserted.rowAutoHeightSpans, [{ start: 3, end: 4, value: false }, { start: 6, end: 8, value: true }])
  const deleted = transformSheetAxisDimensions({ ...sheet, ...inserted }, { axis: 'row', kind: 'delete', start: 3, count: 2 })
  assert.deepEqual(deleted.rowHeights, original.rowHeights); assert.deepEqual(deleted.rowHeightSpans, original.rowHeightSpans)
  assert.deepEqual(deleted.rowAutoHeightSpans, original.rowAutoHeightSpans); assert.deepEqual(deleted.rowAutoHeights, original.rowAutoHeights)
  assert.deepEqual(sheet, original)
  assert.deepEqual(transformSheetAxisSpans([{ start: 0, end: 1_048_575, value: 32 }], { axis: 'row', kind: 'delete', start: 100, count: 10 }), [{ start: 0, end: 1_048_565, value: 32 }])
})

test('view adapters and internal snapshots retain dimension fields, zero overrides and explicit modes', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  Object.assign(sheet, { rowHeights: { 7: 0 }, rowHeightSpans: [{ start: 8, end: 50, value: 32.125 }], columnWidthSpans: [{ start: 2, end: 6, value: 0 }],
    rowAutoHeightSpans: [{ start: 60, end: 60, value: true }], rowAutoHeights: { 60: 48.5 }, defaultRowAutoHeight: false })
  const converted = toExcelWorkbookView(workbook), restored = applyExcelWorkbookView(workbook, converted).sheets[0]
  for (const key of ['rowHeights', 'rowHeightSpans', 'columnWidthSpans', 'rowAutoHeightSpans', 'rowAutoHeights', 'defaultRowAutoHeight']) assert.deepEqual(restored[key], sheet[key])
  assert.equal(converted.sheets[0].axisUnit, 'cssPixel')
})

let vite, sdk, useSheetAxisDimensions
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  sdk = await vite.ssrLoadModule('/packages/office-studio/src/index.ts')
  ;({ useSheetAxisDimensions } = await vite.ssrLoadModule('/packages/office-studio/src/components/spreadsheet/use-sheet-axis-dimensions.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
async function withStudio(workbook, run) {
  workbook.id = 'axis-command-contract'; workbook.calculationMode = 'manual'
  const studio = new sdk.OfficeStudio({ workbook, sheetTextMeasure: options.measure })
  try { await run(studio, sdk.FOffice.newAPI(studio)) } finally { studio.dispose() }
}
const undoCount = studio => studio.editor.history.units[studio.activeWorkbookId]?.undo.length ?? 0

test('Facade forced sizes and shared command aliases target inactive sheets with one history and dimension event', async () => {
  const workbook = createBlankWorkbook(), second = createBlankSheet('Second'); workbook.sheets.push(second)
  await withStudio(workbook, (studio, api) => {
    const events = []; studio.events.addEvent('SheetDimensionsChanged', event => events.push(copy(event)))
    const target = api.getActiveWorkbook().getSheetBySheetId(second.id)
    target.setRowHeightsForced(2, 3, 42.125).setColumnWidths(4, 2, 0)
    assert.equal(target.getRowHeight(2), 42.125); assert.equal(target.getRowIsAutoHeight(2), false)
    assert.equal(target.getColumnWidth(4), 0); assert.equal(undoCount(studio), 2)
    assert.equal(events[0].sheetId, second.id); assert.equal(events[0].axis, 'row'); assert.equal(events[0].automatic, false)
    assert.equal(studio.getWorkbookModel(workbook.id).sheets[0].rowHeightSpans?.length ?? 0, 0)
    target.setRowHeightsForced(2, 3, 42.125); assert.equal(undoCount(studio), 2)
    studio.executeCommandSync('sheet.command.set-col-width', { unitId: workbook.id, subUnitId: second.id, ranges: [range(0, 4, second.rowCount - 1, 5)], value: 18.75 })
    assert.equal(target.getColumnWidth(5), 18.75); assert.equal(undoCount(studio), 3)
    api.getActiveWorkbook().undo(); assert.equal(target.getColumnWidth(4), 0)
    api.getActiveWorkbook().redo(); assert.equal(target.getColumnWidth(4), 18.75)
    const saved = api.getActiveWorkbook().save(), reopened = api.createWorkbook({ ...saved, id: 'axis-reopened' })
    assert.equal(reopened.getSheetBySheetId(second.id).getRowHeight(3), 42.125)
    assert.equal(reopened.getSheetBySheetId(second.id).getColumnWidth(4), 18.75)
  })
})

test('permission refusal, command cancellation, stale revisions and invalid sizes never publish partial history', async () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 'keep')
  await withStudio(workbook, (studio, api) => {
    const input = { unitId: workbook.id, sheetId: sheet.id, ranges: [range(0, 0, 0, sheet.columnCount - 1)], value: 0 }
    const initial = copy(studio.getWorkbookModel(workbook.id)), revision = studio.getWorkbookRevision(workbook.id)
    const cancellation = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === SHEET_AXIS_DIMENSION_COMMANDS.rowHeight) event.cancel = true })
    assert.throws(() => studio.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.rowHeight, input), /canceled/)
    assert.deepEqual(studio.getWorkbookModel(workbook.id), initial); assert.equal(undoCount(studio), 0); cancellation.dispose()
    api.getActiveWorkbook().getActiveSheet().setColumnWidth(0, 20)
    const changed = copy(studio.getWorkbookModel(workbook.id))
    assert.throws(() => studio.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.rowHeight, { ...input, expectedRevision: revision }), /已更改/)
    assert.throws(() => studio.executeCommandSync(SHEET_AXIS_DIMENSION_COMMANDS.rowHeight, { ...input, value: -1 }), /非负/)
    assert.deepEqual(studio.getWorkbookModel(workbook.id), changed); assert.equal(undoCount(studio), 1)
  })
  sheet.worksheetProtection = { id: 'locked', permissionId: 'locked-permission', unitType: 'worksheet', ownerUserId: 'other', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
  await withStudio(workbook, (studio, api) => {
    const initial = copy(studio.getWorkbookModel(workbook.id))
    assert.throws(() => api.getActiveWorkbook().getActiveSheet().setRowHeightsForced(0, 1, 0), /保护|拒绝|权限|不可/)
    assert.deepEqual(studio.getWorkbookModel(workbook.id), initial); assert.equal(undoCount(studio), 0)
  })
})

test('list-only additions, range moves, mode changes and removal refresh sparse automatic rows while manual heights stay fixed', () => {
  const sheet = view({ columnWidths: { 0: 60 }, rowHeights: { 1: 40 }, cells: {
    A1: { value: 'A,B,C', rawValue: 'A,B,C', dataType: 'text' }, A2: { value: 'A,B,C', rawValue: 'A,B,C', dataType: 'text' },
    A3: { value: 'A,B,C', rawValue: 'A,B,C', dataType: 'text' },
  } })
  refreshSheetViewAutoRowHeights(sheet, undefined, options)
  let previous = copy(sheet)
  sheet.dataValidations = [{ id: 'tags', type: 'listMultiple', formula1: '"A,B,C"', ranges: [range(0, 0)], renderMode: 'custom' }]
  assert.equal(refreshSheetViewAutoRowHeights(sheet, previous, options), true)
  assert.equal(sheetRowHeight(sheet, 0), 78); assert.equal(sheetRowHeight(sheet, 1), 40); assert.equal(sheetRowHeight(sheet, 2), 26)
  previous = copy(sheet); sheet.dataValidations[0].ranges = [range(1, 0, 2, 0)]
  assert.equal(refreshSheetViewAutoRowHeights(sheet, previous, options), true)
  assert.equal(sheetRowHeight(sheet, 0), 26); assert.equal(sheetRowHeight(sheet, 1), 40); assert.equal(sheetRowHeight(sheet, 2), 78)
  previous = copy(sheet); sheet.dataValidations[0].showDropDown = false
  assert.equal(refreshSheetViewAutoRowHeights(sheet, previous, options), true); assert.equal(sheetRowHeight(sheet, 2), 54)
  previous = copy(sheet); sheet.dataValidations[0].renderMode = 'text'
  assert.equal(refreshSheetViewAutoRowHeights(sheet, previous, options), true); assert.equal(sheetRowHeight(sheet, 2), 26)
  previous = copy(sheet); sheet.dataValidations[0].renderMode = 'custom'
  refreshSheetViewAutoRowHeights(sheet, previous, options); previous = copy(sheet); sheet.dataValidations = []
  assert.equal(refreshSheetViewAutoRowHeights(sheet, previous, options), true)
  assert.equal(sheetRowHeight(sheet, 2), 26); assert.equal(sheetRowHeight(sheet, 1), 40)
})

test.each(['authored', 'reference', 'cleared', 'theme'])('%s style changes invalidate the measured height without rewriting the scalar or projected font cache', source => {
  const cell = { value: 'A', rawValue: 'A', dataType: 'text', fontSize: 11, authoredStyle: {} }
  const sheet = view({ cells: { A1: cell }, styles: { large: { font: { size: 24 } } },
    dataValidations: [{ id: 'tag', type: 'list', formula1: '"A"', ranges: [range(0, 0)] }] })
  if (source === 'cleared') { sheet.workbookDefaultStyle = { font: { size: 30 } }; cell.authoredStyleReference = 'large' }
  refreshSheetViewAutoRowHeights(sheet, undefined, options)
  const previous = copy(sheet), oldHeight = sheetRowHeight(sheet, 0)
  if (source === 'authored') cell.authoredStyle = { font: { size: 24 } }
  if (source === 'reference') cell.authoredStyleReference = 'large'
  if (source === 'cleared') cell.clearedStyleProperties = ['font.size']
  if (source === 'theme') cell.themeStyle = 'large'
  assert.equal(refreshSheetViewAutoRowHeights(sheet, previous, options), true)
  assert.equal(sheetRowHeight(sheet, 0), source === 'cleared' ? 54 : 46); assert.notEqual(sheetRowHeight(sheet, 0), oldHeight)
  assert.equal(cell.fontSize, 11); assert.equal(cell.rawValue, 'A'); assert.equal(cell.value, 'A')
})

test('a workbook catalog change remeasures a shared inherited-style sheet and preserves the previous history object', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.defaultStyle = 'base'; workbook.styles = { base: { font: { size: 11 } } }
  sheet.cells['0:0'] = createCell(0, 0, 'A'); sheet.cells['0:0'].authoredStyle = {}
  sheet.dataValidations = [{ id: 'tag', type: 'list', formula1: '"A"', ranges: [range(0, 0)] }]
  refreshWorkbookAutoRowHeights(workbook, undefined, options)
  const before = copy(workbook), next = { ...workbook, styles: { base: { font: { size: 24 } } } }
  assert.equal(sheetRowHeight(sheet, 0), 29)
  refreshWorkbookAutoRowHeights(next, workbook, options)
  assert.equal(sheetRowHeight(next.sheets[0], 0), 46); assert.notEqual(next.sheets[0], sheet)
  assert.deepEqual(workbook, before)
})

test('table-filter visibility changes the inherited range-theme band height in a shared worksheet', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['2:0'] = createCell(2, 0, 'A'); sheet.cells['2:0'].authoredStyle = {}
  sheet.rangeThemeRules = [{ id: 'theme', themeName: 'bands', range: range(0, 0, 4, 0) }]
  workbook.rangeThemes = { bands: { wholeStyle: { font: { size: 10 } }, firstRowStyle: { font: { size: 30 } },
    secondRowStyle: { font: { size: 10 } } } }
  workbook.tables = [{ id: 'table', sheetId: sheet.id, name: 'Table1', range: range(0, 0, 4, 0), showHeaderRow: true,
    columns: [{ id: 'column', name: 'Value' }], autoFilter: { range: range(0, 0, 4, 0), columns: [], filteredRows: [] } }]
  refreshWorkbookAutoRowHeights(workbook, undefined, options); assert.equal(sheetRowHeight(sheet, 2), 26)
  const before = copy(workbook), next = { ...workbook, tables: [{ ...workbook.tables[0],
    autoFilter: { ...workbook.tables[0].autoFilter, filteredRows: [1] } }] }
  refreshWorkbookAutoRowHeights(next, workbook, options)
  assert.equal(sheetRowHeight(next.sheets[0], 2), 54); assert.notEqual(next.sheets[0], sheet); assert.deepEqual(workbook, before)
})

test('an unchanged shared sheet avoids projecting or measuring its sparse cells on an unrelated workbook change', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, 'A')
  const next = { ...workbook, name: 'Renamed workbook', modified: true }
  Object.defineProperty(sheet.cells['0:0'], 'value', { enumerable: true, get: () => { throw new Error('Unchanged cells must not be scanned') } })
  refreshWorkbookAutoRowHeights(next, workbook, { ...options, measure: () => { throw new Error('Unchanged cells must not be measured') } })
  assert.equal(next.sheets, workbook.sheets); assert.equal(next.sheets[0], sheet)
})

test('a referenced option label change refreshes automatic heights in a different shared target sheet', () => {
  const workbook = createBlankWorkbook(), target = workbook.sheets[0], source = createBlankSheet('Source')
  target.cells['0:0'] = createCell(0, 0, 'A')
  target.dataValidations = [{ id: 'source-list', type: 'list', formula1: '=Source!$A$1', ranges: [range(0, 0)] }]
  source.cells['0:0'] = createCell(0, 0, 'A'); workbook.sheets.push(source)
  refreshWorkbookAutoRowHeights(workbook, undefined, options); assert.equal(sheetRowHeight(target, 0), 29)
  const before = copy(workbook), next = { ...workbook, sheets: [target, { ...source, cells: { '0:0': createCell(0, 0, 'A\n') } }] }
  refreshWorkbookAutoRowHeights(next, workbook, options)
  assert.equal(sheetRowHeight(next.sheets[0], 0), 44); assert.notEqual(next.sheets[0], target); assert.deepEqual(workbook, before)
})

test('a real validation command commits tag heights in the same undo entry and internal reopen retains automatic/manual behavior', async () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.defaultColumnWidth = 60; sheet.cells['0:0'] = createCell(0, 0, 'A,B,C')
  await withStudio(workbook, (studio, api) => {
    const before = copy(studio.getWorkbookModel(workbook.id)), target = api.getActiveWorkbook().getActiveSheet()
    assert.equal(studio.executeCommandSync('sheet.command.set-data-validation', { unitId: workbook.id, sheetId: sheet.id,
      rule: { id: 'tags', type: 'listMultiple', formula1: '"A,B,C"', ranges: [range(0, 0)], renderMode: 'custom' } }), true)
    assert.equal(target.getRowHeight(0), 78); assert.equal(target.getRowIsAutoHeight(0), true); assert.equal(undoCount(studio), 1)
    api.getActiveWorkbook().undo(); assert.deepEqual(studio.getWorkbookModel(workbook.id), before)
    api.getActiveWorkbook().redo(); assert.equal(target.getRowHeight(0), 78)
    const reopened = api.createWorkbook({ ...api.getActiveWorkbook().save(), id: 'validation-axis-reopened' }), restored = reopened.getActiveSheet()
    assert.equal(restored.getRowHeight(0), 78); assert.equal(restored.getRowIsAutoHeight(0), true)
    assert.equal(reopened.save().sheets[0].cells['0:0'].value, 'A,B,C')
    restored.setRowHeightsForced(0, 1, 38); restored.getRange('A1').setValue('A')
    assert.equal(restored.getRowHeight(0), 38); assert.equal(restored.getRowIsAutoHeight(0), false)
    reopened.undo(); assert.equal(reopened.save().sheets[0].cells['0:0'].value, 'A,B,C'); assert.equal(restored.getRowHeight(0), 38)
    reopened.undo(); assert.equal(restored.getRowHeight(0), 78); assert.equal(restored.getRowIsAutoHeight(0), true)
    reopened.redo(); reopened.redo(); assert.equal(restored.getRowHeight(0), 38); assert.equal(reopened.save().sheets[0].cells['0:0'].value, 'A')
  })
})

test('Facade auto-column width measures inherited workbook font and rotation, then retains them through undo, redo and internal reopen', async () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.defaultStyle = { font: { size: 24 }, rotation: 90 }
  sheet.cells['0:0'] = createCell(0, 0, 'A'); sheet.cells['0:0'].authoredStyle = {}
  sheet.dataValidations = [{ id: 'inherited-tag', type: 'list', formula1: '"A"', ranges: [range(0, 0)] }]
  await withStudio(workbook, (studio, api) => {
    const book = api.getActiveWorkbook(), target = book.getActiveSheet(), before = copy(studio.getWorkbookModel(workbook.id))
    target.setColumnAutoWidth(0)
    // 24pt -> 32px text height + 6px tag padding, rotated 90 degrees, plus 8px cell padding and 22px arrow.
    assert.equal(target.getColumnWidth(0), 68); assert.equal(undoCount(studio), 1)
    assert.deepEqual(studio.getWorkbookModel(workbook.id).defaultStyle, before.defaultStyle)
    assert.deepEqual(studio.getWorkbookModel(workbook.id).sheets[0].cells['0:0'], before.sheets[0].cells['0:0'])
    book.undo(); assert.deepEqual(studio.getWorkbookModel(workbook.id), before)
    book.redo(); assert.equal(target.getColumnWidth(0), 68)
    const saved = book.save(), reopened = api.createWorkbook({ ...saved, id: 'inherited-auto-width-reopened' })
    assert.equal(reopened.getActiveSheet().getColumnWidth(0), 68); assert.deepEqual(reopened.save().defaultStyle, saved.defaultStyle)
    reopened.getActiveSheet().setColumnAutoWidth(0); assert.equal(reopened.getActiveSheet().getColumnWidth(0), 68)
  })
})

// These composable drafts use a real snapshot/command path and dispose every synchronous watcher.
function standaloneDimensionsHost({ measure, configure = () => {} } = {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'standalone-axis-dimensions'; model.activeSheetId = sheet.id = 'axes'; model.calculationMode = 'manual'
  sheet.rowCount = 20; sheet.columnCount = 8; sheet.cells['0:0'] = createCell(0, 0, 'A'); configure(model, sheet)
  const studio = measure ? new sdk.OfficeStudio({ workbook: model, sheetTextMeasure: measure }) : undefined
  const converted = toExcelWorkbookView(studio?.getWorkbookModel(model.id) ?? model); delete converted.id
  const state = shallowReactive({ studio, model: converted, sheetId: sheet.id, userId: 'local-user', region: 'en-US', active: true })
  const calls = { prepared: 0, applied: 0, changed: [], focused: 0 }, scope = effectScope()
  const context = { studio: () => state.studio, workbook: () => state.model, sheetId: () => state.sheetId,
    ranges: () => [range(0, 0)], userId: () => state.userId, region: () => state.region, active: () => state.active, finishEdit: () => true,
    prepareView: next => { calls.prepared++; return next }, applyView: next => { state.model = next; calls.applied++ },
    changed: options => calls.changed.push(options), focus: () => { calls.focused++ } }
  const ui = scope.run(() => useSheetAxisDimensions(context))
  return { state, context, calls, ui, dispose: () => { scope.stop(); studio?.dispose() } }
}
const changeDimensionContext = (host, property) => { host.state[property] = property === 'userId' ? 'peer' : 'de-DE' }
const setStandaloneHeight = host => host.ui.runAxisDimensionCommand(SHEET_AXIS_DIMENSION_COMMANDS.rowHeight, { value: 38 })

test('standalone auto-column width uses the inherited style and formatted source candidate without a workbook ID', () => {
  const h = standaloneDimensionsHost({ measure: options.measure, configure: (workbook, sheet) => {
    workbook.defaultStyle = { font: { size: 24 }, rotation: 90 }; sheet.cells['0:0'].authoredStyle = {}
    const source = createBlankSheet('Source'); source.cells['0:0'] = createCell(0, 0, 'A'); workbook.sheets.push(source)
    sheet.dataValidations = [{ id: 'source-tag', type: 'list', formula1: '=Source!$A$1', ranges: [range(0, 0)] }]
  } })
  try {
    const before = copy(h.state.model)
    assert.equal(h.ui.runAxisDimensionCommand(SHEET_AXIS_DIMENSION_COMMANDS.columnAutoWidth), true)
    assert.equal(sheetColumnWidth(h.state.model.sheets[0], 0), 68); assert.equal(h.state.model.id, undefined)
    assert.equal(h.calls.prepared, 1); assert.equal(h.calls.applied, 1); assert.equal(h.calls.changed.length, 1)
    assert.equal(h.calls.changed[0].current(), true); assert.deepEqual(h.state.model.sheets[1], before.sheets[1])
    assert.deepEqual(h.state.model.sheets[0].cells, before.sheets[0].cells)
  } finally { h.dispose() }
})

test.each(['userId', 'region'])('%s changes during standalone source getters abort before preparation or publication', property => {
  const h = standaloneDimensionsHost()
  try {
    const original = h.state.model, cell = original.sheets[0].cells.A1, raw = cell.rawValue
    Object.defineProperty(cell, 'rawValue', { enumerable: true, get: () => { changeDimensionContext(h, property); return raw } })
    assert.equal(setStandaloneHeight(h), false); assert.equal(h.state.model, original)
    assert.equal(sheetRowHeight(original.sheets[0], 0), 26)
    assert.equal(h.calls.prepared, 0); assert.equal(h.calls.applied, 0); assert.deepEqual(h.calls.changed, [])
  } finally { h.dispose() }
})

test.each(['userId', 'region'])('%s changes during standalone tag measurement discard the complete candidate before preparing it', property => {
  let h, measurements = 0
  h = standaloneDimensionsHost({ measure: text => {
    if (h) { measurements++; changeDimensionContext(h, property) }
    return options.measure(text)
  }, configure: (_workbook, sheet) => {
    sheet.dataValidations = [{ id: 'tag', type: 'list', formula1: '"A"', ranges: [range(0, 0)] }]
  } })
  try {
    const original = h.state.model, before = copy(original)
    assert.equal(h.ui.runAxisDimensionCommand(SHEET_AXIS_DIMENSION_COMMANDS.columnAutoWidth), false); assert.ok(measurements > 0)
    assert.equal(h.state.model, original); assert.deepEqual(h.state.model, before)
    assert.equal(h.calls.prepared, 0); assert.equal(h.calls.applied, 0); assert.deepEqual(h.calls.changed, [])
  } finally { h.dispose() }
})

test.each(['userId', 'region'])('%s changes while preparing standalone dimensions reject publication and preserve the source', property => {
  const h = standaloneDimensionsHost()
  try {
    const before = copy(h.state.model), prepare = h.context.prepareView
    h.context.prepareView = next => { const result = prepare(next); changeDimensionContext(h, property); return result }
    assert.equal(setStandaloneHeight(h), false); assert.deepEqual(h.state.model, before)
    assert.equal(h.calls.prepared, 1); assert.equal(h.calls.applied, 0); assert.deepEqual(h.calls.changed, [])
  } finally { h.dispose() }
})

test.each(['userId', 'region'])('%s changes during standalone publication suppress later notifications without losing the applied size', property => {
  const h = standaloneDimensionsHost()
  try {
    const publish = h.context.applyView
    h.context.applyView = next => { publish(next); changeDimensionContext(h, property) }
    assert.equal(setStandaloneHeight(h), true); assert.equal(sheetRowHeight(h.state.model.sheets[0], 0), 38)
    assert.equal(h.calls.applied, 1); assert.deepEqual(h.calls.changed, [])
  } finally { h.dispose() }
})

test.each(['userId', 'region'])('%s changes invalidate saved publication guards and scheduled focus even after restoring the original setting', async property => {
  const h = standaloneDimensionsHost()
  try {
    const original = h.state[property]
    assert.equal(setStandaloneHeight(h), true); assert.equal(h.calls.changed[0].current(), true)
    h.ui.afterAxisSizeClose(); changeDimensionContext(h, property); h.state[property] = original
    await nextTick(); assert.equal(h.calls.changed[0].current(), false); assert.equal(h.calls.focused, 0)
    assert.equal(h.calls.applied, 1); assert.equal(sheetRowHeight(h.state.model.sheets[0], 0), 38)
  } finally { h.dispose() }
})

test.each(['userId', 'region'])('%s changes cancel captured dimension dialogs and require a new dialog before a forced size can publish', async property => {
  const h = standaloneDimensionsHost()
  try {
    const original = h.state[property], before = copy(h.state.model)
    h.ui.openAxisSizeDialog('row'); assert.ok(h.ui.pending.value)
    changeDimensionContext(h, property); assert.equal(h.ui.pending.value, null)
    h.state[property] = original; h.ui.confirmAxisSize(40)
    assert.deepEqual(h.state.model, before); assert.equal(h.calls.applied, 0)
    h.ui.openAxisSizeDialog('row'); h.ui.confirmAxisSize(40); h.ui.afterAxisSizeClose(); await nextTick()
    assert.equal(sheetRowHeight(h.state.model.sheets[0], 0), 40); assert.equal(h.ui.pending.value, null)
    assert.equal(h.calls.applied, 1); assert.equal(h.calls.changed.length, 1); assert.equal(h.calls.focused, 1)
  } finally { h.dispose() }
})
