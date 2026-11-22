import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { assignSheetAxisSpan, canonicalSheetAxisSpans, effectiveSheetRowSizes, sheetColumnWidth,
  sheetRowHeight, sheetRowIsAutoHeight, SparseAxisIndex } from '../src/domain/sheet-axis.ts'
import { transformSheetAxisDimensions, transformSheetAxisSpans } from '../src/domain/sheet-axis-transform.ts'
import { normalizeWorkbookAxisDimensions } from '../src/parser/workbookAxisDimensions.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { calculateSheetAxisAutoFit, SHEET_AUTO_FIT_MAX_ROW_HEIGHT, SHEET_AUTO_FIT_MAX_COLUMN_WIDTH } from '../src/features/sheet-workbench/sheet-auto-fit.ts'
import { applySheetAxisDimensionCommand, applySheetViewAxisDimensionCommand, refreshSheetViewAutoRowHeights,
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

let vite, sdk
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  sdk = await vite.ssrLoadModule('/packages/office-studio/src/index.ts')
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
