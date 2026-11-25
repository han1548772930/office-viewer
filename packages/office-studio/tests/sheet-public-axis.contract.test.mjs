import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { sheetAxisSpanAt, sheetColumnWidth, sheetRowHeight, sheetRowIsAutoHeight } from '../src/domain/sheet-axis.ts'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { applySheetPublicAxisInput, normalizeSheetPublicAxisData, transformSheetPublicAxisSources } from '../src/features/sheet-workbench/sheet-public-axis.ts'
import { planWorkbookAxisStructureCommand } from '../src/features/sheet-workbench/sheet-axis-structure-command.ts'
import { isSheetColumnVisible, isSheetRowVisible } from '../src/features/sheet-workbench/sheet-row-visibility.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'

// Drafts only: unified frontend/browser/Rust/WASM verification remains deferred.
const input = () => ({ id: 'public-axis-book', calculationMode: 'manual', styles: { axis: { ff: 'Georgia', fs: 15 } }, sheetOrder: ['data'],
  sheets: { data: { id: 'data', name: 'Data', rowCount: 40, columnCount: 12, defaultRowHeight: 28, defaultColumnWidth: 120,
    rowData: { 0: { h: 48, ia: 0, hd: 0, s: 'axis', custom: { keep: 'row' } }, 1: { h: 32, hd: 1 },
      2: { h: 30, ia: 1, ah: 54 }, 3: { h: 36, ah: 60 }, 4: { h: 0, ia: 0 }, 5: { h: 33, ia: 1, ah: 0 } },
    columnData: { 0: { w: 180, hd: 0, custom: { keep: 'column' } }, 1: { w: 0, hd: 0 }, 2: { w: 90, hd: 1 } },
    cellData: { 0: { 0: { v: 'source' } } } } } })

test('public manual and automatic row heights enter the authoritative CSS-pixel axis model', () => {
  const model = workbookFromInput(input()), sheet = model.sheets[0]
  assert.equal(sheet.axisUnit, 'cssPixel'); assert.equal(sheetRowHeight(sheet, 0), 48); assert.equal(sheetRowIsAutoHeight(sheet, 0), false)
  assert.equal(sheetRowHeight(sheet, 2), 54); assert.equal(sheetRowIsAutoHeight(sheet, 2), true)
  assert.equal(sheetRowHeight(sheet, 3), 60); assert.equal(sheetRowHeight(sheet, 5), 33)
  assert.equal(sheetRowHeight(sheet, 6), 28); assert.equal(sheetRowIsAutoHeight(sheet, 6), true)
  assert.equal(sheet.cells['0:0'].style.font.family, 'Georgia'); assert.equal(sheet.rowStyles[0], 'axis')
})

test('public widths and manual hides remain separate from zero size and preserve geometry', () => {
  const sheet = workbookFromInput(input()).sheets[0]
  assert.equal(sheetColumnWidth(sheet, 0), 180); assert.equal(sheetColumnWidth(sheet, 3), 120)
  assert.equal(isSheetRowVisible(sheet, 1), false); assert.equal(sheetRowHeight(sheet, 1), 32)
  assert.equal(isSheetRowVisible(sheet, 4), false); assert.equal(sheetAxisSpanAt(sheet.rowHiddenSpans, 4), undefined)
  assert.equal(isSheetColumnVisible(sheet, 1), false); assert.equal(sheetAxisSpanAt(sheet.columnHiddenSpans, 1).value, false)
  assert.equal(isSheetColumnVisible(sheet, 2), false); assert.equal(sheetColumnWidth(sheet, 2), 90)
})

test('public row source extensions and explicit flags survive without sharing caller-owned objects', () => {
  const original = input(), model = workbookFromInput(original), sources = model.sheets[0].metadata.publicAxisData
  assert.deepEqual(sources.rowData, original.sheets.data.rowData); assert.deepEqual(sources.columnData, original.sheets.data.columnData)
  original.sheets.data.rowData[0].custom.keep = 'changed'; original.sheets.data.columnData[0].w = 999
  assert.equal(sources.rowData[0].custom.keep, 'row'); assert.equal(sheetColumnWidth(model.sheets[0], 0), 180)
})

test('adjacent public hide and automatic flags coalesce without allocating blank cells', () => {
  const sheet = createBlankSheet(), rowData = Object.fromEntries(Array.from({ length: 30 }, (_, row) => [row, { hd: 1, ia: 0 }]))
  applySheetPublicAxisInput(sheet, { rowData })
  assert.deepEqual(sheet.rowHiddenSpans, [{ start: 0, end: 29, value: true }])
  assert.deepEqual(sheet.rowAutoHeightSpans, [{ start: 0, end: 29, value: false }]); assert.deepEqual(sheet.cells, {})
})

test('explicit local points and mode spans override public axis defaults without dropping raw sources', () => {
  const source = input(); Object.assign(source.sheets.data, { rowHeights: { 0: 70 }, columnWidths: { 0: 210 },
    rowAutoHeightSpans: [{ start: 2, end: 2, value: false }], rowHiddenSpans: [{ start: 1, end: 1, value: false }], hiddenRows: [7] })
  const sheet = workbookFromInput(source).sheets[0]
  assert.equal(sheetRowHeight(sheet, 0), 70); assert.equal(sheetColumnWidth(sheet, 0), 210); assert.equal(sheetRowHeight(sheet, 2), 30)
  assert.equal(isSheetRowVisible(sheet, 1), true); assert.equal(isSheetRowVisible(sheet, 7), false)
  assert.equal(sheet.metadata.publicAxisData.rowData[0].h, 48)
})

test('invalid public axis indices, dimensions and flags reject import without mutating source data', () => {
  const original = input(), before = cloneModel(original)
  for (const rowData of [{ '-1': {} }, { '01': {} }, { 1_048_576: {} }, { 0: { h: -1 } }, { 0: { ah: NaN } },
    { 0: { h: '40' } }, { 0: { ia: 2 } }, { 0: { hd: null } }, { 0: [] }]) {
    assert.throws(() => normalizeSheetPublicAxisData(rowData, 'row'))
  }
  for (const columnData of [{ 16_384: {} }, { 0: { w: Infinity } }, { 0: { hd: '1' } }]) assert.throws(() => normalizeSheetPublicAxisData(columnData, 'column'))
  const invalid = input(); invalid.sheets.data.defaultRowHeight = 0; assert.throws(() => workbookFromInput(invalid))
  assert.deepEqual(original, before)
})

test('malformed local axis spans and point overrides reject rather than bypassing public import validation', () => {
  for (const patch of [{ rowHeights: { 0: -1 } }, { hiddenColumns: [16_384] }, { defaultRowAutoHeight: 1 },
    { defaultRowHeight: 0 }, { axisUnit: 'point' },
    { columnWidthSpans: [{ start: 0, end: 16_384, value: 30 }] }, { rowHiddenSpans: [{ start: 0, end: 1, value: 1 }] }]) {
    const source = input(); Object.assign(source.sheets.data, patch); assert.throws(() => workbookFromInput(source))
    const sheet = createBlankSheet(), before = cloneModel(sheet)
    assert.throws(() => applySheetPublicAxisInput(sheet, source.sheets.data)); assert.deepEqual(sheet, before)
  }
})

test('structural insertion moves raw public custom fields together with dimensions, styles and hidden modes', () => {
  const model = workbookFromInput(input()), original = cloneModel(model)
  const plan = planWorkbookAxisStructureCommand(model, 'data', [{ axis: 'row', kind: 'insert', start: 0, count: 1 }])
  const sheet = plan.workbook.sheets[0]
  assert.equal(sheet.metadata.publicAxisData.rowData[1].custom.keep, 'row'); assert.equal(sheet.rowStyles[1], 'axis')
  assert.equal(sheetRowHeight(sheet, 1), 48); assert.equal(isSheetRowVisible(sheet, 2), false)
  assert.equal(sheet.metadata.publicAxisData.rowData[0], undefined); assert.deepEqual(model, original)
})

test('copying and moving axes preserve source custom fields at their exact new coordinates', () => {
  let model = workbookFromInput(input())
  model = planWorkbookAxisStructureCommand(model, 'data', [{ axis: 'row', kind: 'copy', start: 0, count: 1, destination: 6 }]).workbook
  let sheet = model.sheets[0]
  assert.equal(sheet.metadata.publicAxisData.rowData[6].custom.keep, 'row'); assert.equal(sheetRowHeight(sheet, 6), 48)
  sheet.metadata.publicAxisData.rowData[6].custom.keep = 'copied'; assert.equal(sheet.metadata.publicAxisData.rowData[0].custom.keep, 'row')
  model = planWorkbookAxisStructureCommand(model, 'data', [{ axis: 'column', kind: 'move', start: 0, count: 1, destination: 4 }]).workbook
  sheet = model.sheets[0]
  assert.equal(sheet.metadata.publicAxisData.columnData[3].custom.keep, 'column'); assert.equal(sheetColumnWidth(sheet, 3), 180)
  assert.equal(isSheetColumnVisible(sheet, 1), false)
})

test('deleting axes removes matching public sources while leaving the other axis and metadata intact', () => {
  const model = workbookFromInput(input()), sheet = model.sheets[0], columns = cloneModel(sheet.metadata.publicAxisData.columnData)
  sheet.metadata.vendor = { keep: true }; transformSheetPublicAxisSources(sheet, { axis: 'row', kind: 'delete', start: 0, count: 2 })
  assert.equal(sheet.metadata.publicAxisData.rowData[0].ah, 54); assert.equal(sheet.metadata.publicAxisData.rowData[0].custom, undefined)
  assert.deepEqual(sheet.metadata.publicAxisData.columnData, columns); assert.deepEqual(sheet.metadata.vendor, { keep: true })
})

test('view and saved model round trips use current dimensions rather than reapplying stale raw axis fields', () => {
  const model = workbookFromInput(input()), view = toExcelWorkbookView(model)
  view.sheets[0].columnWidths[0] = 230
  const mapped = applyExcelWorkbookView(model, view), reopened = workbookFromInput(cloneModel(mapped)), sheet = reopened.sheets[0]
  assert.equal(sheetColumnWidth(sheet, 0), 230); assert.equal(sheet.metadata.publicAxisData.columnData[0].w, 180)
  assert.equal(sheetRowHeight(sheet, 2), 54); assert.equal(isSheetColumnVisible(sheet, 2), false)
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const withStudio = run => {
  const studio = new OfficeStudio(), api = FOffice.newAPI(studio), book = api.createWorkbook(input())
  try { return run(studio, api, book) } finally { studio.dispose() }
}

test('Facade imports public axis geometry and uses the shared dimension and visibility content commands', () => withStudio((studio, _api, book) => {
  const sheet = book.getActiveSheet(), before = cloneModel(book.save())
  assert.equal(sheet.getRowHeight(0), 48); assert.equal(sheet.getRowIsAutoHeight(0), false); assert.equal(sheet.getColumnWidth(0), 180)
  sheet.setRowHeightsForced(0, 1, 64); sheet.setColumnWidth(0, 200); sheet.showRows(1)
  assert.equal(sheet.getRowHeight(0), 64); assert.equal(sheet.getColumnWidth(0), 200)
  assert.equal(isSheetRowVisible(studio.getWorkbookModel(book.getId()).sheets[0], 1), true)
  book.undo(); book.undo(); book.undo(); assert.deepEqual(book.save().sheets, before.sheets)
  book.redo(); book.redo(); book.redo(); assert.equal(sheet.getRowHeight(0), 64)
}))

test('Facade structure edits and snapshot reopening preserve public source custom fields and current geometry', () => withStudio((studio, api, book) => {
  const sheet = book.getActiveSheet(); sheet.insertRowsBefore(0, 1)
  assert.equal(book.save().sheets[0].metadata.publicAxisData.rowData[1].custom.keep, 'row')
  book.undo(); assert.equal(book.save().sheets[0].metadata.publicAxisData.rowData[0].custom.keep, 'row'); book.redo()
  const reopened = api.createWorkbook({ ...book.save(), id: 'public-axis-reopened' }), actual = reopened.getActiveSheet()
  assert.equal(actual.getRowHeight(1), 48); assert.equal(actual.getColumnWidth(0), 180)
  assert.equal(studio.getWorkbookModel(reopened.getId()).sheets[0].metadata.publicAxisData.rowData[1].s, 'axis')
}))

test('invalid public imports and protected structure changes do not publish partial dimensions or history', () => withStudio((studio, api, book) => {
  const before = cloneModel(book.save()), invalid = input(); invalid.sheets.data.columnData[0].w = -1
  invalid.id = 'invalid-public-axis'; assert.throws(() => api.createWorkbook(invalid))
  assert.equal(studio.getWorkbookModel(invalid.id), undefined); assert.deepEqual(book.save(), before)
  studio.editor.workbook.model.sheets[0].rangeProtectionRules = [{ id: 'locked', ranges: [{ startRow: 0, endRow: 3, startCol: 0, endCol: 11 }], allowedUserIds: ['other-user'] }]
  const protectedBefore = cloneModel(book.save())
  assert.throws(() => book.getActiveSheet().insertRowsBefore(0, 1), /保护/); assert.deepEqual(book.save(), protectedBefore)
  assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))
