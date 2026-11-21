import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { canSetSheetHyperlink, normalizeExternalSheetHyperlink, removeSheetCellHyperlink, setSheetCellHyperlink } from '../src/features/sheet-workbench/sheet-hyperlink.ts'

test('external hyperlinks accept web and mail targets and reject unsafe protocols or credentials', () => {
  assert.equal(normalizeExternalSheetHyperlink(' https://example.com/path '), 'https://example.com/path')
  assert.equal(normalizeExternalSheetHyperlink('mailto:team@example.com'), 'mailto:team@example.com')
  assert.equal(normalizeExternalSheetHyperlink('javascript:alert(1)'), undefined)
  assert.equal(normalizeExternalSheetHyperlink('https://user:secret@example.com'), undefined)
})

test('hyperlink edits preserve existing cell content unless a display label is supplied', () => {
  const model = createBlankWorkbook()
  const sheet = model.sheets[0]
  assert.ok(sheet)
  sheet.cells['0:0'] = createCell(0, 0, 'Quarter total')
  setSheetCellHyperlink(sheet, 0, 0, { target: 'https://example.com', display: '' })
  assert.equal(sheet.cells['0:0']?.value, 'Quarter total')
  setSheetCellHyperlink(sheet, 0, 0, { target: 'https://example.com', display: 'Report' })
  assert.equal(sheet.cells['0:0']?.value, 'Report')
})

test('hyperlink target obeys upstream formula and validation restrictions', () => {
  const model = createBlankWorkbook()
  const sheet = model.sheets[0]
  assert.ok(sheet)
  sheet.cells['0:0'] = { ...createCell(0, 0, '=SUM(B1:B4)'), formula: '=SUM(B1:B4)' }
  assert.equal(canSetSheetHyperlink(sheet, 0, 0, sheet.cells['0:0']?.formula), false)
  assert.throws(() => setSheetCellHyperlink(sheet, 0, 0, { target: 'https://example.com', display: 'Blocked' }), /公式/)
  assert.equal(sheet.cells['0:0']?.formula, '=SUM(B1:B4)')
  sheet.cells['0:0'] = createCell(0, 0, 'input')
  sheet.dataValidations = [{ id: 'list', type: 'list', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }], formula1: '"input"' }]
  assert.equal(canSetSheetHyperlink(sheet, 0, 0), false)
})

test('cell hyperlinks roundtrip through the normalized workbook view and preserve cell style', () => {
  const model = createBlankWorkbook()
  const sheet = model.sheets[0]
  assert.ok(sheet)
  sheet.cells['0:0'] = { ...createCell(0, 0, 'Quarter total'), style: { ...createCell(0, 0).style, font: { ...createCell(0, 0).style.font, bold: true } } }
  setSheetCellHyperlink(sheet, 0, 0, { target: 'example.com/report', display: '' })
  const view = toExcelWorkbookView(model)
  assert.equal(view.sheets[0]?.cells.A1?.hyperlink, 'http://example.com/report')
  const restored = applyExcelWorkbookView(model, view)
  assert.equal(restored.sheets[0]?.cells['0:0']?.hyperlink, 'http://example.com/report')
  assert.equal(restored.sheets[0]?.cells['0:0']?.value, 'Quarter total')
  assert.equal(restored.sheets[0]?.cells['0:0']?.style.font.bold, true)
  view.sheets[0].cells.A1.hyperlink = null
  const removed = applyExcelWorkbookView(model, view)
  assert.equal(removed.sheets[0]?.cells['0:0']?.hyperlink, null)
  removeSheetCellHyperlink(sheet, 0, 0)
  assert.equal(sheet.cells['0:0']?.hyperlink, null)
  assert.equal(sheet.cells['0:0']?.value, 'Quarter total')
})
