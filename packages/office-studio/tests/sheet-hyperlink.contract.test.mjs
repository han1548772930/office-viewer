import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { canSetSheetHyperlink, normalizeExternalSheetHyperlink, normalizeSheetHyperlinkTarget, parseSheetHyperlinkTarget, removeSheetCellHyperlink, setSheetCellHyperlink } from '../src/features/sheet-workbench/sheet-hyperlink.ts'

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
  setSheetCellHyperlink(model, sheet, 0, 0, { target: 'https://example.com', display: '' })
  assert.equal(sheet.cells['0:0']?.value, 'Quarter total')
  setSheetCellHyperlink(model, sheet, 0, 0, { target: 'https://example.com', display: 'Report' })
  assert.equal(sheet.cells['0:0']?.value, 'Report')
})

test('hyperlink labels replace numeric and boolean types and remain text through view roundtrip and removal', () => {
  const model = createBlankWorkbook()
  const sheet = model.sheets[0]
  assert.ok(sheet)
  const samples = [
    { value: 42.5, dataType: 'number', display: '', expected: '42.5' },
    { value: false, dataType: 'boolean', display: '', expected: 'false' },
    { value: 7, dataType: 'number', display: '=SUM(1,2)', expected: '=SUM(1,2)' },
  ]
  for (const [row, sample] of samples.entries()) {
    sheet.cells[`${row}:0`] = { ...createCell(row, 0, sample.value), dataType: sample.dataType, cachedDataType: sample.dataType }
    setSheetCellHyperlink(model, sheet, row, 0, { target: 'https://example.com/report', display: sample.display })
    assert.equal(sheet.cells[`${row}:0`].value, sample.expected)
    assert.equal(sheet.cells[`${row}:0`].display, sample.expected)
    assert.equal(sheet.cells[`${row}:0`].formula, null)
    assert.equal(sheet.cells[`${row}:0`].dataType, 'text')
    assert.equal(sheet.cells[`${row}:0`].cachedDataType, null)
  }

  const view = toExcelWorkbookView(model)
  const restored = applyExcelWorkbookView(model, view)
  for (const [row, sample] of samples.entries()) {
    assert.equal(view.sheets[0].cells[`A${row + 1}`].dataType, 'text')
    const restoredSheet = restored.sheets[0]
    removeSheetCellHyperlink(restoredSheet, row, 0)
    assert.equal(restoredSheet.cells[`${row}:0`].hyperlink, null)
    assert.equal(restoredSheet.cells[`${row}:0`].value, sample.expected)
    assert.equal(restoredSheet.cells[`${row}:0`].formula, null)
    assert.equal(restoredSheet.cells[`${row}:0`].dataType, 'text')
    assert.equal(restoredSheet.cells[`${row}:0`].cachedDataType, null)
  }
})

test('hyperlink target obeys upstream formula and validation restrictions', () => {
  const model = createBlankWorkbook()
  const sheet = model.sheets[0]
  assert.ok(sheet)
  sheet.cells['0:0'] = { ...createCell(0, 0, '=SUM(B1:B4)'), formula: '=SUM(B1:B4)' }
  assert.equal(canSetSheetHyperlink(sheet, 0, 0, sheet.cells['0:0']?.formula), false)
  assert.throws(() => setSheetCellHyperlink(model, sheet, 0, 0, { target: 'https://example.com', display: 'Blocked' }), /公式/)
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
  setSheetCellHyperlink(model, sheet, 0, 0, { target: 'example.com/report', display: '' })
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

test('internal links validate sheet ids and cell ranges, and normalize to Univer payload syntax', () => {
  const model = createBlankWorkbook()
  const first = model.sheets[0]
  assert.ok(first)
  const target = { ...first, id: 'target-sheet', name: "Quarter's Report", cells: {} }
  model.sheets.push(target)
  const payload = normalizeSheetHyperlinkTarget("#'Quarter''s Report'!$B$2:C4", model)
  assert.equal(payload, '#gid=target-sheet&range=B2%3AC4')
  assert.deepEqual(parseSheetHyperlinkTarget(payload, model), {
    type: 'internal', sheetId: 'target-sheet', sheetName: "Quarter's Report",
    range: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
  })
  assert.equal(normalizeSheetHyperlinkTarget('#gid=missing&range=A1', model), undefined)
  assert.equal(normalizeSheetHyperlinkTarget('#gid=target-sheet&range=XFE1', model), undefined)
  assert.equal(normalizeSheetHyperlinkTarget('#gid=target-sheet&range=B2:A1', model), undefined)
})
