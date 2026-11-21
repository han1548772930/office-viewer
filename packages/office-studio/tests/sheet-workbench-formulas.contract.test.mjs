import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { applyExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { recalculateSheetViewFormulas } from '../src/features/sheet-workbench/sheet-view-formulas.ts'

function workbook() {
  return {
    id: 'book', name: '计算', activeSheetId: 'a',
    sheets: [
      { id: 'a', name: '统计', rowCount: 7, columnCount: 4, cells: {
        A1: { value: '2', rawValue: 2, bold: true },
        B1: { value: '', formula: "='源 数据'!A1+A1", cachedValue: 'old', italic: true, rawNumberFormat: '#,##0.00' },
        C1: { value: '', formula: '=B1*2', cachedValue: 'old' },
      }, merges: ['A3:B3'], freeze: { rows: 1, columns: 0 }, hiddenRows: [4], columnWidths: { 0: 123 } },
      { id: 'b', name: '源 数据', rowCount: 9, columnCount: 3, cells: {
        A1: { value: '8', rawValue: 8 },
      }, visible: false },
    ],
  }
}

test('the workbench projects authoritative cross-sheet formula results without expanding its view', () => {
  const source = workbook()
  const snapshot = structuredClone(source)
  const result = recalculateSheetViewFormulas(source)
  assert.deepEqual(result.diagnostics, [])
  assert.equal(result.workbook.sheets[0].cells.B1.cachedValue, '10')
  assert.equal(result.workbook.sheets[0].cells.B1.rawValue, 10)
  assert.equal(result.workbook.sheets[0].cells.C1.cachedValue, '20')
  assert.equal(result.workbook.sheets[0].rowCount, 7)
  assert.equal(result.workbook.sheets[0].columnCount, 4)
  assert.deepEqual(result.workbook.sheets[0].merges, ['A3:B3'])
  assert.deepEqual(result.workbook.sheets[0].hiddenRows, [4])
  assert.equal(result.workbook.sheets[0].cells.B1.rawNumberFormat, '#,##0.00')
  assert.equal(result.workbook.sheets[0].cells.B1.italic, true)
  assert.equal(result.workbook.sheets[0].cells.A1, source.sheets[0].cells.A1)
  assert.equal(result.workbook.sheets[1], source.sheets[1])
  assert.deepEqual(source, snapshot)
})

test('formula dependency changes update caches and preserve source formulas', () => {
  const source = workbook()
  const original = applyExcelWorkbookView(createBlankWorkbook(), source)
  source.sheets[1].cells.A1.value = '13'
  const result = recalculateSheetViewFormulas(source, original)
  assert.equal(result.workbook.sheets[0].cells.B1.rawValue, 15)
  assert.equal(result.workbook.sheets[0].cells.C1.rawValue, 30)
  assert.equal(result.workbook.sheets[0].cells.B1.formula, "='源 数据'!A1+A1")
  assert.equal(original.sheets[1].cells['0:0'].value, 8)
})

test('standalone imported text remains text and does not become numeric because it looks numeric', () => {
  const source = workbook()
  source.sheets[0].cells.A1 = { value: '123', rawValue: '123' }
  source.sheets[0].cells.B1 = { value: '', formula: '=ISNUMBER(A1)', cachedValue: '' }
  source.sheets[0].cells.C1 = { value: '', formula: '=ISTEXT(A1)', cachedValue: '' }
  const result = recalculateSheetViewFormulas(source)
  assert.equal(result.workbook.sheets[0].cells.B1.rawValue, false)
  assert.equal(result.workbook.sheets[0].cells.C1.rawValue, true)
})

test('unsupported imported formulas preserve typed cache and formatted display and report diagnostics', () => {
  const source = workbook()
  source.sheets[0].cells.B1 = { value: '', formula: '=CUSTOM_UNSUPPORTED(A1)', rawValue: 1234.5, cachedValue: '1,234.50', rawNumberFormat: '#,##0.00' }
  const original = applyExcelWorkbookView(createBlankWorkbook(), source)
  original.sheets[0].cells['0:1'].value = 1234.5
  const result = recalculateSheetViewFormulas(source, original)
  assert.equal(result.workbook.sheets[0].cells.B1.rawValue, 1234.5)
  assert.equal(result.workbook.sheets[0].cells.B1.cachedValue, '1,234.50')
  assert.equal(result.workbook.sheets[0].cells.B1.formula, '=CUSTOM_UNSUPPORTED(A1)')
  assert.ok(result.diagnostics.some(item => item.code === 'FORMULA_UNSUPPORTED_CACHE_PRESERVED'))
})

test('formula errors and cycles remain visible results without throwing or removing formula source', () => {
  const source = workbook()
  source.sheets[0].cells.B1.formula = '=C1'
  source.sheets[0].cells.C1.formula = '=B1'
  const cyclic = recalculateSheetViewFormulas(source)
  assert.equal(cyclic.workbook.sheets[0].cells.B1.cachedValue, '#CIRCULAR!')
  assert.equal(cyclic.workbook.sheets[0].cells.C1.cachedValue, '#CIRCULAR!')
  source.sheets[0].cells.B1.formula = '=1/0'
  source.sheets[0].cells.C1.formula = '=B1'
  const error = recalculateSheetViewFormulas(source)
  assert.equal(error.workbook.sheets[0].cells.B1.rawValue, '#DIV/0!')
  assert.equal(error.workbook.sheets[0].cells.C1.rawValue, '#DIV/0!')
})

test('unchanged results preserve imported formatting and repeated recalculation is stable', () => {
  const source = workbook()
  source.sheets[0].cells.B1 = { value: '', formula: '=A1/8', rawValue: 0.25, cachedValue: '25.00%', displayValue: '25.00%', numberFormat: 'percent' }
  source.sheets[0].cells.C1 = { value: '', formula: '=B1*2', rawValue: 0.5, cachedValue: '0.5', displayValue: '0.5' }
  const result = recalculateSheetViewFormulas(source)
  assert.equal(result.workbook, source)
  assert.equal(result.workbook.sheets[0].cells.B1.cachedValue, '25.00%')
  assert.deepEqual(result.diagnostics, [])
})
