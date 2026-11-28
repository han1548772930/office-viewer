import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { evaluateFormula, recalculateWorkbookFormulas } from '../src/features/formula/formula-engine.ts'
import { formulaCalendarSerial, readFormulaCalendarValue, rewriteFormulaImportCell } from '../src/features/formula/formula-import-context.ts'
import { remapChartSourceDateCell } from '../src/features/charts/chart-source-import-dates.ts'

// Deferred date-system contracts, including values native Date cannot represent.
const error = code => ({ kind: 'error', code })
function sample(date1904 = false) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'calendar'; workbook.date1904 = date1904; sheet.id = 'data'; sheet.name = 'Data'; workbook.activeSheetId = sheet.id
  return { workbook, sheet }
}
function marked(sheet, row, column, formula, sourceDate1904, dateResult, value = null) {
  const cell = { ...createCell(row, column, value), formula, style: { ...createCell(row, column, value).style, numberFormat: dateResult ? 'yyyy-mm-dd' : 'General' },
    metadata: { chartDependencyFormulaSource: { version: 1, formula, date1904: sourceDate1904, dateResult } } }
  sheet.cells[`${row}:${column}`] = cell
  return cell
}

test('YEAR MONTH and DAY use the same serial 0, serial 60 fractions and maximum-date fields as number formatting', () => {
  const { workbook, sheet } = sample()
  for (const [serial, expected] of [[0, [1900, 1, 0]], [0.5, [1900, 1, 0]], [59, [1900, 2, 28]],
    [60.5, [1900, 2, 29]], [61, [1900, 3, 1]], [2958465.75, [9999, 12, 31]]]) {
    for (const [index, name] of ['YEAR', 'MONTH', 'DAY'].entries()) assert.equal(evaluateFormula(`=${name}(${serial})`, workbook, sheet.id), expected[index])
  }
  for (const serial of [-1, 2958466]) assert.deepEqual(evaluateFormula(`=DAY(${serial})`, workbook, sheet.id), error('#NUM!'))
  workbook.date1904 = true
  assert.equal(evaluateFormula('=YEAR(0)', workbook, sheet.id), 1904)
  assert.equal(evaluateFormula('=DAY(0)', workbook, sheet.id), 1)
  assert.equal(evaluateFormula('=YEAR(2957003.75)', workbook, sheet.id), 9999)
  assert.deepEqual(evaluateFormula('=DAY(2957004)', workbook, sheet.id), error('#NUM!'))
})

test('DATE normalizes from the first month day so overflow crosses the fictitious leap day once', () => {
  const { workbook, sheet } = sample()
  for (const [formula, expected] of [['=DATE(1900,1,0)', 0], ['=DATE(1900,2,29)', 60], ['=DATE(1900,2,30)', 61],
    ['=DATE(1900,3,0)', 60], ['=DATE(2024.9,1.9,1.9)', 45292], ['=DATE(9999,12,31)', 2958465]]) {
    assert.equal(evaluateFormula(formula, workbook, sheet.id), expected, formula)
  }
  for (const formula of ['=DATE(-1,1,1)', '=DATE(1900,1,-1)', '=DATE(9999,12,32)']) assert.deepEqual(evaluateFormula(formula, workbook, sheet.id), error('#NUM!'))
  workbook.date1904 = true
  assert.equal(evaluateFormula('=DATE(1904,1,1)', workbook, sheet.id), 0)
  assert.deepEqual(evaluateFormula('=DATE(1900,2,29)', workbook, sheet.id), error('#NUM!'))
})

test('early physical dates use 1461 days while modern dates use 1462 and preserve time fractions', () => {
  for (const [serial, projected] of [[0, -1461], [1.25, -1459.75], [59.5, -1401.5], [61, -1401], [1462, 0], [45292.5, 43830.5]]) {
    assert.equal(formulaCalendarSerial(serial, false, true), projected)
    assert.equal(formulaCalendarSerial(projected, true, false), serial)
  }
  assert.equal(formulaCalendarSerial(60.5, false, true), -1401.5)
})

test('exact imported originals disambiguate the fictitious leap day from its physical date carrier', () => {
  const { workbook, sheet } = sample(true)
  for (const [row, serial] of [[0, 59.5], [1, 60.5]]) {
    const value = formulaCalendarSerial(serial, false, true)
    sheet.cells[`${row}:0`] = { ...createCell(row, 0, value), style: { ...createCell(row, 0, value).style, numberFormat: 'yyyy-mm-dd' },
      metadata: { chartDependencyDateSource: { version: 1, date1904: false, value: serial, targetDate1904: true, targetValue: value } } }
    marked(sheet, row, 1, `=DAY(A${row + 1})`, false, false)
  }
  const result = recalculateWorkbookFormulas(workbook).workbook.sheets[0]
  assert.equal(result.cells['0:1'].value, 28); assert.equal(result.cells['1:1'].value, 29)
  assert.equal(readFormulaCalendarValue(sheet.cells['1:0'], -1401.5, true).value, 60.5)
  assert.deepEqual(evaluateFormula('=DAY(A1)', workbook, sheet.id), error('#NUM!'))
})

test('fresh formula results retain their source serial before dependent evaluation and metadata publication', () => {
  const { workbook, sheet } = sample(true)
  marked(sheet, 0, 0, '=DATE(1900,3,0)', false, true)
  marked(sheet, 0, 1, '=DAY(A1)', false, false)
  const next = recalculateWorkbookFormulas(workbook).workbook
  assert.equal(next.sheets[0].cells['0:0'].value, -1402)
  assert.equal(next.sheets[0].cells['0:0'].metadata.chartDependencyDateSource.value, 60)
  assert.equal(next.sheets[0].cells['0:1'].value, 29)
  assert.equal(evaluateFormula('=DAY(A1)', next, sheet.id, { formulaCell: { row: 0, column: 1 } }), 29)
  assert.equal(sheet.cells['0:0'].value, null)
})

test('returning an imported early date to its original workbook calendar recovers serial 60 exactly', () => {
  const cell = { ...createCell(0, 0, -1401.5), style: { ...createCell(0, 0, 0).style, numberFormat: 'yyyy-mm-dd hh:mm' },
    metadata: { chartDependencyDateSource: { version: 1, date1904: false, value: 60.5, targetDate1904: true, targetValue: -1401.5 } } }
  remapChartSourceDateCell(cell, 1462, false)
  assert.equal(cell.value, 60.5)
})

test('automatic formula rewrites preserve exact date ownership while manual replacements invalidate it', () => {
  const { sheet } = sample(true), cell = marked(sheet, 0, 0, '=A2', false, true, -1402)
  cell.metadata.chartDependencyDateSource = { version: 1, date1904: false, value: 60, targetDate1904: true, targetValue: -1402, targetFormula: '=A2' }
  rewriteFormulaImportCell(cell, '=A3')
  assert.equal(cell.metadata.chartDependencyDateSource.targetFormula, '=A3')
  assert.equal(readFormulaCalendarValue(cell, -1402, true).value, 60)
  cell.formula = '=40000'; rewriteFormulaImportCell(cell, '=40001')
  assert.equal(readFormulaCalendarValue(cell, -1402, true), undefined)
})

test('ordinary numeric quantities and time-only cells do not acquire a date-system offset', () => {
  for (const pattern of ['General', '0.00', 'hh:mm:ss', '[h]:mm:ss']) {
    const cell = { ...createCell(0, 0, 2.5), style: { ...createCell(0, 0, 0).style, numberFormat: pattern } }
    remapChartSourceDateCell(cell, -1462, true)
    assert.equal(cell.value, 2.5, pattern)
  }
})
