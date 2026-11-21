import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { evaluateFormula, recalculateWorkbookFormulas } from '../src/features/formula/formula-engine.ts'

function fixture() {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.id = 'data'
  sheet.name = 'Data'
  const values = ['123', true, false, null, 8, -2, '', 'hello', 0]
  for (let index = 0; index < values.length; index += 1) sheet.cells[`${index}:0`] = createCell(index, 0, values[index])
  sheet.cells['0:1'] = { ...createCell(0, 1, '#DIV/0!'), dataType: 'error' }
  sheet.cells['1:1'] = { ...createCell(1, 1, null), formula: '=""' }
  sheet.cells['2:1'] = { ...createCell(2, 1, null), formula: '=TRUE' }
  sheet.cells['3:1'] = { ...createCell(3, 1, null), formula: '=1/0' }
  const other = createBlankSheet()
  other.id = 'other'
  other.name = 'Other Data'
  other.cells = {
    '0:0': createCell(0, 0, '99'),
    '1:0': createCell(1, 0, true),
    '2:0': createCell(2, 0, 12),
  }
  workbook.sheets.push(other)
  return workbook
}

function evaluate(source, workbook = fixture()) { return evaluateFormula(`=${source}`, workbook, 'data') }
const error = code => ({ kind: 'error', code })

test('single-cell numeric text and boolean references obey the same rules as 1x1 ranges', () => {
  for (const address of ['A1', 'A2', 'A3']) {
    for (const name of ['SUM', 'COUNT', 'MIN', 'MAX']) {
      assert.equal(evaluate(`${name}(${address})`), 0, `${name}(${address})`)
      assert.equal(evaluate(`${name}(${address}:${address})`), 0, `${name}(1x1 range)`)
    }
    assert.deepEqual(evaluate(`AVERAGE(${address})`), error('#DIV/0!'))
    assert.equal(evaluate(`COUNTA(${address})`), 1)
  }
})

test('direct numeric text, booleans and scalar expressions are coerced by numeric aggregates', () => {
  assert.equal(evaluate('SUM("123",TRUE,FALSE)'), 124)
  assert.equal(evaluate('AVERAGE("123",TRUE,FALSE)'), 124 / 3)
  assert.equal(evaluate('MIN("123",TRUE,FALSE)'), 0)
  assert.equal(evaluate('MAX("123",TRUE,FALSE)'), 123)
  assert.equal(evaluate('COUNT("123",TRUE,FALSE)'), 3)
  assert.equal(evaluate('COUNTA("123",TRUE,FALSE)'), 3)
  assert.equal(evaluate('SUM(IF(TRUE,"2","text"),TRUE)'), 3)
  assert.equal(evaluate('COUNT(IF(TRUE,"2","text"),FALSE)'), 2)
})

test('a text or empty-string literal errors for numeric aggregates but COUNTA counts it and COUNT ignores it', () => {
  for (const text of ['"hello"', '""', '"   "']) {
    for (const name of ['SUM', 'AVERAGE', 'MIN', 'MAX']) assert.deepEqual(evaluate(`${name}(${text})`), error('#VALUE!'))
    assert.equal(evaluate(`COUNT(${text})`), 0)
    assert.equal(evaluate(`COUNTA(${text})`), 1)
  }
})

test('sparse ranges ignore null, text and boolean values while numeric zeros still participate', () => {
  assert.equal(evaluate('SUM(A1:A9)'), 6)
  assert.equal(evaluate('AVERAGE(A1:A9)'), 2)
  assert.equal(evaluate('COUNT(A1:A9)'), 3)
  assert.equal(evaluate('COUNTA(A1:A9)'), 8)
  assert.equal(evaluate('MIN(A1:A9)'), -2)
  assert.equal(evaluate('MAX(A1:A9)'), 8)
  assert.equal(evaluate('COUNT(A100:A1048576)'), 0)
  assert.equal(evaluate('COUNTA(A100:A1048576)'), 0)
})

test('COUNT ignores errors and COUNTA includes errors consistently for direct expressions, cells and ranges', () => {
  for (const expression of ['B1', 'B1:B1', '#DIV/0!', 'NA()', '1/0']) {
    assert.equal(evaluate(`COUNT(${expression})`), 0, expression)
    assert.equal(evaluate(`COUNTA(${expression})`), 1, expression)
    for (const name of ['SUM', 'AVERAGE', 'MIN', 'MAX']) {
      const value = evaluate(`${name}(${expression})`)
      assert.equal(value.kind, 'error', `${name}(${expression})`)
    }
  }
  assert.equal(evaluate('COUNT(B1,NA(),1/0,5)'), 1)
  assert.equal(evaluate('COUNTA(B1,NA(),1/0,5)'), 4)
})

test('formula result cells keep reference aggregation rules including empty strings, booleans and errors', () => {
  assert.equal(evaluate('COUNT(B2:B4)'), 0)
  assert.equal(evaluate('COUNTA(B2:B4)'), 3)
  assert.equal(evaluate('COUNTA(B2)'), 1)
  assert.equal(evaluate('SUM(B3)'), 0)
  assert.equal(evaluate('COUNT(B3)'), 0)
  assert.deepEqual(evaluate('SUM(B4)'), error('#DIV/0!'))
  assert.equal(evaluate('COUNT(B4)'), 0)
})

test('blank referenced cells are excluded instead of inserting zero into average or extrema', () => {
  assert.equal(evaluate('AVERAGE(A4,A5)'), 8)
  assert.equal(evaluate('MIN(A4,A5)'), 8)
  assert.equal(evaluate('MAX(A4,A6)'), -2)
  assert.equal(evaluate('COUNTA(A4,A100)'), 0)
  assert.deepEqual(evaluate('AVERAGE(A4,A100)'), error('#DIV/0!'))
})

test('mixed reference, range and literal arguments preserve each source type and count repeated arguments', () => {
  assert.equal(evaluate('SUM(A1,A2,A5:A9,TRUE,"5",2)'), 14)
  assert.equal(evaluate('COUNT(A1,A2,A5:A9,TRUE,"5",2)'), 6)
  assert.equal(evaluate('SUM(A5,A5,A5:A5)'), 24)
  assert.equal(evaluate('COUNTA(A5,A5,A5:A5)'), 3)
  assert.equal(evaluate("SUM('Other Data'!A1,'Other Data'!A2,'Other Data'!A3)"), 12)
  assert.equal(evaluate("COUNT('Other Data'!A1,'Other Data'!A2,'Other Data'!A3)"), 1)
})

test('direct numeric/date strings use the workbook date system, preserving scalar coercion boundaries', () => {
  assert.equal(evaluate('SUM("1,234.5","2e1","25%")'), 1254.75)
  assert.equal(evaluate('COUNT("1,234.5","2e1","25%","text")'), 3)
  const workbook = fixture()
  assert.equal(evaluate('SUM("1904-1-1")', workbook), 1462)
  workbook.date1904 = true
  assert.equal(evaluate('SUM("1904-1-1")', workbook), 0)
})

test('public aggregate arity is 1 through 255 arguments', () => {
  for (const name of ['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX']) {
    assert.deepEqual(evaluate(`${name}()`), error('#VALUE!'))
    assert.deepEqual(evaluate(`${name}(${Array(256).fill('1').join(',')})`), error('#VALUE!'))
  }
  assert.equal(evaluate(`COUNT(${Array(255).fill('1').join(',')})`), 255)
})

test('workbook recalculation uses the corrected aggregate rules and retains the dependency graph', () => {
  const workbook = fixture()
  const sheet = workbook.sheets[0]
  sheet.cells['0:2'] = { ...createCell(0, 2, null), formula: '=SUM(A1,A2,A5)' }
  sheet.cells['1:2'] = { ...createCell(1, 2, null), formula: '=COUNT(B1,B4,A5)' }
  sheet.cells['2:2'] = { ...createCell(2, 2, null), formula: '=COUNTA(B1,B2,A4)' }
  const result = recalculateWorkbookFormulas(workbook)
  assert.equal(result.workbook.sheets[0].cells['0:2'].value, 8)
  assert.equal(result.workbook.sheets[0].cells['1:2'].value, 1)
  assert.equal(result.workbook.sheets[0].cells['2:2'].value, 2)
  assert.ok(result.graph.dependents.get('data!0:0').has('data!0:2'))
  assert.equal(workbook.sheets[0].cells['0:2'].value, null)
})
