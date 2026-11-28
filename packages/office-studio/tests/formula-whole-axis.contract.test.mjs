import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createFormulaDependencyGraph, evaluateFormula } from '../src/features/formula/formula-engine.ts'
import { parseFormula } from '../src/features/formula/formula-parser.ts'
import { formulaRelativeReferenceSignature, translateFormulaReferences } from '../src/features/formula/formula-references.ts'

function sample() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.name = 'Data'
  for (const [row, column, value] of [[0, 0, 2], [1, 0, 3], [1_048_575, 0, 7], [0, 1, 11], [0, 16_383, 13]]) {
    sheet.cells[`${row}:${column}`] = createCell(row, column, value)
  }
  const other = createBlankSheet("Other's data")
  other.cells['0:0'] = createCell(0, 0, 17)
  workbook.sheets.push(other)
  return workbook
}

test('whole-axis tokens preserve physical bounds and mixed absolute markers in the existing range AST', () => {
  const columns = parseFormula('=SUM($A:C)').arguments[0].reference
  assert.equal(columns.start.row, 0)
  assert.equal(columns.end.row, 1_048_575)
  assert.equal(columns.start.absoluteColumn, true)
  assert.equal(columns.end.absoluteColumn, false)
  assert.equal(columns.start.absoluteRow, true)
  const rows = parseFormula('=SUM(1:$3)').arguments[0].reference
  assert.equal(rows.start.column, 0)
  assert.equal(rows.end.column, 16_383)
  assert.equal(rows.start.absoluteRow, false)
  assert.equal(rows.end.absoluteRow, true)
  assert.equal(rows.start.absoluteColumn, true)
})

test('aggregate and dimension functions include sparse boundary cells across a complete axis', () => {
  const workbook = sample(), sheetId = workbook.sheets[0].id
  for (const [formula, expected] of [
    ['=SUM(A:A)', 12], ['=SUM(A:B)', 23], ['=SUM(1:1)', 26], ['=SUM($1:$2)', 29],
    ['=ROWS(A:C)', 1_048_576], ['=COLUMNS(1:3)', 16_384], ["=SUM('Other''s data'!A:A)", 17],
  ]) assert.equal(evaluateFormula(formula, workbook, sheetId), expected, formula)
})

test('INDEX whole-axis reference results compose with the range operator and preserve zero-axis meaning', () => {
  const workbook = sample(), sheetId = workbook.sheets[0].id
  assert.equal(evaluateFormula('=SUM(INDEX(A:C,0,1):INDEX(A:C,0,2))', workbook, sheetId), 23)
  assert.equal(evaluateFormula('=SUM(INDEX(1:3,1,0):INDEX(1:3,2,0))', workbook, sheetId), 29)
  assert.deepEqual(evaluateFormula('=INDEX(A:C,1048577,1)', workbook, sheetId), { kind: 'error', code: '#REF!' })
})

test('unpaired names strings and malformed mixed axes are never interpreted as complete references', () => {
  assert.deepEqual(parseFormula('="A:C"'), { kind: 'literal', value: 'A:C' })
  assert.deepEqual(parseFormula('=3'), { kind: 'literal', value: 3 })
  assert.deepEqual(parseFormula('=A'), { kind: 'name', name: 'A', sheetName: undefined })
  for (const formula of ['=SUM(A:3)', '=SUM($0:$2)', '=SUM(1:1048577)']) {
    assert.throws(() => parseFormula(formula), undefined, formula)
  }
  const workbook = sample()
  for (const formula of ['=SUM(A1:A)', '=SUM(A:XFE)']) {
    assert.equal(evaluateFormula(formula, workbook, workbook.sheets[0].id).code, '#NAME?', formula)
  }
  assert.equal(translateFormulaReferences('=SUM(A,3,Table[A:C])+A1', 1, 1), '=SUM(A,3,Table[A:C])+B2')
})

test('copying a whole axis translates only its selected dimension and retains relative signatures', () => {
  for (const [formula, rowDelta, columnDelta, expected] of [
    ['=SUM(A:C,$B:$D,1:3,$2:$4)', 2, 1, '=SUM(B:D,$B:$D,3:5,$2:$4)'],
    ["=SUM('Other''s data'!1:3)", 1, -4, "=SUM('Other''s data'!2:4)"],
    ['=SUM(A:A)', -5, 0, '=SUM(A:A)'], ['=SUM(1:1)', 0, -5, '=SUM(1:1)'],
  ]) {
    const changed = translateFormulaReferences(formula, rowDelta, columnDelta)
    assert.equal(changed, expected)
    assert.equal(formulaRelativeReferenceSignature(formula, 5, 5), formulaRelativeReferenceSignature(changed, 5 + rowDelta, 5 + columnDelta))
  }
})

test('dependency collection stores one full-axis interval and connects a sparse formula at its boundary', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  sheet.cells['1048575:0'] = { ...sheet.cells['1048575:0'], formula: '=7' }
  sheet.cells['3:3'] = { ...createCell(3, 3, null), formula: '=SUM(A:A)' }
  const graph = createFormulaDependencyGraph(workbook)
  const references = graph.dependencies.get(`${sheet.id}!3:3`)
  assert.equal(references.length, 1)
  assert.equal(references[0].kind, 'range')
  assert.equal(references[0].end.row, 1_048_575)
  assert.equal(graph.dependents.get(`${sheet.id}!1048575:0`).has(`${sheet.id}!3:3`), true)
})
