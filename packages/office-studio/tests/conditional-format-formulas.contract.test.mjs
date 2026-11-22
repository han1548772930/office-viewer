import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { evaluateFormula } from '../src/features/formula/formula-engine.ts'
import { parseFormula } from '../src/features/formula/formula-parser.ts'
import { formulaRelativeReferenceSignature, translateFormulaReferences } from '../src/features/formula/formula-references.ts'
import { conditionalFormatFormulaAt, reanchorConditionalFormatRule } from '../src/features/conditional-formatting/conditional-format-formulas.ts'
import cases from '../../../tests/fixtures/conditional-formula-projection.json'

const area = (startRow, startCol, endRow = startRow, endCol = startCol) => ({ startRow, startCol, endRow, endCol })
function sample() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.name = 'Data'
  for (const [row, col, value] of [[0, 0, 7], [0, 1, 3], [1, 0, 2], [1, 1, 4], [1_048_575, 0, 19], [0, 16_383, 23]]) sheet.cells[`${row}:${col}`] = createCell(row, col, value)
  const other = createBlankSheet('Other')
  for (const [row, col, value] of [[0, 0, 2], [0, 1, 3], [1, 0, 4], [1, 1, 5]]) other.cells[`${row}:${col}`] = createCell(row, col, value)
  workbook.sheets.push(other)
  return workbook
}
const evaluateAt = (formula, workbook, row, column) => evaluateFormula(formula, workbook, workbook.sheets[0].id, { formulaCell: { row, column } })
const fixture = name => cases.find(entry => entry.name === name)

test('shared standard-projection fixtures agree with ordinary reference translation and semantic signatures', () => {
  for (const entry of cases.filter(value => value.ordinary)) {
    assert.equal(translateFormulaReferences(entry.formula, entry.target.row - entry.origin.row, entry.target.col - entry.origin.col), entry.projected, entry.name)
    assert.equal(formulaRelativeReferenceSignature(entry.formula, entry.origin.row, entry.origin.col), formulaRelativeReferenceSignature(entry.projected, entry.target.row, entry.target.col), entry.name)
  }
})

test('relative signatures preserve reference offsets, absolute dimensions and literal text without REF projection', () => {
  assert.equal(formulaRelativeReferenceSignature('=a1+$B$2', 1, 1), formulaRelativeReferenceSignature('=C3+$B$2', 3, 3))
  assert.notEqual(formulaRelativeReferenceSignature('=A1+$B$2', 1, 1), formulaRelativeReferenceSignature('=$A$1+$B$2', 1, 1))
  assert.notEqual(formulaRelativeReferenceSignature('=A1="A1"', 1, 1), formulaRelativeReferenceSignature('=B2="B2"', 2, 2))
  assert.notEqual(formulaRelativeReferenceSignature('=A1', 1, 1), formulaRelativeReferenceSignature('=#REF!', 0, 0))
})

test('safe rebasing removes an explicit origin while preserving every formula-bearing operand and unknown data', () => {
  const rule = { id: 'formula', type: 'expression', operator: '', formula1: '=A1>0', value: '=A1>0', ranges: [area(0, 0)], formulaAnchor: { row: 1, col: 0 },
    style: { fill: '#123456' }, priority: 1, stopIfTrue: false, config: { min: { type: 'formula', value: '=A1' }, extra: ['A1'] } }
  const moved = reanchorConditionalFormatRule(rule, [area(3, 2)])
  assert.equal(moved.formulaAnchor, undefined); assert.equal(moved.formula1, '=C3>0'); assert.equal(moved.value, '=C3>0'); assert.equal(moved.config.min.value, '=C3')
  assert.deepEqual(moved.config.extra, ['A1']); assert.equal(rule.formula1, '=A1>0')
  assert.equal(conditionalFormatFormulaAt(moved.formula1, moved, 4, 2), '=C4>0')
})

test('rebasing at the same origin retains exact imported formula spelling without rewriting operands', () => {
  const rule = { id: 'formula', type: 'expression', operator: '', formula1: '=a1+$B$2', ranges: [area(1, 1)],
    style: { fill: '#123456' }, priority: 1, stopIfTrue: false, config: { min: { type: 'formula', value: '=UNSUPPORTED([VendorColumn])' } } }
  const moved = reanchorConditionalFormatRule(rule, rule.ranges)
  assert.deepEqual(moved, rule)
  moved.config.min.value = '=A1'; assert.equal(rule.config.min.value, '=UNSUPPORTED([VendorColumn])')
})

test('explicit evaluation coordinates give ROW and COLUMN their actual conditional cell without changing the model', () => {
  const workbook = sample(), before = JSON.stringify(workbook)
  assert.equal(evaluateAt('=ROW()+COLUMN()', workbook, 3, 4), 9)
  assert.equal(evaluateAt('=IF(TRUE,ROW(),0)', workbook, 8, 2), 9)
  assert.equal(evaluateFormula('=ROW()', workbook, workbook.sheets[0].id), 1)
  assert.equal(JSON.stringify(workbook), before)
  assert.deepEqual(evaluateAt('=ROW()', workbook, -1, 0), { kind: 'error', code: '#REF!' })
})

test('standard INDEX projection retains invalid first rows and later valid scalar references independently', () => {
  const workbook = sample(), entry = fixture('partly valid rows')
  for (let row = 0; row < 4; row++) {
    const expected = evaluateAt(translateFormulaReferences(entry.formula, row - entry.origin.row, 0), workbook, row, 0)
    assert.deepEqual(evaluateAt(entry.projected, workbook, row, 0), expected)
  }
  assert.deepEqual(evaluateAt(entry.projected, workbook, 0, 0), { kind: 'error', code: '#REF!' })
  assert.equal(evaluateAt(entry.projected, workbook, 1, 0), true)
  assert.equal(evaluateAt(entry.projected, workbook, 2, 0), false)
})

test('column underflow becomes REF rather than the whole INDEX column and recovers at later cells', () => {
  const workbook = sample(), entry = fixture('partly valid columns')
  assert.deepEqual(evaluateAt(entry.projected, workbook, 0, 0), { kind: 'error', code: '#REF!' })
  assert.equal(evaluateAt(entry.projected, workbook, 0, 1), 7)
  assert.equal(evaluateAt(entry.projected, workbook, 0, 2), 3)
})

test('absolute operands and IFERROR remain valid while each relative reference can overflow', () => {
  const workbook = sample(), entry = fixture('absolute operand survives an invalid relative operand')
  assert.equal(evaluateAt(entry.projected, workbook, 0, 0), 3)
  assert.equal(evaluateAt(entry.projected, workbook, 1, 0), 10)
})

test('upper row and column references recover at legal coordinates and use sparse full-sheet INDEX ranges', () => {
  const workbook = sample(), rows = fixture('upper row boundary'), columns = fixture('upper column boundary')
  assert.equal(evaluateAt(rows.projected, workbook, 0, 0), 19)
  assert.deepEqual(evaluateAt(rows.projected, workbook, 1, 0), { kind: 'error', code: '#REF!' })
  assert.equal(evaluateAt(columns.projected, workbook, 0, 0), 23)
  assert.deepEqual(evaluateAt(columns.projected, workbook, 0, 1), { kind: 'error', code: '#REF!' })
})

test('projected qualified ranges evaluate their two INDEX endpoints on the same sheet and propagate overflow', () => {
  const workbook = sample(), entry = fixture('qualified range inherits its sheet at both dynamic endpoints')
  assert.deepEqual(evaluateAt(entry.projected, workbook, 0, 0), { kind: 'error', code: '#REF!' })
  assert.equal(evaluateAt(entry.projected, workbook, 1, 0), 14)
  assert.equal(evaluateAt(entry.projected, workbook, 2, 0), 9)
})

test('reference-returning INDEX range endpoints preserve ordinary scalar and whole-axis INDEX behavior', () => {
  const workbook = sample()
  assert.equal(evaluateAt('=SUM(INDEX($A$1:$XFD$1048576,1,1):INDEX($A$1:$XFD$1048576,2,2))', workbook, 0, 0), 16)
  assert.equal(evaluateAt('=INDEX(A1:B2,2,2)', workbook, 0, 0), 4)
  assert.equal(evaluateAt('=SUM(INDEX(A1:B2,0,2))', workbook, 0, 0), 7)
  assert.equal(evaluateAt('=SUM(A1:INDEX(A1:B2,2,2))', workbook, 0, 0), 16)
  assert.equal(evaluateAt('=SUM(INDEX(A1:B2,1,1):B2)', workbook, 0, 0), 16)
})

test('error tokens remain errors while quoted error-looking strings and digit-suffixed functions remain text and functions', () => {
  const workbook = sample()
  assert.deepEqual(parseFormula('=#REF!'), { kind: 'error', code: '#REF!' })
  assert.deepEqual(parseFormula('="\u0023REF!"'), { kind: 'literal', value: '#REF!' })
  assert.equal(evaluateAt('=IFERROR(#REF!,99)', workbook, 0, 0), 99)
  assert.equal(evaluateAt('=IFERROR("#REF!",99)', workbook, 0, 0), '#REF!')
  assert.equal(parseFormula('=LOG10(A1)').kind, 'function')
})

test('dynamic INDEX endpoints propagate invalid source references through IFERROR', () => {
  const workbook = sample()
  assert.deepEqual(evaluateAt('=SUM(INDEX(Missing!A1:B2,1,1):B2)', workbook, 0, 0), { kind: 'error', code: '#REF!' })
  assert.equal(evaluateAt('=IFERROR(SUM(INDEX(Missing!A1:B2,1,1):B2),17)', workbook, 0, 0), 17)
})
