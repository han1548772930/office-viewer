import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createBlankSheet, createCell } from '../src/domain/model.ts'
import { evaluateFormula, evaluateFormulaReference, recalculateWorkbookFormulas } from '../src/features/formula/formula-engine.ts'

// Deferred behavior contracts. The runner and browser remain stopped.
function sample() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], privateSheet = createBlankSheet('Private')
  sheet.id = 'public'; sheet.name = 'Data'; privateSheet.id = 'private'
  workbook.id = 'read-book'; workbook.activeSheetId = sheet.id; workbook.sheets.push(privateSheet)
  sheet.cells['0:0'] = createCell(0, 0, 7); sheet.cells['0:1'] = createCell(0, 1, 11)
  privateSheet.cells['0:0'] = createCell(0, 0, 999)
  const reads = []
  const policy = { canReadCell: (source, row, column) => { reads.push([source.id, row, column]); return source.id !== privateSheet.id },
    canReadRange: source => source.id !== privateSheet.id }
  return { workbook, sheet, privateSheet, reads, policy }
}
const error = code => ({ kind: 'error', code })
const range = (startRow, endRow, startCol, endCol) => ({ startRow, endRow, startCol, endCol })
const reference = (sheet, bounds) => ({ kind: 'reference', sheetId: sheet.id, sheetName: sheet.name, range: bounds })

test('scalar source access and recursively evaluated formulas never use an inaccessible cached value', () => {
  const { workbook, sheet, policy } = sample()
  sheet.cells['1:0'] = { ...createCell(1, 0, 999), formula: '=Private!A1', cachedDataType: 'number' }
  assert.deepEqual(evaluateFormula('=A2+1', workbook, sheet.id, policy), error('#REF!'))
  assert.deepEqual(evaluateFormula('=Private!A1', workbook, sheet.id, policy), error('#REF!'))
  assert.equal(evaluateFormula('=A1+B1', workbook, sheet.id, policy), 18)
})

test('reference results preserve geometry for names, INDEX and a single empty cell without reading values', () => {
  const { workbook, sheet, policy } = sample()
  workbook.metadata = { definedNames: [{ name: 'PublicRange', refersTo: '=Data!A1:B2' }] }
  for (const [formula, bounds] of [['=PublicRange', range(0, 1, 0, 1)], ['=INDEX(PublicRange,0,2)', range(0, 1, 1, 1)],
    ['=INDEX(A1,1,1)', range(0, 0, 0, 0)], ['=C9', range(8, 8, 2, 2)]]) {
    assert.deepEqual(evaluateFormulaReference(formula, workbook, sheet.id, policy), reference(sheet, bounds), formula)
  }
  assert.equal(evaluateFormula('=INDEX(A1,1,1)', workbook, sheet.id, policy), 7)
  assert.equal(evaluateFormula('=INDEX(C9,1,1)', workbook, sheet.id, policy), null)
})

test('IF and CHOOSE retain the selected reference and do not read unselected private branches', () => {
  const { workbook, sheet, policy, reads } = sample()
  for (const formula of ['=IF(TRUE,A1:B2,Private!A1)', '=CHOOSE(1,A1:B2,Private!A1)', '=CHOOSE(1.9,A1:B2,1/0)']) {
    assert.deepEqual(evaluateFormulaReference(formula, workbook, sheet.id, policy), reference(sheet, range(0, 1, 0, 1)), formula)
  }
  assert.equal(evaluateFormula('=SUM(IF(TRUE,A1:B2,Private!A1))', workbook, sheet.id, policy), 18)
  assert.ok(reads.every(read => read[0] === sheet.id))
  assert.deepEqual(evaluateFormulaReference('=IF(FALSE,A1:B2,Private!A1)', workbook, sheet.id, policy), error('#REF!'))
})

test('INDIRECT and OFFSET share the same source range gate as static references', () => {
  const { workbook, sheet, policy } = sample()
  assert.deepEqual(evaluateFormulaReference('=INDIRECT("A1:B2")', workbook, sheet.id, policy), reference(sheet, range(0, 1, 0, 1)))
  assert.deepEqual(evaluateFormulaReference('=OFFSET(A1,1,1,2,1)', workbook, sheet.id, policy), reference(sheet, range(1, 2, 1, 1)))
  assert.deepEqual(evaluateFormula('=SUM(INDIRECT("Private!A1:A2"))', workbook, sheet.id, policy), error('#REF!'))
  assert.deepEqual(evaluateFormulaReference('=OFFSET(Private!A1,0,0)', workbook, sheet.id, policy), error('#REF!'))
})

test('a range policy covers protected blank cells in whole-axis queries without dense per-cell calls', () => {
  const { workbook, sheet } = sample(), calls = []
  const options = { canReadCell: () => { throw new Error('No data read is needed') }, canReadRange: (source, bounds) => {
    calls.push(bounds); return source.id === sheet.id && bounds.endRow < 100
  } }
  assert.deepEqual(evaluateFormulaReference('=A:A', workbook, sheet.id, options), error('#REF!'))
  assert.deepEqual(evaluateFormula('=SUM(A:A)', workbook, sheet.id, options), error('#REF!'))
  assert.ok(calls.length < 10); assert.equal(calls[0].endRow, 1_048_575)
})

test('a cell-only policy proves bounded references and refuses an unbounded permission expansion', () => {
  const { workbook, sheet } = sample()
  const options = { canReadCell: (_source, row, column) => row !== 1 || column !== 1 }
  assert.deepEqual(evaluateFormulaReference('=A1:B2', workbook, sheet.id, options), error('#REF!'))
  assert.deepEqual(evaluateFormulaReference('=A:A', workbook, sheet.id, options), error('#REF!'))
  assert.deepEqual(evaluateFormulaReference('=A1:B1', workbook, sheet.id, options), reference(sheet, range(0, 0, 0, 1)))
})

test('a merged source requires View on its entire geometry, including protected lower and right subcells', () => {
  const { workbook, sheet } = sample()
  sheet.mergedRanges = [range(0, 1, 0, 1)]
  const options = { canReadRange: (_source, bounds) => bounds.endRow === 0 && bounds.endCol === 0 }
  assert.deepEqual(evaluateFormula('=A1', workbook, sheet.id, options), error('#REF!'))
  assert.deepEqual(evaluateFormulaReference('=A1', workbook, sheet.id, options), error('#REF!'))
})

test('structured references check the resolved table data range and preserve selectors', () => {
  const { workbook, sheet, policy } = sample()
  workbook.tables = [{ id: 'sales', name: 'Sales', sheetId: sheet.id, range: range(0, 2, 0, 1), showHeaderRow: true,
    showTotalsRow: false, styleName: 'TableStyleMedium2', columns: [{ id: 'label', name: 'Label' }, { id: 'value', name: 'Value' }] }]
  assert.deepEqual(evaluateFormulaReference('=Sales[Value]', workbook, sheet.id, policy), reference(sheet, range(1, 2, 1, 1)))
  assert.deepEqual(evaluateFormulaReference('=Sales[Value]', workbook, sheet.id, { canReadRange: () => false }), error('#REF!'))
})

test('imported INDIRECT aliases do not bypass a denied destination source', () => {
  const { workbook, sheet, privateSheet, policy } = sample()
  const formula = '=INDIRECT("Original!A1")'
  sheet.cells['1:0'] = { ...createCell(1, 0, 999), formula,
    metadata: { chartDependencyFormulaSource: { version: 1, formula, date1904: false, dateResult: false, aliasSheetId: sheet.id } } }
  sheet.metadata = { chartDependencyImport: { formulaAliases: { sheets: { original: privateSheet.name }, names: {}, tables: {} } } }
  assert.deepEqual(evaluateFormula(formula, workbook, sheet.id, { ...policy, formulaCell: { row: 1, column: 0 } }), error('#REF!'))
})

test('permissioned recalculation preserves the caller and does not revive unsupported cached outputs', () => {
  const { workbook, sheet, policy } = sample()
  sheet.cells['1:0'] = { ...createCell(1, 0, 999), formula: '=UNIMPLEMENTED()', cachedDataType: 'number' }
  const before = JSON.stringify(workbook)
  const result = recalculateWorkbookFormulas(workbook, workbook, policy)
  assert.equal(result.workbook.sheets[0].cells['1:0'].value, '#NAME?')
  assert.equal(JSON.stringify(workbook), before)
})

test('invalid reference expressions, bounds and throwing permission providers return explicit errors', () => {
  const { workbook, sheet } = sample()
  assert.deepEqual(evaluateFormulaReference('=1+2', workbook, sheet.id), error('#VALUE!'))
  assert.deepEqual(evaluateFormulaReference('=#REF!', workbook, sheet.id), error('#REF!'))
  assert.deepEqual(evaluateFormulaReference('=A1', workbook, sheet.id, { canReadRange: () => { throw new Error('Revoked') } }), error('#REF!'))
  assert.deepEqual(evaluateFormula('=A1', workbook, sheet.id, { formulaCell: { row: -1, column: 0 } }), error('#REF!'))
})
