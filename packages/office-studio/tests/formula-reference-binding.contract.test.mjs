import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { parseFormula } from '../src/features/formula/formula-parser.ts'
import { bindFormulaReferences } from '../src/features/formula/formula-reference-binding.ts'
import { indexFormulaDefinedNames, readFormulaDefinedNames } from '../src/features/formula/formula-defined-names.ts'
import { evaluateFormula, createFormulaDependencyGraph, recalculateWorkbookFormulas } from '../src/features/formula/formula-engine.ts'
import { mapFormulaIdentifiers } from '../src/features/formula/formula-identifiers.ts'
import { rewriteFormulaImportCell } from '../src/features/formula/formula-import-context.ts'
import { rewriteWorkbookWorksheetReferences } from '../src/features/sheet-workbench/sheet-worksheet-references.ts'

// Deferred model and calculation drafts. No runner, type check, build or browser was started.
function sample() {
  const workbook = createBlankWorkbook(), first = workbook.sheets[0], other = createBlankSheet("Other's data")
  workbook.id = 'names-book'; first.id = 'first'; first.name = 'Data'; other.id = 'other'; workbook.activeSheetId = first.id
  for (const [row, col, value] of [[0, 0, 2], [1, 0, 3], [2, 0, 5], [0, 1, 7], [1, 1, 11], [2, 1, 13]]) first.cells[`${row}:${col}`] = createCell(row, col, value)
  other.cells['0:0'] = createCell(0, 0, 17); workbook.sheets.push(other)
  return { workbook, first, other }
}
const error = code => ({ kind: 'error', code })
const value = (workbook, sheet, formula, formulaCell) => evaluateFormula(formula, workbook, sheet.id, { formulaCell, now: new Date('2024-01-02T12:00:00Z') })

test('native index scopes and public string scopes resolve before global names, including Unicode identifiers', () => {
  const { workbook, first, other } = sample()
  workbook.metadata = { definedNames: [{ name: 'Rate', refersTo: '=2' }, { name: 'Rate', refersTo: '=A1', localSheetId: 1 },
    { name: '税率', formulaOrRefString: '=3', localSheetId: 'first' }] }
  assert.equal(value(workbook, first, '=Rate*税率'), 6)
  assert.equal(value(workbook, other, '=Rate'), 17)
  assert.equal(value(workbook, first, "='Other''s data'!Rate"), 17)
  assert.equal(readFormulaDefinedNames(workbook).find(name => name.name === '税率').owner, first.id)
})

test('missing explicit local scope cannot bind to an active sheet or a destination global name', () => {
  const { workbook, first } = sample()
  workbook.metadata = { definedNames: [{ name: 'Lost', refersTo: '=1', localSheetId: 'closed' }] }
  assert.equal(value(workbook, first, '=Lost').code, '#NAME?')
  first.metadata = { definedNames: [{ name: 'Global', refersTo: '=5', localSheetId: 'AllDefaultWorkbook' }] }
  assert.equal(readFormulaDefinedNames(workbook).find(name => name.name === 'Global').owner, undefined)
})

test('long linear name chains use an explicit compiler stack while a cycle retains its error identity', () => {
  const { workbook, first } = sample()
  workbook.metadata = { definedNames: Array.from({ length: 2048 }, (_, index) => ({ name: `Name_${index}`,
    refersTo: index === 2047 ? '=A1' : `=Name_${index + 1}` })) }
  assert.equal(value(workbook, first, '=Name_0+Name_0'), 4)
  workbook.metadata.definedNames[2047].refersTo = '=Name_0'
  assert.deepEqual(value(workbook, first, '=Name_0'), error('#CIRCULAR!'))
})

test('duplicate name expressions are tolerated only when the scope and expression agree', () => {
  const { workbook, first } = sample()
  workbook.metadata = { definedNames: [{ name: 'Tax', refersTo: '=A1', localSheetId: first.id }] }
  first.metadata = { localDefinedNames: [{ name: 'tax', refersTo: '=A1' }] }
  assert.equal(indexFormulaDefinedNames(workbook).size, 1)
  first.metadata.localDefinedNames[0].refersTo = '=B1'
  assert.throws(() => indexFormulaDefinedNames(workbook), /作用域/)
})

function table(workbook, first) {
  workbook.tables = [{ id: 'sales', name: 'Sales', sheetId: first.id, range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 },
    showHeaderRow: true, showTotalsRow: false, styleName: 'TableStyleMedium2', columns: [{ id: 'a', name: 'A:B' }, { id: 'b', name: 'Amount' }] }]
  return workbook.tables[0]
}

test('structured selectors follow headers, data, totals and table column identities after a range change', () => {
  const { workbook, first } = sample(), sales = table(workbook, first)
  assert.equal(value(workbook, first, '=SUM(Sales[Amount])'), 24)
  assert.equal(value(workbook, first, '=SUM(Sales[[#Headers],[Amount]])'), 7)
  assert.equal(value(workbook, first, '=Sales[@Amount]', { row: 1, column: 0 }), 11)
  assert.equal(value(workbook, first, '=SUM([A:B])', { row: 1, column: 1 }), 8)
  sales.showTotalsRow = true
  assert.equal(value(workbook, first, '=SUM(Sales[Amount])'), 11)
  assert.equal(value(workbook, first, '=SUM(Sales[[#Totals],[Amount]])'), 13)
})

test('escaped bracket and error-shaped column captions are preserved by both binding and identifier rewrites', () => {
  const { workbook, first } = sample(), sales = table(workbook, first)
  sales.columns[0].name = 'X]Name'; sales.columns[1].name = '#REF!'
  const formula = "=SUM(Sales[X']Name])+SUM(Sales['#REF!])"
  assert.equal(value(workbook, first, formula), 32)
  assert.equal(mapFormulaIdentifiers(formula, identifier => identifier.kind === 'function' ? undefined : identifier.kind === 'table' ? 'CopiedSales' : 'Incorrect'),
    "=SUM(CopiedSales[X']Name])+SUM(CopiedSales['#REF!])")
})

test('named whole-axis ranges stay sparse and contribute their real formula dependencies', () => {
  const { workbook, first } = sample()
  workbook.metadata = { definedNames: [{ name: 'ColumnData', refersTo: '=A:A' }] }
  first.cells['3:3'] = { ...createCell(3, 3, null), formula: '=SUM(ColumnData)' }
  const graph = createFormulaDependencyGraph(workbook), dependencies = graph.dependencies.get('first!3:3')
  assert.equal(dependencies.length, 1); assert.equal(dependencies[0].end.row, 1_048_575)
  assert.equal(value(workbook, first, '=SUM(ColumnData)'), 10)
})

test('reference-returning names compose with the existing range operator and INDEX path', () => {
  const { workbook, first } = sample()
  workbook.metadata = { definedNames: [{ name: 'Start', refersTo: '=A1' }, { name: 'End', refersTo: '=B2' }] }
  assert.equal(value(workbook, first, '=SUM(Start:End)'), 23)
  assert.equal(value(workbook, first, '=SUM(INDEX(Start:End,0,2))'), 18)
  assert.equal(bindFormulaReferences(parseFormula('=Start'), workbook, first, indexFormulaDefinedNames(workbook)).reference.sheetName, first.name)
})

test('INDIRECT resolves A1, names, tables and relative R1C1 without executing ref_text expressions', () => {
  const { workbook, first } = sample(); table(workbook, first)
  workbook.metadata = { definedNames: [{ name: 'Top', refersTo: '=A1' }] }
  for (const [formula, expected] of [['=INDIRECT("A2")', 3], ['=INDIRECT("R[-1]C[-1]",FALSE)', 2],
    ['=SUM(INDIRECT("A:A"))', 10], ['=INDIRECT("Top")', 2], ['=SUM(INDIRECT("Sales[Amount]"))', 24]]) {
    assert.equal(value(workbook, first, formula, { row: 1, column: 1 }), expected, formula)
  }
  for (const formula of ['=INDIRECT("1+2")', '=INDIRECT("A0")', '=INDIRECT("R[-2]C",FALSE)', '=INDIRECT("=A1")']) {
    assert.deepEqual(value(workbook, first, formula, { row: 1, column: 1 }), error('#REF!'), formula)
  }
})

test('OFFSET preserves a reference for aggregates and rejects zero size and out-of-workbook bounds', () => {
  const { workbook, first } = sample()
  assert.equal(value(workbook, first, '=SUM(OFFSET(A1,1,0,2,2))'), 32)
  assert.equal(value(workbook, first, '=SUM(OFFSET(B3,0,0,-2,-2))'), 32)
  for (const formula of ['=OFFSET(A1,-1,0)', '=OFFSET(A1,0,0,0,1)', '=OFFSET(A1,1048576,0)']) {
    assert.deepEqual(value(workbook, first, formula), error('#REF!'), formula)
  }
})

test('retained source calendar runs DATE and serial constants once and converts date references through the same resolver', () => {
  const { workbook, first } = sample(); workbook.date1904 = false
  first.cells['1:0'] = { ...createCell(1, 0, 43830), formula: '=DATE(2024,1,1)', style: { ...first.cells['1:0'].style, numberFormat: 'yyyy-mm-dd' },
    metadata: { chartDependencyFormulaSource: { version: 1, formula: '=DATE(2024,1,1)', date1904: true, dateResult: true } } }
  first.cells['1:1'] = { ...createCell(1, 1, null), formula: '=YEAR(A2)+DAY(0)+40000',
    metadata: { chartDependencyFormulaSource: { version: 1, formula: '=YEAR(A2)+DAY(0)+40000', date1904: true, dateResult: false } } }
  const recalculated = recalculateWorkbookFormulas(workbook, undefined, { now: new Date('2024-01-02T12:00:00Z') }).workbook
  assert.equal(recalculated.sheets[0].cells['1:0'].value, 45292)
  assert.equal(recalculated.sheets[0].cells['1:1'].value, 42025)
})

test('automatic rewrites retain a valid calendar marker but never revive one after a manual formula change', () => {
  const cell = { formula: '=A1+1', metadata: { chartDependencyFormulaSource: { version: 1, formula: '=A1+1', date1904: true, dateResult: false } } }
  rewriteFormulaImportCell(cell, '=A2+1'); assert.equal(cell.metadata.chartDependencyFormulaSource.formula, '=A2+1')
  cell.formula = '=A2+2'; rewriteFormulaImportCell(cell, '=A3+2')
  assert.equal(cell.metadata.chartDependencyFormulaSource.formula, '=A2+1')
})

test('worksheet rename updates imported formula ownership and name alias scopes in one detached candidate', () => {
  const { workbook, first, other } = sample()
  first.cells['3:0'] = { ...createCell(3, 0, null), formula: "='Other''s data'!A1",
    metadata: { chartDependencyFormulaSource: { version: 1, formula: "='Other''s data'!A1", date1904: true, dateResult: false } } }
  first.metadata = { chartDependencyImport: { formulaAliases: { sheets: { original: other.name }, names: { global: 'Global_2' }, tables: {}, nameSheets: { global: other.name } } } }
  const next = cloneModel(workbook); next.sheets[1].name = 'Renamed'
  rewriteWorkbookWorksheetReferences(next, workbook, { sheetId: other.id, oldName: other.name, newName: 'Renamed' })
  assert.equal(next.sheets[0].cells['3:0'].metadata.chartDependencyFormulaSource.formula, "='Renamed'!A1")
  assert.equal(next.sheets[0].metadata.chartDependencyImport.formulaAliases.nameSheets.global, 'Renamed')
  assert.equal(first.metadata.chartDependencyImport.formulaAliases.nameSheets.global, other.name)
})
