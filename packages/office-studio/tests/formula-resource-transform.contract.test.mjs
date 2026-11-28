import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createBlankSheet, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { readFormulaDefinedNames } from '../src/features/formula/formula-defined-names.ts'
import { evaluateFormula } from '../src/features/formula/formula-engine.ts'
import { readFormulaImportSource, readFormulaCalendarValue } from '../src/features/formula/formula-import-context.ts'
import { planWorkbookAxisStructureCommand } from '../src/features/sheet-workbench/sheet-axis-structure-command.ts'
import { transformWorkbookCellShiftResources } from '../src/features/sheet-workbench/sheet-cell-shift-resources.ts'
import { planWorksheetCopy } from '../src/features/sheet-workbench/sheet-worksheet-copy.ts'
import { moveSheetClipboardCellReferences, moveSheetClipboardDefinedNames } from '../src/features/sheet-workbench/sheet-clipboard-references.ts'
import { pasteSheetCell } from '../src/features/sheet-workbench/sheet-paste.ts'

// Deferred contracts: no runner, type check, build, WASM or browser was started.
const range = (startRow, endRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
function fixture() {
  const workbook = createBlankWorkbook(), first = workbook.sheets[0], other = createBlankSheet('Other')
  workbook.id = 'transform-book'; first.id = 'data'; first.name = 'Data'; other.id = 'other'; workbook.activeSheetId = first.id
  first.rowCount = 100; first.columnCount = 20; other.rowCount = 100; other.columnCount = 20; workbook.sheets.push(other)
  first.cells = { '0:0': createCell(0, 0, 2), '1:0': createCell(1, 0, 3),
    '2:1': { ...createCell(2, 1, 5), formula: '=SUM(A1:A2)' } }
  workbook.metadata = { definedNames: [{ name: 'GlobalData', refersTo: '=Data!A1:A2', formulaOrRefString: '=Data!A1:A2' },
    { name: 'LocalData', formulaOrRefString: '=A1:A2', localSheetId: first.id },
    { name: 'OtherData', formulaOrRefString: '=A1', localSheetId: 1 }] }
  other.cells['0:0'] = createCell(0, 0, 11)
  return { workbook, first, other }
}
function imported(cell, owner) {
  cell.metadata = { chartDependencyFormulaSource: { version: 1, formula: cell.formula, date1904: false, dateResult: false, aliasSheetId: owner },
    chartDependencyDateSource: { version: 1, date1904: false, value: 60.5, targetDate1904: true, targetValue: cell.value, targetFormula: cell.formula } }
  return cell
}
const name = (workbook, caption, owner) => readFormulaDefinedNames(workbook).find(item => item.name === caption && item.owner === owner)
const result = (workbook, sheetId, formula) => evaluateFormula(formula, workbook, sheetId)

test('axis insertion transforms public/native name expressions in their true scopes and preserves exact calendar ownership', () => {
  const { workbook, first, other } = fixture(), before = JSON.stringify(workbook)
  imported(first.cells['2:1'], first.id)
  const source = JSON.stringify(workbook), next = planWorkbookAxisStructureCommand(workbook, first.id, [{ kind: 'insert', axis: 'row', start: 0, count: 1 }]).workbook
  assert.equal(JSON.stringify(workbook), source); assert.notEqual(source, before)
  assert.equal(result(next, first.id, '=SUM(GlobalData)'), 5); assert.equal(result(next, first.id, '=SUM(LocalData)'), 5)
  assert.equal(name(next, 'GlobalData').expression, '=Data!A2:A3')
  assert.equal(name(next, 'GlobalData').record.refersTo, name(next, 'GlobalData').record.formulaOrRefString)
  assert.equal(name(next, 'LocalData', first.id).expression, '=A2:A3'); assert.equal(name(next, 'OtherData', other.id).expression, '=A1')
  const moved = next.sheets[0].cells['3:1']
  assert.equal(moved.formula, '=SUM(A2:A3)'); assert.equal(readFormulaImportSource(moved).formula, moved.formula)
  assert.equal(moved.metadata.chartDependencyDateSource.targetFormula, moved.formula)
  assert.equal(readFormulaCalendarValue(moved, moved.value, true).value, 60.5)
})

test('deleting the referent preserves a real REF error and does not rebind a missing explicit name scope', () => {
  const { workbook, first } = fixture()
  workbook.metadata.definedNames.push({ name: 'Removed', refersTo: '=Data!A1' }, { name: 'Missing', formulaOrRefString: '=A1', localSheetId: 'closed' })
  const next = planWorkbookAxisStructureCommand(workbook, first.id, [{ kind: 'delete', axis: 'row', start: 0, count: 1 }]).workbook
  assert.equal(result(next, first.id, '=Removed').code, '#REF!')
  assert.equal(name(next, 'Missing', 'missing-scope:closed').expression, '=A1')
  assert.equal(result(next, first.id, '=Missing').code, '#NAME?')
})

test('local cell shifts update cross-sheet and public name references only inside the perpendicular band', () => {
  const { workbook, first, other } = fixture()
  workbook.metadata.definedNames.push({ name: 'Outside', formulaOrRefString: '=Data!B1' })
  other.cells['1:1'] = imported({ ...createCell(1, 1, 5), formula: '=SUM(Data!A1:A2)' }, first.id)
  const next = cloneModel(workbook), target = next.sheets[0]
  target.cells['1:0'] = { ...target.cells['0:0'], row: 1, id: '1:0' }; target.cells['2:0'] = { ...workbook.sheets[0].cells['1:0'], row: 2, id: '2:0' }; delete target.cells['0:0']
  transformWorkbookCellShiftResources(next, workbook, first.id, { kind: 'insert', axis: 'row', start: 0, count: 1, bandStart: 0, bandEnd: 0 })
  assert.equal(result(next, first.id, '=SUM(LocalData)'), 5); assert.equal(result(next, other.id, '=SUM(GlobalData)'), 5)
  assert.equal(name(next, 'Outside').expression, '=Data!B1')
  const formula = next.sheets[1].cells['1:1']
  assert.equal(formula.formula, '=SUM(Data!A2:A3)'); assert.equal(readFormulaImportSource(formula).aliasSheetId, first.id)
  assert.equal(formula.metadata.chartDependencyDateSource.targetFormula, formula.formula)
})

test('worksheet copies duplicate local string and numeric scopes, keep globals and remap imported alias owners', () => {
  const { workbook, first, other } = fixture()
  first.metadata = { chartDependencyImport: { formulaAliases: { sheets: { data: 'Data' }, names: {}, tables: {} } } }
  imported(first.cells['2:1'], first.id)
  const before = JSON.stringify(workbook), next = planWorksheetCopy(workbook, first.id, { userId: 'local-user', targetSheetId: 'copy', targetSheetName: 'Data Copy' }).workbook
  assert.equal(JSON.stringify(workbook), before)
  assert.equal(name(next, 'LocalData', 'copy').expression, '=A1:A2'); assert.equal(name(next, 'LocalData', first.id).expression, '=A1:A2')
  assert.equal(name(next, 'OtherData', other.id).record.localSheetId, 2)
  assert.equal(readFormulaDefinedNames(next).filter(item => item.name === 'GlobalData').length, 1)
  assert.equal(result(next, 'copy', '=SUM(LocalData)'), 5)
  const copied = next.sheets.find(item => item.id === 'copy'), formula = copied.cells['2:1']
  assert.equal(readFormulaImportSource(formula).aliasSheetId, copied.id)
  assert.equal(copied.metadata.chartDependencyImport.formulaAliases.sheets.data, 'Data Copy')
  assert.equal(readFormulaCalendarValue(formula, formula.value, true).value, 60.5)
})

test('worksheet copies remap structured table identifiers without changing literal text or escaped column captions', () => {
  const { workbook, first } = fixture()
  workbook.tables = [{ id: 'sales', name: 'Sales', sheetId: first.id, range: range(0, 1, 0, 1), showHeaderRow: true,
    showTotalsRow: false, styleName: 'TableStyleMedium2', columns: [{ id: 'label', name: 'Label' }, { id: 'value', name: "X]Value" }] }]
  first.cells['0:1'] = createCell(0, 1, 'X]Value'); first.cells['1:1'] = createCell(1, 1, 7)
  first.cells['3:1'] = { ...createCell(3, 1, null), formula: "=SUM(Sales[X']Value])+LEN(\"Sales[X']Value]\")" }
  workbook.metadata.definedNames.push({ name: 'TableData', formulaOrRefString: "=Sales[X']Value]", localSheetId: first.id })
  const next = planWorksheetCopy(workbook, first.id, { userId: 'local-user', targetSheetId: 'copy', targetSheetName: 'Copied Data' }).workbook
  const copiedTable = next.tables.find(item => item.sheetId === 'copy'), copied = next.sheets.find(item => item.id === 'copy')
  assert.notEqual(copiedTable.name, 'Sales'); assert.notEqual(copiedTable.id, 'sales')
  assert.equal(name(next, 'TableData', 'copy').expression, `=${copiedTable.name}[X']Value]`)
  assert.equal(copied.cells['3:1'].formula, `=SUM(${copiedTable.name}[X']Value])+LEN("Sales[X']Value]")`)
  assert.equal(result(next, 'copy', '=SUM(TableData)'), 7)
})

test('cross-sheet cuts relocate names and date-marked formulas to the actual referents without mutating the source DTO', () => {
  const { workbook, first, other } = fixture(), source = imported(first.cells['2:1'], first.id), before = JSON.stringify(source)
  const move = { sourceSheetId: first.id, targetSheetId: other.id, sourceSheetName: first.name, targetSheetName: other.name,
    sourceRange: range(0, 1), targetRange: range(2, 3, 2) }
  const moved = moveSheetClipboardCellReferences(source, first.name, first.name, workbook, move), next = cloneModel(workbook)
  next.metadata = moveSheetClipboardDefinedNames(workbook.metadata, workbook, move)
  delete next.sheets[0].cells['0:0']; delete next.sheets[0].cells['1:0']; next.sheets[0].cells['2:1'] = moved
  next.sheets[1].cells['2:2'] = createCell(2, 2, 2); next.sheets[1].cells['3:2'] = createCell(3, 2, 3)
  assert.equal(JSON.stringify(source), before); assert.equal(moved.formula, '=SUM(Other!C3:C4)')
  assert.equal(result(next, first.id, moved.formula), 5); assert.equal(result(next, first.id, '=SUM(GlobalData)'), 5)
  assert.equal(result(next, first.id, '=SUM(LocalData)'), 5)
  assert.equal(readFormulaImportSource(moved).formula, moved.formula)
  assert.equal(moved.metadata.chartDependencyDateSource.targetFormula, moved.formula)
})

test('formula copy translates relative operands and preserves owned provenance while stale markers stay invalid', () => {
  const source = { value: '5', rawValue: 5, formula: '=SUM(A1:$A$2)', metadata: {
    chartDependencyFormulaSource: { version: 1, formula: '=SUM(A1:$A$2)', date1904: false, dateResult: false },
    chartDependencyDateSource: { version: 1, date1904: false, value: 60.5, targetDate1904: true, targetValue: 5, targetFormula: '=SUM(A1:$A$2)' } } }
  const before = JSON.stringify(source), pasted = pasteSheetCell(source, { value: '' }, 'all', 2, 1)
  assert.equal(JSON.stringify(source), before); assert.equal(pasted.formula, '=SUM(B3:$A$2)')
  assert.equal(readFormulaImportSource(pasted).formula, pasted.formula)
  assert.equal(pasted.metadata.chartDependencyDateSource.targetFormula, pasted.formula)
  const stale = { ...source, formula: '=SUM(A1:A2)' }, invalid = pasteSheetCell(stale, { value: '' }, 'all', 2, 1)
  assert.equal(readFormulaImportSource(invalid), undefined)
  assert.equal(invalid.metadata.chartDependencyDateSource.targetFormula, source.formula)
})
