import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { applyWorkbookFormulaCalendarSource, normalizeWorkbookFormulaCalendarSource } from '../src/parser/workbookFormulaCalendar.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'

// Deferred parser-boundary contracts. Native Office recalculation is a separate
// later interoperability test; these sources have already been native-admitted.
function sample() {
  const workbook = createBlankWorkbook()
  workbook.id = 'new-workbook'; workbook.date1904 = true; workbook.calculationMode = 'manual'
  const other = { ...createBlankSheet('Other'), id: 'new-other' }
  const helper = { ...createBlankSheet('Calc'), id: 'new-calc', visible: false, metadata: { hiddenState: 'veryHidden' } }
  const author = { ...createBlankSheet('Data'), id: 'new-data' }
  author.cells['0:0'] = { ...createCell(0, 0, 29), formula: "='Calc'!A1" }
  helper.cells['0:0'] = { ...createCell(0, 0, 29), formula: '=DAY(A2)' }
  workbook.sheets = [other, helper, author]; workbook.activeSheetId = helper.id
  workbook.metadata.localDefinedNames = [
    { name: 'DropHelper', refersTo: '=Calc!A1', localSheetId: 1 },
    { name: 'KeepAuthor', refersTo: '=Data!A1', localSheetId: 2 },
  ]
  workbook.tables = [{ id: 'native-table', officeTableId: 7, name: 'DataTable', sheetId: author.id,
    range: { startRow: 3, endRow: 4, startCol: 0, endCol: 0 }, columns: [{ id: 'column', name: 'Amount' }],
    showHeaderRow: true, showTotalsRow: false, styleName: 'TableStyleMedium2' }]
  workbook.slicers = [{ id: 'slicer', tableId: 'native-table', sourceFieldId: 'column', type: 'slicer' }]
  workbook.slicerCaches = [{ id: 'cache', source: { kind: 'table', subUnitId: author.id, tableId: 'native-table', columnId: 'column' } }]
  const source = { version: 1, date1904: true, sheets: [
    { name: 'Other', sheetId: 'old-other', officeSheetId: 2, cells: [] },
    { name: 'Data', sheetId: 'old-data', officeSheetId: 1, import: { version: 1,
      formulaAliases: { sheets: { original: 'Data' }, names: {}, tables: {} } }, cells: [
      { row: 0, col: 0, formula: '=DAY(A2)', projectedFormula: "='Calc'!A1", calculationSheet: 'Calc', calculationFormula: '=DAY(A2)',
        formulaSource: { version: 1, formula: '=DAY(A2)', date1904: false, dateResult: false, aliasSheetId: 'old-data' } },
    ] },
  ], helpers: [{ name: 'Calc', sheetId: 'old-calc', officeSheetId: 3, date1904: false,
    sourceSheetId: 'old-data', fingerprint: '0123456789abcdef', reclaim: true }],
  tables: [{ name: 'DataTable', tableId: 'stable-table', officeTableId: 7, sheetId: 'old-data' }] }
  const rawSheets = workbook.sheets.map(sheet => ({ name: sheet.name, cells: Object.values(sheet.cells).map(cell => ({
    row: cell.row, col: cell.col, formula: cell.formula, value: cell.value,
  })) }))
  const restore = () => applyWorkbookFormulaCalendarSource(workbook, rawSheets, [], normalizeWorkbookFormulaCalendarSource(source))
  return { workbook, author, helper, other, source, rawSheets, restore }
}

test('native reorder restores authors and aliases, remaps table consumers, and removes only admitted veryHidden helpers', () => {
  const h = sample(); h.restore()
  assert.equal(h.author.cells['0:0'].formula, '=DAY(A2)')
  assert.equal(h.author.cells['0:0'].metadata.chartDependencyFormulaSource.aliasSheetId, h.author.id)
  assert.deepEqual(h.workbook.sheets.map(sheet => sheet.id), [h.other.id, h.author.id])
  assert.equal(h.workbook.tables[0].id, 'stable-table')
  assert.equal(h.workbook.slicers[0].tableId, 'stable-table')
  assert.equal(h.workbook.slicerCaches[0].source.tableId, 'stable-table')
  assert.deepEqual(h.workbook.metadata.localDefinedNames, [{ name: 'KeepAuthor', refersTo: '=Data!A1', localSheetId: 1 }])
  assert.equal(h.workbook.activeSheetId, h.other.id)
  assert.equal(h.workbook.metadata.formulaCalendarImport.restoredCells, 1)
  assert.equal(h.workbook.metadata.formulaCalendarImport.reclaimedSheets, 1)
})

test.each(['author', 'helper'])('a late %s formula change remains native and retains the calculation family', target => {
  const h = sample(), index = target === 'author' ? 2 : 1
  h.rawSheets[index].cells[0].formula = '=42+1'
  h.workbook.sheets[index].cells['0:0'].formula = '=42+1'
  h.restore()
  assert.equal(h.workbook.sheets[index].cells['0:0'].formula, '=42+1')
  assert.equal(h.workbook.sheets.length, 3)
  assert.equal(h.workbook.metadata.formulaCalendarImport.unmatchedCells, 1)
  assert.equal(h.workbook.metadata.formulaCalendarImport.restoredCells, 0)
})

test('retained alias ownership in a calculation sheet keeps its live family after ID remapping', () => {
  const h = sample(); h.source.sheets[1].cells[0].formulaSource.aliasSheetId = 'old-calc'; h.restore()
  assert.equal(h.author.cells['0:0'].metadata.chartDependencyFormulaSource.aliasSheetId, h.helper.id)
  assert.ok(h.workbook.sheets.includes(h.helper))
  assert.equal(h.workbook.metadata.formulaCalendarImport.reclaimedSheets, 0)
})

test('an empty missing source sheet is still a failed identity match and prevents family reclamation', () => {
  const h = sample(); h.source.sheets[0].name = 'Missing'; h.restore()
  assert.equal(h.workbook.metadata.formulaCalendarImport.missingSheets, 1)
  assert.ok(h.workbook.sheets.includes(h.helper))
})

test('visible calculation sheets are retained despite an admitted reclaim bit', () => {
  const h = sample(); h.helper.visible = true; delete h.helper.metadata.hiddenState; h.restore()
  assert.ok(h.workbook.sheets.includes(h.helper))
})

test('exact serial provenance is restored only while native value and formula still match', () => {
  const h = sample(), saved = h.source.sheets[1].cells[0]
  saved.formulaSource.dateResult = true
  saved.dateSource = { value: 60.5, date1904: false, targetValue: -1401.5, targetDate1904: true, targetFormula: '=DAY(A2)' }
  h.author.cells['0:0'].value = -1401.5; h.rawSheets[2].cells[0].value = -1401.5; h.restore()
  assert.equal(h.author.cells['0:0'].metadata.chartDependencyDateSource.value, 60.5)
  assert.equal(h.author.cells['0:0'].metadata.chartDependencyDateSource.targetFormula, '=DAY(A2)')
  const changed = sample(); changed.source.sheets[1].cells[0].dateSource = saved.dateSource; changed.restore()
  assert.equal(changed.author.cells['0:0'].formula, "='Calc'!A1")
  assert.ok(changed.workbook.sheets.includes(changed.helper))
})

test('source normalization clones JSON extensions and rejects duplicate identities and unresolved alias owners', () => {
  const h = sample(), normalized = normalizeWorkbookFormulaCalendarSource(h.source)
  assert.deepEqual(normalized, h.source); assert.notEqual(normalized, h.source)
  h.source.sheets[1].import.formulaAliases.sheets.original = 'Changed'
  assert.equal(normalized.sheets[1].import.formulaAliases.sheets.original, 'Data')
  const duplicate = cloneModel(normalized); duplicate.helpers[0].sheetId = duplicate.sheets[0].sheetId
  assert.throws(() => normalizeWorkbookFormulaCalendarSource(duplicate), /FORMULA_CALENDAR_SOURCE_INVALID/)
  const unresolved = cloneModel(normalized); unresolved.sheets[1].cells[0].formulaSource.aliasSheetId = 'missing'
  assert.throws(() => normalizeWorkbookFormulaCalendarSource(unresolved), /FORMULA_CALENDAR_SOURCE_INVALID/)
})

test('extension bytes stop inspection at 16 MiB before later values or conversion hooks can run', () => {
  const h = sample(); let invoked = 0
  h.source.sheets[1].import.extension = 'x'.repeat(16 * 1024 * 1024)
  Object.defineProperty(h.source.sheets[1].import, 'late', { enumerable: true, get() { invoked++; return 1 } })
  assert.throws(() => normalizeWorkbookFormulaCalendarSource(h.source), /FORMULA_CALENDAR_SOURCE_LIMIT/)
  assert.equal(invoked, 0)
})

test.each(['accessor', 'conversion', 'cycle', 'nonfinite', 'sparse'])('%s extension is rejected before cloning or user code execution', kind => {
  const h = sample(); let invoked = 0
  if (kind === 'accessor') Object.defineProperty(h.source.sheets[1].import, 'extension', { enumerable: true, get() { invoked++; return 1 } })
  if (kind === 'conversion') h.source.sheets[1].import.extension = { toJSON() { invoked++; return 1 } }
  if (kind === 'cycle') h.source.sheets[1].import.extension = h.source
  if (kind === 'nonfinite') h.source.sheets[1].import.extension = Number.NaN
  if (kind === 'sparse') h.source.sheets[1].import.extension = Array(100)
  assert.throws(() => normalizeWorkbookFormulaCalendarSource(h.source), /FORMULA_CALENDAR_SOURCE_INVALID/)
  assert.equal(invoked, 0)
})

test('the complete workbook normalizer retains native table/column IDs and source XML beside editable data', () => {
  const xml = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="7" name="DataTable" ref="A1:A2"><tableColumns count="1"><tableColumn id="9" name="Amount"/></tableColumns></table>'
  const workbook = normalizeWorkbook({ format: 'xlsx', parserVersion: 'fixture', fileName: 'table.xlsx', payload: { data: {
    sheets: [{ name: 'Data', visible: 'visible', cells: [{ row: 0, col: 0, value: 'Amount', dataType: 'text' }, { row: 1, col: 0, value: 12, dataType: 'number' }] }],
    tables: [{ id: 'native-table', name: 'DataTable', sheetName: 'Data', sheetIndex: 0, officeTableId: 7, officeSheetId: 3,
      range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 }, columns: [{ id: 'column', name: 'Amount', officeColumnId: 9 }],
      officeSource: { partName: 'xl/tables/table7.xml', xml } }],
  } } })
  assert.equal(workbook.tables[0].officeTableId, 7); assert.equal(workbook.tables[0].officeSheetId, 3)
  assert.equal(workbook.tables[0].columns[0].officeColumnId, 9)
  assert.equal(workbook.tables[0].officeSource.xml, xml)
  assert.equal(workbook.sheets[0].cells['1:0'].value, 12)
})
