import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { restoreWorkbookTableIdentities } from '../src/parser/workbookTableIdentity.ts'
import { normalizeWorkbookFormulaCalendarSource } from '../src/parser/workbookFormulaCalendar.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'

// Deferred contracts for native IDs, stable domain IDs and their consumers.
function table(id = 'native-table', officeTableId = 31, name = 'Orders') {
  return { id, officeTableId, name, sheetId: 'sheet', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 },
    columns: [{ id: 'native-b', officeColumnId: 9, name: 'Amount' }, { id: 'native-a', officeColumnId: 7, name: 'Product' }],
    showHeaderRow: true, showTotalsRow: false, styleName: 'TableStyleMedium2' }
}
function sample() {
  const workbook = createBlankWorkbook(); workbook.sheets[0].id = 'sheet'; workbook.activeSheetId = 'sheet'
  workbook.tables = [table()]
  workbook.slicers = [{ id: 'slicer', type: 'slicer', tableId: 'native-table', sourceFieldId: 'native-b' }]
  workbook.slicerCaches = [{ id: 'cache', source: { kind: 'table', subUnitId: 'sheet', tableId: 'native-table', columnId: 'native-a' } }]
  const sources = [{ tableId: 'stable-table', officeTableId: 31, name: 'Orders', sheetId: 'original-sheet', columns: [
    { columnId: 'stable-a', officeColumnId: 7 }, { columnId: 'stable-b', officeColumnId: 9 },
  ] }]
  const restore = () => restoreWorkbookTableIdentities(workbook, [], sources, new Map([['original-sheet', 'sheet']]))
  return { workbook, sources, restore }
}

test('native column reordering restores identities and updates slicer and cache fields together', () => {
  const h = sample(); h.restore()
  assert.equal(h.workbook.tables[0].id, 'stable-table')
  assert.deepEqual(h.workbook.tables[0].columns.map(column => column.id), ['stable-b', 'stable-a'])
  assert.equal(h.workbook.slicers[0].tableId, 'stable-table'); assert.equal(h.workbook.slicers[0].sourceFieldId, 'stable-b')
  assert.equal(h.workbook.slicerCaches[0].source.tableId, 'stable-table'); assert.equal(h.workbook.slicerCaches[0].source.columnId, 'stable-a')
})

test('table manifests match native identity after Office reorders physical parts', () => {
  const h = sample(); const unrelated = table('foreign-table', 3, 'Foreign')
  h.workbook.tables.unshift(unrelated); h.restore()
  assert.equal(h.workbook.tables[0].id, 'foreign-table'); assert.equal(h.workbook.tables[1].id, 'stable-table')
})

test('foreign tables cannot consume a later restored domain identity', () => {
  const h = sample(); const foreign = table('stable-table', 6, 'Foreign')
  h.workbook.tables.unshift(foreign)
  h.workbook.slicers.push({ id: 'foreign-slicer', type: 'slicer', tableId: foreign.id, sourceFieldId: 'native-a' })
  h.restore()
  assert.equal(h.workbook.tables[1].id, 'stable-table'); assert.notEqual(foreign.id, 'stable-table')
  assert.equal(h.workbook.slicers[1].tableId, foreign.id)
  assert.equal(h.workbook.slicers[0].tableId, 'stable-table')
})

test('new native columns cannot steal the stable ID restored into another column', () => {
  const h = sample(); h.workbook.tables[0].columns.push({ id: 'stable-a', officeColumnId: 11, name: 'New' })
  h.workbook.slicers.push({ id: 'new-slicer', type: 'slicer', tableId: 'native-table', sourceFieldId: 'stable-a' })
  h.restore()
  const [b, a, added] = h.workbook.tables[0].columns
  assert.deepEqual([b.id, a.id], ['stable-b', 'stable-a']); assert.notEqual(added.id, a.id)
  assert.equal(h.workbook.slicers[1].sourceFieldId, added.id)
})

test('a table with a changed native owner is kept native instead of taking a saved identity', () => {
  const h = sample(); h.workbook.tables[0].sheetId = 'different-sheet'; h.restore()
  assert.equal(h.workbook.tables[0].id, 'native-table')
  assert.equal(h.workbook.slicers[0].sourceFieldId, 'native-b')
})

test('the legacy manifest without columns restores the table and keeps original field IDs', () => {
  const h = sample(); delete h.sources[0].columns; h.restore()
  assert.equal(h.workbook.tables[0].id, 'stable-table')
  assert.deepEqual(h.workbook.tables[0].columns.map(column => column.id), ['native-b', 'native-a'])
})

test('all column identity plans are admitted before tables or consumers are mutated', () => {
  const h = sample(), invalid = table('later', 5, 'Later')
  invalid.columns[1].id = invalid.columns[0].id; h.workbook.tables.push(invalid)
  const before = cloneModel(h.workbook)
  assert.throws(h.restore, /TABLE_SOURCE_INVALID/)
  assert.deepEqual(h.workbook, before)
})

test.each(['domain', 'native'])('duplicate %s column manifest identities are rejected at the DTO boundary', kind => {
  const h = sample(), columns = h.sources[0].columns
  if (kind === 'domain') columns[1].columnId = columns[0].columnId
  else columns[1].officeColumnId = columns[0].officeColumnId
  assert.throws(() => normalizeWorkbookFormulaCalendarSource({ version: 1, date1904: false, helpers: [],
    sheets: [{ name: 'Data', sheetId: 'original-sheet', officeSheetId: 1, cells: [] }], tables: h.sources }), /FORMULA_CALENDAR_SOURCE_INVALID/)
})

test('normalization retains formulas, null removals, style flags and original table XML', () => {
  const source = { partName: 'xl/tables/table31.xml', xml: '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="31" ref="A1:B4"><tableColumns count="2"><tableColumn id="9" name="Amount"/><tableColumn id="7" name="Product"/></tableColumns></table>' }
  const input = table(); input.columns[0] = { ...input.columns[0], formula: '=1+2', formulaIsArray: true,
    totalsRowFormula: null, totalsRowLabel: '合计', totalsRowFunction: 'sum', showFilterButton: false }
  Object.assign(input, { officeSource: source, autoFilter: null, showFirstColumn: true, showLastColumn: false, showRowStripes: false, showColumnStripes: true })
  const workbook = normalizeWorkbook({ format: 'xlsx', fileName: 'tables.xlsx', parserVersion: 'fixture', payload: { data: {
    sheets: [{ name: 'Data', visible: 'visible', cells: [] }], tables: [{ ...input, sheetIndex: 0 }],
  } } })
  const actual = workbook.tables[0]
  assert.equal(actual.autoFilter, null); assert.deepEqual(actual.officeSource, source)
  assert.equal(actual.showFirstColumn, true); assert.equal(actual.showLastColumn, false)
  assert.equal(actual.showRowStripes, false); assert.equal(actual.showColumnStripes, true)
  assert.deepEqual(actual.columns[0], input.columns[0])
})
