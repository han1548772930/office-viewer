import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, defaultCellStyle } from '../src/domain/model.ts'
import { createWorksheetCopyName, planWorksheetCopy, remapWorksheetCopyFormula } from '../src/features/sheet-workbench/sheet-worksheet-copy.ts'
import { copyWorksheetCommandInput } from '../src/api/sheet-lifecycle-commands.ts'

function fixture() {
  const workbook = createBlankWorkbook()
  workbook.id = 'copy-unit'
  const source = workbook.sheets[0]
  source.id = 'source'; source.name = '数据'; workbook.activeSheetId = source.id
  source.rowCount = 1_048_576; source.columnCount = 16_384
  source.rowHiddenSpans = [{ start: 50, end: 1_048_575, value: true }]
  source.columnWidthSpans = [{ start: 10, end: 16_383, value: 100 }]
  source.cells = {
    '0:0': { id: '0:0', row: 0, col: 0, value: 7, formula: null, display: '7', style: defaultCellStyle(), metadata: { unknown: { keep: true } } },
    '1:0': { id: '1:0', row: 1, col: 0, value: 7, formula: "=SUM('数据'!$A$1,Sales[金额])", display: '7', style: defaultCellStyle() },
    '10:0': { id: '10:0', row: 10, col: 0, value: 7, formula: null, display: '7', style: defaultCellStyle(), metadata: { pivotTableId: 'pivot', pivotGenerated: true } },
  }
  workbook.sheets.push({ ...structuredClone(source), id: 'other', name: '其他', cells: {} })
  workbook.tables = [{ id: 'table', name: 'Sales', sheetId: 'source', range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 },
    columns: [{ id: 'column', name: '金额' }], showHeaderRow: true, showTotalsRow: false, styleName: 'TableStyleMedium2' }]
  workbook.pivotTables = [{ id: 'pivot', source: { unitId: workbook.id, subUnitId: 'source', sheetName: '数据', range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 } },
    anchor: { unitId: workbook.id, subUnitId: 'source', row: 10, col: 0 }, fields: [{ id: 'amount', sourceColumn: 0, name: '金额', area: 'value', position: 0, aggregation: 'sum' }],
    options: { repeatRowLabels: false, showRowGrandTotal: true, showColGrandTotal: true, showRowSubtotals: true, showColSubtotals: true, valuePosition: 'column' },
    outputRange: { startRow: 10, startCol: 0, endRow: 10, endCol: 0 }, generatedCellKeys: ['10:0'],
    drillDownMap: { '10:0': { valueFieldId: 'amount', fieldFilters: { amount: 7 } } }, revision: 3 }]
  workbook.charts = [{ id: 'chart', sheetId: 'source', type: 'bar', source: { sheetId: 'source', range: workbook.pivotTables[0].outputRange },
    pivotTableId: 'pivot', position: { x: 100, y: 200 }, size: { width: 320, height: 200 }, revision: 2 }]
  workbook.slicers = [{ id: 'slicer', unitId: workbook.id, hostSheetId: 'source', tableId: 'table', sourceFieldId: 'column', type: 'slicer', caption: '金额',
    position: { row: 20, column: 2 }, size: { rows: 4, columns: 3 }, selectedItems: ['7'], revision: 2 }]
  workbook.metadata = { definedNames: [{ name: 'Local', sheetId: 'source', refersTo: "='数据'!$A$1" },
    { name: 'Other', localSheetId: 1, refersTo: "='其他'!$A$1" }, { name: 'Global', refersTo: "='数据'!$A$1" }], unknown: { keep: true } }
  return workbook
}

test('worksheet copy keeps sparse axes and complete resources with independent identities and remapped bindings', () => {
  const source = fixture(), before = structuredClone(source)
  const plan = planWorksheetCopy(source, 'source', { userId: 'owner', targetSheetId: 'copy' })
  const next = plan.workbook, sheet = next.sheets[1]
  assert.deepEqual(source, before)
  assert.equal(sheet.id, 'copy'); assert.equal(sheet.name, '数据 (2)'); assert.equal(next.activeSheetId, 'copy')
  assert.deepEqual(sheet.rowHiddenSpans, source.sheets[0].rowHiddenSpans)
  assert.equal(sheet.rowCount, 1_048_576); assert.equal(Object.keys(sheet.cells).length, 3)
  assert.deepEqual(plan.resources, { tables: 1, pivotTables: 1, charts: 1, slicers: 1 })
  const table = next.tables[1], pivot = next.pivotTables[1], chart = next.charts[1], slicer = next.slicers[1]
  assert.notEqual(table.id, 'table'); assert.notEqual(table.columns[0].id, 'column'); assert.equal(table.name, 'Sales_2')
  assert.equal(sheet.cells['1:0'].formula, "=SUM('数据 (2)'!$A$1,Sales_2[金额])")
  assert.notEqual(pivot.id, 'pivot'); assert.equal(pivot.source.subUnitId, 'copy'); assert.equal(pivot.anchor.subUnitId, 'copy')
  assert.equal(pivot.drillDownMap['10:0'].valueFieldId, pivot.fields[0].id)
  assert.equal(sheet.cells['10:0'].metadata.pivotTableId, pivot.id)
  assert.equal(chart.pivotTableId, pivot.id); assert.equal(chart.source.sheetId, 'copy'); assert.deepEqual(chart.position, { x: 100, y: 200 })
  assert.equal(slicer.tableId, table.id); assert.equal(slicer.sourceFieldId, table.columns[0].id)
  assert.equal(slicer.hostSheetId, 'copy'); assert.deepEqual(slicer.selectedItems, ['7'])
  sheet.cells['0:0'].metadata.unknown.keep = false
  assert.equal(source.sheets[0].cells['0:0'].metadata.unknown.keep, true)
  assert.equal(next.metadata.definedNames.find(entry => entry.sheetId === 'copy').refersTo, "='数据 (2)'!$A$1")
  assert.equal(next.metadata.definedNames.find(entry => entry.name === 'Other').localSheetId, 2)
  assert.equal(next.metadata.definedNames.find(entry => entry.name === 'Global').refersTo, "='数据'!$A$1")
})

test('objects hosted on the copy retain cross-sheet and external data sources', () => {
  const workbook = fixture()
  workbook.pivotTables[0].source = { ...workbook.pivotTables[0].source, subUnitId: 'other', sheetName: '其他' }
  workbook.charts[0].pivotTableId = undefined
  workbook.charts[0].source.sheetId = 'other'
  workbook.slicers[0].hostSheetId = 'other'
  const next = planWorksheetCopy(workbook, 'source', { userId: 'owner' }).workbook
  assert.equal(next.pivotTables[1].source.subUnitId, 'other')
  assert.equal(next.charts[1].source.sheetId, 'other')
  assert.equal(next.slicers.length, 1)
})

test('formula copy preserves text, other-sheet, external, 3D and structured column literals', () => {
  const formula = '=SUM(数据!A1,其他!$B$2,\'[book.xlsx]数据\'!A1,数据:其他!A1,Sales[数据!A1],"数据!A1")'
  assert.equal(remapWorksheetCopyFormula(formula, '数据', '数据 (2)', new Map([['sales', 'Sales_2']])),
    '=SUM(\'数据 (2)\'!A1,其他!$B$2,\'[book.xlsx]数据\'!A1,数据:其他!A1,Sales_2[数据!A1],"数据!A1")')
})

test('copy naming respects case-insensitive collisions and UTF-16 limits', () => {
  const source = '12345678901234567890123456😀xx'
  const name = createWorksheetCopyName([{ id: 's', name: source }], source)
  assert.ok(name.length <= 31); assert.equal(/\ud83d(?!\ude00)/u.test(name), false)
  assert.equal(createWorksheetCopyName([{ id: 's', name: 'Data' }, { id: 'x', name: 'data (2)' }], 'Data'), 'Data (3)')
})

test('private protected source ranges block whole-sheet copy before any candidate becomes observable', () => {
  const workbook = fixture()
  workbook.sheets[0].rangeProtectionRules = [{ id: 'rule', permissionId: 'permission', ownerUserId: 'owner', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'noOneElseCanView', ranges: [{ startRow: 900000, startCol: 15000, endRow: 900001, endCol: 15001 }] }]
  const before = structuredClone(workbook)
  assert.throws(() => planWorksheetCopy(workbook, 'source', { userId: 'visitor' }), /权限/)
  assert.deepEqual(workbook, before)
  const next = planWorksheetCopy(workbook, 'source', { userId: 'owner' }).workbook.sheets[1]
  assert.notEqual(next.rangeProtectionRules[0].permissionId, 'permission')
})

test('copy command accepts fixed subUnitId and rejects conflicting or stale parameter shapes', () => {
  assert.equal(copyWorksheetCommandInput({ unitId: 'u', subUnitId: 's' }).subUnitId, 's')
  for (const input of [null, [], { sheetId: 's', subUnitId: 'x' }, { expectedRevision: -1 }, { expectedHistoryGeneration: 0 }, { host: { unitId: 'u' } }]) {
    assert.throws(() => copyWorksheetCommandInput(input))
  }
})
