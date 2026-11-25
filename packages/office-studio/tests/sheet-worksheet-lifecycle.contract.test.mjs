import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createBlankSheet, createCell } from '../src/domain/model.ts'
import { WORKSHEET_LIFECYCLE_COMMANDS, planWorksheetLifecycle, planWorksheetInsert, planWorksheetRemove,
  planWorksheetRename, planWorksheetOrder, planWorksheetHiddenState, planWorksheetTabColor,
  worksheetHiddenState } from '../src/features/sheet-workbench/sheet-worksheet-lifecycle.ts'
import { invalidateRemovedTableReferences, rewriteWorksheetFormulaReferences } from '../src/features/sheet-workbench/sheet-worksheet-references.ts'
import { sheetTabDropOrder } from '../src/features/sheet-workbench/sheet-tab-drag-sort.ts'
import { worksheetLifecycleCommandInput } from '../src/api/sheet-lifecycle-commands.ts'
import { getPivotSourceFields } from '../src/features/pivot/services/pivot-engine.ts'

const range = { startRow: 0, startCol: 0, endRow: 1, endCol: 0 }
function fixture() {
  const workbook = createBlankWorkbook(); workbook.id = 'lifecycle-unit'
  workbook.sheets = ['Data', 'Middle', 'End', 'Host'].map(name => ({ ...createBlankSheet(name), id: name.toLowerCase(), rowCount: 20, columnCount: 10 }))
  workbook.activeSheetId = 'data'
  workbook.metadata = { definedNames: [{ name: 'Global', refersTo: '=Data!$A$1' },
    { name: 'Local', localSheetId: 0, refersTo: '=Data!$A$1' },
    { name: 'MiddleLocal', localSheetId: 1, refersTo: '=Middle!$B$2' }],
    localDefinedNames: [{ name: 'LocalId', sheetId: 'data', sheetName: 'Data', refersTo: '=Data!$A$1' }],
    unknown: { preserved: true } }
  const host = workbook.sheets[3]
  host.cells['0:0'] = { ...createCell(0, 0), value: 3, display: '3',
    formula: '=SUM(Data!A1,Data!Local,Sales[Amount],"Data!A1")',
    sparkline: { type: 'line', sourceRange: '=Data!A1:A2' }, hyperlink: '#Data!A1',
    richText: [{ text: 'source', hyperlink: '#gid=data&range=A1', style: { bold: true } }, { text: ' web', hyperlink: 'https://example.com/' }] }
  host.dataValidations = [{ id: 'formula-rule', type: 'custom', ranges: [range], formula1: '=Data!A1>0' },
    { id: 'list-rule', type: 'list', ranges: [range], formula1: 'Data!A1,Keep' }]
  host.conditionalFormats = [{ id: 'conditional', type: 'expression', ranges: [range], formula1: '=Data!A1>0',
    formula2: '', operator: 'equal', value: '=Data!A1>0', config: { min: { type: 'formula', value: '=Data!A1' } } }]
  workbook.tables = [{ id: 'table', name: 'Sales', sheetId: 'data', range, columns: [{ id: 'amount', name: 'Amount' }], showHeaderRow: true }]
  const source = { unitId: workbook.id, subUnitId: 'data', sheetName: 'Data', range }
  workbook.pivotTables = ['data', 'host'].map((id, index) => ({ id: 'pivot-' + index, source: { ...source },
    anchor: { unitId: workbook.id, subUnitId: id, row: 10, col: 0 }, fields: [],
    options: { repeatRowLabels: false, showRowGrandTotal: true, showColGrandTotal: true, showRowSubtotals: true, showColSubtotals: true, valuePosition: 'column' }, revision: 1 }))
  workbook.charts = [{ id: 'hosted-chart', sheetId: 'data', type: 'bar', source: { sheetId: 'data', range } },
    { id: 'source-chart', sheetId: 'host', type: 'bar', source: { sheetId: 'data', range } },
    { id: 'pivot-chart', sheetId: 'host', type: 'bar', pivotTableId: 'pivot-0', source: { sheetId: 'data', range } }]
  workbook.slicers = [{ id: 'table-slicer', hostSheetId: 'host', tableId: 'table' },
    { id: 'hosted-slicer', hostSheetId: 'data', pivotTableId: 'pivot-1' },
    { id: 'pivot-slicer', hostSheetId: 'host', pivotTableId: 'pivot-0' },
    { id: 'retained-slicer', hostSheetId: 'host', pivotTableId: 'pivot-1' }]
  return workbook
}

test('rename updates every modeled reference owner while preserving source identity and literal text', () => {
  const source = fixture(), before = structuredClone(source)
  const next = planWorksheetRename(source, 'data', "Budget's data", 'owner').workbook
  assert.deepEqual(source, before)
  const cell = next.sheets[3].cells['0:0']
  assert.equal(cell.formula, '=SUM(\'Budget\'\'s data\'!A1,\'Budget\'\'s data\'!Local,Sales[Amount],"Data!A1")')
  assert.equal(cell.sparkline.sourceRange, "='Budget''s data'!A1:A2")
  assert.equal(cell.hyperlink, '#gid=data&range=A1')
  assert.equal(cell.richText[0].hyperlink, '#gid=data&range=A1')
  assert.equal(next.sheets[3].dataValidations[0].formula1, "='Budget''s data'!A1>0")
  assert.equal(next.sheets[3].dataValidations[1].formula1, 'Data!A1,Keep')
  assert.equal(next.sheets[3].conditionalFormats[0].config.min.value, "='Budget''s data'!A1")
  assert.equal(next.metadata.definedNames[0].refersTo, "='Budget''s data'!$A$1")
  assert.equal(next.metadata.localDefinedNames[0].sheetName, "Budget's data")
  assert.equal(next.pivotTables[1].source.sheetName, "Budget's data")
  assert.equal(next.charts[1].source.sheetId, 'data')
  assert.deepEqual(next.metadata.unknown, { preserved: true })
  assert.equal(planWorksheetRename(next, 'data', "Budget's data", 'owner').changed, false)
})

test('qualified names, complete axes and mixed external range endpoints follow lexical sheet identity', () => {
  const book = fixture(), change = { sheetId: 'data', oldName: 'Data', newName: 'Report' }
  const formula = '=SUM(Data!$A$1,Data!$A:$C,Data!1:3,Data!Local,Data:End!A1,Data!A1:\'[book.xlsx]End\'!B2,\'[book.xlsx]Data\'!A1,Sales[Data!A1],"Data!A1")'
  assert.equal(rewriteWorksheetFormulaReferences(formula, book, change),
    '=SUM(\'Report\'!$A$1,\'Report\'!$A:$C,\'Report\'!1:3,\'Report\'!Local,\'Report:End\'!A1,\'Report\'!A1:\'[book.xlsx]End\'!B2,\'[book.xlsx]Data\'!A1,Sales[Data!A1],"Data!A1")')
  assert.equal(rewriteWorksheetFormulaReferences('=SUM(Data!A1,Data!Local,End!B2:Data!C3)', book, { sheetId: 'data', oldName: 'Data' }), '=SUM(#REF!,#REF!,#REF!)')
  assert.equal(rewriteWorksheetFormulaReferences('=SUM(Data:End!A1,\'End:Data\'!A1,Data:Middle!A1)', book, { sheetId: 'data', oldName: 'Data' }),
    "=SUM('Middle:End'!A1,'End:Middle'!A1,'Middle'!A1)")
  assert.equal(rewriteWorksheetFormulaReferences('=Data:End!A1', book, { sheetId: 'middle', oldName: 'Middle' }), '=Data:End!A1')
})

test('removed structured tables invalidate whole references without editing column literals or external workbooks', () => {
  const formula = '=SUM(Sales[[#Data],[Amount]],Sales[Data!A1],Other[Sales],\'[book.xlsx]Data\'!Sales[Amount],"Sales[Amount]")'
  assert.equal(invalidateRemovedTableReferences(formula, new Set(['sales'])),
    '=SUM(#REF!,#REF!,Other[Sales],\'[book.xlsx]Data\'!Sales[Amount],"Sales[Amount]")')
})

test('remove cleans hosted resources and scoped names while retaining explicit broken sources on surviving objects', () => {
  const source = fixture(), before = structuredClone(source), next = planWorksheetRemove(source, 'data', 'owner').workbook
  assert.deepEqual(source, before)
  assert.equal(next.activeSheetId, 'middle')
  assert.equal(next.sheets[2].cells['0:0'].formula, '=SUM(#REF!,#REF!,#REF!,"Data!A1")')
  assert.equal(next.sheets[2].cells['0:0'].hyperlink, null)
  assert.deepEqual(next.sheets[2].cells['0:0'].richText[0], { text: 'source', style: { bold: true } })
  assert.equal(next.sheets[2].cells['0:0'].richText[1].hyperlink, 'https://example.com/')
  assert.deepEqual(next.tables, [])
  assert.deepEqual(next.pivotTables.map(pivot => pivot.id), ['pivot-1'])
  assert.equal(next.pivotTables[0].source.subUnitId, 'data')
  assert.throws(() => getPivotSourceFields(next, next.pivotTables[0]), /source worksheet not found/)
  assert.deepEqual(next.charts.map(chart => chart.id), ['source-chart'])
  assert.equal(next.charts[0].source.sheetId, 'data')
  assert.deepEqual(next.slicers.map(slicer => slicer.id), ['retained-slicer'])
  assert.deepEqual(next.metadata.definedNames.map(entry => entry.name), ['Global', 'MiddleLocal'])
  assert.equal(next.metadata.definedNames[0].refersTo, '=#REF!')
  assert.equal(next.metadata.definedNames[1].localSheetId, 0)
  assert.deepEqual(next.metadata.localDefinedNames, [])
})

test('insert and persisted reordering rebind positional local names by stable sheet ID', () => {
  const source = fixture(), inserted = planWorksheetInsert(source, { userId: 'owner', index: 1, name: 'DATA', sheet: { id: 'new-sheet' } }).workbook
  assert.equal(inserted.sheets[1].name, 'DATA1')
  assert.equal(inserted.activeSheetId, 'new-sheet')
  assert.equal(inserted.metadata.definedNames[2].localSheetId, 2)
  const moved = planWorksheetOrder(inserted, 'host', 0, 'owner').workbook
  assert.deepEqual(moved.sheets.map(sheet => sheet.id), ['host', 'data', 'new-sheet', 'middle', 'end'])
  assert.equal(moved.metadata.definedNames[1].localSheetId, 1)
  assert.equal(moved.metadata.definedNames[2].localSheetId, 3)
  assert.equal(planWorksheetOrder(moved, 'host', 0, 'owner').changed, false)
  assert.throws(() => planWorksheetInsert(source, { userId: 'owner', sheet: { id: 'data' } }), /ID.*唯一/)
  assert.throws(() => planWorksheetInsert(source, { userId: 'owner', name: 'Bad:Name' }), /名称/)
  assert.throws(() => planWorksheetOrder(source, 'data', 4, 'owner'), /顺序/)
})

test('hidden and veryHidden states retain one visible sheet and restore an accessible active neighbor', () => {
  let source = fixture()
  source = planWorksheetHiddenState(source, 'data', 2, 'owner').workbook
  assert.equal(worksheetHiddenState(source.sheets[0]), 2); assert.equal(source.sheets[0].metadata.hiddenState, 2)
  assert.equal(source.activeSheetId, 'middle')
  source = planWorksheetHiddenState(source, 'middle', 1, 'owner').workbook
  source = planWorksheetHiddenState(source, 'end', 1, 'owner').workbook
  const before = structuredClone(source)
  assert.throws(() => planWorksheetHiddenState(source, 'host', 1, 'owner'), /唯一可见/)
  assert.throws(() => planWorksheetRemove(source, 'host', 'owner'), /可见工作表/)
  assert.deepEqual(source, before)
  const shown = planWorksheetHiddenState(source, 'data', 0, 'owner').workbook
  assert.equal(shown.activeSheetId, 'data'); assert.equal(shown.sheets[0].visible, true)
  assert.equal(worksheetHiddenState(shown.sheets[0]), 0)
  assert.equal(planWorksheetLifecycle(source, WORKSHEET_LIFECYCLE_COMMANDS.hidden, { subUnitId: 'data' }, 'owner').changed, false)
})

test('worksheet permissions and remote protected ranges reject lifecycle edits before publishing a candidate', () => {
  const source = fixture()
  source.sheets[0].rowCount = 1_048_576; source.sheets[0].columnCount = 16_384
  source.sheets[0].rangeProtectionRules = [{ id: 'private', permissionId: 'private-id', ownerUserId: 'owner', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'noOneElseCanView', ranges: [{ startRow: 900000, endRow: 900001, startCol: 15000, endCol: 15001 }] }]
  const before = structuredClone(source)
  assert.throws(() => planWorksheetRemove(source, 'data', 'visitor'), /保护区域/)
  assert.deepEqual(source, before)
  source.sheets[0].worksheetProtection = { id: 'worksheet-rule', permissionId: 'worksheet-permission', unitType: 'worksheet', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
  for (const action of [() => planWorksheetRename(source, 'data', 'Other', 'visitor'),
    () => planWorksheetOrder(source, 'data', 1, 'visitor'), () => planWorksheetTabColor(source, 'data', '#123456', 'visitor')]) assert.throws(action, /权限/)
})

test('visible tab drop computes the full persisted target while retaining hidden and inaccessible entries', () => {
  const order = ['a', 'hidden', 'b', 'private', 'c'], visible = ['a', 'b', 'c']
  assert.equal(sheetTabDropOrder(order, visible, 'a', 'c', true), 4)
  assert.equal(sheetTabDropOrder(order, visible, 'a', 'c', false), 2)
  assert.equal(sheetTabDropOrder(order, visible, 'c', 'a', false), 0)
  assert.equal(sheetTabDropOrder(order, visible, 'c', 'a', true), 2)
  assert.equal(sheetTabDropOrder(order, visible, 'b', 'b', false), 2)
  assert.equal(sheetTabDropOrder(order, visible, 'hidden', 'a', true), undefined)
  assert.equal(sheetTabDropOrder(order, visible, 'a', 'missing', true), undefined)
})

test('command DTO rejects conflicting identifiers and invalid lifecycle payloads', () => {
  assert.equal(worksheetLifecycleCommandInput({ subUnitId: 'data', hidden: 2 }).hidden, 2)
  for (const input of [null, [], { sheetId: 'data', subUnitId: 'host' }, { name: 1 }, { value: false }, { index: 1.5 },
    { order: -1 }, { hidden: 3 }, { sheet: [] }, { expectedRevision: -1 }, { expectedHistoryGeneration: 0 }]) assert.throws(() => worksheetLifecycleCommandInput(input))
})
