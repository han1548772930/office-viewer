import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server, api, panel
before(async () => {
  server = await createServer({ configFile: resolve('vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  panel = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-sort-filter/sheet-data-panel-state.ts')
})
after(async () => { await server?.close() })
function setup() {
  const studio = new api.OfficeStudio(), univer = api.FUniver.newAPI(studio), workbook = univer.getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C5').setValues([['地区', '销售', '序号'], ['东', 10, 'a'], ['西', 20, 'b'], ['东', 30, 'c'], ['西', 40, 'd']])
  const range = { startRow: 0, startCol: 0, endRow: 4, endCol: 2 }
  return { studio, univer, workbook, sheet, range, target: { unitId: workbook.getId(), sheetId: sheet.getSheetId() } }
}

test('panel ranges validate explicit worksheet bounds and preserve real header labels', () => {
  const { studio, workbook, sheet, range } = setup()
  try {
    const model = studio.getWorkbookModel(workbook.getId()).sheets[0]
    assert.deepEqual(panel.parseDataPanelRange(model, ' $C$5:$A$1 '), range)
    assert.equal(panel.dataPanelAddress(range), 'A1:C5')
    assert.deepEqual(panel.readDataPanelFields(model, range).map(field => field.label), ['A · 地区', 'B · 销售', 'C · 序号'])
    assert.deepEqual(panel.readDataPanelFields(model, range, false).map(field => field.label), ['列 A', '列 B', '列 C'])
    for (const address of ['A1', 'A1:C1', 'A:C', '0:3', 'A0:C4', 'Sheet!A1:C4', 'A1:XFE9', 'A1:B9007199254740993']) assert.throws(() => panel.parseDataPanelRange(model, address), RangeError)
    assert.equal(sheet.getRange('A1').getValue(), '地区')
  } finally { studio.dispose() }
})

test('value choices retain number/text/boolean distinctions, sparse blanks and formula caches', () => {
  const { studio, workbook, sheet } = setup()
  try {
    sheet.getRange('A1:A8').setValues([['类型'], [1], ['1'], [true], [''], [null], [{ v: 2, f: '=1+1' }], [{ v: '#N/A', t: 'e' }]])
    const model = studio.getWorkbookModel(workbook.getId()).sheets[0]
    model.cells['7:0'].dataType = 'e'
    const values = panel.readDataPanelValues(model, { startRow: 0, endRow: 8, startCol: 0, endCol: 0 }, 0)
    assert.deepEqual(values.map(item => item.value), [1, '1', true, '', null, 2])
    assert.equal(new Set(values.map(item => item.key)).size, values.length)
    assert.throws(() => panel.readDataPanelValues(model, { startRow: 0, endRow: 8, startCol: 0, endCol: 0 }, 1), RangeError)
    assert.equal(panel.parseDataPanelScalar('01', 'text'), '01')
    assert.equal(panel.parseDataPanelScalar(' 1.5 ', 'number'), 1.5)
    assert.equal(panel.parseDataPanelScalar('TRUE', 'boolean'), true)
    assert.equal(panel.parseDataPanelScalar('false', 'boolean'), false)
    for (const value of ['', 'NaN', 'Infinity', '1x']) assert.throws(() => panel.parseDataPanelScalar(value, 'number'))
    assert.throws(() => panel.parseDataPanelScalar('yes', 'boolean'))
  } finally { studio.dispose() }
})

test('configuring multiple filter columns is one undoable command and retains source rows', () => {
  const { studio, workbook, sheet, range, target } = setup()
  try {
    const before = JSON.parse(JSON.stringify(studio.getWorkbookModel(workbook.getId())))
    const revision = studio.getWorkbookRevision(workbook.getId())
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: { 0: { kind: 'values', selectedItems: ['东'] }, 1: { kind: 'condition', operator: 'greaterThan', value: 15 } }, expectedRevision: revision })
    assert.deepEqual(sheet.getFilteredRows(), [1, 2, 4])
    assert.deepEqual(sheet.getRange('A1:C5').getValues(), [['地区', '销售', '序号'], ['东', 10, 'a'], ['西', 20, 'b'], ['东', 30, 'c'], ['西', 40, 'd']])
    assert.equal(studio.getWorkbookRevision(workbook.getId()), revision + 1)
    workbook.undo()
    assert.equal(sheet.getAutoFilter(), null)
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].cells, before.sheets[0].cells)
    workbook.redo()
    assert.deepEqual(sheet.getFilteredRows(), [1, 2, 4])
    const saved = JSON.parse(JSON.stringify(workbook.save()))
    const reopened = new api.OfficeStudio({ workbook: saved })
    assert.deepEqual(api.FUniver.newAPI(reopened).getActiveWorkbook().getActiveSheet().getFilteredRows(), [1, 2, 4])
    reopened.dispose()
  } finally { studio.dispose() }
})

test('filter preflight, stale revisions and cancelled commands preserve model and redo history', () => {
  const { studio, workbook, sheet, range, target } = setup()
  try {
    const before = JSON.parse(JSON.stringify(studio.getWorkbookModel(workbook.getId())))
    const bad = { ...target, range, columns: { 0: { kind: 'values', selectedItems: ['东'] }, 9: { kind: 'values', selectedItems: [] } } }
    assert.throws(() => panel.configureFilterFromPanel(studio.getWorkbookModel(workbook.getId()), bad), /inside/)
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()), before)
    assert.throws(() => studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: {}, expectedRevision: -1 }), /数据已更改/)
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()), before)
    const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === panel.CONFIGURE_SHEET_FILTER_COMMAND) event.cancel = true })
    assert.throws(() => studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: {} }), /canceled/)
    cancel.dispose()
    assert.equal(sheet.getAutoFilter(), null)
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()), before)
  } finally { studio.dispose() }
})

test('empty choices hide all rows while clear and remove keep manual hiding separate', () => {
  const { studio, workbook, sheet, range, target } = setup()
  try {
    studio.getWorkbookModel(workbook.getId()).sheets[0].hiddenRows = [2]
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: { 0: { kind: 'values', selectedItems: [] } } })
    assert.deepEqual(sheet.getFilteredRows(), [1, 2, 3, 4])
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: {} })
    assert.deepEqual(sheet.getFilteredRows(), [])
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range: null, columns: {} })
    assert.equal(sheet.getAutoFilter(), null)
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].hiddenRows, [2])
  } finally { studio.dispose() }
})

test('multi-key sort validates its fields, preserves stable payloads and uses one history transaction', async () => {
  const { studio, workbook, sheet, range } = setup()
  try {
    const criteria = panel.prepareDataPanelSort(range, [{ column: 0, direction: 'ascending' }, { column: 1, direction: 'descending' }])
    assert.throws(() => panel.prepareDataPanelSort(range, []))
    assert.throws(() => panel.prepareDataPanelSort(range, [{ column: 0, direction: 'ascending' }, { column: 0, direction: 'descending' }]))
    assert.throws(() => panel.prepareDataPanelSort(range, [{ column: 3, direction: 'ascending' }]))
    const before = sheet.getRange('A1:C5').getValues()
    await sheet.sortRange('A1:C5', criteria, { hasHeader: true, locale: 'zh-CN' })
    assert.deepEqual(sheet.getRange('A1:C5').getValues(), [['地区', '销售', '序号'], ['东', 30, 'c'], ['东', 10, 'a'], ['西', 40, 'd'], ['西', 20, 'b']])
    workbook.undo()
    assert.deepEqual(sheet.getRange('A1:C5').getValues(), before)
    workbook.redo()
    assert.equal(sheet.getRange('C2').getValue(), 'c')
  } finally { studio.dispose() }
})
