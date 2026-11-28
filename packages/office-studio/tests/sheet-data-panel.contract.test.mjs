import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server, api, panel
let dataRange
beforeAll(async () => {
  server = await createServer({ configFile: resolve('vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  panel = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-sort-filter/sheet-data-panel-state.ts')
  dataRange = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-sort-filter/sheet-data-range.ts')
})
afterAll(async () => { await server?.close() })
function setup() {
  const studio = new api.OfficeStudio(), officeApi = api.FOffice.newAPI(studio), workbook = officeApi.getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C5').setValues([['地区', '销售', '序号'], ['东', 10, 'a'], ['西', 20, 'b'], ['东', 30, 'c'], ['西', 40, 'd']])
  const range = { startRow: 0, startCol: 0, endRow: 4, endCol: 2 }
  return { studio, officeApi, workbook, sheet, range, target: { unitId: workbook.getId(), sheetId: sheet.getSheetId() } }
}

test('public empty cell bodies retain data-region presence across snapshot import, clear, API writes and reopen', () => {
  const studio = new api.OfficeStudio(), officeApi = api.FOffice.newAPI(studio)
  try {
    const book = officeApi.createWorkbook({ id: 'public-body-region', sheetOrder: ['data'], sheets: { data: {
      id: 'data', name: 'Data', rowCount: 30, columnCount: 10,
      cellData: { 0: { 0: { v: 'head' } }, 1: { 0: { p: {} } }, 2: { 0: { p: { body: { dataStream: '\r\n' } } } },
        3: { 0: { v: null, s: { bg: { rgb: '#ffeeaa' } } } } },
    } } })
    const source = { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }, sheet = book.getActiveSheet()
    const model = () => studio.getWorkbookModel(book.getId()).sheets[0]
    assert.equal(dataRange.expandSheetDataRange(model(), source).endRow, 2)
    assert.equal(dataRange.sheetCellHasData(model().cells['3:0']), false)
    const publicSnapshot = book.getSnapshot(), reopened = officeApi.createWorkbook({ ...publicSnapshot, id: 'public-body-region-reopened' })
    assert.equal(dataRange.expandSheetDataRange(studio.getWorkbookModel(reopened.getId()).sheets[0], source).endRow, 2)
    sheet.activate(); sheet.getRange('A2').setValue({ p: null, v: null })
    assert.equal(dataRange.expandSheetDataRange(model(), source).endRow, 0)
    sheet.getRange('A2').setValue({ p: {} })
    assert.equal(dataRange.expandSheetDataRange(model(), source).endRow, 2)
    sheet.getRange('A2').clearContent()
    assert.equal(dataRange.sheetCellHasData(model().cells['1:0']), false)
    const nativeReopened = officeApi.createWorkbook({ ...book.save(), id: 'public-body-region-native-reopened' })
    assert.equal(dataRange.expandSheetDataRange(studio.getWorkbookModel(nativeReopened.getId()).sheets[0], source).endRow, 0)
  } finally { studio.dispose() }
})

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
    assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getFilteredRows(), [1, 2, 4])
    reopened.dispose()
  } finally { studio.dispose() }
})

test('color filtering retains source styles, content history, redo and native snapshot reopening', () => {
  const { studio, workbook, sheet, range, target } = setup()
  try {
    sheet.getRange('B2').setBackground('#ff0000')
    sheet.getRange('B4').setBackground('#ff0000')
    const source = structuredClone(studio.getWorkbookModel(workbook.getId()).sheets[0].cells)
    const depth = studio.editor.history.units[workbook.getId()].undo.length
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range,
      columns: { 1: { kind: 'color', cellFillColors: ['rgb(255,0,0)'] } }, expectedRevision: studio.getWorkbookRevision(workbook.getId()) })
    assert.deepEqual(sheet.getFilteredRows(), [2, 4])
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].cells, source)
    assert.equal(studio.editor.history.units[workbook.getId()].undo.length, depth + 1)
    workbook.undo(); assert.equal(sheet.getAutoFilter(), null)
    workbook.redo(); assert.deepEqual(sheet.getFilteredRows(), [2, 4])
    const reopened = new api.OfficeStudio({ workbook: JSON.parse(JSON.stringify(workbook.save())) })
    try { assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getFilteredRows(), [2, 4]) }
    finally { reopened.dispose() }
  } finally { studio.dispose() }
})

test('a rejected merged group crossing the filter end survives content history and native snapshot reopening', () => {
  const { studio, workbook, sheet, target } = setup()
  try {
    sheet.getRange('A2:A5').merge({ isForceMerge: true })
    const cells = structuredClone(studio.getWorkbookModel(workbook.getId()).sheets[0].cells)
    const depth = studio.editor.history.units[workbook.getId()].undo.length
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target,
      range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 }, columns: { 0: { kind: 'values', selectedItems: [] } },
      expectedRevision: studio.getWorkbookRevision(workbook.getId()) })
    assert.deepEqual(sheet.getFilteredRows(), [1, 2, 3, 4])
    assert.equal(sheet.getAutoFilter().range.endRow, 2)
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].cells, cells)
    assert.equal(studio.editor.history.units[workbook.getId()].undo.length, depth + 1)
    workbook.undo(); assert.equal(sheet.getAutoFilter(), null)
    workbook.redo(); assert.deepEqual(sheet.getFilteredRows(), [1, 2, 3, 4])
    const reopened = new api.OfficeStudio({ workbook: JSON.parse(JSON.stringify(workbook.save())) })
    try { assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getFilteredRows(), [1, 2, 3, 4]) }
    finally { reopened.dispose() }
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

test('detached filter clearing resets derived rows without relying on a workbook commit', () => {
  const { studio, workbook, range, target } = setup()
  try {
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: { 0: { kind: 'values', selectedItems: ['东'] } } })
    const detached = JSON.parse(JSON.stringify(workbook.save()))
    detached.sheets[0].hiddenRows = [2]
    const cells = JSON.parse(JSON.stringify(detached.sheets[0].cells))
    const before = JSON.stringify(studio.getWorkbookModel(workbook.getId()))
    assert.deepEqual(detached.sheets[0].autoFilter.filteredRows, [2, 4])
    panel.configureFilterFromPanel(detached, { ...target, range, columns: {} })
    assert.deepEqual(detached.sheets[0].autoFilter.range, range)
    assert.deepEqual(detached.sheets[0].autoFilter.columns, {})
    assert.deepEqual(detached.sheets[0].autoFilter.filteredRows, [])
    assert.deepEqual(detached.sheets[0].hiddenRows, [2])
    assert.deepEqual(detached.sheets[0].cells, cells)
    assert.equal(JSON.stringify(studio.getWorkbookModel(workbook.getId())), before)
  } finally { studio.dispose() }
})

test('empty choices hide all rows while clear and remove keep manual hiding separate', () => {
  const { studio, workbook, sheet, range, target } = setup()
  try {
    studio.getWorkbookModel(workbook.getId()).sheets[0].hiddenRows = [2]
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: { 0: { kind: 'values', selectedItems: [] } } })
    assert.deepEqual(sheet.getFilteredRows(), [1, 2, 3, 4])
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: {} })
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].autoFilter.filteredRows, [])
    assert.deepEqual(sheet.getFilteredRows(), [])
    studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range: null, columns: {} })
    assert.equal(sheet.getAutoFilter(), null)
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].hiddenRows, [2])
  } finally { studio.dispose() }
})

test('filter toggle and clearing are undoable while reapply preserves content history', () => {
  const { studio, workbook, sheet, range, target } = setup()
  try {
    studio.executeCommandSync('sheet.command.toggle-auto-filter', { ...target, range })
    assert.equal(sheet.getAutoFilter().range.endRow, 4)
    studio.executeCommandSync('sheet.command.set-filter-criteria', { ...target, column: 0, criteria: { kind: 'values', selectedItems: ['东'] } })
    assert.deepEqual(sheet.getFilteredRows(), [2, 4])

    studio.executeCommandSync('sheet.command.clear-filter', target)
    assert.deepEqual(sheet.getFilteredRows(), [])
    assert.deepEqual(sheet.getAutoFilter().columns, {})
    workbook.undo()
    assert.deepEqual(sheet.getFilteredRows(), [2, 4])

    const modelSheet = studio.getWorkbookModel(workbook.getId()).sheets[0]
    modelSheet.autoFilter.filteredRows = []
    const beforeReapply = JSON.parse(JSON.stringify(studio.editor.history.units[workbook.getId()]))
    const modified = workbook.save().modified
    studio.executeCommandSync('sheet.command.reapply-filter', target)
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].autoFilter.filteredRows, [2, 4])
    assert.deepEqual(studio.editor.history.units[workbook.getId()].undo, beforeReapply.undo)
    assert.deepEqual(studio.editor.history.units[workbook.getId()].redo, beforeReapply.redo)
    assert.equal(workbook.save().modified, modified)
    workbook.redo()
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].autoFilter.columns, {})
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].autoFilter.filteredRows, [])
    workbook.undo()
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].autoFilter.filteredRows, [2, 4])

    studio.executeCommandSync('sheet.command.toggle-auto-filter', { ...target, range })
    assert.equal(sheet.getAutoFilter(), null)
    workbook.undo()
    assert.deepEqual(sheet.getFilteredRows(), [2, 4])
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

test('a numeric computation that resolves after cancellation cannot fall back to synchronous sorting or create history', async () => {
  const { studio, workbook, sheet } = setup()
  try {
    const before = JSON.stringify(studio.getWorkbookModel(workbook.getId())), depth = studio.editor.history.units[workbook.getId()].undo.length
    const controller = new AbortController()
    let release
    studio.numericIndexService.sortRows = () => new Promise(resolve => { release = resolve })
    const pending = sheet.sortRange('B2:B5', [{ column: 1, direction: 'descending' }], { signal: controller.signal })
    assert.equal(typeof release, 'function'); controller.abort(); release(undefined)
    await assert.rejects(pending, error => error?.name === 'AbortError')
    assert.equal(JSON.stringify(studio.getWorkbookModel(workbook.getId())), before)
    assert.equal(studio.editor.history.units[workbook.getId()].undo.length, depth)
  } finally { studio.dispose() }
})

test('sort host guards are checked again after numeric computation and inside before-command dispatch', async () => {
  for (const stage of ['computed', 'before-command']) {
    const { studio, workbook, sheet } = setup()
    try {
      const before = JSON.stringify(studio.getWorkbookModel(workbook.getId())), depth = studio.editor.history.units[workbook.getId()].undo.length
      let current = true, release
      studio.numericIndexService.sortRows = () => new Promise(resolve => { release = resolve })
      const guard = () => { if (!current) throw new Error('host retired') }
      const subscription = studio.events.addEvent('BeforeCommandExecute', event => { if (stage === 'before-command' && event.id === 'sheet.command.sort-range') current = false })
      const pending = sheet.sortRange('B2:B5', [{ column: 1, direction: 'descending' }], { assertCurrent: guard })
      if (stage === 'computed') current = false
      release(undefined); await assert.rejects(pending, /host retired/); subscription.dispose()
      assert.equal(JSON.stringify(studio.getWorkbookModel(workbook.getId())), before)
      assert.equal(studio.editor.history.units[workbook.getId()].undo.length, depth)
    } finally { studio.dispose() }
  }
})

test('filter configuration rejects a before-command switch to a different workbook without publishing a stale mutation', () => {
  const { studio, officeApi, workbook, range, target } = setup()
  try {
    const other = officeApi.createWorkbook({ id: 'filter-neighbor' }, { makeCurrent: false })
    const before = JSON.stringify(studio.getWorkbookModel(workbook.getId())), depth = studio.editor.history.units[workbook.getId()].undo.length
    const subscription = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === panel.CONFIGURE_SHEET_FILTER_COMMAND) studio.activateWorkbook(other.getId()) })
    assert.throws(() => studio.executeCommandSync(panel.CONFIGURE_SHEET_FILTER_COMMAND, { ...target, range, columns: {} }))
    subscription.dispose()
    assert.equal(JSON.stringify(studio.getWorkbookModel(workbook.getId())), before)
    assert.equal(studio.editor.history.units[workbook.getId()].undo.length, depth)
  } finally { studio.dispose() }
})
