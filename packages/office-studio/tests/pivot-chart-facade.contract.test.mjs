import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Real registry/Facade/history integration drafts. Deferred unified validation;
// this file has not been executed in the implementation phase.
let server, api, fixture, connections
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture, connections] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-chart-engine-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/slicer-connections-fixture.mjs'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const model = f => f.studio.getWorkbookModel(f.unitId)
const history = f => f.studio.editor.history.units[f.unitId]?.undo.length ?? 0
const cells = f => json(model(f).sheets.map(sheet => ({ id: sheet.id, cells: sheet.cells })))
const config = () => ({ fields: fixture.pivotChartFields(), options: json(fixture.pivotChartReportOptions) })
function setup(options = { ownership: 'none' }) {
  const studio = new api.OfficeStudio({ workbook: fixture.pivotChartEngineWorkbook(options) })
  const office = api.FOffice.newAPI(studio), book = office.getActiveWorkbook(), unitId = book.getId()
  return { studio, office, book, unitId, data: book.getSheetBySheetId('data'), host: book.getSheetBySheetId('reports'),
    dashboard: book.getSheetBySheetId('dashboard'), chart: book.getPivotChartById('chart-a'), dispose: () => studio.dispose() }
}
const rangeInfo = f => f.host.newPivotChart('column').setSource(f.data.getRange('D3:H7'))
  .setAbsolutePosition(300, 40).setSize(240, 160)
const context = f => ({ expectedRevision: f.studio.getWorkbookRevision(f.unitId),
  expectedHistoryGeneration: f.studio.editor.history.captureSavepoint(f.unitId).generation,
  expectedUserId: f.studio.currentUserId, host: 'workbench' })

test.each([false, true])('range creation with configured=%s owns a report without reserving worksheet cells', async configured => {
  const f = setup(); let reopened
  try {
    const before = cells(f), baseline = history(f), builder = rangeInfo(f)
    if (configured) builder.setInitialFieldsConfig(config())
    const chart = await f.host.insertPivotChart(builder.build(), context(f))
    assert.ok(chart instanceof api.FSheetPivotChart); assert.equal(chart.getSourceKind(), 'range')
    assert.equal(chart.isLinkedToPivotTable(), false); assert.equal(chart.getPivotTable(), null)
    assert.deepEqual(chart.getSourceRangeInfo(), fixture.pivotChartSourceInput)
    assert.equal(chart.getUnitId(), f.unitId); assert.equal(chart.getSubUnitId(), 'reports')
    assert.equal(model(f).pivotTables.length, 0); assert.equal(model(f).sheets.length, 3)
    assert.deepEqual(cells(f), before); assert.equal(history(f), baseline + 1)
    const report = model(f).charts.find(item => item.id === chart.getId()).ownedPivot
    assert.equal(report.outputRange, undefined); assert.equal(report.generatedCellKeys, undefined); assert.equal(report.drillDownMap, undefined)
    assert.deepEqual(chart.getData().series.map(series => series.values), configured ? [[40, 20, 40]] : [])
    const saved = json(f.book.save())
    f.book.undo(); assert.equal(f.host.getPivotChart(chart.getId()), null); assert.deepEqual(cells(f), before)
    f.book.redo(); assert.equal(f.host.getPivotCharts().length, 1)
    reopened = new api.OfficeStudio({ workbook: saved })
    const reopenedChart = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById(chart.getId())
    assert.deepEqual(reopenedChart.getFieldsConfig(), chart.getFieldsConfig())
    assert.deepEqual(reopenedChart.getSourceRangeInfo(), chart.getSourceRangeInfo())
    assert.deepEqual(reopenedChart.getData().series.map(series => series.values), configured ? [[40, 20, 40]] : [])
  } finally { reopened?.dispose(); f.dispose() }
})

test('an empty owned report exposes actual unconfigured source members and can be configured later', async () => {
  const f = setup()
  try {
    const chart = await f.host.insertPivotChart(rangeInfo(f).build()), before = cells(f)
    assert.deepEqual(chart.getFieldItems('region-source'), ['North', 'South', 'West'])
    assert.deepEqual(chart.getFieldItems('segment-source'), ['A', 'B'])
    const metadata = chart.getSourceFieldsInfo()
    assert.deepEqual(metadata.map(field => field.id), ['region-source', 'amount-source', 'date-source', 'segment-source', 'score-source'])
    metadata[0].name = 'Caller mutation'; assert.equal(chart.getSourceFieldsInfo()[0].name, 'Region')
    const baseline = history(f)
    assert.equal(await chart.setFieldsConfig(config()), true)
    assert.equal(history(f), baseline + 1); assert.deepEqual(cells(f), before)
    assert.deepEqual(chart.getData().categories, ['North', 'South', 'West'])
    assert.deepEqual(chart.getData().series[0].values, [40, 20, 40])
    const fields = chart.getFieldsConfig(); fields.fields[0].displayName = 'Mutated draft'
    assert.equal(chart.getFieldSetting('region-local').displayName, undefined)
  } finally { f.dispose() }
})

test('string, FRange and fully qualified source builders resolve the same cross-sheet source', async () => {
  const f = setup()
  try {
    const text = f.host.newPivotChart('line').setSource("'Data'!D3:H7").build()
    const range = f.host.newPivotChart('line').setSource(f.data.getRange('D3:H7')).build()
    const qualified = f.host.newPivotChart('line').setSource(fixture.pivotChartSourceInput).build()
    assert.deepEqual(text.source, range.source)
    assert.deepEqual(qualified.source.sourceRangeInfo, fixture.pivotChartSourceInput)
    const other = f.studio.createWorkbook({ name: 'Other workbook' }, false)
    const otherRange = new api.FWorksheet(f.studio, other.id, other.activeSheetId).getRange('A1:B2')
    assert.throws(() => rangeInfo(f).setSource(otherRange), /同一工作簿/)
    assert.throws(() => f.host.newPivotChart('line').build(), /来源/)
  } finally { f.dispose() }
})

test('linked creation reuses its real report and derives drawing source from the report output', async () => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    const before = cells(f), count = model(f).pivotTables.length, baseline = history(f)
    const pivot = f.book.getPivotTables().find(item => item.getPivotTableId() === 'linked-report')
    const chart = await f.dashboard.insertPivotChart(f.dashboard.newPivotChart('line').setPivotTable(pivot).setAbsolutePosition(40, 20).build())
    assert.equal(chart.getSourceKind(), 'pivotTable'); assert.equal(chart.isLinkedToPivotTable(), true)
    assert.deepEqual(chart.getPivotTableReference(), { unitId: f.unitId, subUnitId: 'reports', pivotTableId: 'linked-report' })
    assert.equal(chart.getSubUnitId(), 'dashboard'); assert.equal(chart.getPivotTable().getPivotTableId(), 'linked-report')
    assert.equal(model(f).pivotTables.length, count); assert.deepEqual(cells(f), before)
    assert.equal(history(f), baseline + 1); assert.deepEqual(chart.getData().series[0].values, [40, 20, 40])
    assert.deepEqual(chart.getInfo().source.range, model(f).pivotTables[0].outputRange)
    assert.deepEqual(f.dashboard.getPivotCharts().map(item => item.getId()), ['chart-b', chart.getId()])
    assert.equal(f.host.getPivotChart(chart.getId()), null)
  } finally { f.dispose() }
})

test('legacy linked ChartConfig creation retains ownership and normalizes its drawing source', async () => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    const chart = await f.host.insertPivotChart(f.host.newChart('line').setPivotTableId('linked-report').setSource('Z1:AA2').build())
    assert.deepEqual(chart.getInfo().source.range, model(f).pivotTables[0].outputRange)
    assert.equal(chart.getInfo().ownedPivot, undefined); assert.equal(chart.getSourceKind(), 'pivotTable')
    assert.deepEqual(chart.getData().series[0].values, [40, 20, 40])
  } finally { f.dispose() }
})

test.each(['BeforePivotChartChange', 'BeforePivotTableChange'])('linked initial fields honor %s cancellation before changing an existing report', async eventName => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    const before = json(model(f)), baseline = history(f), events = []
    f.studio.events.addEvent('BeforePivotChartChange', event => { events.push(['chart', event]); if (eventName === 'BeforePivotChartChange') event.cancel = true })
    f.studio.events.addEvent('BeforePivotTableChange', event => { events.push(['pivot', event]); if (eventName === 'BeforePivotTableChange') event.cancel = true })
    const pivot = f.book.getPivotTables()[0], fields = config(); fields.fields[1].aggregation = 'count'
    await assert.rejects(f.host.insertPivotChart(f.host.newPivotChart('line').setPivotTable(pivot).setInitialFieldsConfig(fields).build()), /取消/)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
    assert.deepEqual(events.map(([kind]) => kind), eventName === 'BeforePivotChartChange' ? ['chart'] : ['chart', 'pivot'])
    if (eventName === 'BeforePivotTableChange') { assert.equal(events[1][1].pivotTableId, 'linked-report'); assert.equal(events[1][1].subUnitId, 'reports') }
  } finally { f.dispose() }
})

test('linked initial fields and new drawing are published together in one history entry', async () => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    const before = json(model(f)), baseline = history(f), fields = config(); fields.fields[1].aggregation = 'count'
    const chart = await f.host.insertPivotChart(f.host.newPivotChart('line').setPivotTable(f.book.getPivotTables()[0]).setInitialFieldsConfig(fields).build())
    assert.equal(history(f), baseline + 1)
    assert.deepEqual(chart.getData().series[0].values, [2, 1, 1])
    assert.equal(f.chart.getFieldSetting('amount-local').aggregation, 'count')
    assert.equal(f.book.getPivotChartById('chart-b').getFieldSetting('amount-local').aggregation, 'count')
    f.book.undo(); assert.deepEqual(json(model(f)), before)
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('configure combines report and visual edits for %s in one history/revision', async ownership => {
  const f = setup({ ownership, configured: true })
  try {
    const before = json(model(f)), baseline = history(f), chartRevision = f.chart.getInfo().revision
    const reportRevision = ownership === 'owned' ? f.chart.getInfo().ownedPivot.revision : model(f).pivotTables[0].revision
    const fields = config(); fields.fields[1].aggregation = 'count'; fields.fields[0].displayName = 'Territory'
    assert.equal(await f.chart.configure({ patch: { type: 'line', title: 'Configured together' }, fieldsConfig: fields }, context(f)), true)
    assert.equal(history(f), baseline + 1); assert.equal(f.chart.getInfo().revision, chartRevision + 1)
    assert.equal(f.chart.getFieldSetting('amount-local').aggregation, 'count')
    assert.equal(f.chart.getFieldSetting('region-local').displayName, 'Territory')
    assert.equal(f.chart.getInfo().title, 'Configured together'); assert.equal(f.chart.getInfo().type, 'line')
    assert.deepEqual(f.chart.getData().series[0].values, [2, 1, 1])
    const report = ownership === 'owned' ? f.chart.getInfo().ownedPivot : model(f).pivotTables[0]
    assert.equal(report.revision, reportRevision + 1)
    if (ownership === 'owned') assert.deepEqual(cells(f), before.sheets.map(sheet => ({ id: sheet.id, cells: sheet.cells })))
    else assert.equal(f.book.getPivotChartById('chart-b').getFieldSetting('amount-local').aggregation, 'count')
    const after = json(model(f)); f.book.undo(); assert.deepEqual(json(model(f)), before)
    f.book.redo(); assert.deepEqual(json(model(f)), after)
  } finally { f.dispose() }
})

test('report-only editing survives the source-aware chart authoring pass', async () => {
  const f = setup({ ownership: 'owned', configured: true })
  try {
    const baseline = history(f), before = cells(f)
    await f.chart.setSubtotalType('amount-local', 'count')
    assert.equal(f.chart.getFieldSetting('amount-local').aggregation, 'count')
    assert.deepEqual(f.chart.getData().series[0].values, [2, 1, 1]); assert.deepEqual(cells(f), before)
    assert.equal(history(f), baseline + 1)
    await f.chart.setLabelManualFilter('region-local', ['West'])
    assert.deepEqual(f.chart.getFieldSetting('region-local').selectedItems, ['West'])
    assert.deepEqual(f.chart.getData().categories, ['West']); assert.deepEqual(f.chart.getData().series[0].values, [1])
  } finally { f.dispose() }
})

test('equivalent configure and satisfied report options return true without history or settled notifications', async () => {
  const f = setup({ ownership: 'owned', configured: true })
  try {
    const before = json(model(f)), baseline = history(f), revision = f.studio.getWorkbookRevision(f.unitId), settled = []
    f.studio.events.addEvent('PivotChartChanged', event => settled.push(event))
    assert.equal(await f.chart.configure({ patch: {}, fieldsConfig: f.chart.getFieldsConfig() }), true)
    assert.equal(await f.chart.setReportOptions(f.chart.getReportOptions()), true)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
    assert.equal(f.studio.getWorkbookRevision(f.unitId), revision); assert.deepEqual(settled, [])
  } finally { f.dispose() }
})

test.each(['source', 'pivotTableId', 'ownedPivot'])('configure rejects %s replacement atomically', async key => {
  const f = setup({ ownership: 'owned', configured: true })
  try {
    const before = json(model(f)), baseline = history(f)
    const values = { source: { sheetId: 'data', range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 }, orientation: 'columns' },
      pivotTableId: 'linked-report', ownedPivot: undefined }
    await assert.rejects(f.chart.configure({ patch: { [key]: values[key] } }), /替换/)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('invalid visual settings roll back the staged fields and worksheet output', async () => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    const before = json(model(f)), baseline = history(f), fields = config(); fields.fields[1].aggregation = 'count'
    await assert.rejects(f.chart.configure({ patch: { size: { width: 0, height: 160 } }, fieldsConfig: fields }), /size|尺寸/i)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test.each(['host', 'source', 'output'])('the %s permission boundary rejects report edits before cancellable events', async boundary => {
  const value = fixture.pivotChartEngineWorkbook({ ownership: 'linked', configured: true })
  const sheet = value.sheets.find(item => item.id === (boundary === 'source' ? 'data' : 'reports'))
  if (boundary === 'output') sheet.rangeProtectionRules = [{ id: 'output-lock', permissionId: 'output-lock', ownerUserId: 'peer', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [json(value.pivotTables[0].outputRange)] }]
  else sheet.worksheetProtection = { id: 'sheet-lock', permissionId: 'sheet-lock', unitType: 'worksheet', ownerUserId: 'peer', allowedUserIds: [],
    editState: 'onlyMe', viewState: boundary === 'source' ? 'noOneElseCanView' : 'othersCanView' }
  const studio = new api.OfficeStudio({ workbook: value }), book = api.FOffice.newAPI(studio).getActiveWorkbook()
  try {
    const before = json(studio.getWorkbookModel(book.getId())), events = []
    studio.events.addEvent('BeforePivotChartChange', event => events.push(event))
    await assert.rejects(book.getPivotChartById('chart-a').setSubtotalType('amount-local', 'count'), /权限|保护|查看|编辑/)
    assert.deepEqual(events, []); assert.deepEqual(json(studio.getWorkbookModel(book.getId())), before)
    assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
  } finally { studio.dispose() }
})

test('new linked output colliding with existing cells rejects every staged report and chart change', async () => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    f.host.getRange('B14').setValue('Occupied below the output')
    const before = json(model(f)), baseline = history(f), fields = config()
    fields.fields.push({ id: 'segment-local', sourceFieldId: 'segment-source', sourceColumn: 6, name: 'Segment', area: 'row', position: 1 })
    fields.options.showRowSubtotals = true
    await assert.rejects(f.chart.configure({ patch: { title: 'Must not publish' }, fieldsConfig: fields }), /conflict|覆盖|冲突|数据/i)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test.each(['BeforeCommandExecute', 'BeforePivotChartChange', 'BeforePivotTableChange'])('throwing %s listeners retain the original error and publish no candidate', async eventName => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    const before = json(model(f)), baseline = history(f), failure = new Error('before-listener-failed')
    f.studio.events.addEvent(eventName, () => { throw failure })
    await assert.rejects(f.chart.setSubtotalType('amount-local', 'count'), error => error === failure)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('Before mutation of the caller request cannot change the snapshotted field edit', async () => {
  const f = setup({ ownership: 'owned', configured: true })
  try {
    const request = { unitId: f.unitId, sheetId: 'reports', pivotChartId: 'chart-a', edit: { type: 'setFieldSetting', fieldId: 'amount-local', setting: { subtotalType: 'count' } } }
    f.studio.events.addEvent('BeforePivotChartChange', () => { request.edit.setting.subtotalType = 'sum'; request.sheetId = 'dashboard' })
    assert.equal(await f.studio.executeCommand(api.PIVOT_CHART_COMMANDS.edit, request), true)
    assert.equal(f.chart.getFieldSetting('amount-local').aggregation, 'count')
    assert.deepEqual(f.chart.getData().series[0].values, [2, 1, 1])
  } finally { f.dispose() }
})

test('reentrant content edits from Before prevent the stale candidate and retain the intervening content', async () => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    const baseline = history(f), pivots = []
    f.studio.events.addEvent('BeforePivotChartChange', () => { f.book.setName('Intervening edit') })
    f.studio.events.addEvent('BeforePivotTableChange', event => pivots.push(event))
    await assert.rejects(f.chart.setSubtotalType('amount-local', 'count'), /更改/)
    assert.equal(model(f).name, 'Intervening edit'); assert.equal(f.chart.getFieldSetting('amount-local').aggregation, 'sum')
    assert.equal(history(f), baseline + 1); assert.deepEqual(pivots, [])
  } finally { f.dispose() }
})

test('settled chart and report payloads are captured and isolated from observer mutation and failure', async () => {
  const f = setup({ ownership: 'linked', configured: true })
  try {
    const changes = [], reports = [], errors = [], baseline = history(f)
    f.studio.events.addEvent('PivotChartChanged', event => { event.pivotChartId = 'Corrupted'; throw new Error('observer') })
    f.studio.events.addEvent('PivotChartChanged', event => changes.push(json(event)))
    f.studio.events.addEvent('PivotTableRendered', event => { if (event.outputRange) event.outputRange.startRow = 999 })
    f.studio.events.addEvent('PivotTableRendered', event => reports.push(json(event)))
    f.studio.events.addEvent('EventListenerError', event => errors.push(event))
    await f.chart.setSubtotalType('amount-local', 'count')
    assert.deepEqual(new Set(changes.map(event => event.pivotChartId)), new Set(['chart-a', 'chart-b']))
    assert.equal(reports.length, 1); assert.deepEqual(reports[0].outputRange, model(f).pivotTables[0].outputRange)
    assert.ok(errors.length >= 2); assert.equal(history(f), baseline + 1)
    assert.equal(changes.find(event => event.pivotChartId === 'chart-a').reportRevision, model(f).pivotTables[0].revision)
  } finally { f.dispose() }
})

test.each(['expectedRevision', 'expectedHistoryGeneration', 'expectedUserId', 'host'])('stale or invalid %s context rejects without changing content', async key => {
  const f = setup({ ownership: 'owned', configured: true })
  try {
    const before = json(model(f)), baseline = history(f), values = { expectedRevision: 999, expectedHistoryGeneration: 999, expectedUserId: 'another-user', host: 'another-host' }
    await assert.rejects(f.chart.refresh({ ...context(f), [key]: values[key] }), /更改|无效/)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('a bound Facade edits its inactive unit while workbench context refuses a changed host', async () => {
  const f = setup({ ownership: 'owned', configured: true })
  try {
    const other = f.studio.createWorkbook({ name: 'Other workbook' }, true)
    assert.equal(await f.chart.setSubtotalType('amount-local', 'count'), true)
    const before = json(model(f)), baseline = history(f)
    await assert.rejects(f.chart.refresh(context(f)), /宿主/)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
    assert.equal(f.studio.getWorkbookModel(other.id).charts?.length ?? 0, 0)
    assert.equal(f.studio.editor.history.units[other.id]?.undo.length ?? 0, 0)
  } finally { f.dispose() }
})

test('refresh reads changed source values once and keeps the owned report free of worksheet output', async () => {
  const f = setup({ ownership: 'owned', configured: true })
  try {
    f.data.getRange('E4').setValue(50)
    const before = cells(f), baseline = history(f), revision = f.chart.getInfo().revision
    assert.equal(await f.chart.refresh(), true)
    assert.deepEqual(f.chart.getData().series[0].values, [80, 20, 40]); assert.deepEqual(cells(f), before)
    assert.equal(history(f), baseline + 1); assert.equal(f.chart.getInfo().revision, revision + 1)
    assert.equal(f.chart.getInfo().ownedPivot.outputRange, undefined)
  } finally { f.dispose() }
})

test('reset produces an editable empty chart and delete/undo restore its independently owned report', async () => {
  const f = setup({ ownership: 'owned', configured: true })
  try {
    const before = cells(f)
    assert.equal(await f.chart.resetFields(), true); assert.deepEqual(f.chart.getFieldsConfig().fields, [])
    assert.deepEqual(f.chart.getData().series, []); assert.deepEqual(f.chart.getFieldItems('region-source'), ['North', 'South', 'West'])
    const reset = json(model(f)), baseline = history(f), events = []
    f.studio.events.addEvent('PivotChartRemoved', event => events.push(json(event)))
    assert.equal(await f.chart.remove(context(f)), true)
    assert.equal(f.book.getPivotChartById('chart-a'), null); assert.equal(history(f), baseline + 1)
    assert.deepEqual(cells(f), before); assert.deepEqual(events, [{ unitId: f.unitId, subUnitId: 'reports', pivotChartId: 'chart-a' }])
    f.book.undo(); assert.deepEqual(json(model(f)), reset)
  } finally { f.dispose() }
})

test('owned chart filters synchronize connected reports and views through the common candidate transaction', async () => {
  const studio = new api.OfficeStudio({ workbook: connections.slicerConnectionsWorkbook({ affine: false }) })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), ids = connections.connectionFixtureIds
  try {
    const control = book.getSlicer(ids.slicer)
    await control.setConnections({ pivotTables: [{ subUnitId: ids.host, pivotTableId: ids.primary }, { subUnitId: ids.other, pivotTableId: ids.secondary }],
      ownedPivotCharts: [{ subUnitId: ids.host, pivotChartId: ids.owned }] })
    const before = json(book.save()), baseline = studio.editor.history.units[book.getId()].undo.length
    await book.getPivotChartById(ids.owned).setLabelManualFilter('0', ['North'])
    assert.equal(studio.editor.history.units[book.getId()].undo.length, baseline + 1)
    for (const id of [ids.primary, ids.secondary, ids.owned]) assert.deepEqual(connections.connectionReportSelection(studio.getWorkbookModel(book.getId()), id), ['North'])
    assert.deepEqual(control.getInfo().selectedItems, ['North']); assert.deepEqual(book.getSlicer(ids.shared).getInfo().selectedItems, ['North'])
    assert.deepEqual(book.getPivotChartById(ids.owned).getData().series[0].values, [40])
    book.undo(); assert.deepEqual(json(book.save()), before)
  } finally { studio.dispose() }
})

test('removing a connected owned chart retains surviving report connections and undo restores its report', async () => {
  const studio = new api.OfficeStudio({ workbook: connections.slicerConnectionsWorkbook({ affine: false }) })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), ids = connections.connectionFixtureIds
  try {
    const control = book.getSlicer(ids.slicer), linked = { subUnitId: ids.host, pivotTableId: ids.primary }
    await control.setConnections({ pivotTables: [linked], ownedPivotCharts: [{ subUnitId: ids.host, pivotChartId: ids.owned }] })
    const before = json(book.save()), baseline = studio.editor.history.units[book.getId()].undo.length
    await book.getPivotChartById(ids.owned).remove()
    assert.equal(book.getPivotChartById(ids.owned), null); assert.deepEqual(control.getConnections(), { pivotTables: [linked] })
    assert.equal(studio.editor.history.units[book.getId()].undo.length, baseline + 1)
    assert.equal(studio.getWorkbookModel(book.getId()).pivotTables.length, before.pivotTables.length)
    book.undo(); assert.deepEqual(json(book.save()), before)
  } finally { studio.dispose() }
})
