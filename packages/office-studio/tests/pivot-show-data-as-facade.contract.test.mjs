import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// New Facade/command/projection drafts, all deferred to unified validation.
let server, api, fixture, charts
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture, charts] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-chart-engine-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-engine.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const rule = (type, baseFieldId = '', baseItem = '', extra = {}) => ({ type, baseFieldId, baseItem, ...extra })
function setup(ownership = 'owned') {
  const studio = new api.OfficeStudio({ workbook: fixture.pivotChartEngineWorkbook({ ownership: ownership === 'pivotTable' ? 'linked' : ownership }) })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), unitId = book.getId(), chart = book.getPivotChartById('chart-a')
  return { studio, book, unitId, chart, target: ownership === 'pivotTable' ? book.getPivotTables()[0] : chart, ownership,
    model: () => studio.getWorkbookModel(unitId), history: () => studio.editor.history.units[unitId]?.undo.length ?? 0,
    dispose: () => studio.dispose() }
}
const context = f => ({ expectedRevision: f.studio.getWorkbookRevision(f.unitId), expectedUserId: f.studio.currentUserId,
  expectedHistoryGeneration: f.studio.editor.history.captureSavepoint(f.unitId).generation, host: 'workbench' })

test.each(['owned', 'linked', 'pivotTable'])('%s Show Values As API changes real values and reset/undo/save/reopen preserve the report', async ownership => {
  const f = setup(ownership); let reopened
  try {
    const before = json(f.book.save()), baseline = f.history(), desired = rule('percentOfGrandTotal')
    await f.target.setShowDataAs('amount-local', desired, context(f))
    assert.deepEqual(f.target.getShowDataAs('amount-local'), desired)
    assert.deepEqual(f.chart.getData().series[0].values, [.4, .2, .4]); assert.equal(f.history(), baseline + 1)
    if (ownership === 'owned') assert.deepEqual(f.model().sheets.map(sheet => sheet.cells), before.sheets.map(sheet => sheet.cells))
    else {
      assert.deepEqual(f.book.getPivotChartById('chart-b').getData().series[0].values, [.4, .2, .4])
      const report = f.model().pivotTables[0], cell = f.model().sheets.find(sheet => sheet.id === 'reports').cells[`${report.anchor.row + 1}:${report.anchor.col + 1}`]
      assert.equal(cell.style.numberFormat, '0.00%'); assert.match(cell.display, /40/)
    }
    const selected = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
    f.book.redo(); assert.deepEqual(json(f.book.save()), selected)
    reopened = new api.OfficeStudio({ workbook: selected })
    const reopenedChart = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById('chart-a')
    assert.deepEqual(reopenedChart.getShowDataAs('amount-local'), desired)
    const read = f.target.getShowDataAs('amount-local'); read.type = 'index'
    assert.deepEqual(f.target.getShowDataAs('amount-local'), desired)
    await f.target.resetShowDataAs('amount-local'); assert.equal(f.target.getShowDataAs('amount-local'), undefined)
    assert.deepEqual(f.chart.getData().series[0].values, [40, 20, 40])
  } finally { reopened?.dispose(); f.dispose() }
})

test('show rule, aggregation and visual type are configured in a single owned-chart transaction', async () => {
  const f = setup()
  try {
    const before = json(f.book.save()), baseline = f.history(), fields = f.chart.getFieldsConfig()
    fields.fields.find(field => field.id === 'amount-local').aggregation = 'count'
    fields.fields.find(field => field.id === 'amount-local').showDataAs = rule('percentOfGrandTotal')
    await f.chart.configure({ patch: { type: 'line', title: 'Counts as percentages' }, fieldsConfig: fields }, context(f))
    assert.equal(f.history(), baseline + 1); assert.deepEqual(f.chart.getData().series[0].values, [.5, .25, .25])
    assert.equal(f.chart.getFieldSetting('amount-local').aggregation, 'count')
    assert.equal(f.chart.getInfo().title, 'Counts as percentages')
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test.each(['owned', 'pivotTable'])('%s setting combines format/aggregation and Show Values As in the existing field-setting command', async ownership => {
  const f = setup(ownership)
  try {
    const baseline = f.history()
    await f.target.setFieldSetting('amount-local', { subtotalType: 'count', format: '0.0%', showDataAs: rule('percentOfGrandTotal') }, context(f))
    assert.deepEqual(f.chart.getData().series[0].values, [.5, .25, .25]); assert.equal(f.history(), baseline + 1)
    assert.deepEqual(f.target.getShowDataAs('amount-local'), rule('percentOfGrandTotal'))
    assert.equal(f.chart.getFieldSetting('amount-local').format, '0.0%')
  } finally { f.dispose() }
})

test('parameterized difference and cumulative rules use actual base members and report ordering', async () => {
  const f = setup()
  try {
    assert.deepEqual(f.chart.getShowDataAsBaseItems('region-local'), ['North', 'South', 'West'])
    await f.chart.setShowDataAs('amount-local', rule('differenceFrom', 'region-local', 'South'))
    assert.deepEqual(f.chart.getData().series[0].values, [20, null, 20])
    await f.chart.setShowDataAs('amount-local', rule('runningTotal', 'region-local'))
    assert.deepEqual(f.chart.getData().series[0].values, [40, 60, 100])
    await f.chart.setShowDataAs('amount-local', rule('percentRunningTotal', 'region-local'))
    assert.deepEqual(f.chart.getData().series[0].values, [.4, .6, 1])
    await f.chart.setShowDataAs('amount-local', rule('rankDescending', 'region-local'))
    assert.deepEqual(f.chart.getData().series[0].values, [1, 3, 1])
  } finally { f.dispose() }
})

test('a satisfied rule does not add history, report revisions or settled events', async () => {
  const f = setup()
  try {
    await f.chart.setShowDataAs('amount-local', rule('percentOfGrandTotal'))
    const before = json(f.model()), baseline = f.history(), events = []
    f.studio.events.addEvent('PivotChartChanged', event => events.push(event))
    assert.equal(await f.chart.setShowDataAs('amount-local', rule('percentOfGrandTotal')), true)
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline); assert.deepEqual(events, [])
  } finally { f.dispose() }
})

test.each([
  rule('unknown-mode'), rule('percentOf', 'missing-field', 'North'), rule('percentOf', 'region-local', 'Missing member'),
  rule('percentOfGrandTotal', 'region-local'), rule('runningTotal', 'region-local', 'North'),
])('unsupported or incompatible rule %j rejects without saving ignored configuration', async desired => {
  const f = setup()
  try {
    const before = json(f.model()), baseline = f.history()
    await assert.rejects(f.chart.setShowDataAs('amount-local', desired))
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test('removing a required base field rejects the entire fields/visual candidate', async () => {
  const f = setup()
  try {
    await f.chart.setShowDataAs('amount-local', rule('runningTotal', 'region-local'))
    const before = json(f.model()), baseline = f.history(), fields = f.chart.getFieldsConfig()
    fields.fields = fields.fields.filter(field => field.id !== 'region-local')
    await assert.rejects(f.chart.configure({ patch: { title: 'Must not publish' }, fieldsConfig: fields }), /基准|字段|来源/)
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test('direct PivotTable Show Values As emits both linked chart changes with captured revisions', async () => {
  const f = setup('pivotTable')
  try {
    const before = f.chart.getInfo().revision, sibling = f.book.getPivotChartById('chart-b').getInfo().revision, changes = [], reports = []
    f.studio.events.addEvent('PivotChartChanged', event => changes.push(json(event)))
    f.studio.events.addEvent('PivotTableRendered', event => reports.push(json(event)))
    await f.target.setShowDataAs('amount-local', rule('percentOfGrandTotal'))
    assert.equal(f.chart.getInfo().revision, before + 1); assert.equal(f.book.getPivotChartById('chart-b').getInfo().revision, sibling + 1)
    assert.deepEqual(new Set(changes.map(event => event.pivotChartId)), new Set(['chart-a', 'chart-b']))
    assert.equal(reports.length, 1)
    for (const event of changes) assert.equal(event.reportRevision, reports[0].revision)
  } finally { f.dispose() }
})

test.each(['cancel', 'stale', 'permission'])('direct report %s boundary preserves data and history', async boundary => {
  const f = setup('pivotTable')
  try {
    if (boundary === 'permission') f.model().sheets.find(sheet => sheet.id === 'reports').worksheetProtection = {
      id: 'readonly-report', permissionId: 'readonly-report', unitType: 'worksheet', ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
    if (boundary === 'cancel') f.studio.events.addEvent('BeforePivotTableChange', event => { event.cancel = true })
    const before = json(f.model()), baseline = f.history(), request = context(f)
    if (boundary === 'stale') request.expectedRevision += 1
    await assert.rejects(f.target.setShowDataAs('amount-local', rule('percentOfGrandTotal'), request))
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test('error-only Pivot output retains categories, raw errors and value-series bindings in both value directions', async () => {
  const f = setup()
  try {
    f.book.getSheetBySheetId('data').getRange('E4:E7').setValues([[0], [0], [0], [0]])
    await f.chart.setShowDataAs('amount-local', rule('percentOfGrandTotal'))
    const data = f.chart.getData()
    assert.deepEqual(data.categories, ['North', 'South', 'West']); assert.deepEqual(data.series[0].values, [null, null, null])
    assert.deepEqual(data.records.map(row => row[1]), ['#DIV/0!', '#DIV/0!', '#DIV/0!'])
    await f.chart.updateValuePosition('row')
    const rows = f.chart.getData()
    assert.deepEqual(rows.categories, ['North', 'South', 'West']); assert.equal(rows.series.length, 1)
    assert.deepEqual(rows.records.map(row => row[1]), ['#DIV/0!', '#DIV/0!', '#DIV/0!'])
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s retains formatted percentage records for hierarchy and heatmap projection', async ownership => {
  const f = setup(ownership)
  try {
    await f.chart.setShowDataAs('amount-local', rule('percentOfGrandTotal'))
    await f.chart.update({ type: 'heatmap' })
    const records = charts.readChartRecords(f.model(), f.chart.getInfo())
    assert.deepEqual(records.records.map(row => row[1]), [.4, .2, .4])
    assert.deepEqual(records.formattedRecords.map(row => row[1]), ['40.00%', '20.00%', '40.00%'])
    assert.deepEqual(f.chart.getData().series[0].formattedValues, ['40.00%', '20.00%', '40.00%'])
  } finally { f.dispose() }
})

test.each([false, true])('common report lifecycle projects source and increments a linked chart once when pre-incremented=%s', preIncremented => {
  const before = fixture.pivotChartEngineWorkbook({ ownership: 'linked' }), after = json(before), visuals = json(after.charts.map(chart => ({ id: chart.id, title: chart.title, position: chart.position, size: chart.size })))
  after.pivotTables[0].revision += 1
  after.pivotTables[0].outputRange.endRow += 1
  if (preIncremented) after.charts[0].revision += 1
  api.synchronizeLinkedPivotChartReports(before, after)
  for (const chart of after.charts) {
    assert.deepEqual(chart.source.range, after.pivotTables[0].outputRange)
    assert.equal(chart.revision, before.charts.find(previous => previous.id === chart.id).revision + 1)
  }
  assert.deepEqual(after.charts.map(chart => ({ id: chart.id, title: chart.title, position: chart.position, size: chart.size })), visuals)
})

test('common lifecycle leaves unchanged reports, owned reports and missing imported references intact', () => {
  const before = fixture.pivotChartEngineWorkbook({ ownership: 'linked' }), after = json(before)
  after.charts.push({ ...json(after.charts[0]), id: 'missing-import', pivotTableId: 'missing-report' })
  const owned = fixture.pivotChartEngineWorkbook({ ownership: 'owned' }).charts[0]
  after.charts.push({ ...json(owned), id: 'owned-independent' })
  const snapshot = json(after)
  api.synchronizeLinkedPivotChartReports(before, after)
  assert.deepEqual(after, snapshot)
})
