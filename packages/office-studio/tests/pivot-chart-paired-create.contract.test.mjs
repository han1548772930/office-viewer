import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Real paired report/chart transaction drafts. Execution is deferred to unified validation.
let server, api, fixture
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-chart-engine-fixture.mjs'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function setup() {
  const studio = new api.OfficeStudio({ workbook: fixture.pivotChartEngineWorkbook({ ownership: 'none' }) })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), unitId = book.getId()
  return { studio, book, unitId, host: book.getSheetBySheetId('reports'), data: book.getSheetBySheetId('data'),
    model: () => studio.getWorkbookModel(unitId), history: () => studio.editor.history.units[unitId]?.undo.length ?? 0,
    dispose: () => studio.dispose() }
}
const info = (f, configured = true) => {
  const builder = f.host.newPivotChart('column').setSource(json(fixture.pivotChartSourceInput))
    .setAbsolutePosition(300, 40).setSize(240, 160)
  if (configured) builder.setInitialFieldsConfig({ fields: fixture.pivotChartFields(), options: json(fixture.pivotChartReportOptions) })
  return builder.build()
}
const options = (f, sheetId = 'reports') => ({ pivotTable: { pivotTableId: 'paired-report',
  anchor: { unitId: f.unitId, subUnitId: sheetId, row: 8, col: 1 } } })
const context = f => ({ expectedRevision: f.studio.getWorkbookRevision(f.unitId), expectedUserId: f.studio.currentUserId,
  expectedHistoryGeneration: f.studio.editor.history.captureSavepoint(f.unitId).generation, host: 'workbench' })

test.each([false, true])('paired creation with configured=%s publishes one report/chart transaction and reopens it', async configured => {
  const f = setup(); let reopened
  try {
    const before = json(f.book.save()), baseline = f.history(), sourceCells = json(f.model().sheets.find(sheet => sheet.id === 'data').cells)
    const chart = await f.host.insertPivotChart(info(f, configured), context(f), options(f))
    assert.equal(chart.getSourceKind(), 'pivotTable'); assert.equal(chart.isLinkedToPivotTable(), true)
    assert.equal(chart.getPivotTable().getPivotTableId(), 'paired-report')
    assert.equal(f.model().pivotTables.length, 1); assert.equal(f.model().charts.length, 1); assert.equal(f.history(), baseline + 1)
    assert.equal(f.model().charts[0].ownedPivot, undefined)
    assert.deepEqual(f.model().sheets.find(sheet => sheet.id === 'data').cells, sourceCells)
    assert.deepEqual(chart.getData().series.map(series => series.values), configured ? [[40, 20, 40]] : [])
    const saved = json(f.book.save())
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
    f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
    reopened = new api.OfficeStudio({ workbook: saved })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById(chart.getId())
    assert.equal(restored.getPivotTable().getPivotTableId(), 'paired-report')
    assert.deepEqual(restored.getData().series.map(series => series.values), configured ? [[40, 20, 40]] : [])
  } finally { reopened?.dispose(); f.dispose() }
})

test('paired destination can differ from chart host without changing source cells', async () => {
  const f = setup()
  try {
    const chart = await f.host.insertPivotChart(info(f), context(f), options(f, 'dashboard'))
    assert.equal(chart.getSubUnitId(), 'reports')
    assert.deepEqual(chart.getPivotTableReference(), { unitId: f.unitId, subUnitId: 'dashboard', pivotTableId: 'paired-report' })
    assert.equal(f.model().charts[0].source.sheetId, 'dashboard')
    assert.deepEqual(f.model().charts[0].source.range, f.model().pivotTables[0].outputRange)
    assert.deepEqual(chart.getData().series[0].values, [40, 20, 40])
  } finally { f.dispose() }
})

test.each(['chart', 'fields', 'anchor'])('invalid %s leaves neither paired resource nor output cells behind', async boundary => {
  const f = setup()
  try {
    const request = info(f), paired = options(f)
    if (boundary === 'chart') request.size.width = 0
    if (boundary === 'fields') request.initialFieldsConfig.fields[1].sourceColumn = 9000
    if (boundary === 'anchor') f.host.getRange('B9').setValue('Occupied')
    const before = json(f.model()), baseline = f.history()
    await assert.rejects(f.host.insertPivotChart(request, context(f), paired))
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test.each(['BeforePivotChartChange', 'BeforePivotTableChange', 'reentry'])('%s preserves paired atomicity', async boundary => {
  const f = setup()
  try {
    const baseline = f.history(), settled = []
    f.studio.events.addEvent('PivotChartAdded', event => settled.push(event))
    if (boundary === 'reentry') f.studio.events.addEvent('BeforePivotChartChange', () => f.host.getRange('J30').setValue('Observer edit'))
    else f.studio.events.addEvent(boundary, event => { event.cancel = true })
    await assert.rejects(f.host.insertPivotChart(info(f), context(f), options(f)))
    assert.equal(f.model().pivotTables.length, 0); assert.equal(f.model().charts.length, 0); assert.deepEqual(settled, [])
    assert.equal(f.history(), baseline + (boundary === 'reentry' ? 1 : 0))
    if (boundary === 'reentry') assert.equal(f.host.getRange('J30').getValue(), 'Observer edit')
    assert.equal(f.model().sheets.find(sheet => sheet.id === 'reports').cells['8:1'], undefined)
  } finally { f.dispose() }
})

test.each(['source', 'output', 'expanded-output', 'chart'])('paired %s protection rejects every staged resource', async boundary => {
  const f = setup()
  try {
    const sheet = f.model().sheets.find(sheet => sheet.id === (boundary === 'source' ? 'data' : 'reports'))
    if (boundary === 'source') sheet.worksheetProtection = { id: 'locked-source', permissionId: 'locked-source', unitType: 'worksheet',
      ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: 'noOneElseCanView' }
    else {
      const [row, col] = boundary === 'output' ? [8, 1] : boundary === 'expanded-output' ? [10, 2] : [2, 4]
      sheet.rangeProtectionRules = [{ id: 'paired-lock', permissionId: 'paired-lock', ownerUserId: 'peer', allowedUserIds: [],
        editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [{ startRow: row, endRow: row, startCol: col, endCol: col }] }]
    }
    const before = json(f.model()), baseline = f.history()
    await assert.rejects(f.host.insertPivotChart(info(f), context(f), options(f)), /权限|保护|查看|编辑/)
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test.each(['expectedRevision', 'expectedHistoryGeneration', 'expectedUserId', 'host'])('paired stale/invalid %s is rejected before registration', async key => {
  const f = setup()
  try {
    const request = context(f), before = json(f.model()), baseline = f.history()
    request[key] = key === 'host' ? 'invalid-host' : key === 'expectedUserId' ? 'someone-else' : request[key] + 1
    await assert.rejects(f.host.insertPivotChart(info(f), request, options(f)))
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test('caller edits in Before listeners cannot alter the captured paired destination or initial fields', async () => {
  const f = setup()
  try {
    const request = info(f), paired = options(f), targets = []
    f.studio.events.addEvent('BeforePivotChartChange', () => { paired.pivotTable.anchor.row = 20; request.initialFieldsConfig.fields[1].aggregation = 'count' })
    f.studio.events.addEvent('BeforePivotTableChange', event => targets.push(json(event)))
    const chart = await f.host.insertPivotChart(request, context(f), paired)
    assert.equal(f.model().pivotTables[0].anchor.row, 8); assert.deepEqual(chart.getData().series[0].values, [40, 20, 40])
    assert.deepEqual(targets.map(event => [event.subUnitId, event.pivotTableId]), [['reports', 'paired-report']])
  } finally { f.dispose() }
})

test('an existing linked source rejects extra paired-table options rather than silently ignoring them', async () => {
  const f = setup()
  try {
    const chart = await f.host.insertPivotChart(info(f), {}, options(f)), before = json(f.model()), baseline = f.history()
    const linked = f.host.newPivotChart('line').setPivotTable(chart.getPivotTable()).build()
    await assert.rejects(f.host.insertPivotChart(linked, {}, options(f)), /区域来源/)
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test('paired reports drive linked charts and survive removal of the drawing', async () => {
  const f = setup()
  try {
    const chart = await f.host.insertPivotChart(info(f), {}, options(f)), report = chart.getPivotTable()
    await report.setShowDataAs('amount-local', { type: 'percentOfGrandTotal', baseFieldId: '', baseItem: '' })
    assert.deepEqual(chart.getData().series[0].values, [.4, .2, .4])
    const output = json(f.model().sheets.find(sheet => sheet.id === 'reports').cells), baseline = f.history()
    await chart.remove(); assert.equal(f.model().charts.length, 0); assert.equal(f.model().pivotTables.length, 1)
    assert.deepEqual(f.model().sheets.find(sheet => sheet.id === 'reports').cells, output); assert.equal(f.history(), baseline + 1)
    f.book.undo(); assert.deepEqual(f.book.getPivotChartById(chart.getId()).getData().series[0].values, [.4, .2, .4])
  } finally { f.dispose() }
})
