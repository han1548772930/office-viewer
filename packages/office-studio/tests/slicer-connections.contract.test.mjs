import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Deferred contract drafts. No test/build/browser/benchmark has been executed for this batch.
let server, api, engine, connections, fixture, charts
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, engine, connections, fixture, charts] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-engine.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-connections.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/slicer-connections-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-engine.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const clone = value => structuredClone(value)
const view = (workbook, key = 'slicer') => workbook.slicers.find(control => control.id === fixture.connectionFixtureIds[key])
const refs = (...keys) => ({ pivotTables: keys.filter(key => key !== 'owned' && key !== 'linked').map(key => ({
  subUnitId: fixture.connectionFixtureIds.host, pivotTableId: fixture.connectionFixtureIds[key] })),
  ownedPivotCharts: keys.filter(key => key === 'owned' || key === 'linked').map(key => ({
    subUnitId: fixture.connectionFixtureIds.host, pivotChartId: fixture.connectionFixtureIds[key] })) })
const selection = (workbook, key, fieldId = '0') => fixture.connectionReportSelection(workbook, fixture.connectionFixtureIds[key], fieldId)

test('connection candidates include compatible tables and owned reports once, with linked owner normalization', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
  const candidates = connections.getSlicerConnectionCandidates(workbook, control)
  assert.deepEqual(candidates.map(candidate => [candidate.kind, candidate.id]), [
    ['pivotTable', fixture.connectionFixtureIds.primary], ['pivotTable', fixture.connectionFixtureIds.secondary], ['ownedPivotChart', fixture.connectionFixtureIds.owned],
  ])
  assert.deepEqual(connections.getSlicerConnectionCandidates(workbook, view(workbook, 'table')), [])
  assert.equal(connections.getSlicerConnections(workbook, view(workbook, 'table')), undefined)
})

test('connection snapshots are defensive and the source field stays locked', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook), before = clone(workbook)
  const snapshot = connections.getSlicerConnections(workbook, control)
  snapshot.pivotTables[0].pivotTableId = 'changed'; snapshot.pivotTables.push({ subUnitId: 'x', pivotTableId: 'y' })
  assert.deepEqual(workbook, before)
  assert.throws(() => connections.planSlicerConnections(workbook, control, { ...refs('primary'), sourceFieldId: '2' }),
    error => error instanceof connections.SlicerConnectionError && error.code === 'invalid-target')
})

test('replacement deduplicates linked owners while preserving an independently owned field instance', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook), plan = connections.planSlicerConnections(workbook, control,
    { pivotTables: [...refs('primary').pivotTables, ...refs('primary').pivotTables], ownedPivotCharts: refs('linked', 'owned', 'linked').ownedPivotCharts })
  assert.deepEqual(plan.connections, refs('primary', 'owned'))
  assert.equal(plan.reports.find(report => report.kind === 'ownedPivotChart').field.sourceFieldId, '0')
  assert.notEqual(plan.reports.find(report => report.kind === 'ownedPivotChart').field.id, '0')
})

for (const invalid of ['other-cache', 'wrong-sheet', 'ordinary-chart', 'unknown-report', 'foreign-unit', 'invalid-list']) {
  test(`invalid ${invalid} target refuses the entire connection collection without dropping valid targets`, () => {
    const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
    const input = refs('primary')
    if (invalid === 'other-cache') input.pivotTables.push(...refs('incompatible').pivotTables)
    if (invalid === 'wrong-sheet') input.pivotTables.push({ subUnitId: fixture.connectionFixtureIds.source, pivotTableId: fixture.connectionFixtureIds.secondary })
    if (invalid === 'ordinary-chart') { const chart = clone(workbook.charts[0]); chart.id = 'ordinary'; delete chart.pivotTableId; workbook.charts.push(chart); input.ownedPivotCharts.push({ subUnitId: chart.sheetId, pivotChartId: chart.id }) }
    if (invalid === 'unknown-report') input.pivotTables.push({ subUnitId: fixture.connectionFixtureIds.host, pivotTableId: 'missing' })
    if (invalid === 'foreign-unit') input.pivotTables[0].unitId = 'foreign-workbook'
    if (invalid === 'invalid-list') input.pivotTables = 'primary'
    const before = clone(workbook)
    assert.throws(() => engine.setSlicerConnections(workbook, control, input), error => error instanceof connections.SlicerConnectionError)
    assert.deepEqual(workbook, before)
  })
}

test('connection identity compares quoted tuple keys, including IDs containing colons', () => {
  assert.equal(connections.sameConnections({ pivotTables: [{ subUnitId: 'a:b', pivotTableId: 'c' }] },
    { pivotTables: [{ subUnitId: 'a', pivotTableId: 'b:c' }] }), false)
})

test('one source selection fans out to two tables, an owned report and every shared view', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
  engine.setSlicerConnections(workbook, control, refs('primary', 'secondary', 'owned'))
  const presentation = workbook.slicers.slice(0, 2).map(({ selectedItems, timelineRange, revision, ...author }) => clone(author))
  engine.setSlicerItems(workbook, control, ['North'])
  for (const key of ['primary', 'secondary', 'owned']) assert.deepEqual(selection(workbook, key), ['North'])
  for (const key of ['slicer', 'shared']) assert.deepEqual(engine.getSlicer(workbook, view(workbook, key).id).selectedItems, ['North'])
  assert.deepEqual(workbook.slicers.slice(0, 2).map(({ selectedItems, timelineRange, revision, ...author }) => author), presentation)
  engine.clearSlicerFilter(workbook, control)
  for (const key of ['primary', 'secondary', 'owned']) assert.equal(selection(workbook, key), undefined)
})

test('timeline ranges fan out independently of a different source field and clear together', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook, 'timeline')
  engine.setSlicerConnections(workbook, control, refs('primary', 'secondary', 'owned'))
  engine.setSlicerTimelineRange(workbook, control, { start: '2026-02-01', end: '2026-02-28', level: 'month' })
  for (const key of ['primary', 'secondary', 'owned']) assert.deepEqual(selection(workbook, key, '1'), ['2026-02-02', '2026-02-20'])
  assert.equal(selection(workbook, 'secondary', '0'), undefined)
  engine.clearSlicerFilter(workbook, control)
  for (const key of ['primary', 'secondary', 'owned']) assert.equal(selection(workbook, key, '1'), undefined)
})

test('empty replacement keeps source members and selected values for a later reconnection', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
  engine.setSlicerItems(workbook, control, ['North'])
  engine.setSlicerConnections(workbook, control, { pivotTables: [] })
  assert.deepEqual(engine.getSlicerItems(workbook, control), ['North', 'South'])
  assert.deepEqual(engine.readSlicerSelection(workbook, control), ['North'])
  engine.setSlicerItems(workbook, control, ['South'])
  assert.deepEqual(selection(workbook, 'primary'), ['North'])
  engine.setSlicerConnections(workbook, control, refs('secondary', 'owned'))
  assert.deepEqual(selection(workbook, 'secondary'), ['South']); assert.deepEqual(selection(workbook, 'owned'), ['South'])
})

test('replacement removes old influence and retains the detached report last filter', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
  engine.setSlicerConnections(workbook, control, refs('primary', 'secondary'))
  engine.setSlicerItems(workbook, control, ['North'])
  engine.setSlicerConnections(workbook, control, refs('owned'))
  engine.setSlicerItems(workbook, control, ['South'])
  assert.deepEqual(selection(workbook, 'primary'), ['North']); assert.deepEqual(selection(workbook, 'secondary'), ['North'])
  assert.deepEqual(connections.getSlicerConnections(workbook, control), { pivotTables: [], ownedPivotCharts: refs('owned').ownedPivotCharts })
})

test('an output conflict on the second report publishes no partial source selection or cell output', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
  engine.setSlicerConnections(workbook, control, refs('primary', 'secondary')); engine.setSlicerItems(workbook, control, ['North'])
  workbook.sheets.find(sheet => sheet.id === fixture.connectionFixtureIds.host).cells['19:8'] = api.createCell(19, 8, 'occupied')
  const before = clone(workbook)
  assert.throws(() => engine.clearSlicerFilter(workbook, control), /overlap|overwrite|conflict|occupied/i)
  assert.deepEqual(workbook, before)
})

test('owned filtering uses the Pivot calculation without creating worksheet report cells', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
  engine.setSlicerConnections(workbook, control, refs('owned'))
  const beforeCells = clone(workbook.sheets.map(sheet => sheet.cells))
  engine.setSlicerItems(workbook, control, ['North'])
  assert.deepEqual(workbook.sheets.map(sheet => sheet.cells), beforeCells)
  const chart = workbook.charts.find(chart => chart.id === fixture.connectionFixtureIds.owned), records = charts.readChartRecords(workbook, chart)
  assert.ok(records.records.some(record => record[0] === 'North' && record[1] === 40))
  assert.deepEqual(chart.ownedPivot.generatedCellKeys, [])
})

test('filtering every owned report member out produces an empty chart source rather than stale categories', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
  engine.setSlicerConnections(workbook, control, refs('owned')); engine.setSlicerItems(workbook, control, [])
  const chart = workbook.charts.find(chart => chart.id === fixture.connectionFixtureIds.owned)
  assert.deepEqual(charts.readChartRecords(workbook, chart).records, [])
  assert.deepEqual(charts.readChartData(workbook, chart).categories, [])
})

test('a retained nonnumeric source field remains readable after all reports are disconnected', () => {
  const workbook = fixture.slicerConnectionsWorkbook(), control = view(workbook)
  const cache = workbook.slicerCaches.find(cache => cache.id === control.cacheId)
  const pivotCache = workbook.pivotCaches.find(cache => cache.id === cache.source.pivotCacheId)
  cache.source.sourceFieldId = 'stable-region'; control.sourceFieldId = 'stable-region'; view(workbook, 'shared').sourceFieldId = 'stable-region'
  pivotCache.fields.push({ id: 'stable-region', sourceColumn: 0, name: 'Region' })
  for (const pivot of workbook.pivotTables) if (pivot.pivotCacheId === pivotCache.id) pivot.fields.find(field => field.sourceColumn === 0).sourceFieldId = 'stable-region'
  engine.setSlicerConnections(workbook, control, { pivotTables: [] })
  assert.deepEqual(engine.getSlicerItems(workbook, control), ['North', 'South'])
  pivotCache.fields = pivotCache.fields.filter(field => field.id !== 'stable-region')
  assert.throws(() => engine.getSlicerItems(workbook, control), error => error.code === 'field-missing')
})
