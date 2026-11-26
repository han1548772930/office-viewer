import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Deferred model/projection/layout contracts. No test or build was run while writing this batch.
let server, engine, pivot, fixture, axis, metadata, graphics, frame
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[engine, pivot, fixture, axis, metadata, graphics, frame] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-engine.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-engine.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-axis-state-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-pivot-axis-layout.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-pivot-axis.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/echarts-option.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-layout.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function setup(options = {}) {
  const workbook = fixture.pivotAxisWorkbook(options), chart = workbook.charts[0]
  chart.size = { width: 900, height: 500 }
  chart.valueFields = []
  return { workbook, chart, data: () => engine.readChartData(workbook, chart), report: () => pivot.getPivotChartReport(workbook, chart) }
}
function labels(f) {
  const data = f.data(), layout = frame.chartPlotFrame(f.chart, data.series, data.series.length, axis.pivotChartCategoryAxisDepth(data, f.chart))
  return axis.pivotChartCategoryAxis(data, f.chart, layout.frame, layout.plot, '#123456')
}

test.each(['owned', 'linked'].flatMap(ownership => ['row', 'column'].map(valuePosition => [ownership, valuePosition])))
  ('%s/%s projection excludes subtotal and grand-total metadata while retaining actual leaf aggregates', (ownership, valuePosition) => {
    const f = setup({ ownership, valuePosition, totals: true }), data = f.data()
    assert.equal(data.categories.length, 8)
    assert.deepEqual(data.series[0].values, [10, 20, 30, 40, 20, 10, 10, 60])
    assert.equal(data.pivotAxis.recordRows.length, data.categories.length)
    assert.ok(data.pivotAxis.recordRows.every(row => row.kind === 'detail'))
    assert.ok(data.pivotAxis.valueColumns.every(column => column.kind === 'detail'))
    assert.equal(labels(f).depth, 3)
    assert.deepEqual(labels(f).labels.filter(label => label.level === 0).map(label => label.text), ['North', 'South'])
  })

test.each(['owned', 'linked'])('%s collapsed categories carry their original typed prefix and omit unavailable descendants', ownership => {
  const f = setup({ ownership })
  pivot.applyPivotChartEdit(f.workbook, f.chart, { type: 'setCollapse', fieldId: 'region-label', collapse: true, item: 'North' })
  const data = f.data(), north = data.pivotAxis.recordRows.find(row => row.labels[0] === 'North')
  assert.equal(north.kind, 'collapsed'); assert.deepEqual(north.labels, ['North'])
  assert.equal(data.categories.length, 5); assert.equal(data.series[0].values[0], 100)
  assert.equal(labels(f).labels.filter(label => label.path[0] === 'North').length, 1)
  assert.ok(labels(f).labels.some(label => label.fieldId === 'product-label' && label.path[0] === 'South'))
})

test.each(['owned', 'linked'])('%s real Grand Total and blank members survive the exact authoritative output metadata', ownership => {
  const records = [['Grand Total', 'A', 'first', 2024, 'Q1', 'Retail', 10],
    ['', 'A', 'second', 2024, 'Q1', 'Retail', 20]]
  const f = setup({ ownership, records, totals: true }), data = f.data()
  assert.equal(data.categories.length, 2); assert.equal(data.series[0].values.reduce((sum, value) => sum + value, 0), 30)
  const parents = labels(f).labels.filter(label => label.level === 0)
  assert.ok(parents.some(label => label.text === 'Grand Total' && label.item === 'Grand Total'))
  assert.ok(parents.some(label => label.text === '（空白）' && label.item === ''))
})

test('equal displayed text from different scalar types retains two native category identities', () => {
  const records = [[1, 'A', 'first', 2024, 'Q1', 'Retail', 10], ['1', 'A', 'second', 2024, 'Q1', 'Retail', 20]]
  const f = setup({ records }), parents = labels(f).labels.filter(label => label.level === 0)
  assert.equal(parents.length, 2); assert.equal(parents[0].text, parents[1].text)
  assert.notEqual(parents[0].key, parents[1].key)
  assert.notEqual(typeof parents[0].path[0], typeof parents[1].path[0])
  assert.deepEqual(parents.map(label => [label.startIndex, label.endIndex]), [[0, 0], [1, 1]])
})

test('synthetic value rows preserve hierarchy fields separately from duplicate value names and local measure IDs', () => {
  const f = setup({ valuePosition: 'row', multipleValues: true, totals: true })
  pivot.applyPivotChartEdit(f.workbook, f.chart, { type: 'setFieldSetting', fieldId: 'amount-average', setting: { displayName: 'Amount' } })
  const data = f.data()
  assert.equal(data.series.length, 8)
  assert.equal(data.pivotAxis.rowFields.length, 3)
  assert.ok(data.pivotAxis.rowFields.every(field => field.field > 8))
  assert.equal(data.categories.length, 8)
  assert.equal(labels(f).labels.filter(label => label.level === 0).length, 2)
})

test('metadata copies isolate original source-member arrays and scalar paths from a caller', () => {
  const entry = { index: 1, kind: 'detail', level: 0, labels: ['North'], fieldFilters: { 'region-label': 'North' },
    fieldMembers: { 'region-label': ['North', 'N'] }, valueFieldId: 'amount-value' }
  const result = metadata.pivotChartAxisMetadata(fixture.pivotAxisFields().filter(field => field.area === 'row'), [], [entry], [entry])
  result.recordRows[0].labels[0] = 'Changed'; result.recordRows[0].fieldMembers['region-label'].push('New')
  result.valueColumns[0].fieldFilters['region-label'] = 'Changed'; result.rowFields[0].name = 'Changed'
  assert.deepEqual(entry.labels, ['North']); assert.deepEqual(entry.fieldMembers['region-label'], ['North', 'N'])
  assert.equal(entry.fieldFilters['region-label'], 'North')
})

test.each(['column', 'bar-horizontal', 'line', 'area', 'combination', 'waterfall'])('%s draws actual typed axis commands through the shared native graphics adapter', type => {
  const f = setup(); f.chart.type = type
  const result = labels(f), first = result.labels.find(label => label.level === 0)
  assert.ok(first); assert.equal(result.horizontal, type === 'bar-horizontal')
  const option = { graphic: { elements: graphics.chartGraphics([first.command], 'axis-label') } }
  const group = option.graphic.elements[0], text = group.children[0]
  assert.equal(text.type, 'text'); assert.equal(text.silent, false)
  assert.equal(text.info.officePivotAxisLabel.fieldId, 'region-label')
  assert.deepEqual(text.info.officePivotAxisLabel.path, first.path)
  assert.ok(group.clipPath.shape.width > 0)
})

test.each([false, true])('reversed=%s axis centers follow source spans without changing member command identity', reverse => {
  const f = setup(); f.chart.xAxis.reverse = reverse
  const parents = labels(f).labels.filter(label => label.level === 0)
  assert.equal(parents[0].item, 'North'); assert.equal(parents[1].item, 'South')
  assert.equal(parents[0].command.x < parents[1].command.x, !reverse)
})

test('manual plot rectangles remain author state while category bands clip to the actual frame', () => {
  const f = setup(); f.chart.options = { plotArea: { left: 0.2, top: 0.15, width: 0.55, height: 0.4 } }
  const without = frame.chartPlotFrame(f.chart), withHierarchy = frame.chartPlotFrame(f.chart, f.data().series, 4, 3)
  assert.deepEqual(withHierarchy.plot, without.plot)
  assert.ok(labels(f).labels.every(label => label.command.clip.x === f.chart.position.x))
})

test.each(['one-field', 'custom-mapping', 'date-axis', 'hidden-labels', 'three-d'])('%s retains its valid projection while excluding unsupported multi-level axis routing', boundary => {
  const f = setup()
  if (boundary === 'one-field') pivot.applyPivotChartEdit(f.workbook, f.chart, { type: 'removeFields', fieldIds: ['product-label', 'item-label'] })
  if (boundary === 'custom-mapping') f.chart.options = { mapping: { categoryIndexes: [1] } }
  if (boundary === 'date-axis') f.chart.options = { mapping: { useDateAxis: true } }
  if (boundary === 'hidden-labels') f.chart.xAxis.label = { visible: false }
  if (boundary === 'three-d') f.chart.type = 'column3D'
  assert.equal(axis.pivotChartCategoryAxisDepth(f.data(), f.chart), 0)
})
