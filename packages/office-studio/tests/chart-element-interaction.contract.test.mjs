import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartCanvasElementAsSeries, chartCanvasElementKey, chartCanvasElementRoute, chartCanvasNativeElementRuntime,
  chartCanvasNativePointOrder, chartCanvasRuntimeSeriesId, chartCanvasSeriesBindings, clipChartCanvasPolygon, readChartCanvasElement } from '../src/features/charts/chart-element-interaction.ts'
import { chartDateAxis, prepareChartDateData } from '../src/features/charts/chart-date-axis.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { chartLegendViewData } from '../src/features/charts/chart-legend.ts'

// P19 drafts only. Execution, type checks and browser acceptance are deferred.
const model = (type = 'line', patch = {}) => ({ id: 'elements', sheetId: 'sheet', type, revision: 0,
  categoryField: 0, valueFields: [1], source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } },
  position: { x: 0, y: 0 }, size: { width: 480, height: 320 }, legend: { visible: true, position: 'bottom' },
  xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, ...patch })
const source = () => ({ categories: ['A', 'B', 'C'], series: [{ field: 1, name: 'Displayed', sourceName: 'Original', values: [2, 4, 6] }] })
const json = value => JSON.parse(JSON.stringify(value))

test('runtime positions map to existing field IDs while a display rename retains the original source identity', () => {
  const data = source(), binding = chartCanvasSeriesBindings(data)[0]
  assert.deepEqual(binding, { seriesId: 'field:1', selector: 1, field: 1, sourcePosition: 0, sourceName: 'Original' })
  data.series[0].name = 'Renamed again'
  assert.deepEqual(chartCanvasSeriesBindings(data)[0], binding)
  assert.equal(chartCanvasRuntimeSeriesId('a:0', 2), 'office-chart-series:["a:0",2]')
  assert.notEqual(chartCanvasRuntimeSeriesId('a:0', 2), binding.seriesId)
  const named = chartCanvasSeriesBindings(data, { series: { 1: { color: '#111111' } }, namedSeries: [{ name: 'Original', style: { color: '#222222' } }] })[0]
  assert.deepEqual(named.selector, { name: 'Original' }); assert.equal(named.seriesId, 'name:"Original"')
  assert.equal(chartCanvasNativeElementRuntime(data, model('line', { options: { namedSeries: [{ name: 'Original', style: { point: { size: 10 } } }] } })).series[0].binding.seriesId, named.seriesId)
})

test('duplicate source fields use exact unique source names and doubly ambiguous sources acquire no direct author target', () => {
  const data = { series: [{ field: 3, name: 'Same display', sourceName: ' Alpha ', values: [1] },
    { field: 3, name: 'Same display', sourceName: 'Beta', values: [2] }] }
  const bindings = chartCanvasSeriesBindings(data)
  assert.deepEqual(bindings.map(binding => binding.selector), [{ name: ' Alpha ' }, { name: 'Beta' }])
  assert.deepEqual(bindings.map(binding => binding.seriesId), ['name:" Alpha "', 'name:"Beta"'])
  data.series[1].sourceName = ' Alpha '
  assert.deepEqual(chartCanvasSeriesBindings(data), [undefined, undefined])
})

test('native dates retain null slots and explicitly map chronological input order back to semantic point indices', () => {
  const chart = model('line', { options: { mapping: { useDateAxis: true } } })
  const data = prepareChartDateData({ ...source(), categoryValues: [63, 59, 61], series: [{ field: 1, name: 'Units', values: [6, null, 4] }] }, chart)
  const dates = chartDateAxis(data, chart), runtime = chartCanvasNativeElementRuntime(data, chart, dates)
  assert.deepEqual(chartCanvasNativePointOrder(data, chart, 0, dates), [1, 2, 0])
  assert.deepEqual(runtime.series[0].dataIndices, [1, 2, 0])
  assert.equal(data.series[0].values[runtime.series[0].dataIndices[0]], null)
  assert.equal(runtime.series[0].binding.seriesId, 'field:1')
})

test('scatter and bubble identity maps follow the same finite, log-axis and size filtering as native incoming data', () => {
  const data = { categories: ['A', 'B', 'C', 'D'], series: [{ field: 2, name: 'Points', values: [5, 6, null, 8],
    xValues: [1, 0, 3, 4], sizeValues: [2, 2, 1, 0] }] }
  const scatter = model('scatter', { xAxis: { scaleType: 'log' }, yAxis: { scaleType: 'linear' } })
  assert.deepEqual(chartCanvasNativePointOrder(data, scatter, 0), [0, 3])
  assert.deepEqual(chartCanvasNativePointOrder(data, { ...scatter, type: 'bubble' }, 0), [0])
  assert.deepEqual(chartCanvasNativeElementRuntime(data, { ...scatter, type: 'bubble' }).series[0].dataIndices, [0])
  assert.deepEqual(chartCanvasNativePointOrder(data, scatter, 8), [])
})

test('renderer metadata reconstructs validated targets and rejects inconsistent selectors, runtime IDs and malformed points', () => {
  const binding = chartCanvasSeriesBindings(source())[0], value = { kind: 'dataPoint', binding, dataIndex: 2 }
  const read = readChartCanvasElement(value)
  assert.deepEqual(read, value); assert.notEqual(read.binding, binding)
  for (const invalid of [null, { kind: 'title' }, { kind: 'axis', axisId: 'depthAxis', part: 'line' },
    { ...value, dataIndex: -1 }, { ...value, dataIndex: 1.5 }, { ...value, binding: { ...binding, selector: 7 } },
    { ...value, binding: { ...binding, seriesId: chartCanvasRuntimeSeriesId('elements', 0) } },
    { ...value, binding: { ...binding, selector: { name: 'Wrong' }, seriesId: 'name:"Wrong"' } }]) assert.equal(readChartCanvasElement(invalid), undefined)
  assert.deepEqual(readChartCanvasElement({ kind: 'axisTitle', axisId: 'rightYAxis' }), { kind: 'axisTitle', axisId: 'rightYAxis' })
})

test('selection bounds clip rotated rectangles in both winding orders and reject disjoint native clip regions', () => {
  const rect = [{ x: 0, y: 0 }, { x: 10, y: 0 }, { x: 10, y: 10 }, { x: 0, y: 10 }]
  const diamond = [{ x: 5, y: -2 }, { x: 12, y: 5 }, { x: 5, y: 12 }, { x: -2, y: 5 }]
  const clipped = clipChartCanvasPolygon(rect, diamond)
  assert.equal(clipped.length, 8)
  assert.ok(clipped.every(point => point.x >= 0 && point.x <= 10 && point.y >= 0 && point.y <= 10))
  const unordered = points => points.map(point => `${point.x.toFixed(6)},${point.y.toFixed(6)}`).sort()
  assert.deepEqual(unordered(clipChartCanvasPolygon(rect, [...diamond].reverse())), unordered(clipped))
  assert.deepEqual(clipChartCanvasPolygon(rect, rect.map(point => ({ x: point.x + 20, y: point.y }))), [])
})

test('horizontal and secondary axes map to semantic axis IDs while native radar data stays series-only', () => {
  const data = source(), chart = model('bar')
  const runtime = chartCanvasNativeElementRuntime(data, chart, undefined, { xAxis: [{}, {}], yAxis: {} })
  assert.deepEqual(runtime.axes, [{ component: 'xAxis', index: 0, axisId: 'yAxis' },
    { component: 'xAxis', index: 1, axisId: 'rightYAxis' }, { component: 'yAxis', index: 0, axisId: 'xAxis' }])
  assert.equal(chartCanvasNativeElementRuntime(data, model('radar')).series[0].seriesOnly, true)
  const pareto = chartCanvasNativeElementRuntime(data, model('pareto')).series
  assert.deepEqual(pareto.map(entry => ({ kind: entry.directEdit.kind, part: entry.directEdit.part })), [{ kind: 'paretoStyle', part: 'bar' }, { kind: 'paretoStyle', part: 'line' }])
  assert.ok(pareto.every(entry => !Object.hasOwn(entry, 'binding')))
})

test('legend masking preserves source positions and duplicate-field selectors, and element More routes address their real scope', () => {
  const chart = model(), data = { categories: ['A'], series: [{ field: 1, name: 'Alpha', values: [1] }, { field: 1, name: 'Beta', values: [2] }] }
  const derived = deriveChartData(data, chart, { preparedData: data }), before = json(data)
  const view = chartLegendViewData(derived, chart, { Alpha: false }, ['#123456', '#abcdef'])
  assert.equal(view.data.series.length, 2); assert.deepEqual(view.data.series[0].values, [null]); assert.deepEqual(data, before)
  assert.deepEqual(chartCanvasSeriesBindings(view.data), chartCanvasSeriesBindings(data))
  const point = { kind: 'dataPoint', binding: chartCanvasSeriesBindings(data)[1], dataIndex: 0 }
  assert.equal(chartCanvasElementKey(point), 'dataPoint:name:"Beta":0')
  assert.deepEqual(chartCanvasElementAsSeries(point), { kind: 'series', binding: point.binding })
  assert.deepEqual(chartCanvasElementRoute(point), { tab: 'style', section: 'series', seriesId: 'name:"Beta"' })
  assert.deepEqual(chartCanvasElementRoute({ ...point, kind: 'dataLabel' }, 'pie'), { tab: 'style', section: 'pieStyle' })
})
