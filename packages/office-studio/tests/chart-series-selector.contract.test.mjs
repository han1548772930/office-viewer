import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { waterfallComputedFixture } from './fixtures/chart-waterfall-fixtures.mjs'
import { init as initNativeChart, use as useNativeChart } from 'echarts/core'
import { ScatterChart } from 'echarts/charts'
import { GridComponent, LegendComponent } from 'echarts/components'
import { SVGRenderer } from 'echarts/renderers'

// Draft for unified validation; these contracts and native SSR have not run in this phase.
let server, selectors, cartesian, options, projection, catalog, derived, layout, draw, adapter, controls, settings, conversion, gaugeControls
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[selectors, cartesian, options, projection, catalog, derived, layout, draw, adapter, controls, settings, conversion, gaugeControls] = await Promise.all([
    'chart-series-selector', 'chart-cartesian-api', 'chart-options', 'chart-record-projection', 'chart-catalog', 'chart-derived',
    'chart-layout', 'chart-draw', 'echarts-option', 'chart-settings-controls', 'chart-series-settings', 'chart-type-conversion', 'chart-gauge-settings-controls',
  ].map(name => server.ssrLoadModule(`/packages/office-studio/src/features/charts/${name}.ts`)))
  useNativeChart([ScatterChart, GridComponent, LegendComponent, SVGRenderer])
})
afterAll(async () => { await server?.close() })

function chart(type = 'bubble', authored = {}) {
  return { id: 'chart', sheetId: 'sheet', type, source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 4, endCol: 4 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 320 }, categoryField: 0, valueFields: [1], seriesToCategory: {},
    legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear', min: 0, max: 5 }, yAxis: { scaleType: 'linear', min: 0, max: 50 }, revision: 0,
    options: { mapping: { bubble: { xIndex: 0, yIndex: 1, sizeIndex: 2, labelIndex: 3, seriesIndex: 4 } }, ...authored } }
}
function source(groups = ['North', 'South', 'North', 'South']) {
  return { fields: ['X', 'Y', 'Size', 'Label', 'Group'], records: groups.map((group, index) => [index + 1, (index + 1) * 10, (index + 2) ** 2, `Point ${index + 1}`, group]) }
}
function effective(model, series, index) { return options.chartSeriesStyle(model, series.field, index, undefined, series.sourceName ?? series.name) }
function prepared(data, model) { return derived.deriveChartData(data, model, { preparedData: data }) }
function option(data, model) { return adapter.chartEChartsOption({ sheets: [] }, model, {}, { preparedData: data }) }

test('numeric and exact source-name selectors remain separate, including case, whitespace, empty and prototype-like names', () => {
  const names = ['A', 'a', ' A ', '', '__proto__'], colors = ['red', 'blue', 'green', 'orange', 'purple']
  const model = chart('bubble', { series: { 1: { color: 'black' } }, namedSeries: names.map((name, index) => ({ name, style: { color: colors[index], name: `Display ${index}` } })) })
  const data = projection.projectChartRecords(source(names), model)
  assert.deepEqual(data.series.map(series => series.name), names)
  data.series.forEach((series, index) => {
    assert.equal(effective(model, series, index).color, colors[index])
    assert.equal(effective(model, series, index).name, `Display ${index}`)
  })
  assert.equal(options.chartSeriesStyle(model, 1, 0).color, 'black')
  assert.equal(options.chartSeriesStyle(model, 1, 0, undefined, 'A ').color, 'black')
  assert.equal(options.chartSeriesStyle(model, 1, 0, undefined, 'Display 0').color, 'black')
  assert.equal(selectors.chartSeriesOverride(model.options, { name: '__proto__' }).color, 'purple')
  assert.equal({}.color, undefined)
})

test('name patches detach inputs, merge inherited point styles by source index and reset to field and shared values', () => {
  const model = chart('bubble', { allSeriesStyle: { color: 'gray', border: { width: 3, opacity: 0.5 }, point: { size: 8 }, label: { visible: true },
    dataPoints: { 0: { color: 'green', border: { width: 6 } }, 1: { size: 14 } } },
  series: { 1: { color: 'blue', border: { color: 'navy' }, dataPoints: { 0: { color: 'orange' }, 2: { shape: 'diamond' } } } } })
  const input = { color: 'red', border: { color: null }, point: { size: 12 }, dataPoints: { 0: { size: 18, border: { color: 'purple' } } } }
  model.options = options.mergeChartOptions(model.options, cartesian.patchChartSeriesOverride(model.options, { name: 'North' }, cartesian.chartSeriesStyleValues(input)))
  input.point.size = 99; input.dataPoints[0].border.color = 'black'
  const style = effective(model, projection.projectChartRecords(source(), model).series[0], 0)
  assert.equal(style.color, 'red'); assert.deepEqual(style.border, { width: 3, opacity: 0.5, color: 'navy' }); assert.equal(style.point.size, 12)
  const point = derived.chartPointStyle(style, 0)
  assert.equal(point.color, 'orange'); assert.equal(point.point.size, 18); assert.deepEqual(point.border, { width: 6, opacity: 0.5, color: 'purple' })
  assert.equal(derived.chartPointStyle(style, 1).point.size, 14); assert.equal(derived.chartPointStyle(style, 2).point.shape, 'diamond')
  model.options = options.mergeChartOptions(model.options, cartesian.patchChartSeriesOverride(model.options, { name: 'North' }, cartesian.chartSeriesStyleValues({ color: null, point: null, dataPoints: null })))
  const reset = effective(model, projection.projectChartRecords(source(), model).series[0], 0)
  assert.equal(reset.color, 'blue'); assert.equal(reset.point.size, 8); assert.equal(derived.chartPointStyle(reset, 0).point.size, 8)
  assert.equal(derived.chartPointStyle(reset, 0).color, 'orange')
})

test('authored name collections replace explicitly while missing collections and inactive selectors survive other patches', () => {
  const original = { series: { 1: { color: 'blue' } }, namedSeries: [{ name: 'North', style: { color: 'red' } }, { name: 'Absent', style: { color: 'green' } }] }
  const retained = options.mergeChartOptions(original, { bubble: { scale: 2 } })
  assert.deepEqual(retained.namedSeries, original.namedSeries)
  selectors.validateChartNamedSeriesBindings(retained, [{ name: 'North' }])
  const replaced = options.mergeChartOptions(retained, { namedSeries: [{ name: 'South', style: { color: 'purple' } }] })
  assert.deepEqual(replaced.namedSeries.map(entry => entry.name), ['South']); assert.deepEqual(replaced.series, original.series)
  const removed = cartesian.patchChartSeriesOverride(retained, { name: 'North' }, undefined)
  assert.deepEqual(removed.namedSeries.map(entry => entry.name), ['Absent']); assert.equal(original.namedSeries.length, 2)
  assert.deepEqual(options.mergeChartOptions(retained, { namedSeries: [] }).namedSeries, [])
})

test('name validation rejects duplicate authors, sparse or malformed arrays, invalid selectors and ambiguous active source names', () => {
  for (const selector of [-1, 1_048_576, NaN, Infinity]) assert.throws(() => selectors.validateChartSeriesSelector(selector), RangeError)
  for (const selector of ['1', {}, { name: 1 }, { name: 'North', index: 1 }, [], Object.create({ name: 'North' })]) assert.throws(() => selectors.validateChartSeriesSelector(selector), TypeError)
  for (const namedSeries of [null, {}, Array(2), [undefined], [{ name: 1, style: {} }], [{ name: 'North', style: null }]]) assert.throws(() => options.cloneChartOptions({ namedSeries }))
  assert.throws(() => options.cloneChartOptions({ namedSeries: [{ name: 'North', style: {} }, { name: 'North', style: {} }] }), RangeError)
  assert.throws(() => options.cloneChartOptions({ namedSeries: [{ name: 'North', style: { point: { size: -1 } } }] }), RangeError)
  assert.throws(() => options.cloneChartOptions({ mapping: { bubble: { xIndex: 0 } } }), RangeError)
  const authored = { namedSeries: [{ name: 'North', style: { color: 'red' } }] }
  assert.throws(() => selectors.validateChartNamedSeriesBindings(authored, [{ name: 'North' }, { name: 'North' }]), /匹配多个/)
  selectors.validateChartNamedSeriesBindings(authored, [{ name: 'North' }, { name: 'north' }])
  selectors.validateChartNamedSeriesBindings(authored, [{ name: 'Display', sourceName: 'North' }])
  selectors.validateChartNamedSeriesBindings(undefined, [{ name: 'North' }, { name: 'North' }])
})

test('Bubble names drive Canvas and native scatter styles, labels, legend names and independent value axes', () => {
  const model = chart('bubble', { allSeriesStyle: { point: { size: 8 }, label: { visible: true, contentType: 'series' } }, series: { 1: { color: 'blue' } },
    rightYAxis: { scaleType: 'linear', min: 0, max: 100 }, namedSeries: [{ name: 'North', style: { name: 'North display', color: 'red', rightAxis: true, point: { size: 12 }, dataPoints: { 0: { color: 'purple', size: 18, shape: 'rect' } } } }] })
  const data = projection.projectChartRecords(source(), model), result = layout.layoutChartData(data, model, { preparedData: data })
  assert.deepEqual(data.series.map(series => series.name), ['North', 'South'])
  assert.deepEqual(result.series.map(series => [series.name, series.color, Boolean(series.rightAxis)]), [['North display', 'red', true], ['South', 'blue', false]])
  assert.equal(result.series[0].points[0].y, result.plot.y + result.plot.height * 0.9)
  assert.equal(result.series[1].points[0].y, result.plot.y + result.plot.height * 0.6)
  const commands = draw.drawDerivedChartData(prepared(data, model), model)
  assert.ok(commands.some(command => command.kind === 'rect' && command.width === 18 && command.fill === 'rgba(128, 0, 128, 1)'))
  assert.ok(commands.some(command => command.kind === 'text' && command.text === 'North display'))
  const nativeOption = option(data, model)
  assert.deepEqual(nativeOption.legend.data, ['North display', 'South'])
  assert.deepEqual(nativeOption.series.map(series => [series.name, series.yAxisIndex, series.data.length]), [['North display', 1, 2], ['South', 0, 2]])
  const native = initNativeChart(null, null, { renderer: 'svg', ssr: true, width: 480, height: 320 })
  delete nativeOption.officeChartKey
  try {
    native.setOption(nativeOption, { notMerge: true })
    const first = native.getModel().getSeriesByIndex(0), second = native.getModel().getSeriesByIndex(1)
    assert.equal(first.name, 'North display'); assert.equal(first.getData().count(), 2); assert.equal(second.getData().count(), 2)
    assert.equal(first.getData().getItemVisual(0, 'style').fill, 'purple')
    assert.equal(first.getData().getItemVisual(0, 'symbol'), 'rect'); assert.equal(first.getData().getItemVisual(0, 'symbolSize'), 18)
    assert.equal(first.coordinateSystem.getAxis('y').index, 1); assert.equal(second.coordinateSystem.getAxis('y').index, 0)
    assert.ok(native.renderToSVGString().includes('North display'))
  } finally { native.dispose() }
})

test('inactive names and explicit primary overrides do not allocate a secondary-axis plot margin', () => {
  const data = projection.projectChartRecords(source(), chart()), plain = chart()
  const inactive = chart('bubble', { namedSeries: [{ name: 'Absent', style: { rightAxis: true } }] })
  assert.deepEqual(layout.chartPlotFrame(inactive, data.series), layout.chartPlotFrame(plain, data.series))
  const restored = chart('bubble', { series: { 1: { rightAxis: true } }, namedSeries: [{ name: 'North', style: { rightAxis: false } }, { name: 'South', style: { rightAxis: false } }] })
  assert.deepEqual(layout.chartPlotFrame(restored, data.series), layout.chartPlotFrame(plain, data.series))
})

test('settings expose independent source names, retain inactive targets and limit Bubble point editing to the selected group', () => {
  const model = chart('bubble', { series: { 1: { color: 'blue', point: { size: 8 } } }, namedSeries: [{ name: 'Absent', style: { color: 'green' } }] })
  const data = projection.projectChartRecords(source(), model), fields = [{ value: 1, label: 'Y' }]
  const targets = settings.chartSeriesSettingTargets(model, fields, data)
  assert.deepEqual(targets.map(target => target.selector), [1, { name: 'North' }, { name: 'South' }, { name: 'Absent' }])
  assert.equal(new Set(targets.map(target => target.value)).size, 4)
  assert.deepEqual(settings.chartSeriesSettingPoints(model, { name: 'North' }, data.categories, fields, data).map(item => item.value), [0, 2])
  assert.deepEqual(settings.chartSeriesSettingPoints(model, { name: 'South' }, data.categories, fields, data).map(item => item.value), [1, 3])
  assert.equal(settings.chartSeriesSettingPoints(model, 1, data.categories, fields, data).length, 4)
  assert.equal(settings.chartSeriesSettingPoints(model, { name: 'Absent' }, data.categories, fields, data).length, 0)
  const north = targets.find(target => target.selector.name === 'North')
  const fieldsForNorth = controls.chartSeriesControls(north.selector, false, true, false, north.source)
  assert.equal(fieldsForNorth.find(control => control.key.endsWith('.color')).read(model), 'blue')
  fieldsForNorth.find(control => control.key.endsWith('.point.size')).write(model, 16)
  assert.equal(selectors.chartSeriesOverride(model.options, { name: 'North' }).point.size, 16)
  assert.equal(model.options.series[1].point.size, 8)
  const duplicate = { ...data, series: [{ ...data.series[0], name: 'Duplicate' }, { ...data.series[1], name: 'Duplicate' }] }
  assert.equal(settings.chartSeriesSettingTargets(model, fields, duplicate).find(target => target.selector.name === 'Duplicate').disabled, true)
  const recoverable = chart('bubble', { namedSeries: [{ name: 'Duplicate', style: { color: 'red' } }] })
  const recoverableTarget = settings.chartSeriesSettingTargets(recoverable, fields, duplicate).find(target => target.selector.name === 'Duplicate')
  assert.equal(recoverableTarget.ambiguous, true); assert.equal(recoverableTarget.disabled, false)
  assert.deepEqual(settings.chartSeriesSettingPoints(recoverable, { name: 'Duplicate' }, duplicate.categories, fields, duplicate), [])
  const recovered = options.mergeChartOptions(recoverable.options, cartesian.patchChartSeriesOverride(recoverable.options, { name: 'Duplicate' }, undefined))
  selectors.validateChartNamedSeriesBindings(recovered, duplicate.series)
})

test('Boxplot renamed labels retain source-name identities for distribution styles, axes and point settings', () => {
  const data = { categories: ['A', 'B', 'C', 'D', 'E', 'F'], series: [{ field: 9, name: 'First samples', values: [1, 2, 3, 4, 5, 100] },
    { field: 2, name: 'Second samples', values: [10, 20, 30, 40, 50, 1000] }] }
  const model = chart('boxplot', { rightYAxis: { scaleType: 'linear', min: 0, max: 2000 }, series: { 2: { color: 'blue' } },
    namedSeries: [{ name: 'Second samples', style: { name: 'Second display', color: 'purple', rightAxis: true, label: { visible: true, contentType: 'series' } } }] })
  const result = prepared(data, model), scene = layout.layoutDerivedChartData(result, model)
  assert.equal(result.boxes[1].sourceName, 'Second samples'); assert.equal(result.boxes[1].name, 'Second display')
  assert.equal(scene.series[1].name, 'Second display'); assert.equal(scene.series[1].color, 'purple'); assert.equal(scene.series[1].rightAxis, true)
  assert.ok(draw.drawDerivedChartData(result, model).some(command => command.kind === 'text' && command.text === 'Second display'))
  const sampleChoices = settings.chartSeriesSettingPoints(model, { name: 'Second samples' }, data.series.map(series => series.name), [{ value: 9, label: 'First' }, { value: 2, label: 'Second' }], data)
  assert.deepEqual(sampleChoices.map(item => item.value), [0, 1, 2, 3, 4, 5])
  assert.equal(sampleChoices.at(-1).label, '6 · F · 1000')
})

test('named combination types control stacking and statistical eligibility and convert with field-style axis rules', () => {
  const data = { categories: ['A', 'B', 'C', 'D'], series: [{ field: 1, name: 'North', values: [1, 2, 3, 4] }, { field: 2, name: 'South', values: [2, 4, 6, 8] }] }
  const model = chart('combination', { series: { 2: { type: 'line' } }, namedSeries: [{ name: 'South', style: { type: 'areaStacked', rightAxis: true } }],
    trendlines: [{ seriesIndex: 1, type: 'linear' }, { seriesIndex: 2, type: 'linear' }] })
  assert.deepEqual(derived.chartSeriesKind(model, 2, 1, 'South'), { type: 'area', stacked: true })
  assert.equal(catalog.supportsChartStatistics(model, 2, 1, 'South'), false)
  assert.equal(catalog.supportsChartStatistics(model, 1, 0, 'North'), true)
  assert.deepEqual(prepared(data, model).trends.map(trend => trend.field), [1])
  const converted = conversion.convertChartType(model, 'radar')
  assert.equal(converted.options.namedSeries[0].style.rightAxis, false)
  assert.equal(converted.options.namedSeries[0].name, 'South'); assert.equal(model.options.namedSeries[0].style.rightAxis, true)
})

test('named solid colors survive specialized color transforms and Gauge controls inherit and restore the selected source font', () => {
  const authored = { namedSeries: [{ name: 'Weight', style: { color: 'purple' } }] }
  const palette = ['blue', 'red'], hierarchy = chart('sunburst', authored), heatmap = chart('heatmap', authored), candle = chart('candlestick', authored)
  assert.equal(derived.chartHierarchyPointStyle(hierarchy, 1, { pointIndex: 0, rootIndex: 0, depth: 3 }, palette, 'Weight').color, 'purple')
  assert.equal(derived.chartHeatmapPointStyle(heatmap, 1, 0, 0, 5, { min: 0, max: 10 }, palette, 'Weight').color, 'purple')
  assert.equal(derived.chartCandlestickPointStyle(candle, 1, 0, true, palette, 'Weight').color, 'purple')
  const waterfall = chart('waterfall', { ...authored, waterfall: { useSubtotal: false } }), values = { categories: ['A', 'B', 'C'], series: [{ field: 1, name: 'Weight', values: [2, -1, 4] }] }
  const waterfallScene = draw.drawDerivedChartData(derived.deriveChartData(values, waterfall, waterfallComputedFixture(values, 'named-changes')), waterfall)
  assert.equal(waterfallScene.filter(command => command.kind === 'rect' && command.fill === 'rgba(128, 0, 128, 1)').length, 3)
  const gauge = chart('gauge', { appearance: { font: { family: 'Aptos', fontSize: 16 } }, series: { 1: { label: { color: 'blue' } } },
    namedSeries: [{ name: 'Weight', style: { label: { family: 'Georgia', fontSize: 22, color: 'purple' } } }] })
  const read = (key) => gaugeControls.chartGaugeSettingsControls(gauge, 5, 'Weight').find(control => control.key === key).read(gauge)
  assert.equal(read('gauge.detail.family'), 'Georgia'); assert.equal(read('gauge.detail.fontSize'), 22); assert.equal(read('gauge.detail.color'), 'purple')
  gauge.options = options.mergeChartOptions(gauge.options, cartesian.patchChartSeriesOverride(gauge.options, { name: 'Weight' }, cartesian.chartSeriesStyleValues({ label: null })))
  assert.equal(read('gauge.detail.family'), 'Aptos'); assert.equal(read('gauge.detail.fontSize'), 16); assert.equal(read('gauge.detail.color'), 'blue')
})
