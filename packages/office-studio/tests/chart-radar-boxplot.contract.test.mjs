import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { init as initNativeChart, use as useNativeChart } from 'echarts/core'
import { BoxplotChart, RadarChart, ScatterChart } from 'echarts/charts'
import { GridComponent, RadarComponent } from 'echarts/components'
import { SVGRenderer } from 'echarts/renderers'

// Draft for unified validation; neither these contracts nor native SSR have run in this phase.
let server, radar, boxplot, cartesian, options, derived, layout, draw, adapter, controls, color
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[radar, boxplot, cartesian, options, derived, layout, draw, adapter, controls, color] = await Promise.all([
    'chart-radar', 'chart-boxplot', 'chart-cartesian-api', 'chart-options', 'chart-derived', 'chart-layout',
    'chart-draw', 'echarts-option', 'chart-radar-settings-controls', 'chart-color',
  ].map(name => server.ssrLoadModule(`/packages/office-studio/src/features/charts/${name}.ts`)))
  useNativeChart([BoxplotChart, RadarChart, ScatterChart, GridComponent, RadarComponent, SVGRenderer])
})
afterAll(async () => { await server?.close() })

function chart(type = 'radar', patch = {}) {
  return { id: 'chart', sheetId: 'sheet', type, source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 6, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 320 }, categoryField: 0, valueFields: [1, 2],
    legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, revision: 0, ...patch }
}
const radarSource = () => ({ categories: ['North', 'East', 'South', 'West'], series: [{ field: 1, name: 'Score', values: [0, 25, 50, 100] }] })
const boxSource = () => ({ categories: ['A', 'B', 'C', 'D', 'E', 'F'], series: [
  { field: 9, name: 'First samples', values: [1, 2, 3, 4, 5, 100] },
  { field: 2, name: 'Second samples', values: [10, 20, 30, 40, 50, 1000] },
] })
function prepared(source, model) { return derived.deriveChartData(source, model, { preparedData: source }) }
function scene(source, model) { return draw.drawDerivedChartData(prepared(source, model), model) }
function option(source, model) { return adapter.chartEChartsOption({ sheets: [] }, model, {}, { preparedData: source }) }
function near(actual, expected, tolerance = 1e-8) { assert.ok(Math.abs(actual - expected) < tolerance, `${actual} ≈ ${expected}`) }
function assertFiniteScene(commands) {
  for (const command of commands) {
    const coordinates = command.kind === 'polygon' ? command.points.flatMap(point => [point.x, point.y])
      : command.kind === 'rect' ? [command.x, command.y, command.width, command.height]
        : command.kind === 'line' ? [command.x1, command.y1, command.x2, command.y2] : [command.x, command.y]
    assert.ok(coordinates.every(Number.isFinite), command.kind)
  }
}
function nativeChart(model, source) {
  const instance = initNativeChart(null, null, { renderer: 'svg', ssr: true, width: model.size.width, height: model.size.height })
  const value = option(source, model)
  delete value.officeChartKey
  try { instance.setOption(value, { notMerge: true }) }
  catch (error) { instance.dispose(); throw error }
  return instance
}

test('Radar nullable label patches detach inputs and restore appearance and value-axis inheritance', () => {
  const input = { axisLabel: { fontSize: null, color: null, italic: false } }
  const current = { shape: 'circle', fill: true, axisLabel: { visible: true, fontSize: 22, color: 'red', family: 'Georgia', bold: true } }
  const patched = radar.patchChartRadarOptions(current, input)
  assert.deepEqual(patched, { shape: 'circle', fill: true, axisLabel: { visible: true, fontSize: undefined, color: undefined, family: 'Georgia', bold: true, italic: false } })
  assert.deepEqual(input, { axisLabel: { fontSize: null, color: null, italic: false } })
  input.axisLabel.italic = true
  assert.equal(patched.axisLabel.italic, false)
  const model = chart('radar', { options: { appearance: { font: { family: 'Aptos', fontSize: 16, color: 'navy', italic: true } }, radar: patched },
    yAxis: { scaleType: 'linear', format: '0.00', label: { family: undefined, fontSize: undefined, color: 'green' } } })
  assert.deepEqual(radar.chartRadarAxisLabel(model), { family: 'Georgia', fontSize: 16, color: 'green', italic: false, format: '0.00', visible: true, bold: true })
  model.options.radar = radar.patchChartRadarOptions(patched, { axisLabel: null })
  assert.equal(radar.chartRadarAxisLabel(model).visible, true)
  assert.equal(radar.chartRadarAxisLabel(model).family, 'Aptos')
  assert.equal(radar.chartRadarAxisLabel(model).fontSize, 16)
  assert.equal(radar.patchChartRadarOptions(current, undefined), undefined)
})

test('Radar controls show numeric-label format and font only while visible and constrain the radial scale to linear', () => {
  const model = chart('radar', { options: { appearance: { font: { family: 'Georgia', fontSize: 16 } }, radar: { axisLabel: { visible: false } } } })
  const hiddenFields = controls.chartRadarSettingsControls(model)
  assert.deepEqual(hiddenFields.map(field => field.key), ['radar.shape', 'radar.fill', 'radar.axisLabel.visible'])
  hiddenFields.find(field => field.key === 'radar.axisLabel.visible').write(model, true)
  const fields = controls.chartRadarSettingsControls(model)
  assert.deepEqual(fields.map(field => field.key), ['radar.shape', 'radar.fill', 'radar.axisLabel.visible', 'radar.axisLabel.format',
    'radar.axisLabel.family', 'radar.axisLabel.fontSize', 'radar.axisLabel.color', 'radar.axisLabel.bold', 'radar.axisLabel.italic'])
  const font = fields.find(field => field.key === 'radar.axisLabel.fontSize')
  assert.equal(font.read(model), 16)
  font.write(model, 22)
  assert.equal(model.options.radar.axisLabel.fontSize, 22)
  font.write(model, undefined)
  assert.equal(font.read(model), 16)
  fields.find(field => field.key === 'radar.axisLabel.format').write(model, '0.0%')
  assert.equal(radar.chartRadarAxisLabel(model).format, '0.0%')
  fields.find(field => field.key === 'radar.axisLabel.visible').write(model, false)
  assert.deepEqual(controls.chartRadarSettingsControls(model).map(field => field.key), hiddenFields.map(field => field.key))
  assert.equal(model.options.radar.axisLabel.format, '0.0%')
  model.yAxis.label = { visible: true }
  assert.equal(controls.chartRadarSettingsControls(model).length, 3)
  model.options.radar.axisLabel.visible = undefined
  assert.equal(controls.chartRadarSettingsControls(model).length, 9)
  const categoryFields = controls.chartRadarCategoryAxisControls(), valueFields = controls.chartRadarValueAxisControls()
  assert.ok(categoryFields.some(field => field.key === 'xAxis.reverse'))
  assert.ok(categoryFields.some(field => field.key === 'xAxis.label.rotate'))
  assert.ok(!valueFields.some(field => field.key === 'yAxis.scaleType' || field.key === 'yAxis.logBase'))
  valueFields.find(field => field.key === 'yAxis.min').write(model, -10)
  valueFields.find(field => field.key === 'yAxis.max').write(model, 90)
  assert.deepEqual(radar.chartRadarExtent(radarSource(), model.yAxis), { min: -10, max: 90 })
  assert.throws(() => radar.chartRadarExtent(radarSource(), { scaleType: 'log' }), /线性/)
})

test('Radar uses one signed range and preserves category indices through missing values and both reversals', () => {
  const source = { categories: ['North', 'East', 'South', 'West'], series: [{ field: 1, name: 'Score', values: [-20, 30, null, 80] }] }
  const model = chart('radar', { yAxis: { scaleType: 'linear', min: -20, max: 80 } })
  const result = prepared(source, model), geometry = layout.layoutDerivedChartData(result, model), polar = radar.chartRadarGeometry(source, model, geometry.plot)
  assert.deepEqual(polar.ticks, [-20, 0, 20, 40, 60, 80])
  assert.deepEqual(geometry.series[0].points.map(point => point.categoryIndex), [0, 1, 3])
  near(geometry.series[0].points[0].x, polar.center.x)
  near(geometry.series[0].points[0].y, polar.center.y)
  near(geometry.series[0].points[1].x, polar.center.x + polar.radius / 2)
  near(geometry.series[0].points[1].y, polar.center.y)
  model.xAxis.reverse = true
  const reversed = radar.chartRadarGeometry(source, model, geometry.plot)
  near(reversed.point(1, 30).x, polar.center.x - polar.radius / 2)
  model.yAxis.reverse = true
  const radial = radar.chartRadarGeometry(source, model, geometry.plot)
  near(radial.point(0, -20).y, polar.center.y - polar.radius)
  near(radial.point(3, 80).x, polar.center.x)
  near(radial.point(3, 80).y, polar.center.y)
  assert.deepEqual(source.series[0].values, [-20, 30, null, 80])
  const extreme = chart('radar', { yAxis: { scaleType: 'linear', min: -Number.MAX_VALUE, max: Number.MAX_VALUE } })
  const extremeSource = { ...source, series: [{ field: 1, name: 'Finite', values: [-Number.MAX_VALUE, 0, Number.MAX_VALUE, 1] }] }
  assertFiniteScene(scene(extremeSource, extreme))
  assert.ok(option(extremeSource, extreme).graphic)
})

test('native Radar draws scale-label text and uses the same bounds, direction and radius', () => {
  const source = radarSource(), model = chart('radar', { xAxis: { scaleType: 'linear', reverse: true }, yAxis: { scaleType: 'linear', min: 0, max: 100 },
    options: { radar: { shape: 'circle', axisLabel: { visible: true, format: '0.0', family: 'Georgia', color: 'navy', bold: true } } } })
  const value = option(source, model), frame = layout.chartPlotFrame(model), polar = radar.chartRadarGeometry(source, model, frame.plot)
  assert.equal(value.radar.axisLabel.show, true)
  assert.equal(value.radar.axisLabel.formatter(25), '25.0')
  assert.equal(value.radar.clockwise, false)
  near(value.radar.radius, polar.radius)
  const native = nativeChart(model, source)
  try {
    const coordinate = native.getModel().getComponent('radar').coordinateSystem
    for (const axis of coordinate.getIndicatorAxes()) assert.deepEqual(axis.scale.getExtent(), [0, 100])
    const east = coordinate.dataToPoint(25, 1), expected = polar.point(1, 25)
    near(east[0], expected.x, 1e-4); near(east[1], expected.y, 1e-4)
    const display = native.getZr().storage.getDisplayList(), labels = display.filter(element => element.type === 'tspan')
    assert.ok(labels.some(element => element.style.text === '20.0' && /Georgia/.test(element.style.font ?? '')))
    assert.ok(labels.some(element => element.style.text === '100.0'))
    assert.ok(display.some(element => element.type === 'circle' && Math.abs(element.shape.r - polar.radius) < 1e-4))
    assert.ok(!/NaN|Infinity/.test(native.renderToSVGString()))
  } finally { native.dispose() }
  model.options.radar.axisLabel.visible = false
  assert.equal(option(source, model).radar.axisLabel.show, false)
})

test('Radar graphic rendering consumes category rotation, radial reversal and per-point style and value labels', () => {
  const source = radarSource(), model = chart('radar', { xAxis: { scaleType: 'linear', label: { rotate: 23 } }, yAxis: { scaleType: 'linear', reverse: true, min: 0, max: 100 },
    options: { radar: { fill: true, axisLabel: { visible: true, format: '0.0', color: 'navy' } },
      series: { 1: { label: { visible: true, format: '0.00', color: 'green' }, dataPoints: { 2: { color: 'magenta', shape: 'rect', size: 10 } } } } } })
  const commands = scene(source, model)
  assertFiniteScene(commands)
  assert.ok(commands.some(command => command.kind === 'text' && command.text === 'East' && command.rotation === 23))
  assert.ok(commands.some(command => command.kind === 'text' && command.text === '25.00' && command.color === 'rgba(0, 128, 0, 1)'))
  assert.ok(commands.some(command => command.kind === 'text' && command.text === '0.0' && command.color === 'rgba(0, 0, 128, 1)'))
  assert.ok(commands.some(command => command.kind === 'rect' && command.width === 10 && command.height === 10 && command.fill === 'rgba(255, 0, 255, 1)'))
  assert.equal(option(source, model).series, undefined)
  assert.ok(option(source, model).graphic.elements.length > 0)
})

test('Boxplot distributions use raw source-field identity and independent fill, border and outlier styles', () => {
  const source = boxSource(), model = chart('boxplot', { valueFields: [9, 2], options: { boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr' }, allSeriesStyle: { color: 'gold', fillOpacity: 0.6,
    border: { color: 'navy', width: 3, opacity: 0.5 }, point: { shape: 'rect', size: 10, color: 'magenta' } },
  series: { 9: { name: 'Authored first', color: 'red', fillOpacity: 0.25, border: { width: 5, opacity: 0.8 } } } } })
  const result = prepared(source, model), geometry = layout.layoutDerivedChartData(result, model), value = option(source, model)
  assert.deepEqual(result.boxes.map(box => [box.field, box.name]), [[9, 'Authored first'], [2, 'Second samples']])
  assert.deepEqual(geometry.series.map(series => series.field), [9, 2])
  assert.deepEqual(geometry.series.map(series => series.points[0].value), [3.5, 35])
  assert.deepEqual(value.series[0].data[0].value, [1, 2.25, 3.5, 4.75, 5])
  assert.equal(value.series[0].data[0].name, 'Authored first')
  const style = value.series[0].data[0].itemStyle
  assert.equal(style.color, 'rgba(255, 0, 0, 0.25)')
  assert.equal(style.borderColor, 'rgba(0, 0, 128, 0.8)')
  assert.equal(style.borderWidth, 5); assert.equal(style.opacity, 1)
  const outlier = value.series[1].data[0]
  assert.deepEqual(outlier.value, [0, 100]); assert.equal(outlier.symbol, 'rect'); assert.equal(outlier.symbolSize, 10)
  assert.equal(outlier.itemStyle.color, 'rgba(255, 0, 255, 0.25)')
  assert.equal(outlier.itemStyle.borderColor, 'rgba(0, 0, 128, 0.8)'); assert.equal(outlier.itemStyle.opacity, 1)
  const commands = scene(source, model), box = commands.find(command => command.kind === 'rect' && command.fill === 'rgba(255, 0, 0, 0.25)')
  assert.ok(box)
  const endpoints = commands.filter(command => command.kind === 'line' && command.color === 'rgba(0, 0, 128, 0.8)' && command.y1 === command.y2)
  assert.ok(endpoints.some(command => Math.abs(Math.abs(command.x2 - command.x1) - box.width) < 1e-8), 'Whisker caps span the box width')
  assert.ok(commands.some(command => command.kind === 'rect' && command.width === 10 && command.fill === 'rgba(255, 0, 255, 0.25)'))
  assertFiniteScene(commands)
})

test('native Boxplot emits independently styled paths and full-width whisker caps', () => {
  const model = chart('boxplot', { valueFields: [9, 2], options: { series: { 9: { color: 'red', fillOpacity: 0.25, border: { color: 'navy', opacity: 0.8, width: 5 } } } } })
  const native = nativeChart(model, boxSource())
  try {
    const data = native.getModel().getSeriesByIndex(0).getData(), path = data.getItemGraphicEl(0), ends = data.getItemLayout(0).ends
    assert.equal(color.normalizeChartColor(path.style.fill), 'rgba(255, 0, 0, 0.25)')
    assert.equal(color.normalizeChartColor(path.style.stroke), 'rgba(0, 0, 128, 0.8)')
    assert.equal(path.style.opacity, 1); assert.equal(path.style.lineWidth, 5)
    const bodyWidth = Math.abs(ends[1][0] - ends[0][0])
    near(Math.abs(ends[9][0] - ends[8][0]), bodyWidth, 1e-4)
    near(Math.abs(ends[11][0] - ends[10][0]), bodyWidth, 1e-4)
    assert.ok(!/NaN|Infinity/.test(native.renderToSVGString()))
  } finally { native.dispose() }
})

test('Boxplot median anchors and labels follow independent primary and secondary axes and category reversal', () => {
  const source = { categories: ['A', 'B', 'C'], series: [{ field: 9, name: 'Left', values: [10, 20, 30] }, { field: 2, name: 'Right', values: [100, 200, 300] }] }
  const model = chart('boxplot', { valueFields: [9, 2], xAxis: { scaleType: 'linear', reverse: true }, yAxis: { scaleType: 'linear', min: 0, max: 100 },
    options: { rightYAxis: { scaleType: 'linear', min: 100, max: 500, reverse: true }, series: { 2: { rightAxis: true, label: { visible: true, position: 'right', format: '0.0', color: 'green' } } } } })
  const result = prepared(source, model), geometry = layout.layoutDerivedChartData(result, model), left = geometry.series[0].points[0], right = geometry.series[1].points[0]
  assert.deepEqual([geometry.min, geometry.max, geometry.rightMin, geometry.rightMax], [0, 100, 100, 500])
  assert.ok(left.x > right.x)
  near(left.y, geometry.plot.y + geometry.plot.height * 0.8)
  near(right.y, geometry.plot.y + geometry.plot.height * 0.25)
  const label = scene(source, model).find(command => command.kind === 'text' && command.text === '200.0' && command.color === 'rgba(0, 128, 0, 1)')
  assert.ok(label); near(label.y, right.y)
  assert.ok(label.x > right.x)
  assert.ok(option(source, model).graphic)
})

test('Boxplot logarithmic projection retains original statistics and excludes only non-drawable summaries and outliers', () => {
  const source = { categories: ['A', 'B', 'C', 'D', 'E', 'F'], series: [
    { field: 1, name: 'Crosses zero', values: [-10, -5, 0, 5, 10, 1000] },
    { field: 2, name: 'Positive whiskers', values: [1, 2, 3, 4, 5, -100] },
  ] }, original = structuredClone(source)
  const model = chart('boxplot', { yAxis: { scaleType: 'log', logBase: 10 }, options: { boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr' } } }), result = prepared(source, model), originalStatistics = structuredClone(result.boxes)
  const geometry = layout.layoutDerivedChartData(result, model), value = option(source, model)
  assert.equal(result.boxes[0].statistics.min, -10)
  assert.equal(result.boxes[0].statistics.median, 2.5)
  assert.deepEqual(result.boxes[0].statistics.outliers, [1000])
  assert.deepEqual(geometry.series[0].points, [])
  assert.equal(geometry.series[1].points[0].value, 2.5)
  assert.deepEqual([geometry.min, geometry.max], [1, 1000])
  assert.deepEqual(value.series[0].data[0].value, [])
  assert.deepEqual(value.series[0].data[1].value, [1, 1.25, 2.5, 3.75, 5])
  assert.deepEqual(value.series[1].data.map(item => item.value), [[0, 1000]])
  const projected = boxplot.chartBoxplotGeometry(model, geometry, {}, 0, 2)
  assert.equal(projected.y(0), undefined); assert.equal(projected.y(-10), undefined)
  assert.equal(projected.summary(result.boxes[0].statistics), undefined)
  assertFiniteScene(draw.drawDerivedChartData(result, model))
  assert.deepEqual(result.boxes, originalStatistics)
  assert.deepEqual(source, original)
  model.options = { allSeriesStyle: { label: { visible: true, contentType: 'series', color: 'green' } } }
  const commands = scene(source, model)
  assert.ok(!commands.some(command => command.kind === 'text' && command.text === 'Crosses zero' && command.color === 'rgba(0, 128, 0, 1)'))
  assert.ok(commands.some(command => command.kind === 'text' && command.text === 'Positive whiskers' && command.color === 'rgba(0, 128, 0, 1)'))
  assertFiniteScene(commands)
})

test('nullable style normalization detaches atomic arrays and rejects unsafe objects and sparse arrays', () => {
  const input = { border: { color: null, width: 3 }, retained: [{ value: null, nested: [1, 2] }] }
  const patch = cartesian.chartNullableValues(input)
  assert.equal(patch.border.color, undefined)
  assert.deepEqual(patch.retained, [{ value: null, nested: [1, 2] }])
  input.retained[0].nested[0] = 99
  assert.equal(patch.retained[0].nested[0], 1)
  for (const invalid of [null, [], Object.create({ value: 1 }), JSON.parse('{"__proto__":{"polluted":true}}')]) {
    assert.throws(() => cartesian.chartNullableValues(invalid), TypeError)
  }
  assert.throws(() => cartesian.chartNullableValues({ retained: Array(1) }), /稀疏/)
  assert.throws(() => cartesian.chartNullableValues({ retained: [Object.create({ value: 1 })] }), TypeError)
  assert.equal({}.polluted, undefined)
})
