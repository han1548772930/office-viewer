import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData, updateChart } from '../src/features/charts/chart-engine.ts'
import { CHART_TYPE_CATALOG } from '../src/features/charts/chart-catalog.ts'
import { CHART_GRADIENT_MAX_STOPS, chartColorChannels, chartColorRepresentative, chartLinearGradient, isChartLinearGradient, validateChartColor } from '../src/features/charts/chart-color.ts'
import { appendChartGradientStop, chartColorForMode, chartColorMode, patchChartGradientCoordinate, patchChartGradientStop, removeChartGradientStop } from '../src/features/charts/chart-color-editing.ts'
import { chartGradientPointBounds, chartSeriesGradientBounds, createChartMarkGradient, supportsChartExplicitGradient, validateChartGradientAuthoring, validateChartGradientSeriesPatch } from '../src/features/charts/chart-gradient.ts'
import { chartEChartsAutomaticGradient, chartGradientColorAt, createChartAutomaticGradient } from '../src/features/charts/chart-auto-gradient.ts'
import { chartNullableValues } from '../src/features/charts/chart-cartesian-api.ts'
import { chartPointStyle } from '../src/features/charts/chart-derived.ts'
import { chartSeriesStyle, cloneChartOptions, mergeChartOptions, validateChartOptions, validateChartPalette } from '../src/features/charts/chart-options.ts'
import { chartColorControl, chartFillPatternControls, chartSeriesControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartRect } from '../src/features/charts/chart-geometry.ts'
import { drawChartData } from '../src/features/charts/chart-draw.ts'
import { chartPlotFrame } from '../src/features/charts/chart-layout.ts'
import { chartPieLayout, chartPieOuterRadius } from '../src/features/charts/chart-pie-layout.ts'
import { chartEChartsOption, chartGraphics } from '../src/features/charts/echarts-option.ts'
import { CHART_GRADIENT_COMMAND_BUDGET } from '../src/features/charts/chart-gradient-geometry.ts'
import { waterfallComputedFixture } from './fixtures/chart-waterfall-fixtures.mjs'

// P16 implementation drafts. No tests, rendering engines or benchmarks have run.
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const input = () => ({ start: { x: -0.25, y: 0 }, end: { x: 1.25, y: 1 }, stops: [
  { offset: 1, color: 'rgba(255, 0, 0, 0.8)', opacity: 0.5 }, { offset: 0, color: 'blue', opacity: 0.75 },
  { offset: 0.5, color: 'green' }, { offset: 0.5, color: 'purple', opacity: 0.6 },
] })
const linear = () => chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 0 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] })
const area = points => Math.abs(points.reduce((sum, point, index) => { const next = points[(index + 1) % points.length]; return sum + point.x * next.y - next.x * point.y }, 0)) / 2
const paintedArea = commands => commands.filter(command => command.kind === 'polygon').reduce((sum, command) => sum + area(command.points), 0)
function colorAt(commands, x, y) {
  const command = commands.find(command => {
    if (command.kind !== 'polygon') return false
    const sides = command.points.map((point, index) => { const next = command.points[(index + 1) % command.points.length]; return (next.x - point.x) * (y - point.y) - (next.y - point.y) * (x - point.x) })
    return sides.every(value => value >= -1e-8) || sides.every(value => value <= 1e-8)
  })
  assert.ok(command, `paint exists at ${x}, ${y}`)
  return chartColorChannels(command.fill)
}
function freeze(value) {
  if (value && typeof value === 'object') { Object.values(value).forEach(freeze); Object.freeze(value) }
  return value
}
function model(type = 'column', options = {}) {
  return { id: 'explicit-gradient', sheetId: 'sheet', type, revision: 0, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } }, position: { x: 0, y: 0 },
    size: { width: 360, height: 240 }, legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
function fixture(type, options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = Object.fromEntries([['X', 'Units'], [1, 10], [2, 20], [3, 30]].flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  const authored = type === 'waterfall' ? { ...options, waterfall: { useSubtotal: true, pointRoles: [{ seriesIndex: 1, dataPointIndex: 0, role: 'change' }] } } : options
  return { workbook, chart: createChart(workbook, sheet.id, model(type, authored)) }
}

test('the constructor detaches plain authored data, retains stop order and accepts finite coordinates outside a mark', () => {
  const authored = freeze(input()), before = structuredClone(authored), first = chartLinearGradient(authored), second = chartLinearGradient(authored)
  assert.equal(isChartLinearGradient(first), true)
  assert.equal(first.type, 'linear-gradient')
  assert.deepEqual(first.stops.map(stop => stop.offset), [1, 0, 0.5, 0.5])
  first.start.x = -1
  first.stops[0].opacity = 0.2
  assert.deepEqual(authored, before)
  assert.equal(second.start.x, -0.25)
  assert.equal(second.stops[0].opacity, 0.5)
  assert.deepEqual(chartColorChannels(chartColorRepresentative(second)), [0, 0, 1, 0.75])
  assert.equal(chartColorRepresentative('red'), 'red')
})

test('malformed gradients reject at their authoring boundary, including sparse, hidden and accessor metadata', () => {
  const invalid = [null, [], { type: 'radial-gradient' }, { ...linear(), start: { x: NaN, y: 0 } },
    { ...linear(), end: { x: Infinity, y: 1 } }, { ...linear(), stops: [{ offset: 0, color: 'red' }] },
    { ...linear(), stops: Array(CHART_GRADIENT_MAX_STOPS + 1).fill({ offset: 0, color: 'red' }) },
    { ...linear(), stops: [{ offset: -0.1, color: 'red' }, { offset: 1, color: 'blue' }] },
    { ...linear(), stops: [{ offset: 0, color: 'red', opacity: 1.1 }, { offset: 1, color: 'blue' }] },
    { ...linear(), stops: [{ offset: 0, color: 'not-a-color' }, { offset: 1, color: 'blue' }] },
    { ...linear(), stops: [{ offset: 0, color: 'red' }, , { offset: 1, color: 'blue' }] },
    Object.assign(Object.create({ inherited: true }), linear()), { ...linear(), nativePaint: true },
  ]
  const hidden = linear(), accessor = linear(), arrayMetadata = linear(), symbolMetadata = linear()
  Object.defineProperty(hidden, 'nativePaint', { value: true })
  Object.defineProperty(accessor.start, 'x', { get() { throw new Error('accessor was evaluated') } })
  arrayMetadata.stops.extra = 'unsupported'
  symbolMetadata[Symbol('paint')] = true
  for (const value of [...invalid, hidden, accessor, arrayMetadata, symbolMetadata]) assert.throws(() => validateChartColor(value), error => error instanceof TypeError || error instanceof RangeError)
  assert.doesNotThrow(() => validateChartColor({ ...linear(), stops: Array.from({ length: CHART_GRADIENT_MAX_STOPS }, (_, index) => ({ offset: index / (CHART_GRADIENT_MAX_STOPS - 1), color: 'red' })) }))
})

test('type-specific public styles distinguish fill series, fill data points and solid symbols across the catalog', () => {
  const color = linear()
  for (const entry of CHART_TYPE_CATALOG) {
    const series = () => validateChartGradientSeriesPatch(entry.value, { color })
    const point = () => validateChartGradientSeriesPatch(entry.value, { dataPoints: { 0: { color } } })
    if (supportsChartExplicitGradient(entry.value)) assert.doesNotThrow(series); else assert.throws(series, TypeError)
    if (supportsChartExplicitGradient(entry.value, 'dataPoint')) assert.doesNotThrow(point); else assert.throws(point, TypeError)
    assert.equal(supportsChartExplicitGradient(entry.value, 'symbol'), false)
  }
  assert.equal(CHART_TYPE_CATALOG.filter(entry => supportsChartExplicitGradient(entry.value)).length, 31)
  assert.equal(CHART_TYPE_CATALOG.filter(entry => supportsChartExplicitGradient(entry.value, 'dataPoint')).length, 25)
  for (const style of [{ point: { color } }, { border: { color } }]) assert.throws(() => validateChartOptions({ allSeriesStyle: style }), TypeError)
  assert.throws(() => validateChartPalette([color]), TypeError)
})

test('clone, nullable updates and group inheritance replace a complete color atom without merging vectors or stops', () => {
  const color = chartLinearGradient(input()), replacement = linear(), options = freeze({ allSeriesStyle: { color, dataPoints: { 0: { color } } } })
  const copy = cloneChartOptions(options), patch = chartNullableValues({ color: replacement }), merged = mergeChartOptions(options, { allSeriesStyle: patch })
  copy.allSeriesStyle.color.start.x = 5
  patch.color.stops[0].color = 'yellow'
  assert.deepEqual(merged.allSeriesStyle.color, replacement)
  assert.equal(options.allSeriesStyle.color.start.x, -0.25)
  assert.equal(options.allSeriesStyle.dataPoints[0].color.stops[0].color, 'rgba(255, 0, 0, 0.8)')
  const reset = mergeChartOptions(options, { allSeriesStyle: chartNullableValues({ color: null }) })
  assert.equal(reset.allSeriesStyle.color, undefined)
  assert.deepEqual(mergeChartOptions(options, { allSeriesStyle: { color: 'lime' } }).allSeriesStyle.color, 'lime')
  assert.throws(() => mergeChartOptions(options, { allSeriesStyle: { color: { type: 'linear-gradient', start: { x: 0, y: 0 } } } }), TypeError)
  const inherited = chartSeriesStyle({ options: { allSeriesStyle: { color }, series: { 1: { color: replacement } }, namedSeries: [{ name: 'Units', style: { color } }] } }, 1, 0, undefined, 'Units')
  assert.deepEqual(inherited.color, color)
})

test('explicit paint wins over the automatic switch, retains authored direction and consumes stop/fill alpha once', () => {
  const color = chartLinearGradient(input()), style = freeze({ color, fillOpacity: 0.5 }), chart = model('bar-horizontal', { autoGradientFill: false }), before = structuredClone(style)
  const bounds = { x: 10, y: 20, width: 80, height: 40 }, gradient = createChartMarkGradient(chart, style, 'bar', { negative: true, reverse: true, horizontal: true }, bounds)
  assert.equal(gradient.authored, true)
  assert.deepEqual(gradient.start, color.start)
  assert.deepEqual(gradient.end, color.end)
  assert.deepEqual(gradient.stops.map(stop => stop.offset), [0, 0.5, 0.5, 1])
  near(chartColorChannels(gradient.stops[0].color)[3], 0.375)
  near(chartColorChannels(gradient.stops.at(-1).color)[3], 0.2)
  assert.deepEqual(style, before)
  assert.equal(createChartAutomaticGradient(chart, style, 'bar'), undefined)
  const enabled = createChartMarkGradient(model('bar-horizontal', { autoGradientFill: true }), style, 'bar')
  assert.deepEqual(enabled.stops, gradient.stops)
  assert.equal(createChartMarkGradient(chart, { ...style, fillPattern: { preset: 'cross', foregroundColor: 'red', backgroundColor: 'white' } }, 'bar'), undefined)
  assert.equal(createChartMarkGradient(chart, style, 'line', { symbol: true }), undefined)
  assert.equal(createChartMarkGradient(model('scatter'), style, 'line'), undefined)
  near(chartColorChannels(chartGradientColorAt(gradient, 0.5))[3], 0.3)
  assert.deepEqual(chartColorChannels(chartGradientColorAt(gradient, -10)), chartColorChannels(gradient.stops[0].color))
  assert.deepEqual(chartColorChannels(chartGradientColorAt(gradient, 10)), chartColorChannels(gradient.stops.at(-1).color))
})

test('3D shading preserves every stop and its alpha while authored vectors and shared projected bounds remain unchanged', () => {
  const color = linear(), chart = model('column3D'), bounds = { x: 5, y: 8, width: 90, height: 120 }
  const paint = createChartMarkGradient(chart, { color, fillOpacity: 0.4 }, 'column', { shade: 0.6, negative: true, reverse: true }, bounds)
  assert.deepEqual(paint.start, color.start)
  assert.deepEqual(paint.end, color.end)
  assert.deepEqual(paint.bounds, bounds)
  assert.deepEqual(chartColorChannels(paint.stops[0].color), [0.6, 0, 0, 0.4])
  assert.deepEqual(chartColorChannels(paint.stops[1].color), [0, 0, 0.6, 0.4])
  assert.deepEqual(chartGradientPointBounds([{ x: 5, y: 8 }, { x: 95, y: 128 }, { x: NaN, y: 0 }]), bounds)
  assert.equal(chartGradientPointBounds([]), undefined)
})

test('hard and narrow stops preserve mark area and their transition boundaries without overlapping alpha fills', () => {
  const color = chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 0 }, stops: [
    { offset: 0, color: 'red', opacity: 0.4 }, { offset: 0.499, color: 'red', opacity: 0.4 },
    { offset: 0.499, color: 'lime', opacity: 0.4 }, { offset: 0.5, color: 'lime', opacity: 0.4 },
    { offset: 0.5, color: 'blue', opacity: 0.4 }, { offset: 1, color: 'blue', opacity: 0.4 },
  ] }), commands = [], gradient = createChartMarkGradient(model(), { color }, 'column')
  chartRect(commands, { x: 0, y: 0, width: 100, height: 100 }, 'red', undefined, undefined, undefined, gradient)
  near(paintedArea(commands), 10_000)
  assert.deepEqual(colorAt(commands, 49.8, 23), [1, 0, 0, 0.4])
  assert.deepEqual(colorAt(commands, 49.95, 23), [0, 1, 0, 0.4])
  assert.deepEqual(colorAt(commands, 50.1, 23), [0, 0, 1, 0.4])
  assert.deepEqual(chartColorChannels(chartGradientColorAt(gradient, 0.499)), [0, 1, 0, 0.4])
  assert.deepEqual(chartColorChannels(chartGradientColorAt(gradient, 0.5)), [0, 0, 1, 0.4])
})

test('degenerate vectors with authored intermediate stops paint one midpoint layer and budget rejection is atomic', () => {
  const color = chartLinearGradient({ start: { x: 0.5, y: 0.5 }, end: { x: 0.5, y: 0.5 }, stops: [
    { offset: 0, color: 'red' }, { offset: 0.2, color: 'lime' }, { offset: 0.5, color: 'purple', opacity: 0.4 }, { offset: 1, color: 'blue' },
  ] }), gradient = createChartMarkGradient(model(), { color }, 'column'), commands = []
  chartRect(commands, { x: 0, y: 0, width: 100, height: 100 }, 'red', undefined, undefined, undefined, gradient)
  near(paintedArea(commands), 10_000)
  assert.ok(commands.every(command => command.fill === chartGradientColorAt(gradient, 0.5)))
  const sentinel = { kind: 'rect', x: 0, y: 0, width: 1, height: 1, fill: 'black' }, full = Array(CHART_GRADIENT_COMMAND_BUDGET - 1).fill(sentinel)
  assert.throws(() => chartRect(full, { x: 0, y: 0, width: 20, height: 20 }, 'red', undefined, undefined, undefined, gradient), /预算/)
  assert.equal(full.length, CHART_GRADIENT_COMMAND_BUDGET - 1)
  assert.equal(full.at(-1), sentinel)
})

test('all fill-capable chart families actually paint explicit gradients with automatic fill disabled and preserve data', () => {
  for (const entry of CHART_TYPE_CATALOG.filter(entry => supportsChartExplicitGradient(entry.value))) {
    const { workbook, chart } = fixture(entry.value, { autoGradientFill: false, allSeriesStyle: { color: linear(), fillOpacity: 0.7, point: { shape: 'none' }, border: { width: 2 } } })
    const before = structuredClone(chart), data = readChartData(workbook, chart), solid = structuredClone(chart)
    updateChart(workbook, solid, { options: { allSeriesStyle: { color: 'red' } } })
    const computed = entry.family === 'waterfall' ? waterfallComputedFixture(data, 'catalog-changes') : undefined
    assert.notDeepEqual(drawChartData(data, chart, {}, computed), drawChartData(data, solid, {}, computed), entry.id)
    assert.notDeepEqual(chartEChartsOption(workbook, chart, {}, computed), chartEChartsOption(workbook, solid, {}, computed), entry.id)
    assert.deepEqual(readChartData(workbook, solid), data)
    assert.deepEqual(chart, before)
  }
})

test('native positive/negative/reversed bars use the same authored vector and native item opacity is already consumed', () => {
  const color = linear(), { workbook, chart } = fixture('bar-horizontal', { autoGradientFill: false, allSeriesStyle: { color, fillOpacity: 0.4 } })
  workbook.sheets[0].cells['2:1'] = createCell(2, 1, -20)
  const first = chartEChartsOption(workbook, chart).series[0].data
  const reversed = structuredClone(chart)
  updateChart(workbook, reversed, { yAxis: { scaleType: 'linear', reverse: true } })
  const second = chartEChartsOption(workbook, reversed).series[0].data
  for (const datum of [...first, ...second]) {
    assert.deepEqual([datum.itemStyle.color.x, datum.itemStyle.color.y, datum.itemStyle.color.x2, datum.itemStyle.color.y2], [0, 0, 1, 0])
    assert.equal(datum.itemStyle.color.global, false)
    assert.equal(datum.itemStyle.opacity, 1)
    near(chartColorChannels(datum.itemStyle.color.colorStops[0].color)[3], 0.4)
  }
})

test('native pie gradients resolve actual sector circles with automatic fill off and preserve holes and exploded offsets', () => {
  const color = linear(), { workbook, chart } = fixture('doughnut', { autoGradientFill: false, allSeriesStyle: { color }, pie: { explosion: 10, doughnutHole: 0.6, rosePie: true } })
  const option = chartEChartsOption(workbook, chart), data = readChartData(workbook, chart), { plot } = chartPlotFrame(chart), sectors = chartPieLayout(data, chart, plot).sectors
  for (const sector of sectors) {
    const paint = option.series[0].data[sector.part.index].itemStyle.color
    assert.equal(paint.global, true)
    near(paint.x, sector.center.x - sector.outer)
    near(paint.y, sector.center.y - sector.outer)
    near(paint.x2, sector.center.x + sector.outer)
    near(paint.y2, sector.center.y - sector.outer)
  }
  const radius = chartPieOuterRadius(plot, chart.options.pie.explosion)
  assert.deepEqual(option.series[0].radius, [radius * 0.6, radius])
})

test('explicit line/area and combination fragments use the common Canvas scene and source-wide bounds including smooth extrema', () => {
  for (const type of ['line', 'area', 'combination']) {
    const { workbook, chart } = fixture(type, { autoGradientFill: false, allSeriesStyle: { color: linear(), point: { shape: 'none' } },
      ...(type === 'combination' ? { series: { 1: { type: 'line' } } } : {}), area: { smooth: true } })
    const data = readChartData(workbook, chart), option = chartEChartsOption(workbook, chart)
    assert.ok(option.graphic?.elements?.length > 0)
    assert.deepEqual(option.graphic.elements, chartGraphics(drawChartData(data, chart), chart.id))
    assert.equal(option.series, undefined)
  }
  assert.deepEqual(chartSeriesGradientBounds({ points: [{ x: 10, y: 30, baseY: 50 }, { x: 40, y: 20, baseY: 55 }] }, true, [{ x: 20, y: 15 }]), { x: 10, y: 15, width: 30, height: 40 })
})

test('type switches retain inactive data-point gradients but solid-only symbols resume the active series color', () => {
  const color = linear(), { workbook, chart } = fixture('column', { autoGradientFill: false, series: { 1: { color: 'red', dataPoints: { 0: { color } } } } })
  updateChart(workbook, chart, { type: 'line' })
  assert.deepEqual(chart.options.series[1].dataPoints[0].color, color)
  const style = chartPointStyle(chartSeriesStyle(chart, 1, 0, undefined, 'Units'), 0, chart.type)
  assert.equal(style.color, 'red')
  assert.equal(style.point.color, undefined)
  const solid = structuredClone(chart)
  solid.options.series[1].dataPoints[0] = {}
  assert.deepEqual(drawChartData(readChartData(workbook, chart), chart), drawChartData(readChartData(workbook, solid), solid))
  assert.doesNotThrow(() => validateChartGradientAuthoring('line', cloneChartOptions(chart.options), chart.options))
  const before = structuredClone(chart), invalid = cloneChartOptions(chart.options)
  invalid.series[1].dataPoints[0].color.stops[0].color = 'purple'
  Object.defineProperty(workbook.sheets[0], 'cells', { get() { throw new Error('records were scanned before rejection') } })
  assert.throws(() => updateChart(workbook, chart, { options: invalid }), TypeError)
  assert.deepEqual(chart, before)
  assert.throws(() => createChart(workbook, workbook.sheets[0].id, model('scatter', { allSeriesStyle: { color } })), TypeError)
})

test('the color editor preserves alpha, vector/stop order and inheritance while invalid edits leave its source untouched', () => {
  const color = freeze(chartLinearGradient(input())), before = structuredClone(color)
  assert.equal(chartColorMode(undefined), 'inherit')
  assert.equal(chartColorMode('red'), 'solid')
  assert.equal(chartColorMode(color), 'gradient')
  assert.equal(chartColorForMode(color, 'inherit'), undefined)
  assert.deepEqual(chartColorChannels(chartColorForMode(color, 'solid')), [0, 0, 1, 0.75])
  assert.deepEqual(chartColorForMode(undefined, 'gradient', color), color)
  const seeded = chartColorForMode(undefined, 'gradient', 'rgba(255, 0, 0, 0.5)')
  assert.equal(seeded.stops[0].color, 'rgba(255, 0, 0, 0.5)')
  assert.equal(seeded.stops[1].opacity, 0)
  const changed = patchChartGradientCoordinate(color, 'end', 'x', 2), stop = patchChartGradientStop(changed, 1, { opacity: undefined })
  assert.equal(stop.end.x, 2)
  assert.deepEqual(stop.stops.map(entry => entry.offset), [1, 0, 0.5, 0.5])
  assert.equal(stop.stops[1].opacity, undefined)
  const appended = appendChartGradientStop(stop)
  assert.deepEqual(appended.stops.at(-1), stop.stops.at(-1))
  assert.deepEqual(removeChartGradientStop(appended, appended.stops.length - 1), stop)
  for (const edit of [() => patchChartGradientCoordinate(color, 'start', 'y', NaN), () => patchChartGradientStop(color, 0, { offset: 2 }),
    () => patchChartGradientStop(color, 10, { opacity: 0 }), () => patchChartGradientStop(color, 0, { opacity: null }), () => removeChartGradientStop(linear(), 0)]) assert.throws(edit)
  assert.deepEqual(color, before)
})

test('settings controls distinguish authored color from inherited paint and keep texture colors solid and detached', () => {
  const config = model('column', { allSeriesStyle: { color: linear() }, namedSeries: [{ name: 'Units', style: {} }] })
  const controls = chartSeriesControls({ name: 'Units' }, false, false, false, { field: 1, position: 0, name: 'Units' }), color = controls.find(control => control.kind === 'chart-color')
  assert.equal(color.read(config), undefined)
  assert.deepEqual(color.fallback(config), linear())
  const authored = linear()
  color.write(config, authored)
  authored.stops[0].color = 'yellow'
  assert.equal(config.options.namedSeries[0].style.color.stops[0].color, 'red')
  color.write(config, undefined)
  assert.equal(color.read(config), undefined)
  const pure = model('line'), point = chartColorControl('data.color', '项目颜色', () => undefined, (chart, value) => { chart.options.testColor = value }, 'dataPoint')
  assert.throws(() => point.write(pure, linear()), TypeError)
  assert.equal(Object.hasOwn(pure.options, 'testColor'), false)
  const read = chart => chart.options.allSeriesStyle, write = chart => chart.options.allSeriesStyle ??= {}
  const pattern = chartFillPatternControls('style', read, write)
  pattern[0].write(config, 'cross')
  assert.equal(typeof config.options.allSeriesStyle.fillPattern.foregroundColor, 'string')
  assert.deepEqual(chartColorChannels(config.options.allSeriesStyle.fillPattern.foregroundColor), [1, 0, 0, 1])
  assert.equal(chartEChartsAutomaticGradient(createChartMarkGradient(config, { color: linear() }, 'column')).type, 'linear')
})
