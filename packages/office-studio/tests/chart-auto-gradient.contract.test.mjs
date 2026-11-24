import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData, updateChart } from '../src/features/charts/chart-engine.ts'
import { CHART_TYPE_CATALOG } from '../src/features/charts/chart-catalog.ts'
import { chartAutoGradientFill, chartEChartsAutomaticGradient, chartGradientColorAt, createChartAutomaticGradient, supportsChartAutoGradientControl } from '../src/features/charts/chart-auto-gradient.ts'
import { chartAutoGradientControls } from '../src/features/charts/chart-plot-area-controls.ts'
import { chartColorChannels } from '../src/features/charts/chart-color.ts'
import { chartPointStyle } from '../src/features/charts/chart-derived.ts'
import { chartSeriesStyle, cloneChartOptions } from '../src/features/charts/chart-options.ts'
import { drawChartData, drawChartSector } from '../src/features/charts/chart-draw.ts'
import { chartLine, chartPolygon, chartRect } from '../src/features/charts/chart-geometry.ts'
import { CHART_GRADIENT_COMMAND_BUDGET } from '../src/features/charts/chart-gradient-geometry.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { chartEChartsItemStyle, chartEChartsLineStyle } from '../src/features/charts/echarts-style.ts'
import { waterfallComputedFixture } from './fixtures/chart-waterfall-fixtures.mjs'

// P15 drafts for the unified validation phase. No test in this file has run.
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const triangleArea = points => Math.abs(points.reduce((sum, point, index) => { const next = points[(index + 1) % points.length]; return sum + point.x * next.y - next.x * point.y }, 0)) / 2
const paintedArea = commands => commands.filter(command => command.kind === 'polygon').reduce((sum, command) => sum + triangleArea(command.points), 0)
function colorAt(commands, x, y) {
  const command = commands.find(command => {
    if (command.kind !== 'polygon') return false
    const signs = command.points.map((point, index) => { const next = command.points[(index + 1) % command.points.length]; return (next.x - point.x) * (y - point.y) - (next.y - point.y) * (x - point.x) })
    return signs.every(value => value >= -1e-8) || signs.every(value => value <= 1e-8)
  })
  assert.ok(command, `paint exists at ${x}, ${y}`)
  return chartColorChannels(command.fill)
}
function model(type = 'column', options = {}) {
  return { id: 'gradient', sheetId: 'sheet', type, revision: 0, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } }, position: { x: 0, y: 0 },
    size: { width: 360, height: 240 }, legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
function fixture(type, options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = Object.fromEntries([['X', 'Units'], [1, 10], [2, 20], [3, 30]].flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  const authored = type === 'waterfall' ? { ...options, waterfall: { useSubtotal: true, pointRoles: [{ seriesIndex: 1, dataPointIndex: 0, role: 'change' }] } } : options
  return { workbook, chart: createChart(workbook, sheet.id, model(type, authored)) }
}

test('area-type defaults, explicit off and UI reset preserve author intent across all directory variants', () => {
  for (const entry of CHART_TYPE_CATALOG) {
    const chart = model(entry.value)
    assert.equal(chartAutoGradientFill(chart), entry.family === 'area', entry.id)
    chart.options.autoGradientFill = false
    assert.equal(chartAutoGradientFill(chart), false)
    const control = chartAutoGradientControls()[0]
    assert.equal(control.visible(chart), ['column', 'bar', 'area', 'combination', 'pie', 'waterfall'].includes(entry.family))
    control.reset(chart)
    assert.equal(chart.options.autoGradientFill, undefined)
    assert.equal(chartAutoGradientFill(chart), entry.family === 'area')
  }
  assert.equal(CHART_TYPE_CATALOG.filter(entry => supportsChartAutoGradientControl(entry.value)).length, 25)
  assert.equal(supportsChartAutoGradientControl('line'), false)
  assert.ok(createChartAutomaticGradient(model('line', { autoGradientFill: true }), { color: 'red' }, 'line'))
  assert.equal(createChartAutomaticGradient(model('histogram', { autoGradientFill: true }), { color: 'red' }, 'column'), undefined)
})

test('effective source/point colors feed the preset and real patterns take precedence without mutating inputs', () => {
  const options = { autoGradientFill: true, allSeriesStyle: { color: 'red', fillOpacity: 0.5 },
    series: { 1: { color: 'blue' } }, namedSeries: [{ name: 'Units', style: { color: 'green', fillPattern: null, dataPoints: { 0: { color: 'purple' } } } }] }
  const chart = model('column', options), original = structuredClone(chart), style = chartPointStyle(chartSeriesStyle(chart, 1, 0, undefined, 'Units'), 0)
  const bounds = { x: 5, y: 7, width: 20, height: 40 }, gradient = createChartAutomaticGradient(chart, style, 'column', {}, bounds)
  assert.ok(gradient)
  assert.equal(style.color, 'purple')
  const native = chartEChartsAutomaticGradient(gradient)
  native.colorStops[0].color = 'black'
  bounds.x = 500
  assert.equal(gradient.bounds.x, 5)
  assert.notEqual(gradient.stops[0].color, 'black')
  assert.equal(createChartAutomaticGradient(chart, { ...style, fillPattern: { preset: 'horz', foregroundColor: 'red', backgroundColor: 'yellow' } }, 'column'), undefined)
  assert.deepEqual(chart, original)
  for (const value of [0, 1, 'true', null, {}]) assert.throws(() => cloneChartOptions({ autoGradientFill: value }), TypeError)
})

test('negative and reversed vertical/horizontal bars always run from their rendered base toward their tip', () => {
  const chart = model('column', { autoGradientFill: true }), rectangle = { x: 0, y: 0, width: 100, height: 100 }
  for (const horizontal of [false, true]) for (const negative of [false, true]) for (const reverse of [false, true]) {
    const gradient = createChartAutomaticGradient(chart, { color: 'red' }, horizontal ? 'bar' : 'column', { negative, reverse })
    const commands = []
    chartRect(commands, rectangle, 'red', undefined, undefined, undefined, gradient)
    const start = colorAt(commands, horizontal ? 2 : 50, horizontal ? 50 : 98), end = colorAt(commands, horizontal ? 98 : 50, horizontal ? 50 : 2)
    const brightness = channels => channels[0] + channels[1] + channels[2]
    assert.equal(brightness(start) > brightness(end), negative === reverse)
    near(paintedArea(commands), 10_000)
  }
})

test('fill/line alpha is embedded once and global native paint shares Canvas coordinates', () => {
  const chart = model('area', { autoGradientFill: true }), style = { color: 'rgba(100, 150, 200, 0.5)', fillOpacity: 0.4 }
  const gradient = createChartAutomaticGradient(chart, style, 'area', {}, { x: 10, y: 20, width: 100, height: 80 })
  near(chartColorChannels(gradient.stops[0].color)[3], 0.012)
  near(chartColorChannels(gradient.stops[1].color)[3], 0.2)
  const native = chartEChartsItemStyle(style, gradient)
  assert.equal(native.opacity, 1)
  assert.deepEqual([native.color.x, native.color.y, native.color.x2, native.color.y2, native.color.global], [10, 100, 10, 20, true])
  const line = createChartAutomaticGradient(model('line', { autoGradientFill: true }), { color: style.color, fillOpacity: 0.3 }, 'line')
  assert.equal(chartEChartsLineStyle({ color: style.color, border: { opacity: 0.3 } }, line).opacity, 1)
  near(chartColorChannels(line.stops[0].color)[3], 0.15)
})

test('area fragments conserve geometry and sample one shared gradient, including colors outside the reference frame', () => {
  const chart = model('area'), bounds = { x: 0, y: 0, width: 100, height: 100 }
  const gradient = createChartAutomaticGradient(chart, { color: 'blue', fillOpacity: 0.6 }, 'area', {}, bounds), commands = []
  chartRect(commands, { x: 0, y: 0, width: 50, height: 100 }, 'blue', undefined, undefined, undefined, gradient)
  chartRect(commands, { x: 50, y: 0, width: 50, height: 100 }, 'blue', undefined, undefined, undefined, gradient)
  near(paintedArea(commands), 10_000)
  assert.deepEqual(colorAt(commands, 25, 72), colorAt(commands, 75, 72))
  const outside = []
  chartRect(outside, { x: 0, y: -30, width: 10, height: 10 }, 'blue', undefined, undefined, undefined, gradient)
  assert.ok(outside.every(command => command.fill === chartGradientColorAt(gradient, 1)))
  near(paintedArea(outside), 100)
})

test('dashed line fragments and annular sector triangles use shared coordinates without filling the inner hole', () => {
  const gradient = createChartAutomaticGradient(model('line', { autoGradientFill: true }), { color: 'red' }, 'line', {}, { x: 0, y: 0, width: 100, height: 100 }), lines = []
  chartLine(lines, { x: 10, y: 10 }, { x: 90, y: 90 }, { width: 3, dashType: 'dashed' }, 'red', undefined, gradient)
  assert.ok(lines.length > 2 && new Set(lines.map(line => line.color)).size > 2)
  assert.ok(lines.every(line => line.width === 3 && line.x1 <= line.x2 && line.y1 <= line.y2))
  const sectorGradient = createChartAutomaticGradient(model('pie', { autoGradientFill: true }), { color: 'blue' }, 'pie'), solid = [], painted = [], center = { x: 40, y: 40 }
  drawChartSector(solid, center, 12, 24, 0, Math.PI, 'blue')
  drawChartSector(painted, center, 12, 24, 0, Math.PI, 'blue', undefined, undefined, undefined, sectorGradient)
  near(paintedArea(painted), paintedArea(solid))
  assert.ok(new Set(painted.map(command => command.fill)).size > 2)
  assert.ok(painted.every(command => command.points.every(point => Math.hypot(point.x - center.x, point.y - center.y) >= 12 * Math.cos(Math.PI / 96) - 1e-7)))
})

test('degenerate paint uses the midpoint and a geometry budget failure commits no partial append', () => {
  const gradient = { start: { x: 0.5, y: 0.5 }, end: { x: 0.5, y: 0.5 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }, commands = []
  chartPolygon(commands, [{ x: 0, y: 0 }, { x: 10, y: 0 }, { x: 0, y: 10 }], 'red', undefined, undefined, undefined, gradient)
  assert.ok(commands.every(command => command.fill === chartGradientColorAt(gradient, 0.5)))
  const sentinel = { kind: 'rect', x: 0, y: 0, width: 1, height: 1, fill: 'black' }, full = Array(CHART_GRADIENT_COMMAND_BUDGET).fill(sentinel)
  assert.throws(() => chartRect(full, { x: 0, y: 0, width: 10, height: 10 }, 'red', undefined, undefined, undefined, gradient), /预算/)
  assert.equal(full.length, CHART_GRADIENT_COMMAND_BUDGET)
  assert.equal(full.at(-1), sentinel)
  const almostFull = Array(CHART_GRADIENT_COMMAND_BUDGET - 1).fill(sentinel)
  assert.throws(() => chartRect(almostFull, { x: 0, y: 0, width: 10, height: 10 }, 'red', undefined, undefined, undefined, gradient), /预算/)
  assert.equal(almostFull.length, CHART_GRADIENT_COMMAND_BUDGET - 1)
  assert.equal(almostFull.at(-1), sentinel)
})

test('every supported 2D/3D family actually changes paint while retaining bound data and source authoring', () => {
  for (const entry of CHART_TYPE_CATALOG.filter(entry => ['column', 'bar', 'line', 'area', 'combination', 'pie', 'waterfall'].includes(entry.family))) {
    const { workbook, chart } = fixture(entry.value, { autoGradientFill: true }), before = structuredClone(chart), data = readChartData(workbook, chart)
    const off = structuredClone(chart)
    updateChart(workbook, off, { options: { autoGradientFill: false } })
    const computed = entry.family === 'waterfall' ? waterfallComputedFixture(data, 'catalog-changes') : undefined
    assert.notDeepEqual(drawChartData(data, chart, {}, computed), drawChartData(data, off, {}, computed), entry.id)
    assert.notDeepEqual(chartEChartsOption(workbook, chart, {}, computed), chartEChartsOption(workbook, off, {}, computed), entry.id)
    assert.deepEqual(readChartData(workbook, off), data)
    assert.deepEqual(chart, before)
  }
})

test('native sign overrides and per-segment authored paints follow the same effective options', () => {
  const { workbook, chart } = fixture('bar-horizontal', { autoGradientFill: true })
  const sheet = workbook.sheets[0]
  sheet.cells['2:1'] = createCell(2, 1, -20)
  const option = chartEChartsOption(workbook, chart), values = option.series[0].data
  assert.equal(values[0].itemStyle.color.x, 0)
  assert.equal(values[1].itemStyle.color.x, 1)
  assert.equal(values[0].itemStyle.opacity, 1)
  const before = structuredClone(chart), area = structuredClone(chart)
  updateChart(workbook, area, { type: 'area', options: { allSeriesStyle: { dataPoints: { 1: { color: 'purple', fillOpacity: 0.7 } } } } })
  const authored = chartEChartsOption(workbook, area)
  assert.ok(authored.graphic?.elements?.length > 0)
  assert.equal(authored.series, undefined)
  assert.deepEqual(chart, before)
})
