import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartLinearGradient, normalizeChartColor } from '../src/features/charts/chart-color.ts'
import { chartPointStyle, deriveChartData } from '../src/features/charts/chart-derived.ts'
import { drawChartData } from '../src/features/charts/chart-draw.ts'
import { chartManualPlotFrame, chartPlotAreaFromFrame } from '../src/features/charts/chart-plot-area.ts'
import { layoutDerivedChartData, chartPlotFrame } from '../src/features/charts/chart-layout.ts'
import { describeChartConfiguration } from '../src/features/charts/chart-description.ts'
import { cloneChartOptions, mergeChartOptions, validateChartOptions } from '../src/features/charts/chart-options.ts'
import { clearChartSeriesStyle, patchChartSeriesOverride, ChartSeriesClearTargets } from '../src/features/charts/chart-cartesian-api.ts'
import { assertChartWaterfallType, chartWaterfallAuthoredStyles, chartWaterfallConnector, chartWaterfallConnectorValues,
  chartWaterfallRole, chartWaterfallRoleStyle, chartWaterfallStyleValues, chartWaterfallUsesSecondaryAxis, patchChartWaterfallStyle } from '../src/features/charts/chart-waterfall.ts'
import { chartWaterfallConnectorControls, chartWaterfallPointControls, chartWaterfallStyleControls } from '../src/features/charts/chart-waterfall-settings-controls.ts'
import { validateChartGradientAuthoring } from '../src/features/charts/chart-gradient.ts'
import { waterfallComputedFixture } from './fixtures/chart-waterfall-fixtures.mjs'

// Authoring/rendering drafts only. No runner, WASM, browser or benchmark has run.
const palette = ['#112233', '#223344', '#334455', '#445566']
const pattern = { preset: 'cross', foregroundColor: '#334455', backgroundColor: '#eeeeee' }
const source = { categories: ['A', 'B', 'C'], fields: ['Stage', 'Delta'],
  series: [{ field: 1, name: 'Delta', sourceName: 'Delta', values: [6, -2, 4] }] }
function model(options = {}) {
  return { id: 'waterfall-style', sheetId: 'sheet', revision: 0, type: 'waterfall', categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } },
    position: { x: 0, y: 0 }, size: { width: 600, height: 400 }, palette,
    legend: { visible: true, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
const totalRole = { seriesIndex: 1, dataPointIndex: 2, role: 'total' }
function marks(commands) { return commands.filter(command => command.kind === 'rect' && command.clip && command.width > 10 && command.height > 0) }
function gradient() { return chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 1 }, stops: [{ offset: 0, color: 'red', opacity: 0.4 }, { offset: 1, color: 'blue' }] }) }
function computed(name = 'style-changes') { return waterfallComputedFixture(source, name) }
function render(chart, name = 'style-changes') { return drawChartData(source, chart, {}, computed(name)) }
function geometry(chart, name = 'style-changes') { return layoutDerivedChartData(deriveChartData(source, chart, computed(name)), chart, palette) }
function pointPolygons(commands, point, layout) {
  const halfBand = layout.plot.width / layout.categories.length / 2
  return commands.filter(command => command.kind === 'polygon' && command.clip
    && command.points.every(vertex => vertex.x >= point.x - halfBand - 1e-8 && vertex.x <= point.x + halfBand + 1e-8
      && vertex.y >= Math.min(point.y, point.baseY) - 1e-8 && vertex.y <= Math.max(point.y, point.baseY) + 1e-8))
}
function area(points) { return Math.abs(points.reduce((sum, point, index) => { const next = points[(index + 1) % points.length]; return sum + point.x * next.y - next.x * point.y }, 0)) / 2 }
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)

test('Waterfall style targets classify negative/positive changes separately from a signed total', () => {
  assert.equal(chartWaterfallRole('change', 0), 'positive')
  assert.equal(chartWaterfallRole('change', -4), 'negative')
  assert.equal(chartWaterfallRole('total', -4), 'subtotal')
  assert.throws(() => assertChartWaterfallType('line'), /瀑布图/)
  assert.equal(assertChartWaterfallType('waterfall'), undefined)
})

test('fixed connector omission/false hides it and true applies the declared defaults', () => {
  assert.equal(chartWaterfallConnector(model()), undefined)
  assert.equal(chartWaterfallConnector(model({ waterfall: { connector: false } })), undefined)
  assert.deepEqual(chartWaterfallConnector(model({ waterfall: { connector: true } })), { color: '#8C8C8C', width: 1, opacity: 1, dashType: 'dashed' })
  assert.deepEqual(chartWaterfallConnector(model({ waterfall: { connector: { width: 0, opacity: 0, color: undefined } } })), { color: '#8C8C8C', width: 0, opacity: 0, dashType: 'dashed' })
  assert.equal(chartWaterfallConnector(model({ waterfall: { connector: true } }), '#999999').color, '#999999')
})

test('connector candidates detach, permit nullable resets and reject invalid scalar fields', () => {
  const candidate = { color: 'red', width: 3, dashType: 'dotted' }
  const value = chartWaterfallConnectorValues(candidate)
  value.color = 'blue'
  assert.equal(candidate.color, 'red')
  assert.deepEqual(chartWaterfallConnectorValues({ width: null, opacity: 0 }), { width: undefined, opacity: 0 })
  assert.throws(() => chartWaterfallConnectorValues({ width: -1 }), /线宽/)
  assert.throws(() => chartWaterfallConnectorValues({ opacity: 2 }), /不透明度/)
  assert.throws(() => chartWaterfallConnectorValues('on'), /普通对象/)
})

test('role patches preserve source-field and exact-name scopes and other authored roles', () => {
  const current = { series: { 1: { name: 'Business', color: 'red', waterfallStyles: { positive: { border: { width: 2 }, fillOpacity: 0.4 }, negative: { color: 'blue' } } }, 3: { color: 'green' } },
    namedSeries: [{ name: 'Delta', style: { waterfallStyles: { subtotal: { color: 'orange' } } } }] }
  const before = structuredClone(current)
  const patched = { ...current, ...patchChartWaterfallStyle(current, 1, 'positive', { color: 'purple', border: { color: 'black' } }) }
  assert.deepEqual(current, before)
  assert.equal(patched.series[1].waterfallStyles.positive.border.width, 2)
  assert.equal(patched.series[1].waterfallStyles.positive.border.color, 'black')
  assert.equal(patched.series[1].waterfallStyles.negative.color, 'blue')
  assert.equal(patched.series[3].color, 'green')
  const named = { ...patched, ...patchChartWaterfallStyle(patched, { name: 'Delta' }, 'positive', { color: 'cyan' }) }
  assert.equal(named.namedSeries[0].style.waterfallStyles.subtotal.color, 'orange')
  assert.equal(named.namedSeries[0].style.waterfallStyles.positive.color, 'cyan')
  assert.equal(named.series[1].waterfallStyles.positive.color, 'purple')
})

test('clearing a role and clearing one nullable field resume inheritance without clearing unrelated authorship', () => {
  const current = { allSeriesStyle: { waterfallStyles: { positive: { color: 'green', border: { width: 1 } } } },
    series: { 1: { waterfallStyles: { positive: { color: 'blue', border: { width: 4 } }, negative: { color: 'red' } } } } }
  const cleared = { ...current, ...patchChartWaterfallStyle(current, 1, 'positive', { color: null }) }
  assert.equal(chartWaterfallRoleStyle(model(cleared), 1, 0, 'positive', palette, 'Delta').color, 'green')
  assert.equal(chartWaterfallRoleStyle(model(cleared), 1, 0, 'positive', palette, 'Delta').border.width, 4)
  const reset = { ...current, ...patchChartWaterfallStyle(current, 1, 'positive', undefined) }
  assert.equal(chartWaterfallRoleStyle(model(reset), 1, 0, 'positive', palette, 'Delta').border.width, 1)
  assert.equal(reset.series[1].waterfallStyles.negative.color, 'red')
  const allReset = { ...current, ...patchChartSeriesOverride(current, 1, clearChartSeriesStyle(current.series[1], ChartSeriesClearTargets.WaterfallStyles)) }
  assert.equal(chartWaterfallRoleStyle(model(allReset), 1, 0, 'negative', palette, 'Delta').color, palette[1])
})

test('role merge order is base, global role, all-series role, field role, exact-name role and point', () => {
  const chart = model({ allSeriesStyle: { color: 'black', border: { width: 1 }, waterfallStyles: { positive: { color: 'red', fillOpacity: 0.8 } } },
    waterfall: { positiveColor: 'brown', styles: { positive: { color: 'green', border: { color: 'gray' } } } },
    series: { 1: { waterfallStyles: { positive: { border: { width: 3 }, dataPoints: { 0: { color: 'cyan' } } } } } },
    namedSeries: [{ name: 'Delta', style: { waterfallStyles: { positive: { fillOpacity: 0.2, dataPoints: { 0: { border: { color: 'blue' } } } } } } }] })
  const role = chartWaterfallRoleStyle(chart, 1, 0, 'positive', palette, 'Delta')
  assert.equal(role.color, 'red')
  assert.equal(role.fillOpacity, 0.2)
  assert.deepEqual(role.border, { width: 3, color: 'gray' })
  const point = chartPointStyle(role, 0, 'waterfall')
  assert.equal(point.color, 'cyan')
  assert.equal(point.border.width, 3)
  assert.equal(point.border.color, 'blue')
  assert.equal(chartWaterfallRoleStyle(chart, 1, 0, 'negative', palette, 'Delta').color, 'black')
})

test('explicit patterns and role-point solid opt-out retain independent paint contracts', () => {
  const chart = model({ allSeriesStyle: { waterfallStyles: { positive: { color: gradient(), fillPattern: pattern } } },
    series: { 1: { waterfallStyles: { positive: { dataPoints: { 0: { color: 'cyan', fillPattern: null } } } } } } })
  const role = chartWaterfallRoleStyle(chart, 1, 0, 'positive', palette, 'Delta')
  assert.equal(role.fillPattern.preset, 'cross')
  const point = chartPointStyle(role, 0, 'waterfall')
  assert.equal(point.fillPattern, null)
  assert.equal(point.color, 'cyan')
  const commands = render(chart), layout = geometry(chart)
  assert.equal(marks(commands)[0].fill, normalizeChartColor('cyan'))
  assert.equal(pointPolygons(commands, layout.waterfall.points[0], layout).length, 0)
  const patterned = pointPolygons(commands, layout.waterfall.points[2], layout)
  assert.ok(patterned.some(command => command.fill === normalizeChartColor(pattern.foregroundColor)))
  assert.ok(patterned.some(command => command.fill === normalizeChartColor(pattern.backgroundColor)))
})

test('positive/negative/subtotal colors, role labels and names alter the real Canvas scene', () => {
  const chart = model({ waterfall: { useSubtotal: false, pointRoles: [totalRole], styles: {
    positive: { color: '#010203', label: { visible: true, contentType: 'series' }, name: 'Revenue' },
    negative: { color: '#040506', border: { visible: true, color: '#ffffff', width: 3 }, label: { visible: true, contentType: 'series' }, name: 'Costs' },
    subtotal: { color: '#070809', label: { visible: true, contentType: 'series' }, name: 'Profit' } } } })
  const commands = render(chart, 'style-total-no-subtotal'), bars = marks(commands)
  assert.deepEqual(bars.map(mark => mark.fill), ['#010203', '#040506', '#070809'].map(normalizeChartColor))
  assert.equal(bars[1].stroke, normalizeChartColor('#ffffff'))
  assert.equal(bars[1].strokeWidth, 3)
  const labels = commands.filter(command => command.kind === 'text').map(command => command.text)
  for (const label of ['Revenue', 'Costs', 'Profit']) assert.equal(labels.filter(value => value === label).length, 2)
})

test('new explicit role gradients are painted and unsupported type authoring remains atomic', () => {
  const author = gradient(), chart = model({ series: { 1: { waterfallStyles: { positive: { color: author } } } } })
  const commands = render(chart), layout = geometry(chart), first = layout.waterfall.points[0]
  const painted = pointPolygons(commands, first, layout)
  assert.ok(painted.length > 2 && new Set(painted.map(command => command.fill)).size > 2)
  const width = layout.plot.width / layout.categories.length / 2.5 - 1
  near(painted.reduce((sum, command) => sum + area(command.points), 0), width * Math.abs(first.y - first.baseY))
  assert.throws(() => validateChartGradientAuthoring('radar', chart.options), /纯色/)
  assert.equal(validateChartGradientAuthoring('radar', chart.options, structuredClone(chart.options)), undefined)
  assert.deepEqual(chart.options.series[1].waterfallStyles.positive.color, author)
})

test('role axis assignment changes point coordinates, axis extent and automatic plot margins together', () => {
  const chart = model({ rightYAxis: { scaleType: 'linear', min: -10, max: 10 },
    waterfall: { pointRoles: [totalRole], styles: { negative: { rightAxis: true } } } })
  const derived = deriveChartData(source, chart, computed('style-total')), layout = layoutDerivedChartData(derived, chart, palette)
  assert.equal(layout.series[0].points[0].rightAxis, false)
  assert.equal(layout.series[0].points[1].rightAxis, true)
  assert.equal(layout.rightMin, -10)
  assert.equal(layout.rightMax, 10)
  near(layout.series[0].points[1].y, layout.plot.y + 0.3 * layout.plot.height)
  assert.equal(chartWaterfallUsesSecondaryAxis(chart, source.series), true)
  const ordinary = chartPlotFrame(model(), source.series)
  assert.ok(chartPlotFrame(chart, source.series).plot.width < ordinary.plot.width)
  const author = chartPlotAreaFromFrame(layout.frame, layout.plot)
  assert.deepEqual(chartManualPlotFrame('waterfall', layout.frame, author), layout.plot)
})

test('omitted connectors produce no lines and authored width/opacity/dashes survive scene generation', () => {
  const without = render(model())
  const connected = render(model({ waterfall: { connector: { color: '#fedcba', width: 3, opacity: 0.5, dashType: 'dotted' } } }))
  const links = connected.filter(command => command.kind === 'line' && command.color === 'rgba(254, 220, 186, 0.5)')
  assert.ok(links.length > 2)
  assert.ok(links.every(link => link.width === 3 && Math.hypot(link.x2 - link.x1, link.y2 - link.y1) <= 3 + 1e-8))
  const levels = [...new Set(links.map(link => link.y1))]
  assert.equal(levels.length, 2)
  for (const level of levels) {
    const fragments = links.filter(link => link.y1 === level)
    assert.ok(fragments.length > 1)
    for (let index = 1; index < fragments.length; index++) near(fragments[index].x1 - fragments[index - 1].x2, 6)
  }
  assert.equal(without.filter(command => command.kind === 'line').length + links.length, connected.filter(command => command.kind === 'line').length)
})

test('role descriptions detach authored data, keep inactive/named overrides and omit renderer defaults', () => {
  const config = model({ waterfall: { styles: { positive: { color: gradient() } } },
    series: { 1: { waterfallStyles: { negative: { color: 'red' } } }, 3: { waterfallStyles: { subtotal: { name: 'Retained' } } } },
    namedSeries: [{ name: 'Delta', style: { waterfallStyles: { negative: { fillOpacity: 0.3 } } } }] })
  const description = describeChartConfiguration(config, { id: config.id, fields: { 0: 'Stage', 1: 'Delta' } })
  assert.equal(description.series[1].waterfallStyles.negative.color, 'red')
  assert.equal(description.series[1].waterfallStyles.negative.fillOpacity, 0.3)
  assert.equal(description.series[3].waterfallStyles.subtotal.name, 'Retained')
  description.series[1].waterfallStyles.positive.color.stops[0].color = 'green'
  assert.equal(config.options.waterfall.styles.positive.color.stops[0].color, 'red')
  assert.equal(chartWaterfallAuthoredStyles(model(), 1, 'Delta'), undefined)
})

test('invalid role keys, recursive styles and malformed values reject before cloning', () => {
  for (const invalid of [
    { waterfall: { styles: { rising: { color: 'red' } } } },
    { series: { 1: { waterfallStyles: { positive: { waterfallStyles: {} } } } } },
    { waterfall: { styles: { positive: { fillOpacity: -1 } } } },
    { series: { 1: { waterfallStyles: { negative: { dataPoints: { '-1': { color: 'red' } } } } } } },
  ]) assert.throws(() => cloneChartOptions(invalid))
  assert.throws(() => chartWaterfallStyleValues({ waterfallStyles: {} }), /嵌套/)
  assert.throws(() => patchChartWaterfallStyle({}, 1, 'subtotal', { color: 'not-a-color' }), /颜色/)
  assert.throws(() => patchChartWaterfallStyle({}, 1, 'unknown', { color: 'red' }), /样式目标/)
})

test('draft controls edit detached role and connector values through the production configuration contract', () => {
  const chart = model({ waterfall: { styles: { negative: { color: 'red' } } } }), before = structuredClone(chart)
  const controls = chartWaterfallStyleControls('positive', 1, { field: 1, position: 0, name: 'Delta' })
  const next = { ...chart, options: cloneChartOptions(chart.options) }
  controls.find(control => control.key.endsWith('.fillOpacity')).write(next, 0.5)
  controls.find(control => control.key.endsWith('.color') && control.kind === 'chart-color').write(next, gradient())
  const point = chartWaterfallPointControls('positive', 0, 1)
  point.find(control => control.key.endsWith('.color') && control.kind === 'chart-color').write(next, 'cyan')
  const connector = chartWaterfallConnectorControls()
  connector[0].write(next, true)
  connector.find(control => control.key === 'waterfall.connector.width').write(next, 4)
  validateChartOptions(next.options)
  assert.equal(chartWaterfallConnector(next).width, 4)
  assert.equal(next.options.series[1].waterfallStyles.positive.dataPoints[0].fillPattern, null)
  assert.deepEqual(chart, before)
  connector[0].reset(next)
  assert.equal(chartWaterfallConnector(next), undefined)
  assert.equal(next.options.waterfall.styles.negative.color, 'red')
})

test('snapshot merging preserves unrelated role groups and atomic gradient/texture values', () => {
  const before = { waterfall: { styles: { positive: { color: gradient(), fillPattern: pattern }, negative: { color: 'red' } } } }
  const next = mergeChartOptions(before, { waterfall: { styles: { positive: { color: 'blue', fillPattern: null } } } })
  assert.equal(next.waterfall.styles.positive.color, 'blue')
  assert.equal(next.waterfall.styles.positive.fillPattern, null)
  assert.equal(next.waterfall.styles.negative.color, 'red')
  assert.deepEqual(JSON.parse(JSON.stringify(next)), next)
  assert.equal(typeof before.waterfall.styles.positive.color, 'object')
})
