import assert from 'node:assert/strict'
import { test } from 'vitest'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { drawChartData, drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { layoutChartData, layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { layoutChartWaterfall } from '../src/features/charts/chart-waterfall-layout.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { normalizeChartColor } from '../src/features/charts/chart-color.ts'
import { CHART_WATERFALL_GLOBAL_SERIES } from '../src/features/charts/chart-statistics.ts'
import { waterfallComputedFixture } from './fixtures/chart-waterfall-fixtures.mjs'

// P18 implementation contracts only; no runner, native renderer or browser ran.
const palette = ['#112233', '#223344', '#334455', '#445566']
const source = () => ({ categories: ['A', 'B'], fields: ['Label', 'First', 'Second'],
  series: [{ field: 1, name: 'First', values: [10, -3] }, { field: 2, name: 'Second', values: [4, -2] }] })
const single = () => ({ categories: ['A', 'B', 'C'], series: [{ field: 1, name: 'Delta', values: [6, -2, 4] }] })
function chart(options = {}, overrides = {}) {
  return { id: 'waterfall', sheetId: 'sheet', revision: 0, type: 'waterfall', categoryField: 0, valueFields: [1, 2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 640, height: 400 }, palette,
    legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options, ...overrides }
}
function scene(model = chart(), input = source(), name = 'arrangement') {
  const computed = waterfallComputedFixture(input, name), derived = deriveChartData(input, model, computed)
  return { computed, derived, layout: layoutDerivedChartData(derived, model, palette), commands: drawDerivedChartData(derived, model, {}, palette) }
}
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} ≈ ${expected}`)
const bars = commands => commands.filter(command => command.kind === 'rect' && command.clip && command.width > 10 && command.height > 0)

test('nonempty waterfall data requires the Rust scene across derive, public layout, drawing and ECharts', () => {
  const input = source(), model = chart()
  for (const operation of [() => deriveChartData(input, model), () => layoutChartData(input, model),
    () => drawChartData(input, model), () => chartEChartsOption({ sheets: [] }, model, {}, { preparedData: input })]) {
    assert.throws(operation, error => error.code === 'COMPUTE_NOT_READY')
  }
  const empty = { categories: [], series: [] }, commands = drawChartData(empty, model)
  assert.ok(commands.some(command => command.kind === 'text' && command.text === '暂无可绘制数据'))
})

test('layout metadata for a different source, stack or subtotal request is rejected as stale', () => {
  const input = source(), model = chart()
  for (const change of [value => { value.stackType = 'stacked' }, value => { value.useSubtotal = false },
    value => { value.sourceCategoryCount = 3 }, value => { value.seriesFields.reverse() }, value => { value.seriesFields.pop() }]) {
    const computed = waterfallComputedFixture(input, 'arrangement'); change(computed.waterfallLayout)
    assert.throws(() => deriveChartData(input, model, computed), error => error.code === 'COMPUTE_STALE_RESULT')
  }
})

test('arrangement expands categories and preserves source identities and the shared final subtotal', () => {
  const { derived, layout } = scene()
  assert.deepEqual(layout.categories, ['A', 'B', 'First-(合计)', 'A', 'B', 'Second-(合计)'])
  assert.deepEqual(layout.waterfall.points.map(point => [point.seriesPosition, point.dataPointIndex, point.categoryIndex, point.synthetic]),
    [[0, 0, 0, false], [0, 1, 1, false], [0, 2, 2, true], [1, 0, 3, false], [1, 1, 4, false], [1, 2, 5, true]])
  assert.deepEqual(layout.series.map(series => series.points.filter(point => point.synthetic).map(point => point.value)), [[9], [9]])
  assert.equal(layout.series[1].points[0], layout.waterfall.points[3])
  assert.equal(derived.waterfalls[1][0], derived.waterfallLayout.points[3])
  assert.equal(layout.series[1].points[0].baseValue, 7)
  assert.equal(layout.max, 11)
  assert.equal(layout.waterfall.connectors.length, 3)
  near(layout.waterfall.connectors[1].from.x, layout.waterfall.points[1].x)
  near(layout.waterfall.connectors[1].to.x, layout.waterfall.points[3].x)
})

test('stacked sources share category centers and the sole global subtotal remains outside source series', () => {
  const { layout } = scene(chart({ waterfall: { stackType: 'stacked' } }), source(), 'stacked')
  assert.deepEqual(layout.categories, ['A', 'B', '合计'])
  assert.deepEqual(layout.waterfall.points.map(point => [point.seriesPosition, point.dataPointIndex, point.baseValue, point.value]),
    [[0, 0, 0, 10], [1, 0, 10, 4], [0, 1, 12, -3], [1, 1, 14, -2], [CHART_WATERFALL_GLOBAL_SERIES, 2, 0, 9]])
  assert.deepEqual(layout.series.map(series => series.points.length), [2, 2])
  assert.equal(layout.max, 14)
  for (let row = 0; row < 2; row++) near(layout.series[0].points[row].x, layout.series[1].points[row].x)
  assert.equal(layout.waterfall.points.at(-1).synthetic, true)
  assert.equal(layout.waterfall.connectors.length, 1)
  near(layout.waterfall.connectors[0].from.y, layout.plot.y)
  near(layout.waterfall.connectors[0].to.y, layout.plot.y)
})

test('both arrangements use the full category band rather than divide mark width by the source count', () => {
  for (const [stackType, name] of [['arrangement', 'arrangement'], ['stacked', 'stacked']]) {
    const { commands, layout } = scene(chart({ waterfall: { stackType }, bar: { gapWidth: 150 } }), source(), name)
    const width = layout.plot.width / layout.categories.length / 2.5 - 1
    assert.equal(bars(commands).length, stackType === 'stacked' ? 5 : 6)
    for (const mark of bars(commands)) near(mark.width, width)
  }
})

test('arrangement connector ends use their role axes and cross the inserted subtotal slot', () => {
  const model = chart({ rightYAxis: { scaleType: 'linear', min: 0, max: 40 },
    waterfall: { styles: { negative: { rightAxis: true } }, connector: { color: '#fedcba', dashType: 'solid' } } },
    { yAxis: { scaleType: 'linear', min: 0, max: 20 } })
  const { layout, commands } = scene(model), connector = layout.waterfall.connectors[1]
  near(connector.from.y, layout.plot.y + 0.825 * layout.plot.height)
  near(connector.to.y, layout.plot.y + 0.65 * layout.plot.height)
  const line = commands.filter(command => command.kind === 'line' && command.color === normalizeChartColor('#fedcba'))[1]
  const width = bars(commands)[0].width
  near(line.x1, connector.from.x + width / 2)
  near(line.x2, connector.to.x - width / 2)
  near(line.y1, connector.from.y); near(line.y2, connector.to.y)
  assert.equal(layout.rightMax, 40)
})

test('an explicit total draws from zero and accepts a sloping incoming connector', () => {
  const input = single(); input.series[0].values[2] = 10
  const model = chart({ waterfall: { useSubtotal: false, pointRoles: [{ seriesIndex: 1, dataPointIndex: 2, role: 'total' }],
    connector: { color: '#fedcba', dashType: 'solid' } } }, { valueFields: [1] })
  const { layout, commands } = scene(model, input, 'total-bridge'), total = layout.waterfall.points[2]
  assert.equal(total.baseValue, 0); assert.equal(total.waterfallRole, 'subtotal'); assert.equal(total.synthetic, false)
  near(layout.waterfall.connectors[1].from.y, layout.plot.y + 0.6 * layout.plot.height)
  near(layout.waterfall.connectors[1].to.y, layout.plot.y)
  const links = commands.filter(command => command.kind === 'line' && command.color === normalizeChartColor('#fedcba'))
  assert.equal(links.length, 2); assert.ok(links[1].y1 > links[1].y2)
})

test('stacked connectors use the primary category cumulative even when negative marks use the secondary axis', () => {
  const model = chart({ rightYAxis: { scaleType: 'linear', min: 0, max: 40 },
    waterfall: { stackType: 'stacked', styles: { negative: { rightAxis: true } } } }, { yAxis: { scaleType: 'linear', min: 0, max: 20 } })
  const { layout } = scene(model, source(), 'stacked')
  assert.equal(layout.waterfall.points[2].rightAxis, true)
  near(layout.waterfall.points[2].y, layout.plot.y + 0.775 * layout.plot.height)
  near(layout.waterfall.connectors[0].from.y, layout.plot.y + 0.3 * layout.plot.height)
  near(layout.waterfall.connectors[0].to.y, layout.plot.y + 0.3 * layout.plot.height)
})

test('reversed axes reflect all point and connector coordinates while retaining source order and mark edges', () => {
  const options = { waterfall: { connector: { color: '#fedcba', dashType: 'solid' } } }, model = chart(options)
  const before = scene(model), reversed = scene(chart(options, { xAxis: { scaleType: 'linear', reverse: true }, yAxis: { scaleType: 'linear', reverse: true } }))
  for (let index = 0; index < before.layout.waterfall.points.length; index++) {
    const first = before.layout.waterfall.points[index], second = reversed.layout.waterfall.points[index]
    assert.equal(second.dataPointIndex, first.dataPointIndex)
    near(second.x, before.layout.plot.x * 2 + before.layout.plot.width - first.x)
    near(second.y, before.layout.plot.y * 2 + before.layout.plot.height - first.y)
    near(second.baseY, before.layout.plot.y * 2 + before.layout.plot.height - first.baseY)
  }
  const links = reversed.commands.filter(command => command.kind === 'line' && command.color === normalizeChartColor('#fedcba'))
  assert.equal(links.length, 3)
  for (const link of links) assert.ok(link.x1 > link.x2)
  near(links[1].x1, reversed.layout.waterfall.connectors[1].from.x - bars(reversed.commands)[0].width / 2)
})

test('logarithmic axes omit zero-based marks and adjacent connectors while preserving visible categories', () => {
  const { layout, commands } = scene(chart({ waterfall: { connector: { color: '#fedcba', dashType: 'solid' } } },
    { valueFields: [1], yAxis: { scaleType: 'log', min: 1, max: 10 } }), single(), 'style-changes')
  assert.equal(layout.categories.length, 4)
  assert.deepEqual(layout.waterfall.points.map(point => point.dataPointIndex), [1, 2])
  assert.equal(layout.waterfall.connectors.length, 1)
  assert.equal(bars(commands).length, 2)
  assert.ok(bars(commands).every(mark => [mark.x, mark.y, mark.width, mark.height].every(Number.isFinite)))
})

test('role dataPoints index real source rows and automatic subtotals rather than expanded category positions', () => {
  const model = chart({ series: { 2: { waterfallStyles: { positive: { dataPoints: { 0: { color: 'red' }, 3: { color: 'green' } } },
    subtotal: { dataPoints: { 2: { color: 'purple' }, 5: { color: 'yellow' } } } } } } })
  const { commands } = scene(model), marks = bars(commands)
  assert.equal(marks[3].fill, normalizeChartColor('red'))
  assert.equal(marks[5].fill, normalizeChartColor('purple'))
})

test('global stacked subtotal consumes only shared role/point styles and has one legend item', () => {
  const model = chart({ waterfall: { stackType: 'stacked', styles: { subtotal: { name: 'Grand', color: 'pink',
    dataPoints: { 2: { color: 'orange' } }, label: { visible: true, contentType: 'series' } } } },
    series: { 1: { waterfallStyles: { subtotal: { name: 'First override', color: 'red' } } },
      2: { waterfallStyles: { subtotal: { name: 'Second override', color: 'blue' } } } } }, { legend: { visible: true, position: 'bottom' } })
  const { commands } = scene(model, source(), 'stacked')
  assert.equal(bars(commands).at(-1).fill, normalizeChartColor('orange'))
  const texts = commands.filter(command => command.kind === 'text').map(command => command.text)
  assert.equal(texts.filter(text => text === 'Grand').length, 3) // category + mark label + one legend item
  assert.equal(texts.includes('First override'), false); assert.equal(texts.includes('Second override'), false)
})

test('connector defaults stay neutral under a theme axis override and omit authored false', () => {
  const input = source(), computed = waterfallComputedFixture(input, 'arrangement'), model = chart({ waterfall: { connector: true } })
  const commands = drawChartData(input, model, { axisColor: 'red' }, computed)
  assert.ok(commands.some(command => command.kind === 'line' && command.color === normalizeChartColor('#8C8C8C')))
  const omitted = drawChartData(input, chart({ waterfall: { connector: false } }), { axisColor: 'red' }, computed)
  assert.equal(omitted.some(command => command.kind === 'line' && command.color === normalizeChartColor('#8C8C8C')), false)
})

test('an empty authoritative layout does not invent automatic subtotal slots', () => {
  const layout = layoutChartWaterfall({ categories: [], series: [{ field: 1, name: 'Empty', values: [] }] }, chart(),
    { stackType: 'arrangement', useSubtotal: true, sourceCategoryCount: 0, seriesFields: [1], cumulative: 0, points: [], connectors: [] },
    { x: 0, y: 0, width: 400, height: 300 }, { x: 40, y: 20, width: 340, height: 240 }, palette)
  assert.deepEqual(layout.categories, []); assert.deepEqual(layout.waterfall.points, []); assert.deepEqual(layout.waterfall.connectors, [])
})

test('ECharts graphic preserves full-band stacked marks, global style and connector clipping from the Rust scene', () => {
  const model = chart({ waterfall: { stackType: 'stacked', connector: { color: '#fedcba', width: 3, dashType: 'solid' },
    styles: { subtotal: { color: 'orange', border: { color: 'black', width: 2 } } } } })
  const input = source(), { computed, layout } = scene(model, input, 'stacked'), native = chartEChartsOption({ sheets: [] }, model, {}, computed)
  assert.equal(native.series, undefined)
  const elements = native.graphic.elements, shapes = elements.map(element => element.type === 'group' ? element.children[0] : element)
  const marks = shapes.filter(element => element.type === 'rect' && element.shape.width > 10 && element.style.fill !== '#ffffff' && element.style.fill !== 'rgba(255, 255, 255, 1)')
  assert.equal(marks.length, 5)
  near(marks[0].shape.x, marks[1].shape.x)
  near(marks[0].shape.width, layout.plot.width / 3 / 2.5 - 1)
  assert.equal(marks.at(-1).style.fill, normalizeChartColor('orange'))
  assert.equal(marks.at(-1).style.stroke, normalizeChartColor('black'))
  assert.equal(marks.at(-1).style.lineWidth, 2)
  const linkGroup = elements.find(element => element.type === 'group' && element.children[0].type === 'line'
    && element.children[0].style.stroke === normalizeChartColor('#fedcba'))
  assert.deepEqual(linkGroup.clipPath.shape, layout.plot)
  assert.equal(linkGroup.children[0].style.lineWidth, 3)
  near(linkGroup.children[0].shape.y1, layout.plot.y)
})
