import assert from 'node:assert/strict'
import { test } from 'vitest'
import * as echarts from 'echarts'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { ChartNativeElementIndex } from '../src/features/charts/echarts-element-interaction.ts'
import { chartCanvasCommandElement, chartCanvasElementAsSeries, chartCanvasElementKey, chartCanvasElementRoute,
  chartCanvasParetoLineRuntimeId, chartCanvasParetoMark, chartCanvasSeriesBindings, readChartCanvasElement } from '../src/features/charts/chart-element-interaction.ts'
import { drawChartData } from '../src/features/charts/chart-draw.ts'
import { chartPieLayout } from '../src/features/charts/chart-pie-layout.ts'
import { chartPieAggregateStyle } from '../src/features/charts/chart-pie.ts'
import { chartLinearGradient } from '../src/features/charts/chart-color.ts'
import { waterfallComputedFixture } from './fixtures/chart-waterfall-fixtures.mjs'

// P20 drafts only. Real ECharts SVG SSR supplies native views and ZRender
// containment/clip geometry; Canvas2D, DOM/IME and file acceptance remain deferred.
const size = { width: 600, height: 420 }, palette = ['#112233', '#223344', '#334455', '#445566', '#556677']
const model = (type, options = {}) => ({ id: 'special-marks', sheetId: 'sheet', revision: 0, type,
  categoryField: 0, valueFields: [1], source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 4, endCol: 1 } },
  position: { x: 0, y: 0 }, size, palette, legend: { visible: false, position: 'bottom' },
  xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options })
const data = (values = [8, 5, 3, 2], categories = ['A', 'B', 'C', 'D']) => ({ categories,
  series: [{ field: 1, name: 'Renamed display', sourceName: 'Original', values }] })
const json = value => JSON.parse(JSON.stringify(value))
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const markCommands = (commands, kind, part) => commands.filter(command => {
  const element = chartCanvasCommandElement(command)
  return element?.kind === 'mark' && element.directEdit.kind === kind && (part === undefined || element.directEdit.part === part)
})
function render(input) {
  const { officeChartElements: runtime = { series: [], axes: [] }, officeChartKey, officeChartPointer, officeChartLegend, officeChartPlotArea, ...option } = input
  const instance = echarts.init(null, null, { renderer: 'svg', ssr: true, ...size })
  instance.setOption(option, { notMerge: true, lazyUpdate: false }); instance.renderToSVGString()
  return { input, option, runtime, instance, index: new ChartNativeElementIndex(instance, option, runtime, size.width, size.height), dispose: () => instance.dispose() }
}
function measuredHit(index, element) {
  const target = index.target(element); assert.ok(target, `Missing ${chartCanvasElementKey(element)}`)
  const { bounds } = target, accept = point => { const hit = index.hit(point); return hit?.key === target.key ? { point, hit } : undefined }
  const center = accept({ x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 }); if (center) return center
  for (let y = bounds.y; y <= bounds.y + bounds.height; y += 0.5) for (let x = bounds.x; x <= bounds.x + bounds.width; x += 0.5) {
    const result = accept({ x, y }); if (result) return result
  }
  assert.fail(`Measured target is shielded or clipped: ${target.key}`)
}
function option(chart, source, computed = { preparedData: source }) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.id = chart.sheetId; workbook.activeSheetId = sheet.id
  const rows = [['Category', ...source.series.map(series => series.sourceName ?? series.name)],
    ...source.categories.map((category, index) => [category, ...source.series.map(series => series.values[index] ?? null)])]
  sheet.cells = Object.fromEntries(rows.flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  workbook.charts = [chart]
  return chartEChartsOption(workbook, chart, {}, computed)
}
const pareto = () => {
  const source = data([2, 6, 4, 3], ['C', 'A', 'B', 'A'])
  return { source, computed: { preparedData: source, pareto: { categories: ['A', 'B', 'C'], values: [9, 4, 2], cumulativePercent: [60, 86.66, 100] } } }
}

test('native Pareto bar and cumulative-line views map their real geometry to separate dedicated author groups', () => {
  const { source, computed } = pareto(), chart = model('pareto'), value = render(option(chart, source, computed))
  try {
    assert.equal(value.option.series.at(-1).id, chartCanvasParetoLineRuntimeId(chart.id))
    assert.deepEqual(value.runtime.series.map(entry => entry.directEdit.part), ['bar', 'line'])
    assert.ok(value.runtime.series.every(entry => !Object.hasOwn(entry, 'binding')))
    assert.equal(value.runtime.series[0].directEdit.sourceName, 'Original')
    for (const part of ['bar', 'line']) {
      const measured = measuredHit(value.index, chartCanvasParetoMark(part, source.series[0]))
      assert.equal(measured.hit.element.kind, 'mark'); assert.equal(measured.hit.element.directEdit.part, part)
      assert.equal(value.index.hit(measured.point, true).key, measured.hit.key)
    }
    assert.ok(value.index.list().every(target => target.element.kind !== 'dataPoint' && target.element.kind !== 'series'))
  } finally { value.dispose() }
})

test('graphic Pareto gradients and attached data labels keep the same dedicated identities as native marks', () => {
  const { source, computed } = pareto(), chart = model('pareto', { pareto: { barStyle: { label: { visible: true } },
    cumulativeLineStyle: { color: chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 1 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }),
      label: { visible: true, position: 'top' } } } }), commands = drawChartData(source, chart, {}, computed)
  for (const part of ['bar', 'line']) {
    assert.ok(markCommands(commands, 'paretoStyle', part).some(command => command.kind === 'text'))
    assert.ok(markCommands(commands, 'paretoStyle', part).some(command => command.kind !== 'text'))
  }
  assert.ok(commands.every(command => chartCanvasCommandElement(command)?.kind !== 'dataPoint'))
  const value = render(option(chart, source, computed))
  try { for (const part of ['bar', 'line']) measuredHit(value.index, chartCanvasParetoMark(part, source.series[0])) }
  finally { value.dispose() }
})

test('real and automatic arrangement Waterfall totals share the correct source role without acquiring fake point identities', () => {
  const source = data([6, -2, 4], ['A', 'B', 'C']), chart = model('waterfall', { waterfall: { pointRoles: [{ seriesIndex: 1, dataPointIndex: 2, role: 'total' }], connector: true },
    allSeriesStyle: { label: { visible: true, position: 'inside' } } }), computed = waterfallComputedFixture(source, 'style-total'),
    commands = drawChartData(source, chart, {}, computed), binding = chartCanvasSeriesBindings(source, chart.options)[0]
  const totals = markCommands(commands, 'waterfallStyle', 'subtotal')
  assert.equal(totals.filter(command => command.kind === 'rect').length, 2)
  assert.ok(totals.every(command => chartCanvasCommandElement(command).directEdit.binding.seriesId === binding.seriesId))
  assert.ok(commands.every(command => !['dataPoint', 'series'].includes(chartCanvasCommandElement(command)?.kind)))
  const value = render(option(chart, source, computed))
  try {
    for (const part of ['positive', 'negative', 'subtotal']) {
      const hit = measuredHit(value.index, { kind: 'mark', directEdit: { kind: 'waterfallStyle', part, binding } }).hit
      assert.equal(hit.element.directEdit.binding.sourceName, 'Original'); assert.equal(hit.element.dataIndex, undefined)
    }
  } finally { value.dispose() }
})

test('the stacked Waterfall global subtotal retains Rust synthetic ownership independently of every source-bound role', () => {
  const source = { categories: ['A', 'B'], series: [{ field: 1, name: 'First', values: [10, -3] }, { field: 2, name: 'Second', values: [4, -2] }] },
    chart = model('waterfall', { waterfall: { stackType: 'stacked', useSubtotal: true, connector: true }, allSeriesStyle: { label: { visible: true, position: 'inside' } } }),
    computed = waterfallComputedFixture(source, 'stacked'), commands = drawChartData(source, chart, {}, computed)
  const totals = markCommands(commands, 'waterfallStyle', 'subtotal')
  assert.equal(totals.filter(command => command.kind === 'rect').length, 1)
  assert.ok(totals.every(command => chartCanvasCommandElement(command).directEdit.binding === undefined))
  const value = render(option(chart, source, computed))
  try {
    const target = { kind: 'mark', directEdit: { kind: 'waterfallStyle', part: 'subtotal' } }, measured = measuredHit(value.index, target)
    assert.equal(measured.hit.key, 'mark:waterfall:subtotal:global'); assert.equal(measured.hit.element.directEdit.binding, undefined)
    assert.ok(value.index.list().some(item => item.key === 'mark:waterfall:positive:field:1'))
    assert.ok(value.index.list().some(item => item.key === 'mark:waterfall:negative:field:2'))
  } finally { value.dispose() }
})

test('both composite Pie layouts expose real source slices plus a dedicated aggregate and connection-line target', () => {
  const source = data(), before = json(source)
  for (const type of ['pieOfPie', 'barOfPie']) {
    const chart = model(type, { pie: { label: { visible: true, position: 'inside' }, aggregateLabel: 'Other categories',
      aggregateStyle: { color: '#abcdef', fillOpacity: 0.6 }, secondaryPlot: { type: type === 'barOfPie' ? 'bar' : 'pie', split: { type: 'position', count: 2 } } } }),
      commands = drawChartData(source, chart), aggregate = { kind: 'mark', directEdit: { kind: 'pieAggregateStyle', sourceField: 1, sourceName: 'Original', aggregateIndex: 4 } },
      value = render(option(chart, source))
    try {
      const aggregateHit = measuredHit(value.index, aggregate).hit
      assert.equal(aggregateHit.element.dataIndex, undefined); assert.equal(aggregateHit.element.directEdit.aggregateIndex, 4)
      measuredHit(value.index, { kind: 'mark', directEdit: { kind: 'pieConnectorStyle' } })
      const binding = chartCanvasSeriesBindings(source)[0]
      measuredHit(value.index, { kind: 'mark', directEdit: { kind: 'pieSliceStyle', sourceField: 1, sourceName: 'Original', sourcePointIndex: 2, plot: 'secondary' } })
      assert.equal(value.index.target({ kind: 'dataPoint', binding, dataIndex: 4 }), undefined)
      assert.ok(markCommands(commands, 'pieAggregateStyle').some(command => command.kind === 'text' && command.text === 'Other categories'))
    } finally { value.dispose() }
  }
  assert.deepEqual(json(source), before)
})

test('Pie connector style changes both real segments, and hidden or zero-paint lines provide no selectable graphic target', () => {
  const source = data(), base = model('pieOfPie', { pie: { secondaryPlot: { split: { type: 'position', count: 2 } } } }),
    plain = markCommands(drawChartData(source, base), 'pieConnectorStyle')
  assert.equal(plain.length, 2); assert.ok(plain.every(command => command.kind === 'line' && command.width === 1))
  const styled = model('pieOfPie', { pie: { secondaryPlot: { split: { type: 'position', count: 2 }, connector: { color: '#ee0000', width: 3, opacity: 0.4, dashType: 'dashed' } } } }),
    dashed = markCommands(drawChartData(source, styled), 'pieConnectorStyle')
  assert.ok(dashed.length > 2); assert.ok(dashed.every(command => command.width === 3 && command.color.includes('0.4')))
  for (const hidden of [{ visible: false }, { width: 0 }, { opacity: 0 }]) {
    const chart = model('pieOfPie', { pie: { secondaryPlot: { split: { type: 'position', count: 2 }, connector: hidden } } })
    assert.equal(markCommands(drawChartData(source, chart), 'pieConnectorStyle').length, 0)
    const value = render(option(chart, source))
    try { assert.equal(value.index.target({ kind: 'mark', directEdit: { kind: 'pieConnectorStyle' } }), undefined) }
    finally { value.dispose() }
  }
})

test('aggregate appearance survives repartitioning, and its explosion moves the actual connector endpoints with the generated sector', () => {
  const source = data(), chart = model('pieOfPie', { series: { 1: { dataPoints: { 4: { color: '#111111' } } } }, pie: {
    aggregateStyle: { color: '#abcdef', explosion: 25 }, secondaryPlot: { split: { type: 'position', count: 2 } },
  } }), before = json(chart), plot = { x: 10, y: 10, width: 560, height: 350 }, layout = chartPieLayout(source, chart, plot),
    aggregate = layout.sectors.find(sector => sector.part.index === 4), primary = layout.sectors.find(sector => !sector.secondary && sector.part.index !== 4)
  near(Math.hypot(aggregate.center.x - primary.center.x, aggregate.center.y - primary.center.y), aggregate.outer * 0.25)
  const expectedStarts = [aggregate.start, aggregate.end].map(angle => ({ x: aggregate.center.x + aggregate.outer * Math.cos(angle), y: aggregate.center.y + aggregate.outer * Math.sin(angle) })).sort((a, b) => a.y - b.y)
  for (let index = 0; index < 2; index++) { near(layout.connectors[index][0].x, expectedStarts[index].x); near(layout.connectors[index][0].y, expectedStarts[index].y) }
  assert.equal(chartPieAggregateStyle(chart, source.series[0], 4, palette).color, '#abcdef')
  const more = data([8, 5, 3, 2, 1], ['A', 'B', 'C', 'D', 'E']), later = drawChartData(more, chart)
  assert.ok(markCommands(later, 'pieAggregateStyle').every(command => chartCanvasCommandElement(command).directEdit.aggregateIndex === 5))
  assert.equal(chartPieAggregateStyle(chart, more.series[0], 5, palette).color, '#abcdef'); assert.deepEqual(json(chart), before)
})

test('mark metadata validates role and source ownership, and More refines existing sections with validated local field focus', () => {
  const binding = chartCanvasSeriesBindings(data())[0], waterfall = { kind: 'mark', directEdit: { kind: 'waterfallStyle', part: 'negative', binding } },
    aggregate = { kind: 'mark', directEdit: { kind: 'pieAggregateStyle', sourceField: 1, sourceName: 'Original', aggregateIndex: 4 } }
  assert.deepEqual(readChartCanvasElement(waterfall), waterfall); assert.notEqual(readChartCanvasElement(waterfall).directEdit.binding, binding)
  assert.deepEqual(readChartCanvasElement(aggregate), aggregate); assert.deepEqual(chartCanvasElementAsSeries(aggregate), aggregate)
  for (const directEdit of [{ kind: 'waterfallStyle', part: 'base', binding }, { kind: 'waterfallStyle', part: 'positive' },
    { kind: 'waterfallStyle', part: 'negative', binding: { ...binding, selector: 9 } }, { kind: 'paretoStyle', part: 'unknown' },
    { ...aggregate.directEdit, aggregateIndex: -1 }, { ...aggregate.directEdit, sourceField: 1.5 }, { kind: 'pieAggregateStyle' }])
    assert.equal(readChartCanvasElement({ kind: 'mark', directEdit }), undefined)
  assert.deepEqual(chartCanvasElementRoute(chartCanvasParetoMark('bar')), { tab: 'style', section: 'paretoBarSeries' })
  assert.deepEqual(chartCanvasElementRoute(chartCanvasParetoMark('line')), { tab: 'style', section: 'paretoLineSeries' })
  assert.deepEqual(chartCanvasElementRoute(waterfall), { tab: 'style', section: 'waterfallSeries', seriesId: 'field:1', focus: { kind: 'waterfallRole', role: 'negative' } })
  assert.deepEqual(chartCanvasElementRoute({ kind: 'mark', directEdit: { kind: 'waterfallStyle', part: 'subtotal' } }),
    { tab: 'style', section: 'waterfallSeries', focus: { kind: 'waterfallRole', role: 'subtotal' } })
  assert.deepEqual(chartCanvasElementRoute(aggregate), { tab: 'style', section: 'pieStyle', focus: { kind: 'pieAggregate' } })
  assert.deepEqual(chartCanvasElementRoute({ kind: 'mark', directEdit: { kind: 'pieConnectorStyle' } }),
    { tab: 'style', section: 'pieStyle', focus: { kind: 'pieConnector' } })
})
