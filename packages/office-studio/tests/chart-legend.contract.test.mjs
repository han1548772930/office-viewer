import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createChartDefinition, defaultChartLegend, updateChartDefinition } from '../src/features/charts/chart-engine.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { chartAxisPointerCommands } from '../src/features/charts/chart-axis-pointer.ts'
import { chartCommonControls } from '../src/features/charts/chart-settings-controls.ts'
import { ChartLegendPositions, ChartLegendSelectModes, chartLegendPatch, chartLegendGeometry, chartLegendItems,
  chartLegendRelationScene, chartLegendSankeyScene, chartLegendViewData, chartLegendViewLayout,
  reconcileChartLegendSelection, toggleChartLegendSelection } from '../src/features/charts/chart-legend.ts'
import { chartNativeLegendHit, chartNativeLegendNames } from '../src/features/charts/echarts-legend.ts'

// Written for the later unified verification phase. Nothing in this file has run.
const palette = ['#ff0000', '#008800', '#0000ff']
function model(patch = {}) {
  return { id: 'legend-chart', sheetId: 'sheet', type: 'bar', revision: 0,
    source: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 3 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: 520, height: 360 }, categoryField: 0, valueFields: [1, 2, 3],
    legend: { visible: true, position: 'bottom', selectMode: 'multiple' }, palette,
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, ...patch }
}
function data() {
  return { categories: ['North', 'South'], series: [
    { field: 1, name: 'Revenue', values: [4, 7] }, { field: 2, name: 'Costs', values: [2, 3] }, { field: 3, name: 'Profit', values: [2, 4] },
  ] }
}
const workbook = { sheets: [{ id: 'sheet', cells: {} }] }
const option = (chart, source = data()) => chartEChartsOption(workbook, chart, {}, { preparedData: source })

test('legend author patches validate before change, detach fonts and translate Hide without discarding position', () => {
  const font = { color: '#123456', family: 'Georgia' }, current = { visible: true, position: 'left' }
  const patch = chartLegendPatch(current, { selectMode: 'single', wrap: true, font })
  font.family = 'Arial'
  assert.equal(patch.font.family, 'Georgia')
  const flat = chartLegendPatch(patch, { color: '#abcdef', fontSize: 15, bold: true, italic: false, family: 'Cambria' })
  assert.deepEqual(flat.font, { color: '#abcdef', fontSize: 15, bold: true, italic: false, family: 'Cambria' })
  for (const field of ['color', 'fontSize', 'bold', 'italic', 'family']) assert.equal(Object.hasOwn(flat, field), false)
  assert.deepEqual(chartLegendPatch(flat, { font: { family: 'Georgia', bold: false }, family: 'Arial' }).font, { family: 'Arial', bold: false })
  assert.equal(chartLegendPatch(flat, { font: undefined }).font, undefined)
  assert.deepEqual(chartLegendPatch(patch, false), { ...patch, visible: false })
  assert.equal(chartLegendPatch(patch, { position: ChartLegendPositions.Hide }).position, 'left')
  assert.equal(chartLegendPatch(patch, { position: ChartLegendPositions.Hide, visible: true }).visible, false)
  const before = structuredClone(patch)
  for (const bad of [null, [], { selectMode: 'toggle' }, { wrap: 1 }, { visible: 'yes' }, { position: 'middle' }, { bold: 'yes' }]) {
    assert.throws(() => chartLegendPatch(patch, bad), TypeError)
  }
  assert.throws(() => chartLegendPatch(patch, { font: { fontSize: -1 } }), RangeError)
  assert.throws(() => chartLegendPatch(patch, { fontSize: -1 }), RangeError)
  assert.deepEqual(patch, before)
  const created = createChartDefinition(workbook, 'sheet', { ...model(), legend: patch }, 'legend-chart')
  assert.equal(created.legend.selectMode, ChartLegendSelectModes.Single)
  assert.equal(created.legend.wrap, true)
  assert.throws(() => createChartDefinition(workbook, 'sheet', { ...model(), legend: { visible: true, position: 'bottom', selectMode: 'toggle' } }), TypeError)
  const reset = updateChartDefinition(workbook, created, { legend: { ...defaultChartLegend(created.type), font: undefined, selectMode: undefined, wrap: undefined } })
  assert.deepEqual(reset.legend, { visible: true, position: 'bottom' })
})

test('view selection handles duplicate labels, single mode, removed/new names and prototype-shaped labels without mutating input', () => {
  const previous = { Revenue: true, Costs: false, Removed: false }, names = ['Revenue', '', 'Costs', 'Revenue', '__proto__']
  const multiple = reconcileChartLegendSelection(names, 'multiple', previous)
  assert.deepEqual(multiple, Object.fromEntries([['Revenue', true], ['Costs', false], ['__proto__', true]]))
  assert.equal(Object.getPrototypeOf(multiple), Object.prototype)
  assert.deepEqual(previous, { Revenue: true, Costs: false, Removed: false })
  const single = toggleChartLegendSelection(names, 'single', multiple, 'Costs')
  assert.equal(Object.values(single).filter(Boolean).length, 1)
  assert.equal(single.Costs, true)
  assert.deepEqual(toggleChartLegendSelection(names, 'single', single, 'Costs'), single)
  assert.equal(toggleChartLegendSelection(names, 'multiple', multiple, 'Revenue').Revenue, false)
  assert.equal(reconcileChartLegendSelection(['New', 'Costs'], 'multiple', multiple).New, true)
  assert.equal(Object.values(reconcileChartLegendSelection(names, 'close', previous)).every(Boolean), true)
})

test('settings preserve authored resets and fixed type/position visibility', () => {
  const chart = model(), controls = chartCommonControls(chart), mode = controls.find(control => control.key === 'legend.selectMode'), wrap = controls.find(control => control.key === 'legend.wrap')
  assert.equal(mode.visible(chart), true)
  assert.equal(mode.visible(model({ type: 'heatmap' })), false)
  assert.equal(mode.visible(model({ type: 'pieOfPie' })), false)
  assert.equal(mode.visible(model({ type: 'barOfPie' })), false)
  assert.equal(wrap.visible(chart), true)
  assert.equal(wrap.visible(model({ legend: { visible: true, position: 'right', wrap: true } })), false)
  mode.write(chart, 'single'); wrap.write(chart, true)
  assert.equal(chart.legend.selectMode, 'single'); assert.equal(chart.legend.wrap, true)
  mode.reset(chart); wrap.reset(chart)
  assert.equal(chart.legend.selectMode, undefined); assert.equal(chart.legend.wrap, undefined)
  for (const type of ['treemap', 'sunburst', 'chord', 'wordCloud']) assert.equal(chartCommonControls(model({ type })).some(control => control.key.startsWith('legend.')), false)
})

test('horizontal wrapping reserves multiple bands, side legends and excessive overflow keep pagination', () => {
  const source = { categories: ['A'], series: Array.from({ length: 12 }, (_, index) => ({ field: index + 1, name: 'Series ' + index, values: [index + 1] })) }
  const plain = model({ size: { width: 340, height: 380 }, source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 12 }, orientation: 'columns' }, valueFields: source.series.map(series => series.field), legend: { visible: true, position: 'top', wrap: true } })
  const scrolling = { ...plain, legend: { ...plain.legend, wrap: false } }
  const a = option(plain, source), b = option(scrolling, source)
  assert.equal(a.legend.type, 'plain'); assert.equal(b.legend.type, 'scroll')
  assert.ok(a.grid.top > b.grid.top)
  assert.equal(chartLegendGeometry({ ...plain, legend: { ...plain.legend, position: 'right' } }, 12).wrap, false)
  assert.equal(chartLegendGeometry(plain, 400).wrap, false)
  assert.equal(option({ ...plain, legend: { ...plain.legend, position: 'right' } }, source).legend.type, 'scroll')
  const derived = deriveChartData(source, plain, { preparedData: source }), layout = layoutDerivedChartData(derived, plain)
  const names = new Set(source.series.map(series => series.name)), legendText = drawDerivedChartData(derived, plain).filter(command => command.kind === 'text' && names.has(command.text))
  assert.equal(legendText.length, source.series.length)
  assert.ok(legendText.every(command => command.y < layout.plot.y))
})

test('native legends use fixed closed default and expose names without flattening ECharts data indices', () => {
  const chart = model({ type: 'line', legend: { visible: true, position: 'bottom' } }), native = option(chart)
  assert.equal(native.legend.selectedMode, false)
  assert.equal(native.officeChartLegend.selectMode, 'close')
  const names = chartNativeLegendNames({ legend: { data: ['Same', '', 'Same', 'Next'] } })
  assert.deepEqual(names, ['Same', '', 'Same', 'Next'])
  assert.deepEqual(chartNativeLegendHit({ parent: { __legendDataIndex: 3 } }, names), { kind: 'item', name: 'Next' })
  assert.equal(chartNativeLegendHit({ parent: { __legendDataIndex: 1 } }, names), undefined)
  assert.deepEqual(chartNativeLegendHit({ name: 'pageNext' }, names), { kind: 'page', direction: 'next' })
})

test('native date and radar guides retain the painted bounds when a high-valued series is hidden', () => {
  const source = { categories: ['First', 'Second', 'Third'], categoryValues: [45292, 45323, 45352], series: [
    { field: 1, name: 'Revenue', values: [1, 3, 5] }, { field: 2, name: 'Costs', values: [100, 200, 400] }, { field: 3, name: 'Profit', values: [8, 10, 12] },
  ] }
  for (const type of ['line', 'radar']) {
    const chart = model({ type, options: { allSeriesStyle: { point: { shape: 'none' } },
      ...(type === 'line' ? { mapping: { useDateAxis: true }, series: { 3: { rightAxis: true } } } : {}) } })
    const native = option(chart, source), original = native.officeChartPointer.geometry
    assert.equal(native.officeChartLegend.rebuild, undefined)
    const masked = native.officeChartLegend.preparePointer({ Revenue: true, Costs: false, Profit: true }).geometry
    assert.deepEqual(masked.data.series[1].values, [null, null, null])
    assert.deepEqual(original.data.series[1].values, [100, 200, 400])
    assert.equal(masked.layout, original.layout)
    assert.equal(masked.radar, original.radar)
    const pointer = type === 'radar' ? { x: original.radar.center.x, y: original.radar.center.y - original.radar.radius * 0.4 }
      : { x: original.layout.plot.x + original.layout.plot.width * 0.4, y: original.layout.plot.y + original.layout.plot.height * 0.4 }
    const commands = chartAxisPointerCommands(original, pointer)
    assert.ok(commands.length > 0)
    assert.deepEqual(chartAxisPointerCommands(masked, pointer), commands)
    if (type === 'line') {
      assert.equal(native.yAxis[0].min, masked.layout.min); assert.equal(native.yAxis[0].max, masked.layout.max)
      assert.equal(native.yAxis[1].min, masked.layout.rightMin); assert.equal(native.yAxis[1].max, masked.layout.rightMax)
    } else {
      assert.equal(native.radar.indicator[0].min, masked.radar.min); assert.equal(native.radar.indicator[0].max, masked.radar.max)
    }
  }
})

test('graphic fallbacks filter actual marks, retain styles/data identities and leave source and authors unchanged', () => {
  const chart = model({ options: { allSeriesStyle: { label: { visible: true, position: 'insideEnd' } } } }), source = data(), before = structuredClone([chart, source])
  const native = option(chart, source), runtime = native.officeChartLegend
  assert.equal(typeof runtime.rebuild, 'function')
  const all = runtime.rebuild({ Revenue: true, Costs: true, Profit: true }), filtered = runtime.rebuild({ Revenue: true, Costs: false, Profit: true })
  const rectangles = value => value.option.graphic.elements.flatMap(element => element.type === 'rect' ? [element] : element.children ?? []).filter(element => element.type === 'rect' && element.style.fill !== '#ffffff')
  assert.ok(rectangles(filtered).length < rectangles(all).length)
  const derived = deriveChartData(source, chart, { preparedData: source }), mask = chartLegendViewData(derived, chart, { Costs: false }, palette)
  assert.deepEqual(mask.data.series[1].values, [null, null])
  assert.equal(mask.data.series[1].field, source.series[1].field)
  assert.deepEqual(mask.data.series[2].values, source.series[2].values)
  assert.deepEqual(chartLegendItems(derived, chart, layoutDerivedChartData(derived, chart)).map(item => item.name), ['Revenue', 'Costs', 'Profit'])
  assert.deepEqual([chart, source], before)
})

test('pie masks retain original category slots and Candle graphic fallbacks remove authored-source candles only in the view', () => {
  const source = data(), pie = model({ type: 'pie', valueFields: [1] }), derived = { data: { ...source, series: source.series.slice(0, 1) }, trends: [], errors: [] }
  const mask = chartLegendViewData(derived, pie, { North: false, South: true }, palette)
  assert.deepEqual(mask.data.categories, ['North', 'South']); assert.deepEqual(mask.data.series[0].values, [null, 7])
  assert.deepEqual(derived.data.series[0].values, [4, 7])
  const candle = { ...derived, data: { ...derived.data, candles: [{ open: 1, high: 5, low: 0, close: 3 }, { open: 3, high: 6, low: 2, close: 5 }] } }
  assert.deepEqual(chartLegendViewData(candle, model({ type: 'candlestick' }), { Revenue: false }, palette).data.candles, [null, null])
  assert.equal(candle.data.candles[0].high, 5)
})

test('graph category filtering preserves topology identities, hides incident bands/hits and leaves uncategorized Relation nodes visible', () => {
  const paint = index => ({ kind: 'rect', x: index * 20, y: 0, width: 10, height: 10, fill: '#ff0000' })
  const marks = [{ kind: 'node', index: 0, commands: [paint(0)] }, { kind: 'node', index: 1, commands: [paint(1)] }, { kind: 'link', index: 0, commands: [paint(2)] }]
  const sankey = { layout: { nodes: [{ index: 0, name: 'A', width: 10, height: 20 }, { index: 1, name: 'B', width: 10, height: 20 }], links: [{ index: 0, sourceIndex: 0, targetIndex: 1, width: 2, band: [{ x: 0, y: 0 }] }] }, paint: { nodes: [{ label: { visible: true } }, { label: { visible: true } }] }, marks, commands: marks.flatMap(mark => mark.commands) }
  const masked = chartLegendSankeyScene(sankey, { B: false })
  assert.deepEqual(masked.marks.map(mark => [mark.kind, mark.index]), [['node', 0]])
  assert.equal(masked.layout.nodes[1].width, 0); assert.equal(masked.layout.links[0].width, 0)
  assert.equal(sankey.layout.nodes[1].width, 10)
  const relation = { ...sankey, layout: { ...sankey.layout, nodes: [{ index: 0, category: 'Group' }, { index: 1 }] } }
  const category = chartLegendRelationScene(relation, { Group: false })
  assert.deepEqual(category.marks.map(mark => [mark.kind, mark.index]), [['node', 1]])
  assert.equal(category.layout, relation.layout)
})

test('Waterfall role filtering keeps cached accumulation/category labels and removes only view marks and adjacent connectors', () => {
  const chart = model({ type: 'waterfall' }), source = data(), a = { x: 10, waterfallRole: 'positive', seriesPosition: 0 }, b = { x: 20, waterfallRole: 'negative', seriesPosition: 1 }
  const layout = { categories: ['First', 'Second'], series: [{ points: [a] }, { points: [b] }], waterfall: { points: [a, b], connectors: [{ from: { x: 10 }, to: { x: 20 } }] } }
  const derived = { data: source, trends: [], errors: [], waterfallLayout: { points: ['authoritative'] } }
  const result = chartLegendViewLayout(layout, derived, chart, { Revenue: true, Costs: false }, palette)
  assert.deepEqual(result.categories, layout.categories); assert.deepEqual(result.waterfall.points, [a]); assert.deepEqual(result.waterfall.connectors, [])
  assert.deepEqual(derived.waterfallLayout.points, ['authoritative']); assert.equal(layout.waterfall.points.length, 2)
})
