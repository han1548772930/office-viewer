import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData, updateChart } from '../src/features/charts/chart-engine.ts'
import { chartDataHint } from '../src/features/charts/chart-catalog.ts'
import { chartColorChannels, chartLinearGradient, normalizeChartColor } from '../src/features/charts/chart-color.ts'
import { describeChartConfiguration } from '../src/features/charts/chart-description.ts'
import { chartPointStyle, deriveChartData } from '../src/features/charts/chart-derived.ts'
import { drawChartData } from '../src/features/charts/chart-draw.ts'
import { chartColorOpacity, chartLabelText, chartSymbol } from '../src/features/charts/chart-geometry.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { cloneChartOptions, mergeChartOptions, validateChartOptions } from '../src/features/charts/chart-options.ts'
import { CHART_PARETO_CUMULATIVE_NAME, assertChartParetoType, chartParetoBarStyle, chartParetoBarStyleValues, chartParetoCumulativeExtent, chartParetoCumulativeMarkerStyle,
  chartParetoCumulativePointStyle, chartParetoCumulativeStyle, chartParetoLineStyleValues, chartParetoSettingsCategories } from '../src/features/charts/chart-pareto.ts'
import { chartParetoBarControls, chartParetoDataPointControls, chartParetoLineControls } from '../src/features/charts/chart-pareto-settings-controls.ts'
import { chartConfigFromDraft, cloneChartDraftConfig, createChartSettingsDraft } from '../src/features/charts/chart-settings-draft.ts'
import { chartPareto } from '../src/features/charts/chart-statistics.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'

// P17 implementation drafts. No test runner, renderer, server or benchmark has run.
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const palette = ['#112233', '#223344', '#334455', '#445566', '#556677']
const gradient = () => chartLinearGradient({ start: { x: -0.2, y: 0 }, end: { x: 1.2, y: 1 },
  stops: [{ offset: 1, color: 'blue', opacity: 0.4 }, { offset: 0, color: 'rgba(255, 0, 0, 0.8)' }] })
function model(options = {}) {
  return { id: 'pareto-style', sheetId: 'sheet', revision: 0, type: 'pareto', categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } }, position: { x: 0, y: 0 },
    size: { width: 420, height: 300 }, palette, legend: { visible: true, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options }
}
const source = (categories = ['A', 'B', 'A'], values = [2, 5, 3]) => ({ categories,
  series: [{ field: 1, name: 'Count', sourceName: 'Count', values }], fields: ['Cause', 'Count'] })
const optionForData = (data, chart) => chartEChartsOption(createBlankWorkbook(), chart, {}, { preparedData: data })
function fixture(rows, options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = Object.fromEntries([['Cause', 'Count'], ...rows].flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  const config = model(options)
  config.source.range.endRow = rows.length
  return { workbook, sheet, chart: createChart(workbook, sheet.id, config) }
}
function collectGraphic(elements) {
  return elements.flatMap(element => element.type === 'group' ? collectGraphic(element.children ?? []) : [element])
}

test('signed rows aggregate before filtering, while equal totals retain first-category order and truncated percentages', () => {
  const categories = ['A', 'B', 'A', 'negative', 'zero', 'zero', 'E', 'F', 'missing', 'invalid'], values = [-5, 9, 8, -4, 2, -2, 3, 3, null, NaN]
  const before = structuredClone({ categories, values }), ordinary = chartPareto(categories, values), included = chartPareto(categories, values, true)
  assert.deepEqual(ordinary.categories, ['B', 'A', 'E', 'F'])
  assert.deepEqual(ordinary.values, [9, 3, 3, 3])
  assert.deepEqual(ordinary.cumulativePercent, [50, 66.66, 83.33, 100])
  assert.deepEqual(included.categories, ['B', 'A', 'E', 'F', 'zero'])
  assert.deepEqual(included.values, [9, 3, 3, 3, 0])
  assert.equal(included.cumulativePercent.at(-1), 100)
  assert.deepEqual({ categories, values }, before)
  assert.deepEqual(chartPareto(['A', 'B'], [Number.MAX_VALUE, Number.MAX_VALUE]).cumulativePercent, [50, 100])
  assert.throws(() => chartPareto(['A', 'A'], [Number.MAX_VALUE, Number.MAX_VALUE]), RangeError)
})

test('fixed OOe cumulative values truncate without a max-normalization rounding change and overflow uses a finite fallback', () => {
  const cases = [
    { categories: ['A', 'B'], values: [6, 4], expected: [60, 100] },
    { categories: ['A', 'B', 'C'], values: [4, 4, 4], expected: [33.33, 66.66, 100] },
    { categories: ['A', 'B'], values: [7, 5], expected: [58.33, 100] },
    { categories: ['A', 'B'], values: [101, 99], expected: [50.5, 100] },
    { categories: ['A', 'B', 'zero'], values: [6, 4, 0], expected: [60, 100, 100] },
    { categories: ['A', 'B', 'C'], values: [Number.MAX_VALUE, Number.MAX_VALUE, Number.MAX_VALUE], expected: [33.33, 66.66, 100] },
  ]
  for (const fixture of cases) {
    const input = structuredClone(fixture), result = chartPareto(fixture.categories, fixture.values, true)
    assert.deepEqual(result.cumulativePercent, fixture.expected)
    assert.deepEqual(fixture, input)
  }
})

test('Canvas layout and native Pareto data retain the fixed truncated cumulative tuples', () => {
  const data = source(['A', 'B', 'C'], [4, 4, 4]), chart = model()
  const derived = deriveChartData(data, chart), layout = layoutDerivedChartData(derived, chart), option = optionForData(data, chart)
  assert.deepEqual(derived.pareto.cumulativePercent, [33.33, 66.66, 100])
  assert.deepEqual(layout.series.at(-1).points.map(point => point.value), [33.33, 66.66, 100])
  assert.deepEqual(option.series.at(-1).data.map(point => typeof point === 'number' ? point : point.value), [33.33, 66.66, 100])
})

test('zero totals are opt-in, negative totals never return, and all-zero input has a finite zero cumulative line', () => {
  assert.deepEqual(chartPareto(['A', 'A', 'B', 'negative'], [1, -1, -0, -1]), { categories: [], values: [], cumulativePercent: [] })
  const included = chartPareto(['A', 'A', 'B', 'negative'], [1, -1, -0, -1], true)
  assert.deepEqual(included.categories, ['A', 'B'])
  included.values.forEach(value => assert.equal(value === 0, true))
  assert.deepEqual(included.cumulativePercent, [0, 0])
  const chart = model({ pareto: { includeZeroValues: true } }), data = source(['A', 'B'], [0, 0])
  assert.deepEqual(layoutDerivedChartData(deriveChartData(data, chart), chart).series.at(-1).points.map(point => point.value), [0, 0])
  assert.equal(drawChartData(data, chart).some(command => command.kind === 'text' && command.text === '暂无可绘制数据'), false)
})

test('cumulative log axes and out-of-default single bounds use finite ordered extents without changing retained percentages', () => {
  const data = source(), chart = model({ rightYAxis: { scaleType: 'log', logBase: 10 } }), derived = deriveChartData(data, chart)
  const extent = chartParetoCumulativeExtent(chart, derived.pareto.cumulativePercent), layout = layoutDerivedChartData(derived, chart), native = optionForData(data, chart)
  assert.equal(extent.min > 0 && extent.max > extent.min, true)
  assert.deepEqual([layout.rightMin, layout.rightMax], [extent.min, extent.max])
  assert.deepEqual([native.yAxis[1].min, native.yAxis[1].max], [extent.min, extent.max])
  assert.equal(layout.series.at(-1).points.every(point => Number.isFinite(point.x) && Number.isFinite(point.y)), true)
  assert.deepEqual(derived.pareto.cumulativePercent, [50, 100])
  const empty = model({ rightYAxis: { scaleType: 'log' }, pareto: { includeZeroValues: true } }), zero = source(['A', 'B'], [0, 0])
  assert.deepEqual(layoutDerivedChartData(deriveChartData(zero, empty), empty).series.at(-1).points, [])
  assert.deepEqual(optionForData(zero, empty).series.at(-1).data, [null, null])
  for (const axis of [{ scaleType: 'linear', min: 200 }, { scaleType: 'linear', max: -20 }]) {
    const result = chartParetoCumulativeExtent(model({ rightYAxis: axis }), [50, 100])
    assert.equal(Number.isFinite(result.min) && Number.isFinite(result.max) && result.min < result.max, true)
    if (axis.min !== undefined) assert.equal(result.min, axis.min)
    if (axis.max !== undefined) assert.equal(result.max, axis.max)
  }
})

test('the engine accepts signed source rows and projects exactly one selected business value field', () => {
  const { workbook, chart, sheet } = fixture([['A', -5], ['B', 9], ['A', 8], ['zero', 0], ['negative', -1]])
  const raw = readChartData(workbook, chart), derived = deriveChartData(raw, chart)
  assert.deepEqual(raw.series[0].values, [-5, 9, 8, 0, -1])
  assert.deepEqual(derived.pareto.categories, ['B', 'A'])
  assert.deepEqual(derived.data.series[0].values, [9, 3])
  assert.deepEqual(chartParetoSettingsCategories({ pareto: chartPareto(raw.categories, raw.series[0].values, true) }), ['B', 'A', 'zero'])
  assert.equal(chartDataHint('pareto').includes('合并'), true)
  sheet.cells['0:2'] = createCell(0, 2, 'Other count')
  for (let row = 1; row <= 5; row++) sheet.cells[`${row}:2`] = createCell(row, 2, row)
  const before = structuredClone(chart)
  assert.throws(() => updateChart(workbook, chart, { source: { ...chart.source, range: { ...chart.source.range, endCol: 2 } }, valueFields: [1, 2] }), RangeError)
  assert.deepEqual(chart, before)
})

test('business bars inherit their source style while the cumulative line uses independent fixed defaults', () => {
  const chart = model({ allSeriesStyle: { color: 'purple', fillOpacity: 0.4, border: { width: 18 }, point: { size: 2, shape: 'none' }, label: { visible: true } } })
  const bar = chartParetoBarStyle(chart, 1, 0, palette, 'Count'), cumulative = chartParetoCumulativeStyle(chart, palette)
  assert.equal(bar.color, 'purple')
  assert.equal(bar.fillOpacity, 0.4)
  assert.equal(cumulative.color, palette[3])
  assert.deepEqual(cumulative.border, { width: 2, opacity: 1, dashType: 'solid' })
  assert.deepEqual(cumulative.point, { size: 7, shape: 'circle' })
  assert.equal(cumulative.label.visible, undefined)
  assert.equal(cumulative.rightAxis, true)
  const marker = chartParetoCumulativeMarkerStyle(cumulative)
  assert.deepEqual(marker.border, { visible: false, width: 0 })
  const native = optionForData(source(), chart).series.at(-1)
  assert.equal(native.itemStyle.borderWidth, 0)
  assert.equal(native.symbolSize, 7)
  assert.equal(native.lineStyle.width, 2)
  assert.equal(native.lineStyle.color, palette[3])
})

test('dedicated style patches deep-merge, gradients replace atomically, and nullable resets preserve unrelated authoring', () => {
  const barInput = { color: gradient(), opacity: 0.7, border: { color: 'black', width: 3 }, label: { visible: true, bold: true }, dataPoints: { 0: { color: 'lime', border: { width: 2 } } } }
  const lineInput = { color: gradient(), opacity: 0.5, width: 4, point: { shape: 'diamond', size: 12, color: 'yellow' } }
  const initial = { mapping: { categoryIndexes: [0] }, pareto: { includeZeroValues: true, barStyle: chartParetoBarStyleValues(barInput), cumulativeLineStyle: chartParetoLineStyleValues(lineInput) } }
  const before = structuredClone(initial), replacement = chartLinearGradient({ start: { x: 1, y: 0 }, end: { x: 0, y: 1 }, stops: [{ offset: 0, color: 'white' }, { offset: 1, color: 'black' }] })
  const merged = mergeChartOptions(initial, { pareto: { barStyle: chartParetoBarStyleValues({ color: replacement, border: { opacity: 0.6 }, dataPoints: { 0: { fillOpacity: 0.2 } } }),
    cumulativeLineStyle: chartParetoLineStyleValues({ point: { size: null }, opacity: null, width: 0 }) } })
  assert.deepEqual(merged.pareto.barStyle.color, replacement)
  assert.deepEqual(merged.pareto.barStyle.border, { color: 'black', width: 3, opacity: 0.6 })
  assert.equal(merged.pareto.barStyle.dataPoints[0].color, 'lime')
  assert.equal(merged.pareto.barStyle.dataPoints[0].border.width, 2)
  assert.equal(merged.pareto.cumulativeLineStyle.width, 4)
  assert.equal(merged.pareto.cumulativeLineStyle.point.size, undefined)
  assert.equal(chartParetoCumulativeStyle({ options: merged }, palette).point.size, 7)
  assert.equal(chartParetoCumulativeStyle({ options: merged }, palette).fillOpacity, 1)
  const reset = mergeChartOptions(merged, { pareto: { cumulativeLineStyle: chartParetoLineStyleValues({ width: null }) } })
  assert.equal(Object.hasOwn(reset.pareto.cumulativeLineStyle, 'width'), true)
  assert.equal(chartParetoCumulativeStyle({ options: reset }, palette).border.width, 2)
  const cleared = mergeChartOptions(reset, { pareto: { barStyle: undefined, cumulativeLineStyle: undefined } })
  assert.equal(cleared.pareto.barStyle, undefined)
  assert.equal(cleared.pareto.cumulativeLineStyle, undefined)
  assert.equal(cleared.pareto.includeZeroValues, true)
  assert.deepEqual(cleared.mapping, initial.mapping)
  barInput.color.stops[0].color = 'pink'
  lineInput.point.size = 1
  replacement.end.x = 0.5
  assert.deepEqual(initial, before)
  assert.equal(merged.pareto.barStyle.color.end.x, 0)
  assert.equal(chartParetoLineStyleValues({ width: -1, dashType: 'dotted' }).width, undefined)
})

test('Pareto authored styles validate their real colors, labels, solid marker fields and finite width contract', () => {
  assert.doesNotThrow(() => assertChartParetoType('pareto'))
  assert.throws(() => assertChartParetoType('column'), TypeError)
  assert.doesNotThrow(() => validateChartOptions({ pareto: { barStyle: { color: gradient(), label: { position: 'outside', contentType: 'percentage' } },
    cumulativeLineStyle: { color: gradient(), width: -1, label: { position: 'inside', italic: true }, dataPoints: { 0: { color: gradient(), shape: 'none', border: { width: 0 } } } } } }))
  for (const patch of [{ label: { position: 'outside' } }, { opacity: 1.1 }, { point: { color: gradient() } }, { width: Infinity }, { width: 101 }, { dashType: 'custom' }]) {
    assert.throws(() => chartParetoLineStyleValues(patch), error => error instanceof TypeError || error instanceof RangeError)
  }
  assert.throws(() => chartParetoBarStyleValues({ dataPoints: { 0: { color: { type: 'linear-gradient', start: { x: 0, y: 0 } } } } }), TypeError)
  assert.throws(() => cloneChartOptions({ pareto: { includeZeroValues: 'true' } }), TypeError)
  assert.throws(() => chartParetoLineStyleValues(null), TypeError)
})

test('bar-specific overrides retain source display identity and layer materialized point styles over source-field/name styles', () => {
  const chart = model({ allSeriesStyle: { color: 'red', fillOpacity: 0.4, border: { color: 'black', width: 1 }, dataPoints: { 0: { color: 'cyan', border: { color: 'white' } } } },
    series: { 1: { color: 'blue', border: { width: 2 } } }, namedSeries: [{ name: 'Count', style: { name: 'Business count', color: 'green', label: { bold: true } } }],
    pareto: { barStyle: { color: gradient(), opacity: 0.8, border: { opacity: 0.6 }, label: { visible: true, contentType: 'percentage' }, dataPoints: { 0: { fillOpacity: 0.2, border: { width: 3 } } } } } })
  const bar = chartParetoBarStyle(chart, 1, 0, palette, 'Count'), point = chartPointStyle(bar, 0, 'pareto')
  assert.equal(bar.name, 'Business count')
  assert.deepEqual(bar.color, gradient())
  assert.equal(bar.label.bold, true)
  assert.equal(point.color, 'cyan')
  assert.equal(point.fillOpacity, 0.2)
  assert.deepEqual(point.border, { color: 'white', width: 3, opacity: 0.6 })
  assert.deepEqual(chartParetoCumulativeStyle(chart, palette).color, palette[3])
})

test('native bars and cumulative line consume independent appearance and business versus cumulative percentage labels', () => {
  const chart = model({ pareto: { barStyle: { color: 'purple', opacity: 0.7, border: { color: 'black', width: 3, opacity: 0.8 },
    label: { visible: true, contentType: 'percentage', position: 'outside', fontSize: 12 }, dataPoints: { 0: { fillOpacity: 0.3 } } },
    cumulativeLineStyle: { color: 'rgba(255, 0, 0, 0.8)', opacity: 0.5, width: 5, dashType: 'dashed', point: { shape: 'diamond', size: 12, color: 'green' },
      label: { visible: true, contentType: 'percentage', position: 'top', italic: true } } } })
  const option = optionForData(source(), chart), bar = option.series[0], cumulative = option.series[1]
  assert.equal(bar.type, 'bar')
  near(chartColorChannels(bar.itemStyle.color)[3], 0.7)
  assert.equal(bar.itemStyle.opacity, 1)
  assert.equal(bar.itemStyle.borderWidth, 3)
  near(chartColorChannels(bar.itemStyle.borderColor)[3], 0.8)
  near(chartColorChannels(bar.data[0].itemStyle.color)[3], 0.3)
  assert.equal(bar.data[0].itemStyle.opacity, 1)
  near(chartColorChannels(bar.data[0].itemStyle.borderColor)[3], 0.8)
  assert.deepEqual(bar.data.map(point => point.label.formatter), ['50%', '50%'])
  assert.deepEqual(cumulative.data.map(point => point.label.formatter), ['50%', '100%'])
  assert.equal(cumulative.lineStyle.width, 5)
  assert.equal(cumulative.lineStyle.type, 'dashed')
  assert.equal(cumulative.lineStyle.opacity, 0.5)
  assert.equal(cumulative.itemStyle.color, 'green')
  assert.equal(cumulative.itemStyle.opacity, 0.5)
  assert.equal(cumulative.itemStyle.borderWidth, 0)
  assert.equal(cumulative.symbol, 'diamond')
  assert.equal(cumulative.yAxisIndex, 1)
  assert.deepEqual(option.yAxis.slice(1).map(axis => [axis.min, axis.max]), [[0, 100]])
  assert.deepEqual(option.legend.data, ['Count', CHART_PARETO_CUMULATIVE_NAME])
  const commands = drawChartData(source(), chart)
  assert.equal(commands.filter(command => command.kind === 'text' && command.text === '50%').length, 3)
  assert.equal(commands.filter(command => command.kind === 'text' && command.text === '100%').length, 1)
})

test('percentage label number formats apply one percent scale and suffix to business and cumulative values', () => {
  assert.equal(chartLabelText({ contentType: 'percentage', format: '0.0%' }, 0, 'A', 'Count', 75), '75.0%')
  assert.equal(chartLabelText({ contentType: 'percentage', format: '0.00' }, 0, 'A', 'Count', 75), '75.00%')
  const chart = model({ pareto: { barStyle: { label: { visible: true, contentType: 'percentage', format: '0.0%' } },
    cumulativeLineStyle: { label: { visible: true, contentType: 'percentage', format: '0.0%' } } } })
  const native = optionForData(source(), chart)
  assert.deepEqual(native.series[0].data.map(point => point.label.formatter), ['50.0%', '50.0%'])
  assert.deepEqual(native.series[1].data.map(point => point.label.formatter), ['50.0%', '100.0%'])
  const text = drawChartData(source(), chart).filter(command => command.kind === 'text').map(command => command.text)
  assert.equal(text.filter(value => value === '50.0%').length, 3)
  assert.equal(text.filter(value => value === '100.0%').length, 1)
  assert.equal(text.some(value => value.includes('%%')), false)
})

test('explicit Pareto borders use raw fill-color defaults and zero-width rectangles emit no outline', () => {
  for (const border of [{ width: 3, opacity: 0.75 }, { color: 'green' }, { opacity: 0.75 }, { width: 0 }]) {
    const chart = model({ pareto: { barStyle: { color: 'rgba(255, 0, 0, 0.8)', opacity: 0.5, border } } }), data = source(['A', 'B'], [20, 10])
    const item = optionForData(data, chart).series[0].itemStyle, commands = drawChartData(data, chart)
    const bars = commands.filter(command => command.kind === 'rect' && command.width > 20 && command.fill === normalizeChartColor('rgba(255, 0, 0, 0.4)'))
    assert.equal(bars.length, 2)
    assert.equal(item.opacity, 1)
    near(chartColorChannels(item.color)[3], 0.4)
    assert.equal(item.borderWidth, border.width ?? 1)
    if (border.width === 0) bars.forEach(command => assert.equal(command.stroke, undefined))
    else bars.forEach(command => {
      assert.equal(command.strokeWidth, border.width ?? 1)
      assert.deepEqual(chartColorChannels(command.stroke), chartColorChannels(item.borderColor))
      near(chartColorChannels(command.stroke)[3], border.color ? 1 : 0.8 * (border.opacity ?? 1))
    })
  }
  const dashed = model({ pareto: { barStyle: { color: 'rgba(255, 0, 0, 0.8)', opacity: 0.5, border: { width: 2, opacity: 0.75, dashType: 'dashed' } } } })
  const item = optionForData(source(), dashed).series[0].itemStyle, commands = drawChartData(source(), dashed)
  assert.equal(item.borderType, 'dashed')
  const strokes = commands.filter(command => command.kind === 'line' && command.color === normalizeChartColor(item.borderColor))
  assert.equal(strokes.length > 8, true)
  strokes.forEach(command => { assert.equal(command.width, 2); assert.deepEqual(chartColorChannels(command.color), chartColorChannels(item.borderColor)) })
})

test('explicit bar gradients remain native while cumulative gradients and overrides use the shared graphic scene', () => {
  const authored = gradient(), chart = model({ autoGradientFill: false, pareto: { barStyle: { color: authored, opacity: 0.5 } } }), before = structuredClone(chart)
  const native = optionForData(source(), chart)
  assert.equal(native.series[0].itemStyle.color.type, 'linear')
  assert.deepEqual(native.series[0].itemStyle.color.colorStops.map(stop => stop.offset), [0, 1])
  near(chartColorChannels(native.series[0].itemStyle.color.colorStops[0].color)[3], 0.4)
  chart.options.pareto.cumulativeLineStyle = { color: authored, opacity: 0.5, width: 3 }
  const graphic = optionForData(source(), chart), commands = drawChartData(source(), chart), elements = collectGraphic(graphic.graphic.elements)
  assert.deepEqual(graphic.series ?? [], [])
  assert.equal(elements.some(element => element.type === 'polygon'), true)
  assert.equal(new Set(elements.filter(element => element.type === 'polygon').map(element => element.style.fill)).size > 4, true)
  assert.deepEqual(elements.filter(element => element.type === 'polygon').map(element => element.style.fill), commands.filter(command => command.kind === 'polygon').map(command => command.fill))
  const description = describeChartConfiguration(chart, { id: chart.id, fields: { 0: 'Cause', 1: 'Count' } })
  assert.deepEqual(description.pareto.cumulativeLineStyle.color, authored)
  assert.equal(description.series[1].type, 'column')
  assert.equal(Object.keys(description.series).length, 1)
  description.pareto.cumulativeLineStyle.color.start.x = 0
  assert.equal(chart.options.pareto.cumulativeLineStyle.color.start.x, -0.2)
  assert.deepEqual(chart.options.pareto.barStyle, before.options.pareto.barStyle)
  for (const field of ['authored', 'colorStops', 'bounds']) assert.equal(JSON.stringify(chart.options).includes(`"${field}"`), false)
})

test('materialized point indexes share aggregation/TopN preparation and keep positive curve widths when marker borders are zero', () => {
  const data = source(['low', 'high', 'high', 'zero', 'negative'], [2, 10, -1, 0, -9])
  const chart = model({ pareto: { includeZeroValues: true, cumulativeLineStyle: { color: 'red', width: 4,
    dataPoints: { 1: { color: 'blue', size: 14, shape: 'diamond', border: { width: 0 } } } } } })
  const derived = deriveChartData(data, chart), layout = layoutDerivedChartData(derived, chart), cumulative = chartParetoCumulativeStyle(chart, palette)
  assert.deepEqual(layout.categories, ['high', 'low', 'zero'])
  assert.equal(layout.series.at(-1).points[1].category, 'low')
  assert.equal(chartParetoCumulativePointStyle(cumulative, 1).point.size, 14)
  assert.equal(chartParetoCumulativeMarkerStyle(cumulative, 1).border.width, 0)
  const commands = drawChartData(data, chart)
  assert.equal(commands.some(command => command.kind === 'line' && command.width === 4 && command.color === normalizeChartColor('blue')), true)
  assert.equal(commands.some(command => command.kind === 'polygon' && command.fill === normalizeChartColor('blue')), true)
  const reordered = source(['high', 'low', 'negative', 'high', 'zero'], [10, 2, -9, -1, 0])
  assert.deepEqual(drawChartData(reordered, chart), commands)
  assert.deepEqual(optionForData(data, chart).series ?? [], [])
  const aggregated = source(['A', 'A', 'B', 'C'], [6, -1, 4, 2]), original = structuredClone(aggregated)
  const limited = model({ aggregation: { aggregate: true, topN: { count: 1, includeOthers: true, othersLabel: '其他' } }, pareto: { includeZeroValues: true } })
  assert.deepEqual(deriveChartData(aggregated, limited).data.categories, ['其他', 'A'])
  assert.deepEqual(chartParetoSettingsCategories(deriveChartData(aggregated, limited)), ['其他', 'A'])
  assert.deepEqual(aggregated, original)
  const nullable = source(['missing', 'A'], [null, 3]), zeroed = model({ invalidValueStrategy: 'zero', pareto: { includeZeroValues: true } })
  assert.deepEqual(chartParetoSettingsCategories(deriveChartData(nullable, zeroed)), ['A', 'missing'])
  assert.deepEqual(deriveChartData(nullable, zeroed).data.categories, ['A', 'missing'])
})

test('cumulative markers consume source alpha and opacity once and keep per-point outlines independent of the curve', () => {
  const chart = model({ pareto: { cumulativeLineStyle: { color: 'rgba(255, 0, 0, 0.8)', opacity: 0.5, width: 8,
    dataPoints: { 0: { fillOpacity: 0.25, border: { width: 3, opacity: 0.75 } }, 1: { shape: 'rect', border: { width: 0 } } } } } })
  const style = chartParetoCumulativeStyle(chart, palette), marker = chartParetoCumulativeMarkerStyle(style, 0), commands = []
  chartSymbol(commands, { x: 20, y: 20 }, marker.point.size, chartColorOpacity(marker.color, marker.fillOpacity), marker.point.shape, marker.border)
  commands.filter(command => command.kind === 'polygon').forEach(command => near(chartColorChannels(command.fill)[3], 0.2))
  const outlines = commands.filter(command => command.kind === 'line')
  assert.equal(outlines.length > 0, true)
  outlines.forEach(command => { assert.equal(command.width, 3); near(chartColorChannels(command.color)[3], 0.6) })
  for (const index of [undefined, 1]) {
    const current = chartParetoCumulativeMarkerStyle(style, index), local = []
    chartSymbol(local, { x: 20, y: 20 }, current.point.size, chartColorOpacity(current.color, current.fillOpacity), current.point.shape, current.border)
    assert.equal(local.some(command => command.kind === 'line'), false)
    local.filter(command => command.kind === 'rect').forEach(command => assert.equal(command.stroke, undefined))
  }
  const painted = drawChartData(source(), chart)
  assert.equal(painted.some(command => command.kind === 'line' && command.width === 8 && Math.abs(chartColorChannels(command.color)[3] - 0.4) < 1e-7), true)
  const explicit = chartParetoCumulativeMarkerStyle(chartParetoCumulativeStyle(model({ pareto: { cumulativeLineStyle: { color: gradient(), opacity: 0.5 } } }), palette))
  assert.equal(typeof explicit.color, 'string')
  near(chartColorChannels(chartColorOpacity(explicit.color, explicit.fillOpacity))[3], 0.4)
})

test('real setting controls edit detached drafts, restore line defaults, and expose only consumed bar marker fields', () => {
  const chart = model({ allSeriesStyle: { color: 'purple' }, pareto: { cumulativeLineStyle: { width: 4 } } }), before = structuredClone(chart)
  const draft = createChartSettingsDraft(chart), config = cloneChartDraftConfig(draft.config), line = chartParetoLineControls(), bar = chartParetoBarControls(source().series[0])
  const content = line.find(control => control.key.endsWith('.label.contentType'))
  assert.equal(content.visible(config), false)
  line.find(control => control.key.endsWith('.label.visible')).write(config, true)
  assert.equal(content.visible(config), true)
  const width = line.find(control => control.key === 'pareto.cumulativeLineStyle.width')
  width.write(config, 0)
  assert.equal(config.options.pareto.cumulativeLineStyle.width, 4)
  width.write(config, undefined)
  assert.equal(chartParetoCumulativeStyle(config, palette).border.width, 2)
  const barColor = bar.find(control => control.kind === 'chart-color'), color = gradient()
  assert.equal(barColor.read(config), undefined)
  assert.equal(barColor.fallback(config), 'purple')
  barColor.write(config, color)
  color.end.x = 0.5
  assert.equal(config.options.pareto.barStyle.color.end.x, 1.2)
  assert.equal(line.find(control => control.key.endsWith('.label.position')).choices.some(choice => choice.value === 'outside'), false)
  assert.equal(bar.find(control => control.key.endsWith('.label.position')).choices.some(choice => choice.value === 'outside'), true)
  const point = chartParetoDataPointControls('cumulativeLineStyle', 1), barPoint = chartParetoDataPointControls('barStyle', 1, source().series[0])
  assert.equal(point.some(control => control.key.endsWith('.size')), true)
  assert.equal(barPoint.some(control => control.key.endsWith('.size') || control.key.endsWith('.shape')), false)
  const withoutSource = createChartSettingsDraft(model({ series: { 1: { color: 'orange' } } })).config
  const inherited = chartParetoDataPointControls('barStyle', 1).find(control => control.kind === 'chart-color')
  assert.equal(inherited.fallback(withoutSource), 'orange')
  point.find(control => control.kind === 'chart-color').write(config, gradient())
  assert.equal(config.options.pareto.cumulativeLineStyle.dataPoints[1].fillPattern, null)
  const result = chartConfigFromDraft({ ...draft, config })
  assert.deepEqual(result.options.pareto.cumulativeLineStyle.dataPoints[1].color, gradient())
  assert.deepEqual(chart, before)
})
