import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartDateAxis, chartDateSeriesOrder } from '../src/features/charts/chart-date-axis.ts'
import { chartDateAxisCommands } from '../src/features/charts/chart-date-axis-draw.ts'
import { chartAxisPointerCommands, chartAxisPointerGeometry } from '../src/features/charts/chart-axis-pointer.ts'
import { chartAxisControls, chartDateAxisControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartSettingsDateAxisError } from '../src/features/charts/chart-settings-data.ts'
import { cloneChartDraftConfig } from '../src/features/charts/chart-settings-draft.ts'
import { validateChartAxis } from '../src/features/charts/chart-option-validation.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { normalizeChartColor } from '../src/features/charts/chart-color.ts'

// P22 implementation drafts only. No runner, native renderer, browser, benchmark or Office sample has run.
const JAN = 45292, JAN_31 = 45322, FEB = 45323, FEB_29 = 45351, MAR = 45352, MAR_31 = 45382, APR = 45383, APR_30 = 45412
const month = (year, number) => year * 12 + number - 1
function model(type = 'line', patch = {}) {
  return { id: 'date-axis-draft', sheetId: 'sheet', type, revision: 0, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 4, endCol: 1 } },
    position: { x: 0, y: 0 }, size: { width: 520, height: 360 }, legend: { visible: false, position: 'bottom' },
    xAxis: { scaleType: 'linear', baseTimeUnit: 'months', majorTimeUnit: 'months', majorUnit: 1 },
    yAxis: { scaleType: 'linear', min: 0, max: 100 }, options: { mapping: { useDateAxis: true } }, ...patch }
}
const source = () => ({ categories: ['March row', 'January row', 'February row', 'January duplicate'], categoryValues: [MAR_31, JAN_31, FEB_29, JAN_31],
  series: [{ field: 1, name: 'Units', values: [30, 10, 20, 11] }] })
function geometry(data, chart, computed = {}) {
  const derived = deriveChartData(data, chart, { preparedData: data, ...computed })
  return { derived, layout: layoutDerivedChartData(derived, chart) }
}
const elements = option => (option.graphic?.elements ?? []).flatMap(function flatten(element) { return element.children ? element.children.flatMap(flatten) : [element] })
const dateTexts = option => elements(option).filter(element => element.type === 'text' && /^\d{4}-\d{2}-\d{2}$/.test(element.style.text))
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} ≈ ${expected}`)

test('stable chronological ordering retains duplicate rows, raw dates and per-source marker styles in both paths', () => {
  const data = source(), chart = model('line', { options: { mapping: { useDateAxis: true }, series: { 1: { point: { shape: 'circle' },
    dataPoints: { 1: { color: 'red', point: { size: 9 } }, 3: { color: 'blue', point: { size: 13 } } } } } } })
  const before = structuredClone({ data, chart }), { derived, layout } = geometry(data, chart)
  assert.deepEqual(layout.dateOrder, [1, 3, 2, 0])
  assert.deepEqual(layout.series[0].points.map(point => [point.categoryIndex, point.sequenceIndex, point.xValue, point.value]),
    [[1, 0, JAN_31, 10], [3, 1, JAN_31, 11], [2, 2, FEB_29, 20], [0, 3, MAR_31, 30]])
  assert.equal(layout.series[0].points[0].x, layout.series[0].points[1].x)
  near(layout.series[0].points[2].x - layout.series[0].points[0].x, layout.series[0].points[3].x - layout.series[0].points[2].x)
  assert.deepEqual(derived.data.categories, data.categories)
  assert.deepEqual(derived.data.categoryDateValues, data.categoryValues)
  const native = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: data }).series[0].data
  assert.deepEqual(native.map(point => point.id), [1, 3, 2, 0].map(index => `${chart.id}:series:0:point:${index}`))
  assert.deepEqual(native.slice(0, 2).map(point => [point.name, point.symbolSize, point.itemStyle.color]),
    [['January row', 9, normalizeChartColor('red')], ['January duplicate', 13, normalizeChartColor('blue')]])
  assert.deepEqual({ data, chart }, before)
})

test('a missing Y occupies its chronological slot so break and link use date sequence instead of original row adjacency', () => {
  const data = { categories: ['March', 'January', 'Missing February', 'April'], categoryValues: [MAR_31, JAN_31, FEB_29, APR_30],
    series: [{ field: 1, name: 'Units', values: [30, 10, null, 40] }] }
  for (const strategy of ['break', 'link']) {
    const chart = model('line', { options: { mapping: { useDateAxis: true }, invalidValueStrategy: strategy, series: { 1: { color: '#dc2626', point: { shape: 'none' } } } } })
    const { derived, layout } = geometry(data, chart), points = layout.series[0].points
    assert.deepEqual(points.map(point => [point.categoryIndex, point.sequenceIndex]), [[1, 0], [0, 2], [3, 3]])
    const lines = drawDerivedChartData(derived, chart).filter(command => command.kind === 'line' && command.color === normalizeChartColor('#dc2626'))
    assert.equal(lines.length, strategy === 'break' ? 1 : 2)
    assert.deepEqual([lines.at(-1).x1, lines.at(-1).y1, lines.at(-1).x2, lines.at(-1).y2], [points[1].x, points[1].y, points[2].x, points[2].y])
    const native = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: data }).series[0]
    assert.equal(native.connectNulls, strategy === 'link')
    assert.equal(native.data.length, 4); assert.equal(native.data[1], null)
    assert.deepEqual(native.data.filter(Boolean).map(point => point.id), [1, 0, 3].map(index => `${chart.id}:series:0:point:${index}`))
  }
})

test('native contents and common chrome use identical monthly extents, first-day tick coordinates and reversed anchors', () => {
  const data = source()
  for (const reverse of [false, true]) {
    const chart = model('line', { xAxis: { scaleType: 'linear', baseTimeUnit: 'months', majorUnit: 1, reverse, label: { family: 'Georgia', fontSize: 16, color: 'purple' } } })
    const { derived, layout } = geometry(data, chart), option = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: data }), labels = dateTexts(option)
    assert.deepEqual([option.xAxis.min, option.xAxis.max], [layout.xMin, layout.xMax])
    assert.equal(option.xAxis.type, 'value'); assert.equal(option.xAxis.axisLabel.show, false)
    assert.equal(option.xAxis.axisTick.show, false); assert.equal(option.xAxis.splitLine.show, false)
    assert.equal(option.axisPointer, undefined); assert.equal(option.tooltip.axisPointer, undefined)
    assert.equal(option.officeChartPointer.mode, 'graphic')
    assert.deepEqual(labels.map(element => element.style.text), ['2024-01-01', '2024-02-01', '2024-03-01'])
    const commands = drawDerivedChartData(derived, chart).filter(command => command.kind === 'text' && /^2024-\d{2}-01$/.test(command.text))
    assert.deepEqual(labels.map(element => [element.x, element.y, element.style.font, normalizeChartColor(element.style.fill)]),
      commands.map(command => [command.x, command.y, command.font, command.color]))
    for (const [index, label] of labels.entries()) near(label.x, layout.plot.x + (reverse ? 1 - (index + 0.5) / 3 : (index + 0.5) / 3) * layout.plot.width)
    assert.equal(elements(option).filter(element => String(element.id).startsWith(chart.id + ':date-axis')).every(element => element.z < 0), true)
  }
})

test('date grid defaults stay hidden and explicit grid, ticks, label visibility and frame clipping remain independent', () => {
  const chart = model(), { layout } = geometry(source(), chart)
  const ordinary = chartDateAxisCommands(chart, layout, 'black')
  assert.equal(ordinary.some(command => command.kind === 'line'), false)
  assert.equal(ordinary.filter(command => command.kind === 'text').length, 3)
  chart.xAxis.gridLine = { visible: true, color: 'green', width: 2 }
  chart.xAxis.tick = { visible: true, length: 6, position: 'inside', color: 'orange' }
  chart.xAxis.label = { visible: false }
  const explicit = chartDateAxisCommands(chart, layout, 'black')
  assert.equal(explicit.some(command => command.kind === 'text'), false)
  const grids = explicit.filter(command => command.kind === 'line' && command.color === normalizeChartColor('green'))
  const ticks = explicit.filter(command => command.kind === 'line' && command.color === normalizeChartColor('orange'))
  assert.equal(grids.length, 3); assert.equal(ticks.length, 3)
  assert.equal(grids.every(command => command.width === 2 && command.y1 === layout.plot.y && command.y2 === layout.plot.y + layout.plot.height), true)
  assert.equal(grids.every(command => command.clip === layout.plot), true)
  assert.equal(ticks.every(command => command.y2 === command.y1 - 6 && command.clip === layout.frame), true)
})

test('horizontal Bar and projected 3D charts preserve calendar source identity and render the same first-day labels', () => {
  for (const type of ['bar-horizontal', 'column3D', 'bar3D', 'line3D', 'area3D']) {
    const chart = model(type), { derived, layout } = geometry(source(), chart), option = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: source() })
    assert.deepEqual(layout.dateOrder, [1, 3, 2, 0], type)
    assert.deepEqual(layout.series[0].points.map(point => [point.categoryIndex, point.sequenceIndex, point.xValue]),
      [[1, 0, JAN_31], [3, 1, JAN_31], [2, 2, FEB_29], [0, 3, MAR_31]], type)
    assert.equal(layout.series[0].points.every(point => Number.isFinite(point.x) && Number.isFinite(point.y)), true, type)
    assert.equal(option.series, undefined, type)
    const labels = drawDerivedChartData(derived, chart).filter(command => command.kind === 'text' && /^2024-\d{2}-01$/.test(command.text))
    assert.deepEqual(labels.map(command => command.text), ['2024-01-01', '2024-02-01', '2024-03-01'], type)
    if (type === 'bar-horizontal') {
      assert.equal(layout.horizontal, true)
      assert.equal(layout.series[0].points[0].y, layout.series[0].points[1].y)
      assert.equal(labels.every(command => command.x === layout.plot.x - 8 && command.align === 'right'), true)
    }
  }
})

test('Bubble keeps sorted date tuples, size values and original grouped row identifiers', () => {
  const data = source(); data.series[0].sizeValues = [3, 1, 2, 4]
  const chart = model('bubble'), { layout } = geometry(data, chart), native = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: data })
  assert.deepEqual(native.series[0].data.map(point => point.value),
    [[month(2024, 1), 10, 1], [month(2024, 1), 11, 4], [month(2024, 2), 20, 2], [month(2024, 3), 30, 3]])
  assert.deepEqual(layout.series[0].points.map(point => point.size), [1, 4, 2, 3])
  assert.deepEqual(native.series[0].data.map(point => point.name), ['January row', 'January duplicate', 'February row', 'March row'])
})

test('Boxplot distribution names and Candlestick dates use calendar positions without replacing raw samples or summaries', () => {
  const boxes = { categories: ['Sample 1', 'Sample 2', 'Sample 3'], series: [
    { field: 1, name: '2024-03-31', values: [10, 20, 30] }, { field: 2, name: '2024-01-31', values: [3, 4, 5] }, { field: 3, name: '2024-02-29', values: [6, 7, 8] }] }
  const boxChart = model('boxplot'), before = structuredClone(boxes), result = geometry(boxes, boxChart)
  assert.deepEqual(result.layout.dateOrder, [1, 2, 0])
  assert.deepEqual(result.derived.data.categoryDateValues, [MAR_31, JAN_31, FEB_29])
  assert.deepEqual(result.derived.boxes.map(box => [box.field, box.seriesPosition, box.sourceName, box.statistics.median]),
    [[1, 0, '2024-03-31', 20], [2, 1, '2024-01-31', 4], [3, 2, '2024-02-29', 7]])
  assert.ok(result.layout.series[1].points[0].x < result.layout.series[2].points[0].x && result.layout.series[2].points[0].x < result.layout.series[0].points[0].x)
  assert.equal(chartEChartsOption({ sheets: [] }, boxChart, {}, { preparedData: boxes }).series, undefined)
  assert.deepEqual(boxes, before)
  const candles = { categories: ['March', 'January', 'February'], categoryValues: [MAR_31, JAN_31, FEB_29],
    series: [{ field: 1, name: 'Close', values: [30, 10, 20] }],
    candles: [{ open: 28, close: 30, low: 27, high: 31 }, { open: 8, close: 10, low: 7, high: 11 }, { open: 18, close: 20, low: 17, high: 21 }] }
  const candleChart = model('candlestick'), candleResult = geometry(candles, candleChart)
  assert.deepEqual(candleResult.layout.dateOrder, [1, 2, 0])
  assert.deepEqual(candleResult.derived.data.candles, candles.candles)
  assert.deepEqual(dateTexts(chartEChartsOption({ sheets: [] }, candleChart, {}, { preparedData: candles })).map(element => element.style.text),
    ['2024-01-01', '2024-02-01', '2024-03-01'])
})

test('Pareto cumulative values retain their weight sequence when calendar coordinates run in reverse chronological order', () => {
  const data = { categories: ['Earlier', 'Later', 'Middle'], categoryValues: [JAN_31, MAR_31, FEB_29], series: [{ field: 1, name: 'Weight', values: [2, 9, 3] }] }
  const chart = model('pareto'), { derived, layout } = geometry(data, chart), axis = chartDateAxis(derived.data, chart)
  assert.deepEqual(axis.order, [2, 1, 0])
  assert.deepEqual(chartDateSeriesOrder(axis, chart), [0, 1, 2])
  assert.deepEqual(layout.dateOrder, [0, 1, 2])
  assert.deepEqual(derived.data.categories, ['Later', 'Middle', 'Earlier'])
  assert.deepEqual(layout.series[1].points.map(point => [point.categoryIndex, point.sequenceIndex, point.xValue, point.value]),
    derived.pareto.cumulativePercent.map((value, index) => [index, index, [MAR_31, FEB_29, JAN_31][index], value]))
  assert.ok(layout.series[1].points[0].x > layout.series[1].points[1].x && layout.series[1].points[1].x > layout.series[1].points[2].x)
  const option = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: data })
  assert.deepEqual(option.series[1].data.map(point => point.value), derived.pareto.cumulativePercent.map((value, index) => [month(2024, 3) - index, value]))
  assert.deepEqual(option.series[1].data.map(point => point.id), [0, 1, 2].map(index => `${chart.id}:pareto-cumulative:${index}`))
})

test('Histogram final bin centers supply real date identities and budget validation runs on derived coordinates', () => {
  const data = { categories: ['Raw 1', 'Raw 2'], series: [{ field: 1, name: 'Dates', values: [JAN_31, APR_30] }] }
  const histogram = { validSampleCount: 2, bins: [{ kind: 'regular', lowerBound: JAN, upperBound: FEB, count: 1 }, { kind: 'regular', lowerBound: FEB, upperBound: MAR, count: 1 }] }
  const result = geometry(data, model('histogram'), { histogram })
  assert.deepEqual(result.derived.data.categoryDateValues, [(JAN + FEB) / 2, (FEB + MAR) / 2])
  assert.deepEqual(result.layout.dateTicks.map(tick => tick.serial), [JAN, FEB])
  const dense = model('histogram', { xAxis: { scaleType: 'linear', majorTimeUnit: 'days', majorUnit: 0.001 } })
  assert.throws(() => geometry(data, dense, { histogram }), /超过10000/)
  // This draft proves final-layout validation, not pre-command-publication admission.
})

test('raw-serial trends and X error bounds expand monthly layout before tick generation and remain raw in derived results', () => {
  const data = { categories: ['January', 'February'], categoryValues: [JAN_31, FEB_29], series: [{ field: 1, name: 'Units', values: [10, 20], sizeValues: [1, 2] }] }
  const options = { mapping: { useDateAxis: true }, trendlines: [{ seriesIndex: 1, type: 'linear' }],
    errorBars: [{ seriesIndex: 1, direction: 'x', type: 'both', valueType: 'fixedVal', value: 1 }] }
  const trend = { points: [[JAN_31, 10], [APR_30, 40]] }, errors = [{ categoryIndex: 0, x: JAN_31, y: 10, minus: 31, plus: 60 }]
  const result = geometry(data, model('bubble', { options }), { trendlines: [{ configIndex: 0, seriesPosition: 0, statistics: trend }], errorBars: [{ configIndex: 0, seriesPosition: 0, intervals: errors }] })
  assert.deepEqual([result.layout.xMin, result.layout.xMax], [month(2023, 12) - 0.5, month(2024, 4) + 0.5])
  assert.deepEqual([result.layout.dateSourceMin, result.layout.dateSourceMax], [45291, APR_30])
  assert.deepEqual(result.layout.dateTicks.map(tick => tick.serial), [45261, JAN, FEB, MAR, APR])
  assert.deepEqual(result.derived.trends[0].statistics, trend)
  assert.deepEqual(result.derived.errors[0].intervals, errors)
  const narrow = { ...data, categoryValues: [JAN_31, JAN_31 + 1] }
  const dense = model('bubble', { xAxis: { scaleType: 'linear', majorUnit: 0.001 }, options })
  assert.doesNotThrow(() => geometry(narrow, model('bubble', { xAxis: dense.xAxis, options: { mapping: { useDateAxis: true } } })))
  assert.throws(() => geometry(narrow, dense, { trendlines: [{ configIndex: 0, seriesPosition: 0, statistics: trend }], errorBars: [{ configIndex: 0, seriesPosition: 0, intervals: errors }] }), /超过10000/)
})

test('calendar pointers show original dates and raw spans rather than the first-day labels or month ordinals', () => {
  const data = source(); data.categoryValues[3] = JAN
  const chart = model('line', { position: { x: 90, y: 70 } }), before = structuredClone(chart), { derived, layout } = geometry(data, chart)
  const pointer = chartAxisPointerGeometry(derived.data, chart, layout), point = layout.series[0].points[0]
  for (const offset of [-1, 0, 1]) {
    const commands = chartAxisPointerCommands(pointer, { x: point.x + offset, y: layout.plot.y + layout.plot.height / 2 })
    assert.equal(commands.find(command => command.kind === 'text').text, '2024-01-31')
  }
  assert.equal(chartDateAxisCommands(chart, layout, 'black').find(command => command.kind === 'text').text, '2024-01-01')
  assert.deepEqual(pointer.categories.filter(category => category.position === layout.categoryPositions[1]).map(category => category.index), [1, 3])
  assert.deepEqual(chartAxisPointerCommands(pointer), [])
  assert.deepEqual(chart, before)
})

test('date unit fields reset independently, preserve unrelated axes and inherit effective major units without author writes', () => {
  const original = model('line', { xAxis: { scaleType: 'linear', min: JAN, max: APR_30, reverse: true, format: 'mmm yyyy', baseTimeUnit: 'months', majorTimeUnit: 'years', majorUnit: 2 } }), before = structuredClone(original)
  const draft = cloneChartDraftConfig(original), controls = chartAxisControls('xAxis', true, true), fields = ['baseTimeUnit', 'majorTimeUnit', 'majorUnit']
  const units = fields.map(field => controls.find(control => control.key === 'xAxis.' + field))
  assert.deepEqual(units.map(control => control.label), ['日期基本单位', '日期主刻度单位', '日期主单位间隔'])
  assert.equal(units.every(control => control.visible(draft) && control.canReset(draft)), true)
  units[1].reset(draft)
  assert.equal(units[1].read(draft), 'months'); assert.equal(Object.hasOwn(draft.xAxis, 'majorTimeUnit'), true)
  assert.equal(draft.xAxis.majorTimeUnit, undefined)
  assert.equal(draft.xAxis.baseTimeUnit, 'months'); assert.equal(draft.xAxis.majorUnit, 2)
  units[0].reset(draft)
  assert.equal(units[0].read(draft), undefined); assert.equal(units[0].fallback, 'days')
  assert.equal(units[1].read(draft), 'days'); assert.equal(draft.xAxis.majorUnit, 2)
  units[2].reset(draft)
  assert.deepEqual(draft.xAxis, { scaleType: 'linear', min: JAN, max: APR_30, reverse: true, format: 'mmm yyyy', baseTimeUnit: undefined, majorTimeUnit: undefined, majorUnit: undefined })
  assert.equal(units.every(control => !control.canReset(draft)), true)
  assert.deepEqual(draft.yAxis, original.yAxis); assert.deepEqual(original, before)
})

test('date off and toggle reset preserve calendar metadata while numeric axes accept fractional intervals and independent resets', () => {
  const chart = model(), dates = chartDateAxisControls()[0], controls = chartAxisControls('xAxis', true, true), units = controls.filter(control => ['xAxis.baseTimeUnit', 'xAxis.majorTimeUnit', 'xAxis.majorUnit'].includes(control.key)), axisBefore = structuredClone(chart.xAxis)
  dates.write(chart, false)
  assert.equal(chart.options.mapping.useDateAxis, false)
  assert.deepEqual(chart.xAxis, axisBefore)
  assert.equal(units.every(control => !control.visible(chart)), true)
  dates.reset(chart)
  assert.equal(Object.hasOwn(chart.options.mapping, 'useDateAxis'), true)
  assert.equal(chart.options.mapping.useDateAxis, undefined)
  assert.deepEqual(chart.xAxis, axisBefore); assert.equal(dates.canReset(chart), false)
  const numeric = model('scatter', { xAxis: { scaleType: 'linear', min: 0, max: 1, baseTimeUnit: 'months', majorTimeUnit: 'years', majorUnit: 0.25 }, options: { mapping: { useDateAxis: false } } })
  assert.doesNotThrow(() => validateChartAxis(numeric.xAxis, 'X'))
  assert.equal(chartDateAxis(source(), numeric), undefined)
  const data = { categories: ['A', 'B'], series: [{ field: 1, name: 'Y', values: [1, 2], xValues: [0.25, 0.75] }] }
  const option = chartEChartsOption({ sheets: [] }, numeric, {}, { preparedData: data })
  assert.deepEqual([option.xAxis.interval, option.xAxis.minInterval, option.xAxis.maxInterval], [0.25, 0.25, 0.25])
  const interval = chartAxisControls('xAxis').find(control => control.key === 'xAxis.majorUnit')
  assert.equal(interval.label, '主刻度间隔'); interval.reset(numeric)
  assert.deepEqual(numeric.xAxis, { scaleType: 'linear', min: 0, max: 1, baseTimeUnit: 'months', majorTimeUnit: 'years', majorUnit: undefined })
})

test('settings report the same active calendar interval and budget errors without changing the document draft', () => {
  const chart = model('line', { xAxis: { scaleType: 'linear', baseTimeUnit: 'months', majorUnit: 0.5 } }), before = structuredClone(chart)
  assert.match(chartSettingsDateAxisError(chart, { categories: [], series: [] }), /主单位/)
  assert.match(chartSettingsDateAxisError(model('line', { xAxis: { scaleType: 'linear', majorUnit: 0.001 } }), source()), /超过10000/)
  assert.equal(chartSettingsDateAxisError({ ...chart, options: { mapping: { useDateAxis: false } } }, source()), undefined)
  assert.deepEqual(chart, before)
})
