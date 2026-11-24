import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, updateChart } from '../src/features/charts/chart-engine.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { chartPlotFrame, layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { chartDateAxis, chartDateAxisLabel, chartDatePosition, chartDateSerial, prepareChartDateData, prepareChartDateSource } from '../src/features/charts/chart-date-axis.ts'
import { chartAxisTitleCommand, chartAxisTitleFont } from '../src/features/charts/chart-axis-title.ts'
import { chartAxisPointerCommands, chartAxisPointerGeometry, chartAxisPointerLocalPoint, supportsChartAxisPointer } from '../src/features/charts/chart-axis-pointer.ts'
import { chartAxisControls, chartAxisPointerControls, chartDateAxisControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartSettingsDateAxisError } from '../src/features/charts/chart-settings-data.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'

// P12 drafts for the unified verification phase. No test in this file has been executed in this phase.
function chart(type = 'line', patch = {}) {
  return { id: 'axes', sheetId: 'sheet', type, revision: 0, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 320 }, legend: { visible: false, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear', min: 0, max: 100 }, ...patch }
}
const ordinary = () => ({ categories: ['A', 'B', 'C'], series: [{ field: 1, name: 'Units', values: [10, 20, 30] }] })
function geometry(source, model) {
  const derived = deriveChartData(source, model, { preparedData: source })
  const layout = layoutDerivedChartData(derived, model)
  return { derived, layout, pointer: chartAxisPointerGeometry(derived.data, model, layout) }
}
function labels(commands) { return commands.filter(command => command.kind === 'text').map(command => command.text) }
function near(actual, expected) { assert.ok(Math.abs(actual - expected) < 1e-9, `${actual} ≈ ${expected}`) }

test('Excel serial 59/60/61 and 1904 dates retain their separate calendar identities', () => {
  assert.deepEqual(['1900-02-28', '1900-02-29', '1900-03-01'].map(value => chartDateSerial(value)), [59, 60, 61])
  assert.deepEqual([59, 60, 61].map(value => chartDateAxisLabel(value, {})), ['1900-02-28', '1900-02-29', '1900-03-01'])
  assert.equal(chartDateSerial('1904-01-01', true), 0)
  assert.equal(chartDateSerial('1900-02-29', true), null)
  assert.equal(chartDateAxisLabel(0, {}, true), '1904-01-01')
  assert.equal(chartDateSerial('60'), 60)
  for (const value of ['2024-02-30', '2024-13-01', '2024-01-01T25:00', '2024-01-01T00:00+24:00', '1/2/2024', '', null, true, NaN, Infinity, -1]) {
    assert.equal(chartDateSerial(value), null, String(value))
  }
})

test('real ISO UTC offsets skip the fictional leap day while an explicit fictional date keeps serial arithmetic', () => {
  near(chartDateSerial('1900-03-01T00:30+01:00'), 59 + 23.5 / 24)
  near(chartDateSerial('1900-02-28T23:30-01:00'), 61 + 0.5 / 24)
  near(chartDateSerial('1900-02-29T23:30-01:00'), 61 + 0.5 / 24)
  near(chartDateSerial('1904-01-01T02:30+02:00', true), 0.5 / 24)
  assert.equal(chartDateSerial('1904-01-01T00:00+01:00', true), null)
})

test('date positions reflect elapsed days, authored bounds and reversal instead of category ordinals', () => {
  const source = { ...ordinary(), categoryValues: [59, 60, 63] }
  const model = chart('line', { options: { mapping: { useDateAxis: true } }, xAxis: { scaleType: 'linear', min: 58, max: 64 } })
  const prepared = prepareChartDateData(source, model), dates = chartDateAxis(prepared, model)
  assert.deepEqual(prepared.categoryDateValues, [59, 60, 63])
  assert.deepEqual(prepared.series[0].xValues, [59, 60, 63])
  near(chartDatePosition(dates, model.xAxis, 0), 1 / 6)
  near(chartDatePosition(dates, model.xAxis, 1), 2 / 6)
  near(chartDatePosition(dates, model.xAxis, 2), 5 / 6)
  near(chartDatePosition(dates, { ...model.xAxis, reverse: true }, 1), 4 / 6)
  const { layout } = geometry(source, model)
  near(layout.series[0].points[2].x - layout.series[0].points[1].x, (layout.series[0].points[1].x - layout.series[0].points[0].x) * 3)
  near(layout.categoryBandwidth, layout.plot.width / 6)
})

test('missing dates omit their points and a numeric source needs at least one matching date/value pair', () => {
  const model = chart('line', { options: { mapping: { useDateAxis: true } } })
  const source = { ...ordinary(), categoryValues: [59, null, 61] }
  assert.deepEqual(geometry(source, model).layout.series[0].points.map(point => point.categoryIndex), [0, 2])
  assert.throws(() => prepareChartDateData({ categories: ['Valid', 'Missing'], categoryValues: [59, null],
    series: [{ field: 1, name: 'Value', values: [null, 9] }] }, model), /与数值对应/)
  assert.throws(() => prepareChartDateData({ ...source, categoryValues: [59, 'invalid', 61] }, model), /第 2 个分类/)
  assert.throws(() => prepareChartDateData(source, { ...model, xAxis: { scaleType: 'log' } }), /线性/)
  assert.throws(() => prepareChartDateData(source, { ...model, xAxis: { scaleType: 'linear', min: -1 } }), /边界/)
})

test('Pareto keeps raw date identity through weight sorting and rejects conflicting grouped dates before model mutation', () => {
  const model = chart('pareto', { options: { mapping: { useDateAxis: true } } })
  const source = { categories: ['Earlier', 'Middle', 'Later'], categoryValues: [59, 60, 63],
    series: [{ field: 1, name: 'Weight', values: [2, 9, 3] }] }
  const result = geometry(source, model).derived
  assert.deepEqual(result.data.categories, ['Middle', 'Later', 'Earlier'])
  assert.deepEqual(result.data.categoryDateValues, [60, 63, 59])
  assert.deepEqual(result.pareto.values, [9, 3, 2])
  const ambiguous = { categories: ['Same', 'Same'], categoryValues: [59, 60], series: [{ field: 1, name: 'Weight', values: [1, 2] }] }
  assert.throws(() => prepareChartDateSource(ambiguous, model), /同名分类/)
  assert.throws(() => geometry(ambiguous, model), /同名分类/)
})

test('Bubble ISO dates become real X coordinates during candidate validation and invalid dates leave the model unchanged', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = Object.fromEntries([['Date', 'Y', 'Size'], ['1904-01-01', 4, 1], ['1904-01-03', 8, 2]]
    .flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  workbook.date1904 = true
  const model = createChart(workbook, sheet.id, { ...chart('bubble'), valueFields: [1, 2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 } },
    options: { mapping: { useDateAxis: true, bubble: { xIndex: 0, yIndex: 1, sizeIndex: 2 } } } })
  const before = structuredClone(model)
  assert.throws(() => updateChart(workbook, model, { xAxis: { scaleType: 'log' } }), /线性/)
  assert.deepEqual(model, before)
  sheet.cells['1:0'].value = 'bad date'
  assert.throws(() => updateChart(workbook, model, { title: 'Rejected' }), /无法作为日期轴值/)
  assert.deepEqual(model, before)
})

test('axis title fonts inherit defined fields and physical start/end remain stable under axis reversal', () => {
  const model = chart('line', { options: { appearance: { font: { family: 'Courier New', fontSize: 24, bold: true, color: '#112233' } } } })
  const axis = { title: 'Axis', titleStyle: { fontSize: 18, bold: undefined, italic: true }, titleAlignment: 'start' }
  assert.deepEqual(chartAxisTitleFont(model, axis), { family: 'Courier New', fontSize: 18, bold: true, color: '#112233', italic: true })
  const start = { x: 20, y: 200 }, end = { x: 20, y: 30 }
  const command = chartAxisTitleCommand(model, axis, start, end, 'black', -90)
  assert.deepEqual([command.x, command.y, command.align, command.maxWidth], [20, 30, 'right', 170])
  assert.equal(command.font, 'italic 600 18px Courier New')
  const reversed = chartAxisTitleCommand(model, { ...axis, reverse: true }, start, end, 'black', -90)
  assert.deepEqual([reversed.x, reversed.y, reversed.align], [20, 30, 'right'])
  const centered = chartAxisTitleCommand(model, { ...axis, titleAlignment: 'center' }, start, end, 'black', -90)
  assert.deepEqual([centered.x, centered.y, centered.align], [20, 115, 'center'])
})

test('title size changes the content margins and Canvas/native/3D paths consume the authored font and alignment', () => {
  const source = ordinary(), model = chart('line', { title: 'Heading', subtitle: 'Subheading',
    options: { appearance: { font: { family: 'Georgia', fontSize: 48, bold: true } } },
    yAxis: { scaleType: 'linear', title: 'Y title', titleStyle: { fontSize: 32, italic: true }, titleAlignment: 'end' } })
  const frame = chartPlotFrame(model, source.series), smaller = chartPlotFrame({ ...model, yAxis: { ...model.yAxis, titleStyle: { fontSize: 11 } } }, source.series)
  near(frame.plot.x - smaller.plot.x, 21)
  assert.ok(frame.plot.y >= 12 + 62 + 55)
  const { derived } = geometry(source, model), scene = drawDerivedChartData(derived, model)
  const subtitle = scene.find(command => command.kind === 'text' && command.text === 'Subheading')
  assert.ok(subtitle.y + 24 < frame.plot.y)
  assert.match(subtitle.font, /48px Georgia/)
  const native = chartEChartsOption({ sheets: [] }, model, {}, { preparedData: source })
  assert.equal(native.yAxis.name, undefined)
  assert.deepEqual([native.yAxis.nameTextStyle.fontSize, native.yAxis.nameTextStyle.fontFamily, native.yAxis.nameTextStyle.fontWeight, native.yAxis.nameTextStyle.fontStyle], [32, 'Georgia', 'bold', 'italic'])
  const title = native.graphic.elements.map(element => element.type === 'text' ? element : element.children?.[0]).find(element => element?.style.text === 'Y title')
  assert.equal(title.y, frame.plot.y + frame.plot.height)
  assert.match(title.style.font, /32px Georgia/)
  const threeD = chart('column3D', { ...model, type: 'column3D', xAxis: { scaleType: 'linear', title: 'Projected title', titleStyle: { family: 'Georgia', fontSize: 22 }, titleAlignment: 'start' } })
  assert.ok(drawDerivedChartData(geometry(source, threeD).derived, threeD).some(command => command.kind === 'text' && command.text === 'Projected title' && command.font.includes('22px Georgia')))
})

test('categorical pointers consume all four style parameters, secondary formats and cached geometry without scanning values', () => {
  const source = ordinary()
  source.series.push({ field: 2, name: 'Ratio', values: [0.2, 0.4, 0.6] })
  const model = chart('line', { yAxis: { scaleType: 'linear', min: 0, max: 100, format: '$0.0' }, options: {
    axisPointer: { indicatorLineColor: 'blue', indicatorLineType: 'dashed', indicatorLabelColor: 'black', indicatorLabelTextColor: 'white' },
    series: { 2: { rightAxis: true } }, rightYAxis: { scaleType: 'linear', min: 0, max: 1, format: '0.00' } } })
  const { layout, pointer } = geometry(source, model)
  Object.defineProperty(pointer.data.series[0], 'values', { get() { throw new Error('Hover scanned source values') } })
  const commands = chartAxisPointerCommands(pointer, { x: layout.plot.x + layout.plot.width / 6, y: layout.plot.y + layout.plot.height / 2 })
  assert.deepEqual(labels(commands), ['A', '$50.0', '0.50'])
  assert.ok(commands.filter(command => command.kind === 'line').length > 2)
  assert.ok(commands.filter(command => command.kind === 'line').every(command => command.color === 'rgba(0, 0, 255, 1)'))
  assert.ok(commands.filter(command => command.kind === 'rect').every(command => command.fill === 'rgba(0, 0, 0, 1)'))
  assert.ok(commands.filter(command => command.kind === 'text').every(command => command.color === 'rgba(255, 255, 255, 1)'))
  for (const rect of commands.filter(command => command.kind === 'rect')) {
    assert.ok(rect.x >= layout.frame.x && rect.y >= layout.frame.y)
    assert.ok(rect.x + rect.width <= layout.frame.x + layout.frame.width && rect.y + rect.height <= layout.frame.y + layout.frame.height)
  }
  assert.deepEqual(chartAxisPointerCommands(pointer), [])
  assert.deepEqual(chartAxisPointerCommands(pointer, { x: -1, y: -1 }), [])
  assert.deepEqual(chartAxisPointerCommands(pointer, { x: NaN, y: 0 }), [])
})

test('date pointers snap to real dates including serial 60 and numeric pointers respect logarithmic reversal', () => {
  const dated = chart('line', { options: { mapping: { useDateAxis: true } }, xAxis: { scaleType: 'linear', min: 58, max: 64, reverse: true } })
  const result = geometry({ ...ordinary(), categoryValues: [59, 60, 63] }, dated)
  const point = { x: result.layout.plot.x + result.layout.categoryPositions[1] * result.layout.plot.width, y: result.layout.plot.y + result.layout.plot.height / 2 }
  assert.equal(labels(chartAxisPointerCommands(result.pointer, point))[0], '1900-02-29')
  const scatter = chart('scatter', { xAxis: { scaleType: 'log', min: 1, max: 100, reverse: true }, yAxis: { scaleType: 'log', min: 1, max: 100 } })
  const numeric = geometry({ ...ordinary(), series: [{ field: 1, name: 'Pairs', values: [1, 10, 100], xValues: [1, 10, 100] }] }, scatter)
  assert.deepEqual(labels(chartAxisPointerCommands(numeric.pointer, { x: numeric.layout.plot.x + numeric.layout.plot.width / 2, y: numeric.layout.plot.y + numeric.layout.plot.height / 2 })), ['10', '10'])
})

test('horizontal and Heatmap pointers identify their actual categories and reversed category rows', () => {
  const horizontal = geometry(ordinary(), chart('bar-horizontal', { xAxis: { scaleType: 'linear', reverse: true } }))
  assert.deepEqual(labels(chartAxisPointerCommands(horizontal.pointer, { x: horizontal.layout.plot.x + horizontal.layout.plot.width / 2, y: horizontal.layout.plot.y + horizontal.layout.plot.height / 6 })), ['C', '50'])
  const heatmap = ordinary(); heatmap.series.push({ field: 2, name: 'Other row', values: [4, 5, 6] })
  const result = geometry(heatmap, chart('heatmap', { yAxis: { scaleType: 'linear', reverse: true } }))
  assert.deepEqual(labels(chartAxisPointerCommands(result.pointer, { x: result.layout.plot.x + result.layout.plot.width / 6, y: result.layout.plot.y + result.layout.plot.height / 4 })), ['Units', 'A'])
})

test('Radar pointers follow the radial geometry and stop beyond the polygon or circle', () => {
  for (const shape of ['polygon', 'circle']) {
    const model = chart('radar', { options: { radar: { shape } }, xAxis: { scaleType: 'linear', reverse: true }, yAxis: { scaleType: 'linear', min: 0, max: 100, reverse: true } })
    const result = geometry(ordinary(), model), radar = result.pointer.radar, vertex = radar.vertices[1]
    const point = { x: radar.center.x + (vertex.x - radar.center.x) / 4, y: radar.center.y + (vertex.y - radar.center.y) / 4 }
    assert.deepEqual(labels(chartAxisPointerCommands(result.pointer, point)), ['B', '75'])
    assert.deepEqual(chartAxisPointerCommands(result.pointer, { x: result.layout.plot.x, y: result.layout.plot.y }), [])
  }
})

test('native pointer and tooltip models receive the same authorship and specialized Gauge tooltip stays offscreen', () => {
  const source = ordinary(), model = chart('line', { options: { axisPointer: { indicatorLineColor: 'red', indicatorLineType: 'dotted', indicatorLabelColor: 'black', indicatorLabelTextColor: 'white' } } })
  const option = chartEChartsOption({ sheets: [] }, model, {}, { preparedData: source })
  assert.equal(option.officeChartPointer.mode, 'native')
  assert.equal(option.tooltip.showContent, false)
  assert.equal(option.tooltip.triggerOn, 'none')
  assert.deepEqual(option.tooltip.axisPointer.label, option.axisPointer.label)
  assert.deepEqual(option.tooltip.axisPointer.lineStyle, option.axisPointer.lineStyle)
  assert.deepEqual(option.tooltip.axisPointer.crossStyle, option.axisPointer.lineStyle)
  assert.equal(option.axisPointer.label.backgroundColor, 'rgba(0, 0, 0, 1)')
  const gauge = chartEChartsOption({ sheets: [] }, chart('gauge'), {}, { preparedData: source })
  assert.equal(gauge.officeChartPointer, undefined)
  assert.equal(gauge.tooltip.showContent, false)
  assert.equal(gauge.tooltip.triggerOn, 'none')
  for (const type of ['column3D', 'pie', 'candlestick', 'histogram']) assert.equal(supportsChartAxisPointer(chart(type)), false, type)
})

test('date-axis form controls preserve explicit off intent and share real candidate validation', () => {
  const model = chart('line', { xAxis: { scaleType: 'log', logBase: 2 } }), toggle = chartDateAxisControls()[0]
  toggle.write(model, true)
  assert.deepEqual(model.xAxis, { scaleType: 'linear', logBase: undefined })
  toggle.write(model, false)
  assert.equal(model.options.mapping.useDateAxis, false)
  const controls = chartAxisControls('xAxis', true, true)
  assert.ok(controls.some(control => control.key === 'xAxis.min'))
  assert.ok(controls.some(control => control.key === 'xAxis.titleStyle.family'))
  assert.ok(controls.some(control => control.key === 'xAxis.titleAlignment'))
  assert.ok(!controls.some(control => ['xAxis.scaleType', 'xAxis.logBase'].includes(control.key)))
  assert.deepEqual(chartAxisPointerControls().map(control => control.key), ['axisPointer.indicatorLineColor', 'axisPointer.indicatorLineType', 'axisPointer.indicatorLabelColor', 'axisPointer.indicatorLabelTextColor'])
  const box = chart('boxplot', { options: { mapping: { useDateAxis: true } } })
  const source = { categories: ['Sample 1', 'Sample 2'], series: [{ field: 1, name: '1904-01-01', values: [null, 9] }, { field: 2, name: '1904-01-03', values: [8, null] }] }
  assert.equal(chartSettingsDateAxisError(box, source, true), undefined)
  assert.match(chartSettingsDateAxisError({ ...box, xAxis: { min: -1 } }, source, true), /边界/)
  assert.match(chartSettingsDateAxisError(box, { ...source, series: [{ field: 1, name: 'Not a date', values: [1, 2] }] }, true), /无法作为日期轴值/)
})

test('CSS pointer coordinates apply zoom, scroll and object origin once and leave DPR to the surface', () => {
  assert.deepEqual(chartAxisPointerLocalPoint({ x: 220, y: 170 }, { x: 100, y: 50 }, { zoom: 200, scrollLeft: 40, scrollTop: 30 }), { x: 30, y: 50 })
  assert.deepEqual(chartAxisPointerLocalPoint({ x: 90, y: 35 }, { x: 100, y: 50 }, { zoom: 50, scrollLeft: 10, scrollTop: 15 }), { x: 100, y: 50 })
})
