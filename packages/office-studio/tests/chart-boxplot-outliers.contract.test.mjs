import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData, updateChart } from '../src/features/charts/chart-engine.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { chartBoxplotGeometry, chartBoxplotOutliers, chartBoxplotOutlierStyle, chartBoxplotStyle } from '../src/features/charts/chart-boxplot.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { normalizeChartColor } from '../src/features/charts/chart-color.ts'
import { chartPalette } from '../src/features/charts/chart-options.ts'
import { chartSeriesStyleValues, patchChartSeriesOverride } from '../src/features/charts/chart-cartesian-api.ts'
import { mergeChartOptions } from '../src/features/charts/chart-options.ts'
import { chartSettingsPatch, createChartSettingsDraft } from '../src/features/charts/chart-settings-draft.ts'
import { chartBoxplotStatisticOptions } from '../src/features/charts/chart-boxplot-options.ts'
import { chartTypeSettingsControls } from '../src/features/charts/chart-type-settings-controls.ts'
import { CHART_BOXPLOT_SAMPLE_CHOICES_LIMIT, chartSeriesSettingPoints, chartSeriesSettingSample, chartSeriesSettingSampleCount, chartSeriesSettingTargets } from '../src/features/charts/chart-series-settings.ts'

// P23 drafts only. No test runner, renderer, browser, Rust/WASM or benchmark has run.
// Quartiles/IQR/outliers below are the local Office extension, not fixed Xce/Yce equivalence.
const samples = [100, null, -100, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 100, -100]
const palette = ['green', 'orange']
const json = value => JSON.parse(JSON.stringify(value))
function source(second = false) {
  return { categories: samples.map((_value, index) => `Sample ${index + 1}`), series: [
    { field: 9, name: 'First samples', values: [...samples] },
    ...(second ? [{ field: 2, name: 'Second samples', values: samples.map(value => value === null ? null : value * 10) }] : []),
  ] }
}
function chart(patch = {}) {
  return { id: 'boxplot-source-points', sheetId: 'sheet', type: 'boxplot', categoryField: 0, valueFields: [9],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: samples.length, endCol: 9 } },
    position: { x: 90, y: 70 }, size: { width: 560, height: 400 }, legend: { visible: false, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear', min: -120, max: 120 }, revision: 0, ...patch,
    options: { ...patch.options, boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr', ...patch.options?.boxplot } } }
}
function prepare(data, model) { return deriveChartData(data, model, { preparedData: data }) }
function computed(data, derived) {
  return { preparedData: data, boxplots: derived.boxes.map(box => ({ field: box.field, seriesPosition: box.seriesPosition, statistics: box.statistics })) }
}
function native(data, model, derived = prepare(data, model)) { return chartEChartsOption({ sheets: [] }, model, {}, computed(data, derived)) }
function markers(data, model, derived = prepare(data, model)) { return drawDerivedChartData(derived, model, {}, palette) }
function outliers(option) { assert.ok(Array.isArray(option.series)); assert.equal(option.series[1].type, 'scatter'); return option.series[1].data }
function namedPaint(points = {}) {
  return { series: { 9: { color: 'green', point: { shape: 'rect', size: 4 }, dataPoints: points } } }
}
const pointPaint = {
  0: { color: 'red', size: 10, shape: 'rect', fillOpacity: 0.25, border: { color: 'navy', width: 2, opacity: 0.8 } },
  2: { color: 'orange', size: 8, shape: 'diamond' },
  13: { color: 'blue', size: 14, shape: 'square', fillOpacity: 0.75 },
  14: { color: 'purple', size: 12, shape: 'triangle' },
}

test('the local raw-sample distribution retains duplicate outlier identities and null slots in sorted value order', () => {
  const data = source(), original = structuredClone(data), model = chart(), result = prepare(data, model), box = result.boxes[0]
  assert.deepEqual(box.statistics, { min: 0, q1: 1.25, median: 4.5, q3: 7.75, max: 9,
    outliers: [-100, -100, 100, 100], outlierSourceIndices: [2, 14, 0, 13] })
  assert.deepEqual([...chartBoxplotOutliers(box)], [
    { value: -100, sourceIndex: 2, outlierIndex: 0 }, { value: -100, sourceIndex: 14, outlierIndex: 1 },
    { value: 100, sourceIndex: 0, outlierIndex: 2 }, { value: 100, sourceIndex: 13, outlierIndex: 3 },
  ])
  assert.deepEqual(data, original)
})

test('sample zero does not override the box summary, median, axis or distribution legend', () => {
  const data = source(), model = chart({ legend: { visible: true, position: 'bottom' }, options: namedPaint(pointPaint) }), result = prepare(data, model)
  const style = chartBoxplotStyle(model, result.boxes[0], 0, palette), layout = layoutDerivedChartData(result, model, palette), option = native(data, model, result)
  assert.equal(style.color, 'green'); assert.equal(style.fillOpacity, undefined)
  assert.deepEqual([layout.min, layout.max, layout.series[0].points[0].value], [-120, 120, 4.5])
  assert.equal(option.series[0].data[0].itemStyle.color, 'rgba(0, 128, 0, 0.45)')
  assert.equal(option.series[0].data[0].itemStyle.borderColor, 'rgba(0, 128, 0, 1)')
  const commands = markers(data, model, result), legend = commands.find(command => command.kind === 'text' && command.text === '样本分布')
  assert.ok(legend)
  assert.ok(commands.some(command => command.kind === 'rect' && command.x === legend.x - 15 && command.fill === 'rgba(0, 128, 0, 1)'))
})

test('native outliers use original source IDs and per-sample size, fill and border after duplicate-value sorting', () => {
  const data = source(), model = chart({ options: namedPaint(pointPaint) }), before = structuredClone(model), items = outliers(native(data, model))
  assert.deepEqual(items.map(item => [item.id, item.value]), [2, 14, 0, 13].map(index => [`${model.id}:box:0:outlier:${index}`, [0, samples[index]]]))
  assert.deepEqual(items.map(item => [item.symbol, item.symbolSize]), [['diamond', 8], ['triangle', 12], ['rect', 10], ['rect', 14]])
  assert.deepEqual(items.map(item => item.itemStyle.color), ['rgba(255, 165, 0, 1)', 'rgba(128, 0, 128, 1)', 'rgba(255, 0, 0, 0.25)', 'rgba(0, 0, 255, 0.75)'])
  assert.equal(items[2].itemStyle.borderColor, 'rgba(0, 0, 128, 0.8)'); assert.equal(items[2].itemStyle.borderWidth, 2)
  assert.equal(items.every(item => item.itemStyle.opacity === 1), true)
  assert.deepEqual(model, before)
})

test('shared Canvas outliers consume the same source styles at the distribution x and original raw value y', () => {
  const data = source(), model = chart({ options: namedPaint(pointPaint) }), result = prepare(data, model), layout = layoutDerivedChartData(result, model, palette)
  const style = chartBoxplotStyle(model, result.boxes[0], 0, palette), geometry = chartBoxplotGeometry(model, layout, style, 0, 1), commands = markers(data, model, result)
  for (const [index, size, fill] of [[0, 10, 'rgba(255, 0, 0, 0.25)'], [13, 14, 'rgba(0, 0, 255, 0.75)']]) {
    const command = commands.find(command => command.kind === 'rect' && command.width === size && command.fill === fill)
    assert.ok(command, String(index)); assert.equal(command.x + size / 2, geometry.x); assert.equal(command.y + size / 2, geometry.y(samples[index]))
    assert.deepEqual(command.clip, layout.plot)
  }
  assert.ok(commands.some(command => command.kind === 'line' && command.color === 'rgba(0, 0, 128, 0.8)' && command.width === 2))
  assert.deepEqual(result.boxes[0].statistics.outlierSourceIndices, [2, 14, 0, 13])
})

test('shared, field and source-name overrides inherit independently for each original sample and nullable name resets preserve lower layers', () => {
  const data = source(), model = chart({ options: {
    allSeriesStyle: { color: 'green', point: { size: 4 }, dataPoints: { 0: { size: 7, border: { width: 3 } } } },
    series: { 9: { color: 'blue', dataPoints: { 0: { color: 'red' }, 13: { size: 11 } } } },
    namedSeries: [{ name: 'First samples', style: { name: 'Display name', point: { color: 'orange' }, dataPoints: { 0: { size: 10, border: { color: 'purple' } }, 13: { color: 'navy' } } } }],
  } }), before = structuredClone(model.options), result = prepare(data, model)
  const base = chartBoxplotStyle(model, result.boxes[0], 0, palette), first = chartBoxplotOutlierStyle(base, 0), duplicate = chartBoxplotOutlierStyle(base, 13)
  assert.equal(base.name, 'Display name'); assert.equal(base.color, 'blue'); assert.equal(first.color, 'red'); assert.equal(first.point.size, 10)
  assert.deepEqual(first.border, { width: 3, color: 'purple' }); assert.equal(duplicate.color, 'navy'); assert.equal(duplicate.point.size, 11)
  model.options = mergeChartOptions(model.options, patchChartSeriesOverride(model.options, { name: 'First samples' }, chartSeriesStyleValues({ dataPoints: { 0: null } })))
  const reset = chartBoxplotOutlierStyle(chartBoxplotStyle(model, prepare(data, model).boxes[0], 0, palette), 0)
  assert.equal(reset.color, 'red'); assert.equal(reset.point.size, 7); assert.deepEqual(reset.border, { width: 3 })
  assert.equal(before.namedSeries[0].style.dataPoints[0].size, 10)
})

test('source sample styles remain tied to their field and name when distribution output order changes', () => {
  const data = source(true), model = chart({ valueFields: [2, 9], options: { series: { 9: { color: 'green', dataPoints: { 0: { color: 'red', size: 10 } } },
    2: { color: 'orange', dataPoints: { 13: { color: 'blue', size: 12 } } } } } })
  data.series.reverse()
  const items = outliers(native(data, model)), first = items.filter(item => item.id.includes(':box:0:')), second = items.filter(item => item.id.includes(':box:1:'))
  assert.equal(first.find(item => item.id.endsWith(':13')).itemStyle.color, 'rgba(0, 0, 255, 1)')
  assert.equal(second.find(item => item.id.endsWith(':0')).itemStyle.color, 'rgba(255, 0, 0, 1)')
  assert.deepEqual(first.map(item => item.value[0]), [0, 0, 0, 0]); assert.deepEqual(second.map(item => item.value[0]), [1, 1, 1, 1])
})

test('hidden and zero-sized sample markers remove only their own duplicate and leave statistics unchanged', () => {
  const data = source(), model = chart({ options: namedPaint({ ...pointPaint, 0: { shape: 'none' }, 14: { size: 0 } }) }), result = prepare(data, model), original = structuredClone(result.boxes)
  const items = outliers(native(data, model, result))
  assert.deepEqual(items.map(item => item.id), [2, 13].map(index => `${model.id}:box:0:outlier:${index}`))
  const commands = markers(data, model, result)
  assert.equal(commands.some(command => command.kind === 'rect' && command.fill === 'rgba(255, 0, 0, 0.25)'), false)
  assert.deepEqual(result.boxes, original)
  model.options.boxplot = { ...model.options.boxplot, showOutliers: false }
  assert.deepEqual(outliers(native(data, model, result)), [])
})

test('logarithmic omission keeps positive outliers bound to original rows while retaining negative summary statistics', () => {
  const data = source(), model = chart({ yAxis: { scaleType: 'log', logBase: 10, min: 1, max: 200 }, options: namedPaint(pointPaint) }), result = prepare(data, model)
  const option = native(data, model, result), items = outliers(option)
  assert.deepEqual(option.series[0].data[0].value, [])
  assert.deepEqual(items.map(item => [item.id, item.itemStyle.color]), [[`${model.id}:box:0:outlier:0`, 'rgba(255, 0, 0, 0.25)'], [`${model.id}:box:0:outlier:13`, 'rgba(0, 0, 255, 0.75)']])
  assert.deepEqual(result.boxes[0].statistics.outliers, [-100, -100, 100, 100]); assert.equal(result.boxes[0].statistics.min, 0)
})

test('only active drawable outlier shapes and patterns request shared graphics; inlier and hidden overrides remain authored', () => {
  const data = source(), model = chart({ options: namedPaint({ 3: { shape: 'emptyPin', fillPattern: { preset: 'ltUpDiag', foregroundColor: 'black', backgroundColor: 'white' } } }) })
  assert.ok(Array.isArray(native(data, model).series))
  model.options.series[9].dataPoints[0] = { shape: 'emptySquare', size: 8 }
  assert.equal(native(data, model).series, undefined)
  model.options.series[9].dataPoints[0].size = 0
  assert.ok(Array.isArray(native(data, model).series))
  model.options.series[9].dataPoints[0].size = 8; model.options.boxplot = { ...model.options.boxplot, showOutliers: false }
  assert.ok(Array.isArray(native(data, model).series))
  assert.equal(model.options.series[9].dataPoints[3].shape, 'emptyPin')
})

test('sample textures can be inherited and explicitly disabled without painting or changing a distribution summary', () => {
  const data = source(), pattern = { preset: 'ltUpDiag', foregroundColor: 'black', backgroundColor: 'transparent' }
  const model = chart({ options: { allSeriesStyle: { color: 'green', dataPoints: { 0: { fillPattern: pattern } } }, series: { 9: { point: { shape: 'rect', size: 10 } } } } })
  const result = prepare(data, model), base = chartBoxplotStyle(model, result.boxes[0], 0, palette)
  assert.equal(base.fillPattern, undefined); assert.deepEqual(chartBoxplotOutlierStyle(base, 0).fillPattern, pattern)
  assert.equal(native(data, model, result).series, undefined)
  assert.ok(markers(data, model, result).some(command => command.kind === 'polygon' && normalizeChartColor(command.fill) === 'rgba(0, 0, 0, 1)'))
  model.options.series[9].dataPoints = { 0: { fillPattern: null } }
  assert.equal(chartBoxplotOutlierStyle(chartBoxplotStyle(model, prepare(data, model).boxes[0], 0, palette), 0).fillPattern, null)
  assert.ok(Array.isArray(native(data, model).series))
})

test('missing or malformed outlier identities reject rather than inventing sorted point indices', () => {
  const model = chart(), box = prepare(source(), model).boxes[0]
  assert.throws(() => [...chartBoxplotOutliers({ ...box, statistics: { ...box.statistics, outlierSourceIndices: undefined } })], /原始样本索引不匹配/)
  assert.throws(() => [...chartBoxplotOutliers({ ...box, statistics: { ...box.statistics, outlierSourceIndices: [2] } })], /原始样本索引不匹配/)
  assert.throws(() => [...chartBoxplotOutliers({ ...box, statistics: { ...box.statistics, outlierSourceIndices: [2, -1, 0, 13] } })], /有效的原始样本索引/)
  assert.throws(() => chartBoxplotOutlierStyle({}, 0.5), /有效的原始样本索引/)
  assert.deepEqual([...chartBoxplotOutliers({ ...box, statistics: undefined })], [])
})

test('settings list actual source samples, omit empty slots and preserve field/name scopes and formatted values', () => {
  const data = source(true), model = chart(), fields = [{ value: 9, label: 'First' }, { value: 2, label: 'Second' }]
  data.series[1].formattedValues = samples.map(value => value === null ? undefined : `${value * 10} kg`)
  const choices = chartSeriesSettingPoints(model, { name: 'Second samples' }, ['First distribution', 'Second distribution'], fields, data)
  assert.deepEqual(choices.map(item => item.value), [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14])
  assert.equal(choices[0].label, '1 · Sample 1 · 1000 kg'); assert.equal(choices.at(-1).label, '15 · Sample 15 · -1000 kg')
  assert.equal(chartSeriesSettingSample(2, 1, data), undefined); assert.equal(chartSeriesSettingSampleCount(2, data), samples.length)
  assert.equal(chartSeriesSettingTargets(model, fields, data)[1].source.position, 1)
  data.series.push({ ...data.series[1], field: 3 })
  assert.deepEqual(chartSeriesSettingPoints(model, { name: 'Second samples' }, [], fields, data), [])
})

test('bounded sample options retain direct access to a valid sample beyond the list and never move source indices', () => {
  const index = CHART_BOXPLOT_SAMPLE_CHOICES_LIMIT + 600, values = Array(index + 1).fill(1), model = chart()
  values[index] = 100
  const data = { categories: [], series: [{ field: 9, name: 'Large raw samples', values }] }
  assert.equal(chartSeriesSettingPoints(model, 9, [], [], data).length, CHART_BOXPLOT_SAMPLE_CHOICES_LIMIT)
  assert.equal(chartSeriesSettingSample(9, index, data).label, `${index + 1} · 100`)
  const selected = chartSeriesSettingPoints(model, 9, [], [], data, index)
  assert.equal(selected.length, CHART_BOXPLOT_SAMPLE_CHOICES_LIMIT + 1); assert.equal(selected.at(-1).value, index)
  assert.equal(chartSeriesSettingSampleCount(9, data), index + 1); assert.equal(chartSeriesSettingSample(9, index + 1, data), undefined)
  assert.equal(values[index], 100)
})

test('settings deletion of a sample override survives actual series collection replacement and calendar box ordering preserves sample identities', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], table = [['First samples'], ...samples.map(value => [value])]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const model = createChart(workbook, sheet.id, { type: 'boxplot', source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: samples.length, endCol: 0 } },
    valueFields: [0], size: { width: 560, height: 400 }, position: { x: 90, y: 70 }, legend: { visible: false, position: 'bottom' },
    options: { boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr' }, series: { 0: { color: 'green', dataPoints: { 0: pointPaint[0], 13: pointPaint[13] } } } } })
  const before = structuredClone(model), draft = createChartSettingsDraft(model)
  delete draft.config.options.series[0].dataPoints[0]
  updateChart(workbook, model, chartSettingsPatch(draft))
  assert.equal(model.options.series[0].dataPoints[0], undefined); assert.deepEqual(json(model.options.series[0].dataPoints[13]), json(before.options.series[0].dataPoints[13]))
  assert.equal(model.revision, before.revision + 1)
  const prepared = readChartData(workbook, model), items = outliers(native(prepared, model))
  assert.equal(items.find(item => item.id.endsWith(':outlier:0')).itemStyle.color, 'rgba(0, 128, 0, 1)')
  const data = source(true), dated = chart({ valueFields: [9, 2], xAxis: { scaleType: 'linear', baseTimeUnit: 'months', majorUnit: 1 }, options: { ...namedPaint(pointPaint), mapping: { useDateAxis: true } } })
  data.series[0].name = '2024-03-31'; data.series[1].name = '2024-01-31'
  const result = prepare(data, dated), layout = layoutDerivedChartData(result, dated, chartPalette(dated))
  assert.deepEqual(layout.dateOrder, [1, 0]); assert.ok(layout.series[0].points[0].x > layout.series[1].points[0].x)
  assert.deepEqual(result.boxes[0].statistics.outlierSourceIndices, [2, 14, 0, 13])
  assert.equal(native(data, dated, result).series, undefined)
  assert.ok(markers(data, dated, result).some(command => command.kind === 'rect' && command.width === 10 && command.fill === 'rgba(255, 0, 0, 0.25)'))
})

test('fixed defaults use half medians and full-sample whiskers without activating authored raw-sample marker overrides', () => {
  const data = source(), model = chart({ options: { ...namedPaint(pointPaint), boxplot: { quartileMethod: undefined, whiskerMethod: undefined } } }), before = structuredClone(model)
  assert.deepEqual(chartBoxplotStatisticOptions(model.options.boxplot), { quartileMethod: 'medianOfHalves', whiskerMethod: 'range', whiskerMultiplier: 1.5 })
  const result = prepare(data, model)
  assert.deepEqual(result.boxes[0].statistics, { min: -100, q1: 1, median: 4.5, q3: 8, max: 100, outliers: [], outlierSourceIndices: [] })
  const option = native(data, model, result)
  assert.deepEqual(option.series[0].data[0].value, [-100, 1, 4.5, 8, 100]); assert.deepEqual(outliers(option), [])
  assert.equal(option.series[0].data[0].itemStyle.color, 'rgba(0, 128, 0, 0.45)')
  assert.equal(markers(data, model, result).some(command => command.kind === 'rect' && command.width === 10 && command.fill === 'rgba(255, 0, 0, 0.25)'), false)
  const singleton = { categories: ['Only sample'], series: [{ field: 9, name: 'First samples', values: [100] }] }
  assert.equal(prepare(singleton, model).boxes[0].statistics, undefined)
  assert.deepEqual(native(singleton, model).series[0].data[0].value, [])
  assert.deepEqual(model, before)
})

test('Boxplot controls resolve legacy IQR inheritance and independent resets through the real settings merge without losing sample author styles', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], table = [['First samples'], ...samples.map(value => [value])]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const model = createChart(workbook, sheet.id, { type: 'boxplot', source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: samples.length, endCol: 0 } },
    valueFields: [0], options: { boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr', whiskerMultiplier: 2, showOutliers: false },
      series: { 0: { color: 'green', dataPoints: { 0: pointPaint[0], 13: pointPaint[13] } } } } })
  const before = structuredClone(model), draft = createChartSettingsDraft(model), controls = chartTypeSettingsControls(draft.config)
  assert.deepEqual(controls.map(control => control.key), ['boxplot.quartileMethod', 'boxplot.whiskerMethod', 'boxplot.showOutliers', 'boxplot.whiskerMultiplier'])
  const field = key => { const item = controls.find(control => control.key === 'boxplot.' + key); assert.ok(item?.reset, key); return item }
  field('whiskerMethod').reset(draft.config)
  assert.equal(Object.hasOwn(draft.config.options.boxplot, 'whiskerMethod'), true)
  assert.equal(field('whiskerMethod').read(draft.config), 'iqr'); assert.equal(field('whiskerMethod').canReset(draft.config), false)
  field('quartileMethod').reset(draft.config)
  assert.equal(field('quartileMethod').read(draft.config) ?? field('quartileMethod').fallback, 'medianOfHalves')
  assert.equal(field('whiskerMethod').read(draft.config), 'iqr'); assert.equal(field('whiskerMultiplier').visible(draft.config), true)
  field('whiskerMultiplier').reset(draft.config)
  assert.equal(field('whiskerMethod').read(draft.config), 'range'); assert.equal(field('whiskerMultiplier').visible(draft.config), false)
  field('showOutliers').reset(draft.config)
  assert.equal(field('showOutliers').read(draft.config) ?? field('showOutliers').fallback, true)
  assert.equal(controls.every(control => !control.canReset(draft.config)), true)
  assert.deepEqual(model, before)
  updateChart(workbook, model, chartSettingsPatch(draft))
  assert.deepEqual(json(model.options.boxplot), {})
  assert.deepEqual(model.options.series, before.options.series); assert.equal(model.revision, before.revision + 1)
  const data = readChartData(workbook, model), result = prepare(data, model)
  assert.deepEqual(result.boxes[0].statistics.outliers, []); assert.deepEqual(result.boxes[0].statistics.outlierSourceIndices, [])
  assert.deepEqual(chartTypeSettingsControls(model).filter(control => control.visible?.(model) !== false).map(control => control.key),
    ['boxplot.quartileMethod', 'boxplot.whiskerMethod', 'boxplot.showOutliers'])
})
