import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, updateChart } from '../src/features/charts/chart-engine.ts'
import { createChartSettingsDraft, chartSettingsPatch } from '../src/features/charts/chart-settings-draft.ts'
import { chartAxisControls, chartCommonControls, chartDateAxisControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartHeatmapSettingsControls } from '../src/features/charts/chart-heatmap-settings-controls.ts'
import { chartRadarSettingsControls } from '../src/features/charts/chart-radar-settings-controls.ts'
import { clearChartSeriesSettingPoint } from '../src/features/charts/chart-series-settings.ts'
import { chartSeriesStyle } from '../src/features/charts/chart-options.ts'
import { chartPointStyle } from '../src/features/charts/chart-derived.ts'

// P22/P25 regression drafts only. No runner, model command, browser or Office sample has been executed.
function fixture(type, patch = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], table = [['Date', 'Units'], [45322, 10], [45351, 20], [45382, 30]]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } },
    position: { x: 90, y: 70 }, size: { width: 520, height: 360 }, legend: { visible: false, position: 'bottom' }, ...patch })
  return { workbook, chart }
}
const control = (controls, key) => { const result = controls.find(item => item.key === key); assert.ok(result?.reset, key); return result }
const json = value => JSON.parse(JSON.stringify(value))

test('date unit, toggle, numeric interval and whole-axis title resets survive the actual settings patch merge', () => {
  const { workbook, chart } = fixture('line', { options: { mapping: { useDateAxis: true }, appearance: { font: { family: 'Georgia' } } },
    xAxis: { scaleType: 'linear', baseTimeUnit: 'months', majorTimeUnit: 'years', majorUnit: 2, reverse: true,
      title: 'Dates', titleAlignment: 'end', titleStyle: { fontSize: 20, color: 'red' }, line: { width: 2 } },
    yAxis: { scaleType: 'linear', min: 0, max: 100, majorUnit: 25, title: 'Units' } })
  const before = structuredClone(chart), draft = createChartSettingsDraft(chart), axes = chartAxisControls('xAxis', true, true)
  for (const key of ['xAxis.baseTimeUnit', 'xAxis.majorTimeUnit', 'xAxis.majorUnit', 'xAxis.title']) control(axes, key).reset(draft.config)
  chartDateAxisControls()[0].reset(draft.config)
  control(chartAxisControls('yAxis'), 'yAxis.majorUnit').reset(draft.config)
  assert.deepEqual(chart, before)
  const patch = chartSettingsPatch(draft)
  for (const field of ['baseTimeUnit', 'majorTimeUnit', 'majorUnit', 'title', 'titleAlignment', 'titleStyle']) {
    assert.equal(Object.hasOwn(patch.xAxis, field), true, field)
    assert.equal(patch.xAxis[field], undefined, field)
  }
  assert.equal(Object.hasOwn(patch.options.mapping, 'useDateAxis'), true)
  updateChart(workbook, chart, patch)
  assert.deepEqual(json(chart.xAxis), { scaleType: 'linear', reverse: true, line: { width: 2 } })
  assert.deepEqual(json(chart.yAxis), { scaleType: 'linear', min: 0, max: 100, title: 'Units' })
  assert.deepEqual(json(chart.options), { mapping: {}, appearance: { font: { family: 'Georgia' } } })
  assert.equal(chart.revision, before.revision + 1)
  assert.deepEqual(chart.source, before.source)
  for (const key of ['xAxis.baseTimeUnit', 'xAxis.majorTimeUnit', 'xAxis.majorUnit', 'xAxis.title']) assert.equal(control(axes, key).canReset(draft.config), false, key)
})

test('Heatmap field and label resets replace prior overrides through deep options merge while unrelated author settings remain', () => {
  const { workbook, chart } = fixture('heatmap', { options: { heatmap: { visualMapType: 'piecewise', min: 1, max: 40, unit: 'kg',
    label: { visible: true, position: 'bottom', fontSize: 18, format: '0.00' } }, appearance: { background: 'white' } } })
  const before = structuredClone(chart), draft = createChartSettingsDraft(chart), controls = chartHeatmapSettingsControls(), common = chartCommonControls(draft.config)
  for (const key of ['heatmap.min', 'heatmap.unit', 'heatmap.label.fontSize', 'heatmap.label.position']) control(controls, key).reset(draft.config)
  control(common, 'heatmap.visualMapType').reset(draft.config)
  updateChart(workbook, chart, chartSettingsPatch(draft))
  assert.deepEqual(json(chart.options.heatmap), { max: 40, label: { visible: true, format: '0.00' } })
  assert.deepEqual(chart.options.appearance, before.options.appearance)
  assert.equal(control(common, 'heatmap.visualMapType').canReset(draft.config), false)
  assert.equal(control(controls, 'heatmap.label.fontSize').canReset(draft.config), false)
  const wholeLabel = createChartSettingsDraft(chart)
  control(controls, 'heatmap.label.visible').reset(wholeLabel.config)
  updateChart(workbook, chart, chartSettingsPatch(wholeLabel))
  assert.deepEqual(json(chart.options.heatmap), { max: 40 })
  assert.equal(chart.revision, before.revision + 2)
})

test('Radar leaf and group resets clear old deep overrides and preserve inherited fonts through actual updates', () => {
  const { workbook, chart } = fixture('radar', { options: { radar: { shape: 'circle', fill: true,
    axisLabel: { visible: true, format: '0.00', family: 'Aptos', fontSize: 20 } }, appearance: { font: { family: 'Georgia', fontSize: 16 } } } })
  const before = structuredClone(chart), draft = createChartSettingsDraft(chart)
  control(chartRadarSettingsControls(draft.config), 'radar.axisLabel.fontSize').reset(draft.config)
  control(chartRadarSettingsControls(draft.config), 'radar.axisLabel.format').reset(draft.config)
  updateChart(workbook, chart, chartSettingsPatch(draft))
  assert.deepEqual(json(chart.options.radar), { shape: 'circle', fill: true, axisLabel: { visible: true, family: 'Aptos' } })
  const groups = createChartSettingsDraft(chart)
  for (const key of ['radar.shape', 'radar.fill', 'radar.axisLabel.visible']) control(chartRadarSettingsControls(groups.config), key).reset(groups.config)
  updateChart(workbook, chart, chartSettingsPatch(groups))
  assert.deepEqual(json(chart.options.radar), {})
  assert.deepEqual(chart.options.appearance, before.options.appearance)
  assert.equal(chart.revision, before.revision + 2)
  for (const key of ['radar.shape', 'radar.fill', 'radar.axisLabel.visible']) assert.equal(control(chartRadarSettingsControls(groups.config), key).canReset(groups.config), false, key)
})

test('the actual field-point reset replaces its author map and resumes the shared point without deleting other points', () => {
  const { workbook, chart } = fixture('column', { options: {
    allSeriesStyle: { dataPoints: { 0: { color: 'red', fillOpacity: 0.4, border: { color: 'navy', width: 2 } } } },
    series: { 1: { color: 'green', dataPoints: { 0: { color: 'blue', size: 12 }, 1: { color: 'gold', size: 18 } } },
      0: { dataPoints: { 0: { color: 'black' } } } }, appearance: { background: 'white' },
  } })
  const before = structuredClone(chart), draft = createChartSettingsDraft(chart)
  clearChartSeriesSettingPoint(draft.config, 1, 0)
  assert.deepEqual(chart, before)
  assert.equal(Object.hasOwn(draft.config.options.series[1].dataPoints, 0), false)
  updateChart(workbook, chart, chartSettingsPatch(draft))
  assert.equal(Object.hasOwn(chart.options.series[1].dataPoints, 0), false)
  assert.deepEqual(chart.options.series[1].dataPoints[1], before.options.series[1].dataPoints[1])
  assert.deepEqual(chart.options.series[0], before.options.series[0])
  assert.deepEqual(chart.options.allSeriesStyle, before.options.allSeriesStyle)
  const effective = chartPointStyle(chartSeriesStyle(chart, 1, 0, undefined, 'Units'), 0, chart.type)
  assert.equal(effective.color, 'red'); assert.equal(effective.fillOpacity, 0.4)
  assert.deepEqual(effective.border, { color: 'navy', width: 2 })
  assert.equal(chart.revision, before.revision + 1)
})

test('the actual named-point reset replaces its author array and resumes field/shared layers with precise source identity', () => {
  const { workbook, chart } = fixture('line', { options: {
    allSeriesStyle: { dataPoints: { 0: { color: 'red', border: { color: 'navy', width: 2 } } } },
    series: { 1: { dataPoints: { 0: { color: 'green', fillOpacity: 0.6 } } } },
    namedSeries: [{ name: 'Units', style: { name: 'Measurements', dataPoints: { 0: { color: 'blue' }, 1: { color: 'gold' } } } },
      { name: 'Inactive', style: { dataPoints: { 0: { color: 'black' } } } }],
  } })
  const before = structuredClone(chart), draft = createChartSettingsDraft(chart)
  clearChartSeriesSettingPoint(draft.config, { name: 'Units' }, 0)
  assert.deepEqual(chart, before)
  updateChart(workbook, chart, chartSettingsPatch(draft))
  assert.equal(Object.hasOwn(chart.options.namedSeries[0].style.dataPoints, 0), false)
  assert.deepEqual(chart.options.namedSeries[0].style.dataPoints[1], { color: 'gold' })
  assert.deepEqual(chart.options.namedSeries[1], before.options.namedSeries[1])
  assert.deepEqual(chart.options.series, before.options.series)
  const effective = chartPointStyle(chartSeriesStyle(chart, 1, 0, undefined, 'Units'), 0, chart.type)
  assert.equal(effective.color, 'green'); assert.equal(effective.fillOpacity, 0.6)
  assert.deepEqual(effective.border, { color: 'navy', width: 2 })
  assert.equal(effective.name, 'Measurements')
  assert.equal(chart.revision, before.revision + 1)
})
