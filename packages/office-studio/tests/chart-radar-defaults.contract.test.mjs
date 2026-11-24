import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartRadarAxisLabel, chartRadarGeometry, patchChartRadarOptions } from '../src/features/charts/chart-radar.ts'
import { chartRadarSettingsControls } from '../src/features/charts/chart-radar-settings-controls.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { drawDerivedChartData } from '../src/features/charts/chart-draw.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { chartAxisPointerCommands, chartAxisPointerGeometry } from '../src/features/charts/chart-axis-pointer.ts'
import { cloneChartDraftConfig } from '../src/features/charts/chart-settings-draft.ts'

// P10/P11 implementation drafts only. No runner, build, browser or benchmark has run.
const source = () => ({ categories: ['North', 'East', 'South', 'West'], series: [{ field: 1, name: 'Score', values: [0, 25, 50, 100] }] })
function model(patch = {}) {
  return { id: 'radar-defaults', sheetId: 'sheet', type: 'radar', revision: 0, categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 4, endCol: 1 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 320 }, legend: { visible: false, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear', min: 0, max: 100 }, ...patch }
}
const control = (config, key) => chartRadarSettingsControls(config).find(field => field.key === key)
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)

test('fixed Radar defaults show numeric labels at 12px, while explicit and legacy false remain authoritative', () => {
  const chart = model(), before = structuredClone(chart)
  assert.deepEqual(chartRadarAxisLabel(chart), { fontSize: 12, format: undefined, visible: true })
  assert.equal(chartRadarSettingsControls(chart).length, 9)
  chart.yAxis.label = { visible: false, family: 'Georgia', fontSize: 16 }
  assert.equal(chartRadarAxisLabel(chart).visible, false)
  assert.equal(chartRadarSettingsControls(chart).length, 3)
  chart.options = { radar: { axisLabel: { visible: true } } }
  assert.deepEqual(chartRadarAxisLabel(chart), { fontSize: 16, family: 'Georgia', format: undefined, visible: true })
  delete chart.yAxis.label; delete chart.options
  assert.deepEqual(chart, before)
})

test('whole and field resets restore fixed defaults and inherited font without changing nullable author inputs', () => {
  const chart = model({ options: { appearance: { font: { family: 'Aptos', fontSize: 18 } }, radar: { axisLabel: { visible: false, fontSize: 24, format: '0.000' } } } })
  const input = { axisLabel: { visible: null, fontSize: null } }, before = structuredClone(input)
  chart.options.radar = patchChartRadarOptions(chart.options.radar, input)
  assert.equal(chartRadarAxisLabel(chart).visible, true)
  assert.equal(chartRadarAxisLabel(chart).fontSize, 18)
  assert.equal(chartRadarAxisLabel(chart).format, '0.000')
  assert.deepEqual(input, before)
  chart.options.radar = patchChartRadarOptions(chart.options.radar, { axisLabel: null })
  assert.equal(chartRadarAxisLabel(chart).visible, true)
  assert.equal(chartRadarAxisLabel(chart).family, 'Aptos')
  assert.equal(chartRadarAxisLabel(chart).format, undefined)
})

test('shared radial ticks have five intervals for signed and extreme finite ranges and retain author bounds', () => {
  const values = source(), chart = model({ yAxis: { scaleType: 'linear', min: -20, max: 80 } }), before = structuredClone(chart)
  const geometry = chartRadarGeometry(values, chart, { x: 40, y: 20, width: 300, height: 200 })
  assert.deepEqual(geometry.ticks, [-20, 0, 20, 40, 60, 80])
  chart.yAxis.min = -Number.MAX_VALUE; chart.yAxis.max = Number.MAX_VALUE
  const extreme = chartRadarGeometry(values, chart, { x: 40, y: 20, width: 300, height: 200 })
  assert.equal(extreme.ticks.length, 6)
  assert.equal(extreme.ticks.every(Number.isFinite), true)
  assert.equal(extreme.ticks[0], -Number.MAX_VALUE); assert.equal(extreme.ticks.at(-1), Number.MAX_VALUE)
  chart.yAxis = before.yAxis
  assert.deepEqual(chart, before)
})

test('final automatic radius reserves chart-font padding, manual plots keep the exact radius and tiny plots have no pointer', () => {
  const values = source(), plot = { x: 40, y: 20, width: 300, height: 200 }, chart = model()
  near(chartRadarGeometry(values, chart, plot).radius, 76)
  chart.options = { appearance: { font: { fontSize: 18 } } }
  near(chartRadarGeometry(values, chart, plot).radius, 70)
  chart.options.plotArea = { left: 0.1, top: 0.2, width: 0.5, height: 0.5 }
  near(chartRadarGeometry(values, chart, plot).radius, 100)
  delete chart.options.plotArea
  const layout = { frame: { x: 0, y: 0, width: 20, height: 20 }, plot: { x: 2, y: 2, width: 16, height: 16 }, min: 0, max: 100, categories: values.categories, series: [] }
  const geometry = chartAxisPointerGeometry(values, chart, layout)
  assert.equal(geometry.radar.radius, 0)
  assert.deepEqual(chartAxisPointerCommands(geometry, geometry.radar.center), [])
})

test('all nine dedicated settings have real resets, preserve hidden fields and edit detached drafts', () => {
  const original = model({ options: { appearance: { font: { family: 'Aptos', fontSize: 18 } }, radar: { shape: 'circle', fill: true,
    axisLabel: { visible: true, format: '0.000', family: 'Georgia', fontSize: 22, color: 'red', bold: true, italic: true } } } })
  const before = structuredClone(original), draft = cloneChartDraftConfig(original)
  assert.equal(chartRadarSettingsControls(draft).every(field => typeof field.reset === 'function'), true)
  control(draft, 'radar.axisLabel.fontSize').reset(draft)
  assert.equal(chartRadarAxisLabel(draft).fontSize, 18)
  assert.equal(control(draft, 'radar.axisLabel.fontSize').canReset(draft), false)
  control(draft, 'radar.axisLabel.format').reset(draft)
  assert.equal(chartRadarAxisLabel(draft).format, undefined)
  control(draft, 'radar.axisLabel.visible').write(draft, false)
  assert.equal(chartRadarSettingsControls(draft).length, 3)
  assert.equal(draft.options.radar.axisLabel.family, 'Georgia')
  control(draft, 'radar.axisLabel.visible').reset(draft)
  assert.equal(draft.options.radar.axisLabel, undefined)
  assert.equal(chartRadarSettingsControls(draft).length, 9)
  assert.equal(chartRadarAxisLabel(draft).visible, true)
  assert.equal(chartRadarAxisLabel(draft).family, 'Aptos')
  control(draft, 'radar.shape').reset(draft); control(draft, 'radar.fill').reset(draft)
  assert.equal(control(draft, 'radar.shape').read(draft) ?? control(draft, 'radar.shape').fallback, 'polygon')
  assert.equal(control(draft, 'radar.fill').read(draft) ?? control(draft, 'radar.fill').fallback, false)
  assert.deepEqual(original, before)
})

test('native option and shared Canvas expose the fixed five intervals, font and visible labels without author writes', () => {
  const values = source(), chart = model(), before = structuredClone(chart)
  const option = chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: values })
  assert.equal(option.radar.splitNumber, 5)
  assert.equal(option.radar.axisLabel.show, true)
  assert.equal(option.radar.axisLabel.fontSize, 12)
  assert.equal(option.radar.axisLabel.margin, 4)
  assert.equal(option.radar.axisLabel.hideOverlap, true)
  const derived = deriveChartData(values, chart, { preparedData: values }), layout = layoutDerivedChartData(derived, chart)
  const commands = drawDerivedChartData(derived, chart)
  const labels = commands.filter(command => command.kind === 'text' && ['20', '40', '60', '80', '100'].includes(command.text))
  assert.equal(labels.length, 20)
  assert.equal(labels.every(command => /12px/.test(command.font)), true)
  near(option.radar.radius, chartRadarGeometry(values, chart, layout.plot).radius)
  assert.deepEqual(chart, before)
  chart.options = { radar: { axisLabel: { visible: false } } }
  assert.equal(chartEChartsOption({ sheets: [] }, chart, {}, { preparedData: values }).radar.axisLabel.show, false)
  assert.equal(drawDerivedChartData(deriveChartData(values, chart, { preparedData: values }), chart).some(command => command.kind === 'text' && command.text === '20'), false)
})
