import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart, readChartData, updateChart } from '../src/features/charts/chart-engine.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { buildChartHistogram, chartHistogramBinGap, chartHistogramCategories, chartHistogramRuntimeOptions,
  chartHistogramSampleStandardDeviation } from '../src/features/charts/chart-histogram.ts'
import { prepareChartComputeInputs } from '../src/features/charts/chart-input-preparation.ts'
import { prepareChartStatisticsTask, decodeChartStatistics } from '../src/features/charts/chart-statistics-compute.ts'
import { chartHistogramDataControls, chartHistogramStyleControls } from '../src/features/charts/chart-histogram-settings-controls.ts'
import { chartSettingsBindingKey, chartSettingsPreparedStatistics } from '../src/features/charts/chart-settings-data.ts'
import { chartSettingsPatch, createChartSettingsDraft } from '../src/features/charts/chart-settings-draft.ts'
import { validateChartOptions } from '../src/features/charts/chart-options.ts'

// P25 contract drafts. No test, compiler, Worker, browser or benchmark ran.
const json = value => JSON.parse(JSON.stringify(value))
function fixture(values = [0, 1, 2, 3, 4], histogram = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'histogram-draft'
  const table = [['Label', 'Sample'], ...values.map((value, index) => ['Row ' + index, value])]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type: 'histogram', categoryField: 0, valueFields: [1],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: table.length - 1, endCol: 1 } },
    position: { x: 0, y: 0 }, size: { width: 420, height: 300 }, options: { histogram } }, 'histogram-draft')
  return { workbook, chart, data: readChartData(workbook, chart) }
}
function packet(values, indices, count, available = count > 0) {
  return { kind: 'chartStatistics', values: new Float64Array(values), indices: new Uint32Array(indices),
    segments: [{ kind: 'histogram', jobIndex: 0, field: 1, seriesPosition: 0, slot: 0,
      valueOffset: 0, valueLength: values.length, indexOffset: 0, indexLength: indices.length, validSampleCount: count, available }] }
}
const getControl = (controls, key) => { const value = controls.find(item => item.key === key); assert.ok(value, key); return value }

test('Scott uses sample deviation, while count bins cover the observed range and include edge bins in the total', () => {
  assert.equal(chartHistogramSampleStandardDeviation([0, 1, 2]), 1)
  assert.deepEqual(buildChartHistogram([0, 1, 2]), { validSampleCount: 3,
    bins: [{ kind: 'regular', lowerBound: 0, upperBound: 2, count: 3 }] })
  const samples = [0, 1, 2, 3, 4], options = { binning: { mode: 'count', count: 4 }, underflowThreshold: 1, overflowThreshold: 3 }
  const result = buildChartHistogram(samples, options)
  assert.deepEqual(result.bins, [{ kind: 'underflow', upperBound: 1, count: 2 },
    { kind: 'regular', lowerBound: 2, upperBound: 2.5, count: 1 },
    { kind: 'regular', lowerBound: 2.5, upperBound: 3, count: 1 }, { kind: 'overflow', lowerBound: 3, count: 1 }])
  assert.deepEqual(chartHistogramCategories(result), ['≤ 1', '[2, 2.5]', '(2.5, 3]', '> 3'])
  const capped = buildChartHistogram(samples, { ...options, binning: { mode: 'count', count: 1000 } })
  assert.equal(capped.bins.length, 1000); assert.equal(capped.bins.filter(bin => bin.kind === 'regular').length, 998)
  assert.equal(capped.bins.reduce((total, bin) => total + bin.count, 0), 5)
})

test('ineffective finite author values survive validation and fall back without rewriting the configuration', () => {
  const auto = buildChartHistogram([0, 1, 2])
  for (const binning of [{ mode: 'count', count: 0 }, { mode: 'count', count: 1.5 }, { mode: 'count', count: 1001 },
    { mode: 'width', width: 0 }, { mode: 'width', width: -2 }]) {
    const authored = { binning, underflowThreshold: 10, overflowThreshold: 0, binGap: 2 }, before = structuredClone(authored)
    assert.doesNotThrow(() => validateChartOptions({ histogram: authored }))
    assert.deepEqual(buildChartHistogram([0, 1, 2], authored), auto)
    assert.equal(chartHistogramBinGap(authored), 0.007); assert.deepEqual(authored, before)
  }
  for (const key of ['binGap', 'underflowThreshold', 'overflowThreshold']) assert.throws(() => validateChartOptions({ histogram: { [key]: Infinity } }))
  assert.throws(() => validateChartOptions({ histogram: { binning: { mode: 'count', count: NaN } } }))
  assert.deepEqual(chartHistogramRuntimeOptions({ underflowThreshold: 2, overflowThreshold: 2 }), {})
})

test('width uses the threshold origin, right-closed boundaries, truncated final bins and a bounded Auto fallback', () => {
  const result = buildChartHistogram([1, 4, 7], { binning: { mode: 'width', width: 5 }, underflowThreshold: 0, overflowThreshold: 7 })
  assert.deepEqual(result.bins, [{ kind: 'underflow', upperBound: 0, count: 0 }, { kind: 'regular', lowerBound: 0, upperBound: 5, count: 2 },
    { kind: 'regular', lowerBound: 5, upperBound: 7, count: 1 }, { kind: 'overflow', lowerBound: 7, count: 0 }])
  assert.deepEqual(chartHistogramCategories(result), ['≤ 0', '(0, 5]', '(5, 7]', '> 7'])
  assert.deepEqual(buildChartHistogram([0, 1, 2], { binning: { mode: 'width', width: 0.000001 } }), buildChartHistogram([0, 1, 2]))
  assert.deepEqual(buildChartHistogram([5, 5], { binning: { mode: 'width', width: 2 } }).bins,
    [{ kind: 'regular', lowerBound: 5, upperBound: 7, count: 2 }])
})

test('constant, all-edge, empty and extreme observations retain the actual samples without padding or phantom edge bins', () => {
  for (const value of [7, Number.MAX_VALUE, Number.MIN_VALUE]) {
    const result = buildChartHistogram([value, null, value], { binning: { mode: 'count', count: 5 } })
    assert.deepEqual(result, { validSampleCount: 2, bins: [{ kind: 'regular', lowerBound: value, upperBound: value, count: 2 }] })
  }
  assert.deepEqual(buildChartHistogram([-Number.MAX_VALUE, Number.MAX_VALUE]).bins,
    [{ kind: 'regular', lowerBound: -Number.MAX_VALUE, upperBound: Number.MAX_VALUE, count: 2 }])
  assert.deepEqual(buildChartHistogram([1, 2, 3], { underflowThreshold: 5, overflowThreshold: 10 }).bins,
    [{ kind: 'underflow', upperBound: 5, count: 3 }])
  assert.deepEqual(buildChartHistogram([20, 30], { underflowThreshold: 5, overflowThreshold: 10 }).bins,
    [{ kind: 'overflow', lowerBound: 10, count: 2 }])
  assert.deepEqual(buildChartHistogram([null, NaN, Infinity], { underflowThreshold: 5, overflowThreshold: 10 }), { validSampleCount: 0, bins: [] })
})

test('Histogram null samples remain absent even when zero is the authored invalid-value strategy', async () => {
  const { chart, data } = fixture([1, null, 3])
  chart.options.invalidValueStrategy = 'zero'
  const prepared = await prepareChartComputeInputs(data, chart)
  assert.deepEqual(prepared.series[0].values, [1, null, 3])
  const result = deriveChartData(prepared, chart)
  assert.equal(result.histogram.validSampleCount, 2)
})

test('materialized Histogram decoding accepts constant point bins and verifies source counts instead of regenerating boundaries', async () => {
  const { chart, data } = fixture([7, 7, 7], { binning: { mode: 'count', count: 4 } })
  const plan = await prepareChartStatisticsTask(data, chart), decoded = await decodeChartStatistics(plan, packet([7, 7], [1, 3], 3))
  assert.deepEqual(decoded.histogram, { validSampleCount: 3, bins: [{ kind: 'regular', lowerBound: 7, upperBound: 7, count: 3 }] })
  for (const malformed of [packet([7, 8], [1, 3], 3), packet([7, 7], [1, 2], 3), packet([7, 7], [1, 3], 3, false)]) {
    await assert.rejects(() => decodeChartStatistics(plan, malformed), reason => reason.code === 'COMPUTE_PROTOCOL_MISMATCH')
  }
  const ranged = fixture([0, 1, 2, 3, 4], { binning: { mode: 'count', count: 4 }, underflowThreshold: 1, overflowThreshold: 3 })
  const rangePlan = await prepareChartStatisticsTask(ranged.data, ranged.chart)
  const output = packet([0, 1, 2, 2.5, 2.5, 3, 3, 0], [0, 2, 1, 1, 1, 1, 2, 1], 5)
  assert.equal((await decodeChartStatistics(rangePlan, output)).histogram.bins.length, 4)
  output.indices.set([0, 1, 1, 2, 1, 1, 2, 1])
  await assert.rejects(() => decodeChartStatistics(rangePlan, output), reason => reason.code === 'COMPUTE_PROTOCOL_MISMATCH')
})

test('data/style resets clear the actual merged model and binding identities only change for statistical inputs', () => {
  const { workbook, chart } = fixture([0, 1, 2, 3, 4], { binning: { mode: 'count', count: 4 }, underflowThreshold: 1, overflowThreshold: 3, binGap: 0.25 })
  assert.equal(chart.legend.visible, false)
  const before = structuredClone(chart), draft = createChartSettingsDraft(chart), key = chartSettingsBindingKey(draft)
  getControl(chartHistogramStyleControls(), 'histogram.binGap').reset(draft.config)
  assert.equal(chartSettingsBindingKey(draft), key)
  getControl(chartHistogramDataControls(draft.config), 'histogram.binning.count').reset(draft.config)
  assert.notEqual(chartSettingsBindingKey(draft), key)
  for (const property of ['underflowThreshold', 'overflowThreshold']) getControl(chartHistogramDataControls(draft.config), 'histogram.' + property).reset(draft.config)
  assert.deepEqual(chart, before)
  updateChart(workbook, chart, chartSettingsPatch(draft))
  assert.deepEqual(json(chart.options.histogram), {}); assert.equal(chart.revision, before.revision + 1)
  const binning = getControl(chartHistogramDataControls(draft.config), 'histogram.binning.mode')
  binning.write(draft.config, 'count')
  assert.deepEqual(draft.config.options.histogram.binning, { mode: 'count', count: 10 })
  assert.ok(chartHistogramDataControls(draft.config).some(control => control.key === 'histogram.binning.count'))
  binning.write(draft.config, 'width')
  assert.deepEqual(draft.config.options.histogram.binning, { mode: 'width', width: 1 })
  const statistics = { preparedData: readChartData(workbook, chart), histogram: buildChartHistogram([0, 1, 2, 3, 4]) }
  assert.equal(chartSettingsPreparedStatistics(createChartSettingsDraft(chart), statistics).statistics, statistics)
  assert.throws(() => chartSettingsPreparedStatistics(createChartSettingsDraft(chart), { preparedData: statistics.preparedData }), /分箱/)
})
