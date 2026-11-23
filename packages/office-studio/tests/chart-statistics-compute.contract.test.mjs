import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { prepareChartStatisticsTask, decodeChartStatistics } from '../src/features/charts/chart-statistics-compute.ts'
import { readChartRecordsAsync } from '../src/features/charts/chart-source-compute.ts'
import { prepareChartComputeInputs } from '../src/features/charts/chart-input-preparation.ts'
import { projectChartRecords } from '../src/features/charts/chart-record-projection.ts'
import { consumeChartPreparation } from '../src/features/charts/chart-compute-lifecycle.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'

// Draft contracts for the unified verification phase, not evidence of a run.
function chart(extra = {}) {
  return { id: 'statistics-chart', sheetId: 'sheet', type: 'line', revision: 0,
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 5, endCol: 1 } },
    position: { x: 20, y: 20 }, size: { width: 360, height: 240 },
    legend: { visible: true, position: 'bottom' }, categoryField: 0, valueFields: [1],
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, ...extra }
}
function data(values = [0, 1, 2, 3, 4]) {
  return { categories: values.map((_value, index) => String(index)), fields: ['Label', 'Value'],
    series: [{ field: 1, name: 'Value', values }], records: values.map((value, index) => [String(index), value]) }
}
function segment(job, extra = {}) {
  return { kind: job.kind, jobIndex: 0, field: 1, seriesPosition: 0, slot: job.slot, valueOffset: 0, indexOffset: 0,
    valueLength: 0, indexLength: 0, validSampleCount: 0, available: false, ...extra }
}
const histogramChart = () => chart({ type: 'histogram', options: { histogram: {
  binning: { mode: 'count', count: 2 }, underflowThreshold: 1, overflowThreshold: 3 } } })
const histogramOutput = () => ({ kind: 'chartStatistics',
  values: new Float64Array([0, 1, 1, 2, 2, 3, 3, 0]), indices: new Uint32Array([0, 2, 1, 1, 1, 1, 2, 1]),
  segments: [segment({ kind: 'histogram', slot: 0 }, { valueLength: 8, indexLength: 8, validSampleCount: 5, available: true })] })

test('source preparation reads authoritative raw numbers in either orientation, including their original precision', async () => {
  const number = 1.234567890123
  for (const orientation of ['columns', 'rows']) {
    const matrix = orientation === 'columns' ? [['Label', 'Value'], ['A', number], ['B', 2]]
      : [['Label', 'A', 'B'], ['Value', number, 2]]
    const workbook = { ...createBlankWorkbook(), id: 'book' }, sheet = workbook.sheets[0]
    sheet.id = 'sheet'; sheet.rowCount = matrix.length; sheet.columnCount = matrix[0].length
    sheet.cells = Object.fromEntries(matrix.flatMap((row, rowIndex) => row.map((value, column) => {
      const cell = createCell(value)
      if (value === number) cell.display = '1.23'
      return [rowIndex + ':' + column, cell]
    })))
    const source = chart({ source: { orientation, range: { startRow: 0, startCol: 0,
      endRow: matrix.length - 1, endCol: matrix[0].length - 1 } } })
    const result = await readChartRecordsAsync(workbook, source)
    assert.deepEqual(result.fields, ['Label', 'Value'])
    assert.deepEqual(result.records, [['A', number], ['B', 2]])
    assert.equal(sheet.cells[orientation === 'columns' ? '1:1' : '1:1'].display, '1.23')
  }
})

test('planner keeps authored slots, shares requested Y/X buffers and ignores the unused custom error side', async () => {
  const source = { categories: ['A', 'A', 'B'], fields: ['Label', 'Value', 'Plus'], records: [['A', 2, 1], ['A', null, 2], ['B', 6, 3]],
    series: [{ field: 1, name: 'Value', values: [2, null, 6], xValues: [10, 20, 30] }] }
  const model = chart({ type: 'scatter', options: {
    trendlines: [{ seriesIndex: 9, type: 'linear' }, { seriesIndex: 1, type: 'linear' }],
    errorBars: [{ seriesIndex: 1, direction: 'y', type: 'plus', valueType: 'cust', plusIndex: 2, minusIndex: 999 }],
  } })
  const plan = await prepareChartStatisticsTask(source, model)
  assert.equal(plan.preparedData, source)
  assert.equal(plan.trendlineCount, 2); assert.equal(plan.errorBarCount, 1)
  assert.deepEqual(plan.categoryDictionary, ['A', 'B'])
  assert.deepEqual([...plan.task.categoryIds], [0, 0, 1])
  assert.deepEqual(plan.task.params.series, [{ field: 1, seriesPosition: 0, column: 0, xColumn: 1 }])
  assert.deepEqual(plan.task.params.jobs.map(job => [job.kind, job.slot]), [['trendline', 1], ['errorBars', 0]])
  assert.equal(plan.task.params.jobs[1].plusColumn, 2)
  assert.equal(Object.hasOwn(plan.task.params.jobs[1], 'minusColumn'), false)
  assert.deepEqual([...plan.task.values], [2, 0, 6, 10, 20, 30, 1, 2, 3])
  assert.deepEqual([...plan.task.states], [1, 0, 1, 1, 1, 1, 1, 1, 1])
})

test('hidden auxiliary error columns are encoded as missing, while a disabled direction needs no column', async () => {
  const source = { fields: ['Label', 'Y', 'Plus', 'Minus'], records: [['A', 4, 1, 2], ['B', 8, 3, 4]], visibleFields: [0, 1, 3] }
  const model = chart({ options: { errorBars: [{ seriesIndex: 1, direction: 'y', type: 'plus', valueType: 'cust',
    plusIndex: 2, minusIndex: 3 }] } })
  const projected = projectChartRecords(source, model), plan = await prepareChartStatisticsTask(projected, model)
  assert.deepEqual(projected.records, [['A', 4, null, 2], ['B', 8, null, 4]])
  assert.equal(plan.task.params.columnCount, 2)
  assert.deepEqual([...plan.task.states], [1, 1, 0, 0])
  assert.equal(Object.hasOwn(plan.task.params.jobs[0], 'minusColumn'), false)
})

test('Top-N ties preserve source order and percentage normalization stays finite at extreme magnitudes', async () => {
  const source = { categories: ['A', 'B', 'A', 'C'], series: [{ field: 1, name: 'Value', values: [1, 5, 2, 5] }],
    records: [['A', 1], ['B', 5], ['A', 2], ['C', 5]], fields: ['Label', 'Value'] }
  const aggregated = await prepareChartComputeInputs(source, chart({ options: { aggregation: { aggregate: true,
    topN: { count: 2, includeOthers: true, othersLabel: 'Rest' } } } }))
  assert.deepEqual(aggregated.categories, ['B', 'C', 'Rest'])
  assert.deepEqual(aggregated.series[0].values, [5, 5, 3])
  assert.equal(aggregated.records, undefined)
  const percent = await prepareChartComputeInputs({ categories: ['A', 'B'], series: [
    { field: 1, name: 'First', values: [Number.MAX_VALUE, -Number.MAX_VALUE] },
    { field: 2, name: 'Second', values: [Number.MAX_VALUE, Number.MAX_VALUE] },
  ] }, chart({ type: 'columnPercentStacked' }))
  assert.deepEqual(percent.series.map(series => series.values), [[50, -50], [50, 50]])
})

test('empty or nonstatistical charts allocate no Worker task but still validate the preparation budget', async () => {
  assert.equal((await prepareChartStatisticsTask(data(), chart())).task, null)
  assert.equal((await prepareChartStatisticsTask({ categories: [], series: [] }, histogramChart())).task, null)
  await assert.rejects(() => prepareChartStatisticsTask({ categories: [], series: [] }, chart(), { memoryBudgetBytes: 0 }),
    error => error.code === 'COMPUTE_INVALID_INPUT')
})

test('histogram decoding enforces complete counts and threshold identities', async () => {
  const plan = await prepareChartStatisticsTask(data(), histogramChart())
  const decoded = await decodeChartStatistics(plan, histogramOutput())
  assert.equal(decoded.preparedData, plan.preparedData)
  assert.deepEqual(decoded.histogram, { validSampleCount: 5, bins: [
    { kind: 'underflow', upperBound: 1, count: 2 }, { kind: 'regular', lowerBound: 1, upperBound: 2, count: 1 },
    { kind: 'regular', lowerBound: 2, upperBound: 3, count: 1 }, { kind: 'overflow', lowerBound: 3, count: 1 },
  ] })
  for (const mutate of [output => { output.indices[1] = 6 }, output => { output.values[1] = 0.5 },
    output => { output.values[2] = NaN }, output => { output.segments[0].field = 2 }]) {
    const output = histogramOutput(); mutate(output)
    await assert.rejects(() => decodeChartStatistics(plan, output), error => error.code === 'COMPUTE_PROTOCOL_MISMATCH')
  }
})

test('boxplot decoding accepts authored zero whisker multiplier where min is above Q1', async () => {
  const model = chart({ type: 'boxplot', options: { boxplot: { quartileMethod: 'inclusive', whiskerMultiplier: 0 } } })
  const plan = await prepareChartStatisticsTask(data([1, 2, 3, 4]), model)
  const output = { kind: 'chartStatistics', values: new Float64Array([2, 1.75, 2.5, 3.25, 3, 1, 4]), indices: new Uint32Array(),
    segments: [segment(plan.task.params.jobs[0], { valueLength: 7, validSampleCount: 4, available: true })] }
  const decoded = await decodeChartStatistics(plan, output)
  assert.deepEqual(decoded.boxplots[0].statistics, { min: 2, q1: 1.75, median: 2.5, q3: 3.25, max: 3, outliers: [1, 4] })
})

test('an unavailable fit occupies its authored slot and is consumed without a synchronous fallback', async () => {
  const model = chart({ options: { trendlines: [{ seriesIndex: 9, type: 'linear' }, { seriesIndex: 1, type: 'linear' }] } })
  const plan = await prepareChartStatisticsTask(data([2]), model)
  const decoded = await decodeChartStatistics(plan, { kind: 'chartStatistics', values: new Float64Array(), indices: new Uint32Array(),
    segments: [segment(plan.task.params.jobs[0], { validSampleCount: 1 })] })
  assert.deepEqual(decoded.trendlines, [{ configIndex: 1, seriesPosition: 0, statistics: { points: [] } }])
  assert.deepEqual(deriveChartData(plan.preparedData, model, decoded).trends[0].statistics, { points: [] })
})

test('grouped series sharing a Y field retain their X samples and expand each authored statistical config', async () => {
  const values = [2, 4, 6], source = { categories: ['A', 'B', 'C'], series: [
    { field: 1, name: 'First group', values, xValues: [1, 2, 3], sizeValues: [1, 1, 1] },
    { field: 1, name: 'Second group', values, xValues: [10, 20, 30], sizeValues: [1, 1, 1] },
  ] }
  const model = chart({ type: 'bubble', options: {
    trendlines: [{ seriesIndex: 1, type: 'movingAvg', period: 2 }],
    errorBars: [{ seriesIndex: 1, direction: 'y', type: 'both', valueType: 'fixedVal', value: 1 }],
  } })
  const plan = await prepareChartStatisticsTask(source, model)
  assert.deepEqual(plan.task.params.series, [
    { field: 1, seriesPosition: 0, column: 0, xColumn: 1 }, { field: 1, seriesPosition: 1, column: 0, xColumn: 2 },
  ])
  assert.deepEqual(plan.task.params.jobs.map(job => [job.kind, job.slot, job.seriesIndex]),
    [['trendline', 0, 0], ['trendline', 0, 1], ['errorBars', 0, 0], ['errorBars', 0, 1]])
  assert.deepEqual([...plan.task.values], [2, 4, 6, 1, 2, 3, 10, 20, 30])
  const output = { kind: 'chartStatistics', values: new Float64Array([
    2, 3, 3, 5, 20, 3, 30, 5,
    1, 2, 1, 1, 2, 4, 1, 1, 3, 6, 1, 1,
    10, 2, 1, 1, 20, 4, 1, 1, 30, 6, 1, 1,
  ]), indices: new Uint32Array([0, 1, 2, 0, 1, 2]), segments: [
    segment(plan.task.params.jobs[0], { valueLength: 4, validSampleCount: 3, available: true }),
    segment(plan.task.params.jobs[1], { jobIndex: 1, seriesPosition: 1, valueOffset: 4, valueLength: 4, validSampleCount: 3, available: true }),
    segment(plan.task.params.jobs[2], { jobIndex: 2, valueOffset: 8, valueLength: 12, indexLength: 3, validSampleCount: 3, available: true }),
    segment(plan.task.params.jobs[3], { jobIndex: 3, seriesPosition: 1, valueOffset: 20, valueLength: 12,
      indexOffset: 3, indexLength: 3, validSampleCount: 3, available: true }),
  ] }
  const decoded = await decodeChartStatistics(plan, output), derived = deriveChartData(source, model, decoded)
  assert.deepEqual(decoded.trendlines, [
    { configIndex: 0, seriesPosition: 0, statistics: { points: [[2, 3], [3, 5]] } },
    { configIndex: 0, seriesPosition: 1, statistics: { points: [[20, 3], [30, 5]] } },
  ])
  assert.deepEqual(derived.trends.map(trend => [trend.seriesPosition, trend.statistics.points]), [[0, [[2, 3], [3, 5]]], [1, [[20, 3], [30, 5]]]])
  assert.deepEqual(derived.errors.map(error => [error.seriesPosition, error.intervals.map(interval => interval.x)]), [[0, [1, 2, 3]], [1, [10, 20, 30]]])
  const malformed = { ...output, segments: output.segments.map(segment => ({ ...segment })) }
  malformed.segments[1].seriesPosition = 0
  await assert.rejects(() => decodeChartStatistics(plan, malformed), error => error.code === 'COMPUTE_PROTOCOL_MISMATCH')
})

test('decoder refuses subviews and insufficient output memory before returning any statistics', async () => {
  const plan = await prepareChartStatisticsTask(data(), histogramChart()), output = histogramOutput()
  const parent = new Float64Array(output.values.length + 1); parent.set(output.values)
  await assert.rejects(() => decodeChartStatistics(plan, { ...output, values: parent.subarray(0, output.values.length) }),
    error => error.code === 'COMPUTE_PROTOCOL_MISMATCH')
  await assert.rejects(() => decodeChartStatistics(plan, output, { memoryBudgetBytes: 1 }),
    error => error.code === 'COMPUTE_MEMORY_LIMIT')
})

test('cancellation at a macro-task boundary closes the retained generator without publishing a prefix', async () => {
  const controller = new AbortController()
  let released = false, resumed = false
  function* retained() {
    try { yield; resumed = true; return ['partial'] } finally { released = true }
  }
  const result = consumeChartPreparation(retained(), { signal: controller.signal })
  controller.abort()
  await assert.rejects(result, error => error.code === 'COMPUTE_CANCELLED')
  assert.equal(released, true); assert.equal(resumed, false)
})
