import assert from 'node:assert/strict'
import { test } from 'vitest'
import { CHART_COMPUTE_LIMITS } from '../src/compute/chart-protocol.ts'
import { chartOutputCapacity, estimateChartTaskMemory, isChartStatisticsParams, validateChartTask } from '../src/compute/chart-tasks.ts'
import { readChartStatisticsOutput } from '../src/compute/chart-results.ts'
import { prepareChartStatisticsTask, decodeChartStatistics } from '../src/features/charts/chart-statistics-compute.ts'

// Draft protocol/source contracts for unified verification. No checks ran in
// this implementation phase; these fixtures are independent expected tuples.
function params(extra = {}) {
  return { rowCount: 2, columnCount: 4, categoryCount: 2,
    series: [{ field: 1, seriesPosition: 7, column: 0 }, { field: 2, seriesPosition: 2, column: 1 }],
    jobs: [{ kind: 'waterfallLayout', seriesIndex: 1, slot: 0, seriesIndices: [1, 0],
      stackType: 'arrangement', useSubtotal: true, totalColumns: [null, 3] }], ...extra }
}
function descriptor(p, values, indices, extra = {}) {
  const job = p.jobs[0], series = p.series[job.seriesIndex]
  return { kind: job.kind, jobIndex: 0, field: series.field, seriesPosition: series.seriesPosition, slot: job.slot,
    valueOffset: 0, indexOffset: 0, valueLength: values.length, indexLength: indices.length,
    validSampleCount: p.rowCount * job.seriesIndices.length, available: p.rowCount > 0, ...extra }
}
function output(p = params()) {
  const values = new Float64Array([7, 0, 5, 5, 5, 7, 2, 0, 7, 7,
    0, 10, 10, 10, 7, -3, 0, 7, 7, 5, 5, 7, 10, 10, 10])
  const indices = new Uint32Array([6, 3, 2, 0, 0, 0, 2, 1, 1, 0, 2, 2, 2, 3,
    7, 0, 3, 1, 7, 1, 4, 0, 7, 2, 5, 3, 0, 1, 1, 3, 3, 4])
  return { kind: 'chartStatistics', values, indices, segments: [descriptor(p, values, indices)] }
}

test('layout control accepts nonordinal source identities and validates selected anchors and total columns', () => {
  assert.equal(isChartStatisticsParams(params()), true)
  for (const patch of [
    { seriesIndices: [] }, { seriesIndices: [1, 1] }, { seriesIndices: [0, 1] }, { seriesIndices: [1, 2] },
    { totalColumns: [null] }, { totalColumns: [null, 4] }, { totalColumns: [null, -1] },
    { totalColumns: [null, undefined] }, { totalColumns: [null, Infinity] },
    { stackType: 'unknown' }, { useSubtotal: 1 }, { extra: true },
  ]) {
    const p = params(); Object.assign(p.jobs[0], patch)
    assert.equal(isChartStatisticsParams(p), false, JSON.stringify(patch))
  }
  const p = params()
  validateChartTask({ kind: 'chartStatistics', params: p, values: new Float64Array(8), states: new Uint8Array(8), categoryIds: new Uint32Array(2) })
  assert.throws(() => validateChartTask({ kind: 'chartStatistics', params: p,
    values: new Float64Array(7), states: new Uint8Array(8), categoryIds: new Uint32Array(2) }), error => error.code === 'COMPUTE_INVALID_INPUT')
})

test('sparse and inherited control entries are rejected before any indexed read or serde conversion', () => {
  for (const mutate of [
    p => { delete p.series[0] }, p => { delete p.jobs[0] },
    p => { delete p.jobs[0].seriesIndices[1] }, p => { delete p.jobs[0].totalColumns[1] },
    p => { const inherited = Object.create(Array.prototype); inherited[1] = 0;
      const selected = [1, 0]; delete selected[1]; Object.setPrototypeOf(selected, inherited); p.jobs[0].seriesIndices = selected },
    p => { p.jobs[0].seriesIndices = new Array(CHART_COMPUTE_LIMITS.maxSeries + 1) },
  ]) {
    const p = params(); mutate(p)
    assert.equal(isChartStatisticsParams(p), false)
    assert.throws(() => validateChartTask({ kind: 'chartStatistics', params: p,
      values: new Float64Array(8), states: new Uint8Array(8), categoryIds: new Uint32Array(2) }), error => error.code === 'COMPUTE_INVALID_INPUT')
  }
})

test('layout capacity and admission account for automatic points, connectors, both buffers and control arrays', () => {
  const p = params(), memory = estimateChartTaskMemory(p)
  assert.deepEqual(chartOutputCapacity(p.jobs[0], 2), { values: 25, indices: 32, scratch: 0 })
  assert.equal(memory.inputBytes, 2 * (4 * 9 + 4))
  assert.equal(memory.outputBytes, 25 * 8 + 32 * 4)
  assert.equal(memory.estimatedPeakBytes, memory.inputBytes * 2 + memory.outputBytes * 2
    + CHART_COMPUTE_LIMITS.scratchAllowance + CHART_COMPUTE_LIMITS.runtimeAllowance + 2 * 32 + 512 + 2 * 4 + 2 * 8)
  assert.deepEqual(chartOutputCapacity({ ...p.jobs[0], stackType: 'stacked' }, 2), { values: 18, indices: 24, scratch: 0 })
  assert.deepEqual(chartOutputCapacity({ ...p.jobs[0], useSubtotal: false }, 2), { values: 19, indices: 24, scratch: 0 })
  assert.deepEqual(chartOutputCapacity(p.jobs[0], 0), { values: 1, indices: 2, scratch: 0 })
})

test('descriptor checks require complete headers, independent buffers and valid counts before decoding', () => {
  const p = params(), complete = output(p)
  assert.ok(readChartStatisticsOutput(complete, p))
  for (const mutate of [
    o => { o.indices[0] = 5 }, o => { o.indices[1] = 2 }, o => { o.values[0] = NaN },
    o => { o.segments[0].validSampleCount = 5 }, o => { o.segments[0].available = false },
    o => { o.segments[0].seriesPosition = 1 }, o => { o.segments[0].valueLength-- },
    o => { o.segments[0].fit = { coefficients: [1, 2], xScale: 1, xMean: 0, xSpread: 1, yScale: 1, yMean: 0 } },
    o => { delete o.segments[0] },
    o => { const padded = new Float64Array(o.values.length + 1); padded.set(o.values); o.values = padded.subarray(0, o.values.length) },
    o => { const backing = new ArrayBuffer(o.values.byteLength + o.indices.byteLength);
      const values = new Float64Array(backing, 0, o.values.length), indices = new Uint32Array(backing, o.values.byteLength, o.indices.length);
      values.set(o.values); indices.set(o.indices); o.values = values; o.indices = indices },
  ]) {
    const malformed = output(p); mutate(malformed)
    assert.equal(readChartStatisticsOutput(malformed, p), null)
  }
  const empty = params({ rowCount: 0, categoryCount: 0 }), values = new Float64Array([0]), indices = new Uint32Array([0, 0])
  const emptyOutput = { kind: 'chartStatistics', values, indices, segments: [descriptor(empty, values, indices)] }
  assert.ok(readChartStatisticsOutput(emptyOutput, empty))
  emptyOutput.values[0] = 1
  assert.equal(readChartStatisticsOutput(emptyOutput, empty), null)
})

function chart(options = {}) {
  return { id: 'waterfall', sheetId: 'sheet', type: 'waterfall', revision: 0, source: { orientation: 'columns',
    range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 } }, position: { x: 0, y: 0 }, size: { width: 420, height: 280 },
    categoryField: 0, valueFields: [1, 2], xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' },
    legend: { visible: true, position: 'bottom' }, options }
}
function source() {
  return { categories: ['A', 'B'], fields: ['Label', 'First', 'Second'], records: [['A', 10, 5], ['B', -3, 2]],
    series: [{ field: 1, name: 'First', values: [10, -3] }, { field: 2, name: 'Second', values: [5, 2] }] }
}
function plannedOutput(p) {
  const values = new Float64Array([7, 0, 10, 10, 10, 7, -3, 0, 7, 7,
    0, 5, 5, 5, 7, 2, 0, 7, 7, 10, 10, 7, 5, 5, 5])
  const indices = new Uint32Array([6, 3, 0, 0, 0, 0, 0, 1, 1, 0, 0, 2, 2, 3,
    1, 0, 3, 1, 1, 1, 4, 0, 1, 2, 5, 3, 0, 1, 1, 3, 3, 4])
  return { kind: 'chartStatistics', values, indices, segments: [descriptor(p, values, indices)] }
}

test('one layout job covers all series and source roles without retaining transferred input views', async () => {
  const model = chart({ waterfall: { pointRoles: [{ seriesIndex: 2, dataPointIndex: 0, role: 'total' }] } })
  const plan = await prepareChartStatisticsTask(source(), model), p = plan.task.params
  assert.equal(p.jobs.length, 1)
  assert.deepEqual(p.jobs[0], { kind: 'waterfallLayout', seriesIndex: 0, seriesIndices: [0, 1], slot: 0,
    stackType: 'arrangement', useSubtotal: true, totalColumns: [null, 2] })
  assert.deepEqual([...plan.task.values], [10, -3, 5, 2, 1, 0])
  const result = plannedOutput(p)
  structuredClone(plan.task, { transfer: [plan.task.values.buffer, plan.task.states.buffer, plan.task.categoryIds.buffer] })
  assert.equal(plan.task.values.byteLength, 0)
  const decoded = (await decodeChartStatistics(plan, result)).waterfallLayout
  assert.equal(decoded.cumulative, 7)
  assert.deepEqual(decoded.points.filter(p => p.synthetic).map(p => [p.seriesPosition, p.dataPointIndex, p.value]), [[0, 2, 7], [1, 2, 7]])
  assert.deepEqual(decoded.connectors, [
    { from: { categoryIndex: 0, value: 10 }, to: { categoryIndex: 1, value: 10 } },
    { from: { categoryIndex: 1, value: 7 }, to: { categoryIndex: 3, value: 5 } },
    { from: { categoryIndex: 3, value: 5 }, to: { categoryIndex: 4, value: 5 } },
  ])
})

test('malformed roles, source values, subtotal flags, geometry and connectors reject the entire decoded result', async () => {
  const model = chart({ waterfall: { pointRoles: [{ seriesIndex: 2, dataPointIndex: 0, role: 'total' }] } })
  const plan = await prepareChartStatisticsTask(source(), model), p = plan.task.params
  for (const mutate of [
    o => { o.indices[2] = 1 }, o => { o.indices[3] = 2 }, o => { o.indices[4] = 1 },
    o => { o.indices[5] = 4 }, o => { o.indices[13] = 1 }, o => { o.indices[17] = 0 },
    o => { o.values[3] = 11 }, o => { o.values[4] = 9 }, o => { o.values[9] = 6 },
    o => { o.values[0] = 8 }, o => { o.values[20] = 11 }, o => { o.indices[28] = 2 },
    o => { o.segments[0].validSampleCount = 3 }, o => { o.values[2] = Infinity },
  ]) {
    const malformed = plannedOutput(p); mutate(malformed)
    await assert.rejects(() => decodeChartStatistics(plan, malformed), error => error.code === 'COMPUTE_PROTOCOL_MISMATCH')
  }
})
