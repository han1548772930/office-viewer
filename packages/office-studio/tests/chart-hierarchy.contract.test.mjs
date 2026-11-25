import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'vitest'
import { chartHierarchyData, chartHierarchyNumber, chartHierarchyResultAsync, chartHierarchyTokenKey, createChartHierarchyPlanSteps } from '../src/features/charts/chart-hierarchy-data.ts'
import { decodeChartHierarchySteps, executeChartHierarchyTaskSteps } from '../src/features/charts/chart-hierarchy-compute.ts'
import { chartHierarchySourceFormatter } from '../src/features/charts/chart-hierarchy-format.ts'
import { consumeChartPreparationSync } from '../src/features/charts/chart-compute-lifecycle.ts'
import { projectChartRecords, projectChartRecordsAsync } from '../src/features/charts/chart-record-projection.ts'
import { estimateChartTaskMemory, isChartStatisticsParams, validateChartTask } from '../src/compute/chart-tasks.ts'
import { readChartStatisticsOutput } from '../src/compute/chart-results.ts'
import { TREEMAP_RECORDS } from './fixtures/chart-treemap-fixtures.mjs'

// P26 contract drafts. No test, compiler, Worker, WASM, browser or benchmark ran.
const shared = JSON.parse(readFileSync(new URL('../../../tests/fixtures/chart-hierarchy.json', import.meta.url), 'utf8')).cases
const consume = steps => consumeChartPreparationSync(steps)
const binding = (depth = 2, valueField = depth) => ({ type: 'treemap', categoryField: 0, valueFields: [valueField], options: { mapping: { hierarchyIndexes: Array.from({ length: depth }, (_, index) => index) } } })
const source = (records = TREEMAP_RECORDS) => ({ fields: ['Region', 'Category', 'Weight'], records })
function numericFixture(fixture) {
  const rows = fixture.rows.length, columns = fixture.depth * 2 + 1, values = new Float64Array(rows * columns), states = new Uint8Array(values.length)
  fixture.rows.forEach((row, r) => {
    if (row.weight !== null) { values[r] = row.weight; states[r] = 1 }
    row.keys.forEach((key, depth) => {
      if (key === null) return
      for (const column of [depth + 1, fixture.depth + depth + 1]) { const offset = column * rows + r; values[offset] = key; states[offset] = 1 }
    })
  })
  return { kind: 'chartStatistics', values, states, categoryIds: new Uint32Array(rows), params: { rowCount: rows, columnCount: columns, categoryCount: rows ? 1 : 0,
    series: [{ field: fixture.depth, seriesPosition: 0, column: 0 }], jobs: [{ kind: 'hierarchy', seriesIndex: 0, slot: 0,
      hierarchyColumns: Array.from({ length: fixture.depth }, (_, depth) => depth + 1), labelColumns: Array.from({ length: fixture.depth }, (_, depth) => fixture.depth + depth + 1),
      keyCount: fixture.keyCount, labelCount: fixture.keyCount }] } }
}
function copy(output) { return { ...output, values: new Float64Array(output.values), indices: new Uint32Array(output.indices), segments: output.segments.map(segment => ({ ...segment })) } }
function prepare(records = source(), chart = binding()) {
  const plan = consume(createChartHierarchyPlanSteps(records, chart)), output = consume(executeChartHierarchyTaskSteps(plan.task))
  return { plan, output, result: consume(decodeChartHierarchySteps(plan, output)) }
}

for (const fixture of shared) test(`portable numeric hierarchy and the shared Rust fixture: ${fixture.name}`, () => {
  const task = numericFixture(fixture), output = consume(executeChartHierarchyTaskSteps(task))
  assert.deepEqual(Array.from(output.values), fixture.values); assert.deepEqual(Array.from(output.indices), fixture.indices)
  assert.equal(output.segments[0].validSampleCount, fixture.validSampleCount); assert.equal(output.segments[0].available, fixture.available)
  assert.ok(readChartStatisticsOutput(output, task.params))
})

test('typed identities distinguish equal labels, -0, whitespace, booleans and dates, while numeric strings stay invalid weights', () => {
  const records = [[1, 3], ['1', 2], [-0, 4], [0, 5], [true, 6], ['true', 7], [' ', 8], ['', 9], [new Date(1234), 10], [new Date(5678), 11], ['Ignored', '100']]
  const input = { fields: ['Name', 'Weight'], records, formattedRecords: records.map(([name, value]) => [name instanceof Date ? 'Same date label' : String(name), String(value)]) }
  const nodes = chartHierarchyData(input, binding(1)), values = new Map(nodes.map(node => [node.id, node.value]))
  assert.equal(values.get(JSON.stringify(['number:-0'])), 4); assert.equal(values.get(JSON.stringify(['number:0'])), 5)
  assert.equal(values.get(JSON.stringify(['number:1'])), 3); assert.equal(values.get(JSON.stringify(['string:"1"'])), 2)
  assert.equal(values.get(JSON.stringify(['boolean:true'])), 6); assert.equal(values.get(JSON.stringify(['string:"true"'])), 7)
  assert.equal(values.get(JSON.stringify(['string:" "'])), 17)
  assert.equal(nodes.filter(node => node.name === 'Same date label').length, 2)
  assert.equal(values.has(JSON.stringify(['string:"Ignored"'])), false)
  assert.equal(chartHierarchyNumber('3'), null); assert.equal(chartHierarchyNumber(Infinity), null); assert.equal(chartHierarchyTokenKey(-0), 'number:-0')
})

test('blank carry follows ancestor scope and invalid-weight rows update later scope before they are rejected', () => {
  const records = [[null, 'Bad', 'Leaf', 10], ['A', 'B', 'L1', 5], [null, null, 'L2', 7], ['C', null, 'Leak', 10],
    [null, 'D', null, 3], ['C', 'D', 'Seed', 'invalid'], [null, null, null, 4], ['C', 'E', null, -2]]
  const data = chartHierarchyData({ fields: ['Root', 'Group', 'Leaf', 'Weight'], records }, binding(3))
  assert.deepEqual(data.map(node => [node.name, node.value]), [['A', 12], ['C', 7]])
  assert.deepEqual(data[0].children[0].children.map(node => [node.name, node.value]), [['L2', 7], ['L1', 5]])
  assert.equal(data[1].children[0].selfValue, 3); assert.deepEqual(data[1].children[0].children[0].path, ['C', 'D', 'Seed'])
  assert.equal(data[1].children[0].children[0].value, 4)
})

test('signed duplicate paths retain parent-owned values, prune non-positive descendants and decouple root palette order from weight order', () => {
  const { result } = prepare(), [north, south] = result.nodes
  assert.deepEqual(result.nodes.map(node => [node.name, node.value, node.selfValue, node.paletteIndex]), [['North', 13, -5, 1], ['South', 9, 3, 0]])
  assert.deepEqual(north.children.map(node => [node.name, node.value]), [['Leaf B', 12], ['Leaf C', 6]])
  assert.deepEqual(south.children.map(node => [node.name, node.value]), [['Leaf A', 6]])
  assert.deepEqual(result.paletteOrder, [south.id, north.id])
  const empty = prepare(source([['A', 'B', -2], [null, null, 2], ['C', null, 0]]))
  assert.deepEqual(empty.result.nodes, []); assert.match(empty.result.diagnostic, /最终为正/)
})

test('source number inference retains decimal grouping, unit spaces and percentages, and applies to signed self totals', () => {
  assert.equal(chartHierarchySourceFormatter('$1,234.50 USD')(12345.25), '$12,345.25 USD')
  assert.equal(chartHierarchySourceFormatter('1\u202f234,50 EUR')(12345.25), '12\u202f345,25 EUR')
  assert.equal(chartHierarchySourceFormatter('12.50%')(0.375), '37.50%')
  assert.equal(chartHierarchySourceFormatter('1e+20'), undefined)
  const input = source(), formattedRecords = input.records.map(row => row.map((value, index) => index === 2 ? `$${Number(value).toFixed(2)} USD` : String(value ?? '')))
  const { result } = prepare({ ...input, formattedRecords })
  assert.equal(result.nodes[0].formattedValue, '$13.00 USD'); assert.equal(result.nodes[0].formattedSelfValue, '$-5.00 USD')
})

test('hierarchy defaults select one strict numeric field and disjoint ordered roles reject malformed authoring', () => {
  const input = { fields: ['Root', 'Leaf', 'String number', 'Weight', 'Other'], records: [['A', 'B', '20', 3, 4]] }
  const chart = { ...binding(), valueFields: [] }, projected = projectChartRecords(input, chart)
  assert.deepEqual(projected.series.map(series => series.field), [3]); assert.deepEqual(projected.sourceNumericFields, [3, 4])
  for (const options of [{ hierarchyIndexes: [] }, { hierarchyIndexes: [0, 0] }, { hierarchyIndexes: [0, 2] }]) {
    assert.throws(() => prepare(source(), { ...binding(), options: { mapping: options } }))
  }
  assert.throws(() => prepare(source(), { ...binding(), valueFields: [1, 2] }))
})

test('materialized decoder rejects bad identity, parent preorder, totals, source creation order and root palette tuples', () => {
  const { plan, output } = prepare(), bad = [
    candidate => { candidate.indices[2] = 0 },
    candidate => { candidate.indices[3] = plan.keys.length },
    candidate => { candidate.indices[4] = plan.labels.length },
    candidate => { candidate.indices[5] = 0xffff_ffff },
    candidate => { candidate.values[0]++ },
    candidate => { candidate.indices[candidate.indices.length - 1] = candidate.indices[candidate.indices.length - 2] },
    candidate => { candidate.segments[0].validSampleCount = 0 },
    candidate => { candidate.segments[0].available = false },
  ]
  for (const mutate of bad) { const candidate = copy(output); mutate(candidate); assert.throws(() => consume(decodeChartHierarchySteps(plan, candidate)), error => error.code === 'COMPUTE_PROTOCOL_MISMATCH') }
  const parameters = structuredClone(plan.task.params); parameters.jobs[0].labelColumns[0] = parameters.jobs[0].hierarchyColumns[0]
  assert.equal(isChartStatisticsParams(parameters), false)
  assert.throws(() => consume(decodeChartHierarchySteps({ ...plan, labels: plan.labels.slice(1) }, output)), error => error.code === 'COMPUTE_PROTOCOL_MISMATCH')
})

test('only numeric buffers transfer, retained dictionaries decode after detachment and the consumed input cannot be reused', async () => {
  const plan = consume(createChartHierarchyPlanSteps(source(), binding())), keys = [...plan.keys], labels = [...plan.labels]
  const moved = structuredClone(plan.task, { transfer: [plan.task.values.buffer, plan.task.states.buffer, plan.task.categoryIds.buffer] })
  assert.equal(plan.task.values.byteLength, 0); assert.equal(plan.task.states.byteLength, 0); assert.equal(plan.task.categoryIds.byteLength, 0)
  assert.deepEqual(plan.keys, keys); assert.deepEqual(plan.labels, labels)
  const output = consume(executeChartHierarchyTaskSteps(moved)), decoded = consume(decodeChartHierarchySteps(plan, output))
  assert.equal(decoded.nodes[0].value, 13); assert.throws(() => validateChartTask(plan.task))
  let submits = 0
  const result = await projectChartRecordsAsync(source(), binding(), { runChartTask: async task => { submits++; assert.equal(task.params.jobs[0].kind, 'hierarchy'); return consume(executeChartHierarchyTaskSteps(task)) } })
  assert.equal(submits, 1); assert.equal(result.hierarchy[0].value, 13)
})

test('retained source admission, cancellation and revision checks prevent a partial or stale hierarchy result', async () => {
  const plan = consume(createChartHierarchyPlanSteps(source(), binding())), bytes = estimateChartTaskMemory(plan.task.params)
  assert.ok(bytes.estimatedPeakBytes - bytes.inputBytes * 2 - bytes.outputBytes * 2 >= source().records.length * 2 * 256)
  assert.throws(() => consumeChartPreparationSync(createChartHierarchyPlanSteps(source(), binding(), { memoryBudgetBytes: 1 }), { memoryBudgetBytes: 1 }), error => error.code === 'COMPUTE_MEMORY_LIMIT')
  const controller = new AbortController(), large = source(Array.from({ length: 10000 }, (_, index) => ['A', 'B' + index, 1]))
  const pending = chartHierarchyResultAsync(large, binding(), { signal: controller.signal })
  setTimeout(() => controller.abort(), 0)
  await assert.rejects(pending, error => error.code === 'COMPUTE_CANCELLED')
  await assert.rejects(() => chartHierarchyResultAsync(source(), binding(), { assertCurrent: () => { throw Object.assign(new Error('new revision'), { code: 'COMPUTE_STALE_RESULT' }) } }), error => error.code === 'COMPUTE_STALE_RESULT')
})
