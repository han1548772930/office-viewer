import assert from 'node:assert/strict'
import { test, vi } from 'vitest'
import { SlicerMemberInputBuilder } from '../src/features/slicer/slicer-member-input.ts'
import { SlicerMemberQueryTransport } from '../src/features/slicer/slicer-member-transport.ts'
import { prepareSlicerTextRanks } from '../src/compute/slicer-member.worker.ts'
import { cloneTask, estimateTaskMemory, taskViews, transferTaskBuffers, validateTask } from '../src/compute/tasks.ts'
import { readSlicerMembersOutput } from '../src/compute/slicer-member-results.ts'
import { readResult, progressUnits } from '../src/compute/guards.ts'
import { initializeComputeWasm } from '../src/compute/wasm.worker.ts'
import { OfficeComputeClient } from '../src/compute/client.ts'
import { createWasmTaskControl } from '../src/platform/wasm/task-protocol.ts'
import { createSlicerMemberTestClient, decodeSlicerMemberTask, decodeSlicerMemberCaptions, slicerComputeFixtureResult } from './fixtures/slicer-member-compute-fixture.mjs'

// Contract drafts only: no tests, workers, builds or browsers executed yet.
const defaults = { mode: 'items', start: 0, count: 100, locale: 'en-US', sortOrder: 'ascending',
  customListSort: false, crossFilter: 'showItemsWithNoData', hideItemsWithNoData: false }
const rows = [{ type: 'blank' },{ type: 'number', value: 1 },{ type: 'date', value: 1 },{ type: 'string', value: '1' },
  { type: 'boolean', value: true },{ type: 'string', value: 'true' },{ type: 'error', value: '#N/A' },
  { type: 'string', value: '#N/A' },{ type: 'string', value: '' },{ type: 'number', value: 1 }]
async function input(values = rows, options = {}, availability = values.map(() => true), budget, canvasSources) {
  const builder = new SlicerMemberInputBuilder(values.length, { ...defaults, ...options }, budget)
  values.forEach((value, index) => builder.append(value, availability[index], canvasSources?.[index]))
  return builder.finish(async () => {})
}
const signal = () => new AbortController().signal

test('input preserves all six typed kinds and exact raw values while only dictionary text is interned', async () => {
  const task = await input()
  assert.deepEqual(decodeSlicerMemberTask(task), rows)
  assert.equal(task.params.rowCount, 10); assert.equal(task.params.dictionaryCount, 4)
  assert.equal(task.textIds[6], task.textIds[7]); assert.notEqual(task.states[6], task.states[7])
  assert.equal(task.textIds[3], 0); assert.equal(task.textIds[8], 3)
  assert.equal(taskViews(task).length, 7)
  assert.doesNotThrow(() => validateTask(task))
  assert.equal(taskViews(task).reduce((total, view) => total + view.byteLength, 0), estimateTaskMemory(task).inputBytes)
})

test('UTF-16LE preserves embedded nulls, astral pairs and distinct lone surrogates', async () => {
  const values = ['a\0b','😀','\ud800','\ud801',''].map(value => ({ type: 'string', value }))
  const task = await input(values)
  assert.deepEqual(decodeSlicerMemberTask(task), values)
  assert.equal(task.params.dictionaryCount, 5)
  await prepareSlicerTextRanks(task, async () => {})
  assert.equal(decodeSlicerMemberTask(task)[2].value, '\ud800')
})

test('metadata encodes only typed Date rows and preserves exact fractional serials', async () => {
  const task = await input([{ type: 'number', value: -1000 },{ type: 'string', value: '1900-01-01' },
    { type: 'date', value: -2.25 },{ type: 'date', value: 10.5 },{ type: 'error', value: '#N/A' }], { mode: 'metadata' })
  assert.equal(task.params.dictionaryCount, 0); assert.equal(task.textBytes.length, 0)
  assert.deepEqual(task.states, new Uint8Array([0,0,2,2,0]))
  assert.deepEqual(task.values, new Float64Array([0,0,-2.25,10.5,0]))
  assert.deepEqual(progressUnits(task), { total: 5, unit: 'rows' })
})

for (const locale of ['en-US','zh-CN','de-DE']) test(`${locale} rank preparation retains current numeric/base and optional calendar aliases`, async () => {
  const month = new Intl.DateTimeFormat(locale, { month: 'long', timeZone: 'UTC' }), short = new Intl.DateTimeFormat(locale, { month: 'short', timeZone: 'UTC' })
  const values = ['item10','item2','ALPHA','alpha',month.format(new Date(Date.UTC(2020,1,1))),
    short.format(new Date(Date.UTC(2020,0,1))),month.format(new Date(Date.UTC(2020,0,1)))].map(value => ({ type: 'string', value }))
  const task = await input(values, { locale, customListSort: true })
  await prepareSlicerTextRanks(task, async () => {})
  assert.ok(task.textRanks[task.textIds[1]] < task.textRanks[task.textIds[0]])
  assert.equal(task.textRanks[task.textIds[2]], task.textRanks[task.textIds[3]])
  assert.equal(task.textRanks[task.textIds[5]], task.textRanks[task.textIds[6]])
  assert.ok(task.textRanks[task.textIds[5]] < task.textRanks[task.textIds[4]])
})

test('rank preparation yields during long UTF-16 decoding and bounded merge writes', async () => {
  const task = await input(Array.from({ length: 2050 }, (_, index) => ({ type: 'string', value: `${2050-index}:${'x'.repeat(40)}` })))
  let checkpoints = 0
  await prepareSlicerTextRanks(task, async () => { checkpoints++ })
  assert.ok(checkpoints > 2)
  let stopped = false
  await assert.rejects(() => prepareSlicerTextRanks(task, async () => { stopped = true; throw new Error('cancel at real checkpoint') }), /real checkpoint/)
  assert.equal(stopped, true)
})

test('shape and budget failures reject before transfer, and builder consumption is explicit', async () => {
  const task = await input()
  for (const mutate of [value => value.params.rowCount++,value => value.params.dictionaryCount++,
    value => value.params.start = Number.MAX_SAFE_INTEGER+1,value => value.params.count = 0,
    value => value.params.locale = '',value => value.params.crossFilter = { toString: () => 'none' },
    value => value.params.extra = true,value => value.availability = new Uint8Array(1)]) {
    const candidate = cloneTask(task); mutate(candidate)
    assert.throws(() => validateTask(candidate), error => error.code === 'COMPUTE_INVALID_INPUT')
    assert.ok(candidate.values.byteLength)
  }
  const estimate = estimateTaskMemory(task)
  assert.throws(() => validateTask(task, estimate.estimatedPeakBytes-1), error => error.code === 'COMPUTE_MEMORY_LIMIT')
  assert.throws(() => new SlicerMemberInputBuilder(100, defaults, 1024), error => error.code === 'COMPUTE_MEMORY_LIMIT')
  const builder = new SlicerMemberInputBuilder(1, defaults)
  await assert.rejects(() => builder.finish(async () => {}), error => error.code === 'COMPUTE_INVALID_INPUT')
  builder.append({ type: 'blank' }); await builder.finish(async () => {})
  assert.throws(() => builder.append({ type: 'blank' }), error => error.code === 'COMPUTE_INVALID_INPUT')
  await assert.rejects(() => builder.finish(async () => {}), error => error.code === 'COMPUTE_INVALID_INPUT')
})

test('exclusive typed buffers transfer once and detached or aliased inputs are rejected', async () => {
  const task = await input(), copy = cloneTask(task)
  const transferred = structuredClone(task, { transfer: transferTaskBuffers(task) })
  assert.equal(task.values.buffer.byteLength, 0); assert.equal(task.textOffsets.buffer.byteLength, 0)
  assert.deepEqual(decodeSlicerMemberTask(transferred), rows)
  assert.deepEqual(decodeSlicerMemberTask(copy), rows)
  assert.throws(() => validateTask(task), error => error.code === 'COMPUTE_INVALID_INPUT')
  const aliased = cloneTask(copy); aliased.availability = aliased.states
  assert.throws(() => transferTaskBuffers(aliased), error => error.code === 'COMPUTE_INVALID_INPUT')
})

test('output guards accept only the bounded projected window with owned unique row offsets and real date bounds', async () => {
  const task = await input(), expected = slicerComputeFixtureResult(task, { documentId: 'book', modelRevision: 4 }, 'fixture-1').output
  assert.ok(readSlicerMembersOutput(expected, task.params))
  for (const mutate of [value => value.total = 11,value => value.indices[0] = 99,value => value.indices[1] = value.indices[0],
    value => value.hasData[0] = 2,value => value.dateCount = 0,value => value.minDate = 2,value => value.maxDate = Infinity,
    value => value.indices = new Uint32Array(value.indices.buffer, 4)]) {
    const candidate = structuredClone(expected); mutate(candidate)
    assert.equal(readSlicerMembersOutput(candidate, task.params), null)
  }
})

test('empty member windows require real independent buffers and reject detached empty views', async () => {
  const task = await input([], { mode: 'metadata' })
  const output = { kind: 'slicerMembers', indices: new Uint32Array(), hasData: new Uint8Array(),
    total: 0, dateCount: 0, minDate: null, maxDate: null }
  assert.ok(readSlicerMembersOutput(output, task.params))
  structuredClone(output, { transfer: [output.indices.buffer, output.hasData.buffer] })
  assert.equal(readSlicerMembersOutput(output, task.params), null)
})

test('client result guards include all seven input copies and only two effective window output copies', async () => {
  const task = await input(), options = { documentId: 'book', modelRevision: 4 }, requestId = 'fixture-1'
  const response = slicerComputeFixtureResult(task, options, requestId), control = response.context
  assert.ok(readResult(response, { ...options, requestId }, control, task, estimateTaskMemory(task), task.values.length))
  for (const mutate of [value => value.metrics.boundary.inputCopies--,value => value.metrics.boundary.inputBytes--,
    value => value.metrics.boundary.outputCopies++,value => value.metrics.boundary.outputBytes++,
    value => value.metrics.boundary.rustBufferClones++,value => value.metrics.transport.inputTransfers--,
    value => value.metrics.transport.outputTransfers--,value => value.metrics.transport.structuredResultMessages = 0]) {
    const candidate = structuredClone(response); mutate(candidate)
    assert.equal(readResult(candidate, { ...options, requestId }, control, task, estimateTaskMemory(task), task.values.length), null)
  }
})

test('transport enforces source revision, consumes exclusive inputs and keeps injected shared clients alive on dispose', async () => {
  const client = createSlicerMemberTestClient(), transport = new SlicerMemberQueryTransport(client), task = await input()
  const result = await transport.query(task, { unitId: 'book', baseRevision: 4, signal: signal(), timeoutMs: 1000 })
  assert.equal(task.values.buffer.byteLength, 0); assert.equal(result.output.total, 9)
  assert.equal(client.submitted[0].options.ownership, 'transfer'); assert.equal(client.submitted[0].options.priority, 'interactive')
  transport.dispose()
  assert.equal(client.isCurrent({ documentId: 'book', modelRevision: 4 }), true)
  const after = await input()
  await assert.rejects(() => transport.query(after, { unitId: 'book', baseRevision: 4, signal: signal(), timeoutMs: 1000 }), error => error.code === 'COMPUTE_DISPOSED')
  client.dispose()
})

test('loading rejects protocol-2 artifacts that omit the new typed member binding', async () => {
  const module = { default: async () => ({ memory: new WebAssembly.Memory({ initial: 1 }) }), compute_protocol_version: () => 2,
    compute_version: () => 'fixture', compute_yield: async () => {}, create_numeric_aggregate() {},
    create_numeric_sort() {}, create_numeric_filter() {}, create_chart_statistics() {} }
  await assert.rejects(() => initializeComputeWasm(module), error => error.code === 'COMPUTE_WASM_API_MISSING')
  assert.equal((await initializeComputeWasm({ ...module, create_slicer_members() {} })).module.compute_protocol_version(), 2)
})

class WorkerPort {
  onmessage = null; onerror = null; onmessageerror = null; messages = []; terminated = false
  postMessage(message, transfer = []) { this.messages.push(structuredClone(message, { transfer })) }
  terminate() { this.terminated = true }
  send(data) { this.onmessage?.({ data }) }
}
function workerResult(request) {
  const response = slicerComputeFixtureResult(request.task, { documentId: request.context.unitId, modelRevision: request.context.baseRevision }, request.context.requestId)
  return { ...response, context: request.context }
}
function ready(worker) { worker.send({ type: 'ready', protocolVersion: 2, engineVersion: 'test-port-no-wasm' }) }

test('real compute client reuses one Worker across member queries and rejects late results after cancellation', async () => {
  const workers = [], client = new OfficeComputeClient({ workerFactory: () => { const worker = new WorkerPort(); workers.push(worker); return worker }, wasmModuleUrl: 'https://office.test/compute.js' })
  try {
    const first = client.submit(await input(), { documentId: 'book', modelRevision: 4, ownership: 'transfer' })
    const worker = workers[0]; ready(worker); const oldRequest = worker.messages.at(-1)
    const rejected = assert.rejects(first.promise, error => error.code === 'COMPUTE_CANCELLED'); first.cancel(); await rejected
    const next = client.submit(await input(), { documentId: 'book', modelRevision: 4, ownership: 'transfer' })
    worker.send({ type: 'cancelled', context: oldRequest.context }); const nextRequest = worker.messages.at(-1)
    worker.send(workerResult(oldRequest)); worker.send(workerResult(nextRequest))
    assert.equal((await next.promise).output.total, 9); assert.equal(workers.length, 1); assert.equal(worker.terminated, false)
  } finally { client.dispose() }
})

test('source revision change rejects old work and an unresponsive member Worker is killed before queued work resumes', async () => {
  vi.useFakeTimers()
  const workers = [], client = new OfficeComputeClient({ workerFactory: () => { const worker = new WorkerPort(); workers.push(worker); return worker }, wasmModuleUrl: 'https://office.test/compute.js', cancellationGraceMs: 20 })
  try {
    const first = client.submit(await input(), { documentId: 'book', modelRevision: 4, ownership: 'transfer' })
    const worker = workers[0]; ready(worker); const oldRequest = worker.messages.at(-1), oldHandler = worker.onmessage
    const rejected = assert.rejects(first.promise, error => error.code === 'COMPUTE_STALE_RESULT'); client.setDocumentRevision('book', 5); await rejected
    const next = client.submit(await input(), { documentId: 'book', modelRevision: 5, ownership: 'transfer' })
    await vi.advanceTimersByTimeAsync(20)
    assert.equal(worker.terminated, true); assert.equal(workers.length, 2)
    oldHandler({ data: workerResult(oldRequest) }); ready(workers[1]); workers[1].send(workerResult(workers[1].messages.at(-1)))
    assert.equal((await next.promise).modelRevision, 5)
  } finally { client.dispose(); vi.useRealTimers() }
})

test('Canvas shares seven buffers while retaining raw error identity separately from every source caption and selection bit', async () => {
  const task = await input([{ type: 'error', value: '#N/A' }], { mode: 'canvas' }, [false], undefined, [{ caption: '0', selected: true }])
  assert.equal(task.params.dictionaryCount, 2); assert.equal(task.textIds.length, 2); assert.equal(task.availability[0], 2)
  assert.deepEqual(decodeSlicerMemberTask(task), [{ type: 'error', value: '#N/A' }]); assert.deepEqual(decodeSlicerMemberCaptions(task), ['0'])
  assert.equal(taskViews(task).length, 7); assert.doesNotThrow(() => validateTask(task))
  assert.equal(estimateTaskMemory(task).inputBytes, taskViews(task).reduce((sum, view) => sum + view.byteLength, 0))
  const copy = cloneTask(task)
  assert.deepEqual(decodeSlicerMemberTask(copy), decodeSlicerMemberTask(task)); assert.deepEqual(decodeSlicerMemberCaptions(copy), ['0'])
  for (let index = 0; index < 7; index++) assert.notEqual(taskViews(copy)[index].buffer, taskViews(task)[index].buffer)
  const short = cloneTask(task); short.textIds = new Uint32Array([short.textIds[0]])
  assert.throws(() => validateTask(short), error => error.code === 'COMPUTE_INVALID_INPUT')
  for (const mode of ['items', 'metadata']) {
    const mismatched = cloneTask(task); mismatched.params.mode = mode
    assert.throws(() => validateTask(mismatched), error => error.code === 'COMPUTE_INVALID_INPUT')
  }
  const builder = new SlicerMemberInputBuilder(1, { ...defaults, mode: 'canvas' })
  assert.throws(() => builder.append({ type: 'blank' }), error => error.code === 'COMPUTE_INVALID_INPUT')
})

test('Canvas output guards accept packed full/partial states and reject contradictory states or missing available bits', async () => {
  const task = await input([{ type: 'number', value: 1 },{ type: 'string', value: '1' }], { mode: 'canvas' }, [true,false],
    undefined, [{ caption: '1', selected: true },{ caption: '1', selected: false }])
  const output = slicerComputeFixtureResult(task, { documentId: 'canvas', modelRevision: 0 }, 'canvas-1').output
  assert.equal(output.total,1); assert.deepEqual(output.hasData,new Uint8Array([5])); assert.ok(readSlicerMembersOutput(output,task.params))
  for (const flags of [6,7,8,255]) assert.equal(readSlicerMembersOutput({ ...output,hasData:new Uint8Array([flags]) },task.params),null)
  assert.equal(readSlicerMembersOutput({ ...output,hasData:new Uint8Array([4]) },{ ...task.params,hideItemsWithNoData:true }),null)
  assert.equal(readSlicerMembersOutput(output,{ ...task.params,mode:'items' }),null)
})
