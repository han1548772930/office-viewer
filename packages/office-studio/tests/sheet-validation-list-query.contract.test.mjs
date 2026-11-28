import assert from 'node:assert/strict'
import { afterEach, test, vi } from 'vitest'
import { OfficeComputeClient } from '../src/compute/client.ts'
import { ComputeError } from '../src/compute/errors.ts'
import { cloneTask, estimateTaskMemory, taskViews, transferTaskBuffers, validateTask } from '../src/compute/tasks.ts'
import { progressUnits, readResult } from '../src/compute/guards.ts'
import { readValidationListQueryOutput } from '../src/compute/validation-list-query-results.ts'
import { createWasmTaskControl } from '../src/platform/wasm/task-protocol.ts'
import { filterSheetValidationListOptions, prepareSheetValidationListQuery, querySheetValidationList } from '../src/features/sheet-validation/validation-list-query.ts'

// Protocol/encoding/window drafts. The independent JS substring oracle below
// exists only in this test adapter, never as a second production authority.
afterEach(() => { vi.useRealTimers(); vi.restoreAllMocks() })
const options = values => Object.freeze(values.map((label, index) => Object.freeze({ value: label, label, ...(index === 1 ? { color: '#008000' } : {}) })))
function utf16(value) {
  const bytes = new Uint8Array(value.length * 2), view = new DataView(bytes.buffer)
  for (let index = 0; index < value.length; index++) view.setUint16(index * 2, value.charCodeAt(index), true)
  return bytes
}
function text(bytes, start = 0, end = bytes.length) {
  const units = [], view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let position = start; position < end; position += 2) units.push(view.getUint16(position, true))
  return String.fromCharCode(...units)
}
const task = () => ({ kind: 'validationListQuery', offsets: new Uint32Array([0, 4, 8, 10]), textBytes: utf16('abbca'), queryBytes: utf16('a'),
  params: { rowCount: 3, textByteLength: 10, queryByteLength: 2, start: 0, count: 3 } })
function packet(input, control, indices = new Uint32Array([0, 2]), total = 2) {
  const memory = estimateTaskMemory(input), length = indices.byteLength
  return { type: 'result', context: control, engineVersion: 'list-query-protocol-fixture', output: { kind: 'validationListQuery', indices, total },
    metrics: { elapsedMs: 1, ...memory, boundary: { inputCopies: 1 + Number(input.textBytes.length > 0) + Number(input.queryBytes.length > 0),
      inputBytes: memory.inputBytes, outputCopies: Number(length > 0), outputBytes: length, rustBufferClones: 0, controlDtoConversions: 2, jsonConversions: 0 },
    transport: { inputTransfers: 3, inputBytes: memory.inputBytes, outputTransfers: 1, outputBytes: length, structuredResultMessages: 1 } } }
}
function computeOracle(onSubmit = () => {}) {
  const submitted = []
  const compute = { submit: (input, settings) => {
    submitted.push({ input, settings }); onSubmit(input, settings)
    const query = text(input.queryBytes), matches = []
    for (let row = 0; row < input.params.rowCount; row++) {
      if (text(input.textBytes, input.offsets[row], input.offsets[row + 1]).includes(query)) matches.push(row)
    }
    const context = createWasmTaskControl(input.kind, 'oracle-request', settings.documentId, settings.modelRevision)
    const response = packet(input, context, new Uint32Array(matches.slice(input.params.start, input.params.start + input.params.count)), matches.length)
    return { requestId: context.requestId, cancel() {}, promise: Promise.resolve({ type: 'result', requestId: context.requestId,
      documentId: settings.documentId, modelRevision: settings.modelRevision, engineVersion: response.engineVersion, output: response.output, metrics: response.metrics }) }
  } }
  return { compute, submitted }
}
const guard = extra => ({ assertCurrent() {}, ...extra })
class WorkerPort {
  messages = []; terminated = false; onmessage = null; onerror = null; onmessageerror = null
  postMessage(value, transfer = []) { this.messages.push(structuredClone(value, { transfer })) }
  terminate() { this.terminated = true }
  send(data) { this.onmessage?.({ data }) }
  ready() { this.send({ type: 'ready', protocolVersion: 2, engineVersion: 'list-query-protocol-fixture' }) }
}

test('the three-buffer protocol uses even byte offsets and an exact prefix/window admission estimate', () => {
  const input = task(), cloned = cloneTask(input)
  validateTask(input); assert.equal(taskViews(input).length, 3); assert.deepEqual(progressUnits(input), { total: 3, unit: 'rows' })
  assert.deepEqual(estimateTaskMemory(input), { inputBytes: 28, estimatedPeakBytes: 8 * 1024 * 1024 + 56 + 4 + 24 })
  const moved = structuredClone(input, { transfer: transferTaskBuffers(input) })
  assert.equal(input.offsets.byteLength, 0); assert.deepEqual([...cloned.offsets], [0, 4, 8, 10]); assert.equal(text(moved.textBytes), 'abbca')
  assert.notEqual(moved.params, cloned.params); assert.throws(() => validateTask(input), error => error.code === 'COMPUTE_INVALID_INPUT')
})

test.each(['params', 'row-count', 'odd-text', 'odd-query', 'zero-window', 'query-limit', 'shared', 'subview', 'budget'])('admission rejects %s before caller ownership changes', kind => {
  const input = task()
  if (kind === 'params') input.params.extra = true
  if (kind === 'row-count') input.params.rowCount = Number.MAX_SAFE_INTEGER + 1
  if (kind === 'odd-text') { input.textBytes = new Uint8Array(9); input.params.textByteLength = 9 }
  if (kind === 'odd-query') { input.queryBytes = new Uint8Array(1); input.params.queryByteLength = 1 }
  if (kind === 'zero-window') input.params.count = 0
  if (kind === 'query-limit') input.params.queryByteLength = 32_000_002
  if (kind === 'shared') input.queryBytes = new Uint8Array(new SharedArrayBuffer(2))
  if (kind === 'subview') input.offsets = new Uint32Array(new ArrayBuffer(24), 4, 4)
  if (kind === 'subview') assert.throws(() => transferTaskBuffers(input), error => error.code === 'COMPUTE_INVALID_INPUT')
  else assert.throws(() => validateTask(input, kind === 'budget' ? 1 : 256 * 1024 * 1024))
  assert.ok(input.textBytes.byteLength > 0)
})

test.each(['count', 'total', 'first-index', 'last-index', 'subview', 'shared', 'detached'])('result windows reject invalid %s without exposing a partial array', kind => {
  const input = task(), output = { kind: input.kind, total: 2, indices: new Uint32Array([0, 2]) }
  assert.ok(readValidationListQueryOutput(output, input.params))
  if (kind === 'count') output.indices = new Uint32Array([0])
  if (kind === 'total') output.total = 4
  if (kind === 'first-index') output.indices = new Uint32Array([3, 4])
  if (kind === 'last-index') output.indices[1] = 3
  if (kind === 'subview') output.indices = new Uint32Array(new ArrayBuffer(16), 4, 2)
  if (kind === 'shared') output.indices = new Uint32Array(new SharedArrayBuffer(8))
  if (kind === 'detached') structuredClone(output, { transfer: [output.indices.buffer] })
  assert.equal(readValidationListQueryOutput(output, input.params), null)
})

test.each(['inputCopies', 'inputBytes', 'outputCopies', 'outputBytes', 'rustBufferClones', 'jsonConversions', 'controlDtoConversions'])('protocol guards reject an invented %s counter', field => {
  const input = task(), context = { requestId: 'list-1', documentId: 'book', modelRevision: 4 }
  const control = createWasmTaskControl(input.kind, context.requestId, context.documentId, context.modelRevision), result = packet(input, control)
  assert.ok(readResult(result, context, control, input, estimateTaskMemory(input), input.params.rowCount))
  result.metrics.boundary[field] = field === 'controlDtoConversions' ? 0 : result.metrics.boundary[field] + 1
  assert.equal(readResult(result, context, control, input, estimateTaskMemory(input), input.params.rowCount), null)
})

test.each(['copy', 'transfer'])('the real client keeps %s ownership/copy/transfer accounting distinct from WASM boundary copies', async ownership => {
  const worker = new WorkerPort(), client = new OfficeComputeClient({ workerFactory: () => worker }), input = task()
  try {
    const job = client.submit(input, { documentId: 'book', modelRevision: 4, ownership }); worker.ready()
    const request = worker.messages.find(message => message.type === 'run')
    assert.equal(request.context.kind, input.kind); assert.equal(request.context.unitId, 'book'); assert.equal(request.context.baseRevision, 4)
    assert.equal(input.textBytes.length, ownership === 'copy' ? 10 : 0)
    worker.send(packet(request.task, request.context)); const result = await job.promise
    assert.deepEqual([...result.output.indices], [0, 2]); assert.equal(result.metrics.boundary.inputCopies, 3)
    assert.equal(result.metrics.preparedInputCopies, ownership === 'copy' ? 3 : 0)
    assert.equal(result.metrics.preparedInputBytes, ownership === 'copy' ? 28 : 0)
    assert.equal(result.metrics.ownershipTransfers, ownership === 'transfer' ? 3 : 0)
  } finally { client.dispose() }
})

test.each(['abort', 'revision', 'close'])('the actual client rejects %s while queued and cannot publish a late list window', async change => {
  const worker = new WorkerPort(), client = new OfficeComputeClient({ workerFactory: () => worker }), controller = new AbortController()
  try {
    const job = client.submit(task(), { documentId: 'book', modelRevision: 4, ownership: 'copy', signal: controller.signal })
    const rejection = assert.rejects(job.promise, error => error.code === (change === 'abort' ? 'COMPUTE_CANCELLED' : 'COMPUTE_STALE_RESULT'))
    if (change === 'abort') controller.abort()
    if (change === 'revision') client.setDocumentRevision('book', 5)
    if (change === 'close') client.forgetDocument('book')
    await rejection
  } finally { client.dispose() }
})

test('UTF-16 packing preserves ambient case folding, code units, byte offsets and one source array', async () => {
  const source = options(['Alpha', 'İX', '汉字👩‍💻', '\ud800x', 'a\nb', '']), before = source.map(item => item.label)
  const prepared = await prepareSheetValidationListQuery(source, guard())
  assert.equal(prepared.options, source)
  for (let row = 0; row < source.length; row++) assert.equal(text(prepared.textBytes, prepared.offsets[row], prepared.offsets[row + 1]), before[row].toLocaleLowerCase())
  assert.equal(prepared.offsets.at(-1), prepared.textBytes.length)
  assert.equal(prepared.preparation.encodedBytes, prepared.offsets.byteLength + prepared.textBytes.byteLength)
  assert.equal(prepared.preparation.normalizedCodeUnits, prepared.textBytes.length / 2); assert.deepEqual(source.map(item => item.label), before)
})

test.each(['', 'alpha', '👩‍💻', '\ud800', '\udc69', '\n', 'missing'])('small and Worker %j searches have the same spelling/order/window and color semantics', async query => {
  const source = options(['Alpha', 'β ALPHA', '汉字👩‍💻', '\ud800x', 'a\nb', 'İX']), prepared = await prepareSheetValidationListQuery(source, guard())
  const oracle = computeOracle(), all = source.filter(item => item.label.toLocaleLowerCase().includes(query.toLocaleLowerCase()))
  const result = await querySheetValidationList(prepared, query, guard({ compute: oracle.compute, unitId: 'book', baseRevision: 4 }), 1, 2)
  assert.deepEqual(filterSheetValidationListOptions(source, query), all)
  assert.deepEqual(result.options, all.slice(1, 3)); assert.equal(result.total, all.length)
  assert.equal(Object.isFrozen(result.options), true); assert.equal(result.options.every(item => source.includes(item)), true)
  const { input, settings } = oracle.submitted[0]
  assert.equal(settings.ownership, 'copy'); assert.equal(settings.priority, 'interactive'); assert.equal(settings.documentId, 'book'); assert.equal(settings.modelRevision, 4)
  assert.equal(input.textBytes, prepared.textBytes); assert.equal(result.preparation.queryEncodedBytes, query.toLocaleLowerCase().length * 2)
})

test('the same default-locale boundary handles Turkish folding in both query paths without Rust recasing', async () => {
  const original = String.prototype.toLocaleLowerCase, calls = []
  vi.spyOn(String.prototype, 'toLocaleLowerCase').mockImplementation(function (...args) { calls.push(args); return original.call(this, 'tr-TR') })
  const source = options(['I', 'İ', 'xı', 'i']), prepared = await prepareSheetValidationListQuery(source, guard()), oracle = computeOracle()
  const result = await querySheetValidationList(prepared, 'I', guard({ compute: oracle.compute, unitId: 'book', baseRevision: 0 }))
  assert.deepEqual(result.options.map(item => item.value), ['I', 'xı'])
  assert.deepEqual(filterSheetValidationListOptions(source, 'I').map(item => item.value), ['I', 'xı'])
  assert.equal(calls.every(args => args.length === 0), true)
})

test('chunked encoding records real completed work and checks cancellation after the macro-task yield', async () => {
  vi.useFakeTimers()
  const source = options(Array.from({ length: 2048 }, (_, index) => `Item ${index}`)), controller = new AbortController(), progress = []
  const job = prepareSheetValidationListQuery(source, guard({ signal: controller.signal, onProgress: value => progress.push(value) }))
  const rejection = assert.rejects(job, error => error.code === 'COMPUTE_CANCELLED')
  assert.deepEqual(progress[0], { phase: 'preparing', processed: 1024, total: 4096, ratio: 0.25 })
  controller.abort(); await vi.runAllTimersAsync(); await rejection
  assert.equal(progress.some(value => value.ratio === 1), false)
})

test('long individual labels yield during exact encoding and stale progress cannot publish a prepared source', async () => {
  vi.useFakeTimers()
  const source = options(['A'.repeat(70_000)]), preparedJob = prepareSheetValidationListQuery(source, guard())
  await vi.runAllTimersAsync(); const prepared = await preparedJob
  assert.equal(prepared.preparation.yieldCount, 3); assert.equal(text(prepared.textBytes.subarray(0, 32)), 'a'.repeat(16))
  let current = true
  const stale = prepareSheetValidationListQuery(options(['Alpha']), guard({ assertCurrent: () => { if (!current) throw new ComputeError('COMPUTE_STALE_RESULT', 'stale') },
    onProgress: () => { current = false } }))
  await assert.rejects(stale, error => error.code === 'COMPUTE_STALE_RESULT')
})

test.each(['window', 'text-budget', 'query-budget'])('service %s admission fails before invoking compute', async kind => {
  const oracle = computeOracle(), source = options(['Alpha'])
  if (kind === 'text-budget') await assert.rejects(() => prepareSheetValidationListQuery(source, guard({ memoryBudgetBytes: 1 })), error => error.code === 'COMPUTE_MEMORY_LIMIT')
  else {
    const prepared = await prepareSheetValidationListQuery(source, guard())
    await assert.rejects(() => querySheetValidationList(prepared, 'a', guard({ compute: oracle.compute, unitId: 'book', baseRevision: 0,
      ...(kind === 'query-budget' ? { memoryBudgetBytes: 1 } : {}) }), 0, kind === 'window' ? 0 : 1),
    error => error.code === (kind === 'window' ? 'COMPUTE_INVALID_INPUT' : 'COMPUTE_MEMORY_LIMIT'))
  }
  assert.equal(oracle.submitted.length, 0)
})

test.each(['unit', 'revision', 'count', 'middle-bounds', 'middle-order'])('mapped results reject %s before publishing any candidate window', async kind => {
  const source = options(['a0', 'a1', 'a2', 'a3']), prepared = await prepareSheetValidationListQuery(source, guard()), oracle = computeOracle()
  const compute = { submit(input, settings) {
    const job = oracle.compute.submit(input, settings)
    job.promise = job.promise.then(result => {
      if (kind === 'unit') result.documentId = 'other'
      if (kind === 'revision') result.modelRevision++
      if (kind === 'count') result.output.total = 1
      if (kind === 'middle-bounds') result.output.indices[1] = 9
      if (kind === 'middle-order') result.output.indices[2] = 0
      return result
    })
    return job
  } }
  await assert.rejects(() => querySheetValidationList(prepared, 'a', guard({ compute, unitId: 'book', baseRevision: 4 })),
    error => error.code === (kind === 'unit' || kind === 'revision' ? 'COMPUTE_STALE_RESULT' : 'COMPUTE_PROTOCOL_MISMATCH'))
})

test('result mapping can be cancelled at its own yield without publishing a partial UI array', async () => {
  vi.useFakeTimers()
  const source = options(Array.from({ length: 4096 }, (_, index) => `a${index}`)), preparedJob = prepareSheetValidationListQuery(source, guard())
  await vi.runAllTimersAsync(); const prepared = await preparedJob, oracle = computeOracle(), controller = new AbortController()
  const job = querySheetValidationList(prepared, 'a', guard({ compute: oracle.compute, unitId: 'book', baseRevision: 4, signal: controller.signal }))
  const rejection = assert.rejects(job, error => error.code === 'COMPUTE_CANCELLED')
  await Promise.resolve(); await Promise.resolve(); controller.abort(); await vi.runAllTimersAsync(); await rejection
})
