import assert from 'node:assert/strict'
import { afterEach, beforeEach, test, vi } from 'vitest'
import { createWasmTaskControl, isWasmTaskControl, matchesWasmTaskControl, requireIndependentOutput, serializeWasmError } from '../src/platform/wasm/task-protocol.ts'
import { initializeParserWasm, initializeWriterWasm } from '../src/parser/wasm-modules.ts'
import { initializeComputeWasm } from '../src/compute/wasm.worker.ts'
import { OfficeComputeClient } from '../src/compute/client.ts'
import { parseOfficeFile, resetParser } from '../src/services/wasmAdapter.ts'
import { exportOfficeModel, resetOfficeWriter } from '../src/services/exportService.ts'
import { estimateTaskMemory } from '../src/compute/tasks.ts'
import { createBlankDocument } from '../src/domain/model.ts'

// Protocol/client drafts. Test execution is deferred until implementation completes.
const workers = []
class WorkerPort {
  onmessage = null
  onerror = null
  onmessageerror = null
  terminated = false
  messages = []
  constructor() { workers.push(this) }
  postMessage(message, transfer = []) { this.messages.push(structuredClone(message, { transfer })) }
  terminate() { this.terminated = true }
  send(data) { this.onmessage?.({ data }) }
}
const copies = extra => ({ inputCopies: 0, inputBytes: 0, outputCopies: 0, outputBytes: 0,
  rustBufferClones: 0, controlDtoConversions: 2, jsonConversions: 0, ...extra })
const transport = extra => ({ inputTransfers: 0, inputBytes: 0, outputTransfers: 0, outputBytes: 0, structuredResultMessages: 0, ...extra })
const memory = () => new WebAssembly.Memory({ initial: 1 })
const task = () => ({ kind: 'aggregate', values: new Float64Array([4, 7, 9]), states: new Uint8Array([1, 1, 1]) })
const statistics = { numericCount: 3, nonEmptyCount: 3, emptyCount: 0, errorCount: 0, invalidNumericCount: 0,
  sum: 20, mean: 20 / 3, min: 4, max: 9, populationVariance: 38 / 9, sampleVariance: 19 / 3,
  sampleStandardDeviation: Math.sqrt(19 / 3), overflowed: false }
const flush = async () => { await Promise.resolve(); await Promise.resolve(); await Promise.resolve() }
function deferred() {
  let resolve, reject
  const promise = new Promise((yes, no) => { resolve = yes; reject = no })
  return { promise, resolve, reject }
}
function aggregateResult(request) {
  const measured = estimateTaskMemory(request.task)
  return { type: 'result', context: request.context, engineVersion: '2.0.0', output: { kind: 'aggregate', statistics },
    metrics: { elapsedMs: 1, ...measured, boundary: copies({ inputCopies: 2, inputBytes: 27 }),
      transport: transport({ inputTransfers: 2, inputBytes: 27, structuredResultMessages: 1 }) } }
}
function ready(worker) { worker.send({ type: 'ready', protocolVersion: 2, engineVersion: '2.0.0' }) }

beforeEach(() => {
  workers.length = 0
  vi.stubGlobal('Worker', WorkerPort)
  vi.stubGlobal('window', { location: { href: 'https://office.test/' } })
})
afterEach(() => { resetParser(); resetOfficeWriter(); vi.useRealTimers(); vi.unstubAllGlobals() })

test('controls reject old versions, unknown keys, unsafe revisions and altered budgets', () => {
  const context = createWasmTaskControl('parse', 'parse-1', 'document-1', 7)
  assert.equal(isWasmTaskControl(context), true)
  for (const value of [{ ...context, protocolVersion: 1 }, { ...context, baseRevision: Number.MAX_SAFE_INTEGER + 1 },
    { ...context, oldRequestId: 'parse-1' }, { ...context, budget: { ...context.budget, extra: 1 } }]) assert.equal(isWasmTaskControl(value), false)
  assert.equal(matchesWasmTaskControl({ ...context, budget: { ...context.budget, maxInputBytes: 1 } }, context), false)
  const oldError = JSON.stringify({ ok: false, error: { code: 'OLD_ERROR', message: 'old result' } })
  assert.equal(serializeWasmError(oldError, { code: 'NEW_FAILURE', message: 'failed' }).code, 'NEW_FAILURE')
})

test('module initializers require protocol 2 interfaces and a real memory handle', async () => {
  await assert.rejects(() => initializeParserWasm({ default: async () => ({}), parse_office: () => '{}' }), error => error.code === 'WASM_API_MISSING')
  await assert.rejects(() => initializeWriterWasm({ default: async () => ({}), write_office: () => ({ bytes: [] }) }), error => error.code === 'WASM_API_MISSING')
  const parser = { default: async () => ({ memory: memory() }), parser_protocol_version: () => 1, parse_office: async () => ({}) }
  await assert.rejects(() => initializeParserWasm(parser), error => error.code === 'WASM_PROTOCOL_MISMATCH')
  const compute = { default: async () => ({ memory: memory() }), compute_protocol_version: () => 1, compute_version: () => '1.0.0',
    compute_yield: async () => {}, create_numeric_aggregate() {}, create_numeric_sort() {}, create_numeric_filter() {} }
  await assert.rejects(() => initializeComputeWasm(compute), error => error.code === 'COMPUTE_PROTOCOL_MISMATCH')
  await assert.rejects(() => initializeParserWasm({ ...parser, parser_protocol_version: () => 2, default: async () => ({ memory: new ArrayBuffer(8) }) }), error => error.code === 'WASM_API_MISSING')
})

test('bulk outputs reject linear memory, subviews and shared buffers', () => {
  const wasmMemory = memory()
  assert.throws(() => requireIndependentOutput(new Uint8Array(wasmMemory.buffer), wasmMemory), error => error.code === 'WASM_OUTPUT_ALIAS')
  assert.throws(() => requireIndependentOutput(new Uint8Array(new ArrayBuffer(8), 2, 4), wasmMemory), error => error.code === 'WASM_OUTPUT_ALIAS')
  const owned = new Uint32Array([1, 2])
  assert.equal(requireIndependentOutput(owned, wasmMemory), owned.buffer)
})

test('compute clients send nested contexts and reject mismatched results', async () => {
  const client = new OfficeComputeClient({ workerFactory: () => new WorkerPort(), wasmModuleUrl: 'https://office.test/compute.js' })
  try {
    const job = client.submit(task(), { documentId: 'book', modelRevision: 7, ownership: 'copy' })
    const rejection = assert.rejects(job.promise, error => error.code === 'COMPUTE_PROTOCOL_MISMATCH')
    const worker = workers[0]; ready(worker)
    const request = worker.messages.at(-1)
    assert.equal(request.context.unitId, 'book')
    assert.equal(request.context.baseRevision, 7)
    assert.equal(request.context.protocolVersion, 2)
    assert.equal(Object.hasOwn(request, 'documentId'), false)
    worker.send({ ...aggregateResult(request), context: { ...request.context, baseRevision: 8 } })
    await rejection
    assert.equal(worker.terminated, true)
  } finally { client.dispose() }
})

test('late completed-task events cannot settle a later calculation', async () => {
  const client = new OfficeComputeClient({ workerFactory: () => new WorkerPort(), wasmModuleUrl: 'https://office.test/compute.js' })
  try {
    const first = client.submit(task(), { documentId: 'book', modelRevision: 1, ownership: 'copy' })
    const worker = workers[0]; ready(worker)
    const oldRequest = worker.messages.at(-1)
    worker.send(aggregateResult(oldRequest)); await first.promise
    const second = client.submit(task(), { documentId: 'book', modelRevision: 1, ownership: 'copy' })
    const request = worker.messages.at(-1)
    worker.send(aggregateResult(oldRequest))
    worker.send(aggregateResult(request))
    assert.equal((await second.promise).requestId, second.requestId)
  } finally { client.dispose() }
})

test('unresponsive cancellation kills the old worker and recreates it for queued work', async () => {
  vi.useFakeTimers()
  const client = new OfficeComputeClient({ workerFactory: () => new WorkerPort(), wasmModuleUrl: 'https://office.test/compute.js', cancellationGraceMs: 20 })
  try {
    const first = client.submit(task(), { documentId: 'book', modelRevision: 1, ownership: 'copy' })
    const worker = workers[0]; ready(worker)
    const oldRequest = worker.messages.at(-1), lateHandler = worker.onmessage
    const second = client.submit(task(), { documentId: 'other', modelRevision: 1, ownership: 'copy' })
    const rejection = assert.rejects(first.promise, error => error.code === 'COMPUTE_CANCELLED')
    first.cancel(); await rejection
    assert.equal(worker.messages.at(-1).type, 'cancel')
    assert.deepEqual(worker.messages.at(-1).context, oldRequest.context)
    await vi.advanceTimersByTimeAsync(20)
    assert.equal(worker.terminated, true)
    assert.equal(workers.length, 2)
    lateHandler({ data: aggregateResult(oldRequest) })
    const replacement = workers[1]; ready(replacement)
    replacement.send(aggregateResult(replacement.messages.at(-1)))
    assert.equal((await second.promise).documentId, 'other')
  } finally { client.dispose() }
})

test('parse cancellation covers file reading before a Worker exists', async () => {
  const reading = deferred(), controller = new AbortController()
  const promise = parseOfficeFile({ name: 'slow.docx', size: 4, arrayBuffer: () => reading.promise }, { signal: controller.signal })
  const rejection = assert.rejects(promise, error => error.code === 'CANCELLED')
  controller.abort(); await rejection
  reading.resolve(new ArrayBuffer(4)); await flush()
  assert.equal(workers.length, 0)
})

test('parse deadlines include file reading and ignore late file data', async () => {
  vi.useFakeTimers()
  const reading = deferred()
  const promise = parseOfficeFile({ name: 'late.docx', size: 4, arrayBuffer: () => reading.promise }, { timeoutMs: 25 })
  const rejection = assert.rejects(promise, error => error.code === 'PARSE_TIMEOUT')
  await vi.advanceTimersByTimeAsync(25); await rejection
  reading.resolve(new ArrayBuffer(4)); await flush()
  assert.equal(workers.length, 0)
})

test('parse rejects results from a different unit and releases the Worker', async () => {
  const promise = parseOfficeFile({ name: 'file.docx', size: 4, arrayBuffer: async () => new ArrayBuffer(4) }, { unitId: 'expected' })
  const rejection = assert.rejects(promise, error => error.code === 'WASM_PROTOCOL_MISMATCH')
  await flush()
  const worker = workers[0], context = worker.messages[0].context
  worker.send({ type: 'error', context: { ...context, unitId: 'different' }, error: { code: 'INVALID_FILE', message: 'failed' } })
  await rejection
  assert.equal(worker.terminated, true)
})

test('export rejects an invalidated model revision even when the Worker succeeds', async () => {
  const model = { ...createBlankDocument(), id: 'export-document' }
  let current = true
  const promise = exportOfficeModel('document', model, { modelRevision: 8, isCurrent: () => current })
  const rejection = assert.rejects(promise, error => error.code === 'EXPORT_STALE_RESULT')
  const worker = workers[0], request = worker.messages[0]
  assert.equal(request.context.baseRevision, 8)
  current = false
  worker.send({ type: 'success', context: request.context, bytes: new Uint8Array([1]), warnings: [], metrics: {
    elapsedMs: 1, boundary: copies({ outputCopies: 1, outputBytes: 1, jsonConversions: 1 }),
    transport: transport({ outputTransfers: 1, outputBytes: 1 }),
  } })
  await rejection
  assert.equal(worker.terminated, true)
})

test('export accepts one complete owned output buffer and structured warnings', async () => {
  const model = { ...createBlankDocument(), id: 'export-document' }
  const promise = exportOfficeModel('document', model, { modelRevision: 3 })
  const worker = workers[0], request = worker.messages[0], bytes = new Uint8Array([0x50, 0x4b, 3, 4])
  const warnings = [{ code: 'SOURCE_DIAGNOSTIC', message: 'retained source requires review' }]
  worker.send({ type: 'success', context: request.context, bytes, warnings, metrics: { elapsedMs: 1,
    boundary: copies({ outputCopies: 1, outputBytes: 4, jsonConversions: 1 }), transport: transport({ outputTransfers: 1, outputBytes: 4 }) } })
  const result = await promise
  assert.equal(result.bytes, bytes)
  assert.deepEqual(result.context, request.context)
  assert.deepEqual(result.warnings, warnings)
})
