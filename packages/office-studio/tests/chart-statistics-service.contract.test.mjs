import assert from 'node:assert/strict'
import { afterEach, test, vi } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart } from '../src/features/charts/chart-engine.ts'
import { ChartStatisticsService } from '../src/features/charts/services/chart-statistics-service.ts'
import { ComputeError } from '../src/compute/errors.ts'
import { estimateTaskMemory } from '../src/compute/tasks.ts'
import { createWasmTaskControl } from '../src/platform/wasm/task-protocol.ts'

// Drafts for the implementation batch; the Worker itself is covered separately.
const services = new Set()
afterEach(() => { for (const service of services) service.dispose(); services.clear(); vi.useRealTimers() })
function fixture() {
  const workbook = { ...createBlankWorkbook(), id: 'chart-book' }, sheet = workbook.sheets[0]
  sheet.cells = { '0:0': createCell(0, 0, 'Label'), '0:1': createCell(0, 1, 'Value') }
  for (let row = 0; row < 5; row++) {
    sheet.cells[(row + 1) + ':0'] = createCell(row + 1, 0, String(row))
    sheet.cells[(row + 1) + ':1'] = createCell(row + 1, 1, row)
  }
  const chart = createChart(workbook, sheet.id, { type: 'histogram',
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 5, endCol: 1 } },
    position: { x: 20, y: 20 }, size: { width: 360, height: 240 }, categoryField: 0, valueFields: [1],
    options: { histogram: { binning: { mode: 'count', count: 4 }, underflowThreshold: 1, overflowThreshold: 3 } },
  }, 'chart-one')
  workbook.charts = [chart]
  return { workbook, chart, sheet }
}
function deferred() {
  let resolve, reject
  const promise = new Promise((yes, no) => { resolve = yes; reject = no })
  return { promise, resolve, reject }
}
function output() {
  return { kind: 'chartStatistics', values: new Float64Array([0, 1, 2, 2.5, 2.5, 3, 3, 0]),
    indices: new Uint32Array([0, 2, 1, 1, 1, 1, 2, 1]),
    segments: [{ kind: 'histogram', jobIndex: 0, field: 1, seriesPosition: 0, slot: 0, valueOffset: 0, valueLength: 8,
      indexOffset: 0, indexLength: 8, validSampleCount: 5, available: true }] }
}
function client() {
  const calls = [], waiters = new Map(), revisions = new Map()
  return {
    calls,
    setDocumentRevision(id, revision) { revisions.set(id, revision) },
    isCurrent(context) { return revisions.get(context.documentId) === context.modelRevision },
    waitForCall(index) {
      if (calls[index]) return Promise.resolve(calls[index])
      const waiter = deferred(); waiters.set(index, waiter); return waiter.promise
    },
    submit(source, options) {
      const index = calls.length, task = structuredClone(source, { transfer: [source.values.buffer, source.states.buffer, source.categoryIds.buffer] })
      const job = deferred(), requestId = 'chart-compute-' + index
      const abort = () => job.reject(new ComputeError('COMPUTE_CANCELLED', 'Mock Worker acknowledged cancellation'))
      options.signal.addEventListener('abort', abort, { once: true })
      const cleanup = () => options.signal.removeEventListener('abort', abort)
      job.promise.then(cleanup, cleanup)
      const call = { task, source, options,
        complete(result = output()) {
          const memory = estimateTaskMemory(task), outputBytes = result.values.byteLength + result.indices.byteLength
          job.resolve({ type: 'result', requestId, documentId: options.documentId, modelRevision: options.modelRevision,
            context: createWasmTaskControl(task.kind, requestId, options.documentId, options.modelRevision), engineVersion: 'draft-engine',
            output: result, metrics: { elapsedMs: 1, inputBytes: memory.inputBytes, estimatedPeakBytes: memory.estimatedPeakBytes,
              boundary: { inputCopies: 3, inputBytes: memory.inputBytes, outputCopies: 2, outputBytes,
                rustBufferClones: 0, controlDtoConversions: 2, jsonConversions: 0 },
              transport: { inputTransfers: 3, inputBytes: memory.inputBytes, outputTransfers: 2, outputBytes, structuredResultMessages: 1 } } })
        },
        fail(error) { job.reject(error) },
      }
      calls.push(call); waiters.get(index)?.resolve(call); waiters.delete(index)
      return { requestId, promise: job.promise, cancel: abort }
    },
  }
}
function serviceFor(client, options) {
  const service = new ChartStatisticsService(client, options)
  services.add(service); service.setDocumentRevision('chart-book', 7)
  return service
}
const context = extra => ({ documentId: 'chart-book', modelRevision: 7, ...extra })

test('a detached candidate shares the bounded scheduler without replacing the published chart cache', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute), notices = []
  service.subscribe(event => notices.push(event))
  const published = service.request(source.workbook, source.chart, context())
  ;(await compute.waitForCall(0)).complete()
  const cached = await published.promise
  const candidate = structuredClone(source.workbook), candidateChart = candidate.charts[0]
  candidateChart.options.mapping = { useDateAxis: true }
  const request = service.requestCandidate(candidate, candidateChart, context())
  const call = await compute.waitForCall(1)
  assert.equal(service.getCached(source.chart, context()), cached)
  assert.equal(notices.length, 0)
  call.complete(); await request.promise
  assert.equal(service.getCached(source.chart, context()), cached)
  assert.equal(service.getCached(candidateChart, context()), undefined)
  assert.equal(notices.length, 0)
})

test('final invalid date bins reject only the candidate while a published cache remains usable', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const published = service.request(source.workbook, source.chart, context())
  ;(await compute.waitForCall(0)).complete()
  const cached = await published.promise
  const candidate = structuredClone(source.workbook), candidateChart = candidate.charts[0]
  candidateChart.options = { mapping: { useDateAxis: true },
    histogram: { binning: { mode: 'count', count: 2 }, underflowThreshold: -1, overflowThreshold: 1 } }
  for (let row = 0; row < 5; row++) candidate.sheets[0].cells[(row + 1) + ':1'] = createCell(row + 1, 1, row - 2)
  const request = service.requestCandidate(candidate, candidateChart, context())
  const rejection = assert.rejects(request.promise, /日期|date|分类/)
  const call = await compute.waitForCall(1), invalid = output()
  invalid.values = new Float64Array([0, -1, -1, 0, 0, 1, 1, 0])
  call.complete(invalid); await rejection
  assert.equal(service.getCached(source.chart, context()), cached)
  assert.equal(source.workbook.sheets[0].cells['1:1'].value, 0)
})

test('cancelling a queued candidate immediately frees its shared admission slot', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute, { maxPendingCharts: 2 })
  const published = service.request(source.workbook, source.chart, context()), active = await compute.waitForCall(0)
  const candidate = structuredClone(source.workbook)
  const queued = service.requestCandidate(candidate, candidate.charts[0], context())
  const rejection = assert.rejects(queued.promise, error => error.code === 'COMPUTE_CANCELLED')
  assert.throws(() => service.requestCandidate(candidate, candidate.charts[0], context()), error => error.code === 'COMPUTE_MEMORY_LIMIT')
  queued.cancel(); await rejection
  const replacement = service.requestCandidate(candidate, candidate.charts[0], context())
  active.complete(); await published.promise
  ;(await compute.waitForCall(1)).complete(); await replacement.promise
  assert.equal(compute.calls.length, 2)
})

test('two views share one calculation; cancelling one keeps the other and its owned result alive', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context()), second = service.request(source.workbook, source.chart, context())
  const rejection = assert.rejects(first.promise, error => error.code === 'COMPUTE_CANCELLED')
  const call = await compute.waitForCall(0)
  assert.equal(call.source.values.byteLength, 0)
  assert.equal(call.source.states.byteLength, 0)
  assert.equal(call.source.categoryIds.byteLength, 0)
  first.cancel(); await rejection
  assert.equal(call.options.signal.aborted, false); assert.equal(compute.calls.length, 1)
  call.complete()
  const result = await second.promise
  assert.deepEqual(result.statistics.preparedData.series[0].values, [0, 1, 2, 3, 4])
  assert.equal(result.statistics.histogram.validSampleCount, 5)
  assert.equal(service.getCached(source.chart, context()), result)
})

test('the last consumer aborts the active calculation and a later request starts a fresh job', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context()), rejection = assert.rejects(first.promise, error => error.code === 'COMPUTE_CANCELLED')
  const call = await compute.waitForCall(0)
  first.cancel(); await rejection
  assert.equal(call.options.signal.aborted, true)
  assert.equal(service.getCached(source.chart, context()), undefined)
  const second = service.request(source.workbook, source.chart, context()), replacement = await compute.waitForCall(1)
  call.complete(); replacement.complete()
  assert.equal((await second.promise).chartId, source.chart.id)
  assert.equal(compute.calls.length, 2)
})

test('revision changes reject consumers immediately and late Worker data cannot populate the cache', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const request = service.request(source.workbook, source.chart, context()), rejection = assert.rejects(request.promise, error => error.code === 'COMPUTE_STALE_RESULT')
  const call = await compute.waitForCall(0)
  service.setDocumentRevision('chart-book', 8); await rejection
  assert.equal(call.options.signal.aborted, true)
  call.complete(); await Promise.resolve()
  assert.equal(service.getCached(source.chart, context({ modelRevision: 8 })), undefined)
})

test('queued deadlines include waiting for another chart, and expired input never reaches the Worker', async () => {
  vi.useFakeTimers()
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context()), call = await compute.waitForCall(0)
  const secondChart = { ...source.chart, id: 'chart-two' }; source.workbook.charts.push(secondChart)
  const second = service.request(source.workbook, secondChart, context({ timeoutMs: 10 }))
  const rejection = assert.rejects(second.promise, error => error.code === 'COMPUTE_TIMEOUT')
  await vi.advanceTimersByTimeAsync(11); await rejection
  call.complete(); await first.promise; await Promise.resolve(); await Promise.resolve()
  assert.equal(compute.calls.length, 1)
  assert.equal(service.getCached(secondChart, context()), undefined)
})

test('cancelling a queued chart frees its admission slot before the active chart finishes', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute, { maxPendingCharts: 2 })
  const first = service.request(source.workbook, source.chart, context()), active = await compute.waitForCall(0)
  const queuedChart = { ...source.chart, id: 'queued' }, replacementChart = { ...source.chart, id: 'replacement' }
  source.workbook.charts.push(queuedChart, replacementChart)
  const queued = service.request(source.workbook, queuedChart, context())
  const rejection = assert.rejects(queued.promise, error => error.code === 'COMPUTE_CANCELLED')
  assert.throws(() => service.request(source.workbook, replacementChart, context()), error => error.code === 'COMPUTE_MEMORY_LIMIT')
  queued.cancel()
  const replacement = service.request(source.workbook, replacementChart, context())
  await rejection
  assert.equal(compute.calls.length, 1)
  active.complete(); await first.promise
  ;(await compute.waitForCall(1)).complete()
  assert.equal((await replacement.promise).chartId, replacementChart.id)
  assert.equal(compute.calls.length, 2)
  assert.equal(service.getCached(queuedChart, context()), undefined)
})

test('queued interactive charts run before background charts while the active lifecycle completes', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context()), active = await compute.waitForCall(0)
  const backgroundChart = { ...source.chart, id: 'background' }, interactiveChart = { ...source.chart, id: 'interactive' }
  source.workbook.charts.push(backgroundChart, interactiveChart)
  const background = service.request(source.workbook, backgroundChart, context({ priority: 'background' }))
  const interactive = service.request(source.workbook, interactiveChart, context({ priority: 'interactive' }))
  active.complete(); await first.promise
  const next = await compute.waitForCall(1)
  assert.equal(next.options.priority, 'interactive')
  next.complete(); assert.equal((await interactive.promise).chartId, interactiveChart.id)
  const last = await compute.waitForCall(2)
  assert.equal(last.options.priority, 'background')
  last.complete(); assert.equal((await background.promise).chartId, backgroundChart.id)
})

test('a shared consumer has its own deadline without aborting the other view', async () => {
  vi.useFakeTimers()
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context({ timeoutMs: 1000 }))
  const short = service.request(source.workbook, source.chart, context({ timeoutMs: 10 }))
  const rejection = assert.rejects(short.promise, error => error.code === 'COMPUTE_TIMEOUT')
  const call = await compute.waitForCall(0)
  await vi.advanceTimersByTimeAsync(11); await rejection
  assert.equal(call.options.signal.aborted, false)
  assert.equal(compute.calls.length, 1)
  call.complete(); assert.equal((await first.promise).chartId, source.chart.id)
})

test('the original consumer can time out after a longer-lived view attaches to its active calculation', async () => {
  vi.useFakeTimers()
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context({ timeoutMs: 10 }))
  const rejection = assert.rejects(first.promise, error => error.code === 'COMPUTE_TIMEOUT')
  const call = await compute.waitForCall(0)
  const second = service.request(source.workbook, source.chart, context({ timeoutMs: 1000 }))
  await vi.advanceTimersByTimeAsync(11); await rejection
  assert.equal(call.options.signal.aborted, false)
  assert.equal(compute.calls.length, 1)
  call.complete(); assert.equal((await second.promise).chartId, source.chart.id)
})

test('an interactive view promotes its existing queued background calculation without creating a duplicate', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context()), active = await compute.waitForCall(0)
  const backgroundChart = { ...source.chart, id: 'background' }, promotedChart = { ...source.chart, id: 'promoted' }
  source.workbook.charts.push(backgroundChart, promotedChart)
  const background = service.request(source.workbook, backgroundChart, context({ priority: 'background' }))
  const preload = service.request(source.workbook, promotedChart, context({ priority: 'background' }))
  const interactive = service.request(source.workbook, promotedChart, context({ priority: 'interactive' }))
  active.complete(); await first.promise
  const next = await compute.waitForCall(1)
  assert.equal(next.options.priority, 'interactive')
  next.complete(); assert.equal(await interactive.promise, await preload.promise)
  const last = await compute.waitForCall(2)
  assert.equal(last.options.priority, 'background')
  last.complete(); assert.equal((await background.promise).chartId, backgroundChart.id)
  assert.equal(compute.calls.length, 3)
})

test('a completion observer that advances the source revision cannot deliver or cache an old snapshot', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  let invalidate = false
  const request = service.request(source.workbook, source.chart, context({ assertCurrent() {
    if (invalidate) service.setDocumentRevision('chart-book', 8)
  } }))
  const rejection = assert.rejects(request.promise, error => error.code === 'COMPUTE_STALE_RESULT')
  const call = await compute.waitForCall(0)
  invalidate = true
  call.complete()
  // Request.assertCurrent runs again when this consumer receives the snapshot.
  await rejection
  assert.equal(service.getCached(source.chart, context({ modelRevision: 8 })), undefined)
})

test('invalid request controls cannot bypass validation by reusing an existing calculation', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context()), call = await compute.waitForCall(0)
  for (const invalid of [{ timeoutMs: 0 }, { timeoutMs: NaN }, { priority: 'unexpected' }, { memoryBudgetBytes: 0 }]) {
    assert.throws(() => service.request(source.workbook, source.chart, context(invalid)), error => error.code === 'COMPUTE_INVALID_INPUT')
  }
  call.complete(); await first.promise
})

test('cache count eviction and document close release entries without disposing the shared compute client', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute, { maxCachedCharts: 1 })
  const first = service.request(source.workbook, source.chart, context())
  ;(await compute.waitForCall(0)).complete(); await first.promise
  const next = { ...source.chart, id: 'chart-two' }; source.workbook.charts.push(next)
  const second = service.request(source.workbook, next, context())
  ;(await compute.waitForCall(1)).complete(); await second.promise
  assert.equal(service.getCached(source.chart, context()), undefined)
  assert.equal(service.getCached(next, context()).chartId, next.id)
  service.forgetDocument('chart-book')
  assert.equal(service.getCached(next, context()), undefined)
})

test('configuration observer reentry reuses the newly registered request instead of overwriting its slot', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const first = service.request(source.workbook, source.chart, context())
  ;(await compute.waitForCall(0)).complete(); await first.promise
  const changed = { ...source.chart, options: { ...source.chart.options, allSeriesStyle: { color: '#c81e1e' } } }
  source.workbook.charts = [changed]
  let nested, notifications = 0
  service.subscribe(event => {
    if (event.reason === 'configuration') { notifications++; nested = service.request(source.workbook, changed, context()) }
  })
  const outer = service.request(source.workbook, changed, context())
  ;(await compute.waitForCall(1)).complete()
  assert.equal(await outer.promise, await nested.promise)
  assert.equal(notifications, 1); assert.equal(compute.calls.length, 2)
})

test('observer errors do not cancel other consumers or a completed result', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute), seen = []
  const first = service.request(source.workbook, source.chart, context({ onProgress: () => { throw new Error('observer') } }))
  const second = service.request(source.workbook, source.chart, context({ onProgress: progress => seen.push(progress.processed) }))
  const call = await compute.waitForCall(0)
  call.options.onProgress({ type: 'progress', requestId: 'progress', documentId: 'chart-book', modelRevision: 7,
    phase: 'computing', ratio: 0, processed: 0, total: 1, unit: 'jobs' })
  assert.deepEqual(seen, [0]); assert.equal(service.listenerErrors.length, 1)
  call.complete(); assert.equal(await first.promise, await second.promise)
})

test('a progress observer that invalidates the source cannot deliver that progress to detached consumers', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute), seen = []
  const first = service.request(source.workbook, source.chart, context({ onProgress: () => service.setDocumentRevision('chart-book', 8) }))
  const second = service.request(source.workbook, source.chart, context({ onProgress: progress => seen.push(progress.processed) }))
  const firstRejected = assert.rejects(first.promise, error => error.code === 'COMPUTE_STALE_RESULT')
  const secondRejected = assert.rejects(second.promise, error => error.code === 'COMPUTE_STALE_RESULT')
  const call = await compute.waitForCall(0)
  call.options.onProgress({ type: 'progress', requestId: 'progress', documentId: 'chart-book', modelRevision: 7,
    phase: 'computing', ratio: 0, processed: 0, total: 1, unit: 'jobs' })
  await Promise.all([firstRejected, secondRejected])
  assert.deepEqual(seen, [])
  assert.equal(call.options.signal.aborted, true)
})

test('permission, pre-abort and invalid memory controls reject before allocating a Worker request', () => {
  const source = fixture(), compute = client(), service = serviceFor(compute), controller = new AbortController()
  controller.abort()
  assert.throws(() => service.request(source.workbook, source.chart, context({ signal: controller.signal })), error => error.code === 'COMPUTE_CANCELLED')
  source.sheet.worksheetProtection = { id: 'protect', permissionId: 'permission', unitType: 'worksheet', ownerUserId: 'owner',
    editState: 'onlyMe', viewState: 'noOneElseCanView', allowedUserIds: [] }
  assert.throws(() => service.request(source.workbook, source.chart, context({ userId: 'other' })), error => error.code === 'COMPUTE_INVALID_INPUT')
  assert.equal(compute.calls.length, 0)
  assert.throws(() => service.getCached(source.chart, context({ memoryBudgetBytes: 0 })), error => error.code === 'COMPUTE_INVALID_INPUT')
})

test('malformed bulk data rejects the whole prepared result and leaves no successful cache', async () => {
  const source = fixture(), compute = client(), service = serviceFor(compute)
  const request = service.request(source.workbook, source.chart, context()), rejection = assert.rejects(request.promise, error => error.code === 'COMPUTE_PROTOCOL_MISMATCH')
  const invalid = output(); invalid.values[2] = NaN
  ;(await compute.waitForCall(0)).complete(invalid); await rejection
  assert.equal(service.getCached(source.chart, context()), undefined)
})
