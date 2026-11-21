import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let serviceModule
let ComputeError

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  serviceModule = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-compute/numeric-index-service.ts')
  const errors = await server.ssrLoadModule('/packages/office-studio/src/compute/errors.ts')
  ComputeError = errors.ComputeError
})

afterAll(async () => {
  await server?.close()
})

function fakeClient({ current = true, outputKind = 'sort', indices = [0], metrics = {} } = {}) {
  const calls = []
  return {
    calls,
    setDocumentRevision() {},
    isCurrent() { return current },
    submit(task, options) {
      calls.push({ task, options })
      return {
        requestId: 'compute-test-1',
        cancel() {},
        promise: Promise.resolve({
          type: 'result',
          requestId: 'compute-test-1',
          documentId: options.documentId,
          modelRevision: options.modelRevision,
          engineVersion: 'test-engine',
          output: { kind: outputKind, indices: new Uint32Array(indices) },
          metrics: { elapsedMs: 1, inputBytes: 1, estimatedPeakBytes: 1, ...metrics },
        }),
      }
    },
  }
}

function context(overrides = {}) {
  return { documentId: 'workbook-test', modelRevision: 7, ...overrides }
}

test('delegates large numeric sort to the compute task with column-major buffers', async () => {
  const client = fakeClient({ indices: [1, 2, 0] })
  const service = new serviceModule.NumericIndexService(client, { minRows: 2 })
  const result = await service.sortRows({
    rowCount: 3,
    columns: [[30, 10, 20], [3, 1, 2]],
    directions: ['descending', 'ascending'],
    nullsLast: false,
  }, context())

  assert.deepEqual([...result.indices], [1, 2, 0])
  assert.equal(client.calls.length, 1)
  const { task, options } = client.calls[0]
  assert.equal(task.kind, 'sort')
  assert.equal(task.rowCount, 3)
  assert.deepEqual([...task.values], [30, 10, 20, 3, 1, 2])
  assert.deepEqual([...task.states], [1, 1, 1, 1, 1, 1])
  assert.deepEqual([...task.directions], [1, 0])
  assert.equal(task.nullsLast, false)
  assert.equal(options.documentId, 'workbook-test')
  assert.equal(options.modelRevision, 7)
  assert.equal(options.ownership, 'transfer')
})

test('delegates large numeric filter and preserves predicate semantics in the task', async () => {
  const client = fakeClient({ outputKind: 'filter', indices: [0, 2] })
  const service = new serviceModule.NumericIndexService(client, { minRows: 2 })
  const result = await service.filterRows({
    rowCount: 3,
    columns: [[1, 2, 3], [10, 20, 30]],
    predicates: [
      { operator: 'between', lower: 1, upper: 2 },
      { operator: 'isNotEmpty' },
    ],
    match: 'any',
  }, context({ modelRevision: 8 }))

  assert.deepEqual([...result.indices], [0, 2])
  const { task } = client.calls[0]
  assert.equal(task.kind, 'filter')
  assert.equal(task.match, 'any')
  assert.deepEqual([...task.operators], [6, 9])
  assert.deepEqual([...task.lower], [1, 0])
  assert.deepEqual([...task.upper], [2, 0])
})

test('returns null for small or unsupported inputs so the owning feature can use its TS comparator', async () => {
  const client = fakeClient()
  const service = new serviceModule.NumericIndexService(client, { minRows: 4 })
  assert.equal(await service.sortRows({ rowCount: 3, columns: [[3, 1, 2]] }, context()), null)
  assert.equal(await service.sortRows({ rowCount: 4, columns: [['3', '1', '2', '4']] }, context()), null)
  assert.equal(client.calls.length, 0)
})

test('rejects a result when the workbook revision is no longer current', async () => {
  const client = fakeClient({ current: false })
  const service = new serviceModule.NumericIndexService(client, { minRows: 1 })
  await assert.rejects(
    () => service.sortRows({ rowCount: 1, columns: [[1]] }, context()),
    (error) => error instanceof ComputeError && error.code === 'COMPUTE_STALE_RESULT',
  )
})

test('propagates cancellation while packing before any compute task is submitted', async () => {
  const client = fakeClient()
  const service = new serviceModule.NumericIndexService(client, { minRows: 1 })
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(
    () => service.filterRows({ rowCount: 1, columns: [[1]], predicates: [{ operator: 'eq', lower: 1 }] }, context({ signal: controller.signal })),
    (error) => error instanceof ComputeError && error.code === 'COMPUTE_CANCELLED',
  )
  assert.equal(client.calls.length, 0)
})

test('forwards an abort signal to an active compute job', async () => {
  let rejectJob
  let submittedOptions
  const client = {
    calls: [],
    setDocumentRevision() {},
    isCurrent() { return true },
    submit(_task, options) {
      submittedOptions = options
      const promise = new Promise((_resolve, reject) => { rejectJob = reject })
      options.signal.addEventListener('abort', () => rejectJob(new ComputeError('COMPUTE_CANCELLED', 'aborted')), { once: true })
      return { requestId: 'compute-test-2', cancel() {}, promise }
    },
  }
  const service = new serviceModule.NumericIndexService(client, { minRows: 1 })
  const controller = new AbortController()
  const job = service.sortRows({ rowCount: 1, columns: [[1]] }, context({ signal: controller.signal }))
  await Promise.resolve()
  assert.equal(submittedOptions.signal, controller.signal)
  controller.abort()
  await assert.rejects(job, (error) => error instanceof ComputeError && error.code === 'COMPUTE_CANCELLED')
})
