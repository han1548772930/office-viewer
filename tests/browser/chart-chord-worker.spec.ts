import { expect, test } from '@playwright/test'
import type {} from './chart-chord-worker'

// These six P28 drafts require the production protocol 2 WASM artifacts. They
// declare actual Worker/TypedArray behavior; none ran during implementation.
const goldenValues = [5, -Math.PI / 2, Math.PI / 2, 5, Math.PI / 2, Math.PI * 3 / 2,
  5, -Math.PI / 2, Math.PI / 2, Math.PI / 2, Math.PI * 3 / 2]
const goldenIndices = [2, 1, 0, 1, 0, 1]
const zeroCheckpoint = { phase: 'computing', ratio: 0, processed: 0, total: 1, unit: 'jobs' }
function expectGolden(output: { values: number[]; indices: number[] }) {
  expect(output.indices).toEqual(goldenIndices); expect(output.values).toHaveLength(goldenValues.length)
  output.values.forEach((value, index) => expect(value).toBeCloseTo(goldenValues[index]!, 10))
}

test.beforeEach(async ({ page }) => {
  await page.goto('/tests/browser/chart-chord-worker.html')
  await expect(page.getByRole('status')).toHaveText('ready')
})

test('real Chord bulk transfer detaches caller buffers and returns independent aggregate/angle owners with counted boundary copies', async ({ page }) => {
  const result = await page.evaluate(() => window.chordWorkerRuntime.runTransfer())
  expectGolden(result.output)
  expect(result.output.segments).toEqual([{ kind: 'chord', jobIndex: 0, field: 2, seriesPosition: 0, slot: 0,
    valueOffset: 0, valueLength: 11, indexOffset: 0, indexLength: 6, validSampleCount: 2, available: true }])
  expect(result.output.context).toMatchObject({ protocolVersion: 2, unitId: 'chord-worker-transfer', baseRevision: 0, kind: 'chartStatistics' })
  expect(result.output.engineVersion).toBeTruthy(); expect(result.output.independent).toBe(true)
  expect(result.inputBytes).toEqual([48, 6, 8]); expect(result.callerBytesAfterSubmit).toEqual([0, 0, 0])
  expect(result.output.metrics).toMatchObject({ preparedInputCopies: 0, preparedInputBytes: 0, ownershipTransfers: 3,
    boundary: { inputCopies: 3, inputBytes: 62, outputCopies: 2, outputBytes: 112, rustBufferClones: 0, jsonConversions: 0 },
    transport: { inputTransfers: 3, inputBytes: 62, outputTransfers: 2, outputBytes: 112, structuredResultMessages: 1 } })
  expect(result.second.values[0]).toBe(8); expect(result.second.indices).toEqual(goldenIndices)
  expect(result.firstValuesAfterReuse).toEqual(result.output.values); expect(result.firstIndicesAfterReuse).toEqual(result.output.indices)
  expect(result.trace.workers).toHaveLength(1); expect(result.trace.workers[0]!.terminated).toBe(true)
  expect(result.trace.workers[0]!.posts.filter(post => post.type === 'run')).toEqual([
    { type: 'run', requestId: result.output.requestId, transferCount: 3, delivered: true },
    { type: 'run', requestId: result.second.requestId, transferCount: 3, delivered: true },
  ])
})

test('AbortSignal at a real computing checkpoint rejects Chord work and permits a later native Worker request', async ({ page }) => {
  const result = await page.evaluate(() => window.chordWorkerRuntime.runCancellation())
  expect(result.progress[0]).toMatchObject(zeroCheckpoint); expect(result.aborted).toBe(true)
  expect(result.rejected.code).toBe('COMPUTE_CANCELLED'); expectGolden(result.recovered)
  expect(result.trace.failedCancelPosts).toBe(0)
  expect(result.trace.workers.flatMap(worker => worker.posts).some(post => post.type === 'cancel' && post.delivered)).toBe(true)
  expect(result.trace.workers.every(worker => worker.terminated)).toBe(true)
})

test('one injected cancel postMessage failure actually terminates and recreates the native Worker before recovery', async ({ page }) => {
  const result = await page.evaluate(() => window.chordWorkerRuntime.runCancelPostFailure())
  expect(result.progress[0]).toMatchObject(zeroCheckpoint); expect(result.rejected.code).toBe('COMPUTE_CANCELLED')
  expect(result.trace.failedCancelPosts).toBe(1); expect(result.trace.workers).toHaveLength(2)
  expect(result.trace.workers[0]!.posts.at(-1)).toMatchObject({ type: 'cancel', transferCount: 0, delivered: false })
  expect(result.trace.workers[0]!.terminated).toBe(true)
  expect(result.trace.workers[1]!.posts.map(post => post.type)).toEqual(['init', 'run'])
  expect(result.trace.workers[1]!.terminated).toBe(true); expectGolden(result.recovered)
})

test('revision zero Chord results are rejected after revision one is announced and old admissions retain their caller buffers', async ({ page }) => {
  const result = await page.evaluate(() => window.chordWorkerRuntime.runStaleRevision())
  expect(result.progress[0]).toMatchObject(zeroCheckpoint); expect(result.rejected.code).toBe('COMPUTE_STALE_RESULT')
  expect(result.oldAdmission.code).toBe('COMPUTE_STALE_RESULT'); expect(result.oldCallerBytes).toEqual([48, 6, 8])
  expect(result.oldCurrent).toBe(false); expect(result.newCurrent).toBe(true)
  expect(result.recovered.modelRevision).toBe(1); expect(result.recovered.context.baseRevision).toBe(1); expectGolden(result.recovered)
  expect(result.trace.workers.every(worker => worker.terminated)).toBe(true)
})

test('a spreadsheet error weight reaches the real Rust Chord validator, publishes no successful result and leaves the Worker reusable', async ({ page }) => {
  const result = await page.evaluate(() => window.chordWorkerRuntime.runInvalidWeight())
  expect(result.rejected).toMatchObject({ code: 'COMPUTE_INVALID_INPUT', message: 'Chord weight state is invalid' })
  expectGolden(result.recovered); expect(result.trace.workers).toHaveLength(1)
  expect(result.trace.workers[0]!.posts.filter(post => post.type === 'run')).toHaveLength(2)
  expect(result.trace.workers[0]!.posts.every(post => post.delivered)).toBe(true)
  expect(result.trace.workers[0]!.terminated).toBe(true)
})

test('disposing active Chord work terminates its native Worker and rejects subsequent submissions without spawning another owner', async ({ page }) => {
  const result = await page.evaluate(() => window.chordWorkerRuntime.runDispose())
  expect(result.progress[0]).toMatchObject(zeroCheckpoint)
  expect(result.rejected.code).toBe('COMPUTE_DISPOSED'); expect(result.afterDispose.code).toBe('COMPUTE_DISPOSED')
  expect(result.trace.workers).toHaveLength(1); expect(result.trace.workers[0]!.terminated).toBe(true)
  expect(result.trace.workers[0]!.posts.map(post => post.type)).toEqual(['init', 'run'])
})
