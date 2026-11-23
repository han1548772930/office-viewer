import { expect, test } from '@playwright/test'

test.beforeEach(async ({ page }) => {
  await page.goto('/tests/browser/compute-runtime.html')
  await expect(page.locator('#status')).toHaveText('ready')
})

test('loads the built Rust WASM module and returns deterministic sort and filter indices', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runSortAndFilter())
  expect(result.sort.indices).toEqual([3, 1, 2, 0])
  expect(result.filter.indices).toEqual([1, 2, 3])
  expect(result.sort.engineVersion).toBeTruthy()
  expect(result.filter.engineVersion).toBe(result.sort.engineVersion)
})

test('returns finite numeric statistics and preserves empty, error and invalid-numeric categories', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runAggregate())
  expect(result.statistics).toMatchObject({
    numericCount: 3,
    nonEmptyCount: 5,
    emptyCount: 0,
    errorCount: 1,
    invalidNumericCount: 1,
    sum: 13,
    mean: 13 / 3,
    min: 2,
    max: 7,
    overflowed: false,
  })
  expect(result.engineVersion).toBeTruthy()
})

// Drafts for unified validation after rebuilding the protocol 2 WASM artifacts.
test('chart statistics preserve authored slots and report the real bulk ownership transfers', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runChartStatistics())
  expect(result.values).toEqual([4, 6.5, 6.5, 9, 4, 5.5, 7, 8, 9])
  expect(result.indices).toEqual([1, 1, 1, 2])
  expect(result.segments.map(segment => [segment.kind, segment.field, segment.slot])).toEqual([['histogram', 7, 2], ['boxplot', 7, 5]])
  expect(result.callerDetached).toBe(true)
  expect(result.metrics).toMatchObject({ preparedInputCopies: 0, ownershipTransfers: 3,
    boundary: { inputCopies: 3, inputBytes: 39, outputCopies: 2, outputBytes: 88, rustBufferClones: 0, jsonConversions: 0 },
    transport: { inputTransfers: 3, inputBytes: 39, outputTransfers: 2, outputBytes: 88, structuredResultMessages: 1 } })
})

test('the chart service prepares real workbook data, caches its result and recomputes after source edits', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runChartService())
  expect(result.categories).toEqual(['A', 'B', 'C'])
  expect(result.values).toEqual([4, 7, 9])
  expect(result.counts).toEqual([1, 2])
  expect(result.updatedCounts).toEqual([2, 1])
  expect(result.cached).toBe(true)
  expect(result.invalidated).toBe(true)
  expect(result.metrics?.transport.inputTransfers).toBe(3)
})

test('chart cancellation permits a subsequent Worker/WASM chart to finish', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runChartCancellation())
  expect(result.cancelled).toBe('COMPUTE_CANCELLED')
  expect(result.recoveredChart).toEqual([4, 6.5, 6.5, 9, 4, 5.5, 7, 8, 9])
})

test('a revised chart source rejects active stale work and computes its current revision', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runChartStaleRevision())
  expect(result.stale).toBe('COMPUTE_STALE_RESULT')
  expect(result.recoveredChart).toEqual([4, 6.5, 6.5, 9, 4, 5.5, 7, 8, 9])
})

test('cancels an active Worker/WASM task cooperatively', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runCancellation())
  expect(result.cancelled).toBe('COMPUTE_CANCELLED')
  expect(result.recoveredIndices).toEqual([1, 2, 0])
})

test('rejects active results when their workbook revision advances', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runStaleRevision())
  expect(result.stale).toBe('COMPUTE_STALE_RESULT')
})

test('dispose terminates the Worker and rejects its active job', async ({ page }) => {
  const result = await page.evaluate(() => window.computeRuntime.runDispose())
  expect(result.disposed).toBe('COMPUTE_DISPOSED')
})
