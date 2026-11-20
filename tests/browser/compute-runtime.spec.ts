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
