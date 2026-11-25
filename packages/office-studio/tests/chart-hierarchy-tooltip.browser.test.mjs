import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// Written only: no server, runner, browser, pixel check or benchmark has run.
let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/chart-hierarchy-tooltip.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; page = await browser.newPage({ viewport: { width: 1440, height: 900 } }); page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL); await page.waitForFunction(() => window.chartHierarchyTooltipFixture?.state().charts.length === 2)
})
afterEach(async () => { await page?.evaluate(() => window.chartHierarchyTooltipFixture?.dispose()); await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })
async function hover(id, target = 'residual', expected = id === 'treemap-draft' ? 'Weight: 9' : target === 'parent' ? 'Weight: 9 (40.9%)' : 'Weight: 3 (13.6%)') {
  const point = await page.evaluate(({ id, target }) => window.chartHierarchyTooltipFixture.browserPoint(id, target), { id, target })
  await page.mouse.move(point.x, point.y)
  await page.getByRole('tooltip').filter({ hasText: expected }).waitFor({ state: 'visible' })
  return page.getByRole('tooltip').innerText()
}
async function waitClear() { await page.getByRole('tooltip').waitFor({ state: 'hidden' }) }

test('real native residual/parent hits display Ant content with isolated Treemap and Sunburst weight semantics in a readonly sheet', async () => {
  const before = await page.evaluate(() => window.chartHierarchyTooltipFixture.mount({ readonly: true }))
  assert.equal(await hover('treemap-draft'), 'South\nWeight: 9\n自身数值: 3\n子节点合计: 6')
  assert.equal(await hover('sunburst-draft'), 'South\nWeight: 3 (13.6%)\n自身数值: 3\n子节点合计: 6')
  assert.equal(await hover('treemap-draft', 'parent'), 'South\nWeight: 9\n自身数值: 3\n子节点合计: 6')
  assert.equal(await hover('sunburst-draft', 'parent'), 'South\nWeight: 9 (40.9%)\n自身数值: 3\n子节点合计: 6')
  const state = await page.evaluate(() => window.chartHierarchyTooltipFixture.state())
  assert.equal(state.author, before.author); assert.equal(state.revision, before.revision); assert.equal(state.updates, 0); assert.equal(state.removals, 0)
  assert.equal(state.charts.every(chart => chart.readonly), true)
  assert.equal(state.charts.some(chart => chart.words.some(text => /Weight:|自身数值:/.test(text))), false)
  assert.equal(await page.getByRole('tooltip').count(), 1)
})

test('blank Canvas, leave, Escape and blur clear the Ant tooltip while zoom, theme and scroll repaint and preserve chart authors', async () => {
  const before = await page.evaluate(() => window.chartHierarchyTooltipFixture.state())
  await hover('treemap-draft')
  const rect = await page.locator('.chart-hitbox[data-chart-id="treemap-draft"]').boundingBox()
  await page.mouse.move(rect.x + 2, rect.y + 2); await waitClear()
  await hover('treemap-draft'); await page.keyboard.press('Escape'); await waitClear()
  await page.mouse.move(2, 2); await hover('treemap-draft'); await page.evaluate(() => window.dispatchEvent(new Event('blur'))); await waitClear()
  await page.mouse.move(2, 2); await hover('treemap-draft'); await page.evaluate(() => window.chartHierarchyTooltipFixture.setDark(true)); await waitClear()
  await page.evaluate(() => window.chartHierarchyTooltipFixture.setZoom(125)); await waitClear()
  await page.waitForFunction(() => window.chartHierarchyTooltipFixture.state().charts.every(chart => chart.width === chart.expectedWidth))
  await page.mouse.move(2, 2); assert.match(await hover('treemap-draft'), /Weight: 9/)
  await page.evaluate(() => window.chartHierarchyTooltipFixture.scroll(64)); await waitClear()
  await page.mouse.move(2, 2); assert.match(await hover('treemap-draft'), /Weight: 9/)
  await page.mouse.move(2, 2); await waitClear()
  const after = await page.evaluate(() => window.chartHierarchyTooltipFixture.state())
  assert.equal(after.author, before.author); assert.equal(after.revision, before.revision)
  assert.equal(after.dark, true); assert.equal(after.zoom, 125)
})

test('new prepared facts refresh semantic tooltip values at retained proportions without adding author history', async () => {
  const before = await page.evaluate(() => window.chartHierarchyTooltipFixture.state())
  await hover('sunburst-draft'); await page.evaluate(() => window.chartHierarchyTooltipFixture.changePreparedValues()); await waitClear()
  await page.mouse.move(2, 2)
  assert.equal(await hover('sunburst-draft', 'residual', 'Weight: 6 (13.6%)'), 'South\nWeight: 6 (13.6%)\n自身数值: 6\n子节点合计: 12')
  assert.equal(await hover('treemap-draft', 'residual', 'Weight: 18'), 'South\nWeight: 18\n自身数值: 6\n子节点合计: 12')
  const after = await page.evaluate(() => window.chartHierarchyTooltipFixture.state())
  assert.equal(after.author, before.author); assert.equal(after.revision, before.revision)
  for (const chart of after.charts) assert.ok(chart.version > before.charts.find(item => item.id === chart.id).version)
})

test('unmount releases both raster instances and popup state, and a new host can hover the same chart IDs independently', async () => {
  await hover('treemap-draft')
  const released = await page.evaluate(() => window.chartHierarchyTooltipFixture.release()); await waitClear()
  assert.equal(released.length, 2)
  for (const surface of released) { assert.equal(surface.disposed, true); assert.deepEqual(surface.size, [1, 1]) }
  await page.evaluate(() => window.chartHierarchyTooltipFixture.mount({ readonly: false })); await page.mouse.move(2, 2)
  assert.match(await hover('sunburst-draft'), /Weight: 3 \(13\.6%\)/)
  const state = await page.evaluate(() => window.chartHierarchyTooltipFixture.state())
  assert.equal(state.charts.every(chart => !chart.readonly), true); assert.equal(state.updates, 0); assert.equal(state.removals, 0)
})
