import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// Written only: no server/browser/runner has started. Canvas previews use
// prepared data; Apply requires the real OfficeStudio/Worker/WASM platform.
let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/chart-titles.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; page = await browser.newPage(); page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL); await page.waitForFunction(() => Boolean(window.chartTitlesFixture))
})
afterEach(async () => { await page?.evaluate(() => window.chartTitlesFixture?.dispose()); await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })
async function choose(label, option) {
  await page.locator(`[aria-label="${label}"]`).click()
  await page.getByRole('option', { name: option, exact: true }).click()
}

test('real Canvas native/shared content shows four-edge title glyphs, hides the whole title block and preserves document authors', async () => {
  const result = await page.evaluate(async () => {
    const f = window.chartTitlesFixture, before = JSON.stringify(f.state().author), views = []
    for (const position of ['top', 'bottom', 'left', 'right']) {
      await f.setPreview(position); const native = f.nativeText(), commands = f.state().commands
      await f.setPreview(position, true); const shared = f.nativeText()
      views.push({ position, native, shared, commands })
    }
    await f.setPreview('hide')
    return { views, hidden: f.state().commands.length === 0 && !f.nativeText().some(text => ['Main', 'Sub', 'M', 'a', 'i', 'n', 'S', 'u', 'b'].includes(text)), author: JSON.stringify(f.state().author) === before }
  })
  assert.equal(result.hidden, true); assert.equal(result.author, true)
  for (const view of result.views) {
    for (const command of view.commands) { assert.ok(view.native.includes(command.text)); assert.ok(view.shared.includes(command.text)) }
    assert.ok(view.commands.length > 0)
  }
})

test('production Ant title fields stay detached until Apply and preserve real command history and internal saved configuration', async () => {
  const initial = await page.evaluate(() => window.chartTitlesFixture.state())
  await page.getByRole('textbox', { name: '图表标题', exact: true }).fill('Moved title')
  await choose('标题位置', '左侧'); await choose('标题对齐', '靠右'); await choose('副标题对齐', '靠左')
  const draft = await page.evaluate(() => window.chartTitlesFixture.state())
  assert.deepEqual(draft.author, initial.author); assert.equal(draft.draft.title, 'Moved title')
  await page.getByRole('button', { name: '应用标题', exact: true }).click()
  await page.waitForFunction(() => { const state = window.chartTitlesFixture.state(); return state.applies === 1 || !!state.error })
  const authored = await page.evaluate(() => window.chartTitlesFixture.state())
  assert.equal(authored.error, ''); assert.equal(authored.author.title, 'Moved title')
  assert.equal(authored.author.options.titleStyle.position, 'left'); assert.equal(authored.author.options.titleStyle.alignment, 'right')
  assert.equal(authored.author.options.subtitleStyle.alignment, 'left')
  await page.evaluate(() => window.chartTitlesFixture.undo()); assert.deepEqual((await page.evaluate(() => window.chartTitlesFixture.state())).author, initial.author)
  await page.evaluate(() => window.chartTitlesFixture.redo()); assert.deepEqual((await page.evaluate(() => window.chartTitlesFixture.state())).author, authored.author)
  await page.evaluate(() => window.chartTitlesFixture.reopen())
  const reopened = await page.evaluate(() => window.chartTitlesFixture.state())
  assert.equal(reopened.author.title, 'Moved title'); assert.deepEqual(reopened.author.options.titleStyle, authored.author.options.titleStyle)
  await page.getByRole('button', { name: '图表标题恢复默认', exact: true }).click()
  await page.getByRole('button', { name: '应用标题', exact: true }).click()
  await page.waitForFunction(() => window.chartTitlesFixture.state().applies === 2 || !!window.chartTitlesFixture.state().error)
  const cleared = await page.evaluate(() => window.chartTitlesFixture.state())
  assert.equal(cleared.error, ''); assert.equal(cleared.author.title, undefined); assert.equal(cleared.author.options.titleStyle, undefined); assert.equal(cleared.author.subtitle, 'Sub')
})

test('theme and DPR redraw retain title author fields and Ant input focus while publishing new Canvas dimensions', async () => {
  const input = page.getByRole('textbox', { name: '图表标题', exact: true })
  await input.focus(); const before = await page.evaluate(() => window.chartTitlesFixture.state())
  await page.evaluate(async () => { await window.chartTitlesFixture.setDark(true); await window.chartTitlesFixture.setDpr(2) })
  await page.waitForFunction(version => { const state = window.chartTitlesFixture.state(); return state.sourceVersion > version && state.width === state.expectedWidth && state.height === state.expectedHeight }, before.sourceVersion)
  const after = await page.evaluate(() => window.chartTitlesFixture.state())
  assert.deepEqual(after.author, before.author); assert.equal(after.applies, 0)
  assert.equal(await input.evaluate(element => element === document.activeElement), true)
  assert.ok(after.commands.every(command => command.color === '#eeeeee'))
})
