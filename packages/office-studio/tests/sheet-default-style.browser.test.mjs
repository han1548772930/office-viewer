import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Written for the deferred browser batch. No server or browser was started during implementation.
const route = '/__sheet-default-style__', fixture = readFileSync(new URL('./fixtures/sheet-default-style.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-default-style-fixture', configureServer(vite) {
      vite.middlewares.use(route, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(route, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); baseURL = `http://127.0.0.1:${server.httpServer.address().port}${route}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const editor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const snapshot = () => page.evaluate(() => window.styleFixture.snapshot())
async function edit(address) {
  await nameBox().fill(address); await nameBox().press('Enter'); await grid().press('F2'); await expect(editor()).toBeFocused()
}

test('blank cells use their column font in the Ant input bridge without creating content', async () => {
  const before = await snapshot(); await edit('A2')
  const font = await editor().evaluate(element => getComputedStyle(element).fontFamily)
  assert.ok(font.includes('Courier New')); assert.deepEqual(await snapshot(), before)
  await editor().press('Escape'); assert.equal((await snapshot()).sheets[0].cells['1:0'], undefined)
})

test('a row-font update during a blank edit preserves the input and remains inherited after commit', async () => {
  await edit('B2'); await editor().fill('new text')
  await page.evaluate(() => window.styleFixture.row(1, { ff: 'Inter', fs: 20 }))
  await expect(editor()).toHaveValue('new text')
  assert.ok((await editor().evaluate(element => getComputedStyle(element).fontFamily)).includes('Inter'))
  await editor().press('Enter')
  const cell = (await snapshot()).sheets[0].cells['1:1']
  assert.equal(cell.value, 'new text'); assert.equal(cell.style.font.family, 'Inter'); assert.equal(cell.authoredStyle.font, undefined)
})

test('resetting direct and rich-run font formatting restores row and column defaults', async () => {
  await page.evaluate(() => window.styleFixture.clearFont('A1'))
  const cell = (await snapshot()).sheets[0].cells['0:0']
  assert.equal(cell.style.font.family, 'Georgia'); assert.equal(cell.style.font.size, 16)
  assert.equal(cell.richText[0].font?.family, undefined); assert.equal(cell.richText[0].font?.size, undefined)
  assert.equal(cell.richText[1].vendor.keep, 1)
  await edit('A1'); await expect(editor()).toHaveValue('abc'); await editor().press('Escape')
})

test('inherited percentage formatting controls a blank-cell input and survives snapshot reopening', async () => {
  await edit('D2'); await editor().fill('25%'); await editor().press('Enter')
  const before = await snapshot(), cell = before.sheets[0].cells['1:3']
  assert.equal(cell.value, 0.25); assert.equal(cell.style.font.family, 'Arial')
  await page.evaluate(() => window.styleFixture.reopen())
  await expect(grid()).toBeVisible()
  const reopened = await snapshot(); assert.equal(reopened.sheets[0].cells['1:3'].value, 0.25)
  assert.deepEqual(reopened.sheets[0].columnStyles, before.sheets[0].columnStyles)
})

test('theme and Ribbon layout changes keep inherited content and the Ant-owned input host', async () => {
  await edit('A2'); const before = await snapshot()
  await page.evaluate(() => window.styleFixture.appearance(true, 'grid'))
  await expect(editor()).toBeVisible()
  assert.ok((await editor().evaluate(element => element.className)).includes('ant-'))
  assert.deepEqual(await snapshot(), before)
  await page.evaluate(() => window.styleFixture.appearance(false, 'classic'))
  await expect(editor()).toHaveValue(''); assert.deepEqual(await snapshot(), before)
})
