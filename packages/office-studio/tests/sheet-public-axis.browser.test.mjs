import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Draft only. No dev server or browser was started while implementing these scenarios.
const route = '/__sheet-public-axis__', fixture = readFileSync(new URL('./fixtures/sheet-default-style.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-public-axis-fixture', configureServer(vite) {
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
  await page.evaluate(() => window.styleFixture.loadPublicAxes()); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const editor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const snapshot = () => page.evaluate(() => window.styleFixture.snapshot())
async function select(address) { await nameBox().fill(address); await nameBox().press('Enter') }
async function edit(address) { await select(address); await grid().press('F2'); await expect(editor()).toBeFocused() }
const editorBox = () => page.locator('.canvas-cell-input').boundingBox()

test('public h/w/style data controls the actual Canvas input bounds and Ant font bridge', async () => {
  assert.deepEqual(await page.evaluate(() => window.styleFixture.axisDimensions()), { row: 48, auto: false, column: 180 })
  await edit('A1'); await expect(editor()).toHaveValue('public source')
  const box = await editorBox(); assert.ok(box.width >= 178 && box.width <= 184); assert.ok(box.height >= 46 && box.height <= 52)
  const font = await editor().evaluate(element => getComputedStyle(element).fontFamily); assert.ok(font.includes('Georgia'))
  assert.ok((await editor().evaluate(element => element.className)).includes('ant-'))
  await editor().press('Escape'); assert.equal((await snapshot()).sheets[0].metadata.publicAxisData.rowData[0].custom.keep, 'row')
})

test('Canvas keyboard navigation skips public hidden rows, hidden columns and independent zero widths', async () => {
  await select('A1'); await grid().press('ArrowDown'); await expect(nameBox()).toHaveValue('A3')
  await grid().press('ArrowRight'); await expect(nameBox()).toHaveValue('D3')
  await grid().press('ArrowLeft'); await expect(nameBox()).toHaveValue('A3')
  await grid().press('ArrowUp'); await expect(nameBox()).toHaveValue('A1')
  const sheet = (await snapshot()).sheets[0]
  assert.equal(sheet.columnWidths[1], 0); assert.equal(sheet.columnWidths[2], 90)
  assert.equal(sheet.rowHiddenSpans.find(span => span.start === 1).value, true)
})

test('dimension edits, history, reopen and theme changes keep current geometry while retaining raw axis custom fields', async () => {
  await page.evaluate(() => window.styleFixture.axisSize(0, 64, 0, 210))
  assert.deepEqual(await page.evaluate(() => window.styleFixture.axisDimensions()), { row: 64, auto: false, column: 210 })
  await page.evaluate(() => { window.styleFixture.undo(); window.styleFixture.undo() })
  assert.deepEqual(await page.evaluate(() => window.styleFixture.axisDimensions()), { row: 48, auto: false, column: 180 })
  await page.evaluate(() => { window.styleFixture.redo(); window.styleFixture.redo(); window.styleFixture.reopen(); window.styleFixture.appearance(true, 'grid') })
  await edit('A1'); const box = await editorBox(); assert.ok(box.width >= 208 && box.width <= 214); assert.ok(box.height >= 62 && box.height <= 68)
  const sheet = (await snapshot()).sheets[0]
  assert.equal(sheet.metadata.publicAxisData.columnData[0].w, 180); assert.equal(sheet.metadata.publicAxisData.columnData[0].custom.keep, 'column')
  await editor().fill('edited'); await editor().press('Enter'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'edited')
})
