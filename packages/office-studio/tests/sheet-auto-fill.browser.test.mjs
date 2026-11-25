import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Draft only. Browser, IME, rendering backend and file round trips remain unaccepted.
const route = '/__sheet-auto-fill-test__', fixture = readFileSync(new URL('./fixtures/sheet-selection-input.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-auto-fill-fixture', configureServer(vite) {
      vite.middlewares.use(route, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(route, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Browser fixture did not allocate a port')
  baseURL = `http://127.0.0.1:${address.port}${route}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
  await page.evaluate(() => window.selectionInputFixture.seed({ entries: [
    { row: 0, value: 1, style: { font: { bold: true }, fill: '#ffeeaa' } },
    { row: 1, value: 2, style: { font: { bold: true }, fill: '#ffeeaa' } },
    ...[2, 3, 4, 5].map(row => ({ row, value: 100 + row, style: { font: { italic: true }, fill: '#00ff00' } })),
  ] }))
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const button = () => page.locator('.sheet-fill-options')
const snapshot = () => page.evaluate(() => window.selectionInputFixture.snapshot())
const history = () => page.evaluate(() => window.selectionInputFixture.history())
async function select(address) { await nameBox().fill(address); await nameBox().press('Enter'); await expect(grid()).toBeFocused() }
async function drag(lastRow, lastColumn, targetRow, targetColumn) {
  const source = await page.evaluate(({ row, column }) => window.selectionInputFixture.rect(row, column, false), { row: lastRow, column: lastColumn })
  const target = await page.evaluate(({ row, column }) => window.selectionInputFixture.rect(row, column, false), { row: targetRow, column: targetColumn })
  await page.mouse.move(source.x + source.width, source.y + source.height); await page.mouse.down()
  await page.mouse.move(target.x + target.width / 2, target.y + target.height / 2, { steps: 6 }); await page.mouse.up()
}
async function choose(label) {
  await button().click()
  await page.locator('.sheet-fill-options-popup').getByRole('menuitem', { name: label, exact: true }).click()
  await expect(grid()).toBeFocused()
}

test('pointer fill and four library menu modes amend one undo entry and preserve original target data/styles', async () => {
  const original = await snapshot(); await select('A1:A2'); await drag(1, 0, 5, 0)
  await expect.poll(async () => (await snapshot()).sheets[0].cells['5:0'].value).toBe(6)
  await expect(button()).toBeVisible()
  for (const label of ['复制单元格', '仅填充格式', '不带格式填充', '填充序列']) {
    await choose(label); assert.equal((await history()).undo, 1)
    const current = (await snapshot()).sheets[0]
    if (label === '复制单元格') assert.deepEqual([2, 3, 4, 5].map(row => current.cells[`${row}:0`].value), [1, 2, 1, 2])
    if (label === '仅填充格式') { assert.equal(current.cells['2:0'].value, 102); assert.equal(current.cells['2:0'].style.fill, '#ffeeaa') }
    if (label === '不带格式填充') { assert.equal(current.cells['2:0'].value, 3); assert.equal(current.cells['2:0'].style.fill, '#00ff00') }
  }
  const final = await snapshot(); await grid().press('Control+z')
  assert.deepEqual((await snapshot()).sheets[0].cells, original.sheets[0].cells)
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).sheets[0].cells, final.sheets[0].cells)
  await expect(button()).toHaveCount(0)
})

test('inward drag clears trailing values and selection history restores the expanded fill', async () => {
  await select('A1:A2'); await drag(1, 0, 5, 0); await expect(button()).toBeVisible()
  await drag(5, 0, 2, 0)
  await expect.poll(async () => (await snapshot()).sheets[0].cells['5:0'].value).toBe(null)
  let cells = (await snapshot()).sheets[0].cells
  assert.deepEqual([0, 1, 2, 3, 4, 5].map(row => cells[`${row}:0`].value), [1, 2, 3, null, null, null])
  await expect(button()).toHaveCount(0); assert.equal((await history()).undo, 2)
  await grid().press('Control+z')
  cells = (await snapshot()).sheets[0].cells
  assert.deepEqual([0, 1, 2, 3, 4, 5].map(row => cells[`${row}:0`].value), [1, 2, 3, 4, 5, 6])
})

test('copy shortcuts use the same commands, and editing keeps text shortcuts native', async () => {
  await select('A1:A4'); await grid().press('Control+d')
  const copied = (await snapshot()).sheets[0].cells
  assert.deepEqual([0, 1, 2, 3].map(row => copied[`${row}:0`].value), [1, 1, 1, 1])
  await expect(button()).toHaveCount(0)
  await select('A1:C1'); await grid().press('Control+r')
  const current = (await snapshot()).sheets[0]
  assert.deepEqual([0, 1, 2].map(column => current.cells[`0:${column}`].value), [1, 1, 1])
  await grid().press('F2'); await page.locator('.canvas-cell-input input').fill('draft')
  const before = await history(); await page.locator('.canvas-cell-input input').press('Control+d')
  assert.deepEqual(await history(), before)
  await page.locator('.canvas-cell-input input').press('Escape')
})

test('theme, Escape, lost selection and frozen-pane position retain menu focus/visibility ownership', async () => {
  await select('A1:A2'); await drag(1, 0, 5, 0); await expect(button()).toBeVisible()
  await page.evaluate(() => window.selectionInputFixture.appearance(true, 'grid'))
  await button().focus(); await button().press('ArrowDown')
  await expect(page.locator('.sheet-fill-options-popup [role="menu"]')).toBeFocused()
  await page.locator('.sheet-fill-options-popup [role="menu"]').press('Escape'); await expect(grid()).toBeFocused()
  await select('D8'); await expect(button()).toHaveCount(0)
  const session = await page.evaluate(() => window.selectionInputFixture.fillSession())
  await page.evaluate(id => window.selectionInputFixture.fillCommand('sheet.command.refill', { sessionId: id, applyType: 'COPY' }), session.id)
  await expect(button()).toBeVisible()
  await page.evaluate(() => window.selectionInputFixture.seed({ entries: [{ row: 0, value: 1 }, { row: 1, value: 2 }],
    freeze: { rows: 1, columns: 1 } }))
  await select('A1:A2'); await drag(1, 0, 5, 0); await expect(button()).toBeVisible()
  const gridBounds = await grid().boundingBox(), menuBounds = await button().boundingBox()
  assert.ok(menuBounds.x >= gridBounds.x && menuBounds.x + menuBounds.width <= gridBounds.x + gridBounds.width)
  assert.ok(menuBounds.y >= gridBounds.y && menuBounds.y + menuBounds.height <= gridBounds.y + gridBounds.height)
})
