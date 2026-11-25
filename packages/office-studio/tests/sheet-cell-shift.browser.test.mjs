import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Drafts only. These browser/Canvas/Ant scenarios have not been executed.
const route = '/__sheet-cell-shift__', fixture = readFileSync(new URL('./fixtures/sheet-cell-shift.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-cell-shift-fixture', configureServer(vite) {
      vite.middlewares.use(route, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(route, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${route}`; browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1600, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const dialog = () => page.getByRole('dialog', { name: '合并单元格', exact: true })
const snapshot = () => page.evaluate(() => window.cellShiftFixture.snapshot())
const history = () => page.evaluate(() => window.cellShiftFixture.history())
async function select(address = 'B2') {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter'); await expect(grid()).toBeFocused()
}
async function openInsertMenu(primary = false) {
  if (primary) await page.evaluate(() => window.cellShiftFixture.primary()); else await select()
  await grid().click({ button: 'right', position: { x: 240, y: 76 } })
  await page.getByRole('menuitem', { name: '插入', exact: true }).hover()
}
async function insertDown(primary = false) {
  await openInsertMenu(primary)
  await page.getByRole('menuitem', { name: '下移单元格', exact: true }).click()
}

test('real context-menu insertion uses the active non-last range and publishes source, reference and history together', async () => {
  const before = await snapshot(), depth = await history(); await insertDown(true)
  const after = await snapshot(); assert.equal(after.sheets[0].cells['2:1'].value, '移动内容')
  assert.equal(after.sheets[0].cells['1:2'].value, '其他列'); assert.equal(after.sheets[1].cells['0:0'].formula, '=Data!B3')
  assert.deepEqual(after.sheets[0].cells['2:1'].metadata, { note: '保留批注' }); assert.equal(after.sheets[0].rangeThemeRules.length, 3)
  assert.equal((await history()).undo, depth.undo + 1); await expect(grid()).toBeFocused()
  await grid().press('Control+z'); assert.deepEqual((await snapshot()).sheets, before.sheets)
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).sheets, after.sheets)
})

test('canceling the Ant merge confirmation leaves content and history intact and returns focus after closing', async () => {
  await page.evaluate(() => window.cellShiftFixture.merge()); const before = await snapshot(), depth = await history()
  await insertDown(); await expect(dialog()).toBeVisible(); await expect(dialog()).toContainText('A3:C4')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
  await dialog().getByRole('button', { name: '取消', exact: true }).click(); await expect(dialog()).toBeHidden(); await expect(grid()).toBeFocused()
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
})

test('accepting the captured merge request commits one complete edit and undo restores the merge and references', async () => {
  await page.evaluate(() => window.cellShiftFixture.merge()); const before = await snapshot(), depth = await history()
  await insertDown(); await expect(dialog()).toBeVisible(); await dialog().getByRole('button', { name: '继续', exact: true }).click()
  await expect(dialog()).toBeHidden(); await expect(grid()).toBeFocused(); const after = await snapshot()
  assert.deepEqual(after.sheets[0].mergedRanges, []); assert.equal(after.sheets[0].cells['2:1'].value, '移动内容')
  assert.equal(after.sheets[1].cells['0:0'].formula, '=Data!B3'); assert.equal((await history()).undo, depth.undo + 1)
  await grid().press('Control+z'); assert.deepEqual((await snapshot()).sheets, before.sheets)
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).sheets, after.sheets)
})

test('a newer content revision keeps the pending dialog and its visible error without applying the stale move', async () => {
  await page.evaluate(() => window.cellShiftFixture.merge()); await insertDown(); await expect(dialog()).toBeVisible()
  await page.evaluate(() => window.cellShiftFixture.newer()); const newer = await snapshot(), depth = await history()
  await dialog().getByRole('button', { name: '继续', exact: true }).click(); await expect(dialog().getByRole('alert')).toBeVisible()
  await expect(dialog()).toContainText('更改'); assert.deepEqual(await snapshot(), newer); assert.deepEqual(await history(), depth)
  await dialog().getByRole('button', { name: '取消', exact: true }).click(); await expect(grid()).toBeFocused()
})

test('switching worksheet dismisses the captured confirmation and never shifts either replacement content or the old merge', async () => {
  await page.evaluate(() => window.cellShiftFixture.merge()); const before = await snapshot(), depth = await history()
  await insertDown(); await expect(dialog()).toBeVisible(); await page.evaluate(() => window.cellShiftFixture.switchSheet())
  await expect(dialog()).toBeHidden(); const after = await snapshot()
  assert.equal(after.activeSheetId, 'report'); assert.deepEqual(after.sheets, before.sheets); assert.deepEqual(await history(), depth)
})

test('protected tail members disable the context-menu shift before any merge request or content history is published', async () => {
  await page.evaluate(() => { window.cellShiftFixture.merge(); window.cellShiftFixture.protect() })
  const before = await snapshot(), depth = await history(); await openInsertMenu()
  await expect(page.getByRole('menuitem', { name: '下移单元格', exact: true })).toBeDisabled()
  await expect(dialog()).toBeHidden(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
})

test('the real confirm menu rejects a filtered tail before a merge dialog or content update', async () => {
  await page.evaluate(() => { window.cellShiftFixture.merge(); window.cellShiftFixture.filtered() })
  const before = await snapshot(), depth = await history(); await insertDown()
  await expect(dialog()).toBeHidden(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
  assert.ok((await page.evaluate(() => window.cellShiftFixture.notifications())).some(event => /筛选隐藏行/.test(event.message)))
})

test('dark/grid confirmation preserves Ant focus and the Canvas input still honors synthetic composition after cancellation', async () => {
  await page.evaluate(() => window.cellShiftFixture.merge()); await insertDown(); await expect(dialog()).toBeVisible()
  const button = dialog().getByRole('button', { name: '继续', exact: true }); await button.focus()
  const light = await dialog().evaluate(element => getComputedStyle(element.querySelector('.ant-modal-content') ?? element).backgroundColor)
  await page.evaluate(() => window.cellShiftFixture.appearance(true, 'grid')); await expect(button).toBeFocused()
  const dark = await dialog().evaluate(element => getComputedStyle(element.querySelector('.ant-modal-content') ?? element).backgroundColor)
  assert.notEqual(dark, light)
  await dialog().getByRole('button', { name: '取消', exact: true }).click(); await expect(grid()).toBeFocused(); await grid().press('F2')
  const input = page.locator('.canvas-cell-input input, .canvas-cell-input textarea'); await expect(input).toBeFocused()
  assert.match(await input.evaluate(element => element.className), /ant-/)
  await input.dispatchEvent('compositionstart', { data: '中' }); await input.fill('中文编辑'); await input.press('Enter')
  await expect(input).toBeFocused(); assert.equal((await snapshot()).sheets[0].cells['1:1'].value, '移动内容')
  await input.dispatchEvent('compositionend', { data: '中文编辑' }); await input.press('Enter'); await expect(grid()).toBeFocused()
  assert.equal((await snapshot()).sheets[0].cells['1:1'].value, '中文编辑')
})
