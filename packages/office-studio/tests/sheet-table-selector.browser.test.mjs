import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'
import { SHEET_TABLE_COMMANDS } from '../src/features/sheet-tables/sheet-table-create.ts'

// Implementation-phase drafts. No dev server, browser or test runner has been started.
const route = '/__sheet-table-selector__', fixture = readFileSync(new URL('./fixtures/sheet-table-selector.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-table-selector-fixture', configureServer(vite) {
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
  errors = []; page = await browser.newPage({ viewport: { width: 1600, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const dialog = () => page.getByRole('dialog', { name: '创建表格', exact: true })
const rangeInput = () => dialog().getByRole('textbox', { name: '表格范围', exact: true })
const snapshot = () => page.evaluate(() => window.tableFixture.snapshot())
const history = () => page.evaluate(() => window.tableFixture.history())
async function select(address) { await nameBox().fill(address); await nameBox().press('Enter') }
async function open() {
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  await page.getByRole('button', { name: '创建表格', exact: true }).click(); await expect(dialog()).toBeVisible()
}
async function appearance(variant, dark = false) { await page.evaluate(({ variant, dark }) => window.tableFixture.appearance(dark, variant), { variant, dark }) }

for (const variant of ['classic', 'grid']) {
  test(`${variant}: creation expands the last single-cell range, commits once, and survives undo/redo plus internal reopen`, async () => {
    await appearance(variant); await select('B2'); const before = await snapshot(); await open()
    await expect(rangeInput()).toHaveValue('A1:B3'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
    await dialog().getByRole('button', { name: '确认', exact: true }).click(); await expect(dialog()).toBeHidden(); await expect(grid()).toBeFocused()
    const added = await snapshot(); assert.equal(added.tables.length, 1); assert.deepEqual(added.tables[0].columns.map(column => column.name), ['部门', 'TRUE'])
    assert.deepEqual(added.sheets[0].cells, before.sheets[0].cells); assert.equal((await history()).undo, 1)
    await page.evaluate(() => window.tableFixture.undo()); assert.equal((await snapshot()).tables?.length ?? 0, 0)
    await page.evaluate(() => window.tableFixture.redo()); assert.deepEqual((await snapshot()).tables, added.tables)
    await page.evaluate(() => window.tableFixture.reopen()); await expect(grid()).toBeVisible(); assert.deepEqual((await snapshot()).tables, added.tables)
  })

  test(`${variant}: cancel keeps the exact model/history and returns focus after the Ant modal animation`, async () => {
    await appearance(variant, true); await select('A1:B3'); const before = await snapshot(), oldHistory = await history(); await open()
    await rangeInput().fill('D4:E8'); await dialog().getByRole('button', { name: '取消', exact: true }).click()
    await expect(dialog()).toBeHidden(); await expect(grid()).toBeFocused(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await expect(nameBox()).toHaveValue('A1:B3')
  })

  test(`${variant}: multiple selections use the last range and invalid single-row drafts remain correctable`, async () => {
    await appearance(variant)
    await page.evaluate(() => window.tableFixture.selections([
      { startRow: 6, endRow: 7, startColumn: 4, endColumn: 5 }, { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    ]))
    await open(); await expect(rangeInput()).toHaveValue('A1:B3'); await rangeInput().fill('A1:B1')
    await dialog().getByRole('button', { name: '确认', exact: true }).click(); await expect(dialog()).toBeVisible()
    await expect(dialog().locator('.ant-alert-error')).toBeVisible(); assert.equal((await history()).undo, 0)
    await rangeInput().fill('A1:B3'); await dialog().getByRole('button', { name: '确认', exact: true }).click()
    await expect(dialog()).toBeHidden(); assert.equal((await history()).undo, 1)
  })

  test(`${variant}: merges, existing tables and protected cells in the proposed range reject without changing content/history`, async () => {
    await appearance(variant); await page.evaluate(() => window.tableFixture.merge()); await select('D1:E3'); await open()
    const beforeMerge = await snapshot(); await rangeInput().fill('A1:B3'); await dialog().getByRole('button', { name: '确认', exact: true }).click()
    await expect(dialog().locator('.ant-alert-error')).toBeVisible(); assert.deepEqual(await snapshot(), beforeMerge); assert.equal((await history()).undo, 0)
    await dialog().getByRole('button', { name: '取消', exact: true }).click(); await expect(dialog()).toBeHidden()
    await page.reload(); await expect(grid()).toBeVisible(); await appearance(variant); await select('A1:B3'); await open()
    await dialog().getByRole('button', { name: '确认', exact: true }).click(); await expect(dialog()).toBeHidden()
    await select('D1:E3'); await open(); const existing = await snapshot(), existingHistory = await history()
    await rangeInput().fill('B2:C4'); await dialog().getByRole('button', { name: '确认', exact: true }).click()
    await expect(dialog().locator('.ant-alert-error')).toBeVisible(); assert.deepEqual(await snapshot(), existing); assert.deepEqual(await history(), existingHistory)
    await dialog().getByRole('button', { name: '取消', exact: true }).click(); await expect(dialog()).toBeHidden()
    await page.reload(); await expect(grid()).toBeVisible(); await appearance(variant); await page.evaluate(() => window.tableFixture.protectTail())
    await select('D1:E3'); await open(); const protectedModel = await snapshot(); await rangeInput().fill('A1:B3')
    await dialog().getByRole('button', { name: '确认', exact: true }).click(); await expect(dialog().locator('.ant-alert-error')).toBeVisible()
    assert.deepEqual(await snapshot(), protectedModel); assert.equal((await history()).undo, 0)
  })

  test(`${variant}: normal cell drafts commit before opening while active IME keeps the draft and rejects toolbar/API opening`, async () => {
    await appearance(variant); await select('A1'); await grid().focus(); await grid().press('F2')
    const editor = page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
    await editor.fill('新表头'); await open(); await expect(rangeInput()).toHaveValue('A1:B3')
    assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '新表头'); assert.equal((await history()).undo, 1)
    await dialog().getByRole('button', { name: '取消', exact: true }).click(); await expect(dialog()).toBeHidden()
    await select('A1'); await grid().focus(); await grid().press('F2'); await editor.fill('未提交草稿'); await editor.dispatchEvent('compositionstart', { data: '未' })
    await page.getByRole('tab', { name: '数据', exact: true }).click(); await expect(page.getByRole('button', { name: '创建表格', exact: true })).toBeDisabled()
    const before = await snapshot(); await page.evaluate(id => window.tableFixture.command(id), SHEET_TABLE_COMMANDS.openSelector)
    await expect(dialog()).toBeHidden(); await expect(editor).toHaveValue('未提交草稿'); assert.deepEqual(await snapshot(), before)
    await editor.dispatchEvent('compositionend', { data: '未提交草稿' }); await editor.press('Escape')
  })

  test(`${variant}: composition within the range input cannot confirm or cancel until composition finishes`, async () => {
    await appearance(variant); await select('A1:B3'); await open(); await rangeInput().dispatchEvent('compositionstart', { data: 'A' })
    await expect(dialog().getByRole('button', { name: '确认', exact: true })).toBeDisabled()
    await expect(dialog().getByRole('button', { name: '取消', exact: true })).toBeDisabled(); await rangeInput().press('Escape')
    await expect(dialog()).toBeVisible(); assert.equal((await history()).undo, 0)
    await rangeInput().dispatchEvent('compositionend', { data: 'A1:B3' }); await rangeInput().fill('A1:B3')
    await dialog().getByRole('button', { name: '确认', exact: true }).click(); await expect(dialog()).toBeHidden(); assert.equal((await history()).undo, 1)
  })

  test(`${variant}: sheet switches close the owner, and same-ID content/history replacement cannot commit a stale dialog`, async () => {
    await appearance(variant); await select('A1:B3'); await open(); await page.evaluate(() => window.tableFixture.switchSheet())
    await expect(dialog()).toBeHidden(); assert.equal((await snapshot()).tables?.length ?? 0, 0); assert.equal((await history()).undo, 0)
    await page.reload(); await expect(grid()).toBeVisible(); await appearance(variant); await select('A1:B3'); await open()
    await page.evaluate(() => window.tableFixture.clearHistory()); await expect(dialog().getByRole('button', { name: '确认', exact: true })).toBeDisabled()
    await dialog().getByRole('button', { name: '取消', exact: true }).click(); await expect(dialog()).toBeHidden(); assert.equal((await snapshot()).tables?.length ?? 0, 0)
    await open(); await page.evaluate(() => window.tableFixture.replaceSameId()); await expect(dialog().getByRole('button', { name: '确认', exact: true })).toBeDisabled()
    assert.equal((await snapshot()).tables?.length ?? 0, 0)
  })
}
