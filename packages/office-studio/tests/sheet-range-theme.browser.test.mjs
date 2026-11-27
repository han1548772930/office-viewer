import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'
import { OPEN_SHEET_RANGE_THEME_OPERATION, SHEET_RANGE_THEME_COMMANDS } from '../src/features/sheet-workbench/sheet-range-theme-command.ts'

// Drafts only. No dev server or browser was started while writing these scenarios.
const route = '/__sheet-range-theme__', fixture = readFileSync(new URL('./fixtures/sheet-range-theme.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-range-theme-fixture', configureServer(vite) {
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
const dialog = () => page.getByRole('dialog', { name: '范围主题', exact: true })
const snapshot = () => page.evaluate(() => window.rangeThemeFixture.snapshot())
async function select(address) { await nameBox().fill(address); await nameBox().press('Enter') }
async function open() { await page.getByRole('button', { name: '范围主题', exact: true }).click(); await expect(dialog()).toBeVisible() }
async function choose(name) {
  const combobox = dialog().getByRole('combobox', { name: '主题', exact: true })
  await combobox.fill(name); await page.locator('.ant-select-dropdown:visible').getByText(name, { exact: true }).click()
}

test('the local Ribbon entry applies a fixed range theme with one history record and preserved Canvas selection', async () => {
  await select('A1:D6'); await open(); await choose('default-last-row'); await dialog().getByRole('button', { name: '应用', exact: true }).click()
  await expect(dialog()).toBeHidden(); await expect(grid()).toBeFocused(); await expect(nameBox()).toHaveValue('A1:D6')
  const model = await snapshot(); assert.equal(model.sheets[0].rangeThemeRules[0].themeName, 'default-last-row')
  assert.equal(model.sheets[0].cells['0:0'].value, '源内容'); assert.equal(Object.keys(model.sheets[0].cells).length, 1)
  assert.equal((await page.evaluate(() => window.rangeThemeFixture.history())).undo, 1)
  assert.equal((await page.evaluate(() => window.rangeThemeFixture.effective('A6'))).bl, 1)
})

test('the Ant dialog includes registered custom themes and history plus saved reopening retain them', async () => {
  await page.evaluate(() => window.rangeThemeFixture.custom()); await select('A1:B3'); await open(); await choose('自定义 · 自定义蓝')
  await dialog().getByRole('button', { name: '应用', exact: true }).click()
  assert.equal((await page.evaluate(() => window.rangeThemeFixture.effective('A1'))).bg.rgb, '#123456')
  await page.evaluate(() => window.rangeThemeFixture.undo()); assert.equal((await snapshot()).sheets[0].rangeThemeRules?.length ?? 0, 0)
  await page.evaluate(() => { window.rangeThemeFixture.redo(); window.rangeThemeFixture.reopen() })
  assert.equal((await snapshot()).sheets[0].rangeThemeRules[0].themeName, '自定义蓝')
  assert.equal((await page.evaluate(() => window.rangeThemeFixture.effective('A1'))).bg.rgb, '#123456')
})

test('removal from a subrange removes the entire author rule and undo restores its original extent', async () => {
  await page.evaluate(() => window.rangeThemeFixture.theme('A1:D6', 'default')); await select('B2'); await open()
  await expect(dialog()).toContainText('整个应用区域'); await dialog().getByRole('button', { name: '移除主题', exact: true }).click()
  assert.deepEqual((await snapshot()).sheets[0].rangeThemeRules, [])
  await page.evaluate(() => window.rangeThemeFixture.undo())
  assert.deepEqual((await snapshot()).sheets[0].rangeThemeRules[0].range, { startRow: 0, endRow: 5, startCol: 0, endCol: 3 })
})

test('protected portions of an intersecting theme reject removal and keep the Ant modal and error visible', async () => {
  await page.evaluate(() => { window.rangeThemeFixture.theme('A1:D6', 'default'); window.rangeThemeFixture.protectTail() })
  await select('A2'); await open(); const before = await snapshot(), history = await page.evaluate(() => window.rangeThemeFixture.history())
  await dialog().getByRole('button', { name: '移除主题', exact: true }).click(); await expect(dialog()).toBeVisible(); await expect(dialog()).toContainText('保护')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await page.evaluate(() => window.rangeThemeFixture.history()), history)
})

test('a changed model rejects the captured dialog revision and sheet switching closes it without applying stale rules', async () => {
  await select('A1:B3'); await open(); await page.evaluate(() => window.rangeThemeFixture.edit('外部更新'))
  await dialog().getByRole('button', { name: '应用', exact: true }).click(); await expect(dialog()).toContainText('已更改')
  assert.equal((await snapshot()).sheets[0].rangeThemeRules?.length ?? 0, 0)
  await page.evaluate(() => window.rangeThemeFixture.switchSheet()); await expect(dialog()).toBeHidden()
  assert.equal((await snapshot()).activeSheetId, 'other'); assert.equal((await snapshot()).sheets[1].rangeThemeRules?.length ?? 0, 0)
})

test('range-theme input sessions retain drafts and dialog cancellation returns Canvas focus through the Ant afterClose callback', async () => {
  for (const host of ['cell', 'formula']) for (const variant of ['classic', 'grid']) {
    await page.goto(baseURL); await expect(grid()).toBeVisible()
    await page.evaluate(value => window.rangeThemeFixture.appearance(false, value), variant); await select('A1')
    const before = await snapshot(), depth = await page.evaluate(() => window.rangeThemeFixture.history())
    const input = host === 'cell' ? page.locator('.canvas-cell-input input, .canvas-cell-input textarea') : page.getByRole('textbox', { name: '公式编辑', exact: true })
    if (host === 'cell') await grid().press('F2')
    if (host === 'formula') await input.focus()
    await input.fill('未提交主题草稿')
    const rangeTheme = page.locator('[data-ribbon-item="sheet.operation.open-range-theme"]:visible').first().getByRole('button')
    for (const composing of [false, true]) {
      if (composing) {
        await input.dispatchEvent('compositionstart')
        await page.evaluate(value => window.rangeThemeFixture.appearance(true, value), variant)
      }
      await expect(rangeTheme).toBeDisabled(); await rangeTheme.evaluate(element => element.click())
      for (const intent of [
        { commandId: OPEN_SHEET_RANGE_THEME_OPERATION, params: undefined },
        { commandId: SHEET_RANGE_THEME_COMMANDS.set, params: { themeName: 'default' } },
        { commandId: SHEET_RANGE_THEME_COMMANDS.remove, params: undefined },
      ]) {
        await page.getByTestId('workbench-ribbon').evaluate((element, value) => element.__vueParentComponent.emit('execute', value.commandId, value.params), intent)
      }
      await expect(dialog()).toBeHidden(); await expect(input).toHaveValue('未提交主题草稿'); await expect(grid()).not.toBeFocused()
      assert.deepEqual(await snapshot(), before); assert.deepEqual(await page.evaluate(() => window.rangeThemeFixture.history()), depth)
    }
    await input.dispatchEvent('compositionend', { data: '未提交主题草稿' }); await expect(rangeTheme).toBeDisabled()
    await input.press('Escape'); await expect(grid()).toBeFocused(); await expect(rangeTheme).toBeEnabled(); await open()
    await expect(rangeTheme).toBeDisabled()
    await page.getByTestId('workbench-ribbon').evaluate((element, commandId) => element.__vueParentComponent.emit('execute', commandId, { themeName: 'default' }), SHEET_RANGE_THEME_COMMANDS.set)
    await expect(dialog()).toBeVisible(); assert.deepEqual(await snapshot(), before)
    assert.deepEqual(await page.evaluate(() => window.rangeThemeFixture.history()), depth)
    await dialog().getByRole('button', { name: '取消', exact: true }).click()
    await expect(dialog()).toBeHidden(); await expect(grid()).toBeFocused(); await expect(nameBox()).toHaveValue('A1')
    await expect(rangeTheme).toBeEnabled(); assert.deepEqual(await snapshot(), before)
    assert.deepEqual(await page.evaluate(() => window.rangeThemeFixture.history()), depth)
  }
})

test('grid layout, dark theme and Canvas Ant input use the applied style and return focus after Escape', async () => {
  await page.evaluate(() => window.rangeThemeFixture.appearance(true, 'grid')); await select('A1:D6'); await open()
  const background = await dialog().locator('.ant-modal-content').evaluate(element => getComputedStyle(element).backgroundColor)
  assert.notEqual(background, 'rgb(255, 255, 255)'); await dialog().getByRole('button', { name: '应用', exact: true }).click()
  await expect(grid()).toBeFocused(); await select('A1'); await grid().press('F2')
  const input = page.locator('.canvas-cell-input input, .canvas-cell-input textarea'); await expect(input).toBeFocused()
  assert.ok((await input.evaluate(element => element.className)).includes('ant-')); await input.press('Escape')
  await expect(grid()).toBeFocused(); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '源内容')
})
