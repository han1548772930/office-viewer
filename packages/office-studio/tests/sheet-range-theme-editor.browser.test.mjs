import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Browser drafts only. Canvas2D is forced; composition events do not verify real system IME.
const route = '/__sheet-range-theme-editor__', fixture = readFileSync(new URL('./fixtures/sheet-range-theme.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-range-theme-editor-fixture', configureServer(vite) {
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
  errors = []; page = await browser.newPage({ viewport: { width: 1600, height: 1000 } }); page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const dialog = () => page.getByRole('dialog', { name: '范围主题', exact: true })
const nameInput = () => dialog().getByRole('textbox', { name: '主题名称', exact: true })
const save = () => dialog().getByRole('button', { name: '保存并应用', exact: true })
const snapshot = () => page.evaluate(() => window.rangeThemeFixture.snapshot())
const history = () => page.evaluate(() => window.rangeThemeFixture.history())
async function select(address) { const box = page.getByRole('textbox', { name: '名称框', exact: true }); await box.fill(address); await box.press('Enter') }
async function open() { await page.getByRole('button', { name: '范围主题', exact: true }).click(); await expect(dialog()).toBeVisible() }
async function choose(label, value) {
  await dialog().getByRole('combobox', { name: label, exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText(value, { exact: true }).click()
}
async function existingEditor() {
  await page.evaluate(() => window.rangeThemeFixture.editorUsage()); await select('A1:B4'); await open()
  await dialog().getByRole('button', { name: '编辑主题', exact: true }).click(); await expect(nameInput()).toHaveValue('编辑主题')
}

test('the real Ant editor saves all nine sections as a copied theme with one history record', async () => {
  await select('A1:D6'); await open(); await dialog().getByRole('button', { name: '编辑副本', exact: true }).click()
  await nameInput().fill('九区段副本')
  const sections = ['整个区域', '第一列条纹', '第二列条纹', '第一行条纹', '第二行条纹', '首列', '末列', '表头行', '末行']
  for (let index = 0; index < sections.length; index++) {
    await choose('样式区段', sections[index]); await dialog().getByRole('spinbutton', { name: '字号', exact: true }).fill(String(index + 11))
  }
  await choose('样式区段', '表头行'); await choose('加粗', '关闭'); await choose('填充颜色 模式', '无颜色')
  await save().click(); await expect(dialog()).toBeHidden(); await expect(grid()).toBeFocused()
  const model = await snapshot(), theme = model.rangeThemes['九区段副本']
  assert.deepEqual([theme.wholeStyle, theme.firstColumnStyle, theme.secondColumnStyle, theme.firstRowStyle, theme.secondRowStyle,
    theme.headerColumnStyle, theme.lastColumnStyle, theme.headerRowStyle, theme.lastRowStyle].map(style => style.font.size), [11, 12, 13, 14, 15, 16, 17, 18, 19])
  assert.equal(theme.headerRowStyle.font.bold, false); assert.equal(theme.headerRowStyle.fill, null); assert.equal((await history()).undo, 1)
  assert.equal(Object.keys(model.sheets[0].cells).length, 1); assert.equal(model.sheets[0].rangeThemeRules[0].themeName, '九区段副本')
  await page.evaluate(() => { window.rangeThemeFixture.undo(); window.rangeThemeFixture.redo(); window.rangeThemeFixture.reopen() })
  assert.deepEqual((await snapshot()).rangeThemes['九区段副本'], theme)
})

test('editing a used custom theme updates both sheets while preserving rule IDs and unknown source data', async () => {
  await existingEditor(); const before = await snapshot(), records = await history()
  await dialog().getByRole('spinbutton', { name: '字号', exact: true }).fill('25'); await choose('样式区段', '表头行'); await choose('加粗', '关闭')
  await save().click(); await expect(dialog()).toBeHidden()
  const model = await snapshot(); assert.equal((await history()).undo, records.undo + 1)
  assert.deepEqual(model.sheets.map(sheet => sheet.rangeThemeRules), before.sheets.map(sheet => sheet.rangeThemeRules))
  assert.equal((await page.evaluate(() => window.rangeThemeFixture.otherEffective('A2'))).fs, 25)
  assert.equal((await page.evaluate(() => window.rangeThemeFixture.otherEffective('A1'))).bl, 0)
  assert.deepEqual(model.rangeThemes['编辑主题'].wholeStyle.metadata.sourceStyle.ol, { vendor: 'keep' })
})

test('a protected reference on another sheet rejects save-and-apply and keeps the edited draft visible', async () => {
  await existingEditor(); await page.evaluate(() => window.rangeThemeFixture.protectOtherTheme())
  const before = await snapshot(), records = await history(); await dialog().getByRole('spinbutton', { name: '字号', exact: true }).fill('30')
  await save().click(); await expect(dialog()).toContainText('保护'); await expect(nameInput()).toHaveValue('编辑主题')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), records)
})

test('a history reset leaves the model unchanged and rejects the captured editor confirmation', async () => {
  await select('A1:B3'); await open(); await dialog().getByRole('button', { name: '新建主题', exact: true }).click(); await nameInput().fill('迟到主题')
  const before = await snapshot(); await page.evaluate(() => window.rangeThemeFixture.clearHistory()); await save().click()
  await expect(dialog()).toContainText('历史已更改'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})

test('reserved names and active composition prevent committing a draft until the user corrects them', async () => {
  await select('A1:B3'); await open(); await dialog().getByRole('button', { name: '新建主题', exact: true }).click(); await nameInput().fill('default')
  await save().click(); await expect(dialog()).toContainText('内置主题名称已保留'); assert.equal((await history()).undo, 0)
  await nameInput().fill('输入法主题'); await nameInput().dispatchEvent('compositionstart', { data: '中' }); await expect(save()).toBeDisabled()
  assert.equal((await history()).undo, 0); await nameInput().dispatchEvent('compositionend', { data: '中文' }); await expect(save()).toBeEnabled()
  await save().click(); await expect(dialog()).toBeHidden(); assert.ok((await snapshot()).rangeThemes['输入法主题'])
})

test('the grid and dark-theme editor uses Ant inputs and cancellation restores Canvas focus without saving a draft', async () => {
  await page.evaluate(() => window.rangeThemeFixture.appearance(true, 'grid')); await select('A1:B3'); await open()
  await dialog().getByRole('button', { name: '新建主题', exact: true }).click(); await nameInput().fill('取消草稿')
  assert.ok((await nameInput().getAttribute('class')).includes('ant-input'))
  assert.ok((await dialog().getByRole('textbox', { name: '字体', exact: true }).getAttribute('class')).includes('ant-input'))
  const background = await dialog().locator('.ant-modal-content').evaluate(element => getComputedStyle(element).backgroundColor)
  assert.notEqual(background, 'rgb(255, 255, 255)'); await dialog().getByRole('button', { name: '取消', exact: true }).click()
  await expect(dialog()).toBeHidden(); await expect(grid()).toBeFocused(); assert.equal((await history()).undo, 0)
  assert.equal(Object.hasOwn((await snapshot()).rangeThemes ?? {}, '取消草稿'), false)
})

test('restore and clear-section actions remain local until save-and-apply records the final author change', async () => {
  await existingEditor(); const before = await snapshot(), records = await history()
  const size = dialog().getByRole('spinbutton', { name: '字号', exact: true }); await size.fill('42')
  await dialog().getByRole('button', { name: '恢复初始配置', exact: true }).click(); await expect(size).toHaveValue('18')
  await dialog().getByRole('button', { name: '清空区段', exact: true }).click(); await expect(size).toHaveValue('')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), records); await save().click(); await expect(dialog()).toBeHidden()
  assert.equal((await snapshot()).rangeThemes['编辑主题'].wholeStyle, null); assert.equal((await history()).undo, records.undo + 1)
})
