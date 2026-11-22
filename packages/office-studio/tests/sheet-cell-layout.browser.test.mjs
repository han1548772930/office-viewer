import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-cell-layout-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Sheets cell layout and input</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model'
import { toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
import { createSpreadsheetScene } from '/packages/office-studio/src/rendering/spreadsheet/scene'
import { SpreadsheetGeometry } from '/packages/office-studio/src/rendering/spreadsheet/layout'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankWorkbook()
initial.id = 'sheet-cell-layout-fixture'; initial.name = '单元格布局'; initial.calculationMode = 'manual'
const initialSheet = initial.sheets[0]
initialSheet.rowCount = 100; initialSheet.columnCount = 26
initialSheet.defaultRowHeight = 24; initialSheet.defaultColumnWidth = 80
initialSheet.cells['0:0'] = createCell(0, 0, '甲乙丙')
const studio = new OfficeStudio({ workbook: initial, currentUserId: 'local-user' })
const api = FOffice.newAPI(studio)
const view = ref(toExcelWorkbookView(studio.editor.workbook.model))
const dark = ref(false), variant = ref('classic'), notifications = []
watch(() => studio.editor.workbook.model, model => { view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
window.sheetLayoutFixture = {
  snapshot: () => copy(studio.editor.workbook.model),
  history: () => { const unit = studio.editor.history.units[studio.activeWorkbookId]; return { undo: unit.undo.length, redo: unit.redo.length } },
  notifications: () => copy(notifications),
  appearance: (isDark, layout = 'classic') => { dark.value = isDark; variant.value = layout },
  seed: (entries, options = {}) => {
    const next = copy(initial), sheet = next.sheets[0]
    sheet.cells = {}
    Object.assign(sheet, options)
    for (const { row = 0, column = 0, value, style = {} } of entries) {
      const cell = createCell(row, column, value)
      cell.style = { ...cell.style, ...style, font: { ...cell.style.font, ...style.font } }
      sheet.cells[row + ':' + column] = cell
    }
    studio.editor.workbook.replace(next)
    studio.editor.history.clear(studio.activeWorkbookId)
  },
  layout: (method, value, address = 'A1') => { const range = api.getActiveWorkbook().getActiveSheet().getRange(address); range[method](value) },
  list: () => {
    api.getActiveWorkbook().getActiveSheet().setDataValidation('A1', { type: 'list', formula1: '"Alpha,Beta"', allowBlank: false, errorStyle: 'stop', error: '只能输入 Alpha 或 Beta。' })
    studio.editor.history.clear(studio.activeWorkbookId)
  },
  denyEdit: () => {
    const next = copy(studio.editor.workbook.model)
    next.sheets[0].rangeProtectionRules = [{ id: 'deny-input', permissionId: 'deny-input-permission', ownerUserId: 'remote-owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }] }]
    studio.editor.workbook.replace(next)
  },
  reopen: snapshot => { api.createWorkbook({ ...JSON.parse(snapshot), id: 'sheet-cell-layout-reopened' }) },
  scene: () => {
    const element = document.querySelector('.spreadsheet-canvas'), scroll = element.querySelector('.canvas-scroll')
    const geometry = new SpreadsheetGeometry(view.value.sheets[0])
    return createSpreadsheetScene(geometry, { width: element.clientWidth, height: element.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 1 }, { activeCell: 'A1', selectedCells: [] })
  },
}
createApp({ render: () => h(ConfigProvider, {
  theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } },
}, { default: () => h(ExcelWorkbench, {
  modelValue: view.value, studio, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
  'onUpdate:modelValue': value => { studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value) },
  onUndo: () => api.getActiveWorkbook().undo(), onRedo: () => api.getActiveWorkbook().redo(),
  onNotify: value => { notifications.push(copy(value)) },
}) }) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-cell-layout-fixture',
    configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => {
      try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
    }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []
  page = await browser.newPage({ viewport: { width: 1800, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => {
    Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined })
    const fillRect = CanvasRenderingContext2D.prototype.fillRect, fillText = CanvasRenderingContext2D.prototype.fillText
    CanvasRenderingContext2D.prototype.fillRect = function (x, y, width, height) {
      if (x === 0 && y === 0 && width > 100 && height > 100) this.canvas.sheetTextDraws = []
      return fillRect.call(this, x, y, width, height)
    }
    CanvasRenderingContext2D.prototype.fillText = function (text, x, y, maxWidth) {
      const matrix = this.getTransform()
      const list = this.canvas.sheetTextDraws ??= []
      list.push({ text, x, y, font: this.font, maxWidth, matrix: { a: matrix.a, b: matrix.b, c: matrix.c, d: matrix.d, e: matrix.e, f: matrix.f } })
      return fillText.call(this, text, x, y, maxWidth)
    }
  })
  await page.goto(baseURL)
  await expect(grid()).toBeVisible(); await expect(canvas()).toHaveCount(1)
  await canvas().evaluate(element => { element.dataset.layoutTestMount = 'original' })
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const canvas = () => page.locator('.canvas-surface canvas')
const editor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const formula = () => page.getByRole('textbox', { name: '公式编辑', exact: true })
const snapshot = () => page.evaluate(() => window.sheetLayoutFixture.snapshot())
const history = () => page.evaluate(() => window.sheetLayoutFixture.history())
const draws = () => canvas().evaluate(element => element.sheetTextDraws ?? [])
const scene = () => page.evaluate(() => window.sheetLayoutFixture.scene())
async function select(address = 'A1') {
  const name = page.getByRole('textbox', { name: '名称框', exact: true })
  await name.fill(address); await name.press('Enter'); await expect(grid()).toBeFocused()
}
async function edit(address = 'A1') { await select(address); await grid().press('F2'); await expect(editor()).toBeFocused(); return editor() }
async function chooseMenu(id, label) {
  const item = page.locator('[data-ribbon-item="' + id + '"]:visible').first()
  await expect(item).toBeVisible(); await item.getByRole('button').click()
  await page.locator('.ribbon-command-menu:visible').getByRole('menuitem', { name: label, exact: true }).click()
}
async function command(id) { await page.locator('[data-ribbon-item="' + id + '"]:visible').first().getByRole('button').click() }
async function caret(start, end = start) { await editor().evaluate((element, selection) => element.setSelectionRange(selection[0], selection[1]), [start, end]) }
async function caretRange() { return editor().evaluate(element => [element.selectionStart, element.selectionEnd]) }
async function key(info) { return editor().evaluate((element, data) => { const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...data }); element.dispatchEvent(event); return event.defaultPrevented }, info) }

test('Alt+Enter replaces the native selection, switches to TextArea and commits in place with one history entry', async () => {
  await edit(); await caret(1, 2); await editor().press('Alt+Enter')
  await expect(page.locator('.canvas-cell-input textarea')).toBeFocused()
  await expect(editor()).toHaveValue('甲\n丙'); assert.deepEqual(await caretRange(), [2, 2])
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.ok(await page.locator('.canvas-cell-input').evaluate(element => element.getBoundingClientRect().height > 24))
  await editor().press('Control+Enter')
  await expect(editor()).toHaveCount(0); await expect(grid()).toBeFocused()
  await expect(page.getByRole('textbox', { name: '名称框', exact: true })).toHaveValue('A1')
  const saved = await snapshot(); assert.equal(saved.sheets[0].cells['0:0'].value, '甲\n丙')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await command('sheet.command.undo'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '甲乙丙')
  await command('sheet.command.redo'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '甲\n丙')
  await page.evaluate(value => window.sheetLayoutFixture.reopen(JSON.stringify(value)), saved)
  await edit(); await expect(page.locator('.canvas-cell-input textarea')).toBeFocused(); await expect(editor()).toHaveValue('甲\n丙')
  await editor().press('Escape'); await expect(canvas()).toHaveAttribute('data-layout-test-mount', 'original')
})

test('multiline paste inside a cell preserves tabs and blank lines without changing adjacent cells', async () => {
  await page.evaluate(() => window.sheetLayoutFixture.seed([{ value: 'head tail' }, { column: 1, value: 'neighbor' }]))
  await edit(); await caret(5, 9)
  const prevented = await editor().evaluate(element => {
    const data = new DataTransfer(); data.setData('text/plain', 'a\tb\r\nc\rd\n')
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data })
    element.dispatchEvent(event); return event.defaultPrevented
  })
  assert.equal(prevented, true)
  await expect(page.locator('.canvas-cell-input textarea')).toBeFocused()
  await expect(editor()).toHaveValue('head a\tb\nc\nd\n'); assert.deepEqual(await caretRange(), [13, 13])
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await editor().press('Enter')
  const model = await snapshot()
  assert.equal(model.sheets[0].cells['0:0'].value, 'head a\tb\nc\nd\n')
  assert.equal(model.sheets[0].cells['0:1'].value, 'neighbor')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(page.getByRole('textbox', { name: '名称框', exact: true })).toHaveValue('A2')
})

test('removing a draft line break keeps the current TextArea and cancel restores the original single-line bridge', async () => {
  await edit(); await editor().press('Alt+Enter'); await expect(page.locator('.canvas-cell-input textarea')).toBeFocused()
  await editor().fill('single'); await expect(page.locator('.canvas-cell-input textarea')).toBeFocused()
  await editor().press('Escape'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '甲乙丙')
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await edit(); await expect(page.locator('.canvas-cell-input input')).toBeFocused(); await expect(editor()).toHaveValue('甲乙丙')
})

test('formula and canvas bridges keep one draft while multiline changes and theme/layout changes preserve focus', async () => {
  await edit(); await editor().fill('draft'); await formula().focus()
  await expect(formula()).toHaveValue('draft'); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await formula().fill('formula\nbar')
  await expect(formula()).toBeFocused(); await expect(page.locator('.canvas-cell-input textarea')).toHaveValue('formula\nbar')
  await page.evaluate(() => window.sheetLayoutFixture.appearance(true, 'grid'))
  await expect(formula()).toBeFocused(); await expect(formula()).toHaveValue('formula\nbar')
  await editor().click(); await caret(3)
  await page.evaluate(() => window.sheetLayoutFixture.appearance(false, 'classic'))
  await expect(editor()).toBeFocused(); assert.deepEqual(await caretRange(), [3, 3])
  await editor().press('Tab')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'formula\nbar')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('TextArea retains validation choices and an IME Enter cannot accept a highlighted option or commit', async () => {
  await page.evaluate(() => { window.sheetLayoutFixture.seed([{ value: 'Alpha', style: { wrap: true, wrapStrategy: 'wrap' } }]); window.sheetLayoutFixture.list() })
  await edit(); await expect(page.locator('.canvas-cell-input textarea')).toBeFocused()
  await editor().press('ArrowDown'); await expect(page.locator('.ant-select-dropdown:visible')).toBeVisible()
  await editor().evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' })))
  await editor().fill('Beta')
  assert.equal(await key({ key: 'Enter', keyCode: 13, which: 13, isComposing: true }), false)
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'Alpha')
  assert.deepEqual(await history(), { undo: 0, redo: 0 }); await expect(editor()).toHaveValue('Beta')
  await expect(page.getByRole('button', { name: '确认编辑', exact: true })).toBeDisabled()
  await editor().evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'Beta' })))
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'Beta'); assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await edit(); await editor().press('ArrowDown')
  await page.locator('.ant-select-dropdown:visible').getByTitle('Alpha', { exact: true }).click()
  await expect(editor()).toHaveValue('Alpha'); await editor().press('Tab')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'Alpha')
})

test('Alt+Enter bypasses open validation choices and a rejected multiline commit keeps its draft and selection', async () => {
  await page.evaluate(() => { window.sheetLayoutFixture.seed([{ value: 'Alpha', style: { wrap: true, wrapStrategy: 'wrap' } }]); window.sheetLayoutFixture.list() })
  await edit(); await editor().press('ArrowDown'); await editor().press('Alt+Enter')
  await expect(editor()).toHaveValue('Alpha\n'); assert.deepEqual(await caretRange(), [6, 6])
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect(editor()).toHaveValue('Alpha\n')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'Alpha'); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.ok((await page.evaluate(() => window.sheetLayoutFixture.notifications())).some(item => item.message.includes('Alpha')))
  await page.getByRole('button', { name: '取消编辑', exact: true }).click(); await expect(grid()).toBeFocused()
})

for (const variant of ['classic', 'grid']) test(variant + ' wrap menu and API share the Canvas layout, disabled fit state and undo/reopen model', async () => {
  const text = 'alpha beta gamma delta'
  await page.evaluate(({ text, variant }) => {
    window.sheetLayoutFixture.seed([{ value: text, style: { font: { size: 18 }, shrinkToFit: true } }, { column: 3, value: 'blocker' }], { columnWidths: { 0: 65 }, rowHeights: { 0: 130 } })
    window.sheetLayoutFixture.appearance(false, variant)
  }, { text, variant })
  await select(); await chooseMenu('sheet.menu.text-wrap', '自动换行')
  const wrapped = await snapshot(); const style = wrapped.sheets[0].cells['0:0'].style
  assert.equal(style.wrapStrategy, 'wrap'); assert.equal(style.wrap, true); assert.equal(style.shrinkToFit, false)
  await expect(page.locator('[data-ribbon-item="sheet.command.set-shrink-to-fit"]:visible').first().getByRole('button')).toBeDisabled()
  await expect.poll(async () => (await draws()).filter(item => item.text && text.includes(item.text) && item.y >= 26).map(item => item.text).join('')).toBe(text)
  const lines = (await scene()).commands.filter(item => item.kind === 'text' && item.text && text.includes(item.text) && item.clip?.y >= 26)
  assert.ok(lines.length > 1)
  await command('sheet.command.undo'); assert.equal((await snapshot()).sheets[0].cells['0:0'].style.wrapStrategy, 'overflow')
  await page.evaluate(value => window.sheetLayoutFixture.reopen(JSON.stringify(value)), wrapped)
  await edit(); await expect(page.locator('.canvas-cell-input textarea')).toBeFocused()
  await editor().press('Escape'); await expect(canvas()).toHaveAttribute('data-layout-test-mount', 'original')
})

test('API rotation and shrink submit a uniformly smaller font to the real Canvas2D surface', async () => {
  const text = 'longrotatedunbrokenvalue'
  await page.evaluate(text => {
    window.sheetLayoutFixture.seed([{ value: text, style: { font: { size: 18 } } }], { columnWidths: { 0: 80 }, rowHeights: { 0: 48 } })
    window.sheetLayoutFixture.layout('setTextRotation', 45)
    window.sheetLayoutFixture.layout('setShrinkToFit', true)
  }, text)
  await expect.poll(async () => (await draws()).some(item => item.text === text && Math.abs(item.matrix.b) > 0.1)).toBe(true)
  const draw = (await draws()).find(item => item.text === text)
  assert.ok(Number(/([\d.]+)px/u.exec(draw.font)[1]) < 18 * 4 / 3)
  assert.equal(draw.maxWidth, undefined)
  const saved = await snapshot(); assert.equal(saved.sheets[0].cells['0:0'].style.rotation, 45); assert.equal(saved.sheets[0].cells['0:0'].style.shrinkToFit, true)
  await page.evaluate(value => window.sheetLayoutFixture.reopen(JSON.stringify(value)), saved)
  await expect.poll(async () => (await draws()).some(item => item.text === text && Math.abs(item.matrix.b) > 0.1)).toBe(true)
})

test('read-only multiline input cannot alter the model or history and rejected formatting stays disabled', async () => {
  await page.evaluate(() => { window.sheetLayoutFixture.seed([{ value: 'locked\nvalue', style: { wrap: true, wrapStrategy: 'wrap' } }]); window.sheetLayoutFixture.denyEdit() })
  await edit(); await editor().fill('attempt\nchange'); await editor().press('Control+Enter')
  await expect(editor()).toHaveValue('attempt\nchange'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'locked\nvalue')
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await editor().press('Escape')
  await expect(page.locator('[data-ribbon-item="sheet.menu.text-wrap"]:visible').first().getByRole('button')).toBeDisabled()
})
