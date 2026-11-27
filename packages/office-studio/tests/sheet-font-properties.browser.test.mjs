import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-font-properties-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Sheets font family and size</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model'
import { toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankWorkbook()
initial.id = 'sheet-font-properties-fixture'; initial.name = '字体字号'; initial.calculationMode = 'manual'
const sheet = initial.sheets[0]
sheet.id = 'sheet-font-main'; initial.activeSheetId = sheet.id; sheet.rowCount = 200; sheet.columnCount = 26
sheet.cells['0:0'] = createCell(0, 0, 12.25)
sheet.cells['1:0'] = { ...createCell(1, 0, .25), formula: '=1/4', dataType: 'formula', cachedDataType: 'number' }
sheet.cells['2:0'] = { ...createCell(2, 0, '00123'), dataType: 'text' }
for (const [row, size] of [[0, 9], [1, 19], [2, 30]]) { sheet.cells[row + ':0'].style.font.family = 'Courier New'; sheet.cells[row + ':0'].style.font.size = size }
sheet.cells['4:4'] = createCell(4, 4, 'outside')
const studio = new OfficeStudio({ workbook: initial, currentUserId: 'local-user' }), api = FOffice.newAPI(studio)
const view = ref(toExcelWorkbookView(studio.editor.workbook.model)), dark = ref(false), variant = ref('classic'), canceled = ref(false), commands = [], notifications = []
const direction = computed(() => ['ar-SA', 'fa-IR'].includes(studio.regionalSettings.locale.value) ? 'rtl' : 'ltr')
watch(() => studio.editor.workbook.model, model => { view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
studio.events.addEvent('CommandExecuted', event => commands.push(copy(event)))
studio.events.addEvent('BeforeCommandExecute', event => { if (canceled.value && ['sheet.command.set-font-family', 'sheet.command.set-font-size', 'sheet.command.increase-font-size', 'sheet.command.decrease-font-size'].includes(event.id)) event.cancel = true })
window.sheetFontFixture = {
  snapshot: () => copy(studio.editor.workbook.model),
  history: (id = studio.activeWorkbookId) => { const unit = studio.editor.history.units[id]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0 } },
  commands: () => copy(commands), notifications: () => copy(notifications),
  cancel: value => { canceled.value = value },
  appearance: (isDark, language = 'zh-CN') => { dark.value = isDark; studio.regionalSettings.setLocale(language) },
  variant: value => { variant.value = value },
  apiFont: (field, value, address = 'A1') => { const target = api.getActiveWorkbook().getActiveSheet().getRange(address); if (field === 'family') target.setFontFamily(value); else target.setFontSize(value) },
  newWorkbook: () => { api.createWorkbook({ ...copy(initial), id: 'sheet-font-properties-second' }) },
  activate: id => { studio.activateWorkbook(id) },
  reopen: snapshot => { api.createWorkbook({ ...JSON.parse(snapshot), id: 'sheet-font-properties-reopened' }) },
  denyEdit: () => {
    const next = copy(studio.editor.workbook.model)
    next.sheets[0].rangeProtectionRules = [{ id: 'locked-font', permissionId: 'locked-font-permission', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [{ startRow: 0, endRow: 2, startCol: 0, endCol: 0 }] }]
    studio.editor.workbook.replace(next)
  },
  filteredRows: () => {
    const next = copy(studio.editor.workbook.model)
    next.sheets[0].autoFilter = { range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 }, columns: { 0: { kind: 'values', selectedItems: [7] } }, filteredRows: [1] }
    next.sheets[0].hiddenRows = [2]; studio.editor.workbook.replace(next)
  },
}
createApp({ render: () => h(ConfigProvider, {
  direction: direction.value, theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } },
}, { default: () => h(ExcelWorkbench, {
  modelValue: view.value, studio, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
  'onUpdate:modelValue': value => studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value),
  onUndo: () => api.getActiveWorkbook().undo(), onRedo: () => api.getActiveWorkbook().redo(), onNotify: value => notifications.push(copy(value)),
}) }) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-font-properties-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message)); await page.goto(baseURL)
  await expect(grid()).toBeVisible(); await expect(canvas()).toHaveCount(1)
  await canvas().evaluate(element => { element.dataset.fontTestMount = 'original' })
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const canvas = () => page.locator('.canvas-surface canvas')
const editor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const snapshot = () => page.evaluate(() => window.sheetFontFixture.snapshot())
const history = () => page.evaluate(() => window.sheetFontFixture.history())
const commands = () => page.evaluate(() => window.sheetFontFixture.commands())
const fonts = async () => { const model = await snapshot(); return [0, 1, 2].map(row => { const font = model.sheets[0].cells[row + ':0'].style.font; return [font.family, font.size] }) }
const content = model => Object.fromEntries(Object.entries(model.sheets[0].cells).map(([key, { value, formula, display, dataType, cachedDataType }]) => [key, { value, formula, display, dataType, cachedDataType }]))
async function homeItem(id) {
  const item = page.locator('.ribbon-visible-groups [data-ribbon-item="' + id + '"]')
  if (await item.count() && await item.isVisible()) return item
  const overflow = page.locator('.ribbon-overflow-panel:visible')
  if (!await overflow.isVisible()) await page.locator('.ribbon-more-button').click()
  const result = overflow.locator('[data-ribbon-item="' + id + '"]'); await expect(result).toBeVisible(); return result
}
const fontInput = async (field = 'family') => (await homeItem('sheet.command.set-font-' + field)).getByRole('combobox')
async function selectRange(address) {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter'); await expect(input).toHaveValue(address); await expect(grid()).toBeFocused()
}
async function submitFont(field, value) {
  const input = await fontInput(field); await input.fill(value); await input.press('Enter')
}

test('font family and fractional size edits share commands, history and snapshot reopen while retaining the Canvas', async () => {
  await selectRange('A1:A3')
  const before = content(await snapshot())
  await submitFont('family', 'Arial'); await submitFont('size', '14.5')
  await expect.poll(fonts).toEqual([['Arial', 14.5], ['Arial', 14.5], ['Arial', 14.5]])
  assert.deepEqual((await commands()).map(event => event.id), ['sheet.command.set-font-family', 'sheet.command.set-font-size'])
  assert.deepEqual(await history(), { undo: 2, redo: 0 }); assert.deepEqual(content(await snapshot()), before)
  await expect(grid()).toBeFocused(); await expect(canvas()).toHaveAttribute('data-font-test-mount', 'original')
  await (await homeItem('sheet.command.undo')).getByRole('button').click()
  await expect.poll(fonts).toEqual([['Arial', 9], ['Arial', 19], ['Arial', 30]])
  await (await homeItem('sheet.command.redo')).getByRole('button').click()
  const saved = JSON.stringify(await snapshot()); await page.evaluate(value => window.sheetFontFixture.reopen(value), saved)
  await expect.poll(fonts).toEqual([['Arial', 14.5], ['Arial', 14.5], ['Arial', 14.5]])
  assert.deepEqual(content(await snapshot()), before)
})

test('Escape, blur and incomplete selector drafts leave model and history unchanged', async () => {
  await selectRange('A1')
  const before = await snapshot(), family = await fontInput('family')
  await family.fill('Arial'); await family.press('Escape'); await expect(family).toHaveValue('Courier New')
  await family.fill('Arial'); await grid().focus()
  await expect(family).toHaveValue('Courier New')
  await selectRange('A1')
  await submitFont('family', 'Unlisted Test Font'); await expect(await fontInput('family')).toHaveValue('Courier New')
  await submitFont('size', '1e2'); await expect(await fontInput('size')).toHaveValue('9')
  const size = await fontInput('size'); await size.fill('-'); await size.press('Escape'); await expect(size).toHaveValue('9')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0 }); assert.deepEqual(await commands(), [])
})

test('mixed-selection steppers use the active cell instead of a mixed selector fallback', async () => {
  await selectRange('A1:A3')
  await expect(await fontInput('size')).toHaveValue('')
  await (await homeItem('sheet.command.increase-font-size')).getByRole('button').click()
  await expect.poll(fonts).toEqual([['Courier New', 10], ['Courier New', 10], ['Courier New', 10]])
  await (await homeItem('sheet.command.decrease-font-size')).getByRole('button').click()
  await expect.poll(fonts).toEqual([['Courier New', 9], ['Courier New', 9], ['Courier New', 9]])
  assert.deepEqual(await history(), { undo: 2, redo: 0 })
  assert.deepEqual((await commands()).map(event => event.id), ['sheet.command.increase-font-size', 'sheet.command.decrease-font-size'])
})

test('font-size keyboard shortcuts use the active cell and share command history and snapshot reopen', async () => {
  await selectRange('A1:A3')
  const before = content(await snapshot())
  await grid().press('Control+Shift+.')
  await expect.poll(fonts).toEqual([['Courier New', 10], ['Courier New', 10], ['Courier New', 10]])
  await grid().press('Control+Shift+,')
  await expect.poll(fonts).toEqual([['Courier New', 9], ['Courier New', 9], ['Courier New', 9]])
  assert.deepEqual((await commands()).map(event => event.id), ['sheet.command.increase-font-size', 'sheet.command.decrease-font-size'])
  assert.deepEqual(await history(), { undo: 2, redo: 0 }); assert.deepEqual(content(await snapshot()), before)
  await expect(grid()).toBeFocused(); await expect(canvas()).toHaveAttribute('data-font-test-mount', 'original')
  await grid().press('Control+z')
  await expect.poll(fonts).toEqual([['Courier New', 10], ['Courier New', 10], ['Courier New', 10]])
  await grid().press('Control+y')
  await expect.poll(fonts).toEqual([['Courier New', 9], ['Courier New', 9], ['Courier New', 9]])
  const saved = JSON.stringify(await snapshot())
  await page.evaluate(value => window.sheetFontFixture.reopen(value), saved)
  await expect.poll(fonts).toEqual([['Courier New', 9], ['Courier New', 9], ['Courier New', 9]])
  assert.deepEqual(content(await snapshot()), before)
})

test('font-size shortcuts preserve native text selection and publish one rich-text transaction on Enter', async () => {
  await selectRange('A3'); await grid().press('F2'); await expect(editor()).toBeFocused()
  const before = await snapshot(), input = editor()
  await input.evaluate(element => {
    const prefix = element.value.length - '00123'.length
    element.setSelectionRange(prefix + 1, prefix + 3)
    element.dispatchEvent(new Event('select', { bubbles: true }))
  })
  const selection = await input.evaluate(element => [element.selectionStart, element.selectionEnd])
  await input.press('Control+Shift+.'); await expect(await fontInput('size')).toHaveValue('31')
  await input.press('Control+Shift+,'); await expect(await fontInput('size')).toHaveValue('30')
  await input.press('Control+Shift+.'); await expect(await fontInput('size')).toHaveValue('31')
  await expect(input).toBeFocused()
  assert.deepEqual(await input.evaluate(element => [element.selectionStart, element.selectionEnd]), selection)
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.deepEqual((await commands()).map(event => event.id), Array(3).fill('sheet.operation.format-cell-input'))
  await input.press('Enter'); await expect(editor()).toHaveCount(0); await expect(grid()).toBeFocused()
  const saved = await snapshot(), cell = saved.sheets[0].cells['2:0']
  assert.equal(cell.value, '00123'); assert.equal(cell.dataType, 'text'); assert.equal(cell.style.font.size, 30)
  assert.equal(cell.richText.filter(run => run.font?.size === 31).map(run => run.text).join(''), '01')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await grid().press('Control+z'); assert.deepEqual((await snapshot()).sheets[0].cells['2:0'], before.sheets[0].cells['2:0'])
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).sheets[0].cells['2:0'], cell)
  await page.evaluate(value => window.sheetFontFixture.reopen(value), JSON.stringify(saved))
  assert.deepEqual((await snapshot()).sheets[0].cells['2:0'], cell)
})

test('toolbar input bounds and imported API font-size step behavior remain distinct', async () => {
  await selectRange('A1')
  await submitFont('size', '500')
  await expect.poll(async () => (await fonts())[0][1]).toBe(400)
  await page.evaluate(() => window.sheetFontFixture.apiFont('size', 500)); await expect(await fontInput('size')).toHaveValue('500')
  await (await homeItem('sheet.command.decrease-font-size')).getByRole('button').click()
  await expect.poll(async () => (await fonts())[0][1]).toBe(499)
  await (await homeItem('sheet.command.increase-font-size')).getByRole('button').click()
  await expect.poll(async () => (await fonts())[0][1]).toBe(400)
  await submitFont('size', '-12'); await expect.poll(async () => (await fonts())[0][1]).toBe(6)
  const before = await history(); await (await homeItem('sheet.command.decrease-font-size')).getByRole('button').click()
  await expect.poll(async () => (await fonts())[0][1]).toBe(6); assert.deepEqual(await history(), before)
  assert.equal((await commands()).at(-1).id, 'sheet.command.decrease-font-size')
})

test('external API font edits synchronize selectors and discard uncommitted toolbar drafts', async () => {
  await selectRange('A1')
  const family = await fontInput('family'); await family.fill('Arial')
  assert.equal((await fonts())[0][0], 'Courier New')
  await page.evaluate(() => window.sheetFontFixture.apiFont('family', 'Times New Roman'))
  await expect(family).toHaveValue('Times New Roman')
  const size = await fontInput('size'); await size.fill('24')
  await page.evaluate(() => window.sheetFontFixture.apiFont('size', 42))
  await expect(size).toHaveValue('42'); await size.press('Escape')
  assert.deepEqual(await history(), { undo: 2, redo: 0 }); assert.deepEqual((await fonts())[0], ['Times New Roman', 42])
  await expect(canvas()).toHaveAttribute('data-font-test-mount', 'original')
})

test('API null resets retain sparse cells, content and default selector values over a full worksheet', async () => {
  await selectRange('A1')
  const before = await snapshot(), count = Object.keys(before.sheets[0].cells).length
  await page.evaluate(() => { window.sheetFontFixture.apiFont('family', null, 'A1:XFD1048576'); window.sheetFontFixture.apiFont('size', null, 'A1:XFD1048576') })
  await expect.poll(fonts).toEqual([['Aptos', 11], ['Aptos', 11], ['Aptos', 11]])
  await expect(await fontInput('family')).toHaveValue('Aptos'); await expect(await fontInput('size')).toHaveValue('11')
  const after = await snapshot(); assert.equal(Object.keys(after.sheets[0].cells).length, count); assert.deepEqual(content(after), content(before))
  assert.deepEqual(await history(), { undo: 2, redo: 0 })
})

test('canceled property commands preserve model and controls become disabled when edit permission changes', async () => {
  await selectRange('A1')
  const before = await snapshot(); await page.evaluate(() => window.sheetFontFixture.cancel(true))
  await submitFont('family', 'Arial'); await submitFont('size', '24')
  await expect.poll(() => page.evaluate(() => window.sheetFontFixture.notifications())).toHaveLength(2)
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await commands(), []); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await page.evaluate(() => window.sheetFontFixture.denyEdit())
  for (const field of ['family', 'size']) await expect(await fontInput(field)).toBeDisabled()
  for (const id of ['sheet.command.increase-font-size', 'sheet.command.decrease-font-size']) await expect((await homeItem(id)).getByRole('button')).toBeDisabled()
  const button = (await homeItem('sheet.command.increase-font-size')).getByRole('button')
  await button.evaluate(element => element.dispatchEvent(new MouseEvent('click', { bubbles: true })))
  assert.deepEqual(await fonts(), [['Courier New', 9], ['Courier New', 19], ['Courier New', 30]])
  assert.deepEqual(await commands(), []); assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('workbook switches retain independent font values and histories', async () => {
  await selectRange('A1'); await submitFont('family', 'Arial')
  await page.evaluate(() => window.sheetFontFixture.newWorkbook()); await selectRange('A1')
  await expect(await fontInput('family')).toHaveValue('Courier New'); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await submitFont('size', '24'); await page.evaluate(() => window.sheetFontFixture.activate('sheet-font-properties-fixture')); await selectRange('A1')
  await expect(await fontInput('family')).toHaveValue('Arial'); await expect(await fontInput('size')).toHaveValue('9')
  await (await homeItem('sheet.command.undo')).getByRole('button').click(); await expect(await fontInput('family')).toHaveValue('Courier New')
  await page.evaluate(() => window.sheetFontFixture.activate('sheet-font-properties-second')); await selectRange('A1')
  await expect(await fontInput('size')).toHaveValue('24'); assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('font selectors edit manually hidden selected rows while skipping filter-hidden rows', async () => {
  await page.evaluate(() => window.sheetFontFixture.filteredRows()); await selectRange('A1:A3')
  await submitFont('family', 'Arial'); await submitFont('size', '14')
  await expect.poll(fonts).toEqual([['Arial', 14], ['Courier New', 19], ['Arial', 14]])
  assert.deepEqual((await snapshot()).sheets[0].hiddenRows, [2]); assert.deepEqual(await history(), { undo: 2, redo: 0 })
})

test('external family and size commands preserve an active cell draft and its committed font', async () => {
  await selectRange('A1'); await grid().press('F2'); await expect(editor()).toBeVisible(); await editor().fill('42')
  await page.evaluate(() => { window.sheetFontFixture.apiFont('family', 'Arial'); window.sheetFontFixture.apiFont('size', 24) })
  await expect(editor()).toHaveValue('42'); await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('42')
  await editor().press('Enter'); await expect.poll(async () => (await snapshot()).sheets[0].cells['0:0'].value).toBe(42)
  assert.deepEqual((await fonts())[0], ['Arial', 24]); assert.deepEqual(await history(), { undo: 3, redo: 0 })
})

test('narrow grid overflow keeps editable font controls on the shared command path', async () => {
  await selectRange('A1'); await page.evaluate(() => window.sheetFontFixture.variant('grid')); await page.setViewportSize({ width: 260, height: 900 })
  await submitFont('family', 'Arial'); await submitFont('size', '14')
  await expect.poll(async () => (await fonts())[0]).toEqual(['Arial', 14]); assert.deepEqual(await history(), { undo: 2, redo: 0 })
  assert.deepEqual((await commands()).map(event => event.id), ['sheet.command.set-font-family', 'sheet.command.set-font-size']); await expect(grid()).toBeFocused()
})

test('font input ignores composing Enter and Escape until a normal commit key arrives', async () => {
  await selectRange('A1')
  const input = await fontInput('size'); await input.fill('24')
  for (const options of [{ key: 'Enter', isComposing: true }, { key: 'Escape', isComposing: true }, { key: 'Enter', keyCode: 229 }, { key: 'Escape', keyCode: 229 }]) {
    assert.equal(await input.evaluate((element, info) => { const event = new KeyboardEvent('keydown', { ...info, bubbles: true, cancelable: true }); element.dispatchEvent(event); return event.defaultPrevented }, options), false)
    assert.equal((await fonts())[0][1], 9); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  }
  await input.press('Enter'); await expect.poll(async () => (await fonts())[0][1]).toBe(24); assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('controlled composition drafts and font suggestion clicks each submit one shared command', async () => {
  await selectRange('A1')
  const before = content(await snapshot()), input = await fontInput('size')
  await input.fill('')
  await input.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  for (const value of ['2', '24']) {
    await input.evaluate((element, text) => {
      element.value = text
      element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: text }))
      element.dispatchEvent(new InputEvent('input', { bubbles: true, data: text, inputType: 'insertCompositionText', isComposing: true }))
    }, value)
    await expect(input).toHaveValue(value)
    assert.deepEqual((await fonts())[0], ['Courier New', 9])
    assert.deepEqual(await commands(), []); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  }
  await input.evaluate(element => element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', isComposing: true })))
  assert.deepEqual(await commands(), [])
  await input.evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '24' }))
    element.dispatchEvent(new InputEvent('input', { bubbles: true, data: '24', inputType: 'insertText', isComposing: false }))
  })
  await expect(input).toHaveValue('24')
  await input.press('Enter')
  await expect.poll(async () => (await fonts())[0]).toEqual(['Courier New', 24])
  assert.deepEqual((await commands()).map(event => event.id), ['sheet.command.set-font-size'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  const family = await fontInput('family')
  await family.focus(); await family.press('ArrowDown')
  await page.locator('.ant-select-dropdown:visible').getByRole('option', { name: 'Arial', exact: true }).click()
  await expect.poll(async () => (await fonts())[0]).toEqual(['Arial', 24])
  await expect(family).toHaveValue('Arial'); await expect(grid()).toBeFocused()
  assert.deepEqual((await commands()).map(event => event.id), ['sheet.command.set-font-size', 'sheet.command.set-font-family'])
  assert.deepEqual(await history(), { undo: 2, redo: 0 }); assert.deepEqual(content(await snapshot()), before)
})

test('portaled font suggestions follow ConfigProvider dark theme and RTL without changing content', async () => {
  await selectRange('A1')
  const input = await fontInput('family'); await input.focus(); await input.press('ArrowDown')
  const menu = page.locator('.ant-select-dropdown:visible'); await expect(menu).toBeVisible()
  const background = await menu.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.sheetFontFixture.appearance(true, 'ar-SA'))
  await expect(menu).toHaveClass(/ant-select-dropdown-rtl/)
  await expect.poll(() => menu.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(background)
  await input.press('Escape'); await expect(menu).toBeHidden()
  assert.deepEqual((await fonts())[0], ['Courier New', 9]); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await expect(canvas()).toHaveAttribute('data-font-test-mount', 'original')
})
