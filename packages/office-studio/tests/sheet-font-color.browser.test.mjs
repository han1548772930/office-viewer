import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-font-color-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Sheets font and color commands</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model'
import { toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
import { initializeWorkbookCellStyles } from '/packages/office-studio/src/features/sheet-workbench/sheet-style-inheritance'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankWorkbook()
initial.id = 'sheet-font-color-fixture'; initial.name = '字体与颜色'; initial.calculationMode = 'manual'
const sheet = initial.sheets[0]
sheet.rowCount = 200; sheet.columnCount = 26
sheet.cells['0:0'] = createCell(0, 0, 12.25)
sheet.cells['0:0'].style.font.color = '#414657'; sheet.cells['0:0'].style.fill = '#ffffff'
sheet.cells['0:1'] = { ...createCell(0, 1, .25), formula: '=1/4', dataType: 'formula', cachedDataType: 'number' }
sheet.cells['1:0'] = { ...createCell(1, 0, '00123'), dataType: 'text' }
sheet.cells['2:0'] = createCell(2, 0, 7)
sheet.cells['2:0'].style.font.bold = true; sheet.cells['2:0'].style.font.color = '#f05252'
sheet.cells['4:4'] = createCell(4, 4, 'outside')
const studio = new OfficeStudio({ workbook: initial, currentUserId: 'local-user' })
const api = FOffice.newAPI(studio), view = ref(toExcelWorkbookView(studio.editor.workbook.model))
const dark = ref(false), variant = ref('classic'), canceled = ref(false), commands = [], notifications = []
const direction = computed(() => ['ar-SA', 'fa-IR'].includes(studio.regionalSettings.locale.value) ? 'rtl' : 'ltr')
watch(() => studio.editor.workbook.model, model => { view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
studio.events.addEvent('CommandExecuted', event => { commands.push(copy(event)) })
studio.events.addEvent('BeforeCommandExecute', event => { if (canceled.value && /set-(?:color|font-color|background)$/.test(event.id)) event.cancel = true })
window.sheetColorFixture = {
  snapshot: () => copy(studio.editor.workbook.model),
  history: () => { const unit = studio.editor.history.units[studio.activeWorkbookId]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0 } },
  commands: () => copy(commands), notifications: () => copy(notifications),
  cancelColors: value => { canceled.value = value },
  appearance: (isDark, language = 'zh-CN') => { dark.value = isDark; studio.regionalSettings.setLocale(language) },
  variant: value => { variant.value = value },
  apiColor: (field, value, address = 'A1') => { const range = api.getActiveWorkbook().getActiveSheet().getRange(address); if (field === 'color') range.setFontColor(value); else range.setBackground(value) },
  inheritColors: () => {
    const next = copy(studio.editor.workbook.model)
    next.defaultStyle = { font: { color: '#3f83f8' }, fill: '#ebf5ff' }
    next.styles = { colored: { font: { color: '#112233' }, fill: '#ffeecc' } }
    next.sheets[0].cells['1:0'].authoredStyleReference = 'colored'
    next.sheets[0].cells['1:0'].richText = [{ id: 'source', text: '00123', color: '#f05252',
      font: { color: '#f05252', italic: true }, hyperlink: 'https://example.test', vendor: { keep: true } }]
    initializeWorkbookCellStyles(next)
    studio.editor.workbook.replace(next)
  },
  externalColor: value => { const next = copy(studio.editor.workbook.model); next.sheets[0].cells['0:0'].style.font.color = value; studio.editor.workbook.replace(next) },
  newWorkbook: () => { api.createWorkbook({ ...copy(initial), id: 'sheet-font-color-second' }) },
  activate: id => { studio.activateWorkbook(id) },
  reopen: snapshot => { api.createWorkbook({ ...JSON.parse(snapshot), id: 'sheet-font-color-reopened' }) },
  denyEdit: () => {
    const next = copy(studio.editor.workbook.model)
    next.sheets[0].rangeProtectionRules = [{ id: 'locked-color', permissionId: 'locked-color-permission', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [{ startRow: 0, endRow: 2, startCol: 0, endCol: 1 }] }]
    studio.editor.workbook.replace(next)
  },
  filteredRows: () => {
    const next = copy(studio.editor.workbook.model)
    next.sheets[0].autoFilter = { range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 }, columns: { 0: { kind: 'values', selectedItems: [7] } }, filteredRows: [1] }
    next.sheets[0].hiddenRows = [2]
    studio.editor.workbook.replace(next)
  },
}
createApp({ render: () => h(ConfigProvider, {
  direction: direction.value, theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } },
}, { default: () => h(ExcelWorkbench, {
  modelValue: view.value, studio, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
  'onUpdate:modelValue': value => { studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value) },
  onUndo: () => { api.getActiveWorkbook().undo() }, onRedo: () => { api.getActiveWorkbook().redo() },
  onNotify: value => { notifications.push(copy(value)) },
}) }) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-font-color-fixture', configureServer(vite) {
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
  errors = []
  page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(grid()).toBeVisible()
  await expect(canvas()).toHaveCount(1)
  await canvas().evaluate(element => { element.dataset.colorTestMount = 'original' })
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const snapshot = () => page.evaluate(() => window.sheetColorFixture.snapshot())
const history = () => page.evaluate(() => window.sheetColorFixture.history())
const commands = () => page.evaluate(() => window.sheetColorFixture.commands())
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const canvas = () => page.locator('.canvas-surface canvas')
const cellEditor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const ribbonItem = id => page.locator('.ribbon-visible-groups [data-ribbon-item="' + id + '"]')
const colors = (field, rows = [0]) => page.evaluate(({ field, rows }) => rows.map(row => { const cell = window.sheetColorFixture.snapshot().sheets[0].cells[row + ':0']; return field === 'background' ? cell?.style.fill : cell?.style.font.color }), { field, rows })
const content = model => Object.entries(model.sheets[0].cells).map(([key, { value, formula, display, dataType, cachedDataType }]) => [key, { value, formula, display, dataType, cachedDataType }])
async function homeItem(id) {
  const visible = ribbonItem(id)
  if (await visible.count() && await visible.isVisible()) return visible
  const overflow = page.locator('.ribbon-overflow-panel:visible')
  if (!await overflow.isVisible()) await page.locator('.ribbon-more-button').click()
  const result = overflow.locator('[data-ribbon-item="' + id + '"]')
  await expect(result).toBeVisible()
  return result
}
async function selectRange(address) {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter')
  await expect(input).toHaveValue(address)
  await expect(grid()).toBeFocused()
}
async function openColor(field = 'color', keyboard = false) {
  const item = await homeItem(field === 'background' ? 'sheet.menu.fill' : 'sheet.menu.text-color')
  const arrow = item.locator('.sheet-color-arrow button')
  if (keyboard) { await arrow.focus(); await arrow.press('ArrowDown') } else await arrow.click()
  const panel = page.locator('.sheet-color-panel:visible')
  await expect(panel).toBeVisible()
  return panel
}
async function openCustom(panel) {
  await panel.locator('.sheet-color-more button').click()
  const input = panel.getByRole('textbox', { name: 'HEX', exact: true })
  await expect(input).toBeFocused()
  return input
}
async function chooseColorFormat(panel, label) {
  await panel.getByRole('combobox', { name: '颜色格式', exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}
async function mainSwatch(field = 'color') {
  const item = await homeItem(field === 'background' ? 'sheet.menu.fill' : 'sheet.menu.text-color')
  return item.locator('.sheet-color-bar').evaluate(element => getComputedStyle(element).backgroundColor)
}

test('45 presets format Canvas cells through one command and survive undo, redo and snapshot reopen', async () => {
  await selectRange('A1:A3')
  const before = await snapshot(), panel = await openColor('color', true)
  await expect(panel.getByRole('button', { name: '#FFFFFF', exact: true })).toBeFocused()
  assert.deepEqual(await panel.locator('.sheet-color-preset-row').evaluateAll(rows => rows.map(row => row.querySelectorAll('button').length)), [9, 9, 9, 9, 9])
  await expect(panel.locator('.sheet-color-state')).toHaveText('混合')
  await panel.getByRole('button', { name: '#3F83F8', exact: true }).click()
  await expect.poll(() => colors('color', [0, 1, 2])).toEqual(['#3f83f8', '#3f83f8', '#3f83f8'])
  assert.deepEqual(content(await snapshot()), content(before))
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  assert.equal((await commands()).at(-1).id, 'sheet.command.set-color')
  await expect(grid()).toBeFocused()
  await expect(canvas()).toHaveAttribute('data-color-test-mount', 'original')
  await (await homeItem('sheet.command.undo')).getByRole('button').click()
  assert.deepEqual(content(await snapshot()), content(before))
  assert.deepEqual(await colors('color', [0, 1, 2]), ['#414657', null, '#f05252'])
  await (await homeItem('sheet.command.redo')).getByRole('button').click()
  const saved = JSON.stringify(await snapshot())
  await page.evaluate(value => window.sheetColorFixture.reopen(value), saved)
  await expect.poll(() => colors('color', [0, 1, 2])).toEqual(['#3f83f8', '#3f83f8', '#3f83f8'])
  assert.deepEqual(content(await snapshot()), content(before))
})

test('mixed font marks, toolbar buttons and desktop shortcuts share the registered command path', async () => {
  await selectRange('A1:A3')
  const before = content(await snapshot()), bold = (await homeItem('sheet.command.toggle-bold')).getByRole('button')
  await expect(bold).toHaveAttribute('aria-pressed', 'false')
  await bold.click()
  await expect.poll(async () => (await snapshot()).sheets[0].cells['0:0'].style.font.bold).toBe(true)
  await expect(bold).toHaveAttribute('aria-pressed', 'true')
  await grid().press('Control+B')
  await expect.poll(async () => {
    const model = await snapshot()
    return [0, 1, 2].map(row => model.sheets[0].cells[row + ':0'].style.font.bold)
  }).toEqual([false, false, false])
  for (const [key, field, id] of [['Control+I', 'italic', 'toggle-italic'], ['Control+U', 'underline', 'toggle-underline'], ['Control+Shift+X', 'strike', 'toggle-strike']]) {
    await grid().press(key)
    await expect.poll(async () => (await snapshot()).sheets[0].cells['0:0'].style.font[field]).toBe(true)
    assert.equal((await commands()).at(-1).id, 'sheet.command.' + id)
  }
  assert.deepEqual(content(await snapshot()), before)
  assert.deepEqual(await history(), { undo: 5, redo: 0 })
})

test('custom color edits remain draft until confirmation and cancel restores the initial value and focus', async () => {
  await selectRange('A1')
  const before = await snapshot(), panel = await openColor()
  let input = await openCustom(panel)
  const initial = await input.inputValue()
  await input.fill('abcdef')
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await panel.getByRole('button', { name: /取\s*消/ }).click()
  await expect(panel.locator('.sheet-color-more button')).toBeFocused()
  input = await openCustom(panel)
  await expect(input).toHaveValue(initial)
  await input.fill('abcdef')
  await panel.getByRole('button', { name: /确\s*定/ }).click()
  await expect.poll(() => colors('color')).toEqual(['#abcdef'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(grid()).toBeFocused()
})

test('incomplete HEX Enter leaves the custom panel open and never submits the previous valid value', async () => {
  await selectRange('A1')
  const panel = await openColor(), input = await openCustom(panel)
  await input.fill('12'); await input.press('Enter')
  await expect(panel.locator('.sheet-color-custom')).toBeVisible()
  await expect(panel).toContainText('请输入有效的颜色值')
  assert.deepEqual(await colors('color'), ['#414657'])
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await input.fill('0a1b2c'); await input.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#0a1b2c'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('Escape discards custom draft, then closes the palette and returns focus to the arrow', async () => {
  await selectRange('A1')
  const before = await snapshot(), panel = await openColor(), input = await openCustom(panel)
  await input.fill('abcdef'); await input.press('Escape')
  await expect(panel.locator('.sheet-color-more button')).toBeFocused()
  await expect(panel.locator('.sheet-color-presets')).toBeVisible()
  await panel.locator('.sheet-color-more button').press('Escape')
  await expect(page.locator('.sheet-color-panel')).toBeHidden()
  await expect(ribbonItem('sheet.menu.text-color').locator('.sheet-color-arrow button')).toBeFocused()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('IME Enter and Escape retain native defaults and composition disables confirmation until it ends', async () => {
  await selectRange('A1')
  const panel = await openColor(), input = await openCustom(panel)
  await input.fill('112233')
  for (const options of [{ key: 'Enter', isComposing: true }, { key: 'Escape', isComposing: true }, { key: 'Enter', keyCode: 229 }, { key: 'Escape', keyCode: 229 }]) {
    assert.equal(await input.evaluate((element, info) => {
      const event = new KeyboardEvent('keydown', { ...info, bubbles: true, cancelable: true })
      element.dispatchEvent(event); return event.defaultPrevented
    }, options), false)
    await expect(panel.locator('.sheet-color-custom')).toBeVisible()
    assert.deepEqual(await history(), { undo: 0, redo: 0 })
  }
  await input.dispatchEvent('compositionstart')
  await expect(panel.getByRole('button', { name: /确\s*定/ })).toBeDisabled()
  await input.press('Enter')
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await input.dispatchEvent('compositionend')
  await input.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#112233'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('slider keyboard adjustments and RGB channels remain draft and commit their actual value', async () => {
  await selectRange('A1')
  let panel = await openColor(), input = await openCustom(panel)
  await input.fill('336699')
  const slider = panel.getByRole('slider').first()
  await slider.focus(); await slider.press('ArrowRight')
  await expect(input).not.toHaveValue('336699')
  assert.deepEqual(await colors('color'), ['#414657'])
  const adjustedHex = (await input.inputValue()).replace(/^#/, '').toLowerCase()
  await input.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#' + adjustedHex])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  panel = await openColor(); input = await openCustom(panel)
  await chooseColorFormat(panel, 'RGB')
  const red = panel.getByRole('spinbutton', { name: 'R', exact: true })
  await red.fill('10')
  await panel.getByRole('spinbutton', { name: 'G', exact: true }).fill('20')
  await panel.getByRole('spinbutton', { name: 'B', exact: true }).fill('30')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await red.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#0a141e'])
  assert.deepEqual(await history(), { undo: 2, redo: 0 })
})

test('RGB incomplete and out-of-range input cannot confirm, and HSB input converts before applying', async () => {
  await selectRange('A1')
  const panel = await openColor()
  await openCustom(panel); await chooseColorFormat(panel, 'RGB')
  const red = panel.getByRole('spinbutton', { name: 'R', exact: true })
  for (const value of ['', '256']) {
    await red.fill(value); await red.press('Enter')
    await expect(panel.locator('.sheet-color-custom')).toBeVisible()
    assert.deepEqual(await history(), { undo: 0, redo: 0 })
    assert.deepEqual(await colors('color'), ['#414657'])
  }
  await chooseColorFormat(panel, 'HSB')
  const hue = panel.getByRole('spinbutton', { name: 'H', exact: true })
  await hue.fill('120')
  await panel.getByRole('spinbutton', { name: 'S', exact: true }).fill('100')
  await panel.getByRole('spinbutton', { name: 'B', exact: true }).fill('100')
  await hue.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#00ff00'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('the main color remembers the applied choice while the palette follows selection and text reset', async () => {
  await selectRange('A1')
  let panel = await openColor()
  await panel.getByRole('button', { name: '#3F83F8', exact: true }).click()
  await page.evaluate(() => window.sheetColorFixture.externalColor('#f05252'))
  panel = await openColor()
  await expect(panel.getByRole('button', { name: '#F05252', exact: true })).toHaveAttribute('aria-pressed', 'true')
  assert.equal(await mainSwatch(), 'rgb(63, 131, 248)')
  await panel.getByRole('button', { name: '#F05252', exact: true }).press('Escape')
  await (await homeItem('sheet.menu.text-color')).locator('.sheet-color-main').click()
  await expect.poll(() => colors('color')).toEqual(['#3f83f8'])
  panel = await openColor()
  await panel.getByRole('button', { name: '重置文字颜色', exact: true }).click()
  await expect.poll(() => colors('color')).toEqual([null])
  assert.equal(await mainSwatch(), 'rgb(63, 131, 248)')
  panel = await openColor()
  await expect(panel.locator('.sheet-color-state')).toHaveText('自动文字颜色')
  await expect(panel.locator('.sheet-color-preset[aria-pressed="true"]')).toHaveCount(0)
})

test('No Fill clears the stored background and the remembered fill can be applied again', async () => {
  await selectRange('A1')
  assert.equal(await mainSwatch('background'), 'rgb(48, 100, 199)')
  let panel = await openColor('background')
  await panel.getByRole('button', { name: '#FAC815', exact: true }).click()
  panel = await openColor('background')
  await panel.getByRole('button', { name: '无填充', exact: true }).click()
  await expect.poll(() => colors('background')).toEqual([null])
  assert.equal(await mainSwatch('background'), 'rgb(250, 200, 21)')
  await (await homeItem('sheet.menu.fill')).locator('.sheet-color-main').click()
  await expect.poll(() => colors('background')).toEqual(['#fac815'])
  assert.equal((await snapshot()).sheets[0].cells['0:0'].style.font.color, '#414657')
})

test('Ribbon resets override inherited colors and retain the sparse selection through history and reopen', async () => {
  await page.evaluate(() => window.sheetColorFixture.inheritColors())
  await selectRange('A2:B2')
  const before = await snapshot(), textSwatch = await mainSwatch(), fillSwatch = await mainSwatch('background')
  const textPanel = await openColor()
  await textPanel.getByRole('button', { name: '重置文字颜色', exact: true }).click()
  await expect.poll(() => colors('color', [1])).toEqual([null])
  await expect(grid()).toBeFocused()
  const fillPanel = await openColor('background')
  await fillPanel.getByRole('button', { name: '无填充', exact: true }).click()
  await expect.poll(() => colors('background', [1])).toEqual([null])
  await expect(grid()).toBeFocused()
  assert.deepEqual(await history(), { undo: 2, redo: 0 })
  const cleared = await snapshot(), cell = cleared.sheets[0].cells['1:0']
  assert.equal(Object.keys(cleared.sheets[0].cells).length, Object.keys(before.sheets[0].cells).length)
  assert.equal(cleared.sheets[0].cells['1:1'], undefined)
  assert.equal(cell.authoredStyleReference, 'colored')
  assert.equal(cell.authoredStyle.font.color, null)
  assert.equal(cell.authoredStyle.fill, null)
  assert.deepEqual(cell.richText, [{ id: 'source', text: '00123', font: { italic: true },
    hyperlink: 'https://example.test', vendor: { keep: true } }])
  assert.deepEqual(content(cleared), content(before))
  assert.equal(await mainSwatch(), textSwatch)
  assert.equal(await mainSwatch('background'), fillSwatch)
  await (await homeItem('sheet.command.undo')).getByRole('button').click()
  await expect.poll(() => colors('background', [1])).toEqual(['#ffeecc'])
  await (await homeItem('sheet.command.undo')).getByRole('button').click()
  await expect.poll(() => colors('color', [1])).toEqual(['#112233'])
  assert.deepEqual((await snapshot()).sheets[0].cells['1:0'].richText, before.sheets[0].cells['1:0'].richText)
  await (await homeItem('sheet.command.redo')).getByRole('button').click()
  await (await homeItem('sheet.command.redo')).getByRole('button').click()
  const saved = JSON.stringify(await snapshot())
  await page.evaluate(value => window.sheetColorFixture.reopen(value), saved)
  await selectRange('A2:B2')
  await expect.poll(() => colors('color', [1])).toEqual([null])
  await expect.poll(() => colors('background', [1])).toEqual([null])
  assert.deepEqual(content(await snapshot()), content(before))
  const panel = await openColor()
  await expect(panel.locator('.sheet-color-state')).toHaveText('自动文字颜色')
  await expect(panel.locator('.sheet-color-preset[aria-pressed="true"]')).toHaveCount(0)
})

test('API colors update the remembered choice and reapplying the current color adds no history', async () => {
  await selectRange('A1')
  await page.evaluate(() => window.sheetColorFixture.apiColor('color', '#3f83f8'))
  await expect.poll(() => mainSwatch()).toBe('rgb(63, 131, 248)')
  const before = await snapshot(), previousHistory = await history(), panel = await openColor()
  await panel.getByRole('button', { name: '#3F83F8', exact: true }).click()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), previousHistory)
})

test('recent colors and content history remain separate across workbooks', async () => {
  await selectRange('A1')
  let panel = await openColor()
  await panel.getByRole('button', { name: '#3F83F8', exact: true }).click()
  await page.evaluate(() => window.sheetColorFixture.newWorkbook())
  await selectRange('A1')
  assert.equal(await mainSwatch(), 'rgb(0, 0, 0)')
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  panel = await openColor()
  await panel.getByRole('button', { name: '#F05252', exact: true }).click()
  await page.evaluate(() => window.sheetColorFixture.activate('sheet-font-color-fixture'))
  await selectRange('A1')
  assert.equal(await mainSwatch(), 'rgb(63, 131, 248)')
  await (await homeItem('sheet.command.undo')).getByRole('button').click()
  await expect.poll(() => colors('color')).toEqual(['#414657'])
  await page.evaluate(() => window.sheetColorFixture.activate('sheet-font-color-second'))
  await expect.poll(() => colors('color')).toEqual(['#f05252'])
  assert.equal(await mainSwatch(), 'rgb(240, 82, 82)')
})

test('canceled color commands preserve cells, history and the previous remembered color', async () => {
  await selectRange('A1')
  await page.evaluate(() => window.sheetColorFixture.cancelColors(true))
  const before = await snapshot(), swatch = await mainSwatch(), panel = await openColor()
  await panel.getByRole('button', { name: '#3F83F8', exact: true }).click()
  await expect.poll(() => page.evaluate(() => window.sheetColorFixture.notifications())).toHaveLength(1)
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.equal(await mainSwatch(), swatch)
  assert.equal((await commands()).filter(event => event.id === 'sheet.command.set-color').length, 0)
  await expect(ribbonItem('sheet.menu.text-color').locator('.sheet-color-arrow button')).toBeFocused()
})

test('permission changes close an open draft and disable every font or color entry', async () => {
  await selectRange('A1')
  const panel = await openColor(), input = await openCustom(panel)
  await input.fill('112233')
  await page.evaluate(() => window.sheetColorFixture.denyEdit())
  await expect(page.locator('.sheet-color-panel')).toBeHidden()
  for (const id of ['sheet.command.toggle-bold', 'sheet.command.toggle-italic', 'sheet.command.toggle-underline', 'sheet.command.toggle-strike']) await expect((await homeItem(id)).getByRole('button')).toBeDisabled()
  const main = (await homeItem('sheet.menu.text-color')).locator('.sheet-color-main')
  await expect(main).toBeDisabled()
  await main.evaluate(element => element.dispatchEvent(new MouseEvent('click', { bubbles: true })))
  assert.deepEqual(await colors('color'), ['#414657'])
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('portaled palette follows theme and RTL language changes without losing color controls', async () => {
  await selectRange('A1')
  const panel = await openColor(), background = await panel.evaluate(element => getComputedStyle(element).getPropertyValue('--ribbon-background').trim())
  await page.evaluate(() => window.sheetColorFixture.appearance(true, 'ar-SA'))
  await expect(panel).toHaveAttribute('dir', 'rtl')
  await expect(panel.locator('.sheet-color-more button')).toHaveText('المزيد من الألوان')
  await expect.poll(() => panel.evaluate(element => getComputedStyle(element).getPropertyValue('--ribbon-background').trim())).not.toBe(background)
  await expect(panel.locator('.sheet-color-preset')).toHaveCount(45)
  await panel.getByRole('button', { name: '#C81E1E', exact: true }).click()
  await expect.poll(() => colors('color')).toEqual(['#c81e1e'])
  await expect(canvas()).toHaveAttribute('data-color-test-mount', 'original')
})

test('grid overflow retains both color stages and commits a single workbook command', async () => {
  await selectRange('A1')
  await page.evaluate(() => window.sheetColorFixture.variant('grid'))
  await page.setViewportSize({ width: 260, height: 900 })
  const panel = await openColor(), input = await openCustom(panel)
  await input.fill('112233'); await input.press('Enter')
  await expect.poll(() => colors('color')).toEqual(['#112233'])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  assert.equal((await commands()).filter(event => event.id === 'sheet.command.set-color').length, 1)
  await expect(grid()).toBeFocused()
})

test('nested border palette shares the full color control and closes with its parent panel', async () => {
  await selectRange('A1')
  const border = await homeItem('sheet.menu.borders'), arrow = border.getByRole('button', { name: '边框选项', exact: true })
  await arrow.click()
  const parent = page.locator('.border-selector:visible')
  await expect(parent.locator('.border-preset')).toHaveCount(15)
  await parent.getByRole('button', { name: '边框颜色', exact: true }).click()
  let panel = page.locator('.sheet-color-panel:visible')
  await expect(panel.locator('.sheet-color-preset')).toHaveCount(45)
  const draft = await openCustom(panel)
  await draft.fill('abcdef'); await draft.press('Escape')
  await expect(panel.locator('.sheet-color-more button')).toBeFocused()
  await panel.locator('.sheet-color-more button').press('Escape')
  await expect(page.locator('.sheet-color-panel')).toBeHidden()
  await expect(parent).toBeVisible()
  await parent.getByRole('button', { name: '边框颜色', exact: true }).click()
  panel = page.locator('.sheet-color-panel:visible')
  await panel.getByRole('button', { name: '#F05252', exact: true }).click()
  await expect.poll(async () => (await snapshot()).sheets[0].cells['0:0'].style.border.top.color).toBe('#f05252')
  await expect(parent).toBeVisible()
  await parent.getByRole('button', { name: '边框颜色', exact: true }).click()
  panel = page.locator('.sheet-color-panel:visible')
  await openCustom(panel)
  await arrow.click()
  await expect(parent).toBeHidden()
  await expect(page.locator('.sheet-color-panel')).toBeHidden()
  await arrow.click()
  await expect(page.locator('.border-selector:visible')).toBeVisible()
  await expect(page.locator('.sheet-color-panel')).toBeHidden()
})

test('format commands skip filter-hidden rows while manually hidden selected rows remain editable', async () => {
  await page.evaluate(() => window.sheetColorFixture.filteredRows())
  await selectRange('A1:A3')
  const panel = await openColor('background')
  await panel.getByRole('button', { name: '#3F83F8', exact: true }).click()
  await expect.poll(() => colors('background', [0, 1, 2])).toEqual(['#3f83f8', null, '#3f83f8'])
  assert.deepEqual((await snapshot()).sheets[0].hiddenRows, [2])
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('an external color command preserves an active cell draft and its committed text keeps the color', async () => {
  await selectRange('A1')
  await grid().press('F2')
  await expect(cellEditor()).toBeVisible()
  await cellEditor().fill('42')
  await page.evaluate(() => window.sheetColorFixture.apiColor('color', '#3f83f8'))
  await expect(cellEditor()).toHaveValue('42')
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('42')
  await cellEditor().press('Enter')
  await expect.poll(async () => (await snapshot()).sheets[0].cells['0:0'].value).toBe(42)
  assert.deepEqual(await colors('color'), ['#3f83f8'])
  assert.deepEqual(await history(), { undo: 2, redo: 0 })
})
