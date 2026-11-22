import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-number-format-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Sheets number format commands</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model'
import { toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankWorkbook()
initial.id = 'sheet-number-format-fixture'; initial.name = '数字格式'; initial.calculationMode = 'manual'
const sheet = initial.sheets[0]
sheet.rowCount = 200; sheet.columnCount = 26
sheet.cells['0:0'] = createCell(0, 0, 12.25)
sheet.cells['0:1'] = { ...createCell(0, 1, '00123'), dataType: 'text' }
sheet.cells['0:2'] = { ...createCell(0, 2, .1225), formula: '=A1/100', dataType: 'formula', cachedDataType: 'number' }
sheet.cells['0:3'] = { ...createCell(0, 3, null), formula: '=UNSUPPORTED()', dataType: 'formula', cachedDataType: null, display: 'source 25%' }
sheet.cells['1:0'] = createCell(1, 0, .25)
const studio = new OfficeStudio({ workbook: initial, currentUserId: 'local-user' })
const api = FOffice.newAPI(studio)
const view = ref(toExcelWorkbookView(studio.editor.workbook.model))
const dark = ref(false), variant = ref('classic'), canceled = ref(false), commands = [], notifications = []
const direction = computed(() => ['ar-SA', 'fa-IR'].includes(studio.regionalSettings.locale.value) ? 'rtl' : 'ltr')
watch(() => studio.editor.workbook.model, model => { view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
studio.events.addEvent('CommandExecuted', event => { commands.push(copy(event)) })
studio.events.addEvent('BeforeCommandExecute', event => { if (canceled.value && event.id === 'sheet.command.set-number-format') event.cancel = true })
window.sheetFormatFixture = {
  snapshot: () => copy(studio.editor.workbook.model),
  history: () => {
    const unit = studio.editor.history.units[studio.activeWorkbookId]
    return { undo: unit.undo.length, redo: unit.redo.length }
  },
  commands: () => copy(commands),
  notifications: () => copy(notifications),
  cancelFormats: value => { canceled.value = value },
  appearance: (isDark, language = 'zh-CN') => { dark.value = isDark; studio.regionalSettings.setLocale(language) },
  region: value => { studio.regionalSettings.setRegion(value) },
  variant: value => { variant.value = value },
  reopen: snapshot => { api.createWorkbook({ ...JSON.parse(snapshot), id: 'sheet-number-format-reopened' }) },
  restrictFirstCell: () => {
    const workbook = api.getActiveWorkbook(), sheet = workbook.getActiveSheet()
    sheet.getRange('A1').setValue(12)
    sheet.setDataValidation('A1', { type: 'whole', operator: 'between', formula1: '1', formula2: '20',
      allowBlank: false, errorStyle: 'stop', error: '请输入 1 到 20 之间的整数。' })
    studio.editor.history.clear(workbook.getId())
  },
  denyEdit: () => {
    const next = copy(studio.editor.workbook.model)
    next.sheets[0].rangeProtectionRules = [{ id: 'denied-format', permissionId: 'denied-format-permission',
      ownerUserId: 'remote-owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '',
      ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 3 }] }]
    studio.editor.workbook.replace(next)
  },
}
createApp({ render: () => h(ConfigProvider, {
  direction: direction.value, theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#3064c7' } },
}, { default: () => h(ExcelWorkbench, {
  modelValue: view.value, studio, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
  'onUpdate:modelValue': value => { studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value) },
  onUndo: () => { api.getActiveWorkbook().undo() },
  onRedo: () => { api.getActiveWorkbook().redo() },
  onNotify: value => { notifications.push(copy(value)) },
}) }) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-number-format-fixture',
    configureServer(vite) {
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
  await canvas().evaluate(element => { element.dataset.formatTestMount = 'original' })
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const snapshot = () => page.evaluate(() => window.sheetFormatFixture.snapshot())
const history = () => page.evaluate(() => window.sheetFormatFixture.history())
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const canvas = () => page.locator('.canvas-surface canvas')
const cellText = () => grid().getByRole('gridcell')
const cellEditor = () => page.locator('.canvas-cell-input input')
const formula = () => page.getByRole('textbox', { name: '公式编辑', exact: true })
const ribbonItem = id => page.locator('.ribbon-visible-groups [data-ribbon-item="' + id + '"]')
const content = model => Object.entries(model.sheets[0].cells).map(([key, { value, formula, dataType, cachedDataType }]) => [key, { value, formula, dataType, cachedDataType }])
async function homeItem(id) {
  const visible = ribbonItem(id)
  if (await visible.count() && await visible.isVisible()) return visible
  const overflow = page.locator('.ribbon-overflow-panel:visible')
  if (!await overflow.isVisible()) await page.locator('.ribbon-more-button').click()
  const item = overflow.locator('[data-ribbon-item="' + id + '"]')
  await expect(item).toBeVisible()
  return item
}
async function selectRange(address) {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter')
  await expect(input).toHaveValue(address)
  await expect(grid()).toBeFocused()
}
async function openFormats(keyboard = false) {
  const item = await homeItem('sheet.command.set-number-format')
  const trigger = item.locator('.sheet-number-format-trigger')
  if (keyboard) { await trigger.focus(); await trigger.press('ArrowDown') } else await trigger.click()
  const menu = page.locator('.sheet-number-format-menu:visible')
  await expect(menu).toBeVisible()
  return menu
}
async function chooseFormat(label) { const menu = await openFormats(); await menu.getByRole('menuitem', { name: label, exact: false }).click() }
async function openMore() { const menu = await openFormats(); await menu.getByRole('menuitem', { name: '更多格式', exact: true }).click(); const panel = page.locator('.sheet-number-format-panel'); await expect(panel).toBeVisible(); return panel }
async function chooseCategory(panel, label) { await panel.getByRole('combobox', { name: '格式类型', exact: true }).click(); await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click() }
async function customInput(panel) { await chooseCategory(panel, '自定义格式'); return panel.getByRole('textbox', { name: '自定义数字格式代码', exact: true }) }
async function clickCommand(id) { const item = await homeItem(id); await item.getByRole('button').click() }
async function editCell(address) { await selectRange(address); await grid().press('F2'); await expect(cellEditor()).toBeFocused(); return cellEditor() }
async function dispatchKey(input, options) {
  return input.evaluate((element, info) => { const event = new KeyboardEvent('keydown', { ...info, bubbles: true, cancelable: true }); element.dispatchEvent(event); return event.defaultPrevented }, options)
}

test('the ordered format menu and percent button apply their distinct defaults through public history', async () => {
  await selectRange('A1')
  const menu = await openFormats(true)
  assert.deepEqual(await menu.locator('.sheet-number-format-row > span:first-child').allTextContents(), [
    '常规', '文本', '数值', '百分比', '科学记数', '会计', '财务数值', '货币', '货币取整', '日期', '时间', '日期时间', '持续时间', '更多格式',
  ])
  await menu.getByRole('menuitem', { name: /^数值/ }).click()
  await expect(cellText()).toHaveText('A1 12')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].style.numberFormat, '0')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await clickCommand('sheet.command.undo')
  await expect(cellText()).toHaveText('A1 12.25')
  await clickCommand('sheet.command.redo')
  await chooseFormat('百分比')
  await expect(cellText()).toHaveText('A1 1225.00%')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].style.numberFormat, '0.00%')
  await clickCommand('sheet.number-format.percent')
  await expect(cellText()).toHaveText('A1 1225%')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 12.25)
  assert.equal((await snapshot()).sheets[0].cells['0:0'].style.numberFormat, '0%')
  assert.deepEqual(await history(), { undo: 3, redo: 0 })
  await expect(grid()).toBeFocused()
  const commands = await page.evaluate(() => window.sheetFormatFixture.commands())
  assert.deepEqual(commands.filter(event => event.id === 'sheet.command.set-number-format' || event.id === 'sheet.command.set-percent-format').map(event => [event.id, event.kind]), [
    ['sheet.command.set-number-format', 'workbook'], ['sheet.command.set-number-format', 'workbook'], ['sheet.command.set-percent-format', 'workbook'],
  ])
})

test('a range preview leaves text and formula caches intact until confirm, and cancel restores Canvas focus', async () => {
  await selectRange('A1:D1')
  const before = await snapshot()
  let panel = await openMore()
  let input = await customInput(panel)
  await input.fill('0.0')
  await expect(cellText()).toHaveText('A1 12.25')
  await input.press('Tab')
  await expect(cellText()).toHaveText('A1 12.3')
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await panel.getByRole('button', { name: /取.*消/ }).click()
  await expect(panel).toBeHidden()
  await expect(cellText()).toHaveText('A1 12.25')
  await expect(grid()).toBeFocused()
  await expect(canvas()).toHaveAttribute('data-format-test-mount', 'original')
  panel = await openMore(); input = await customInput(panel)
  await expect(input).not.toHaveValue('0.0')
  await input.fill('0.000'); await input.press('Tab')
  assert.deepEqual(await snapshot(), before)
  await panel.getByRole('button', { name: /确.*认/ }).click()
  await expect(panel).toBeHidden()
  await expect(cellText()).toHaveText('A1 12.250')
  const after = await snapshot()
  assert.deepEqual(content(after), content(before))
  assert.equal(after.sheets[0].cells['0:1'].display, '00123')
  assert.equal(after.sheets[0].cells['0:3'].display, 'source 25%')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(grid()).toBeFocused()
  await expect(canvas()).toHaveAttribute('data-format-test-mount', 'original')
  await clickCommand('sheet.command.undo')
  await expect(cellText()).toHaveText('A1 12.25')
  assert.deepEqual((await snapshot()).sheets[0].cells, before.sheets[0].cells)
  await clickCommand('sheet.command.redo')
  await page.evaluate(value => window.sheetFormatFixture.reopen(value), JSON.stringify(await snapshot()))
  await selectRange('A1')
  await expect(cellText()).toHaveText('A1 12.250')
  assert.deepEqual(content(await snapshot()), content(before))
})

test('invalid drafts disable confirm and a canceled public command keeps the panel open without content history', async () => {
  await selectRange('A1')
  const before = await snapshot()
  const panel = await openMore()
  const input = await customInput(panel)
  await input.fill('0"unfinished'); await input.press('Tab')
  await expect(panel.getByRole('button', { name: /确.*认/ })).toBeDisabled()
  await expect(panel.locator('.ant-form-item-explain-error')).toBeVisible()
  assert.deepEqual(await snapshot(), before)
  await input.fill('0.000'); await input.press('Enter')
  await expect(cellText()).toHaveText('A1 12.250')
  await page.evaluate(() => window.sheetFormatFixture.cancelFormats(true))
  await panel.getByRole('button', { name: /确.*认/ }).click()
  await expect(panel).toBeVisible()
  await expect(panel.getByRole('alert')).toContainText('数字格式未应用')
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  const notices = await page.evaluate(() => window.sheetFormatFixture.notifications())
  assert.ok(notices.some(notice => notice.message.includes('canceled')))
  await page.evaluate(() => window.sheetFormatFixture.cancelFormats(false))
  await panel.getByRole('button', { name: /确.*认/ }).click()
  await expect(panel).toBeHidden()
  await expect(cellText()).toHaveText('A1 12.250')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('a permission change discards an open format preview and disables all formatting entry points', async () => {
  await selectRange('A1:D1')
  const before = await snapshot()
  const panel = await openMore()
  const input = await customInput(panel)
  await input.fill('0.000'); await input.press('Tab')
  await expect(cellText()).toHaveText('A1 12.250')
  await page.evaluate(() => window.sheetFormatFixture.denyEdit())
  await expect(panel).toBeHidden()
  await expect(cellText()).toHaveText('A1 12.25')
  const trigger = (await homeItem('sheet.command.set-number-format')).locator('.sheet-number-format-trigger')
  const percent = (await homeItem('sheet.number-format.percent')).getByRole('button')
  await expect(trigger).toBeDisabled(); await expect(percent).toBeDisabled()
  await percent.evaluate(element => element.dispatchEvent(new MouseEvent('click', { bubbles: true })))
  assert.deepEqual((await snapshot()).sheets[0].cells, before.sheets[0].cells)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await expect(canvas()).toHaveAttribute('data-format-test-mount', 'original')
})

test('text format takes precedence over formulas and forced text returns its apostrophe when editing', async () => {
  await selectRange('B1'); await chooseFormat('文本')
  let input = await editCell('B1')
  await expect(input).toHaveValue('00123'); await expect(input).toHaveCSS('text-align', 'left')
  await input.fill('=1+1'); await input.press('Enter')
  let cell = (await snapshot()).sheets[0].cells['0:1']
  assert.equal(cell.value, '=1+1'); assert.equal(cell.formula, null); assert.equal(cell.dataType, 'text')
  await selectRange('B1'); await chooseFormat('常规')
  input = await editCell('B1'); await input.fill("'00123"); await input.press('Enter')
  await selectRange('B1')
  await expect(formula()).toHaveValue("'00123")
  await expect(cellText()).toHaveText('B1 00123')
  input = await editCell('B1')
  await expect(input).toHaveValue("'00123"); await expect(input).toHaveCSS('text-align', 'left')
  await input.fill('42'); await input.press('Enter')
  cell = (await snapshot()).sheets[0].cells['0:1']
  assert.equal(cell.value, 42); assert.equal(cell.dataType, 'number'); assert.equal(cell.cachedDataType, null)
  input = await editCell('B1'); await expect(input).toHaveCSS('text-align', 'right')
  const beforeCancel = await history()
  await input.fill('77'); await input.press('Escape')
  assert.equal((await snapshot()).sheets[0].cells['0:1'].value, 42)
  assert.deepEqual(await history(), beforeCancel)
  await expect(grid()).toBeFocused()
})

test('cell composition and keyCode 229 keep Enter and Escape out of document history until the edit commits', async () => {
  await selectRange('A1')
  const before = await snapshot()
  const input = await editCell('A1')
  await input.fill('12.5')
  await input.dispatchEvent('compositionstart', { data: '12.5' })
  for (const options of [
    { key: 'Enter', isComposing: true }, { key: 'Escape', isComposing: true },
    { key: 'Enter', keyCode: 229 }, { key: 'Escape', keyCode: 229 }, { key: 'Enter' },
  ]) {
    assert.equal(await dispatchKey(input, options), false)
    await expect(input).toBeVisible()
    assert.deepEqual(await snapshot(), before)
    assert.deepEqual(await history(), { undo: 0, redo: 0 })
  }
  await input.dispatchEvent('compositionend', { data: '12.5' }); await input.press('Enter')
  await expect(input).toBeHidden()
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 12.5)
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(grid()).toBeFocused()
})

test('custom format composition leaves the draft open and publishes a preview only after composition ends', async () => {
  await selectRange('A1')
  const before = await snapshot()
  const panel = await openMore()
  const input = await customInput(panel)
  await input.fill('0.000')
  for (const options of [{ key: 'Enter', isComposing: true }, { key: 'Enter', keyCode: 229 }]) {
    assert.equal(await dispatchKey(input, options), false)
    await expect(cellText()).toHaveText('A1 12.25')
  }
  await input.dispatchEvent('compositionstart')
  for (const options of [{ key: 'Enter', isComposing: true }, { key: 'Escape', keyCode: 229 }, { key: 'Escape' }]) {
    assert.equal(await dispatchKey(input, options), false)
    await expect(panel).toBeVisible()
    await expect(cellText()).toHaveText('A1 12.25')
  }
  assert.deepEqual(await snapshot(), before)
  await input.dispatchEvent('compositionend'); await input.press('Enter')
  await expect(cellText()).toHaveText('A1 12.250')
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await input.press('Escape')
  await expect(panel).toBeHidden(); await expect(grid()).toBeFocused()
  await expect(cellText()).toHaveText('A1 12.25')
})

test('menu, sidebar and cell input follow theme and direction without changing content or replacing the Canvas', async () => {
  await selectRange('A1')
  const before = await snapshot()
  const menu = await openFormats()
  const menuBackground = await menu.evaluate(element => getComputedStyle(element).backgroundColor)
  const canvasBackground = await grid().evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.sheetFormatFixture.appearance(true, 'ar-SA'))
  await expect(menu).toHaveAttribute('dir', 'rtl')
  await expect.poll(() => menu.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(menuBackground)
  await expect.poll(() => grid().evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(canvasBackground)
  await page.evaluate(() => window.sheetFormatFixture.appearance(true, 'zh-CN'))
  await menu.getByRole('menuitem').first().press('Escape')
  const panel = await openMore()
  const panelBackground = await panel.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.sheetFormatFixture.appearance(false))
  await expect.poll(() => panel.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(panelBackground)
  await panel.getByRole('button', { name: '关闭', exact: true }).click()
  const input = await editCell('A1')
  await expect(input).toHaveCSS('text-align', 'right')
  const border = await page.locator('.canvas-cell-input').evaluate(element => getComputedStyle(element).borderColor)
  assert.equal(border, 'rgb(48, 100, 199)')
  await input.press('Escape')
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await expect(canvas()).toHaveAttribute('data-format-test-mount', 'original')
})

test('currency controls preview negative styles and decimals, then share toolbar increments and no-op history', async () => {
  await selectRange('A1')
  const before = await snapshot()
  let panel = await openMore()
  await chooseCategory(panel, '货币')
  const decimals = panel.getByRole('spinbutton', { name: '小数位数', exact: true })
  await decimals.fill('3'); await decimals.press('Tab')
  await panel.getByRole('menu', { name: '负数类型', exact: true }).getByRole('menuitem').nth(3).click()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await panel.getByRole('button', { name: /确.*认/ }).click()
  const cell = (await snapshot()).sheets[0].cells['0:0']
  assert.equal(cell.style.numberFormat, '"¥"#,##0.000_);[Red]("¥"#,##0.000)')
  assert.equal(cell.value, 12.25)
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await clickCommand('sheet.command.increase-decimals')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].style.numberFormat, '"¥"#,##0.0000_);[Red]("¥"#,##0.0000)')
  await clickCommand('sheet.command.decrease-decimals')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].style.numberFormat, cell.style.numberFormat)
  const depth = await history()
  panel = await openMore()
  await panel.getByRole('button', { name: /确.*认/ }).click()
  await expect(panel).toBeHidden()
  assert.deepEqual(await history(), depth)
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 12.25)
})

test('name and formula input composition does not navigate or commit, and address errors use the shared Alert theme', async () => {
  await selectRange('A1')
  const before = await snapshot()
  const address = page.getByRole('textbox', { name: '名称框', exact: true })
  await address.fill('B2')
  for (const options of [{ key: 'Enter', isComposing: true }, { key: 'Enter', keyCode: 229 }]) {
    assert.equal(await dispatchKey(address, options), false)
    await expect(cellText()).toHaveText('A1 12.25')
  }
  await address.dispatchEvent('compositionstart')
  assert.equal(await dispatchKey(address, { key: 'Enter' }), false)
  await address.dispatchEvent('compositionend'); await address.press('Enter')
  await expect(cellText()).toHaveText('B2 ')
  await address.fill('XFE1'); await address.press('Enter')
  const error = page.locator('.formula-error.ant-alert')
  await expect(error).toBeVisible(); await expect(address).toHaveAttribute('aria-invalid', 'true')
  const background = await error.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.sheetFormatFixture.appearance(true))
  await expect.poll(() => error.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(background)
  await address.press('Escape')
  await expect(error).toBeHidden(); await expect(address).toHaveValue('B2'); await expect(grid()).toBeFocused()
  const input = formula()
  await input.fill('=1+1'); await input.dispatchEvent('compositionstart')
  for (const options of [{ key: 'Enter', isComposing: true }, { key: 'Escape', keyCode: 229 }, { key: 'Enter' }]) assert.equal(await dispatchKey(input, options), false)
  assert.deepEqual(await snapshot(), before)
  await input.dispatchEvent('compositionend'); await input.press('Escape')
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await expect(grid()).toBeFocused()
})

test('a rejected cell edit keeps its selected address when the name box requests a different cell', async () => {
  await page.evaluate(() => window.sheetFormatFixture.restrictFirstCell())
  const input = await editCell('A1')
  const before = await snapshot()
  await input.fill('999')
  const address = page.getByRole('textbox', { name: '名称框', exact: true })
  await address.fill('B2'); await address.press('Enter')
  await expect(cellText()).toHaveText('A1 12')
  await expect(input).toBeVisible()
  assert.deepEqual(await snapshot(), before)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await input.fill('5'); await input.press('Enter')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 5)
  await selectRange('B2')
  await expect(cellText()).toHaveText('B2 ')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('grid overflow retains the format menu and applies a format to the original Canvas selection', async () => {
  await selectRange('A1')
  await page.evaluate(() => window.sheetFormatFixture.variant('grid'))
  await page.setViewportSize({ width: 360, height: 900 })
  const menu = await openFormats()
  const bounds = await menu.boundingBox()
  assert.ok(bounds && bounds.x >= 0 && bounds.x + bounds.width <= 360)
  await menu.getByRole('menuitem', { name: /^百分比/ }).click()
  await expect(cellText()).toHaveText('A1 1225.00%')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(grid()).toBeFocused()
  await expect(canvas()).toHaveAttribute('data-format-test-mount', 'original')
})
