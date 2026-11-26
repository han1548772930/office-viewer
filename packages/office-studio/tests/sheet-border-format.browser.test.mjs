import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-border-format-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Sheets border commands</title><style>body{margin:0}#app{height:100vh;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model'
import { toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
import { sheetMergedBorderSegments } from '/packages/office-studio/src/features/sheet-workbench/sheet-cell-border'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankWorkbook()
initial.id = 'sheet-border-fixture'; initial.name = '边框'; initial.calculationMode = 'manual'
const sheet = initial.sheets[0]
sheet.id = 'sheet-border-main'; initial.activeSheetId = sheet.id; sheet.rowCount = 200; sheet.columnCount = 26
sheet.cells['0:0'] = { ...createCell(0, 0, '00123'), dataType: 'text' }
sheet.cells['1:0'] = { ...createCell(1, 0, .25), formula: '=1/4', dataType: 'formula', cachedDataType: 'number' }
sheet.cells['2:0'] = createCell(2, 0, 7)
sheet.cells['0:1'] = createCell(0, 1, 'neighbor')
sheet.cells['0:1'].style.border = { left: { style: 'thin', color: '#333333' }, top: { style: 'double', color: '#123456' } }
sheet.cells['1:2'] = createCell(1, 2, 'left of merge')
const merge = { startRow: 0, endRow: 2, startCol: 3, endCol: 3 }
sheet.mergedRanges = [merge]
sheet.cells['0:3'] = createCell(0, 3, '合并')
sheet.cells['0:3'].style.border = { left: { style: 'thin', color: '#3f83f8' }, top: { style: 'thin', color: '#333333' }, bottom: { style: 'thin', color: '#333333' }, right: { style: 'thin', color: '#333333' } }
const studio = new OfficeStudio({ workbook: initial, currentUserId: 'local-user' }), api = FOffice.newAPI(studio)
const view = ref(toExcelWorkbookView(studio.editor.workbook.model)), dark = ref(false), variant = ref('classic'), canceled = ref(false), commands = [], states = [], notifications = []
const direction = computed(() => ['ar-SA', 'fa-IR'].includes(studio.regionalSettings.locale.value) ? 'rtl' : 'ltr')
watch(() => studio.editor.workbook.model, model => { view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
studio.events.addEvent('CommandExecuted', event => commands.push(copy(event)))
studio.events.addEvent('SheetBorderFormatStateChanged', event => states.push(copy(event)))
studio.events.addEvent('BeforeCommandExecute', event => { if (canceled.value && event.id.startsWith('sheet.command.set-border')) event.cancel = true })
window.sheetBorderFixture = {
  snapshot: () => copy(studio.editor.workbook.model),
  history: (id = studio.activeWorkbookId) => { const unit = studio.editor.history.units[id]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0 } },
  state: () => copy(studio.getSheetBorderFormatState(studio.activeWorkbookId, studio.editor.workbook.model.activeSheetId)),
  commands: () => copy(commands), states: () => copy(states), notifications: () => copy(notifications),
  cancel: value => { canceled.value = value },
  appearance: (isDark, language = 'zh-CN') => { dark.value = isDark; studio.regionalSettings.setLocale(language) },
  variant: value => { variant.value = value },
  apiBorder: (type, style, color, address = 'A1') => api.getActiveWorkbook().getActiveSheet().getRange(address).setBorder(type, style, color),
  configure: (field, value, activate = false) => studio.executeCommandSync('sheet.command.set-border-' + field, { value, ...(activate ? { activate: true } : {}) }),
  mergedEdges: () => { const current = studio.editor.workbook.model.sheets[0]; return copy(sheetMergedBorderSegments(merge, (row, col) => current.cells[row + ':' + col]?.style.border)) },
  newWorkbook: () => api.createWorkbook({ ...copy(initial), id: 'sheet-border-second' }),
  activate: id => studio.activateWorkbook(id),
  reopen: snapshot => api.createWorkbook({ ...JSON.parse(snapshot), id: 'sheet-border-reopened' }),
  protect: (row, col) => {
    const next = copy(studio.editor.workbook.model)
    next.sheets[0].rangeProtectionRules = [{ id: 'border-locked', permissionId: 'border-locked-permission', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [{ startRow: row, endRow: row, startCol: col, endCol: col }] }]
    studio.editor.workbook.replace(next)
  },
  hiddenRows: () => {
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
  'onUpdate:modelValue': value => studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value),
  onUndo: () => api.getActiveWorkbook().undo(), onRedo: () => api.getActiveWorkbook().redo(),
  onNotify: value => notifications.push(copy(value)),
}) }) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-border-format-fixture', configureServer(vite) {
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
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(grid()).toBeVisible()
  await expect(canvas()).toHaveCount(1)
  await canvas().evaluate(element => { element.dataset.borderTestMount = 'original' })
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const canvas = () => page.locator('.canvas-surface canvas')
const editor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const snapshot = () => page.evaluate(() => window.sheetBorderFixture.snapshot())
const history = () => page.evaluate(() => window.sheetBorderFixture.history())
const state = () => page.evaluate(() => window.sheetBorderFixture.state())
const executed = () => page.evaluate(() => window.sheetBorderFixture.commands())
const border = async (row = 0, col = 0) => (await snapshot()).sheets[0].cells[row + ':' + col]?.style.border ?? {}
const modelContent = model => Object.fromEntries(Object.entries(model.sheets[0].cells).filter(([, cell]) => cell.value !== '' || cell.formula).map(([key, { value, formula, display, dataType, cachedDataType }]) => [key, { value, formula, display, dataType, cachedDataType }]))
async function homeItem(id) {
  const selector = '[data-ribbon-item="' + id + '"]'
  const item = page.locator('.ribbon-visible-groups ' + selector + ', .ribbon-grid-group:not(.is-hidden) ' + selector)
  if (await item.count() && await item.isVisible()) return item
  const overflow = page.locator('.ribbon-overflow-panel:visible')
  if (!await overflow.isVisible()) await page.locator('.ribbon-more-button:visible, .ribbon-grid-more-button:visible').click()
  const result = overflow.locator(selector); await expect(result).toBeVisible(); return result
}
async function selectRange(address) {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter'); await expect(input).toHaveValue(address); await expect(grid()).toBeFocused()
}
async function openBorder() {
  const item = await homeItem('sheet.menu.borders')
  await item.locator('button').last().click()
  const panel = page.locator('.border-selector:visible'); await expect(panel).toBeVisible(); return panel
}
async function setStyle(panel, label) {
  await panel.getByRole('combobox', { name: '边框样式', exact: true }).click()
  const menu = page.locator('.ant-select-dropdown:visible'); await expect(menu.locator('.ant-select-item-option')).toHaveCount(12)
  await menu.getByText(label, { exact: true }).click()
}
async function openColor(panel) {
  await panel.getByRole('button', { name: '边框颜色', exact: true }).click()
  const color = page.locator('.sheet-color-panel:visible'); await expect(color).toBeVisible(); return color
}
async function contextMenu() {
  await grid().click({ button: 'right', position: { x: 80, y: 42 } })
  const menu = page.locator('.sheet-context-menu:visible'); await expect(menu).toBeVisible(); return menu
}

test('fifteen presets and twelve styles operate on the Canvas through one history entry', async () => {
  await selectRange('A1:A3')
  const before = modelContent(await snapshot()), panel = await openBorder()
  await expect(panel.locator('.border-preset')).toHaveCount(15)
  await setStyle(panel, '双线')
  await expect.poll(async () => (await border()).top).toEqual({ style: 'double', color: '#000000' })
  assert.equal((await executed()).at(-1).id, 'sheet.command.set-border-basic')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  assert.deepEqual(modelContent(await snapshot()), before)
  assert.equal((await border(0, 1)).left, undefined)
  assert.deepEqual((await border(0, 1)).top, { style: 'double', color: '#123456' })
  await expect(canvas()).toHaveAttribute('data-border-test-mount', 'original')
  await expect(page.locator('.excel-workbench table')).toHaveCount(0)
  await expect(page.locator('.excel-workbench [contenteditable="true"]')).toHaveCount(0)
})

test('border custom colors remain draft until confirmation and Escape restores the accepted pen', async () => {
  await selectRange('A1')
  const before = await snapshot(), panel = await openBorder(), color = await openColor(panel)
  await expect(color.locator('.sheet-color-preset')).toHaveCount(45)
  await color.locator('.sheet-color-more button').click()
  const input = color.getByRole('textbox', { name: 'HEX', exact: true })
  await input.fill('abcdef'); assert.deepEqual(await snapshot(), before); await input.press('Escape')
  await expect(color.locator('.sheet-color-more button')).toBeFocused()
  assert.deepEqual(await state(), { type: 'all', style: 'thin', color: '#000000', active: false })
  await color.locator('.sheet-color-more button').click(); await input.fill('abcdef')
  await color.getByRole('button', { name: /确\s*定/ }).click()
  await expect.poll(async () => (await border()).top?.color).toBe('#abcdef')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  assert.equal((await executed()).at(-1).id, 'sheet.command.set-border-basic')
  await expect(panel).toBeVisible(); await expect(page.locator('.sheet-color-panel')).toBeHidden()
})

for (const variant of ['classic', 'grid']) test('the main border icon follows the accepted preset and keeps its repeat action in ' + variant, async () => {
  await page.evaluate(layout => window.sheetBorderFixture.variant(layout), variant)
  await selectRange('A1:A3')
  await expect((await homeItem('sheet.menu.borders')).locator('.ribbon-button-icon .border-preview')).toHaveAttribute('data-border-preset', 'all')
  const panel = await openBorder(); await panel.getByRole('button', { name: '外侧框线', exact: true }).click()
  await expect.poll(async () => (await state()).type).toBe('outside')
  const item = await homeItem('sheet.menu.borders')
  await expect(item.locator('.ribbon-button-icon .border-preview')).toHaveAttribute('data-border-preset', 'outside')
  assert.equal((await executed()).at(-1).id, 'sheet.command.set-borders')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await item.locator('button').first().click()
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  assert.equal((await border(1)).top, undefined)
  assert.equal((await border(1)).bottom, undefined)
  await expect(grid()).toBeFocused()
})

test('right-click presets preserve the selected range and use the registered position command', async () => {
  await selectRange('A1:A3')
  const menu = await contextMenu(); await menu.getByText('边框', { exact: true }).hover()
  const submenu = page.locator('.sheet-context-menu__submenu:visible')
  await expect(submenu.getByRole('menuitem')).toHaveCount(15)
  await submenu.getByText('下框线', { exact: true }).click()
  await expect.poll(async () => (await border(2)).bottom).toEqual({ style: 'thin', color: '#000000' })
  assert.equal((await border()).bottom, undefined)
  assert.equal((await executed()).at(-1).id, 'sheet.command.set-border-position')
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await expect(page.locator('.sheet-context-menu')).toBeHidden(); await expect(grid()).toBeFocused()
})

test('clearing next to an imported merged cell changes one physical segment and survives undo, redo and reopen', async () => {
  await selectRange('C2')
  const before = modelContent(await snapshot()), panel = await openBorder()
  await panel.getByRole('button', { name: '无框线', exact: true }).click()
  const mergedLeft = () => page.evaluate(() => window.sheetBorderFixture.mergedEdges().filter(segment => segment.side === 'left').map(segment => segment.edge))
  const blue = { style: 'thin', color: '#3f83f8' }
  await expect.poll(mergedLeft).toEqual([blue, null, blue])
  assert.deepEqual(modelContent(await snapshot()), before)
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  await (await homeItem('sheet.command.undo')).getByRole('button').click(); await expect.poll(mergedLeft).toEqual([blue, blue, blue])
  await (await homeItem('sheet.command.redo')).getByRole('button').click(); await expect.poll(mergedLeft).toEqual([blue, null, blue])
  const saved = JSON.stringify(await snapshot())
  await page.evaluate(value => { window.sheetBorderFixture.reopen(value) }, saved)
  await expect.poll(mergedLeft).toEqual([blue, null, blue])
  await expect(canvas()).toHaveAttribute('data-border-test-mount', 'original')
})

test('Facade border changes update the open selector and retain their explicit whole-range contract', async () => {
  await selectRange('A1:A3')
  const panel = await openBorder()
  await page.evaluate(() => { window.sheetBorderFixture.apiBorder('outside', 7, '#AbC', 'A1:A3') })
  await expect(panel.getByRole('button', { name: '外侧框线', exact: true })).toHaveAttribute('aria-pressed', 'true')
  await expect((await homeItem('sheet.menu.borders')).locator('.ribbon-button-icon .border-preview')).toHaveAttribute('data-border-preset', 'outside')
  await expect(panel.locator('.ant-select-selection-item')).toHaveText('双线')
  assert.deepEqual(await state(), { type: 'outside', style: 'double', color: '#aabbcc', active: true })
  assert.deepEqual((await executed()).at(-1).params.ranges, [{ startRow: 0, endRow: 2, startCol: 0, endCol: 0 }])
  assert.equal((await border(1)).top, undefined)
})

test('inactive API pens synchronize their displayed labels without adding menu choices or changing cells', async () => {
  await selectRange('A1')
  const before = await snapshot(), panel = await openBorder()
  assert.equal(await page.evaluate(() => window.sheetBorderFixture.configure('style', 12)), false)
  await expect(panel.locator('.ant-select-selection-item')).toHaveText('倾斜点划线')
  assert.equal(await page.evaluate(() => window.sheetBorderFixture.configure('style', 0)), false)
  await expect(panel.locator('.ant-select-selection-item')).toHaveText('无边框线')
  assert.equal(await page.evaluate(() => window.sheetBorderFixture.configure('color', '#AbC')), false)
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await panel.getByRole('combobox', { name: '边框样式', exact: true }).click()
  await expect(page.locator('.ant-select-dropdown:visible .ant-select-item-option')).toHaveCount(12)
  await page.keyboard.press('Escape')
  assert.deepEqual(await state(), { type: 'all', style: 'none', color: '#aabbcc', active: false })
})

test('canceled style and color requests leave the selector and its previous pen unchanged', async () => {
  await selectRange('A1')
  const original = await snapshot(), originalState = await state(), panel = await openBorder()
  await page.evaluate(() => window.sheetBorderFixture.cancel(true))
  await setStyle(panel, '双线')
  await expect(panel.locator('.ant-select-selection-item')).toHaveText('细实线')
  const color = await openColor(panel); await color.getByRole('button', { name: '#F05252', exact: true }).click()
  await expect.poll(async () => page.evaluate(() => window.sheetBorderFixture.notifications().length)).toBe(2)
  assert.deepEqual(await state(), originalState); assert.deepEqual(await snapshot(), original)
  assert.deepEqual(await history(), { undo: 0, redo: 0 }); assert.deepEqual(await executed(), [])
})

test('a protected reverse neighbor rejects UI application without changing content or selected preset', async () => {
  await selectRange('A1')
  await page.evaluate(() => window.sheetBorderFixture.protect(0, 1))
  const original = await snapshot(), panel = await openBorder()
  await expect(panel.getByRole('button', { name: '右框线', exact: true })).toBeEnabled()
  await panel.getByRole('button', { name: '右框线', exact: true }).click()
  await expect.poll(async () => page.evaluate(() => window.sheetBorderFixture.notifications().length)).toBe(1)
  assert.deepEqual(await snapshot(), original); assert.deepEqual(await history(), { undo: 0, redo: 0 })
  assert.deepEqual(await state(), { type: 'all', style: 'thin', color: '#000000', active: false })
  await expect((await homeItem('sheet.menu.borders')).locator('.ribbon-button-icon .border-preview')).toHaveAttribute('data-border-preset', 'all')
  const reopened = await openBorder(); await expect(reopened.getByRole('button', { name: '所有边框', exact: true })).toHaveAttribute('aria-pressed', 'true')
})

test('selection protection disables Ribbon and right-click border choices through the same capability', async () => {
  await selectRange('A1'); await page.evaluate(() => window.sheetBorderFixture.protect(0, 0))
  const item = await homeItem('sheet.menu.borders')
  await expect(item.locator('button').first()).toBeDisabled(); await expect(item.locator('button').last()).toBeDisabled()
  const menu = await contextMenu()
  await expect(menu.getByRole('menuitem', { name: '边框', exact: true })).toHaveAttribute('aria-disabled', 'true')
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
})

test('filtered and manually hidden rows remain included in border operations', async () => {
  await page.evaluate(() => window.sheetBorderFixture.hiddenRows()); await selectRange('A1:A3')
  const panel = await openBorder(); await panel.getByRole('button', { name: '所有边框', exact: true }).click()
  await expect.poll(async () => [await border(0), await border(1), await border(2)].map(value => value.left?.color)).toEqual(['#000000', '#000000', '#000000'])
  assert.deepEqual((await snapshot()).sheets[0].hiddenRows, [2]); assert.deepEqual(await history(), { undo: 1, redo: 0 })
})

test('external API border updates preserve an active input draft and its later committed value', async () => {
  await selectRange('A1'); await grid().press('F2'); await expect(editor()).toBeVisible(); await editor().fill('42')
  await page.evaluate(() => { window.sheetBorderFixture.apiBorder('left', 'thick', '#ff0000') })
  await expect(editor()).toHaveValue('42'); await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('42')
  await editor().press('Enter')
  await expect.poll(async () => (await snapshot()).sheets[0].cells['0:0'].value).toBe(42)
  assert.deepEqual((await border()).left, { style: 'thick', color: '#ff0000' }); assert.deepEqual(await history(), { undo: 2, redo: 0 })
  await expect(grid()).toBeFocused()
})

test('recent border pens and content histories remain isolated across workbooks', async () => {
  await selectRange('A1')
  await page.evaluate(() => { window.sheetBorderFixture.apiBorder('right', 'double', '#ff0000') })
  await page.evaluate(() => { window.sheetBorderFixture.newWorkbook() })
  await expect.poll(state).toEqual({ type: 'all', style: 'thin', color: '#000000', active: false })
  await selectRange('A1'); assert.equal(await page.evaluate(() => window.sheetBorderFixture.configure('style', 'medium')), false)
  assert.deepEqual(await history(), { undo: 0, redo: 0 })
  await page.evaluate(() => window.sheetBorderFixture.activate('sheet-border-fixture'))
  await expect.poll(state).toEqual({ type: 'right', style: 'double', color: '#ff0000', active: true })
  assert.deepEqual(await history(), { undo: 1, redo: 0 })
  const panel = await openBorder(); await expect(panel.locator('.ant-select-selection-item')).toHaveText('双线')
})

test('grid overflow preserves border controls and right-click surfaces follow theme and RTL tokens', async () => {
  await page.evaluate(() => window.sheetBorderFixture.variant('grid')); await page.setViewportSize({ width: 860, height: 900 })
  const panel = await openBorder(); await expect(panel.locator('.border-preset')).toHaveCount(15)
  await setStyle(panel, '双线'); await expect.poll(async () => (await state()).style).toBe('double')
  await page.keyboard.press('Escape'); await grid().focus()
  let menu = await contextMenu(), light = await menu.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.keyboard.press('Escape'); await page.evaluate(() => window.sheetBorderFixture.appearance(true, 'ar-SA'))
  menu = await contextMenu()
  await expect.poll(() => menu.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  assert.equal(await menu.evaluate(element => getComputedStyle(element).direction), 'rtl')
  await expect(menu.locator('.ant-menu')).toBeVisible()
  await expect(canvas()).toHaveAttribute('data-border-test-mount', 'original')
})
