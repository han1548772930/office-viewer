import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-range-editor-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Sheets range editing</title><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model'
import { applyExcelWorkbookView, parseCellKey, toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
import { SpreadsheetGeometry } from '/packages/office-studio/src/rendering/spreadsheet/layout'
const copy = value => JSON.parse(JSON.stringify(value))
const initial = createBlankWorkbook()
initial.id = 'range-edit-fixture'; initial.name = '范围编辑'; initial.calculationMode = 'manual'
const sheet = initial.sheets[0]
sheet.defaultColumnWidth = 90; sheet.defaultRowHeight = 26
sheet.cells['0:0'] = createCell(0, 0, null)
sheet.cells['0:1'] = { ...createCell(0, 1, '首项'), richText: [{ text: '首项', style: { italic: true } }], metadata: { vendor: 'kept' } }
sheet.cells['0:1'].style.font.bold = true; sheet.cells['0:1'].style.fill = '#abcdef'
sheet.cells['0:2'] = { ...createCell(0, 2, 11), formula: '=5+6', display: '11', dataType: 'formula', cachedDataType: 'number', hyperlink: 'https://example.test', metadata: { formulaError: 'obsolete', extension: 'kept' } }
sheet.cells['0:2'].style.font.italic = true
const studio = new OfficeStudio({ workbook: initial }), api = FOffice.newAPI(studio)
const standalone = new URLSearchParams(location.search).has('standalone')
let base = copy(initial)
const view = ref(toExcelWorkbookView(initial)), dark = ref(false), variant = ref('classic'), notices = [], requests = [], emitted = []
watch(() => studio.editor.workbook.model, model => { if (!standalone) view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
studio.events.addEvent('SheetMergeConfirmationRequested', event => requests.push(copy(event)))
const current = () => standalone ? base : studio.editor.workbook.model
window.rangeEditorFixture = {
  snapshot: () => copy(current()), notifications: () => copy(notices), requests: () => copy(requests),
  history: () => { const unit = studio.editor.history.units[studio.activeWorkbookId]; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, emits: emitted.length } },
  appearance: (value, layout) => { dark.value = value; variant.value = layout },
  command: (id, ranges, extra = {}) => studio.executeCommandSync(id, { unitId: current().id, sheetId: current().sheets[0].id, ranges, ...extra }),
  apiClear: (method, address) => api.getActiveWorkbook().getActiveSheet().getRange(address)[method](),
  change: () => api.getActiveWorkbook().getActiveSheet().getRange('D3').setValue('new revision'),
  protect: () => {
    const model = copy(current()); model.sheets[0].rangeProtectionRules = [{ id: 'locked', permissionId: 'locked-p', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 2 }] }]
    if (standalone) { base = model; view.value = toExcelWorkbookView(model) } else studio.editor.workbook.replace(model)
  },
  seed: (mode) => {
    const model = copy(initial), sheet = model.sheets[0]
    if (mode === 'merged') {
      sheet.cells = { '0:0': createCell(0, 0, 'first'), '4:4': createCell(4, 4, 'second') }
      sheet.mergedRanges = [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }, { startRow: 4, endRow: 4, startCol: 4, endCol: 5 }]
      sheet.cells['0:0'].style.font.bold = true
    }
    if (mode === 'filtered') {
      sheet.cells = {}; for (let row = 0; row < 3; row++) sheet.cells[row + ':0'] = createCell(row, 0, 'row ' + row)
      sheet.hiddenRows = [2]; sheet.autoFilter = { range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 }, columns: { 0: { kind: 'values', selectedItems: ['row 2'] } }, filteredRows: [1] }
    }
    if (standalone) { base = model; view.value = toExcelWorkbookView(model) } else { studio.editor.workbook.replace(model); studio.editor.history.clear(studio.activeWorkbookId) }
    emitted.length = 0
  },
  save: () => JSON.stringify(current()),
  reopen: value => api.createWorkbook({ ...JSON.parse(value), id: 'range-edit-reopened' }),
  switchWorkbook: () => { const model = createBlankWorkbook(); model.id = 'range-other'; model.name = '其他工作簿'; api.createWorkbook(model) },
  rect: address => {
    const position = parseCellKey(address), element = document.querySelector('.spreadsheet-canvas'), scroll = element.querySelector('.canvas-scroll')
    const bounds = element.getBoundingClientRect(), geometry = new SpreadsheetGeometry(view.value.sheets[0])
    const rect = geometry.cellRect(position.row, position.col, { width: bounds.width, height: bounds.height, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 1 })
    return { x: bounds.left + rect.x + Math.min(12, rect.width / 2), y: bounds.top + rect.y + rect.height / 2 }
  },
}
createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm } }, {
  default: () => h(ExcelWorkbench, { modelValue: view.value, studio: standalone ? undefined : studio,
    canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { emitted.push(copy(value)); if (standalone) { base = applyExcelWorkbookView(base, value); view.value = value } else studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value) },
    onUndo: () => api.getActiveWorkbook().undo(), onRedo: () => api.getActiveWorkbook().redo(), onNotify: value => notices.push(copy(value)),
  }),
}) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-range-editor-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => {
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
  errors = []; page = await browser.newPage({ viewport: { width: 1800, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const modal = () => page.getByRole('dialog', { name: '合并单元格', exact: true })
const snapshot = () => page.evaluate(() => window.rangeEditorFixture.snapshot())
const history = () => page.evaluate(() => window.rangeEditorFixture.history())
async function select(address) {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter'); await expect(grid()).toBeFocused()
}
async function menu(id, label) {
  const item = page.locator('[data-ribbon-item="' + id + '"]:visible').first()
  await expect(item).toBeVisible(); await item.getByRole('button').click()
  await page.locator('.ribbon-command-menu:visible').getByRole('menuitem', { name: label, exact: true }).click()
}
async function command(id) { await page.locator('[data-ribbon-item="' + id + '"]:visible').first().getByRole('button').click() }
async function contextClear(label, address = 'A1') {
  const point = await page.evaluate(value => window.rangeEditorFixture.rect(value), address)
  await page.mouse.click(point.x, point.y, { button: 'right' })
  await page.getByRole('menuitem', { name: '清除', exact: true }).hover()
  await page.getByRole('menuitem', { name: label, exact: true }).click()
}

test('toolbar merge confirmation cancels without history and restores Canvas focus', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await expect(modal()).toBeVisible(); await expect(modal()).toContainText('其余 1 项内容')
  assert.equal((await history()).undo, 0); assert.equal((await snapshot()).sheets[0].cells['0:2'].formula, '=5+6')
  await modal().getByRole('button', { name: '取消', exact: true }).click()
  await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
  assert.equal((await history()).undo, 0); assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [])
})

test('accepted merge preserves the donor and non-master styles; undo/redo and internal reopen use one history', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
  const saved = await snapshot(), sheet = saved.sheets[0]
  assert.equal(sheet.cells['0:0'].value, '首项'); assert.deepEqual(sheet.cells['0:0'].richText, [{ text: '首项', style: { italic: true } }])
  assert.deepEqual(sheet.cells['0:0'].metadata, { vendor: 'kept' }); assert.equal(sheet.cells['0:2'].style.font.italic, true)
  assert.equal(sheet.cells['0:2'].formula, null); assert.equal(sheet.cells['0:2'].hyperlink, null); assert.equal(sheet.cells['0:2'].metadata.formulaError, undefined)
  assert.equal((await history()).undo, 1)
  await command('sheet.command.undo'); assert.equal((await snapshot()).sheets[0].cells['0:2'].formula, '=5+6')
  await command('sheet.command.redo'); assert.equal((await snapshot()).sheets[0].cells['0:2'].formula, null)
  await page.evaluate(value => window.rangeEditorFixture.reopen(JSON.stringify(value)), saved)
  await select('A1'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '首项')
})

test('Delete, right-click format clear and public clear share content, topology and history behavior', async () => {
  await select('C1'); await grid().press('Delete')
  const cleared = (await snapshot()).sheets[0].cells['0:2']
  assert.equal(cleared.value, null); assert.equal(cleared.formula, null); assert.equal(cleared.display, ''); assert.equal(cleared.style.font.italic, true)
  assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.rangeEditorFixture.seed('merged')); await select('A1')
  await contextClear('清除格式')
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [{ startRow: 4, endRow: 4, startCol: 4, endCol: 5 }])
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'first'); assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.rangeEditorFixture.apiClear('clear', 'E5:F5'))
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, []); assert.equal((await snapshot()).sheets[0].cells['4:4'], undefined)
  assert.equal((await history()).undo, 2)
})

test('range clearing skips filter-hidden rows and processes manually hidden rows using the pre-edit filter scope', async () => {
  await page.evaluate(() => window.rangeEditorFixture.seed('filtered')); await select('A1:A3'); await grid().press('Delete')
  const sheet = (await snapshot()).sheets[0]
  assert.equal(sheet.cells['1:0'].value, 'row 1'); assert.equal(sheet.cells['2:0'].value, null)
  assert.deepEqual(sheet.hiddenRows, [2]); assert.equal((await history()).undo, 1)
})

test('a content revision arriving during confirmation rejects the original request with visible feedback', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await page.evaluate(() => window.rangeEditorFixture.change())
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal().getByRole('alert')).toContainText('工作簿已更改')
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, []); assert.equal((await history()).undo, 1)
  await modal().getByRole('button', { name: '取消', exact: true }).click(); await expect(grid()).toBeFocused()
})

test('a permission change during confirmation never clears the protected contents', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await page.evaluate(() => window.rangeEditorFixture.protect())
  await modal().getByRole('button', { name: '继续合并', exact: true }).click()
  await expect(modal().getByRole('alert')).toContainText('工作簿已更改')
  assert.equal((await snapshot()).sheets[0].cells['0:1'].value, '首项'); assert.equal((await history()).undo, 0)
  await modal().getByRole('button', { name: '取消', exact: true }).click()
  const item = page.locator('[data-ribbon-item="sheet.menu.merge"]:visible').first()
  await expect(item.getByRole('button')).toBeDisabled()
})

test('unmerge handles two disjoint selections instead of only the active master', async () => {
  await page.evaluate(() => window.rangeEditorFixture.seed('merged')); await select('A1')
  const point = await page.evaluate(() => window.rangeEditorFixture.rect('E5'))
  await page.keyboard.down('Control'); await page.mouse.click(point.x, point.y); await page.keyboard.up('Control')
  await menu('sheet.menu.merge', '取消合并')
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, []); assert.equal((await history()).undo, 1)
  assert.equal((await snapshot()).sheets[0].cells['0:1'].style.font.bold, true)
})

test('switching workbooks closes the old confirmation and does not mutate either workbook', async () => {
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await page.evaluate(() => window.rangeEditorFixture.switchWorkbook())
  await expect(modal()).toBeHidden(); await expect(grid()).toBeVisible()
  assert.equal((await snapshot()).id, 'range-other'); assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [])
  assert.equal((await history()).undo, 0)
})

test('the controlled confirmation inherits the dark theme in both Ribbon layouts and returns focus after Escape', async () => {
  for (const variant of ['classic', 'grid']) {
    await page.evaluate(value => window.rangeEditorFixture.appearance(true, value), variant)
    await select('A1:C1'); await menu('sheet.menu.merge', '全部合并'); await expect(modal()).toBeVisible()
    const color = await modal().evaluate(element => getComputedStyle(element.querySelector('.ant-modal-container') ?? element.querySelector('.ant-modal-content') ?? element).backgroundColor)
    assert.notEqual(color, 'rgb(255, 255, 255)')
    await modal().press('Escape'); await expect(modal()).toBeHidden(); await expect(grid()).toBeFocused()
  }
  assert.equal((await history()).undo, 0)
})

test('standalone confirmation and clear keep explicit rich-text and style reset markers through the parent adapter', async () => {
  await page.goto(baseURL + '?standalone=1'); await expect(grid()).toBeVisible()
  await select('A1:C1'); await menu('sheet.menu.merge', '全部合并')
  await modal().getByRole('button', { name: '继续合并', exact: true }).click(); await expect(grid()).toBeFocused()
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '首项'); assert.equal((await snapshot()).sheets[0].cells['0:1'].richText, undefined)
  await select('A1'); await grid().press('Delete')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].richText, undefined)
  await contextClear('清除格式')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].style.font.bold, false)
  assert.deepEqual((await snapshot()).sheets[0].mergedRanges, [])
})
