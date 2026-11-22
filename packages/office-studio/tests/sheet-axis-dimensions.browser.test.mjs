import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-axis-dimensions-test__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>行列尺寸</title><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref, watch } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import { OfficeStudio } from '/packages/office-studio/src/api/office-studio'
import { FOffice } from '/packages/office-studio/src/api/facade'
import ExcelWorkbench from '/packages/office-studio/src/components/ExcelWorkbench.vue'
import { createBlankWorkbook, createCell } from '/packages/office-studio/src/domain/model'
import { applyExcelWorkbookView, toExcelWorkbookView } from '/packages/office-studio/src/services/excelViewAdapter'
import { sheetColumnWidth, sheetRowHeight, sheetRowIsAutoHeight } from '/packages/office-studio/src/domain/sheet-axis'
import { SpreadsheetGeometry } from '/packages/office-studio/src/rendering/spreadsheet/layout'
const copy = value => JSON.parse(JSON.stringify(value))
const original = createBlankWorkbook(); original.id = 'axis-fixture'; original.name = '行列尺寸'; original.calculationMode = 'manual'
original.sheets[0].defaultColumnWidth = 90; original.sheets[0].defaultRowHeight = 26
original.sheets[0].cells['1:0'] = createCell(1, 0, '第一行\\n第二行\\n第三行')
original.sheets[0].cells['1:0'].style.wrap = true
const studio = new OfficeStudio({ workbook: original }), api = FOffice.newAPI(studio)
const standalone = new URLSearchParams(location.search).has('standalone')
let base = copy(original)
const view = ref(toExcelWorkbookView(original)), dark = ref(false), variant = ref('classic'), notifications = [], emitted = []
watch(() => studio.editor.workbook.model, model => { if (!standalone) view.value = toExcelWorkbookView(model) }, { flush: 'sync' })
const current = () => standalone ? base : studio.editor.workbook.model
window.axisFixture = {
  snapshot: () => copy(current()), selection: () => copy(studio.getSelection()),
  history: () => { const value = studio.editor.history.units[studio.activeWorkbookId]; return { undo: value?.undo.length ?? 0, emits: emitted.length } },
  appearance: (value, layout) => { dark.value = value; variant.value = layout },
  height: row => sheetRowHeight(current().sheets[0], row), width: column => sheetColumnWidth(current().sheets[0], column),
  automatic: row => sheetRowIsAutoHeight(current().sheets[0], row),
  save: () => JSON.stringify(current()), reopen: value => api.createWorkbook({ ...JSON.parse(value), id: 'axis-reopened' }).save(),
  change: () => api.getActiveWorkbook().getActiveSheet().getRange('D3').setValue('new revision'),
  seedZero: () => {
    const model = copy(current()), sheet = model.sheets[0]
    sheet.rowHeightSpans = [{ start: 2, end: 5, value: 0 }]; sheet.rowAutoHeightSpans = [{ start: 2, end: 5, value: false }]
    sheet.columnWidthSpans = [{ start: 1, end: 3, value: 0 }]
    if (standalone) { base = model; view.value = toExcelWorkbookView(model) } else studio.editor.workbook.replace(model)
  },
  point: (axis, index, edge = false) => {
    const element = document.querySelector('.spreadsheet-canvas'), scroll = element.querySelector('.canvas-scroll')
    const bounds = element.getBoundingClientRect(), geometry = new SpreadsheetGeometry(view.value.sheets[0])
    const viewport = { width: bounds.width, height: bounds.height, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: view.value.sheets[0].zoom ?? 1 }
    const rect = geometry.cellRect(axis === 'row' ? index : 0, axis === 'column' ? index : 0, viewport)
    return axis === 'row' ? { x: bounds.left + 20, y: bounds.top + rect.y + (edge ? rect.height : rect.height / 2) }
      : { x: bounds.left + rect.x + (edge ? rect.width : rect.width / 2), y: bounds.top + 13 }
  },
}
createApp({ render: () => h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm } }, {
  default: () => h(ExcelWorkbench, { modelValue: view.value, studio: standalone ? undefined : studio,
    canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo, ribbonVariant: variant.value,
    'onUpdate:modelValue': value => { emitted.push(copy(value)); if (standalone) { base = applyExcelWorkbookView(base, value); view.value = value } else studio.editor.replaceWorkbookFromView(value, studio.regionalSettings.region.value) },
    onUndo: () => api.getActiveWorkbook().undo(), onRedo: () => api.getActiveWorkbook().redo(), onNotify: value => notifications.push(copy(value)),
  }),
}) }).mount('#app')
window.addEventListener('pagehide', () => studio.dispose(), { once: true })
</script></body></html>`

let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-axis-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_req, res, next) => {
      try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
    }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
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
const dialog = name => page.getByRole('dialog', { name, exact: true })
const height = row => page.evaluate(value => window.axisFixture.height(value), row)
const width = column => page.evaluate(value => window.axisFixture.width(value), column)
const history = () => page.evaluate(() => window.axisFixture.history())
async function select(address) {
  const input = page.getByRole('textbox', { name: '名称框', exact: true })
  await input.fill(address); await input.press('Enter'); await expect(grid()).toBeFocused()
}
async function axisMenu(axis, index, label) {
  const point = await page.evaluate(({ axis, index }) => window.axisFixture.point(axis, index), { axis, index })
  await page.mouse.click(point.x, point.y, { button: 'right' })
  await page.getByRole('menuitem', { name: label, exact: true }).click()
}

test('row height dialog preserves the header target, fractional size, shared undo and save/reopen', async () => {
  await select('D7'); await axisMenu('row', 2, '行高')
  await expect(dialog('行高')).toBeVisible(); const input = dialog('行高').getByRole('spinbutton')
  await expect(input).toBeFocused(); assert.equal((await history()).undo, 0)
  await input.fill('31.125'); await dialog('行高').getByRole('button', { name: '确定', exact: true }).click()
  await expect(dialog('行高')).toBeHidden(); await expect(grid()).toBeFocused()
  assert.equal(await height(2), 31.125); assert.equal(await height(6), 26); assert.equal((await history()).undo, 1)
  const saved = await page.evaluate(() => window.axisFixture.save())
  const reopened = await page.evaluate(value => window.axisFixture.reopen(value), saved)
  assert.equal(reopened.sheets[0].rowHeightSpans.find(span => span.start <= 2 && span.end >= 2).value, 31.125)
  await grid().press('Control+z'); assert.equal(await height(2), 26)
  await grid().press('Control+y'); assert.equal(await height(2), 31.125)
})

test('manual column widths above the previous UI limit remain editable in dark grid Ribbon', async () => {
  await page.evaluate(() => window.axisFixture.appearance(true, 'grid'))
  await axisMenu('column', 1, '列宽'); const input = dialog('列宽').getByRole('spinbutton')
  await expect(input).toBeFocused(); await input.fill('1050.125'); await input.press('Enter')
  await expect(dialog('列宽')).toBeHidden(); await expect(grid()).toBeFocused()
  assert.equal(await width(1), 1050.125); assert.equal(await width(0), 90)
})

test('composition Enter does not confirm a dimension until composition ends', async () => {
  await axisMenu('row', 3, '行高'); const input = dialog('行高').getByRole('spinbutton')
  await input.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '三' })))
  await input.press('Enter'); await expect(dialog('行高')).toBeVisible(); assert.equal((await history()).undo, 0)
  await input.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '' })))
  await input.fill('35.75'); await input.press('Enter'); await expect(dialog('行高')).toBeHidden()
  assert.equal(await height(3), 35.75)
})

test('cancel and stale model revisions preserve the dialog or Canvas focus without applying dimensions', async () => {
  await axisMenu('row', 2, '行高'); await dialog('行高').getByRole('button', { name: '取消', exact: true }).click()
  await expect(dialog('行高')).toBeHidden(); await expect(grid()).toBeFocused(); assert.equal((await history()).undo, 0)
  await axisMenu('row', 2, '行高'); await dialog('行高').getByRole('spinbutton').fill('45')
  await page.evaluate(() => window.axisFixture.change())
  await dialog('行高').getByRole('button', { name: '确定', exact: true }).click()
  await expect(dialog('行高')).toBeVisible(); await expect(dialog('行高')).toContainText('已更改')
  assert.equal(await height(2), 26); assert.equal((await history()).undo, 1)
})

test('double-clicking a row boundary applies natural multiline height and later forced sizes stay manual', async () => {
  const point = await page.evaluate(() => window.axisFixture.point('row', 1, true))
  await page.mouse.dblclick(point.x, point.y)
  assert.ok(await height(1) > 26); assert.equal(await page.evaluate(() => window.axisFixture.automatic(1)), true)
  await axisMenu('row', 1, '行高'); await dialog('行高').getByRole('spinbutton').fill('30.5')
  await dialog('行高').getByRole('button', { name: '确定', exact: true }).click()
  assert.equal(await height(1), 30.5); assert.equal(await page.evaluate(() => window.axisFixture.automatic(1)), false)
})

test('zero row/column spans participate in keyboard navigation and name-box jumps', async () => {
  await page.evaluate(() => window.axisFixture.seedZero())
  await select('A2'); await grid().press('Tab')
  assert.equal((await page.evaluate(() => window.axisFixture.selection())).startColumn, 4)
  await grid().press('Enter')
  assert.equal((await page.evaluate(() => window.axisFixture.selection())).startRow, 6)
  await select('B3'); const selection = await page.evaluate(() => window.axisFixture.selection())
  assert.equal(selection.startRow, 6); assert.equal(selection.startColumn, 4)
})

test('lost pointer capture cancels a resize preview without committing a history entry', async () => {
  const point = await page.evaluate(() => window.axisFixture.point('row', 2, true))
  await page.mouse.move(point.x, point.y); await page.mouse.down(); await page.mouse.move(point.x, point.y + 40)
  await grid().evaluate(element => element.dispatchEvent(new PointerEvent('lostpointercapture', { bubbles: true, pointerId: 1 })))
  await page.mouse.up(); assert.equal(await height(2), 26); assert.equal((await history()).undo, 0)
})

test('standalone dimensions use the same sparse model and one emitted workbook update', async () => {
  await page.goto(baseURL + '?standalone'); await expect(grid()).toBeVisible()
  await axisMenu('row', 4, '行高'); await dialog('行高').getByRole('spinbutton').fill('0')
  await dialog('行高').getByRole('button', { name: '确定', exact: true }).click()
  await expect(dialog('行高')).toBeHidden(); await expect(grid()).toBeFocused()
  assert.equal(await height(4), 0); assert.equal((await history()).emits, 1)
})
