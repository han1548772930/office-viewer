import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// Draft only: real ExcelWorkbench/Ant/ECharts/Canvas2D/Facade/clipboard/PNG downloads.
// No Vite process, browser, type check, build, worker or benchmark ran in this batch.
let server, browser, context, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No HTTP port was allocated')
  baseURL = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/chart-object-actions.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ acceptDownloads: true, viewport: { width: 1440, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] })
  page = await context.newPage(); page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await page.waitForFunction(() => window.chartObjectFixture?.state().charts.length === 2)
  await page.waitForFunction(() => document.querySelector('.chart-hitbox[data-chart-id="native-object"]')?.getAttribute('data-rendered') === 'true')
}, 60_000)
afterEach(async () => { try { await page?.evaluate(() => window.chartObjectFixture?.dispose()); assert.deepEqual(errors, []) } finally { await context?.close() } })
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.chartObjectFixture.state())
const hitbox = (id = 'native-object') => page.locator(`.chart-hitbox[data-chart-id="${id}"]`)
const menu = () => page.getByRole('menu').filter({ has: page.getByRole('menuitem', { name: /复制图表/ }) }).last()
const drawer = () => page.locator('.ant-drawer').filter({ hasText: '图表设置' })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
async function select(id = 'native-object') { await hitbox(id).click({ position: { x: 8, y: 8 } }) }
async function openMenu(id = 'native-object', keyboard = false) {
  if (keyboard) { await hitbox(id).focus(); await hitbox(id).press('Shift+F10') }
  else await hitbox(id).click({ button: 'right', position: { x: 8, y: 8 } })
  await expect(menu()).toBeVisible()
}
async function action(label) { await menu().getByRole('menuitem', { name: new RegExp(label) }).click() }
async function waitCharts(count) { await page.waitForFunction(count => window.chartObjectFixture.state().charts.length === count, count); return state() }

test('single click selects and the fixed edit command opens the real Drawer; an authored title uses one history entry and reopens', async () => {
  const before = await state(); await select(); await expect(drawer()).toHaveCount(0)
  await openMenu('native-object', true)
  const labels = await menu().getByRole('menuitem').allTextContents()
  assert.deepEqual(labels.map(text => text.replace(/F2|Ctrl\+C|Delete/g, '').trim()), ['编辑图表', '复制图表', '导出图片', '删除图表'])
  await page.keyboard.press('Escape'); await expect(hitbox()).toBeFocused()
  await hitbox().press('F2'); await expect(drawer()).toBeVisible()
  const title = drawer().getByRole('textbox', { name: '图表标题', exact: true })
  await title.fill('Edited object'); await drawer().getByRole('button', { name: '应用', exact: true }).click()
  await page.waitForFunction(() => window.chartObjectFixture.state().charts[0].title === 'Edited object')
  const authored = await state(); assert.equal(authored.history.undo, before.history.undo + 1)
  await page.evaluate(() => window.chartObjectFixture.undo()); assert.equal((await state()).charts[0].title, before.charts[0].title)
  await page.evaluate(() => window.chartObjectFixture.redo()); await page.evaluate(() => window.chartObjectFixture.reopen())
  assert.equal((await state()).charts[0].title, 'Edited object')
})

test('copy through the context root and Ctrl+V create one new ID with live source/style bindings and one content history entry', async () => {
  const before = await state(); await openMenu(); await action('复制图表')
  await page.waitForFunction(() => !!window.chartObjectFixture.state().clipboardId)
  const copied = await state(); assert.equal(copied.charts.length, 2); assert.equal(copied.history.undo, before.history.undo); assert.equal(copied.systemWritten, true)
  await hitbox().focus(); await hitbox().press('Control+v')
  const pasted = await waitCharts(3), added = pasted.charts.find(chart => !before.charts.some(original => original.id === chart.id))
  assert.ok(added); assert.equal(added.source.sheetId, 'source'); assert.equal(added.title, before.charts[0].title)
  assert.deepEqual(added.options, before.charts[0].options); assert.equal(pasted.history.undo, before.history.undo + 1)
  await page.evaluate(() => window.chartObjectFixture.undo()); assert.equal((await state()).charts.length, 2)
  await page.evaluate(() => window.chartObjectFixture.redo()); await page.evaluate(() => window.chartObjectFixture.reopen())
  assert.equal((await state()).charts.find(chart => chart.id === added.id).title, added.title)
})

test('shared-scene export downloads the actual encoded PNG and releases the owned bitmap', async () => {
  await page.waitForFunction(() => document.querySelector('.chart-hitbox[data-chart-id="shared-object"]')?.getAttribute('data-rendered') === 'true')
  const before = await state(); await openMenu('shared-object')
  const downloadPromise = page.waitForEvent('download'); await action('导出图片'); const download = await downloadPromise
  assert.equal(download.suggestedFilename(), 'shared-object.png')
  const path = await download.path(); assert.ok(path); const bytes = await readFile(path)
  assert.deepEqual([...bytes.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10])
  const decoded = await page.evaluate(async bytes => {
    const expected = window.chartObjectFixture.encoding(), bitmap = await createImageBitmap(new Blob([new Uint8Array(bytes)], { type: 'image/png' }))
    const canvas = document.createElement('canvas'); canvas.width = bitmap.width; canvas.height = bitmap.height
    const context = canvas.getContext('2d'); context.drawImage(bitmap, 0, 0); const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
    const equal = pixels.length === expected.pixels.length && pixels.every((value, index) => value === expected.pixels[index])
    bitmap.close(); canvas.width = canvas.height = 0
    return { equal, width: expected.width, height: expected.height }
  }, [...bytes])
  assert.equal(decoded.equal, true); assert.ok(decoded.width > 0 && decoded.height > 0)
  const after = await state(); assert.deepEqual(after.charts, before.charts); assert.equal(after.history.undo, before.history.undo)
  assert.deepEqual(after.owned, [[0, 0]]); assert.equal(after.downloads[0].type, 'image/png')
})

test('focused deletion removes only the captured chart and follows the same undo/redo/save path as keyboard deletion', async () => {
  const before = await state(); await openMenu(); await action('删除图表')
  const removed = await waitCharts(1); assert.equal(removed.charts[0].id, 'shared-object'); assert.equal(removed.history.undo, before.history.undo + 1)
  await page.evaluate(() => window.chartObjectFixture.undo()); await waitCharts(2)
  await hitbox().focus(); await hitbox().press('Delete'); await waitCharts(1)
  await page.evaluate(() => window.chartObjectFixture.undo()); await waitCharts(2)
  await page.evaluate(() => window.chartObjectFixture.redo()); await page.evaluate(() => window.chartObjectFixture.reopen())
  assert.deepEqual((await state()).charts.map(chart => chart.id), ['shared-object'])
})

test('read-only and primary states apply to Ant menu items; popup theme and Escape keep content and focus intact', async () => {
  await page.evaluate(() => window.chartObjectFixture.mount({ readonly: true })); const before = await state()
  await openMenu(); assert.equal(await menu().getByRole('menuitem', { name: /编辑图表/ }).getAttribute('aria-disabled'), 'true')
  assert.equal(await menu().getByRole('menuitem', { name: /删除图表/ }).getAttribute('aria-disabled'), 'true')
  assert.notEqual(await menu().getByRole('menuitem', { name: /复制图表/ }).getAttribute('aria-disabled'), 'true')
  const light = await menu().evaluate(node => getComputedStyle(node).backgroundColor)
  await page.keyboard.press('Escape'); await page.evaluate(() => window.chartObjectFixture.appearance(true, 'grid'))
  await openMenu(); const dark = await menu().evaluate(node => getComputedStyle(node).backgroundColor); assert.notEqual(dark, light)
  await page.keyboard.press('Escape'); await expect(hitbox()).toBeFocused(); assert.deepEqual((await state()).charts, before.charts)
  await page.evaluate(() => window.chartObjectFixture.mount({ primary: true })); await openMenu()
  await expect(menu().getByRole('menuitem', { name: /删除图表/ })).toHaveCount(0)
  assert.equal(await menu().getByRole('menuitem').count(), 3)
})

test('late clipboard/image callbacks and replaced host sessions cannot publish, download or steal the new focus', async () => {
  await page.evaluate(() => window.chartObjectFixture.holdCopy()); await openMenu(); await action('复制图表')
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingCopies === 1)
  await page.evaluate(() => window.chartObjectFixture.switchSheet('destination')); await page.evaluate(() => window.chartObjectFixture.finishCopies())
  assert.equal((await state()).clipboardId, undefined)
  await page.evaluate(() => window.chartObjectFixture.switchSheet('source'))
  await page.waitForFunction(() => document.querySelector('.chart-hitbox[data-chart-id="native-object"]')?.getAttribute('data-rendered') === 'true')
  await page.evaluate(() => window.chartObjectFixture.holdCopy()); await openMenu(); await action('复制图表')
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingCopies === 1)
  await page.evaluate(() => window.chartObjectFixture.reopen()); await grid().focus()
  await page.evaluate(() => window.chartObjectFixture.finishCopies())
  await expect(grid()).toBeFocused(); assert.equal((await state()).clipboardId, undefined)
  await page.waitForFunction(() => document.querySelector('.chart-hitbox[data-chart-id="native-object"]')?.getAttribute('data-rendered') === 'true')
  await page.evaluate(() => window.chartObjectFixture.holdExport()); await openMenu(); await action('导出图片')
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingExports === 1)
  await page.evaluate(() => window.chartObjectFixture.switchSheet('destination')); await page.evaluate(() => window.chartObjectFixture.finishExports())
  const after = await state(); assert.equal(after.downloads.length, 0); assert.ok(after.owned.every(size => size[0] === 0 && size[1] === 0)); assert.equal(after.charts.length, 2)
})

test('grid paste recognizes real chart provenance; text input and IME keep their own clipboard/keyboard behavior', async () => {
  await openMenu(); await action('复制图表'); await page.waitForFunction(() => !!window.chartObjectFixture.state().clipboardId)
  await page.evaluate(() => window.chartObjectFixture.switchSheet('destination')); await grid().focus()
  assert.equal(await page.evaluate(() => window.chartObjectFixture.nativeGridPaste()), true)
  const pasted = await waitCharts(3); assert.equal(pasted.charts.find(chart => chart.sheetId === 'destination').source.sheetId, 'source')
  const nameBox = page.getByRole('textbox', { name: '名称框', exact: true }); await nameBox.fill('A10'); await nameBox.press('Enter'); await grid().press('F2')
  const editor = page.locator('.canvas-cell-input input, .canvas-cell-input textarea'); await expect(editor).toBeFocused(); await editor.fill('输入法文本')
  await editor.evaluate(node => node.dispatchEvent(new KeyboardEvent('keydown', { key: 'v', ctrlKey: true, isComposing: true, bubbles: true, cancelable: true })))
  await expect(editor).toHaveValue('输入法文本'); assert.equal((await state()).charts.length, 3)
  await editor.press('Escape')
})

test('stale chart tokens cannot become cell contents, while a later external text clipboard still follows the grid planner', async () => {
  await openMenu(); await action('复制图表'); await page.waitForFunction(() => !!window.chartObjectFixture.state().clipboardId)
  const before = await state(); await grid().focus()
  assert.equal(await page.evaluate(() => window.chartObjectFixture.nativeGridPaste({ plain: '[office-chart:foreign]', html: '<div data-office-chart-clipboard="foreign"></div>' })), true)
  const rejected = await state(); assert.deepEqual(rejected.charts, before.charts); assert.equal(rejected.history.undo, before.history.undo)
  await expect(page.locator('.chart-render-error')).toContainText('剪贴板来源已失效')
  const nameBox = page.getByRole('textbox', { name: '名称框', exact: true }); await nameBox.fill('D10'); await nameBox.press('Enter')
  await page.evaluate(() => window.chartObjectFixture.nativeGridPaste({ plain: 'external cell', html: '' }))
  await page.waitForFunction(() => window.chartObjectFixture.state().model.sheets.find(sheet => sheet.id === 'source').cells['9:3']?.value === 'external cell')
  assert.deepEqual((await state()).charts.map(chart => chart.id), before.charts.map(chart => chart.id))
})
