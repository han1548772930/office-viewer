import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Drafts only. No dev server, browser, build/type checks or benchmarks were started.
const route = '/__sheet-range-theme-structure__', fixture = readFileSync(new URL('./fixtures/sheet-range-theme.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-range-theme-structure-fixture', configureServer(vite) {
      vite.middlewares.use(route, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(route, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${route}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1600, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
  await page.evaluate(() => window.rangeThemeFixture.structure())
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const snapshot = () => page.evaluate(() => window.rangeThemeFixture.snapshot())
const history = () => page.evaluate(() => window.rangeThemeFixture.history())
const effective = address => page.evaluate(value => window.rangeThemeFixture.effective(value), address)
async function select(address) { await nameBox().fill(address); await nameBox().press('Enter'); await expect(grid()).toBeFocused() }
async function clipboard(kind, content = {}) {
  return grid().evaluate((element, input) => {
    const data = new DataTransfer()
    if (input.content.text !== undefined) data.setData('text/plain', input.content.text)
    if (input.content.html !== undefined) data.setData('text/html', input.content.html)
    const event = new ClipboardEvent(input.kind, { bubbles: true, cancelable: true, clipboardData: data })
    element.dispatchEvent(event)
    return { prevented: event.defaultPrevented, text: data.getData('text/plain'), html: data.getData('text/html') }
  }, { kind, content })
}
async function cut(from, to, other = false) {
  await select(from); const content = await clipboard('cut')
  assert.equal(content.prevented, true)
  if (other) await page.getByRole('tab', { name: 'Other', exact: true }).click()
  await select(to); await clipboard('paste', content)
}

test('real context-menu down insertion preserves the untouched footer and the shifted logical header', async () => {
  await select('B3:C4'); const before = await snapshot(), depth = await history()
  await grid().click({ button: 'right', position: { x: 240, y: 110 } })
  await page.getByRole('menuitem', { name: '插入', exact: true }).hover()
  await page.getByRole('menuitem', { name: '下移单元格', exact: true }).click()
  const moved = await snapshot()
  assert.equal(moved.sheets[0].rangeThemeRules.length, 3)
  assert.equal(moved.sheets[0].cells['4:1'].value, '移动内容')
  assert.equal((await effective('B1')).bg.rgb, '#123456'); assert.equal((await effective('B8')).bg.rgb, '#444444')
  assert.equal((await effective('A6')).bg.rgb, '#444444'); assert.equal((await history()).undo, depth.undo + 1)
  await grid().press('Control+z'); assert.deepEqual((await snapshot()).sheets, before.sheets)
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).sheets, moved.sheets)
})

test('Canvas cut and paste transfer sparse rules and author sources in one content history record', async () => {
  const before = await snapshot(), depth = await history(); await cut('B3:C4', 'F11')
  const moved = await snapshot(), target = moved.sheets[0].cells['10:5']
  assert.equal(target.value, '移动内容'); assert.equal(target.authoredStyle.font.italic, true)
  assert.deepEqual(Object.keys(moved.sheets[0].cells).sort(), ['0:0', '10:5'])
  assert.equal((await effective('F11')).bg.rgb, '#fedcba'); assert.equal((await effective('F12')).bg.rgb, '#abcdef')
  assert.equal((await effective('B3')).bg, undefined); assert.equal((await effective('A3')).bg.rgb, '#fedcba')
  assert.equal((await history()).undo, depth.undo + 1)
  await grid().press('Control+z'); assert.deepEqual((await snapshot()).sheets, before.sheets)
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).sheets, moved.sheets)
})

test('a partial cross-sheet cut to A1 keeps a negative logical origin and footer after JSON reopening', async () => {
  await cut('B5:C6', 'A1', true)
  const moved = await snapshot(), target = moved.sheets[1]
  assert.deepEqual(target.rangeThemeRules[0].styleRange, { startRow: -4, endRow: 1, startCol: -1, endCol: 2 })
  assert.equal(Object.keys(target.cells).length, 0)
  assert.equal((await effective('A1')).bg.rgb, '#fedcba'); assert.equal((await effective('A2')).bg.rgb, '#444444')
  await page.evaluate(() => window.rangeThemeFixture.reopen())
  assert.deepEqual((await snapshot()).sheets[1].rangeThemeRules, target.rangeThemeRules)
  assert.equal((await effective('A2')).bg.rgb, '#444444')
})

test('same-sheet overlapping cuts read the source once and never turn a fragment into another header', async () => {
  await cut('B3:C4', 'B4')
  const moved = await snapshot()
  assert.equal(moved.sheets[0].cells['3:1'].value, '移动内容'); assert.equal(moved.sheets[0].cells['2:1'], undefined)
  assert.equal((await effective('B4')).bg.rgb, '#fedcba'); assert.equal((await effective('B5')).bg.rgb, '#abcdef')
  assert.equal((await effective('B3')).bg, undefined); assert.equal((await effective('A1')).bg.rgb, '#123456')
})

test('a protected destination rejects paste atomically while a protected retained tail permits the cut', async () => {
  await page.evaluate(() => { window.rangeThemeFixture.protectTail(); window.rangeThemeFixture.protectMoveTarget() })
  const before = await snapshot(), depth = await history()
  await cut('B3:C4', 'F11')
  assert.deepEqual((await snapshot()).sheets, before.sheets); assert.deepEqual(await history(), depth)
  assert.ok((await page.evaluate(() => window.rangeThemeFixture.notifications())).some(item => /保护/.test(item.message)))
})

test('a moved rule remains dynamic and Canvas Ant input retains its direct override after theme registration', async () => {
  await page.evaluate(() => window.rangeThemeFixture.appearance(true, 'grid')); await cut('B3:C4', 'F11')
  await page.evaluate(() => window.rangeThemeFixture.updateStructureTheme())
  assert.equal((await effective('F11')).bg.rgb, '#778899'); assert.equal((await effective('F11')).it, 1)
  await select('F11'); await grid().press('F2')
  const input = page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
  await expect(input).toBeFocused(); await expect(input).toHaveValue('移动内容')
  assert.match(await input.evaluate(element => element.className), /ant-/)
  await input.press('Escape'); await expect(grid()).toBeFocused()
  assert.equal((await snapshot()).sheets[0].cells['10:5'].value, '移动内容')
})
