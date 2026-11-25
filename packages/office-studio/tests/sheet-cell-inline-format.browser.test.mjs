import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Draft only: library-owned text/IME hosts and the real workbench; not run.
const route = '/__sheet-cell-inline-format__'
const fixture = readFileSync(new URL('./fixtures/sheet-rich-text-input.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-inline-format-fixture', configureServer(vite) {
      vite.middlewares.use(route, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(route, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + route
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const editor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const formula = () => page.getByRole('textbox', { name: '公式编辑', exact: true })
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const ribbon = id => page.locator('[data-ribbon-item="' + id + '"]:visible').first()
const button = id => ribbon(id).getByRole('button').first()
const snapshot = () => page.evaluate(() => window.richInputFixture.snapshot())
const history = () => page.evaluate(() => window.richInputFixture.history())
async function edit(address = 'A1') {
  await nameBox().fill(address); await nameBox().press('Enter'); await expect(grid()).toBeFocused()
  await grid().press('F2'); await expect(editor()).toBeFocused()
  return editor()
}
async function select(input, start, end = start) {
  await input.evaluate((element, value) => {
    element.focus(); element.setSelectionRange(value.start, value.end)
    element.dispatchEvent(new Event('select', { bubbles: true }))
  }, { start, end })
}
const selection = input => input.evaluate(element => ({ start: element.selectionStart, end: element.selectionEnd }))

test('Ribbon formats selected cell text, restores native offsets and commits one undoable transaction', async () => {
  const input = await edit(), before = await snapshot()
  await select(input, 2, 4); await button('sheet.command.toggle-bold').click()
  await expect(editor()).toBeFocused(); assert.deepEqual(await selection(editor()), { start: 2, end: 4 })
  assert.equal((await history()).undo, 0); assert.deepEqual(await snapshot(), before)
  await editor().press('Enter'); await expect(editor()).toHaveCount(0)
  const saved = await snapshot(), cell = saved.sheets[0].cells['0:0']
  assert.equal(cell.richText.find(run => run.id === 'two').font.bold, true)
  assert.equal(cell.richText.find(run => run.id === 'two').font.italic, true)
  assert.equal(cell.richText.find(run => run.id === 'three').hyperlink, 'https://example.test')
  assert.equal((await history()).undo, 1)
  await grid().press('Control+z'); assert.deepEqual((await snapshot()).sheets[0].cells['0:0'], before.sheets[0].cells['0:0'])
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).sheets[0].cells['0:0'], cell)
  await page.evaluate(value => window.richInputFixture.reopen(value), saved)
  await edit(); await expect(editor()).toHaveValue('ababab'); await editor().press('Escape')
})

test('formula-bar formatting and style undo/redo share the same draft and one final content history', async () => {
  await edit(); await formula().focus(); await select(formula(), 2, 4)
  await formula().press('Control+u'); assert.equal((await history()).undo, 0)
  await formula().press('Control+z'); await formula().press('Control+y')
  await expect(formula()).toBeFocused(); assert.deepEqual(await selection(formula()), { start: 2, end: 4 })
  await formula().press('Enter')
  const cell = (await snapshot()).sheets[0].cells['0:0']
  assert.equal(cell.richText.find(run => run.id === 'two').font.underline, true); assert.equal((await history()).undo, 1)
})

test('font-family and size popup inputs keep the cell session alive and apply to its stored text range', async () => {
  await edit(); await select(editor(), 1, 3)
  const family = ribbon('sheet.command.set-font-family').locator('input')
  await family.fill('Arial'); await family.press('Enter')
  await expect(editor()).toBeFocused(); assert.deepEqual(await selection(editor()), { start: 1, end: 3 })
  const size = ribbon('sheet.command.set-font-size').locator('input')
  await size.fill('18'); await size.press('Enter')
  await expect(editor()).toBeFocused(); assert.deepEqual(await selection(editor()), { start: 1, end: 3 })
  assert.equal((await history()).undo, 0)
  await editor().press('Enter')
  const runs = (await snapshot()).sheets[0].cells['0:0'].richText, styled = runs.filter(run => run.font?.family === 'Arial')
  assert.equal(styled.map(run => run.text).join(''), 'ba'); assert.ok(styled.every(run => run.font.size === 18))
})

test('Canvas pointer drag selects rich text through the library-owned input and cancellation leaves source content intact', async () => {
  await edit('E1'); const bounds = await editor().boundingBox(), before = await snapshot()
  await page.mouse.move(bounds.x + 3, bounds.y + 8); await page.mouse.down()
  await page.mouse.move(bounds.x + bounds.width - 5, bounds.y + 8, { steps: 6 }); await page.mouse.up()
  assert.deepEqual(await selection(editor()), { start: 0, end: 5 })
  await button('sheet.command.toggle-italic').click(); await expect(editor()).toBeFocused()
  await editor().press('Escape'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})

test('Home, End and vertical navigation use the rich font visual lines rather than native uniform-font wrapping', async () => {
  await edit('F1')
  await editor().press('Home'); assert.deepEqual(await selection(editor()), { start: 3, end: 3 })
  await editor().press('ArrowUp'); assert.deepEqual(await selection(editor()), { start: 0, end: 0 })
  await editor().press('End'); assert.deepEqual(await selection(editor()), { start: 2, end: 2 })
  await editor().press('ArrowDown'); const next = await selection(editor())
  assert.ok(next.start >= 3 && next.start <= 5); assert.equal(next.start, next.end)
  await editor().press('Escape'); assert.equal((await history()).undo, 0)
})

test('numeric-looking text can receive a partial font without changing General format or becoming a number', async () => {
  await edit('E2'); await select(editor(), 1, 2)
  await button('sheet.command.toggle-italic').click(); await expect(editor()).toBeFocused(); await editor().press('Enter')
  const cell = (await snapshot()).sheets[0].cells['1:4']
  assert.equal(cell.value, '123'); assert.equal(cell.dataType, 'text'); assert.equal(cell.style.numberFormat, 'General')
  assert.equal(cell.richText.filter(run => run.font?.italic).map(run => run.text).join(''), '2')
  assert.equal((await history()).undo, 1)
})

test('theme/grid changes keep the rich input bridge themed and stale controls cannot publish formatting to a replacement studio', async () => {
  await edit(); await select(editor(), 2, 4)
  await page.evaluate(() => window.richInputFixture.appearance(true, 'grid'))
  await expect(page.locator('.rich-text-editor-active')).toBeVisible()
  assert.equal(await editor().evaluate(element => getComputedStyle(element).color), 'rgba(0, 0, 0, 0)')
  await button('sheet.command.toggle-bold').click(); await expect(editor()).toBeFocused()
  await page.evaluate(() => window.richInputFixture.remember())
  await page.evaluate(() => window.richInputFixture.replaceHost('late input'))
  assert.equal((await history()).undo, 0)
  const cell = (await snapshot()).sheets[0].cells['0:0']
  assert.equal(cell.value, 'ababab'); assert.equal(cell.richText.find(run => run.id === 'two').font.bold, undefined)
})

test('a stationary captured pointer beyond the rich editor keeps scrolling and extending selection without new moves', async () => {
  await edit('G1'); await editor().press('Home')
  const bounds = await editor().boundingBox(), before = await snapshot()
  await page.mouse.move(bounds.x + 6, bounds.y + 10); await page.mouse.down()
  await page.mouse.move(bounds.x + bounds.width + 40, bounds.y + 10)
  const initial = await selection(editor())
  await expect.poll(async () => (await selection(editor())).end).toBeGreaterThan(initial.end + 2)
  const moving = await selection(editor()); assert.equal(moving.start, 0)
  await page.mouse.up()
  const released = await selection(editor())
  await page.evaluate(() => new Promise(resolve => { let remaining = 8; const tick = () => --remaining ? requestAnimationFrame(tick) : resolve(); requestAnimationFrame(tick) }))
  assert.deepEqual(await selection(editor()), released)
  await editor().press('Escape'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})

test('pointer cancellation stops rich drag scrolling and a replacement gesture can select again without publishing content', async () => {
  await edit('G1'); await editor().press('Home')
  const bounds = await editor().boundingBox(), before = await snapshot()
  await grid().evaluate(element => element.addEventListener('pointerdown', event => { element.dataset.testPointerId = String(event.pointerId) }, { capture: true, once: true }))
  await page.mouse.move(bounds.x + 6, bounds.y + 10); await page.mouse.down()
  await page.mouse.move(bounds.x + bounds.width + 40, bounds.y + 10)
  const initial = await selection(editor())
  await expect.poll(async () => (await selection(editor())).end).toBeGreaterThan(initial.end + 2)
  await grid().evaluate(element => element.dispatchEvent(new PointerEvent('pointercancel', { bubbles: true, pointerId: Number(element.dataset.testPointerId) })))
  const canceled = await selection(editor())
  await page.evaluate(() => new Promise(resolve => { let remaining = 8; const tick = () => --remaining ? requestAnimationFrame(tick) : resolve(); requestAnimationFrame(tick) }))
  assert.deepEqual(await selection(editor()), canceled)
  await page.mouse.up(); await editor().press('Home')
  await page.mouse.move(bounds.x + 6, bounds.y + 10); await page.mouse.down()
  await page.mouse.move(bounds.x + 60, bounds.y + 10); await page.mouse.up()
  assert.ok((await selection(editor())).end > 0)
  await editor().press('Escape'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})

test('End on a wrapped rich visual line retains that line through repeated End and Home without changing content', async () => {
  await edit('H1'); const before = await snapshot()
  await editor().press('Home'); await editor().press('ArrowUp'); await editor().press('Home')
  const start = await selection(editor())
  await editor().press('End'); const end = await selection(editor())
  assert.ok(end.end > start.start); assert.equal(end.start, end.end)
  await editor().press('End'); assert.deepEqual(await selection(editor()), end)
  await editor().press('Home'); assert.deepEqual(await selection(editor()), start)
  await editor().press('Escape'); assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})


