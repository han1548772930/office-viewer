import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred Canvas/Ant atom and physical-row browser drafts. No browser or server was started.
// Clipboard and composition events below are synthetic; native OS clipboard/IME still need separate acceptance.
const fixturePath = '/__document-image-inline-test__'
let server, browser, page, baseURL, errors
beforeAll(async () => {
  const fixture = await readFile(new URL('./fixtures/document-images.html', import.meta.url), 'utf8')
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-image-inline-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => {
      try { response.setHeader('Content-Type', 'text/html; charset=utf-8'); response.end(await vite.transformIndexHtml(fixturePath, fixture)) } catch (error) { next(error) }
    }) },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate a port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 960 }, deviceScaleFactor: 1.25 })
  page.on('pageerror', error => errors.push(error.message)); await load()
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
async function load(query = '?inline') {
  await page.goto(baseURL + query); await page.waitForFunction(() => !!window.documentImageFixture)
  await expect(input()).toBeAttached()
}
const input = () => page.getByRole('textbox', { name: '文档文字输入', exact: true })
const images = () => page.evaluate(() => window.documentImageFixture.images())
const location = () => page.evaluate(() => window.documentImageFixture.selection())
const history = () => page.evaluate(() => window.documentImageFixture.history())
const text = () => page.evaluate(() => window.documentImageFixture.snapshot().blocks.find(block => block.id === 'p').runs.map(run => run.text).join(''))
const format = () => page.getByRole('dialog', { name: '图片格式', exact: true })
async function clickImage(options = {}) {
  const point = await page.evaluate(() => window.documentImageFixture.point('inline-image'))
  await page.mouse.click(point.x, point.y, options)
  await expect(page.getByRole('toolbar', { name: '图片操作', exact: true })).toBeVisible()
}

test('two Ant flip switches edit a paragraph atom in one history entry and retain theme, focus and saved flips', async () => {
  await clickImage({ clickCount: 2 }); const panel = format()
  await expect(panel).toBeVisible(); await expect(panel.getByRole('textbox', { name: '替代文字', exact: true })).toBeFocused()
  const horizontal = panel.getByRole('switch', { name: '水平翻转', exact: true }), vertical = panel.getByRole('switch', { name: '垂直翻转', exact: true })
  await expect(horizontal).toHaveAttribute('aria-checked', 'false'); await expect(vertical).toHaveAttribute('aria-checked', 'false')
  await horizontal.click(); await vertical.click()
  const field = panel.getByRole('textbox', { name: '替代文字', exact: true }), light = await field.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.documentImageFixture.dark(true))
  await expect.poll(() => field.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  await panel.getByRole('button', { name: '应用', exact: true }).click(); await expect(panel).toBeHidden(); await expect(input()).toBeFocused()
  assert.deepEqual([(await images())[0].image.flipX, (await images())[0].image.flipY], [true, true]); assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.documentImageFixture.undo()); assert.equal((await images())[0].image.flipX ?? false, false)
  await page.evaluate(() => window.documentImageFixture.redo()); await page.evaluate(() => window.documentImageFixture.saveAndReopen())
  assert.equal((await images())[0].image.flipY, true); assert.equal(await text(), 'LEFT\uFFFCRIGHT')
})

test('clicking an inline image selects its source range and arrow navigation, typing and Backspace preserve adjacent text', async () => {
  await clickImage()
  await expect.poll(async () => { const state = await location(); return [state.selection.anchor.offset, state.selection.focus.offset] }).toEqual([4, 5])
  await input().press('ArrowRight'); await expect.poll(async () => (await location()).selection.focus.offset).toBe(5)
  await input().press('x'); await expect.poll(text).toBe('LEFT\uFFFCxRIGHT')
  assert.equal((await images())[0].imageId, 'inline-image')
  await input().press('ArrowLeft'); await input().press('Backspace'); await expect.poll(text).toBe('LEFTxRIGHT')
  assert.equal((await images()).length, 0)
  await page.evaluate(() => window.documentImageFixture.undo()); await expect.poll(text).toBe('LEFT\uFFFCxRIGHT')
  assert.equal((await images())[0].imageId, 'inline-image')
})

test('rich copy, cut and paste transfer a real IMG and give the pasted atom a new drawing identity', async () => {
  await page.evaluate(() => window.documentImageFixture.selectRange('p', 0, 10)); await input().focus()
  const copied = await page.evaluate(() => window.documentImageFixture.clipboard('copy'))
  assert.equal(copied.text, 'LEFTinline-imageRIGHT'); assert.match(copied.html, /<img /); assert.match(copied.html, /data-office-image="v1"/)
  await page.evaluate(() => window.documentImageFixture.clipboard('cut')); await expect.poll(text).toBe('')
  assert.equal((await images()).length, 0)
  await page.evaluate(data => window.documentImageFixture.clipboard('paste', { 'text/plain': data.text, 'text/html': data.html }), copied)
  await expect.poll(text).toBe('LEFT\uFFFCRIGHT')
  const pasted = (await images())[0]; assert.notEqual(pasted.imageId, 'inline-image'); assert.equal(pasted.image.alt, 'inline-image')
  assert.equal((await history()).undo, 2)
  await page.evaluate(() => window.documentImageFixture.undo()); await expect.poll(text).toBe('')
  await page.evaluate(() => window.documentImageFixture.undo()); await expect.poll(text).toBe('LEFT\uFFFCRIGHT')
  assert.equal((await images())[0].imageId, 'inline-image')
})

test('synthetic composition replaces the selected atom once and undo restores its source range and resource', async () => {
  await clickImage(); const before = await history()
  await input().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    element.value = '中'
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '中' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中' }))
  })
  await expect.poll(text).toBe('LEFT中RIGHT'); assert.equal((await images()).length, 0)
  assert.equal((await history()).undo, before.undo + 1)
  await page.evaluate(() => window.documentImageFixture.undo()); await expect.poll(text).toBe('LEFT\uFFFCRIGHT')
  assert.equal((await images())[0].imageId, 'inline-image')
  await expect(input()).toBeFocused()
})

test('both sides of a square wrap share a physical row and Home, End and Down follow the entire row', async () => {
  await load('?wrap')
  const rows = await page.evaluate(() => window.documentImageFixture.layoutRows().filter(line => line.target.blockId === 'p'))
  const first = rows.find(line => line.fragmentCount > 1), same = rows.filter(line => line.rowKey === first.rowKey)
  assert.equal(same.length, 2); assert.ok(same[0].x < same[1].x); assert.equal(same[0].baseline, same[1].baseline)
  const offset = same[1].start + 1
  await page.evaluate(offset => window.documentImageFixture.selectRange('p', offset, offset), offset); await input().focus()
  await input().press('Home'); await expect.poll(async () => (await location()).selection.focus.offset).toBe(same[0].start)
  await page.evaluate(offset => window.documentImageFixture.selectRange('p', offset, offset), offset); await input().focus()
  await input().press('End'); await expect.poll(async () => (await location()).selection.focus.offset).toBe(same.at(-1).end)
  await page.evaluate(offset => window.documentImageFixture.selectRange('p', offset, offset), offset); await input().focus()
  await input().press('ArrowDown')
  const destination = (await location()).selection.focus.offset, next = rows.find(line => line.y > first.y), nextRow = next.rowKey ? rows.filter(line => line.rowKey === next.rowKey) : [next]
  assert.ok(destination >= nextRow[0].start && destination <= nextRow.at(-1).end)
  assert.equal((await history()).undo, 0)
})
