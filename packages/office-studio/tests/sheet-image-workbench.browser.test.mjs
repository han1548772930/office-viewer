import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred real ExcelWorkbench/Canvas/Ant contracts. Native Windows IME, both rendering backends and Office files remain unaccepted.
const fixturePath = '/__sheet-images-test__'
let server, browser, page, baseURL, errors
beforeAll(async () => {
  const fixture = await readFile(new URL('./fixtures/sheet-images.html', import.meta.url), 'utf8')
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-images-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => {
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
  page.on('pageerror', error => errors.push(error.message)); await page.goto(baseURL)
  await page.waitForFunction(() => !!window.sheetImageFixture)
  await expect(page.locator('.spreadsheet-canvas canvas').first()).toBeAttached()
  await page.waitForFunction(() => {
    const canvas = document.querySelector('.spreadsheet-canvas canvas'), surface = document.querySelector('.canvas-surface')?.getBoundingClientRect()
    return canvas && canvas.width > 1 && canvas.height > 1 && surface && surface.width > 300 && surface.height > 200
  })
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
const images = (sheetId = 'source') => page.evaluate(sheetId => window.sheetImageFixture.images(sheetId), sheetId)
const datum = async id => (await images()).find(image => image.id === id)
const history = () => page.evaluate(() => window.sheetImageFixture.history())
const editor = () => page.getByRole('dialog', { name: '图片属性', exact: true })
const near = (actual, expected, tolerance = 0.15) => assert.ok(Math.abs(actual - expected) < tolerance, `${actual} != ${expected}`)
async function clickImage(id = 'a', options = {}) {
  const point = await page.evaluate(id => window.sheetImageFixture.point(id), id); await page.mouse.click(point.x, point.y, options)
}
async function editImage() { await clickImage('a', { clickCount: 2 }); await expect(editor()).toBeVisible() }

test('double-click opens the Ant image Form and publishes flip/title edits into one history transaction', async () => {
  const before = await datum('a'); await editImage()
  await editor().getByRole('textbox', { name: '图片标题', exact: true }).fill('Canvas edited title')
  await editor().getByRole('checkbox', { name: '水平翻转', exact: true }).uncheck()
  await editor().getByRole('button', { name: '应用', exact: true }).click(); await expect(editor()).toBeHidden()
  const after = await datum('a'); assert.equal(after.title, 'Canvas edited title'); assert.equal(after.flipX, false)
  assert.deepEqual(after.crop, before.crop); assert.equal((await history()).undo, 1)
  for (const key of ['a', 'b', 'c', 'd']) near(after.drawingGeometry.matrix[key], before.drawingGeometry.matrix[key], 1e-7)
  await page.evaluate(() => window.sheetImageFixture.undo()); assert.equal((await datum('a')).title, before.title)
  await page.evaluate(() => window.sheetImageFixture.redo()); await page.evaluate(() => window.sheetImageFixture.reopen())
  assert.equal((await datum('a')).title, 'Canvas edited title'); assert.equal((await datum('a')).flipX, false)
})

test('affine Canvas resize hits the reflected handle and preserves the opposite world corner', async () => {
  await clickImage(); const before = await datum('a'), point = await page.evaluate(() => window.sheetImageFixture.point('a', 'e')),
    m = before.drawingGeometry.matrix, length = Math.hypot(m.a, m.b)
  await page.mouse.move(point.x, point.y); await page.mouse.down()
  await page.mouse.move(point.x + 16 * m.a / length, point.y + 16 * m.b / length, { steps: 4 }); await page.mouse.up()
  const after = await datum('a'); assert.equal((await history()).undo, 1)
  near(after.drawingGeometry.matrix.e, m.e); near(after.drawingGeometry.matrix.f, m.f)
  assert.deepEqual(after.metadata.opaqueAuthor, before.metadata.opaqueAuthor)
  assert.ok(after.size.width > before.size.width)
})

test('Ant Upload reads actual PNG bytes and inserts a new image once with undo and redo', async () => {
  const before = await images(); await page.evaluate(() => window.sheetImageFixture.openInsert())
  const picker = page.getByRole('dialog', { name: '插入浮动图片', exact: true }); await expect(picker).toBeVisible()
  const png = await page.evaluate(() => window.sheetImageFixture.png)
  await picker.locator('input[type=file]').setInputFiles({ name: 'actual.png', mimeType: 'image/png', buffer: Buffer.from(png.split(',')[1], 'base64') })
  await picker.getByRole('button', { name: '插入', exact: true }).click(); await expect(picker).toBeHidden()
  const inserted = (await images()).find(image => !before.some(old => old.id === image.id))
  assert.ok(inserted); assert.equal(inserted.resource.width, 1); assert.equal(inserted.resource.height, 1)
  assert.equal(inserted.resource.source, png); assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.sheetImageFixture.undo()); assert.equal((await images()).some(image => image.id === inserted.id), false)
  await page.evaluate(() => window.sheetImageFixture.redo()); assert.equal((await images()).some(image => image.id === inserted.id), true)
})

test('paste through the Canvas reader restores HTML affine media after the previous Studio broker was disposed', async () => {
  const before = await datum('a'); assert.equal(await page.evaluate(() => window.sheetImageFixture.copyAuthor()), true)
  await page.evaluate(() => window.sheetImageFixture.reopen()); await page.evaluate(() => window.sheetImageFixture.pasteTarget())
  await page.locator('.spreadsheet-canvas').focus(); await page.keyboard.press('Control+v')
  await page.waitForFunction(() => window.sheetImageFixture.images('destination').length === 1)
  const pasted = (await images('destination'))[0]; assert.notEqual(pasted.id, before.id)
  assert.deepEqual(pasted.crop, before.crop); assert.equal(pasted.flipX, before.flipX); assert.equal(pasted.flipY, before.flipY)
  for (const key of ['a', 'b', 'c', 'd']) near(pasted.drawingGeometry.matrix[key], before.drawingGeometry.matrix[key], 1e-7)
  assert.equal((await history()).undo, 1)
})

test('a live Ant dialog follows light/dark tokens and cancel restores Canvas focus without content history', async () => {
  await editImage(); const before = await images(), light = await editor().locator('.ant-modal-content').evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.sheetImageFixture.appearance(true, 'grid'))
  await expect(editor()).toBeVisible()
  await page.waitForFunction(previous => {
    const content = document.querySelector('.ant-modal-content')
    return content && getComputedStyle(content).backgroundColor !== previous
  }, light)
  const dark = await editor().locator('.ant-modal-content').evaluate(element => getComputedStyle(element).backgroundColor)
  assert.notEqual(dark, light)
  await editor().getByRole('button', { name: '取消', exact: true }).click(); await expect(editor()).toBeHidden()
  await expect(page.locator('.spreadsheet-canvas')).toBeFocused(); assert.deepEqual(await images(), before)
  assert.equal((await history()).undo, 0)
})

test('changing the active sheet cancels an image draft and its pending author update', async () => {
  await editImage(); const before = await images()
  await editor().getByRole('textbox', { name: '图片标题', exact: true }).fill('Unpublished draft')
  await page.evaluate(() => window.sheetImageFixture.switchSheet()); await expect(editor()).toBeHidden()
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
})

test('copy/cut inside an Ant text input stays in the text scope and does not capture or remove Canvas images', async () => {
  await editImage(); const before = await images(), input = editor().getByRole('textbox', { name: '图片标题', exact: true })
  await input.fill('中文输入桥'); await input.selectText(); await page.keyboard.press('Control+x')
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
  assert.equal((await page.evaluate(() => window.sheetImageFixture.board())).html, '')
  await editor().getByRole('button', { name: '取消', exact: true }).click()
})
