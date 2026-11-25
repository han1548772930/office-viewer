import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred real-Canvas/antdv-next drafts. Synthetic composition checks guards;
// Windows native IME, both renderer backends and real Office files remain unaccepted.
const fixturePath = '/__document-images-test__'
let server, browser, page, baseURL, errors
beforeAll(async () => {
  const fixture = await readFile(new URL('./fixtures/document-images.html', import.meta.url), 'utf8')
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-images-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => {
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
async function load(query = '') {
  await page.goto(baseURL + query); await page.waitForFunction(() => !!window.documentImageFixture)
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeAttached()
}
const images = host => page.evaluate(host => window.documentImageFixture.images(host), host)
const history = () => page.evaluate(() => window.documentImageFixture.history())
const format = () => page.getByRole('dialog', { name: '图片格式', exact: true })
const crop = () => page.getByRole('dialog', { name: '裁剪图片', exact: true })
const insert = () => page.getByRole('dialog', { name: '插入图片', exact: true })
const apply = panel => panel.getByRole('button', { name: '应用', exact: true })
const datum = async id => (await images()).find(item => item.imageId === id).image
async function clickImage(id = 'image', options = {}, host) {
  const point = await page.evaluate(({ id, host }) => window.documentImageFixture.point(id, undefined, host), { id, host })
  await page.mouse.click(point.x, point.y, options)
}
async function editImage(id = 'image', host) { await clickImage(id, { clickCount: 2 }, host); await expect(format()).toBeVisible() }
async function beginInsert() {
  await page.evaluate(() => window.documentImageFixture.select('p', 1))
  assert.equal(await page.evaluate(() => window.documentImageFixture.openInsert()), true)
  await expect(insert()).toBeVisible(); await expect(insert().getByRole('textbox', { name: '图片地址', exact: true })).toBeFocused()
}
const near = (actual, expected, tolerance = 0.1) => assert.ok(Math.abs(actual - expected) < tolerance, `${actual} != ${expected}`)

test('Ant upload decodes source bytes and inserts one selected Canvas object with one undo and stable public identity', async () => {
  const before = await images(); await beginInsert()
  const png = await page.evaluate(() => window.documentImageFixture.png)
  await insert().locator('input[type=file]').setInputFiles({ name: 'upload.png', mimeType: 'image/png', buffer: Buffer.from(png.split(',')[1], 'base64') })
  await insert().getByRole('textbox', { name: '替代文字', exact: true }).fill('Uploaded image')
  await insert().getByRole('button', { name: '确定', exact: true }).click(); await expect(insert()).toBeHidden()
  const next = (await images()).find(item => !before.some(old => old.imageId === item.imageId))
  assert.ok(next); assert.equal(next.image.alt, 'Uploaded image'); assert.equal(next.image.mimeType, 'image/png')
  assert.deepEqual([next.image.intrinsicWidth, next.image.intrinsicHeight], [120, 80]); assert.equal((await history()).undo, 1)
  await expect(page.getByRole('toolbar', { name: '图片操作', exact: true })).toBeVisible()
  await page.evaluate(() => window.documentImageFixture.undo()); assert.equal((await images()).some(item => item.imageId === next.imageId), false)
  await page.evaluate(() => window.documentImageFixture.redo()); await page.evaluate(() => window.documentImageFixture.saveAndReopen())
  assert.equal((await datum(next.imageId)).alt, 'Uploaded image')
})

test('remote URL bytes are embedded and selection changes cancel an unfinished read without inserting content', async () => {
  const png = await page.evaluate(() => window.documentImageFixture.png), buffer = Buffer.from(png.split(',')[1], 'base64')
  await page.route('https://example.test/image.png', route => route.fulfill({ status: 200, body: buffer,
    headers: { 'content-type': 'image/png', 'access-control-allow-origin': '*' } }))
  const before = await images(); await beginInsert()
  await insert().getByRole('textbox', { name: '图片地址', exact: true }).fill('https://example.test/image.png')
  await insert().getByRole('button', { name: '确定', exact: true }).click(); await expect(insert()).toBeHidden()
  const next = (await images()).find(item => !before.some(old => old.imageId === item.imageId))
  assert.equal(next.image.originalSource, 'https://example.test/image.png'); assert.ok(next.image.src.startsWith('data:image/png;base64,'))
  let pending, received
  const arrived = new Promise(resolve => received = resolve)
  await page.route('https://example.test/slow.png', route => { pending = route; received() })
  await beginInsert(); const beforeCancel = await images(), oldHistory = await history()
  await insert().getByRole('textbox', { name: '图片地址', exact: true }).fill('https://example.test/slow.png')
  await insert().getByRole('button', { name: '确定', exact: true }).click(); await arrived
  await page.evaluate(() => window.documentImageFixture.select('tail', 1))
  await expect(insert().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await pending.fulfill({ status: 200, body: buffer, headers: { 'content-type': 'image/png', 'access-control-allow-origin': '*' } }).catch(() => {})
  assert.deepEqual(await images(), beforeCancel); assert.deepEqual(await history(), oldHistory)
  await insert().getByRole('button', { name: '取消', exact: true }).click(); await expect(insert()).toBeHidden()
})

test('double-click format editing applies ratio, rotation and fractional crop in one transaction and survives save/reopen', async () => {
  await editImage(); const panel = format()
  await panel.getByRole('spinbutton', { name: '宽度（px）', exact: true }).fill('240')
  await expect(panel.getByRole('spinbutton', { name: '高度（px）', exact: true })).toHaveValue('160')
  await panel.getByRole('spinbutton', { name: '旋转（度）', exact: true }).fill('30')
  await panel.getByRole('spinbutton', { name: '左侧裁剪（%）', exact: true }).fill('25')
  await panel.getByRole('textbox', { name: '替代文字', exact: true }).fill('Formatted')
  await apply(panel).click(); await expect(panel).toBeHidden()
  const next = await datum('image'); assert.deepEqual([next.width, next.height, next.angle, next.crop.left, next.alt], [240, 160, 30, 0.25, 'Formatted'])
  assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.documentImageFixture.undo()); assert.equal((await datum('image')).width, 180)
  await page.evaluate(() => window.documentImageFixture.redo()); await page.evaluate(() => window.documentImageFixture.saveAndReopen())
  assert.equal((await datum('image')).crop.left, 0.25); assert.equal((await datum('image')).angle, 30)
})

test('Canvas drag and rotated resize publish once on pointer-up and keep the opposite rotated corner fixed', async () => {
  await clickImage(); const initial = await datum('image'), oldHistory = await history()
  let point = await page.evaluate(() => window.documentImageFixture.point('image'))
  await page.mouse.move(point.x, point.y); await page.mouse.down(); await page.mouse.move(point.x + 40, point.y + 25, { steps: 6 })
  assert.deepEqual(await datum('image'), initial); assert.deepEqual(await history(), oldHistory)
  await page.mouse.up(); const dragged = await datum('image')
  near(dragged.positionH.posOffset, initial.positionH.posOffset + 40); near(dragged.positionV.posOffset, initial.positionV.posOffset + 25)
  assert.equal((await history()).undo, oldHistory.undo + 1)
  await editImage(); await format().getByRole('spinbutton', { name: '旋转（度）', exact: true }).fill('30'); await apply(format()).click(); await expect(format()).toBeHidden()
  await clickImage(); const opposite = await page.evaluate(() => window.documentImageFixture.point('image', 'nw'))
  point = await page.evaluate(() => window.documentImageFixture.point('image', 'se')); const resizedHistory = await history()
  await page.mouse.move(point.x, point.y); await page.mouse.down(); await page.mouse.move(point.x + 35, point.y + 25, { steps: 5 }); await page.mouse.up()
  const after = await page.evaluate(() => window.documentImageFixture.point('image', 'nw'))
  near(after.x, opposite.x, 1); near(after.y, opposite.y, 1); assert.equal((await history()).undo, resizedHistory.undo + 1)
  const resized = await datum('image'); near(resized.width / resized.height, initial.width / initial.height, 0.001)
})

test('Escape cancels an active drag preview and right-click crop/reset/delete operate on the selected object', async () => {
  await clickImage(); const before = await datum('image'), oldHistory = await history()
  const point = await page.evaluate(() => window.documentImageFixture.point('image'))
  await page.mouse.move(point.x, point.y); await page.mouse.down(); await page.mouse.move(point.x + 50, point.y + 30)
  await page.keyboard.press('Escape'); await page.mouse.up()
  assert.deepEqual(await datum('image'), before); assert.deepEqual(await history(), oldHistory)
  await clickImage('image', { button: 'right' }); await page.getByRole('menuitem', { name: '裁剪图片', exact: true }).click(); await expect(crop()).toBeVisible()
  await crop().getByRole('spinbutton', { name: '右侧裁剪（%）', exact: true }).fill('20'); await apply(crop()).click(); await expect(crop()).toBeHidden()
  assert.equal((await datum('image')).crop.right, 0.2)
  await clickImage('image', { button: 'right' }); await page.getByRole('menuitem', { name: '重置尺寸与裁剪', exact: true }).click()
  assert.deepEqual([(await datum('image')).width, (await datum('image')).height, (await datum('image')).angle], [120, 80, 0])
  await clickImage('image', { button: 'right' }); await page.getByRole('menuitem', { name: '删除图片', exact: true }).click()
  assert.equal((await images()).some(item => item.imageId === 'image'), false)
  await page.evaluate(() => window.documentImageFixture.undo()); assert.equal((await datum('image')).width, 120)
})

test('text and numeric Ant fields suppress submission and Escape during composition, adapt theme and restore Canvas focus', async () => {
  await editImage(); const panel = format(), text = panel.getByRole('textbox', { name: '替代文字', exact: true })
  await expect(text).toBeFocused(); await text.fill('中文图片')
  await text.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await text.press('Control+Enter'); await text.press('Escape'); await expect(panel).toBeVisible(); await expect(apply(panel)).toBeDisabled()
  assert.equal((await history()).undo, 0)
  await text.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文图片' })))
  const number = panel.getByRole('spinbutton', { name: '宽度（px）', exact: true })
  await number.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))); await expect(apply(panel)).toBeDisabled()
  await number.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }))); await expect(apply(panel)).toBeEnabled()
  const light = await text.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.documentImageFixture.dark(true)); await expect.poll(() => text.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  await apply(panel).click(); await expect(panel).toBeHidden(); assert.equal((await datum('image')).alt, '中文图片')
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
})

test('canceled or stale drafts retain their inputs and a successful observer navigation keeps the new input target', async () => {
  await editImage(); await format().getByRole('textbox', { name: '替代文字', exact: true }).fill('Draft')
  await page.evaluate(() => window.documentImageFixture.cancelNext()); await apply(format()).click()
  await expect(format()).toBeVisible(); await expect(format().getByRole('alert')).toContainText(/canceled|取消/)
  assert.equal((await datum('image')).alt, 'image'); assert.equal((await history()).undo, 0)
  await page.evaluate(() => window.documentImageFixture.select('tail', 1)); await expect(apply(format())).toBeDisabled()
  await format().getByRole('button', { name: '取消', exact: true }).click(); await expect(format()).toBeHidden()
  await editImage(); await format().getByRole('textbox', { name: '替代文字', exact: true }).fill('Committed')
  await page.evaluate(() => window.documentImageFixture.navigateOnUpdate()); await apply(format()).click(); await expect(format()).toBeHidden()
  assert.equal((await datum('image')).alt, 'Committed')
  assert.equal((await page.evaluate(() => window.documentImageFixture.selection())).selection.focus.blockId, 'tail')
  await load('?standalone'); await editImage(); await format().getByRole('textbox', { name: '替代文字', exact: true }).fill('Standalone draft')
  await page.evaluate(() => window.documentImageFixture.replaceSameSnapshot()); await expect(apply(format())).toBeDisabled()
  await expect(format().getByRole('textbox', { name: '替代文字', exact: true })).toHaveValue('Standalone draft')
})

test('table, header and note images edit through real Canvas hosts and retain separate owner identities on reopening', async () => {
  await load('?auxiliary'); await page.evaluate(() => window.documentImageFixture.select('cell-p'))
  await editImage('cell-image'); await format().getByRole('textbox', { name: '替代文字', exact: true }).fill('Cell edited'); await apply(format()).click(); await expect(format()).toBeHidden()
  const header = { kind: 'story', sectionId: 'first', story: 'header' }
  await page.evaluate(host => window.documentImageFixture.select('header-p', 0, host), header)
  await editImage('header-image', header); await format().getByRole('textbox', { name: '替代文字', exact: true }).fill('Header edited'); await apply(format()).click(); await expect(format()).toBeHidden()
  const noteId = await page.evaluate(() => window.documentImageFixture.noteId), note = { kind: 'note', noteKind: 'footnote', noteId }
  await page.evaluate(host => window.documentImageFixture.select('note-p', 0, host), note)
  await editImage('note-image', note); await format().getByRole('textbox', { name: '替代文字', exact: true }).fill('Note edited'); await apply(format()).click(); await expect(format()).toBeHidden()
  await page.evaluate(() => window.documentImageFixture.saveAndReopen())
  assert.equal((await images(header))[0].image.alt, 'Header edited'); assert.equal((await images(note))[0].image.alt, 'Note edited')
  assert.equal((await datum('cell-image')).alt, 'Cell edited'); assert.equal((await datum('image')).alt, 'image')
})

test('locked picture ranges disable context mutation and KeepAlive deactivation closes pending editors', async () => {
  await load('?locked'); await clickImage('image', { button: 'right' })
  for (const label of ['编辑图片', '裁剪图片', '替换图片', '删除图片']) await expect(page.getByRole('menuitem', { name: label, exact: true })).toHaveClass(/ant-menu-item-disabled/)
  assert.equal((await history()).undo, 0)
  await load(); await editImage(); await format().getByRole('textbox', { name: '替代文字', exact: true }).fill('Uncommitted')
  await page.evaluate(() => window.documentImageFixture.deactivate()); await expect(format()).toBeHidden()
  await page.evaluate(() => window.documentImageFixture.reactivate()); assert.equal((await datum('image')).alt, 'image')
  assert.equal((await history()).undo, 0)
})
