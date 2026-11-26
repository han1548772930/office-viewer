import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred real Canvas / Ant / history drafts. No Vite or browser was started.
// ClipboardEvent and composition below are synthetic; OS clipboard/IME and GPU acceptance remain separate.
const fixturePath = '/__document-image-batch-test__', ids = ['a', 'b', 'c']
let server, browser, page, baseURL, errors
beforeAll(async () => {
  const fixture = await readFile(new URL('./fixtures/document-images.html', import.meta.url), 'utf8')
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-image-batch-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => {
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
async function load(query = '?multi') {
  await page.goto(baseURL + query); await page.waitForFunction(() => !!window.documentImageFixture)
  await expect(input()).toBeAttached()
}
const input = () => page.getByRole('textbox', { name: '文档文字输入', exact: true })
const toolbar = () => page.getByRole('toolbar', { name: '图片操作', exact: true })
const count = value => expect(page.locator('[data-document-image-selection-count]')).toHaveText('已选 ' + value + ' 张')
const images = () => page.evaluate(() => window.documentImageFixture.images())
const history = () => page.evaluate(() => window.documentImageFixture.history())
const rectangles = (selected = ids) => page.evaluate(selected => window.documentImageFixture.imageRectangles(selected), selected)
const point = id => page.evaluate(id => window.documentImageFixture.point(id), id)
const groupPoint = (handle, selected = ids) => page.evaluate(({ handle, selected }) => window.documentImageFixture.groupPoint(selected, handle), { handle, selected })
const near = (actual, expected, tolerance = 0.15) => assert.ok(Math.abs(actual - expected) < tolerance, `${actual} != ${expected}`)
async function clickImage(id, modifier, options = {}) {
  const position = await point(id)
  if (modifier) await page.keyboard.down(modifier)
  try { await page.mouse.click(position.x, position.y, options) } finally { if (modifier) await page.keyboard.up(modifier) }
}
async function select(selected = ids) {
  await input().press('Escape')
  await clickImage(selected[0]); await expect(toolbar()).toBeVisible()
  for (const [index, id] of selected.slice(1).entries()) { await clickImage(id, 'Control'); await count(index + 2) }
  await expect(input()).toBeFocused()
}
async function beginDrag(id = 'a') {
  const position = await point(id); await page.mouse.move(position.x, position.y); await page.mouse.down(); return position
}

test('Ctrl and Shift toggle members while an ordinary click inside the selection keeps every selected object', async () => {
  await select(['a', 'b']); await clickImage('c', 'Shift'); await count(3)
  await clickImage('b', 'Control'); await count(2)
  await clickImage('c'); await count(2)
  await clickImage('d'); await expect(page.locator('[data-document-image-selection-count]')).toHaveCount(0)
  await expect(toolbar().getByRole('button', { name: '编辑图片', exact: true })).toBeVisible()
  assert.equal((await history()).undo, 0)
})

test('the three multi-image roots and eight ordered alignment leaves remain available in grid, with distribution disabled for two images', async () => {
  await load('?multi&ribbon=grid'); await select(['a', 'b'])
  await expect(toolbar().getByRole('button')).toHaveCount(3)
  await expect(toolbar().getByRole('button', { name: '编辑图片', exact: true })).toHaveCount(0)
  await toolbar().getByRole('button', { name: '对齐与分布', exact: true }).click()
  const labels = ['左对齐', '水平居中', '右对齐', '顶部对齐', '垂直居中', '底部对齐', '水平分布', '垂直分布']
  assert.deepEqual(await page.getByRole('menuitem').allTextContents(), labels)
  for (const label of labels.slice(-2)) await expect(page.getByRole('menuitem', { name: label, exact: true })).toHaveClass(/ant-menu-item-disabled/)
  await page.keyboard.press('Escape'); await expect(input()).toBeFocused(); await clickImage('c', 'Control'); await count(3)
  await toolbar().getByRole('button', { name: '对齐与分布', exact: true }).click()
  for (const label of labels) await expect(page.getByRole('menuitem', { name: label, exact: true })).not.toHaveClass(/ant-menu-item-disabled/)
  await page.keyboard.press('Escape'); await expect(input()).toBeFocused(); assert.equal((await history()).undo, 0)
})

test('each of the eight menu leaves changes real Canvas geometry once and undo restores the original objects', async () => {
  const labels = ['左对齐', '水平居中', '右对齐', '顶部对齐', '垂直居中', '底部对齐', '水平分布', '垂直分布']
  for (const [index, label] of labels.entries()) {
    await load(); await select(); const before = await images(), frames = await rectangles()
    await toolbar().getByRole('button', { name: '对齐与分布', exact: true }).click()
    await page.getByRole('menuitem', { name: label, exact: true }).click(); await expect(input()).toBeFocused()
    await expect.poll(async () => (await history()).undo).toBe(1)
    const next = await rectangles(), horizontal = [0, 1, 2, 6].includes(index), coordinate = rectangle => horizontal ? rectangle.x : rectangle.y
    const size = rectangle => horizontal ? rectangle.width : rectangle.height
    if (index < 6) {
      const edge = rectangle => coordinate(rectangle) + size(rectangle) * ([1, 4].includes(index) ? 0.5 : [2, 5].includes(index) ? 1 : 0)
      for (const rectangle of next) near(edge(rectangle), edge(next[0]))
      for (const [position, rectangle] of next.entries()) near(horizontal ? rectangle.y : rectangle.x, horizontal ? frames[position].y : frames[position].x)
    } else {
      const sorted = [...next].sort((a, b) => coordinate(a) - coordinate(b)), gap = coordinate(sorted[1]) - coordinate(sorted[0]) - size(sorted[0])
      near(coordinate(sorted[2]) - coordinate(sorted[1]) - size(sorted[1]), gap)
      near(coordinate(sorted[0]), Math.min(...frames.map(coordinate)))
      near(coordinate(sorted.at(-1)) + size(sorted.at(-1)), Math.max(...frames.map(rectangle => coordinate(rectangle) + size(rectangle))))
    }
    await page.evaluate(() => window.documentImageFixture.undo()); assert.deepEqual(await images(), before)
  }
}, { timeout: 90_000 })

test('a right-click inside a multi-selection preserves it and batch deletion uses one undo before save/reopen', async () => {
  await select(); const before = await images(); await clickImage('a', undefined, { button: 'right' }); await count(3)
  await page.mouse.move(4, 4)
  const menu = page.locator('[aria-label="图片右键菜单"]')
  await expect(menu).toBeVisible(); await expect(menu.getByRole('menuitem')).toHaveCount(3)
  for (const label of ['排列', '对齐与分布', '删除选中的图片']) await expect(menu.getByRole('menuitem', { name: label, exact: true })).toBeVisible()
  await page.keyboard.press('Escape'); await expect(menu).toBeHidden(); await expect(input()).toBeFocused()
  await clickImage('a', undefined, { button: 'right' }); await menu.getByRole('menuitem', { name: '删除选中的图片', exact: true }).click()
  await expect.poll(async () => (await images()).map(image => image.imageId)).toEqual(['d']); assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.documentImageFixture.undo()); assert.deepEqual(await images(), before)
  await page.evaluate(() => window.documentImageFixture.redo()); await page.evaluate(() => window.documentImageFixture.saveAndReopen())
  assert.deepEqual((await images()).map(image => image.imageId), ['d'])
})

test('batch dragging previews without changing content and commits every selected offset once on pointer-up', async () => {
  await select(); const before = await images(), position = await beginDrag()
  await page.mouse.move(position.x + 35, position.y + 25, { steps: 6 })
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
  await page.mouse.up(); await expect.poll(async () => (await history()).undo).toBe(1)
  const next = await images()
  for (const item of next) {
    const old = before.find(candidate => candidate.imageId === item.imageId).image
    if (ids.includes(item.imageId)) { near(item.image.positionH.posOffset, old.positionH.posOffset + 35); near(item.image.positionV.posOffset, old.positionV.posOffset + 25) }
    else assert.deepEqual(item.image, old)
  }
  await page.evaluate(() => window.documentImageFixture.undo()); assert.deepEqual(await images(), before)
  await page.evaluate(() => window.documentImageFixture.redo()); await page.evaluate(() => window.documentImageFixture.saveAndReopen())
  assert.deepEqual(await images(), next)
})

test('the union corner uniformly scales every local image frame and keeps the opposite selection corner fixed', async () => {
  await select(); const before = await images(), opposite = await groupPoint('nw'), position = await groupPoint('se'), scale = 1.18
  await page.mouse.move(position.x, position.y); await page.mouse.down()
  await page.mouse.move(opposite.x + (position.x - opposite.x) * scale, opposite.y + (position.y - opposite.y) * scale, { steps: 6 })
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
  await page.mouse.up(); await expect.poll(async () => (await history()).undo).toBe(1)
  const fixed = await groupPoint('nw'); near(fixed.x, opposite.x, 1); near(fixed.y, opposite.y, 1)
  for (const item of (await images()).filter(item => ids.includes(item.imageId))) {
    const old = before.find(candidate => candidate.imageId === item.imageId).image
    near(item.image.width, old.width * scale, 0.5); near(item.image.height, old.height * scale, 0.5)
    near(item.image.width / item.image.height, old.width / old.height, 0.001); near(item.image.angle ?? 0, old.angle ?? 0)
  }
  await page.evaluate(() => window.documentImageFixture.undo()); assert.deepEqual(await images(), before)
})

test('the rotation handle moves the selection as one unit with Shift snap and preserves each local rotation', async () => {
  await select(); const before = await images(), center = await groupPoint('center'), position = await groupPoint('rotate')
  const radius = center.y - position.y, angle = 22 * Math.PI / 180
  await page.keyboard.down('Shift'); await page.mouse.move(position.x, position.y); await page.mouse.down()
  await page.mouse.move(center.x + radius * Math.sin(angle), center.y - radius * Math.cos(angle), { steps: 6 })
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
  await page.mouse.up(); await page.keyboard.up('Shift'); await expect.poll(async () => (await history()).undo).toBe(1)
  for (const item of (await images()).filter(item => ids.includes(item.imageId))) {
    const old = before.find(candidate => candidate.imageId === item.imageId).image
    near(item.image.angle, (old.angle + 15) % 360); assert.deepEqual([item.image.width, item.image.height, item.image.resourceId], [old.width, old.height, old.resourceId])
  }
  await page.evaluate(() => window.documentImageFixture.undo()); assert.deepEqual(await images(), before)
})

test('Escape cancels an active gesture and a second Escape clears selection without editing content', async () => {
  await select(); const before = await images(), position = await beginDrag()
  await page.mouse.move(position.x + 30, position.y + 40, { steps: 4 }); await input().press('Escape'); await page.mouse.up()
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0); await count(3)
  await input().press('Escape'); await expect(toolbar()).toBeHidden(); assert.deepEqual(await images(), before)
})

test('returning to the starting point removes the preview and avoids creating an obsolete transform history', async () => {
  await select(); const before = await images(), position = await beginDrag()
  await page.mouse.move(position.x + 30, position.y + 40, { steps: 4 }); await page.mouse.move(position.x, position.y, { steps: 4 }); await page.mouse.up()
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0); await count(3)
})

test('keyboard movement uses one-pixel and Shift ten-pixel batch transactions and Delete removes all selected images once', async () => {
  await select(); const before = await images()
  await input().press('ArrowRight'); await input().press('Shift+ArrowDown'); await expect.poll(async () => (await history()).undo).toBe(2)
  for (const item of (await images()).filter(item => ids.includes(item.imageId))) {
    const old = before.find(candidate => candidate.imageId === item.imageId).image
    near(item.image.positionH.posOffset, old.positionH.posOffset + 1); near(item.image.positionV.posOffset, old.positionV.posOffset + 10)
  }
  await input().press('Delete'); await expect.poll(async () => (await images()).map(image => image.imageId)).toEqual(['d'])
  assert.equal((await history()).undo, 3); await page.evaluate(() => window.documentImageFixture.undo())
  assert.equal((await images()).length, 4)
})

test('copy and cut carry every selected IMG and rich paste creates new drawing identities without copying unselected objects', async () => {
  await select(['c', 'a']); const before = await images(), copied = await page.evaluate(() => window.documentImageFixture.clipboard('copy'))
  assert.equal(copied.text, 'c\na'); assert.equal((copied.html.match(/<img\b/g) ?? []).length, 2)
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
  const cut = await page.evaluate(() => window.documentImageFixture.clipboard('cut')); assert.equal(cut.html, copied.html)
  await expect.poll(async () => (await images()).map(image => image.imageId)).toEqual(['b', 'd']); assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.documentImageFixture.undo()); assert.deepEqual(await images(), before)
  await page.evaluate(() => window.documentImageFixture.select('p', 0)); await input().focus()
  await page.evaluate(data => window.documentImageFixture.clipboard('paste', { 'text/plain': data.text, 'text/html': data.html }), copied)
  await expect.poll(async () => (await images()).length).toBe(6)
  const pasted = (await images()).filter(item => !before.some(old => old.imageId === item.imageId))
  assert.deepEqual(pasted.map(item => item.image.alt), ['c', 'a']); assert.ok(pasted.every(item => !['c', 'a'].includes(item.imageId)))
  const next = await images(); await page.evaluate(() => window.documentImageFixture.saveAndReopen()); assert.deepEqual(await images(), next)
})

test('a canceled batch command and a navigation change cannot commit an old pointer preview', async () => {
  await select(); const before = await images()
  await page.evaluate(() => window.documentImageFixture.cancelNext('doc.command.transform-images'))
  let position = await beginDrag(); await page.mouse.move(position.x + 30, position.y + 20, { steps: 4 }); await page.mouse.up()
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
  position = await beginDrag(); await page.mouse.move(position.x + 30, position.y + 20, { steps: 4 })
  await page.evaluate(() => window.documentImageFixture.select('tail', 1)); await page.mouse.up()
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
})

test('an independently committed model change cancels stale geometry while keeping the new content and its own history', async () => {
  await select(); const before = await images(), position = await beginDrag()
  await page.mouse.move(position.x + 30, position.y + 20, { steps: 4 })
  assert.equal(await page.evaluate(() => window.documentImageFixture.updateImage('d', { alt: 'External' })), true); await page.mouse.up()
  for (const item of (await images()).filter(item => ids.includes(item.imageId))) assert.deepEqual(item.image, before.find(old => old.imageId === item.imageId).image)
  assert.equal((await images()).find(item => item.imageId === 'd').image.alt, 'External'); assert.equal((await history()).undo, 1)
})

test('content locks disable all multi-object mutations and reject API geometry without changing an earlier selected object', async () => {
  await load('?multi-locked'); await select(['a', 'b']); const before = await images()
  for (const label of ['删除选中的图片', '排列图片', '对齐与分布']) await expect(toolbar().getByRole('button', { name: label, exact: true })).toBeDisabled()
  await input().press('ArrowRight'); await input().press('Delete')
  const error = await page.evaluate(() => { try { window.documentImageFixture.transformImages(['a', 'b'], { x: 10 }); return '' } catch (cause) { return cause.message } })
  assert.match(error, /锁定/); assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
})

test('Ant floating controls and their menus retain the dark theme and restore the document input focus', async () => {
  await select(); const button = toolbar().getByRole('button', { name: '对齐与分布', exact: true })
  const light = await toolbar().evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.documentImageFixture.dark(true))
  await expect.poll(() => toolbar().evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  await button.click(); await expect(page.getByRole('menuitem', { name: '左对齐', exact: true })).toBeVisible()
  const darkMenu = await page.getByRole('menu').last().evaluate(element => getComputedStyle(element).backgroundColor)
  assert.notEqual(darkMenu, light); await page.keyboard.press('Escape'); await expect(input()).toBeFocused()
  await button.click(); await page.getByRole('menuitem', { name: '左对齐', exact: true }).click(); await expect(input()).toBeFocused()
  assert.equal((await history()).undo, 1)
})

test('synthetic composition cancels active geometry and blocks Delete until the Ant input composition ends', async () => {
  await select(); const before = await images(), position = await beginDrag()
  await page.mouse.move(position.x + 30, position.y + 20, { steps: 4 })
  await input().evaluate(element => { element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    element.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', isComposing: true, bubbles: true, cancelable: true })) })
  await page.mouse.up(); assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
  await input().evaluate(element => { element.value = ''; element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '' })) })
  assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0); await expect(input()).toBeFocused()
})

test('KeepAlive deactivation releases the pending gesture and reactivation uses one current handler for batch movement', async () => {
  await select(); const before = await images(), position = await beginDrag()
  await page.mouse.move(position.x + 30, position.y + 20, { steps: 4 }); await page.evaluate(() => window.documentImageFixture.deactivate()); await page.mouse.up()
  await page.evaluate(() => window.documentImageFixture.reactivate()); assert.deepEqual(await images(), before); assert.equal((await history()).undo, 0)
  await select(); await input().press('ArrowRight'); await expect.poll(async () => (await history()).undo).toBe(1)
  for (const item of (await images()).filter(item => ids.includes(item.imageId))) near(item.image.positionH.posOffset, before.find(old => old.imageId === item.imageId).image.positionH.posOffset + 1)
})

test('a standalone model replacement rejects the pending batch gesture even when its JSON content is identical', async () => {
  await load('?multi&standalone'); await select(); const before = await images(), position = await beginDrag()
  await page.mouse.move(position.x + 30, position.y + 20, { steps: 4 }); await page.evaluate(() => window.documentImageFixture.replaceSameSnapshot()); await page.mouse.up()
  assert.deepEqual(await images(), before)
})
