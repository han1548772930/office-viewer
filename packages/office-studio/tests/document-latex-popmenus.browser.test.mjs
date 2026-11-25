import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred real Ant/Canvas interaction drafts. No runner, dev or browser was
// started here; synthetic composition does not validate the Windows native IME.
const fixturePath = '/__document-latex-popmenus-test__'
let server, browser, page, baseURL, errors
beforeAll(async () => {
  const fixture = await readFile(new URL('./fixtures/document-latex.html', import.meta.url), 'utf8')
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-latex-popmenus-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => {
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
  errors = []; page = await browser.newPage({ viewport: { width: 1600, height: 960 }, deviceScaleFactor: 1.25 })
  page.on('pageerror', error => errors.push(error.message)); await load('?entrances')
})
afterEach(async () => { await page?.close(); assert.deepEqual(errors, []) })
async function load(query = '') {
  await page.goto(baseURL + query); await page.waitForFunction(() => !!window.documentLatexFixture)
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeAttached()
}
const source = () => page.getByRole('textbox', { name: 'LaTeX 源码', exact: true })
const panel = () => page.getByRole('dialog').filter({ has: source() })
const apply = () => panel().getByRole('button', { name: '应用', exact: true })
const cancel = () => panel().getByRole('button', { name: '取消', exact: true })
const menus = () => page.locator('[role="menu"]:visible')
const formulas = host => page.evaluate(host => window.documentLatexFixture.formulas(host), host)
const history = () => page.evaluate(() => window.documentLatexFixture.history())
const snapshot = () => page.evaluate(() => window.documentLatexFixture.snapshot())
async function select(id, start = 0, end = start, host) {
  await page.evaluate(value => window.documentLatexFixture.select(value.id, value.start, value.end, value.host), { id, start, end, host })
  await expect.poll(() => page.evaluate(() => window.documentLatexFixture.selection()?.selection.focus.cellBlockId ?? window.documentLatexFixture.selection()?.selection.focus.blockId)).toBe(id)
}
async function openFormula(id = 'p', host) {
  const item = (await formulas(host)).find(item => item.blockId === id)
  assert.ok(item)
  const segmentId = host ? await page.evaluate(host => window.documentLatexFixture.segmentIdForHost(host), host) : ''
  assert.equal(await page.evaluate(request => window.documentLatexFixture.openPopup(request), { rangeId: item.rangeId, segmentId }), true)
  await expect(source()).toBeFocused(); return item
}
async function openParagraph() {
  await page.locator('[data-document-paragraph-insert]').getByRole('button').click()
  await expect(menus().first()).toBeVisible()
}

test.each(['empty-layout', 'insert-quick-bottom', 'insert-menu', 'edit-quick-bottom', 'float-text-style'])('%s performs an actual selected-range edit through the common popup, undo/redo and save/reopen', async mount => {
  const id = ['empty-layout', 'insert-quick-bottom', 'insert-menu'].includes(mount) ? 'empty' : 'text'
  await select(id, 0, mount === 'float-text-style' ? 'x + y 😀'.length : 0)
  const before = await snapshot(), previous = await history()
  if (mount === 'empty-layout') {
    const point = await page.evaluate(id => window.documentLatexFixture.textPoint(id), id)
    await page.mouse.click(point.x, point.y, { button: 'right' })
    await menus().last().getByRole('menuitem', { name: 'LaTeX 公式', exact: true }).click()
  } else if (mount === 'float-text-style') {
    await page.locator('[data-document-latex-mount="float-text-style"]').click()
  } else {
    await openParagraph()
    if (mount === 'insert-menu') await menus().first().getByRole('menuitem', { name: '插入', exact: true }).hover()
    await menus().last().getByRole('menuitem', { name: 'LaTeX 公式', exact: true }).click()
  }
  await expect(source()).toBeFocused(); await expect(source()).toHaveValue(mount === 'float-text-style' ? 'x + y 😀' : '')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previous)
  await source().fill('x_3'); await apply().click(); await expect(panel()).toBeHidden()
  const created = (await formulas()).find(item => item.blockId === id)
  assert.ok(created); assert.equal(created.formula.latex, 'x_3'); assert.equal((await history()).undo, previous.undo + 1)
  await page.evaluate(() => window.documentLatexFixture.undo()); assert.equal((await formulas()).some(item => item.rangeId === created.rangeId), false)
  await page.evaluate(() => window.documentLatexFixture.redo()); await page.evaluate(() => window.documentLatexFixture.saveAndReopen())
  assert.equal((await formulas()).find(item => item.rangeId === created.rangeId).formula.latex, 'x_3')
})

test('public inclusive indices and a backward exclusive selection navigate to the complete formula source without content history', async () => {
  await select('tail', 1)
  const item = (await formulas()).find(item => item.blockId === 'p')
  const range = await page.evaluate(id => window.documentLatexFixture.formulaRange(id), item.rangeId)
  const before = await snapshot(), previous = await history()
  assert.equal(await page.evaluate(request => window.documentLatexFixture.openPopup(request), { rangeId: item.rangeId,
    startIndex: range.startOffset, endIndex: range.endOffset - 1,
    selection: { startOffset: range.startOffset, endOffset: range.endOffset, collapsed: false, direction: 'backward' } }), true)
  await expect(source()).toHaveValue(item.formula.latex); await expect(source()).toBeFocused()
  const state = await page.evaluate(() => window.documentLatexFixture.selection())
  assert.ok(state.selection.anchor.offset > state.selection.focus.offset); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), previous)
  assert.equal(await page.evaluate(() => window.documentLatexFixture.closePopup(true)), true); await expect(panel()).toBeHidden()
})

test('reopening the same visible popup resets its source draft without changing the formula or duplicating history', async () => {
  const item = await openFormula(); await source().fill('unapplied_2')
  assert.equal(await page.evaluate(request => window.documentLatexFixture.openPopup(request, true), { rangeId: item.rangeId, initialLatex: 'reset_3' }), true)
  await expect(source()).toHaveValue('reset_3'); assert.equal((await formulas()).find(value => value.rangeId === item.rangeId).formula.latex, item.formula.latex)
  assert.equal((await history()).undo, 0); await cancel().click(); await expect(panel()).toBeHidden()
})

test('a canceled close operation keeps the source draft visible with feedback and a later close restores valid input focus', async () => {
  const item = await openFormula(); await source().fill('retained_4')
  await page.evaluate(() => window.documentLatexFixture.cancelNextPopupClose()); await cancel().click()
  await expect(panel()).toBeVisible(); await expect(source()).toHaveValue('retained_4'); await expect(panel().getByRole('alert')).toContainText('canceled')
  assert.equal((await formulas()).find(value => value.rangeId === item.rangeId).formula.latex, item.formula.latex)
  assert.equal((await history()).undo, 0); await cancel().click(); await expect(panel()).toBeHidden()
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
})

test('a queued explicit popup cannot reopen after deactivation/reactivation and a new request can still open', async () => {
  await select('tail', 1)
  const item = (await formulas()).find(item => item.blockId === 'p')
  assert.equal(await page.evaluate(request => window.documentLatexFixture.bouncePopup(request), { rangeId: item.rangeId }), true)
  await expect(panel()).toHaveCount(0); assert.equal((await history()).undo, 0)
  await openFormula(); await expect(source()).toHaveValue(item.formula.latex)
})

test('an input observer that navigates during explicit popup selection keeps its new location and prevents the old popup', async () => {
  await select('tail', 1)
  const item = (await formulas()).find(item => item.blockId === 'p'), before = await snapshot()
  await page.evaluate(() => window.documentLatexFixture.navigateOnPopupSelection())
  assert.equal(await page.evaluate(request => window.documentLatexFixture.openPopup(request), { rangeId: item.rangeId }), false)
  await expect(panel()).toHaveCount(0)
  const state = await page.evaluate(() => window.documentLatexFixture.selection())
  assert.equal(state.selection.focus.blockId, 'tail'); assert.equal(state.selection.focus.offset, 1)
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})

test('synthetic composition rejects public/local close, keyboard apply and stale composition completion without losing the draft', async () => {
  const item = await openFormula(); await source().fill('x_4')
  await source().evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  assert.equal(await page.evaluate(() => window.documentLatexFixture.closePopup()), false)
  assert.equal(await page.evaluate(() => window.documentLatexFixture.closePopup(true)), false)
  await source().press('Escape'); await source().press('Control+Enter'); await expect(panel()).toBeVisible(); await expect(source()).toHaveValue('x_4')
  await source().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'x_4' }))
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
  })
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(resolve)))
  assert.equal(await page.evaluate(() => window.documentLatexFixture.closePopup()), false)
  await source().evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'x_4' })))
  await expect(apply()).toBeEnabled(); assert.equal(await page.evaluate(() => window.documentLatexFixture.closePopup()), true)
  await expect(panel()).toBeHidden(); assert.equal((await formulas()).find(value => value.rangeId === item.rangeId).formula.latex, item.formula.latex)
  assert.equal((await history()).undo, 0)
})

test('a captured paragraph menu rejects a later selection change without moving the new caret', async () => {
  await select('text', 0); await openParagraph(); await select('tail', 1)
  await menus().last().getByRole('menuitem', { name: 'LaTeX 公式', exact: true }).click()
  await expect(panel()).toHaveCount(0)
  const state = await page.evaluate(() => window.documentLatexFixture.selection())
  assert.equal(state.selection.focus.blockId, 'tail'); assert.equal(state.selection.focus.offset, 1); assert.equal((await history()).undo, 0)
})

test('normal and native table text expose the floating formula action while block-range quote text hides it', async () => {
  await select('text', 0, 1); await expect(page.locator('[data-document-latex-mount="float-text-style"]')).toBeVisible()
  await select('quote-text', 0, 6); await expect(page.locator('[data-document-latex-mount="float-text-style"]')).toHaveCount(0)
  await select('cell-text', 0, 9); await expect(page.locator('[data-document-latex-mount="float-text-style"]')).toBeVisible()
  assert.equal((await history()).undo, 0)
})

test('mobile formula, structure and paragraph tools stay disjoint and edit/remove perform real reversible saved commands', async () => {
  await page.setViewportSize({ width: 390, height: 844 }); await select('quote-p', 0)
  const dock = page.locator('[data-document-mobile-quick-tools]'), actions = page.locator('[data-document-mobile-latex-actions]')
  await expect(actions).toBeVisible()
  const boxes = await dock.locator(':scope > div:visible').evaluateAll(elements => elements.map(element => { const rect = element.getBoundingClientRect(); return { top: rect.top, bottom: rect.bottom, left: rect.left, right: rect.right } }))
  assert.ok(boxes.length >= 3)
  for (let index = 1; index < boxes.length; index++) assert.ok(boxes[index].top >= boxes[index - 1].bottom)
  for (const box of boxes) { assert.ok(box.left >= 0); assert.ok(box.right <= 390) }
  const original = (await formulas()).find(item => item.blockId === 'quote-p')
  await actions.locator('[data-document-latex-action="doc-context-latex-edit"]').click()
  await expect(source()).toHaveValue(original.formula.latex); await expect(page.locator('[data-document-latex-mobile-popup]')).toBeVisible()
  const rect = await panel().boundingBox(); assert.ok(rect.x >= 0 && rect.x + rect.width <= 390)
  await source().fill('z_5'); await apply().click(); await expect(panel()).toBeHidden()
  await select('quote-p', 0); await actions.locator('[data-document-latex-action="doc-context-latex-remove"]').click()
  assert.equal((await formulas()).some(item => item.rangeId === original.rangeId), false); assert.equal((await history()).undo, 2)
  await page.evaluate(() => window.documentLatexFixture.undo()); assert.equal((await formulas()).find(item => item.rangeId === original.rangeId).formula.latex, 'z_5')
  await page.evaluate(() => window.documentLatexFixture.redo()); await page.evaluate(() => window.documentLatexFixture.saveAndReopen())
  assert.equal((await formulas()).some(item => item.rangeId === original.rangeId), false)
})

test('a short landscape mobile popup bounds its scrolling body and footer within the viewport', async () => {
  await page.setViewportSize({ width: 568, height: 320 }); await openFormula()
  const rect = await panel().boundingBox(), body = await page.locator('[data-document-latex-mobile-popup]').boundingBox()
  assert.ok(rect.x >= 0 && rect.x + rect.width <= 568); assert.ok(rect.y >= 0 && rect.y + rect.height <= 321)
  assert.ok(body.height <= 100); await expect(cancel()).toBeVisible(); await cancel().click(); await expect(panel()).toBeHidden()
})

test('explicit header/note popup requests preserve the real host, source and saved formula identity through author edits', async () => {
  await load('?auxiliary')
  const header = { kind: 'story', sectionId: 'first', story: 'header' }, noteId = await page.evaluate(() => window.documentLatexFixture.noteId)
  const note = { kind: 'note', noteKind: 'footnote', noteId }, expected = [], previous = await history()
  for (const target of [{ id: 'header-p', host: header }, { id: 'note-p', host: note }]) {
    await select('tail', 0); const item = await openFormula(target.id, target.host)
    const state = await page.evaluate(() => window.documentLatexFixture.selection())
    assert.deepEqual(state.host, target.host); await expect(source()).toHaveValue(item.formula.latex)
    const next = '\\sqrt{' + item.formula.latex + '}'; await source().fill(next); await apply().click(); await expect(panel()).toBeHidden()
    expected.push({ ...target, id: item.rangeId, latex: next })
  }
  assert.equal((await history()).undo, previous.undo + 2)
  await page.evaluate(() => window.documentLatexFixture.undo()); await page.evaluate(() => window.documentLatexFixture.redo())
  await page.evaluate(() => window.documentLatexFixture.saveAndReopen())
  for (const item of expected) assert.equal((await formulas(item.host)).find(value => value.rangeId === item.id).formula.latex, item.latex)
})

test('a locked paragraph quick action stays disabled and cannot author or open an editor', async () => {
  await load('?locked'); await select('p', 7); await openParagraph()
  await expect(menus().last().getByRole('menuitem', { name: 'LaTeX 公式', exact: true })).toHaveAttribute('aria-disabled', 'true')
  await expect(panel()).toHaveCount(0); assert.equal((await history()).undo, 0)
})

test('mobile theme changes update the dock tokens and preserve the source draft and native source selection', async () => {
  await page.setViewportSize({ width: 390, height: 844 }); await select('quote-p', 0)
  const tools = page.locator('[data-document-mobile-latex-actions]'), light = await tools.evaluate(element => getComputedStyle(element).backgroundColor)
  await tools.locator('[data-document-latex-action="doc-context-latex-edit"]').click(); await source().fill('x_6')
  await source().evaluate(element => element.setSelectionRange(0, 1)); await page.evaluate(() => window.documentLatexFixture.dark(true))
  await expect.poll(() => tools.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  await expect(source()).toHaveValue('x_6'); assert.deepEqual(await source().evaluate(element => [element.selectionStart, element.selectionEnd]), [0, 1])
  assert.equal((await history()).undo, 0)
})
