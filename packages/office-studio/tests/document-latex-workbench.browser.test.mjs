import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred browser drafts use the real Canvas and antdv-next editor. Synthetic
// composition is a guard check, not acceptance for the Windows native IME.
const fixturePath = '/__document-latex-test__'
let server, browser, page, baseURL, errors
beforeAll(async () => {
  const fixture = await readFile(new URL('./fixtures/document-latex.html', import.meta.url), 'utf8')
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-latex-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => {
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
  await page.goto(baseURL + query); await page.waitForFunction(() => !!window.documentLatexFixture)
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeAttached()
}
const source = () => page.getByRole('textbox', { name: 'LaTeX 源码', exact: true })
const panel = () => page.getByRole('dialog', { name: '编辑 LaTeX 公式', exact: true })
const apply = () => panel().getByRole('button', { name: '应用', exact: true })
const cancel = () => panel().getByRole('button', { name: '取消', exact: true })
const formulas = () => page.evaluate(() => window.documentLatexFixture.formulas())
async function clickFormula(id = 'p', fraction = 0.5, options = {}, host) {
  const point = await page.evaluate(({ id, fraction, host }) => window.documentLatexFixture.point(id, fraction, host), { id, fraction, host })
  await page.mouse.click(point.x, point.y, options)
}
async function editFormula(id = 'p', host) { await clickFormula(id, 0.5, { clickCount: 2 }, host); await expect(source()).toBeFocused() }

test('Docs initializes its Ribbon and translates a stale formula error after changing the active locale', async () => {
  await expect(page.locator('.document-ribbon')).toBeVisible()
  await editFormula(); await page.evaluate(() => window.documentLatexFixture.select('tail', 1))
  await expect(apply()).toBeDisabled(); await expect(panel().getByRole('alert')).toContainText('文档或公式选区已变化')
  await page.evaluate(() => window.documentLatexFixture.setLocale('en-US'))
  await expect(panel().getByRole('alert')).toContainText('The document or formula selection has changed')
  await expect(apply()).toBeDisabled(); await cancel().click(); await expect(panel()).toBeHidden()
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
})

test('both Canvas halves select the entire formula and double-click editing has one undo and a stable saved ID', async () => {
  const original = (await formulas())[0]
  for (const fraction of [0.1, 0.9]) {
    await clickFormula('p', fraction)
    const selection = await page.evaluate(() => window.documentLatexFixture.selection())
    assert.deepEqual([selection.selection.anchor.offset, selection.selection.focus.offset], [7, 8])
  }
  await editFormula(); await expect(source()).toHaveValue('  x ^2  ')
  await source().fill('  y_2  '); await apply().click(); await expect(panel()).toBeHidden()
  assert.equal((await formulas())[0].formula.latex, '  y_2  ')
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 1)
  await page.evaluate(() => window.documentLatexFixture.undo()); assert.equal((await formulas())[0].formula.latex, '  x ^2  ')
  await page.evaluate(() => window.documentLatexFixture.redo())
  await page.evaluate(() => window.documentLatexFixture.saveAndReopen())
  assert.equal((await formulas())[0].rangeId, original.rangeId); assert.equal((await formulas())[0].formula.latex, '  y_2  ')
})

test('right-click edit, delete and conversion use the real formula target and content history', async () => {
  await clickFormula('p', 0.9, { button: 'right' }); await page.getByRole('menuitem', { name: '编辑公式', exact: true }).click()
  await expect(source()).toHaveValue('  x ^2  '); await cancel().click(); await expect(panel()).toBeHidden()
  await clickFormula('p', 0.2, { button: 'right' }); await page.getByRole('menuitem', { name: '将公式转为源文本', exact: true }).click()
  assert.equal((await formulas()).some(item => item.blockId === 'p'), false)
  assert.ok((await page.evaluate(() => window.documentLatexFixture.publicSnapshot())).body.dataStream.includes('  x ^2  '))
  await page.evaluate(() => window.documentLatexFixture.undo())
  await clickFormula('p', 0.8, { button: 'right' }); await page.getByRole('menuitem', { name: '删除公式', exact: true }).click()
  assert.equal((await formulas()).some(item => item.blockId === 'p'), false)
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 1)
})

test('a canceled submission leaves the Ant modal open with its source and an actionable error', async () => {
  await editFormula(); await source().fill('z^3'); await page.evaluate(() => window.documentLatexFixture.cancelNext())
  await apply().click(); await expect(panel()).toBeVisible(); await expect(source()).toHaveValue('z^3')
  await expect(panel().getByRole('alert')).toContainText(/canceled|取消/)
  assert.equal((await formulas())[0].formula.latex, '  x ^2  ')
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
  await source().fill('z^4'); await apply().click(); await expect(panel()).toBeHidden()
  assert.equal((await formulas())[0].formula.latex, 'z^4')
})

test('stale modal selection cannot submit and observer navigation is preserved after a successful update', async () => {
  await editFormula(); await source().fill('z^2'); await page.evaluate(() => window.documentLatexFixture.select('tail', 1))
  await expect(apply()).toBeDisabled(); await expect(panel().getByRole('alert')).toContainText('选区已变化')
  await cancel().click(); await expect(panel()).toBeHidden()
  assert.equal((await page.evaluate(() => window.documentLatexFixture.selection())).selection.focus.blockId, 'tail')
  await editFormula(); await source().fill('z^3'); await page.evaluate(() => window.documentLatexFixture.navigateOnUpdate())
  await apply().click(); await expect(panel()).toBeHidden()
  assert.equal((await page.evaluate(() => window.documentLatexFixture.selection())).selection.focus.blockId, 'tail')
  assert.equal((await formulas())[0].formula.latex, 'z^3')
})

test('Ant source composition suppresses Escape and Ctrl+Enter until the final input and preserves theme adaptation', async () => {
  await editFormula(); await source().fill('x_2')
  await source().evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await source().press('Control+Enter'); await source().press('Escape'); await expect(panel()).toBeVisible(); await expect(apply()).toBeDisabled()
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
  await source().evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'x_2' })))
  await expect(apply()).toBeEnabled()
  const light = await source().evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.documentLatexFixture.dark(true))
  await expect.poll(() => source().evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  await source().press('Control+Enter'); await expect(panel()).toBeHidden()
  assert.equal((await formulas())[0].formula.latex, 'x_2')
})

test('persisted malformed source is selectable and repairable while locked formulas disable all context mutations', async () => {
  await load('?invalid'); await editFormula(); await expect(source()).toHaveValue('  \\bad')
  await expect(panel()).toContainText('第 3 个位置'); await expect(apply()).toBeDisabled()
  await source().fill('\\frac{a}{b}'); await apply().click(); await expect(panel()).toBeHidden()
  assert.equal((await formulas())[0].formula.latex, '\\frac{a}{b}')
  await load('?locked'); await clickFormula('p', 0.8, { button: 'right' })
  for (const title of ['编辑公式', '删除公式', '将公式转为源文本']) await expect(page.getByRole('menuitem', { name: title, exact: true })).toHaveClass(/ant-menu-item-disabled/)
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
})

test('nested table, header and note formulas edit through their actual Canvas hosts and survive save/reopen', async () => {
  await load('?auxiliary'); await editFormula('cell-p'); await source().fill('c_2'); await apply().click(); await expect(panel()).toBeHidden()
  const header = { kind: 'story', sectionId: 'first', story: 'header' }
  await page.evaluate(host => window.documentLatexFixture.select('header-p', 0, 1, host), header)
  await editFormula('header-p', header); await source().fill('h_2'); await apply().click(); await expect(panel()).toBeHidden()
  const noteId = await page.evaluate(() => window.documentLatexFixture.noteId), note = { kind: 'note', noteKind: 'footnote', noteId }
  await page.evaluate(host => window.documentLatexFixture.select('note-p', 0, 1, host), note)
  await editFormula('note-p', note); await source().fill('n_2'); await apply().click(); await expect(panel()).toBeHidden()
  await page.evaluate(() => window.documentLatexFixture.saveAndReopen())
  assert.equal((await page.evaluate(host => window.documentLatexFixture.formulas(host), header))[0].formula.latex, 'h_2')
  assert.equal((await page.evaluate(host => window.documentLatexFixture.formulas(host), note))[0].formula.latex, 'n_2')
})

test('standalone same-content model replacement leaves the draft visible but invalidates its captured submission target', async () => {
  await load('?standalone'); await editFormula(); await source().fill('draft_2')
  await page.evaluate(() => window.documentLatexFixture.replaceSameSnapshot())
  await expect(panel()).toBeVisible(); await expect(source()).toHaveValue('draft_2'); await expect(apply()).toBeDisabled()
  await expect(panel().getByRole('alert')).toContainText('选区已变化')
  assert.equal((await formulas())[0].formula.latex, '  x ^2  ')
})
