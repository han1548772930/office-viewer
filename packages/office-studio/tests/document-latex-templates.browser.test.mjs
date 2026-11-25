import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Deferred real Ant/Canvas drafts. Synthetic composition does not accept the
// Windows native IME. No browser/dev/test process was started for this batch.
const fixturePath = '/__document-latex-templates-test__'
let server, browser, page, baseURL, errors
beforeAll(async () => {
  const fixture = await readFile(new URL('./fixtures/document-latex.html', import.meta.url), 'utf8')
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'document-latex-templates-fixture', configureServer(vite) { vite.middlewares.use(fixturePath, async (_request, response, next) => {
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
const formulaTab = () => page.getByRole('tab', { name: '公式', exact: true })
const roots = () => page.locator('[data-ribbon-item^="docs-latex-ui.ribbon.menu."]:visible')
const root = category => page.locator('[data-ribbon-item="docs-latex-ui.ribbon.menu.' + category + '"]:visible').getByRole('button').first()
const popup = () => page.locator('.ribbon-component-popup:visible')
const formulas = host => page.evaluate(host => window.documentLatexFixture.formulas(host), host)
const sourceRange = () => source().evaluate(element => [element.selectionStart, element.selectionEnd])
async function selectSource(start = 0, end) { await source().evaluate((element, range) => { element.focus(); element.setSelectionRange(range.start, range.end ?? element.value.length) }, { start, end }) }
async function editFormula(id = 'p', host) {
  const point = await page.evaluate(({ id, host }) => window.documentLatexFixture.point(id, 0.5, host), { id, host })
  await page.mouse.click(point.x, point.y, { clickCount: 2 }); await expect(source()).toBeFocused()
}

test('opening the formula editor activates all 11 contextual roots and closing restores the previous tab', async () => {
  await expect(formulaTab()).toHaveCount(0)
  await page.getByRole('tab', { name: '插入', exact: true }).click(); await editFormula()
  await expect(formulaTab()).toHaveAttribute('aria-selected', 'true'); await expect(roots()).toHaveCount(11)
  assert.deepEqual(await roots().evaluateAll(elements => elements.map(element => element.dataset.ribbonItem)), [
    'fractions', 'scripts', 'radicals', 'integrals', 'large-operators', 'brackets', 'functions', 'accents', 'limits-logarithms', 'operators', 'matrices',
  ].map(category => 'docs-latex-ui.ribbon.menu.' + category))
  await cancel().click(); await expect(formulaTab()).toHaveCount(0)
  await expect(page.getByRole('tab', { name: '插入', exact: true })).toHaveAttribute('aria-selected', 'true')
  await expect(page.getByRole('textbox', { name: '文档文字输入', exact: true })).toBeFocused()
})

test('the top Ribbon remains clickable with the Ant editor open and selected source replacement becomes one saved update', async () => {
  const original = (await formulas())[0]
  await editFormula(); await source().fill('x+y'); await selectSource()
  await root('fractions').click(); await expect(popup()).toBeVisible()
  await popup().locator('[data-latex-template="fraction-basic"]').click()
  const expected = '\\frac{x+y}{b}'
  await expect(source()).toHaveValue(expected); await expect(source()).toBeFocused()
  assert.deepEqual(await sourceRange(), [expected.length, expected.length])
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
  assert.equal((await formulas())[0].formula.latex, original.formula.latex)
  await apply().click(); await expect(panel()).toBeHidden()
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 1)
  await page.evaluate(() => window.documentLatexFixture.undo()); assert.equal((await formulas())[0].formula.latex, original.formula.latex)
  await page.evaluate(() => window.documentLatexFixture.redo()); await page.evaluate(() => window.documentLatexFixture.saveAndReopen())
  assert.equal((await formulas())[0].rangeId, original.rangeId); assert.equal((await formulas())[0].formula.latex, expected)
})

test('the Ant panel gallery selects an empty draft placeholder and typing replaces that placeholder only', async () => {
  await editFormula(); await source().fill('')
  await expect(apply()).toBeDisabled(); await panel().locator('[data-latex-template="fraction-basic"]').click()
  await expect(source()).toHaveValue('\\frac{a}{b}'); await expect(source()).toBeFocused()
  assert.deepEqual(await sourceRange(), [6, 7])
  await page.keyboard.insertText('q'); await expect(source()).toHaveValue('\\frac{q}{b}')
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
})

test('grid contextual selectors expose the complete category counts and dispatch actual source insertions', async () => {
  await load('?ribbon=grid'); await editFormula()
  const categories = ['fractions', 'scripts', 'radicals', 'integrals', 'large-operators', 'brackets', 'functions', 'accents', 'limits-logarithms', 'operators', 'matrices']
  const counts = [4, 4, 3, 6, 6, 8, 8, 8, 6, 12, 6]
  for (let index = 0; index < categories.length; index++) {
    await source().fill('x'); await selectSource(); await root(categories[index]).click()
    await expect(popup().locator('[data-latex-template]')).toHaveCount(counts[index])
    await popup().locator('[data-latex-template]').first().click()
    await expect(source()).not.toHaveValue('x'); await expect(source()).toBeFocused(); await expect(apply()).toBeEnabled()
  }
  assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
})

test('whole emoji source selection stays intact through API template insertion and unknown IDs keep the draft', async () => {
  await editFormula(); await source().fill('\\text{😀}'); await selectSource(6, 8)
  assert.equal(await page.evaluate(() => window.documentLatexFixture.insertTemplate('radical-square')), true)
  await expect(source()).toHaveValue('\\text{\\sqrt{😀}}'); await expect(source()).toBeFocused()
  const previous = await source().inputValue()
  assert.equal(await page.evaluate(() => window.documentLatexFixture.insertTemplate('missing')), false)
  await expect(source()).toHaveValue(previous); assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
})

test('synthetic composition disables template roots and API insertion until the final composition event', async () => {
  await editFormula(); await source().fill('x'); await selectSource()
  await source().evaluate(element => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await expect(root('fractions')).toBeDisabled(); await expect(panel().locator('[data-latex-template="fraction-basic"]')).toBeDisabled()
  assert.equal(await page.evaluate(() => window.documentLatexFixture.insertTemplate('radical-square')), false)
  await source().press('Escape'); await source().press('Control+Enter'); await expect(panel()).toBeVisible(); await expect(source()).toHaveValue('x')
  await source().evaluate(element => {
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'x' }))
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
  })
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(resolve)))
  assert.equal(await page.evaluate(() => window.documentLatexFixture.insertTemplate('radical-square')), false)
  await expect(root('fractions')).toBeDisabled()
  await source().evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'x' })))
  await expect(root('fractions')).toBeEnabled()
  assert.equal(await page.evaluate(() => window.documentLatexFixture.insertTemplate('radical-square')), true)
  await expect(source()).toHaveValue('\\sqrt{x}'); await expect(apply()).toBeEnabled()
})

test('a stale document selection disables all contextual source entry points and keeps its draft visible', async () => {
  await editFormula(); await source().fill('draft_2'); await page.evaluate(() => window.documentLatexFixture.select('tail', 1))
  await expect(root('fractions')).toBeDisabled(); await expect(panel().locator('[data-latex-template="fraction-basic"]')).toBeDisabled()
  assert.equal(await page.evaluate(() => window.documentLatexFixture.insertTemplate('fraction-basic')), false)
  await expect(source()).toHaveValue('draft_2'); await expect(apply()).toBeDisabled(); await expect(panel().getByRole('alert')).toContainText('选区已变化')
  await cancel().click(); assert.equal((await page.evaluate(() => window.documentLatexFixture.selection())).selection.focus.blockId, 'tail')
  assert.equal((await formulas())[0].formula.latex, '  x ^2  ')
})

test('table, header and note Canvas hosts apply template drafts to stable formulas and reopen with exact source', async () => {
  await load('?auxiliary')
  const header = { kind: 'story', sectionId: 'first', story: 'header' }, noteId = await page.evaluate(() => window.documentLatexFixture.noteId)
  const note = { kind: 'note', noteKind: 'footnote', noteId }, expected = []
  for (const target of [{ id: 'cell-p' }, { id: 'header-p', host: header }, { id: 'note-p', host: note }]) {
    if (target.host) await page.evaluate(target => window.documentLatexFixture.select(target.id, 0, 1, target.host), target)
    await editFormula(target.id, target.host)
    const original = await source().inputValue(), originalId = (await formulas(target.host)).find(item => item.blockId === target.id).rangeId
    await selectSource(); assert.equal(await page.evaluate(() => window.documentLatexFixture.insertTemplate('radical-square')), true)
    const next = '\\sqrt{' + original + '}'
    await expect(source()).toHaveValue(next); await apply().click(); await expect(panel()).toBeHidden()
    expected.push({ ...target, rangeId: originalId, source: next })
  }
  await page.evaluate(() => window.documentLatexFixture.undo()); await page.evaluate(() => window.documentLatexFixture.redo())
  await page.evaluate(() => window.documentLatexFixture.saveAndReopen())
  for (const target of expected) {
    const formula = (await formulas(target.host)).find(item => item.rangeId === target.rangeId)
    assert.ok(formula); assert.equal(formula.formula.latex, target.source)
  }
})

test('standalone template insertion uses the same source planner and snapshot replacement makes the draft read-only', async () => {
  await load('?standalone'); await editFormula(); await source().fill('x'); await selectSource()
  await panel().locator('[data-latex-template="fraction-basic"]').click(); await expect(source()).toHaveValue('\\frac{x}{b}')
  await page.evaluate(() => window.documentLatexFixture.replaceSameSnapshot())
  await expect(root('fractions')).toBeDisabled(); await expect(panel().locator('[data-latex-template="fraction-basic"]')).toBeDisabled()
  await expect(source()).toHaveValue('\\frac{x}{b}'); await expect(apply()).toBeDisabled()
  assert.equal((await formulas())[0].formula.latex, '  x ^2  ')
})

test('shared Canvas previews use theme tokens and a theme change preserves source selection and the draft', async () => {
  await editFormula(); await source().fill('\\frac{x}{y}'); await selectSource(6, 7)
  const preview = panel().locator('.latex-panel__preview canvas')
  const pixel = () => preview.evaluate(canvas => Array.from(canvas.getContext('2d').getImageData(0, 0, 1, 1).data).join(','))
  const light = await pixel(); await page.evaluate(() => window.documentLatexFixture.dark(true))
  await expect.poll(pixel).not.toBe(light); await expect(source()).toHaveValue('\\frac{x}{y}')
  assert.deepEqual(await sourceRange(), [6, 7]); assert.equal((await page.evaluate(() => window.documentLatexFixture.history())).undo, 0)
})
