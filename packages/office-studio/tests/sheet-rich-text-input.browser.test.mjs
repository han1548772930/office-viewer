import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Draft only: native events run through the real library controls when the user authorizes validation.
const fixturePath = '/__sheet-rich-text-input-test__'
const fixture = readFileSync(new URL('./fixtures/sheet-rich-text-input.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-rich-text-input-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
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
const snapshot = () => page.evaluate(() => window.richInputFixture.snapshot())
const history = () => page.evaluate(() => window.richInputFixture.history())
const attempts = () => page.evaluate(() => window.richInputFixture.attempts())
const sourceRuns = [['one', 'ab'], ['marker', ''], ['two', 'ab'], ['three', 'ab']]
const identities = cell => cell.richText.map(run => [run.id, run.text])
async function edit(address = 'A1') {
  await nameBox().fill(address); await nameBox().press('Enter'); await expect(grid()).toBeFocused()
  await grid().press('F2'); await expect(editor()).toBeFocused()
  return editor()
}
async function range(input, start, end = start) {
  await input.evaluate((element, value) => { element.focus(); element.setSelectionRange(value.start, value.end) }, { start, end })
}
async function compose(input, start, end, candidates, final = candidates.at(-1), trailing = true) {
  await input.evaluate((element, value) => {
    const before = element.value, prefix = before.slice(0, value.start), suffix = before.slice(value.end)
    element.setSelectionRange(value.start, value.end)
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    for (const data of value.candidates) {
      element.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertCompositionText', data, isComposing: true }))
      element.value = prefix + data + suffix; element.setSelectionRange(value.start + data.length, value.start + data.length)
      element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data, isComposing: true }))
    }
    if (value.final === '') { element.value = before; element.setSelectionRange(value.start, value.end) }
    else { element.value = prefix + value.final + suffix; element.setSelectionRange(value.start + value.final.length, value.start + value.final.length) }
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: value.final }))
    if (value.trailing) element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertFromComposition', data: value.final }))
  }, { start, end, candidates, final, trailing })
}

test('F2 insertion inside repeated text uses the real beforeinput interval and commits one rich transaction', async () => {
  const input = await edit()
  await range(input, 2); await input.pressSequentially('ab'); await input.press('Enter')
  await expect(editor()).toHaveCount(0); await expect(grid()).toBeFocused()
  const cell = (await snapshot()).sheets[0].cells['0:0']
  assert.deepEqual(identities(cell), [['one', 'abab'], ['marker', ''], ['two', 'ab'], ['three', 'ab']])
  assert.equal(cell.richText[0].font.bold, true); assert.equal(cell.richText[3].hyperlink, 'https://example.test')
  const changes = await attempts(); assert.equal(changes.length, 1)
  assert.equal(changes[0].editTrace.source, 'ababab'); assert.equal(changes[0].editTrace.operations[0].start, 2)
  assert.ok(changes[0].editTrace.operations.every(operation => operation.origin === 'beforeinput'))
  assert.equal((await history()).undo, 1); assert.equal((await history()).updates, 1)
})

test('deleting and retyping the same repeated string preserves the changed run affinity through content undo and redo', async () => {
  const input = await edit()
  await range(input, 2, 4); await input.press('Backspace'); await input.pressSequentially('ab')
  await expect(input).toHaveValue('ababab'); await input.press('Enter')
  const saved = await snapshot(), cell = saved.sheets[0].cells['0:0']
  assert.deepEqual(identities(cell), [['one', 'abab'], ['marker', ''], ['three', 'ab']])
  assert.equal(cell.richText[0].font.bold, true); assert.equal((await history()).undo, 1)
  await grid().press('Control+z'); assert.deepEqual(identities((await snapshot()).sheets[0].cells['0:0']), sourceRuns)
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).sheets[0].cells['0:0'], cell)
  await page.evaluate(value => window.richInputFixture.reopen(value), saved)
  await edit(); await expect(editor()).toHaveValue('ababab')
  assert.deepEqual(identities((await snapshot()).sheets[0].cells['0:0']), identities(cell)); await editor().press('Escape')
})

test('collapsed Backspace and Delete delete the intended repeated run rather than the matching suffix', async () => {
  let input = await edit(); await range(input, 4); await input.press('Backspace'); await input.press('Backspace'); await input.press('Enter')
  assert.deepEqual(identities((await snapshot()).sheets[0].cells['0:0']), [['one', 'ab'], ['marker', ''], ['three', 'ab']])
  await grid().press('Control+z')
  input = await edit(); await range(input, 2); await input.press('Delete'); await input.press('Delete'); await input.press('Enter')
  assert.deepEqual(identities((await snapshot()).sheets[0].cells['0:0']), [['one', 'ab'], ['marker', ''], ['three', 'ab']])
})

test('synthetic IME candidates keep the source range and run style with one group and no duplicate trailing input', async () => {
  const input = await edit()
  await compose(input, 2, 4, ['中', '中文']); await expect(input).toHaveValue('ab中文ab'); await input.press('Enter')
  const cell = (await snapshot()).sheets[0].cells['0:0']
  assert.deepEqual(identities(cell), [['one', 'ab'], ['marker', ''], ['two', '中文'], ['three', 'ab']])
  assert.equal(cell.richText[2].font.italic, true); assert.equal(cell.richText[3].vendor.keep, 3)
  const operations = (await attempts())[0].editTrace.operations
  assert.equal(operations.length, 1); assert.equal(operations[0].origin, 'composition')
  assert.deepEqual([operations[0].start, operations[0].end, operations[0].deletedText], [2, 4, 'ab'])
  assert.equal((await history()).undo, 1)
})

test('synthetic canceled IME restores source runs and composition Enter never commits a provisional value', async () => {
  const input = await edit()
  await input.evaluate(element => {
    element.setSelectionRange(2, 4)
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Enter', isComposing: true }))
  })
  assert.equal((await history()).undo, 0); await expect(input).toBeVisible()
  await input.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '' })))
  await compose(input, 2, 4, ['候', '候选'], '', false); await expect(input).toHaveValue('ababab')
  await input.press('Enter'); assert.deepEqual(identities((await snapshot()).sheets[0].cells['0:0']), sourceRuns)
  assert.equal((await history()).undo, 0); assert.equal((await attempts())[0].editTrace.operations.length, 0)
})

test('cell and formula inputs contribute to the same session while preserving native pointer selections', async () => {
  const input = await edit()
  await range(input, 2); await input.pressSequentially('X')
  await formula().focus(); await range(formula(), 3, 5); await formula().pressSequentially('Y'); await formula().press('Enter')
  const change = (await attempts())[0], cell = (await snapshot()).sheets[0].cells['0:0']
  assert.equal(change.editTrace.operations.length, 2); assert.equal(change.editTrace.sessionId, change.session.sessionId)
  assert.deepEqual(identities(cell), [['one', 'abX'], ['marker', ''], ['two', 'Y'], ['three', 'ab']])
  assert.equal(cell.richText[2].font.italic, true); assert.equal((await history()).undo, 1)
})

test('native draft undo and redo revisit recorded ranges before the final content history transaction', async () => {
  const input = await edit()
  await range(input, 2, 4); await input.pressSequentially('X'); await expect(input).toHaveValue('abXab')
  for (let index = 0; index < 3 && await input.inputValue() !== 'ababab'; index++) await input.press('Control+z')
  await expect(input).toHaveValue('ababab'); assert.equal((await history()).undo, 0)
  await input.press('Control+y'); await expect(input).toHaveValue('abXab'); await input.press('Enter')
  const cell = (await snapshot()).sheets[0].cells['0:0']; assert.equal(cell.richText.find(run => run.id === 'two').font.italic, true)
  assert.equal((await history()).undo, 1)
})

test('forced-string escapes, text formats and supplementary characters retain their raw rich-run offsets', async () => {
  let input = await edit('B1'); await expect(input).toHaveValue("'ababab")
  await range(input, 3); await input.pressSequentially('X'); await input.press('Enter')
  let cell = (await snapshot()).sheets[0].cells['0:1']
  assert.equal(cell.value, 'abXabab'); assert.equal(cell.dataType, 'force_string'); assert.equal(cell.richText[0].text, 'abX')
  input = await edit('C1'); await range(input, 0); await input.pressSequentially("'"); await input.press('Enter')
  cell = (await snapshot()).sheets[0].cells['0:2']; assert.equal(cell.value, "'ababab"); assert.equal(cell.richText[0].text, "'ab")
  input = await edit('D1'); await compose(input, 1, 3, ['👩', '👩🏽‍💻']); await input.press('Enter')
  cell = (await snapshot()).sheets[0].cells['0:3']; assert.equal(cell.value, 'A👩🏽‍💻B')
  assert.equal(cell.richText.find(run => run.id === 'emoji').font.italic, true)
})

test('multiline paste and Alt+Enter record programmatic ranges through the AutoComplete to TextArea swap', async () => {
  const input = await edit(); await range(input, 2, 4)
  await input.evaluate(element => {
    const data = new DataTransfer(); data.setData('text/plain', 'X\r\nY')
    element.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data }))
  })
  await expect(page.locator('.canvas-cell-input textarea')).toBeFocused(); await expect(editor()).toHaveValue('abX\nYab')
  await editor().press('Alt+Enter'); await expect(editor()).toHaveValue('abX\nY\nab'); await editor().press('Enter')
  const cell = (await snapshot()).sheets[0].cells['0:0']
  assert.equal(cell.richText.find(run => run.id === 'two').text, 'X\nY\n')
  assert.equal(cell.richText.find(run => run.id === 'two').font.italic, true)
  assert.deepEqual((await attempts())[0].editTrace.operations.map(({ inputType, origin }) => [inputType, origin]),
    [['insertFromPaste', 'programmatic'], ['insertLineBreak', 'programmatic']])
})

test('a canceled public command retains the exact draft trace for retry and same-text external rich changes conflict', async () => {
  const input = await edit(); await range(input, 2, 4); await input.pressSequentially('X')
  await page.evaluate(() => window.richInputFixture.cancel(true)); await input.press('Enter')
  await expect(input).toHaveValue('abXab'); assert.equal((await history()).undo, 0)
  await page.evaluate(() => window.richInputFixture.cancel(false)); await input.press('Enter')
  const changes = await attempts(); assert.equal(changes.length, 2); assert.deepEqual(changes[0].editTrace, changes[1].editTrace)
  await grid().press('Control+z'); const again = await edit(); await range(again, 2); await again.pressSequentially('X')
  await page.evaluate(() => window.richInputFixture.mutateRich()); await again.press('Enter')
  await expect(again).toHaveValue('abXabab'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'ababab')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].richText[1].vendor.keep, false)
  assert.ok((await page.evaluate(() => window.richInputFixture.notifications())).some(item => item.message.includes('文字运行')))
  await again.press('Escape')
})

test('old native host events cannot overwrite a replacement studio session with the same workbook and cell IDs', async () => {
  let input = await edit(); await range(input, 2); await input.pressSequentially('X')
  await page.evaluate(() => window.richInputFixture.remember())
  await page.evaluate(() => window.richInputFixture.replaceHost('stale-connected')); await expect(editor()).toHaveCount(0)
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'ababab'); assert.equal((await attempts()).length, 0)
  input = await edit(); await range(input, 2); await input.pressSequentially('Y')
  await page.evaluate(() => window.richInputFixture.oldInput('stale'))
  await expect(input).toHaveValue('abYabab'); await input.press('Enter')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'abYabab')
  assert.equal((await history()).undo, 1)
})

test('validation AutoComplete choices use an explicit whole-input replacement instead of a value-only inferred update', async () => {
  await page.evaluate(() => window.richInputFixture.validation()); await edit()
  const choice = page.locator('.ant-select-item-option-content').filter({ hasText: /^choice$/u })
  await expect(choice).toBeVisible(); await choice.click(); await expect(editor()).toHaveValue('choice'); await editor().press('Enter')
  const operations = (await attempts()).at(-1).editTrace.operations
  assert.ok(operations.some(edit => edit.origin === 'programmatic' && edit.inputType === 'insertReplacementText'
    && edit.start === 0 && edit.end === 6 && edit.text === 'choice'))
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'choice')
})

test('the same input bridge works in standalone mode and keeps theme, focus and library-owned overlays', async () => {
  await page.goto(baseURL + '?standalone'); await expect(grid()).toBeVisible()
  await page.evaluate(() => window.richInputFixture.appearance(true, 'grid'))
  const input = await edit(); await range(input, 2, 4); await input.press('Delete'); await input.pressSequentially('ab'); await input.press('Enter')
  assert.deepEqual(identities((await snapshot()).sheets[0].cells['0:0']), [['one', 'abab'], ['marker', ''], ['three', 'ab']])
  assert.equal((await history()).updates, 1); await expect(grid()).toBeFocused()
  const darkSurface = await grid().evaluate(element => getComputedStyle(element).backgroundColor)
  const library = await formula().evaluate(element => element.className)
  assert.match(library, /ant-input/u)
  await page.evaluate(() => window.richInputFixture.appearance(false, 'classic'))
  assert.notEqual(await grid().evaluate(element => getComputedStyle(element).backgroundColor), darkSurface)
  const again = await edit(); await again.press('Escape'); await expect(grid()).toBeFocused()
})
