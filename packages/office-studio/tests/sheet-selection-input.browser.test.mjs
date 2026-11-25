import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet-selection-input-test__'
const fixture = readFileSync(new URL('./fixtures/sheet-selection-input.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-selection-input-fixture',
    configureServer(vite) {
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
const snapshot = id => page.evaluate(value => window.selectionInputFixture.snapshot(value), id)
const selection = () => page.evaluate(() => window.selectionInputFixture.selection())
const history = id => page.evaluate(value => window.selectionInputFixture.history(value), id)
const rangeCoordinates = state => state.ranges.map(({ startRow, endRow, startColumn, endColumn }) => [startRow, endRow, startColumn, endColumn])
async function select(address) {
  await nameBox().fill(address); await nameBox().press('Enter'); await expect(grid()).toBeFocused()
}
async function edit(address) {
  await select(address); await grid().press('F2'); await expect(editor()).toBeFocused()
  return editor()
}
async function selectMultiple() {
  await page.evaluate(() => window.selectionInputFixture.select(['A1:C2', 'F6:G7'],
    { activeRangeIndex: 0, primary: { row: 0, column: 1 } }))
  await grid().focus()
  await expect(formula()).toHaveValue('4')
}
async function point(row, column, axis) {
  const rect = await page.evaluate(({ row, column }) => window.selectionInputFixture.rect(row, column, false), { row, column })
  const bounds = await grid().boundingBox()
  return { x: axis === 'row' ? bounds.x + 20 : rect.x + rect.width / 2,
    y: axis === 'column' ? bounds.y + 13 : rect.y + rect.height / 2 }
}

test('API multiple selections project into the Canvas and Enter/Tab traverse them without content history', async () => {
  await selectMultiple()
  const before = await selection(), depth = await history()
  assert.deepEqual(rangeCoordinates(before), [[0, 1, 0, 2], [5, 6, 5, 6]])
  await grid().press('Tab')
  assert.deepEqual((await selection()).primary, { row: 0, column: 2 })
  await grid().press('Enter')
  assert.deepEqual((await selection()).primary, { row: 1, column: 2 })
  await grid().press('Shift+Tab'); await grid().press('Shift+Enter')
  assert.deepEqual((await selection()).primary, { row: 0, column: 1 })
  assert.deepEqual((await selection()).ranges, before.ranges)
  assert.deepEqual(await history(), depth); await expect(editor()).toHaveCount(0)
})

test('a primary cell edit commits once, preserves multiple selections in history and reopens its typed value', async () => {
  await selectMultiple()
  const before = await selection()
  await grid().press('F2'); await editor().fill('12.5%'); await editor().press('Tab')
  await expect(editor()).toHaveCount(0); await expect(grid()).toBeFocused()
  const saved = await snapshot()
  assert.equal(saved.sheets[0].cells['0:1'].value, .125)
  assert.deepEqual((await selection()).ranges, before.ranges); assert.deepEqual((await selection()).primary, { row: 0, column: 2 })
  assert.equal((await history()).undo, 1); assert.equal((await history()).emits, 1)
  const commands = await page.evaluate(() => window.selectionInputFixture.commands())
  assert.ok(commands.some(command => command.id === 'sheet.command.commit-cell-input'))
  await grid().press('Control+z')
  assert.equal((await snapshot()).sheets[0].cells['0:1'].value, 4); assert.deepEqual(await selection(), before)
  await grid().press('Control+y'); assert.equal((await snapshot()).sheets[0].cells['0:1'].value, .125)
  assert.deepEqual(await selection(), before)
  await page.evaluate(value => window.selectionInputFixture.reopen(value), saved)
  const input = await edit('B1'); await expect(input).toHaveValue('12.5%'); await input.press('Escape')
})

test('rejected and canceled input keeps its draft, selection and Tab origin until Escape', async () => {
  await page.evaluate(() => window.selectionInputFixture.seed({ entries: [{ value: 5 }] }))
  await page.evaluate(() => window.selectionInputFixture.validation('stop'))
  let input = await edit('A1')
  const before = await selection(), depth = await history()
  await input.fill('30'); await input.press('Enter'); await input.press('Tab')
  await expect(input).toHaveValue('30'); await expect(input).toBeFocused()
  assert.deepEqual(await selection(), before); assert.deepEqual(await history(), depth)
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 5)
  await input.press('Escape'); await expect(grid()).toBeFocused(); await expect(formula()).toHaveValue('5')
  await select('B2'); await grid().press('Tab')
  const origin = await page.evaluate(() => window.selectionInputFixture.tabOrigin())
  input = editor(); await grid().press('F2'); await input.fill('99')
  await page.evaluate(() => window.selectionInputFixture.cancel(true))
  await input.press('Enter'); await expect(input).toHaveValue('99')
  assert.deepEqual(await page.evaluate(() => window.selectionInputFixture.tabOrigin()), origin)
  assert.deepEqual((await selection()).primary, { row: 1, column: 2 })
  assert.equal((await history()).undo, 0)
  await input.press('Escape'); await expect(grid()).toBeFocused()
})

test('warning validation commits the typed input with the actual warning and one content entry', async () => {
  await page.evaluate(() => window.selectionInputFixture.seed({ entries: [{ value: 5 }] }))
  await page.evaluate(() => window.selectionInputFixture.validation('warning'))
  const input = await edit('A1'); await input.fill('30'); await input.press('Tab')
  await expect(editor()).toHaveCount(0)
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 30); assert.equal((await history()).undo, 1)
  const notices = await page.evaluate(() => window.selectionInputFixture.notifications())
  assert.ok(notices.some(item => item.type === 'warning' && item.message.includes('整数范围')))
  assert.deepEqual((await selection()).primary, { row: 0, column: 1 })
})

test('composition in either Ant Input bridge does not commit until composition ends', async () => {
  await page.evaluate(() => window.selectionInputFixture.seed({ entries: [{ value: '甲' }] }))
  let input = await edit('A1')
  await input.dispatchEvent('compositionstart', { data: '中' }); await input.fill('中文')
  await input.press('Enter'); assert.equal((await history()).undo, 0); await expect(input).toHaveValue('中文')
  await input.dispatchEvent('compositionend', { data: '中文' }); await input.press('Enter')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '中文'); assert.equal((await history()).undo, 1)
  await select('A1'); input = formula(); await input.focus(); await input.fill('组合文本')
  await input.dispatchEvent('compositionstart', { data: '文本' })
  await expect(page.getByRole('button', { name: '确认编辑', exact: true })).toBeDisabled()
  await input.press('Enter'); assert.equal((await history()).undo, 1)
  await input.dispatchEvent('compositionend', { data: '组合文本' })
  await input.press('Enter'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, '组合文本')
  assert.equal((await history()).undo, 2)
})

test('editing a visible representative of a hidden merge owner updates only its persisted main cell', async () => {
  await page.evaluate(() => window.selectionInputFixture.seed({
    entries: [{ value: 'owner' }], merges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }],
    rowHiddenSpans: [{ start: 0, end: 0, value: true }], columnHiddenSpans: [{ start: 0, end: 0, value: true }],
  }))
  const input = await edit('B2')
  await expect(input).toHaveValue('owner'); await expect(formula()).toHaveValue('owner')
  const before = await selection()
  assert.deepEqual(before.primary, { row: 1, column: 1 })
  await input.fill('owner-new'); await input.press('Tab')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'owner-new')
  assert.equal((await snapshot()).sheets[0].cells['1:1'], undefined)
  const next = await selection(); assert.ok(next.primary.row >= 1 && next.primary.column >= 2)
  await grid().press('Control+z'); assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'owner')
  assert.deepEqual(await selection(), before)
})

test('header Ctrl/Shift selection and frozen-pane pointer hits preserve the full API state without content edits', async () => {
  await page.evaluate(() => window.selectionInputFixture.seed({ freeze: { rows: 2, columns: 2 } }))
  const before = await history()
  let target = await point(3, 0, 'row'); await page.mouse.click(target.x, target.y)
  await page.keyboard.down('Control'); target = await point(5, 0, 'row'); await page.mouse.click(target.x, target.y); await page.keyboard.up('Control')
  await page.keyboard.down('Shift'); target = await point(7, 0, 'row'); await page.mouse.click(target.x, target.y); await page.keyboard.up('Shift')
  assert.deepEqual(rangeCoordinates(await selection()), [[3, 3, 0, 25], [5, 7, 0, 25]])
  assert.ok((await selection()).ranges.every(range => range.rangeType === 1))
  const frozen = await page.evaluate(() => window.selectionInputFixture.rect(0, 0))
  await page.locator('.canvas-scroll').evaluate(element => { element.scrollTop = 260; element.scrollLeft = 180; element.dispatchEvent(new Event('scroll')) })
  assert.deepEqual(await page.evaluate(() => window.selectionInputFixture.rect(0, 0)), frozen)
  await page.mouse.click(frozen.x + frozen.width / 2, frozen.y + frozen.height / 2)
  assert.deepEqual((await selection()).primary, { row: 0, column: 0 })
  assert.deepEqual(await history(), before)
})

test('a same-sheet observer selection survives input confirmation and a workbook switch keeps the committed edit in its own history', async () => {
  let input = await edit('A1')
  await input.fill('changed'); await page.evaluate(() => window.selectionInputFixture.observer('reselect'))
  await input.press('Enter'); await expect(editor()).toHaveCount(0); await expect(grid()).toBeFocused()
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 'changed')
  assert.deepEqual((await selection()).primary, { row: 4, column: 4 }); assert.equal((await history()).undo, 1)
  input = await edit('A1'); await input.fill('second edit')
  await page.evaluate(() => window.selectionInputFixture.observer('switch')); await input.press('Tab')
  await expect(editor()).toHaveCount(0)
  assert.equal((await snapshot()).id, 'selection-input-neighbor')
  assert.deepEqual((await selection()).primary, { row: 3, column: 3 })
  assert.equal((await snapshot('selection-input-fixture')).sheets[0].cells['0:0'].value, 'second edit')
  assert.equal((await history('selection-input-fixture')).undo, 2)
  assert.equal((await history('selection-input-neighbor')).undo, 0)
})

test('rich-text edits retain source entities through undo/reopen and direct typing replaces those runs', async () => {
  const before = await snapshot(), input = await edit('A1')
  await input.fill('aXcdef'); await input.press('Enter')
  const saved = await snapshot(), runs = saved.sheets[0].cells['0:0'].richText
  assert.deepEqual(runs.map(run => [run.id, run.text]), [['first-run', 'aX'], ['empty-marker', ''], ['second-run', 'cd'], ['last-run', 'ef']])
  await grid().press('Control+z')
  assert.deepEqual((await snapshot()).sheets[0].cells['0:0'].richText, before.sheets[0].cells['0:0'].richText)
  await page.evaluate(value => window.selectionInputFixture.reopen(value), saved)
  assert.deepEqual((await snapshot()).sheets[0].cells['0:0'].richText, runs)
  await select('A1'); await grid().press('z'); await expect(editor()).toHaveValue('z'); await editor().press('Enter')
  assert.equal((await snapshot()).sheets[0].cells['0:0'].richText, undefined)
})

test('dark/light and classic/grid transitions preserve Ant editor focus, draft, model and content history', async () => {
  const input = await edit('A1'); await input.fill('待提交')
  const model = await snapshot(), depth = await history()
  for (const [dark, variant] of [[true, 'grid'], [false, 'classic']]) {
    await page.evaluate(([dark, variant]) => window.selectionInputFixture.appearance(dark, variant), [dark, variant])
    await expect(input).toBeFocused(); await expect(input).toHaveValue('待提交')
    await expect(page.locator('.canvas-cell-input .ant-input')).toHaveCount(1)
    assert.deepEqual(await snapshot(), model); assert.deepEqual(await history(), depth)
  }
  await input.press('Escape'); await expect(grid()).toBeFocused()
})

test('standalone input uses the same typed mapping and emits one authoritative workbook update', async () => {
  await page.goto(baseURL + '?standalone'); await expect(grid()).toBeVisible()
  const input = await edit('B1'); await input.fill('12.5%'); await input.press('Enter')
  assert.equal((await snapshot()).sheets[0].cells['0:1'].value, .125)
  assert.match((await snapshot()).sheets[0].cells['0:1'].style.numberFormat, /%/)
  assert.equal((await history()).emits, 1); assert.equal((await history()).undo, 0)
})

async function pressEditingFillHandle() {
  const handle = await page.locator('.canvas-cell-fill-handle').boundingBox()
  assert.ok(handle && handle.width > 0 && handle.height > 0)
  // The inner half overlays the Ant Input and was previously ignored by the canvas.
  await page.mouse.move(handle.x + 1, handle.y + 1)
  await page.mouse.down()
}
async function moveFillTo(row, column) {
  const rect = await page.evaluate(({ row, column }) => window.selectionInputFixture.rect(row, column, false), { row, column })
  await page.mouse.move(rect.x + rect.width / 2, rect.y + rect.height / 2, { steps: 4 })
}

test('editing handle commits the draft and fills on the same pointer, with separate undo entries and snapshot reopen', async () => {
  const original = await snapshot(), input = await edit('A1')
  await input.fill('10'); await pressEditingFillHandle()
  await expect(editor()).toHaveCount(0); await expect(grid()).toBeFocused()
  await moveFillTo(3, 0); await page.mouse.up()
  await expect.poll(async () => (await snapshot()).sheets[0].cells['3:0']?.value).toBe(10)
  const saved = await snapshot()
  assert.deepEqual([0, 1, 2, 3].map(row => saved.sheets[0].cells[row + ':0'].value), [10, 10, 10, 10])
  assert.equal((await history()).undo, 2)
  await grid().press('Control+z')
  let restored = await snapshot()
  assert.equal(restored.sheets[0].cells['0:0'].value, 10)
  assert.equal(restored.sheets[0].cells['1:0']?.value, undefined)
  await grid().press('Control+z')
  restored = await snapshot()
  assert.deepEqual(restored.sheets[0].cells['0:0'], original.sheets[0].cells['0:0'])
  await grid().press('Control+y'); await grid().press('Control+y')
  assert.equal((await snapshot()).sheets[0].cells['3:0'].value, 10)
  await page.evaluate(value => window.selectionInputFixture.reopen(value), saved)
  const reopened = await snapshot()
  assert.deepEqual([0, 1, 2, 3].map(row => saved.sheets[0].cells[row + ':0'].value),
    [0, 1, 2, 3].map(row => reopened.sheets[0].cells[row + ':0'].value))
})

test('clicking either half of the editing handle confirms input without filling adjacent cells', async () => {
  for (const offset of [1, 4]) {
    await page.evaluate(() => window.selectionInputFixture.seed({ entries: [{ value: 5 }] }))
    const input = await edit('A1'); await input.fill('9')
    const handle = await page.locator('.canvas-cell-fill-handle').boundingBox()
    assert.ok(handle)
    await page.mouse.click(handle.x + offset, handle.y + offset)
    await expect(editor()).toHaveCount(0)
    const model = await snapshot()
    assert.equal(model.sheets[0].cells['0:0'].value, 9)
    assert.equal(model.sheets[0].cells['1:0']?.value, undefined)
    assert.equal(model.sheets[0].cells['0:1']?.value, undefined)
    assert.equal(model.sheets[0].cells['1:1']?.value, undefined)
    assert.equal((await history()).undo, 1)
  }
})

test('editing fill reads the committed formula and translates its relative references at the destination', async () => {
  await page.evaluate(() => window.selectionInputFixture.seed({ entries: [
    { value: 1 }, { column: 1, value: 2 }, { row: 1, column: 1, value: 3 }, { row: 2, column: 1, value: 4 },
  ] }))
  const input = await edit('A1'); await input.fill('=B1*2'); await pressEditingFillHandle()
  await expect(editor()).toHaveCount(0); await moveFillTo(2, 0); await page.mouse.up()
  const model = await snapshot()
  assert.deepEqual([0, 1, 2].map(row => model.sheets[0].cells[row + ':0'].formula), ['=B1*2', '=B2*2', '=B3*2'])
  assert.equal((await history()).undo, 2)
})

for (const rejection of ['validation', 'command-cancel', 'composition']) {
  test('editing fill preserves draft and model when blocked by ' + rejection, async () => {
    await page.evaluate(() => window.selectionInputFixture.seed({ entries: [{ value: 5 }] }))
    if (rejection === 'validation') await page.evaluate(() => window.selectionInputFixture.validation('stop'))
    const input = await edit('A1'), before = await selection()
    if (rejection === 'composition') await input.dispatchEvent('compositionstart', { data: '三' })
    await input.fill('30')
    if (rejection === 'command-cancel') await page.evaluate(() => window.selectionInputFixture.cancel(true))
    await pressEditingFillHandle(); await moveFillTo(3, 0); await page.mouse.up()
    await expect(input).toHaveValue('30'); await expect(input).toBeFocused()
    const model = await snapshot()
    assert.equal(model.sheets[0].cells['0:0'].value, 5)
    assert.equal(model.sheets[0].cells['3:0']?.value, undefined)
    assert.deepEqual(await selection(), before); assert.equal((await history()).undo, 0)
  })
}

test('observer reselection during edit confirmation cancels the fill and preserves the observer selection', async () => {
  const input = await edit('A1'); await input.fill('7')
  await page.evaluate(() => window.selectionInputFixture.observer('reselect'))
  await pressEditingFillHandle(); await moveFillTo(3, 0); await page.mouse.up()
  await expect(editor()).toHaveCount(0)
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 7)
  assert.equal((await snapshot()).sheets[0].cells['3:0']?.value, undefined)
  assert.deepEqual((await selection()).primary, { row: 4, column: 4 })
  assert.equal((await history()).undo, 1)
})

test('Escape cancels an active fill while preserving its already confirmed source edit', async () => {
  const input = await edit('A1'); await input.fill('9'); await pressEditingFillHandle()
  await expect(editor()).toHaveCount(0); await moveFillTo(3, 0)
  await grid().press('Escape'); await page.mouse.up()
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 9)
  assert.equal((await snapshot()).sheets[0].cells['3:0']?.value, undefined)
  assert.equal((await history()).undo, 1)
})

test('selection changes after a fill begins invalidate its pointer before destination writes', async () => {
  const input = await edit('A1'); await input.fill('11'); await pressEditingFillHandle()
  await expect(editor()).toHaveCount(0)
  await page.evaluate(() => window.selectionInputFixture.select(['C7']))
  await moveFillTo(3, 0); await page.mouse.up()
  assert.equal((await snapshot()).sheets[0].cells['0:0'].value, 11)
  assert.equal((await snapshot()).sheets[0].cells['3:0']?.value, undefined)
  assert.deepEqual((await selection()).primary, { row: 6, column: 2 })
  assert.equal((await history()).undo, 1)
})
