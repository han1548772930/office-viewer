import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'
import { COMMIT_SHEET_DATE_INPUT_COMMAND } from '../src/features/sheet-validation/validation-date-input.ts'
import { COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND } from '../src/features/sheet-validation/validation-list-input.ts'
import { SHEET_DATA_VALIDATION_COMMANDS } from '../src/features/sheet-validation/validation-rule-command.ts'

// Implementation-phase browser drafts. No Vite server, Edge process or test runner has been started.
// Synthetic composition events exercise event guards; they do not establish Windows IME acceptance.
const route = '/__sheet-validation-cell__', fixture = readFileSync(new URL('./fixtures/sheet-validation-cell.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-validation-cell-fixture', configureServer(vite) {
      vite.middlewares.use(route, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(route, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); baseURL = `http://127.0.0.1:${server.httpServer.address().port}${route}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1600, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const datePopup = (key = 'A1') => page.getByRole('dialog', { name: `选择日期或时间 ${key}`, exact: true })
const listPopup = (key = 'B1') => page.getByRole('dialog', { name: `选择列表项 ${key}`, exact: true })
const date = (key = 'A1') => datePopup(key).getByRole('textbox', { name: '日期', exact: true })
const time = key => datePopup(key).getByRole('textbox', { name: '时间', exact: true })
const confirm = key => datePopup(key).getByRole('button', { name: '确定', exact: true })
const snapshot = () => page.evaluate(() => window.validationCellFixture.snapshot())
const history = () => page.evaluate(() => window.validationCellFixture.history())
const panel = () => page.getByRole('complementary', { name: '数据验证', exact: true })
const value = async (row, column = 0) => (await snapshot()).sheets.find(sheet => sheet.id === 'data').cells[`${row}:${column}`]
async function select(address) { await nameBox().fill(address); await nameBox().press('Enter') }
async function appearance(variant, dark = false) { await page.evaluate(({ variant, dark }) => window.validationCellFixture.appearance(dark, variant), { variant, dark }) }
async function openDate(key) { await select(key); await page.getByRole('button', { name: `选择日期或时间 ${key}`, exact: true }).click(); await expect(datePopup(key)).toBeVisible() }
async function openList(key) { await select(key); await page.getByRole('button', { name: `选择列表项 ${key}`, exact: true }).click(); await expect(listPopup(key)).toBeVisible() }
async function changePicker(input, text) { await input.fill(text); await input.press('Enter'); await input.press('Tab') }
async function closeList(key) { await listPopup(key).getByRole('button', { name: '关闭', exact: true }).click(); await expect(listPopup(key)).toBeHidden() }
async function manageRules() {
  await page.getByRole('tab', { name: '数据', exact: true }).click(); await page.getByRole('button', { name: '数据验证', exact: true }).click()
  await page.getByRole('menuitem', { name: '管理验证规则', exact: true }).click(); await expect(panel()).toBeVisible()
}
async function editDateRule() { await openDate('A1'); await datePopup().getByRole('button', { name: '编辑数据验证', exact: true }).click(); await expect(datePopup()).toBeHidden(); await expect(panel()).toBeVisible() }

// Long-list scenarios require a generated protocol-2 artifact with
// create_validation_list_query. A missing build is a failure, never a JS skip.
for (const [variant, dark] of [['classic', false], ['grid', true]]) {
  test(`${variant}/dark=${dark}: hidden-search Rust lists gain initial keyboard focus and Enter commits virtual choices in source order`, async () => {
    await appearance(variant, dark)
    await page.evaluate(() => window.validationCellFixture.setDropdownUi({ showSearchOnDropdown: false, showEditOnDropdown: false }))
    const source = await page.evaluate(() => window.validationCellFixture.longList())
    await openList('B1'); const popup = listPopup(), first = popup.getByRole('checkbox', { name: source.first, exact: true }), last = popup.getByRole('checkbox', { name: source.last, exact: true })
    await expect(popup.getByRole('textbox', { name: '搜索列表项', exact: true })).toHaveCount(0)
    await expect(popup.getByRole('button', { name: '编辑数据验证', exact: true })).toHaveCount(0)
    await expect(first).toBeFocused(); assert.ok(await popup.getByRole('checkbox').count() < 80)
    await first.press('End'); await expect(last).toBeFocused(); await last.press('Enter')
    await expect(last).toBeChecked(); await expect(popup).toBeVisible(); assert.equal((await value(0, 1)).value, source.last)
    await last.press('Home'); await expect(first).toBeFocused(); await first.press('Enter')
    await expect(first).toBeChecked(); assert.equal((await value(0, 1)).value, `${source.first},${source.last}`); assert.equal((await history()).undo, 2)
    await closeList('B1'); await expect(grid()).toBeFocused()
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await value(0, 1)).value, source.last)
    await page.evaluate(() => window.validationCellFixture.redo()); await page.evaluate(() => window.validationCellFixture.reopen())
    await openList('B1'); await expect(first).toBeChecked(); await first.press('End'); await expect(last).toBeChecked()
    assert.equal((await value(0, 1)).value, `${source.first},${source.last}`)
  }, { timeout: 60_000 })

  test(`${variant}/dark=${dark}: real Rust query plus virtual End/Home selection preserves color, history and internal reopen`, async () => {
    await appearance(variant, dark)
    const source = await page.evaluate(() => window.validationCellFixture.longList())
    await openList('B1'); const popup = listPopup(), first = popup.getByRole('checkbox', { name: source.first, exact: true }), last = popup.getByRole('checkbox', { name: source.last, exact: true })
    await expect(first).toBeVisible(); assert.ok(await popup.getByRole('checkbox').count() < 80)
    await expect(first.locator('xpath=ancestor::*[@data-validation-list-index][1]').locator('.ant-tag')).toHaveCSS('background-color', 'rgb(255, 0, 0)')
    await first.focus(); await first.press('End'); await expect(last).toBeFocused(); assert.ok(await popup.getByRole('checkbox').count() < 80)
    await last.check(); await expect(popup).toBeVisible(); assert.equal((await value(0, 1)).value, source.last)
    await last.press('Home'); await expect(first).toBeFocused(); await first.check()
    assert.equal((await value(0, 1)).value, `${source.first},${source.last}`); assert.equal((await history()).undo, 2)
    const records = await page.evaluate(() => window.validationCellFixture.listQueries()), result = records.find(record => record.result)
    assert.equal(result.params.rowCount, source.count); assert.equal(result.metrics.boundary.jsonConversions, 0)
    assert.equal(result.metrics.preparedInputCopies, 3); assert.equal(result.metrics.preparedInputBytes, result.inputBytes)
    assert.equal(result.metrics.boundary.inputCopies, 2); assert.equal(result.metrics.boundary.outputCopies, 1)
    await closeList('B1'); await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await value(0, 1)).value, source.last)
    await page.evaluate(() => window.validationCellFixture.redo()); await page.evaluate(() => window.validationCellFixture.reopen())
    await openList('B1'); await expect(first).toBeChecked(); await first.focus(); await first.press('End'); await expect(last).toBeChecked()
    assert.equal((await value(0, 1)).value, `${source.first},${source.last}`)
  }, { timeout: 60_000 })

  test(`${variant}/dark=${dark}: composition keeps stable search rows and final text selects a result across the virtual window`, async () => {
    await appearance(variant, dark); const source = await page.evaluate(() => window.validationCellFixture.longList())
    await openList('B1'); const popup = listPopup(), search = popup.getByRole('textbox', { name: '搜索列表项', exact: true })
    const first = popup.getByRole('checkbox', { name: source.first, exact: true }), last = popup.getByRole('checkbox', { name: source.last, exact: true })
    await expect(first).toBeVisible(); await search.fill('候选 000'); await expect(first).toBeVisible()
    const firstRow = first.locator('xpath=ancestor::*[@data-validation-list-index][1]')
    await expect(firstRow.locator('[data-validation-list-highlight]')).toHaveText('候选 000')
    await search.evaluate(input => {
      input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
      input.value = '汉字'; input.dispatchEvent(new InputEvent('input', { bubbles: true, data: '汉字', isComposing: true }))
    })
    await expect(first).toBeVisible(); await expect(first).toBeDisabled(); await expect(last).toHaveCount(0)
    await expect(firstRow.locator('[data-validation-list-highlight]')).toHaveText('候选 000')
    await search.press('Escape'); await expect(popup).toBeVisible(); assert.equal((await history()).undo, 0)
    await search.evaluate(input => input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '汉字' })))
    await expect(search).toHaveValue('汉字'); await expect(last).toBeVisible(); await expect(first).toHaveCount(0)
    await expect(popup.locator('[data-validation-list-highlight]')).toHaveText('汉字')
    await expect(popup.locator('[data-validation-list-highlight]')).toHaveCSS('font-weight', '600')
    await last.check(); assert.equal((await value(0, 1)).value, source.last); assert.equal((await history()).undo, 1)
    const queries = await page.evaluate(() => window.validationCellFixture.listQueries())
    assert.ok(queries.some(query => query.result && query.params.queryByteLength === 4))
  }, { timeout: 60_000 })
}

test('a real Worker progress callback switching sheets cancels the old query without changing content/history or reopening its popup', async () => {
  await page.evaluate(() => { window.validationCellFixture.longList(16_384); window.validationCellFixture.switchAtListProgress() })
  await select('B1'); await page.getByRole('button', { name: '选择列表项 B1', exact: true }).click()
  await expect(page.getByRole('tab', { name: 'Other', exact: true })).toHaveAttribute('aria-selected', 'true')
  await expect(listPopup()).toBeHidden(); assert.equal((await value(0, 1)).value, ''); assert.equal((await history()).undo, 0)
  await expect.poll(async () => (await page.evaluate(() => window.validationCellFixture.listQueries())).some(query => query.error)).toBe(true)
  const records = await page.evaluate(() => window.validationCellFixture.listQueries())
  assert.equal(records.some(query => query.result), false); assert.equal(records.some(query => query.aborted), true)
  assert.ok(records.some(query => ['COMPUTE_STALE_RESULT', 'COMPUTE_CANCELLED'].includes(query.error)))
}, { timeout: 60_000 })

for (const variant of ['classic', 'grid']) {
  test(`${variant}: live search/edit visibility keeps composition, resets the hidden query and returns focus to a legal Ant control`, async () => {
    await appearance(variant); await openList('B1')
    const popup = listPopup(), search = popup.getByRole('textbox', { name: '搜索列表项', exact: true })
    await search.fill('Green'); await expect(popup.getByRole('checkbox', { name: 'Green, dark', exact: true })).toBeVisible()
    const before = await snapshot(), oldHistory = await history()
    await search.dispatchEvent('compositionstart', { data: '绿' })
    await page.evaluate(() => window.validationCellFixture.setDropdownUi({ showSearchOnDropdown: false, showEditOnDropdown: false }))
    await expect(search).toBeVisible(); await expect(search).toHaveValue('Green')
    await expect(popup.getByRole('button', { name: '编辑数据验证', exact: true })).toHaveCount(0)
    await search.press('Escape'); await expect(popup).toBeVisible()
    await search.dispatchEvent('compositionend', { data: '绿' }); await expect(search).toHaveCount(0)
    const red = popup.getByRole('checkbox', { name: 'Red', exact: true })
    await expect(red).toBeFocused(); await expect(popup.getByRole('checkbox')).toHaveCount(5)
    await page.evaluate(() => window.validationCellFixture.setDropdownUi({ showSearchOnDropdown: true, showEditOnDropdown: true }))
    await expect(search).toBeFocused(); await expect(search).toHaveValue('')
    await expect(popup.getByRole('button', { name: '编辑数据验证', exact: true })).toBeEnabled()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await closeList('B1'); await expect(grid()).toBeFocused()
  })

  test(`${variant}: a hidden-search single-select checkbox commits Enter once and closes through the shared history and focus path`, async () => {
    await appearance(variant)
    await page.evaluate(() => window.validationCellFixture.setDropdownUi({ showSearchOnDropdown: false, showEditOnDropdown: false }))
    await openList('B2'); const popup = listPopup('B2'), first = popup.getByRole('checkbox', { name: '1', exact: true }), last = popup.getByRole('checkbox', { name: 'Green, dark', exact: true })
    await expect(first).toBeFocused(); await first.press('End'); await expect(last).toBeFocused(); await last.press('Enter')
    await expect(popup).toBeHidden(); await expect(grid()).toBeFocused()
    assert.equal((await value(1, 1)).value, 'Green, dark'); assert.equal((await history()).undo, 1)
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await value(1, 1)).value, '1')
    await page.evaluate(() => window.validationCellFixture.redo()); await page.evaluate(() => window.validationCellFixture.reopen())
    await openList('B2'); await expect(last).toBeChecked(); assert.equal((await value(1, 1)).value, 'Green, dark')
  })

  test(`${variant}: clipping a composing popup retires its Canvas input lease and permits a fresh programmatic show without content changes`, async () => {
    await appearance(variant); await openList('B1')
    const popup = listPopup(), search = popup.getByRole('textbox', { name: '搜索列表项', exact: true }), oldInput = await search.elementHandle()
    await search.fill('Green'); await search.dispatchEvent('compositionstart', { data: '绿' })
    const before = await snapshot(), oldHistory = await history()
    await grid().evaluate(element => { element.scrollTop = 900; element.dispatchEvent(new Event('scroll')) })
    await expect(popup).toBeHidden(); assert.equal(await page.evaluate(() => window.validationCellFixture.activeDropdown()), null)
    await oldInput.evaluate(input => input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '绿' })))
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await grid().evaluate(element => { element.scrollTop = 0; element.dispatchEvent(new Event('scroll')) })
    assert.equal(await page.evaluate(() => window.validationCellFixture.showDropdown(0, 1)), true)
    await expect(popup).toBeVisible(); await expect(search).toHaveValue(''); await expect(search).toBeFocused()
    await expect(popup.getByRole('checkbox', { name: 'Red', exact: true })).toBeEnabled()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
  })

  test(`${variant}: calendar commits one serial plus format, preserves hidden time, and survives undo/redo and internal reopen`, async () => {
    await appearance(variant); await openDate('A1'); await expect(date()).toHaveValue('1900-02-28')
    await changePicker(date(), '1900-03-01'); await confirm('A1').click(); await expect(datePopup()).toBeHidden(); await expect(grid()).toBeFocused()
    const cell = await value(0); assert.equal(cell.value, 61.125); assert.equal(cell.style.numberFormat, 'dd/mm/yyyy'); assert.equal((await history()).undo, 1)
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await value(0)).value, 59.125)
    await page.evaluate(() => window.validationCellFixture.redo()); assert.deepEqual(await value(0), cell)
    await page.evaluate(() => window.validationCellFixture.reopen()); await expect(grid()).toBeVisible(); assert.deepEqual(await value(0), cell)
    await openDate('A1'); await expect(date()).toHaveValue('1900-03-01'); const before = await snapshot(), oldHistory = await history()
    await confirm('A1').click(); await expect(datePopup()).toBeHidden(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
  })

  test(`${variant}: pseudodate anchors stay visible and exact while a clock plus fractional-second edit changes only time`, async () => {
    await appearance(variant)
    for (const [key, row, day, label] of [['A2', 1, 60, '1900-02-29'], ['A3', 2, 0, '1900-01-00']]) {
      await openDate(key); await expect(datePopup(key).getByText(`Excel 兼容日期：${label}`, { exact: true })).toBeVisible()
      await expect(date(key)).toHaveValue(''); await expect(time(key)).toHaveValue('03:00:00')
      await changePicker(time(key), '12:34:56'); await datePopup(key).getByRole('spinbutton', { name: '小数秒', exact: true }).fill('0.25')
      await confirm(key).click(); await expect(datePopup(key)).toBeHidden()
      assert.equal((await value(row)).value, day + (12 * 3600 + 34 * 60 + 56.25) / 86400)
    }
    assert.equal((await history()).undo, 2)
  })

  test(`${variant}: numeric correction, a hidden-date time and elapsed duration use their real Ant inputs`, async () => {
    await appearance(variant); await openDate('A4'); await expect(datePopup('A4').getByText('当前数值无法在日历中表示，请修改日期数值。', { exact: true })).toBeVisible()
    await expect(confirm('A4')).toBeDisabled(); await datePopup('A4').getByRole('spinbutton', { name: '日期数值', exact: true }).fill('60.5')
    await confirm('A4').click(); await expect(datePopup('A4')).toBeHidden(); assert.equal((await value(3)).value, 60.5)
    await openDate('A5'); await expect(date('A5')).toHaveCount(0); await changePicker(time('A5'), '18:00:00'); await confirm('A5').click()
    await expect(datePopup('A5')).toBeHidden(); assert.equal((await value(4)).value, 45292.75)
    await openDate('A6'); await datePopup('A6').getByRole('textbox', { name: '持续时间', exact: true }).fill('72:30:15')
    await confirm('A6').click(); await expect(datePopup('A6')).toBeHidden(); assert.equal((await value(5)).value, (72 * 3600 + 30 * 60 + 15) / 86400)
    assert.equal((await history()).undo, 3)
  })

  test(`${variant}: a date-system conflict disables the old draft and reopening respects the 1904 epoch`, async () => {
    await appearance(variant); await openDate('A1'); await changePicker(date(), '1900-03-01')
    await page.evaluate(() => window.validationCellFixture.setDateSystem(true)); await expect(confirm('A1')).toBeDisabled()
    await datePopup().getByRole('button', { name: '取消', exact: true }).click(); await expect(datePopup()).toBeHidden()
    await page.evaluate(() => window.validationCellFixture.setCell('A1', 0)); await openDate('A1'); await expect(date()).toHaveValue('1904-01-01')
    await changePicker(date(), '1904-01-02'); await confirm('A1').click(); await expect(datePopup()).toBeHidden(); assert.equal((await value(0)).value, 1)
  })

  test(`${variant}: STOP and command cancellation keep the exact date draft, model and history until a valid retry or cancel`, async () => {
    await appearance(variant); await page.evaluate(() => window.validationCellFixture.stopDate()); const before = await snapshot(), oldHistory = await history()
    await openDate('A1'); await changePicker(date(), '1900-03-01'); await confirm('A1').click()
    await expect(datePopup()).toBeVisible(); await expect(date()).toHaveValue('1900-03-01'); await expect(datePopup().locator('.ant-alert-error')).toBeVisible()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await changePicker(date(), '1900-02-28'); await page.evaluate(id => window.validationCellFixture.cancelCommand(id, true), COMMIT_SHEET_DATE_INPUT_COMMAND)
    await confirm('A1').click(); await expect(datePopup()).toBeVisible(); await expect(datePopup().getByText('选择操作已取消，草稿已保留。', { exact: true })).toBeVisible()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await page.evaluate(id => window.validationCellFixture.cancelCommand(id, false), COMMIT_SHEET_DATE_INPUT_COMMAND)
    await confirm('A1').click(); await expect(datePopup()).toBeHidden(); await expect(grid()).toBeFocused(); assert.deepEqual(await history(), oldHistory)
  })

  test(`${variant}: multiple candidate toggles, CSV labels, search, keyboard selection and reopen preserve order and one transaction per change`, async () => {
    await appearance(variant); await openList('B1'); const popup = listPopup()
    await expect(popup.getByRole('checkbox', { name: 'Red', exact: true })).toBeChecked()
    await popup.getByRole('checkbox', { name: 'Green, dark', exact: true }).check(); await expect(popup).toBeVisible()
    assert.equal((await value(0, 1)).value, 'Red,"Green, dark"'); assert.equal((await history()).undo, 1)
    await popup.getByRole('checkbox', { name: 'It"s blue', exact: true }).check(); assert.equal((await value(0, 1)).value, 'Red,"Green, dark",It""s blue')
    await popup.getByRole('textbox', { name: '搜索列表项', exact: true }).fill('green'); await expect(popup.getByRole('checkbox')).toHaveCount(1)
    await popup.getByRole('checkbox').uncheck(); assert.equal((await value(0, 1)).value, 'Red,It""s blue')
    await popup.getByRole('textbox', { name: '搜索列表项', exact: true }).fill(''); const first = popup.getByRole('checkbox', { name: 'Red', exact: true })
    await first.focus(); await first.press('End'); await expect(popup.getByRole('checkbox', { name: '=1+1', exact: true })).toBeFocused()
    await page.keyboard.press('Space'); assert.equal((await value(0, 1)).value, 'Red,It""s blue,=1+1'); assert.equal((await value(0, 1)).formula, null)
    await closeList('B1'); await expect(grid()).toBeFocused(); await openList('B1')
    await expect(listPopup().getByRole('checkbox', { name: '=1+1', exact: true })).toBeChecked()
    await listPopup().getByRole('button', { name: '清除选择', exact: true }).click(); assert.equal((await value(0, 1)).value, ''); assert.equal((await history()).undo, 5)
    await closeList('B1'); await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await value(0, 1)).value, 'Red,It""s blue,=1+1')
    await page.evaluate(() => window.validationCellFixture.redo()); await page.evaluate(() => window.validationCellFixture.reopen())
    await openList('B1'); await expect(listPopup().getByRole('checkbox', { checked: true })).toHaveCount(0)
  })

  test(`${variant}: single values remain complete text while required clearing and canceled list choices do not write`, async () => {
    await appearance(variant); await openList('B2'); await listPopup('B2').getByRole('checkbox', { name: '=1+1', exact: true }).check()
    await expect(listPopup('B2')).toBeHidden(); assert.equal((await value(1, 1)).value, '=1+1'); assert.equal((await value(1, 1)).dataType, 'text'); assert.equal((await value(1, 1)).formula, null)
    await openList('B3'); const before = await snapshot(), oldHistory = await history()
    await listPopup('B3').getByRole('button', { name: '清除选择', exact: true }).click(); await expect(listPopup('B3').locator('.ant-alert-error')).toBeVisible()
    await expect(listPopup('B3').getByRole('checkbox', { name: 'Red', exact: true })).toBeChecked(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await page.evaluate(id => window.validationCellFixture.cancelCommand(id, true), COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND)
    await listPopup('B3').getByRole('checkbox', { name: 'Blue', exact: true }).check(); await expect(listPopup('B3').getByRole('checkbox', { name: 'Blue', exact: true })).not.toBeChecked()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory); await closeList('B3')
  })

  test(`${variant}: live reference candidates reflect source edits and View denial without displaying protected labels`, async () => {
    await appearance(variant); await openList('B4'); await expect(listPopup('B4').getByRole('checkbox', { name: 'Green, dark', exact: true })).toBeVisible()
    await page.evaluate(() => window.validationCellFixture.setCell('D2', 'Approved'))
    await expect(listPopup('B4').getByRole('checkbox', { name: 'Green, dark', exact: true })).toHaveCount(0)
    await listPopup('B4').getByRole('checkbox', { name: 'Approved', exact: true }).check(); assert.equal((await value(3, 1)).value, 'Red,Approved')
    await page.evaluate(() => window.validationCellFixture.protectSource()); await expect(listPopup('B4').getByRole('checkbox')).toHaveCount(0)
    await expect(listPopup('B4').getByText('没有匹配的列表项。', { exact: true })).toBeVisible(); await closeList('B4')
  })

  test(`${variant}: Alt+Down works from Canvas and an existing draft; cancellation restores the owner input and confirmation ends it`, async () => {
    await appearance(variant); await select('A1'); await grid().focus(); await grid().press('Alt+ArrowDown'); await expect(datePopup()).toBeVisible()
    await datePopup().getByRole('button', { name: '取消', exact: true }).click(); await expect(datePopup()).toBeHidden(); await expect(grid()).toBeFocused()
    await grid().press('F2'); const editor = page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
    await editor.fill('uncommitted'); await editor.press('Alt+ArrowDown'); await expect(datePopup()).toBeVisible(); await expect(date()).toHaveValue('1900-02-28')
    await datePopup().getByRole('button', { name: '取消', exact: true }).click(); await expect(datePopup()).toBeHidden(); await expect(editor).toHaveValue('uncommitted'); await expect(editor).toBeFocused()
    await editor.press('Alt+ArrowDown'); await expect(datePopup()).toBeVisible(); await changePicker(date(), '1900-03-01'); await confirm('A1').click()
    await expect(datePopup()).toBeHidden(); await expect(editor).toHaveCount(0); await expect(grid()).toBeFocused(); assert.equal((await value(0)).value, 61.125)
  })

  test(`${variant}: synthetic IME guards picker commit/cancel and list actions until composition ends`, async () => {
    await appearance(variant); await openDate('A1'); await date().dispatchEvent('compositionstart', { data: '2' })
    await expect(confirm('A1')).toBeDisabled(); await expect(datePopup().getByRole('button', { name: '取消', exact: true })).toBeDisabled()
    await date().press('Escape'); await expect(datePopup()).toBeVisible(); assert.equal((await history()).undo, 0)
    await date().dispatchEvent('compositionend', { data: '1900-02-28' }); await datePopup().getByRole('button', { name: '取消', exact: true }).click(); await expect(datePopup()).toBeHidden()
    await openList('B1'); const search = listPopup().getByRole('textbox', { name: '搜索列表项', exact: true })
    await search.dispatchEvent('compositionstart', { data: '绿' }); await expect(listPopup().getByRole('button', { name: '关闭', exact: true })).toBeDisabled()
    await search.press('Escape'); await expect(listPopup()).toBeVisible(); await listPopup().getByRole('checkbox', { name: 'Green, dark', exact: true }).click()
    assert.equal((await history()).undo, 0); await search.dispatchEvent('compositionend', { data: '绿' }); await closeList('B1')
  })

  test(`${variant}: selection/sheet switches and replacement history prevent stale popups from saving or stealing focus`, async () => {
    await appearance(variant); await openDate('A1'); await changePicker(date(), '1900-03-01'); await select('B1'); await expect(datePopup()).toBeHidden()
    await page.getByRole('button', { name: '选择列表项 B1', exact: true }).click(); await expect(listPopup()).toBeVisible(); assert.equal((await history()).undo, 0)
    await page.evaluate(() => window.validationCellFixture.switchSheet()); await expect(listPopup()).toBeHidden(); await expect(nameBox()).toBeVisible()
    await page.reload(); await expect(grid()).toBeVisible(); await appearance(variant); await openDate('A1'); await changePicker(date(), '1900-03-01')
    await page.evaluate(() => window.validationCellFixture.clearHistory()); await expect(datePopup()).toBeHidden(); await openDate('A1')
    await page.evaluate(() => window.validationCellFixture.replaceSameId()); await expect(datePopup()).toBeHidden(); assert.equal((await value(0)).value, 59.125)
  })

  test(`${variant}: unrelated model edits keep a valid owner while target conflicts retain a disabled draft`, async () => {
    await appearance(variant); await openDate('A1'); await changePicker(date(), '1900-03-01'); await page.evaluate(() => window.validationCellFixture.setCell('G10', 'Other'))
    await expect(confirm('A1')).toBeEnabled(); await expect(date()).toHaveValue('1900-03-01')
    await page.evaluate(() => window.validationCellFixture.setCell('A1', 62.125)); await expect(confirm('A1')).toBeDisabled(); await expect(date()).toHaveValue('1900-03-01')
    await datePopup().getByRole('button', { name: '取消', exact: true }).click(); await expect(datePopup()).toBeHidden(); assert.equal((await value(0)).value, 62.125)
  })

  test(`${variant}: switching in a change observer cannot return old popup focus to the former sheet`, async () => {
    await appearance(variant); await openDate('A1'); await changePicker(date(), '1900-03-01'); await page.evaluate(() => window.validationCellFixture.switchAfterChange())
    await confirm('A1').click(); await expect(datePopup()).toBeHidden(); assert.equal((await snapshot()).activeSheetId, 'other'); assert.equal((await value(0)).value, 61.125)
    await expect(page.getByRole('tab', { name: 'Other', exact: true })).toHaveAttribute('aria-selected', 'true')
  })

  test(`${variant}: the date rule checkbox changes showTime in one transaction, preserves metadata, and controls the reopened picker`, async () => {
    await appearance(variant); await editDateRule(); await expect(panel().getByRole('checkbox', { name: '显示时间', exact: true })).not.toBeChecked()
    await panel().getByRole('checkbox', { name: '显示时间', exact: true }).check(); await panel().getByRole('button', { name: '应用', exact: true }).click()
    const rule = (await snapshot()).sheets[0].dataValidations.find(rule => rule.id === 'date-0')
    assert.deepEqual(rule.metadata, { showTime: true, retained: ['fixture', 0] }); assert.equal((await history()).undo, 1)
    await panel().getByRole('button', { name: '关闭', exact: true }).click(); await expect(panel()).toBeHidden(); await expect(grid()).toBeFocused()
    await openDate('A1'); await expect(time('A1')).toHaveValue('03:00:00'); await datePopup().getByRole('button', { name: '取消', exact: true }).click(); await expect(datePopup()).toBeHidden()
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await snapshot()).sheets[0].dataValidations.find(rule => rule.id === 'date-0').metadata.showTime, false)
    await page.evaluate(() => window.validationCellFixture.redo()); await page.evaluate(() => window.validationCellFixture.reopen()); await openDate('A1'); await expect(time('A1')).toHaveValue('03:00:00')
  })

  test(`${variant}: rule cancellation keeps the checkbox and original history, then a valid retry and sidebar add use the same owner`, async () => {
    await appearance(variant); await editDateRule(); await panel().getByRole('checkbox', { name: '显示时间', exact: true }).check()
    const before = await snapshot(), oldHistory = await history()
    await page.evaluate(id => window.validationCellFixture.cancelCommand(id, true), SHEET_DATA_VALIDATION_COMMANDS.replace)
    await panel().getByRole('button', { name: '应用', exact: true }).click(); await expect(panel().locator('.ant-alert-error')).toBeVisible()
    await expect(panel().getByRole('checkbox', { name: '显示时间', exact: true })).toBeChecked(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await page.evaluate(id => window.validationCellFixture.cancelCommand(id, false), SHEET_DATA_VALIDATION_COMMANDS.replace)
    await panel().getByRole('button', { name: '应用', exact: true }).click(); await expect(panel().getByRole('button', { name: '添加规则', exact: true })).toBeVisible()
    await panel().getByRole('button', { name: '添加规则', exact: true }).click(); await expect(panel().getByRole('textbox', { name: '数据验证规则值', exact: true })).toHaveValue('100')
    assert.equal((await history()).undo, 2); await expect(panel()).toBeVisible()
  })

  test(`${variant}: list row edits, color input, ordering and multiline labels save an aligned CSV/color definition and feed the Canvas popup`, async () => {
    await appearance(variant); await openList('B1'); await listPopup().getByRole('button', { name: '编辑数据验证', exact: true }).click(); await expect(panel()).toBeVisible()
    await panel().getByRole('textbox', { name: '列表项 2', exact: true }).fill('Green, deep'); await panel().getByRole('button', { name: '上移列表项 2', exact: true }).click()
    await expect(panel().getByRole('textbox', { name: '列表项 1', exact: true })).toHaveValue('Green, deep')
    await panel().getByRole('button', { name: '列表项颜色 1', exact: true }).click()
    const hex = page.locator('.ant-color-picker-hex-input:visible input, input.ant-color-picker-hex-input:visible'); await expect(hex).toBeVisible(); await hex.fill('123456')
    await panel().getByRole('button', { name: '列表项颜色 1', exact: true }).click(); await expect(hex).toBeHidden()
    await panel().getByRole('button', { name: '删除列表项 3', exact: true }).click(); await panel().getByRole('button', { name: '添加列表项', exact: true }).click()
    await panel().getByRole('textbox', { name: '列表项 5', exact: true }).fill('Line\nbreak'); await panel().getByRole('button', { name: '应用', exact: true }).click()
    const rule = (await snapshot()).sheets[0].dataValidations.find(rule => rule.id === 'multiple')
    assert.equal(rule.formula1, '""Green, deep",Red,1,=1+1,"Line\nbreak""'); assert.equal(rule.formula2.toLowerCase(), '#123456,#ff0000,,,'); assert.equal((await history()).undo, 1)
    await panel().getByRole('button', { name: '关闭', exact: true }).click(); await expect(panel()).toBeHidden(); await openList('B1')
    await listPopup().getByRole('checkbox', { name: 'Green, deep', exact: true }).check(); assert.equal((await value(0, 1)).value, '"Green, deep",Red'); assert.equal((await history()).undo, 2)
  })

  test(`${variant}: hidden rule ranges have no list entry and visible protected definitions cannot be edited, removed or included in clear-all`, async () => {
    await appearance(variant); await page.evaluate(() => window.validationCellFixture.protectRule(true)); await manageRules()
    await expect(panel().getByText('B1, N15', { exact: true })).toHaveCount(0); await expect(panel().getByRole('button', { name: '清除全部规则', exact: true })).toBeDisabled()
    const before = await snapshot(), oldHistory = await history(); await panel().getByRole('button', { name: '关闭', exact: true }).click(); await expect(panel()).toBeHidden()
    await openList('B1'); await expect(listPopup().getByRole('button', { name: '编辑数据验证', exact: true })).toHaveCount(0); await closeList('B1')
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await page.reload(); await expect(grid()).toBeVisible(); await appearance(variant); await page.evaluate(() => window.validationCellFixture.protectRule(false)); await manageRules()
    await expect(panel().getByRole('button', { name: '删除规则 B1, N15', exact: true })).toBeDisabled()
    await expect(panel().locator('.validation-rule-main').filter({ hasText: 'B1, N15' })).toBeDisabled(); await expect(panel().getByRole('button', { name: '清除全部规则', exact: true })).toBeDisabled()
  })

  test(`${variant}: synthetic composition inside a rule field or owned color popup blocks save/close, and external history/content closes the draft`, async () => {
    await appearance(variant); await openList('B1'); await listPopup().getByRole('button', { name: '编辑数据验证', exact: true }).click(); await expect(panel()).toBeVisible()
    const item = panel().getByRole('textbox', { name: '列表项 1', exact: true }); await item.dispatchEvent('compositionstart', { data: '红' })
    await expect(panel().getByRole('button', { name: '应用', exact: true })).toBeDisabled(); await expect(panel().getByRole('button', { name: '关闭', exact: true })).toBeDisabled()
    await item.press('Escape'); await expect(panel()).toBeVisible(); await item.dispatchEvent('compositionend', { data: 'Red' })
    await panel().getByRole('button', { name: '列表项颜色 1', exact: true }).click(); const hex = page.locator('.ant-color-picker-hex-input:visible input, input.ant-color-picker-hex-input:visible')
    await hex.dispatchEvent('compositionstart', { data: 'f' }); await expect(panel().getByRole('button', { name: '应用', exact: true })).toBeDisabled()
    await hex.dispatchEvent('compositionend', { data: 'ff0000' }); await panel().getByRole('button', { name: '列表项颜色 1', exact: true }).click()
    await page.evaluate(() => window.validationCellFixture.clearHistory()); await expect(panel()).toBeHidden(); assert.equal((await history()).undo, 0)
    await editDateRule(); await panel().getByRole('checkbox', { name: '显示时间', exact: true }).check(); await page.evaluate(() => window.validationCellFixture.setCell('G10', 'External'))
    await expect(panel()).toBeHidden(); assert.equal((await snapshot()).sheets[0].dataValidations.find(rule => rule.id === 'date-0').metadata.showTime, false)
  })

  test(`${variant}: a sheet switch in the rule change observer keeps the committed definition and closes the former owner without stale focus`, async () => {
    await appearance(variant); await editDateRule(); await panel().getByRole('checkbox', { name: '显示时间', exact: true }).check()
    await page.evaluate(() => window.validationCellFixture.switchAfterRule()); await panel().getByRole('button', { name: '应用', exact: true }).click(); await expect(panel()).toBeHidden()
    const model = await snapshot(); assert.equal(model.activeSheetId, 'other'); assert.equal(model.sheets[0].dataValidations.find(rule => rule.id === 'date-0').metadata.showTime, true)
    await expect(page.getByRole('tab', { name: 'Other', exact: true })).toHaveAttribute('aria-selected', 'true')
  })

  test(`${variant}: a frozen merged target owns one bounded trigger after zoom/scroll and Ant themes change the popup surface`, async () => {
    await appearance(variant); await page.evaluate(() => window.validationCellFixture.freezeMerge()); await select('B2')
    const trigger = page.getByRole('button', { name: '选择日期或时间 A1', exact: true }); await expect(trigger).toHaveCount(1)
    for (let count = 0; count < 5; count++) await page.getByRole('button', { name: '放大工作表', exact: true }).click()
    await page.locator('.canvas-scroll').evaluate(element => { element.scrollLeft = 350; element.scrollTop = 150; element.dispatchEvent(new Event('scroll')) })
    const host = await grid().boundingBox(), box = await trigger.boundingBox(); assert.ok(box.x >= host.x && box.x + box.width <= host.x + host.width); assert.ok(box.y >= host.y && box.y + box.height <= host.y + host.height)
    await trigger.click(); await expect(datePopup()).toBeVisible(); const surface = page.locator('.sheet-validation-cell-popover .ant-popover-inner')
    const light = await surface.evaluate(element => getComputedStyle(element).backgroundColor)
    await appearance(variant, true); await expect(surface).not.toHaveCSS('background-color', light); await expect(datePopup().locator('.ant-picker')).toBeVisible()
    await datePopup().getByRole('button', { name: '取消', exact: true }).click(); await expect(datePopup()).toBeHidden()
  })
}
