import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'
import { pivotChartFieldButtonsIds as ids } from './fixtures/pivot-chart-field-buttons-fixture.mjs'

// Deferred real Ribbon/table/Ant/Canvas contracts; no tests, dev server or browser have run.
// The planned renderer is Canvas2D. Composition events do not establish native Windows IME behavior.
let server, browser, context, page, url
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No ordinary-field fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/pivot-table-field-filter-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1200 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.pivotTableFieldFilterFixture?.state().model.pivotTables.length === 2)
}, 60_000)
afterEach(async () => {
  try { await page?.evaluate(() => window.pivotTableFieldFilterFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.pivotTableFieldFilterFixture.state())
const main = () => page.getByTestId('pivot-table-panel')
const popup = () => page.getByTestId('pivot-table-field-filter-popup')
const tools = () => page.getByTestId('pivot-field-tools')
const dropdown = () => page.locator('.ant-select-dropdown:visible')
const apply = (kind = 'members') => popup().getByRole('button', { name: kind === 'date' ? '应用透视表日期筛选' : kind === 'label' ? '应用透视表标签条件' : '应用透视表字段筛选', exact: true })
const clear = (kind = 'members') => popup().getByRole('button', { name: kind === 'date' ? '清除透视表日期条件' : kind === 'label' ? '清除透视表标签条件' : '清除透视表字段筛选', exact: true })
const member = key => popup().getByRole('checkbox', { name: `选择透视表标签${key}`, exact: true })
const all = () => popup().getByRole('checkbox', { name: '全选所有项目', exact: true })
const report = value => value.model.pivotTables.find(pivot => pivot.id === value.pivotId)
const field = (value, name) => report(value).fields.find(field => field.name === name)
const total = value => value.data.slice(1).flatMap(row => row.slice(1)).reduce((sum, value) => sum + (typeof value === 'number' ? value : 0), 0)
const sourceCells = value => value.model.sheets.find(sheet => sheet.id === ids.source).cells
async function mount(options = {}) { await page.evaluate(options => window.pivotTableFieldFilterFixture.mount(options), options) }
async function choose(name, label, target = main()) {
  await target.getByRole('combobox', { name, exact: true }).click(); await dropdown().getByText(label, { exact: typeof label === 'string' }).click()
}
async function openMain() {
  if (await main().count() && await main().isVisible()) return
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  const trigger = page.getByRole('toolbar', { name: '数据', exact: true }).getByRole('button', { name: '数据透视表', exact: true })
  await expect(trigger).toBeEnabled(); await trigger.click(); await expect(main()).toBeVisible()
  await choose('选择数据透视表', /^数据透视表 1 ·/)
}
async function fieldMenu(name = 'Region') {
  await openMain(); const item = await page.evaluate(name => window.pivotTableFieldFilterFixture.field(name), name)
  await main().locator(`.pivot-area-grid [data-field-id="${item.id}"]`).getByRole('button', { name: `字段 ${item.name} 操作`, exact: true }).click()
  return item
}
async function open(name = 'Region') {
  const item = await fieldMenu(name); await page.getByRole('menuitem', { name: '筛选字段', exact: true }).click()
  await expect(popup()).toBeVisible(); await expect(popup().getByRole('textbox', { name: '搜索透视表筛选标签', exact: true })).toBeEnabled(); return item
}
async function openTools(name, kind = 'group') {
  await fieldMenu(name); await page.getByRole('menuitem', { name: kind === 'group' ? '分组' : '值筛选', exact: true }).click(); await expect(tools()).toBeVisible()
}
async function waitHistory(count) { await page.waitForFunction(count => window.pivotTableFieldFilterFixture.state().history.undo === count, count) }
async function dateDraft(label = '等于', first = '2026-02-02 00:00:00.000', second) {
  await popup().getByRole('button', { name: '打开透视表日期筛选', exact: true }).click(); await choose('透视表日期条件', label, popup())
  const inputs = popup().locator('.ant-picker input'); await inputs.nth(0).fill(first); await inputs.nth(0).press('Tab')
  if (second !== undefined) { await inputs.nth(1).fill(second); await inputs.nth(1).press('Tab') }
}
async function labelDraft(label = '等于', first = 'North', second) {
  await popup().getByRole('button', { name: '打开透视表标签条件', exact: true }).click(); await choose('透视表标签条件', label, popup())
  await popup().getByRole('textbox', { name: '透视表条件标签', exact: true }).fill(first)
  if (second !== undefined) await popup().getByRole('textbox', { name: '透视表条件结束标签', exact: true }).fill(second)
}
async function changeOwner(kind) {
  await page.evaluate(async kind => {
    const fixture = window.pivotTableFieldFilterFixture
    if (kind === 'source') await fixture.changeSource()
    if (kind === 'history') await fixture.clearHistory()
    if (kind === 'user') await fixture.changeUser()
    if (kind === 'host') await fixture.switchSheet()
  }, kind)
}
async function reread() { await main().getByRole('button', { name: '重新读取', exact: true }).click() }

test.each(['classic', 'grid'])('%s real Ribbon opens ordinary report with the full shared member/condition Ant surface', async variant => {
  await mount({ variant }); await openMain(); const before = await state(); await open('Date')
  await expect(popup().getByRole('button', { name: '打开透视表日期筛选', exact: true })).toBeVisible()
  await expect(popup().getByRole('button', { name: '打开透视表标签条件', exact: true })).toBeVisible()
  await expect(popup().locator('form')).toHaveCount(1); await expect(main().getByRole('switch', { name: '允许字段使用多个筛选', exact: true })).toBeDisabled()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, 0)
})

test('manual confirmation retains the pending destination and options while output and linked chart change once', async () => {
  await openMain(); await main().getByRole('textbox', { name: '透视表起始单元格', exact: true }).fill('K30')
  await main().getByRole('switch', { name: '重复行标签', exact: true }).click(); await open(); const before = await state()
  await member('South').uncheck(); await apply().click(); await waitHistory(1)
  const current = await state(); assert.equal(total(current), 40); assert.deepEqual(sourceCells(current), sourceCells(before))
  await expect(main().getByRole('textbox', { name: '透视表起始单元格', exact: true })).toHaveValue('K30')
  await expect(main().getByRole('switch', { name: '重复行标签', exact: true })).toBeChecked()
  assert.equal(report(current).anchor.row, 24); assert.equal(report(current).anchor.col, 0)
  const commands = current.commands.filter(command => command.id === 'sheet.command.set-pivot-filter'); assert.equal(commands.length, 1); assert.equal(commands[0].params.hostSheetId, ids.host)
  assert.equal(current.linkedData.series.flatMap(series => series.values).reduce((sum, value) => sum + (typeof value === 'number' ? value : 0), 0), 40)
  await page.evaluate(() => window.pivotTableFieldFilterFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

test('member cancel leaves content/history intact and returns focus to the actual field menu trigger', async () => {
  const item = await open(), before = await state(); await member('South').uncheck(); await popup().getByRole('button', { name: '取消', exact: true }).click()
  await expect(popup()).toHaveCount(0); assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  await expect(main().locator(`.pivot-area-grid [data-field-id="${item.id}"]`).getByRole('button', { name: '字段 Region 操作', exact: true })).toBeFocused()
})

test('all applies beyond search visibility and an empty draft cannot publish', async () => {
  await open(); await all().uncheck(); await expect(apply()).toBeDisabled(); const before = await state()
  await popup().getByRole('textbox', { name: '搜索透视表筛选标签', exact: true }).fill('North'); await all().check(); await apply().click()
  assert.equal(total(await state()), 100); assert.deepEqual(sourceCells(await state()), sourceCells(before))
})

const partialCases = [[{ partialDate: true }, 'Date', '2026', 100], [{ typedMembers: true }, 'Region', '1', 30]]
test.each(partialCases)('ordinary partial %j remains exact after confirmation/history/snapshot reopen', async (options, name, key) => {
  await mount(options); await open(name); const before = await state(); await expect(member(key)).not.toBeChecked()
  await expect(popup().getByText(/部分选中 1 项/)).toBeVisible(); await expect(apply()).toBeEnabled(); await apply().click(); await waitHistory(1)
  assert.equal(total(await state()), 10); const saved = await state(); await page.evaluate(() => window.pivotTableFieldFilterFixture.undo()); assert.deepEqual((await state()).model, before.model)
  await page.evaluate(() => window.pivotTableFieldFilterFixture.redo()); assert.deepEqual((await state()).model, saved.model)
  await main().getByRole('button', { name: '关闭', exact: true }).click(); await page.evaluate(() => window.pivotTableFieldFilterFixture.reopen()); await open(name)
  await expect(popup().getByText(/部分选中 1 项/)).toBeVisible(); assert.equal(total(await state()), 10)
})
test.each(partialCases)('checking ordinary partial %j widens the group only by the explicit user action', async (options, name, key, expected) => {
  await mount(options); await open(name); await member(key).check(); await apply().click(); await waitHistory(1); assert.equal(total(await state()), expected)
})

test('date grouping via the real field-tools panel produces year members and refreshes the owner without losing the main draft', async () => {
  await openMain(); await main().getByRole('textbox', { name: '透视表起始单元格', exact: true }).fill('K30'); await openTools('Date')
  await choose('分组方式', '日期分组', tools()); await choose('日期分组依据', '年', tools()); const before = await state()
  await tools().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(before.history.undo + 1); await expect(tools()).toHaveCount(0)
  await expect(main().getByRole('textbox', { name: '透视表起始单元格', exact: true })).toHaveValue('K30'); await open('Date')
  await expect(member('2026')).toBeVisible(); await expect(member('2026-01-15')).toHaveCount(0); assert.equal(total(await state()), 100)
})

test('manual group source items retain numeric/text identity and feed the ordinary partial checklist', async () => {
  await mount({ typedMembers: true }); await openTools('Region'); await choose('分组方式', '组合项目', tools())
  await tools().getByRole('textbox', { name: '新分组名称', exact: true }).fill('Mix')
  await tools().getByRole('combobox', { name: '分组源项目', exact: true }).click()
  await dropdown().getByText('1（数值）', { exact: true }).click(); await dropdown().getByText('1（文本）', { exact: true }).click()
  await tools().getByRole('textbox', { name: '新分组名称', exact: true }).click(); await tools().getByRole('button', { name: '添加组', exact: true }).click()
  await tools().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1); await open('Region')
  await expect(member('Mix')).not.toBeChecked(); await expect(popup().getByText(/部分选中 1 项/)).toBeVisible(); assert.equal(total(await state()), 10)
  await member('Mix').check(); await apply().click(); await waitHistory(2); assert.equal(total(await state()), 30)
})

test('the existing value-filter panel intersects members only when the committed multiple-filter option allows it', async () => {
  await page.evaluate(() => window.pivotTableFieldFilterFixture.setOptions({ allowMultipleFilters: true })); await open(); await member('North').uncheck(); await apply().click()
  await openTools('Region', 'filter'); await choose('值筛选条件', '大于或等于', tools())
  await tools().getByRole('spinbutton', { name: '筛选数值', exact: true }).fill('50'); const before = await state()
  await tools().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(before.history.undo + 1)
  assert.equal(total(await state()), 60); assert.ok(field(await state(), 'Region').groupedItems); assert.ok(field(await state(), 'Region').valueFilter)
})

test.each([['透视表字段升序', 'ascending'], ['透视表字段降序', 'descending'], ['清除透视表字段排序', undefined]])('%s uses one guarded sort and remains undoable', async (label, direction) => {
  if (direction === undefined) { await open(); await popup().getByRole('button', { name: '透视表字段降序', exact: true }).click(); await waitHistory(1) }
  await open(); const before = await state(); await popup().getByRole('button', { name: label, exact: true }).click(); await waitHistory(before.history.undo + 1)
  assert.equal(field(await state(), 'Region').sort, direction); assert.equal((await state()).commands.at(-1).params.expectedUserId, 'local-user')
  await page.evaluate(() => window.pivotTableFieldFilterFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

test.each([true, false])('ordinary multiple=%s is a draft until one configure, with shared cleanup and repeated no-op', async enabled => {
  if (!enabled) await page.evaluate(() => window.pivotTableFieldFilterFixture.multipleFilters())
  await openMain(); const before = await state(), control = main().getByRole('switch', { name: '允许字段使用多个筛选', exact: true })
  await control.click(); assert.deepEqual((await state()).model, before.model)
  await main().getByRole('button', { name: '应用字段配置', exact: true }).click(); await waitHistory(before.history.undo + 1)
  const saved = await state(); assert.equal(report(saved).options.allowMultipleFilters, enabled)
  if (!enabled) { assert.equal(field(saved, 'Region').groupedItems, undefined); assert.equal(field(saved, 'Region').labelFilter, undefined); assert.equal(field(saved, 'Region').valueFilter, undefined); assert.equal(total(saved), 100) }
  await main().getByRole('button', { name: '应用字段配置', exact: true }).click(); assert.equal((await state()).history.undo, saved.history.undo)
  assert.deepEqual((await state()).model, saved.model); await page.evaluate(() => window.pivotTableFieldFilterFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

test('ordinary member all preserves caption/value categories, while clear removes all categories once', async () => {
  await page.evaluate(() => window.pivotTableFieldFilterFixture.multipleFilters()); await open(); const initial = await state()
  await all().check(); await apply().click(); await waitHistory(initial.history.undo + 1)
  assert.ok(field(await state(), 'Region').labelFilter); assert.ok(field(await state(), 'Region').valueFilter); assert.equal(field(await state(), 'Region').groupedItems, undefined)
  await open(); const before = await state(); await clear().click(); await waitHistory(before.history.undo + 1); assert.equal(total(await state()), 100)
  for (const key of ['selectedItems', 'groupedItems', 'partialItems', 'recordSelection', 'dateFilter', 'labelFilter', 'valueFilter']) assert.equal(field(await state(), 'Region')[key], undefined)
})

test.each([
  ['dateBetween', '介于', '2026-02-02', '2026-02-20', 50], ['dateEqual', '等于', '2026-02-02', undefined, 20],
  ['dateNewerThan', '晚于', '2026-02-02', undefined, 70], ['dateNewerThanOrEqual', '晚于或等于', '2026-02-02', undefined, 90],
  ['dateNotBetween', '不介于', '2026-02-02', '2026-02-20', 50], ['dateNotEqual', '不等于', '2026-02-02', undefined, 80],
  ['dateOlderThan', '早于', '2026-02-02', undefined, 10], ['dateOlderThanOrEqual', '早于或等于', '2026-02-02', undefined, 30],
])('ordinary Ant date %s selects actual source records and writes a typed command', async (operator, label, first, second, expected) => {
  await open('Date'); const before = await state(); await dateDraft(label, `${first} 00:00:00.000`, second && `${second} 00:00:00.000`)
  await expect(apply('date')).toBeEnabled(); await apply('date').click(); await waitHistory(1)
  assert.equal(total(await state()), expected); assert.equal(field(await state(), 'Date').dateFilter.operator, operator)
  assert.equal((await state()).commands.filter(command => command.id === 'sheet.command.set-pivot-date-filter').length, 1)
  await page.evaluate(() => window.pivotTableFieldFilterFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

test.each([
  ['captionBeginsWith', '开头是', 'N', undefined, 40], ['captionBetween', '介于', 'North', 'South', 100], ['captionContains', '包含', 'or', undefined, 40],
  ['captionEndsWith', '结尾是', 'orth', undefined, 40], ['captionEqual', '等于', 'North', undefined, 40], ['captionGreaterThan', '大于', 'North', undefined, 60],
  ['captionGreaterThanOrEqual', '大于或等于', 'North', undefined, 100], ['captionLessThan', '小于', 'South', undefined, 40], ['captionLessThanOrEqual', '小于或等于', 'South', undefined, 100],
  ['captionNotBeginsWith', '开头不是', 'N', undefined, 60], ['captionNotBetween', '不介于', 'North', 'South', 0], ['captionNotContains', '不包含', 'or', undefined, 60],
  ['captionNotEndsWith', '结尾不是', 'orth', undefined, 60], ['captionNotEqual', '不等于', 'North', undefined, 60],
])('ordinary caption %s persists real output rather than only showing a condition button', async (operator, label, first, second, expected) => {
  await open(); await labelDraft(label, first, second); await apply('label').click(); await waitHistory(1)
  assert.equal(total(await state()), expected); assert.equal(field(await state(), 'Region').labelFilter.operator, operator)
})

test.each(['missing', 'reversed'])('ordinary %s date operands keep confirmation disabled', async mode => {
  await open('Date'); await popup().getByRole('button', { name: '打开透视表日期筛选', exact: true }).click()
  if (mode === 'reversed') await dateDraft('介于', '2026-03-01 00:00:00.000', '2026-01-15 00:00:00.000')
  await expect(apply('date')).toBeDisabled(); assert.equal((await state()).history.undo, 0); assert.equal(total(await state()), 100)
})

test('renaming a mixed numeric/text label to Date does not invent a date capability', async () => {
  await mount({ typedMembers: true }); await page.evaluate(() => window.pivotTableFieldFilterFixture.fieldSetting('Region', { displayName: 'Date' }))
  await open('Region'); await expect(popup().getByRole('button', { name: '打开透视表日期筛选', exact: true })).toHaveCount(0)
  await expect(popup().getByRole('button', { name: '打开透视表标签条件', exact: true })).toBeVisible()
})

test.each([false, true])('ordinary calendar date1904=%s preserves exact times and whole-day comparison', async date1904 => {
  await mount({ date1904, dateSerials: true }); const serial = sourceCells(await state())['5:5'].value
  await page.evaluate(async serial => { await window.pivotTableFieldFilterFixture.sourceCell('F4', serial); await window.pivotTableFieldFilterFixture.sourceCell('F5', serial + .5) }, serial)
  await open('Date'); await dateDraft('等于', '2026-02-20 00:00:00.000'); await apply('date').click(); assert.equal(total(await state()), 40)
  await open('Date'); await popup().getByRole('button', { name: '打开透视表日期筛选', exact: true }).click()
  await popup().getByRole('checkbox', { name: '按整天比较', exact: true }).check(); await apply('date').click()
  assert.equal(total(await state()), 60); assert.equal(field(await state(), 'Date').dateFilter.expected, serial)
})

test('ordinary saved Excel day 60 uses the shared Ant compatibility input and survives reopen unchanged', async () => {
  await mount({ compatibilityDate: true }); await page.evaluate(() => window.pivotTableFieldFilterFixture.dateFilter({ operator: 'dateEqual', expected: 60.25 }))
  await open('Date'); await popup().getByRole('button', { name: '打开透视表日期筛选', exact: true }).click()
  await expect(popup().getByRole('textbox', { name: '透视表筛选日期', exact: true })).toHaveValue('1900-02-29 06:00:00.000'); await apply('date').click()
  assert.equal(total(await state()), 40); assert.equal(field(await state(), 'Date').dateFilter.expected, 60.25)
  await main().getByRole('button', { name: '关闭', exact: true }).click(); await page.evaluate(() => window.pivotTableFieldFilterFixture.reopen()); await open('Date')
  await popup().getByRole('button', { name: '打开透视表日期筛选', exact: true }).click(); await expect(popup().getByRole('textbox', { name: '透视表筛选日期', exact: true })).toHaveValue('1900-02-29 06:00:00.000')
})

test.each(['date', 'label'])('ordinary clear %s removes only that condition and retains member/other condition state', async kind => {
  await page.evaluate(() => window.pivotTableFieldFilterFixture.multipleFilters('Date')); await open('Date')
  await popup().getByRole('button', { name: kind === 'date' ? '打开透视表日期筛选' : '打开透视表标签条件', exact: true }).click()
  const before = await state(); await clear(kind).click(); await waitHistory(before.history.undo + 1)
  const current = field(await state(), 'Date'); assert.ok(current.groupedItems); assert.ok(kind === 'date' ? current.labelFilter : current.dateFilter)
  assert.equal(kind === 'date' ? current.dateFilter : current.labelFilter, undefined)
  await page.evaluate(() => window.pivotTableFieldFilterFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

test.each([{ readonly: true }, { protectedOutput: true }])('ordinary viewable protection %j preserves browsing and disables edits', async options => {
  await mount(options); await open(); const before = await state()
  await popup().getByRole('textbox', { name: '搜索透视表筛选标签', exact: true }).fill('North'); await expect(member('North')).toBeDisabled(); await expect(apply()).toBeDisabled()
  await popup().getByRole('button', { name: '打开透视表标签条件', exact: true }).click(); await expect(popup().getByRole('textbox', { name: '透视表条件标签', exact: true })).toBeDisabled()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, 0)
})

test.each([{ hiddenSource: true }, { hiddenOutput: true }])('ordinary unreadable source/output %j cannot open members or conditions', async options => {
  await mount(options); await fieldMenu(); await expect(page.getByRole('menuitem', { name: '筛选字段', exact: true })).toHaveAttribute('aria-disabled', 'true')
  await expect(popup()).toHaveCount(0); await expect(main().getByRole('switch', { name: '允许字段使用多个筛选', exact: true })).toBeDisabled(); assert.equal((await state()).history.undo, 0)
})

test('ordinary 130-member pagination and search retain off-page selections', async () => {
  await mount({ manyMembers: true }); await open(); await expect(popup().locator('.member-row')).toHaveCount(80)
  await popup().locator('.ant-pagination-item-2').click(); await member('Region 129').uncheck()
  const search = popup().getByRole('textbox', { name: '搜索透视表筛选标签', exact: true }); await search.fill('Region 000'); await expect(member('Region 000')).toBeChecked()
  await search.fill(''); await expect(popup().locator('.ant-pagination-item-1')).toHaveClass(/ant-pagination-item-active/); await apply().click(); await waitHistory(1)
  assert.equal(total(await state()), 8385); assert.equal(field(await state(), 'Region').groupedItems.length, 129)
})

test.each(['source', 'history', 'user', 'host'])('ordinary pending member read drops a changed %s owner', async change => {
  await openMain(); await page.evaluate(() => window.pivotTableFieldFilterFixture.delayMembers()); await fieldMenu()
  await page.getByRole('menuitem', { name: '筛选字段', exact: true }).click(); await page.waitForFunction(() => window.pivotTableFieldFilterFixture.pendingMembers().length === 1)
  await changeOwner(change); await page.evaluate(() => window.pivotTableFieldFilterFixture.releaseMember()); await expect(popup()).toHaveCount(0)
  assert.equal(field(await state(), 'Region').groupedItems, undefined)
})

test('old ordinary member completion cannot replace a second field or clear its loading', async () => {
  await openMain(); await page.evaluate(() => window.pivotTableFieldFilterFixture.delayMembers()); await fieldMenu()
  await page.getByRole('menuitem', { name: '筛选字段', exact: true }).click(); await popup().getByRole('button', { name: '取消', exact: true }).click()
  await fieldMenu('Segment'); await page.getByRole('menuitem', { name: '筛选字段', exact: true }).click()
  await page.waitForFunction(() => window.pivotTableFieldFilterFixture.pendingMembers().length === 2)
  await page.evaluate(() => window.pivotTableFieldFilterFixture.releaseMember(0, false)); await expect(popup().getByText('正在读取透视表真实筛选成员…', { exact: true })).toBeVisible()
  await page.evaluate(() => window.pivotTableFieldFilterFixture.releaseMember()); await expect(member('A')).toBeVisible(); await expect(member('B')).toBeVisible()
})

test.each(['manual', 'date', 'label', 'sort'].flatMap(kind => ['source', 'history', 'user', 'host'].map(change => [kind, change])))
  ('pending ordinary %s rejects a changed %s context with no partial cells/history', async (kind, change) => {
    await open(kind === 'date' ? 'Date' : 'Region')
    if (kind === 'manual') await member('South').uncheck()
    if (kind === 'date') await dateDraft()
    if (kind === 'label') await labelDraft()
    await page.evaluate(kind => window.pivotTableFieldFilterFixture.delayEdits([kind]), kind)
    if (kind === 'sort') await popup().getByRole('button', { name: '透视表字段降序', exact: true }).click()
    else await apply(kind === 'manual' ? 'members' : kind).click()
    await page.waitForFunction(() => window.pivotTableFieldFilterFixture.pendingEdits().length === 1); await changeOwner(change); const before = await state()
    await page.evaluate(() => window.pivotTableFieldFilterFixture.releaseEdit()); assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  })

test('old ordinary content finally cannot clear a newer busy edit after a history reread', async () => {
  await open(); await member('South').uncheck(); await page.evaluate(() => window.pivotTableFieldFilterFixture.delayEdits(['manual'])); await apply().click()
  await page.waitForFunction(() => window.pivotTableFieldFilterFixture.pendingEdits().length === 1); await page.evaluate(() => window.pivotTableFieldFilterFixture.clearHistory()); await reread()
  await open('Segment'); await member('B').uncheck(); await apply().click(); await page.waitForFunction(() => window.pivotTableFieldFilterFixture.pendingEdits().length === 2)
  await page.evaluate(() => window.pivotTableFieldFilterFixture.releaseEdit(0, false)); await expect(apply()).toHaveClass(/ant-btn-loading/)
  await page.evaluate(() => window.pivotTableFieldFilterFixture.releaseEdit()); await waitHistory(1); assert.equal(total(await state()), 30)
})

test.each(['value', 'dateGroup'].flatMap(kind => ['source', 'history', 'user', 'host'].map(change => [kind, change])))
  ('ordinary tool %s rejects changed %s owner through the same report transaction guard', async (kind, change) => {
    await openTools(kind === 'dateGroup' ? 'Date' : 'Region', kind === 'value' ? 'filter' : 'group')
    if (kind === 'dateGroup') { await choose('分组方式', '日期分组', tools()); await choose('日期分组依据', '年', tools()) }
    else await tools().getByRole('spinbutton', { name: '筛选数值', exact: true }).fill('50')
    await page.evaluate(kind => window.pivotTableFieldFilterFixture.delayEdits([kind]), kind); await tools().getByRole('button', { name: '确定', exact: true }).click()
    await page.waitForFunction(() => window.pivotTableFieldFilterFixture.pendingEdits().length === 1); await changeOwner(change); const before = await state()
    await page.evaluate(() => window.pivotTableFieldFilterFixture.releaseEdit()); assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  })

test.each(['classic', 'grid'])('%s ordinary filter theme/IME preserves a stale input and refuses actions until composition ends', async variant => {
  await mount({ variant }); await open(); const input = popup().getByRole('textbox', { name: '搜索透视表筛选标签', exact: true })
  await input.focus(); await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await expect(apply()).toBeDisabled(); await expect(popup().getByRole('button', { name: '取消', exact: true })).toBeDisabled()
  const light = await popup().evaluate(node => getComputedStyle(node).color); await page.evaluate(variant => window.pivotTableFieldFilterFixture.appearance(true, variant), variant)
  assert.notEqual(await popup().evaluate(node => getComputedStyle(node).color), light)
  await input.evaluate(node => { node.value = '中'; node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '中', isComposing: true, inputType: 'insertCompositionText' })) })
  await page.evaluate(() => window.pivotTableFieldFilterFixture.changeSource()); await expect(input).toHaveValue('中'); await expect(input).toBeEnabled()
  await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中' }))); await expect(popup()).toHaveCount(0)
})

test('ordinary manual-group composition preserves the Ant input while stale apply/close are blocked', async () => {
  await openTools('Region'); const input = tools().getByRole('textbox', { name: '新分组名称', exact: true })
  await input.focus(); await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await expect(tools().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); await expect(tools().getByRole('button', { name: '取消', exact: true })).toBeDisabled()
  await input.evaluate(node => { node.value = '分组'; node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '分组', isComposing: true, inputType: 'insertCompositionText' })) })
  const light = await tools().evaluate(node => getComputedStyle(node).color); await page.evaluate(() => window.pivotTableFieldFilterFixture.appearance(true)); assert.notEqual(await tools().evaluate(node => getComputedStyle(node).color), light)
  await page.evaluate(() => window.pivotTableFieldFilterFixture.changeSource()); await expect(input).toHaveValue('分组'); await expect(input).toBeEnabled()
  await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '分组' })))
  await expect(tools().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); await tools().getByRole('button', { name: '取消', exact: true }).click()
  assert.equal(field(await state(), 'Region').manualGroups, undefined)
})

test('ordinary picker Escape stays in the field popup; outer Escape closes and restores the field trigger', async () => {
  const item = await open('Date'); await popup().getByRole('button', { name: '打开透视表日期筛选', exact: true }).click()
  const date = popup().locator('.ant-picker input').first(); await date.click(); await date.press('Escape'); await expect(popup()).toBeVisible()
  await popup().getByRole('button', { name: '成员筛选', exact: true }).click(); await popup().getByRole('textbox', { name: '搜索透视表筛选标签', exact: true }).press('Escape')
  await expect(popup()).toHaveCount(0); await expect(main().locator(`.pivot-area-grid [data-field-id="${item.id}"]`).getByRole('button', { name: '字段 Date 操作', exact: true })).toBeFocused()
})

test('disposing the actual workbench rejects late ordinary members without attaching a panel', async () => {
  await openMain(); await page.evaluate(() => window.pivotTableFieldFilterFixture.delayMembers()); await fieldMenu(); await page.getByRole('menuitem', { name: '筛选字段', exact: true }).click()
  await page.waitForFunction(() => window.pivotTableFieldFilterFixture.pendingMembers().length === 1); await page.evaluate(() => window.pivotTableFieldFilterFixture.dispose())
  await expect(popup()).toHaveCount(0); await expect(page.locator('#app')).toBeEmpty()
})
