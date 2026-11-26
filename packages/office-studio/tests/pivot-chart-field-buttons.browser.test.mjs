import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'
import { pivotChartFieldButtonsIds as ids } from './fixtures/pivot-chart-field-buttons-fixture.mjs'

// Deferred drafts for production ExcelWorkbench -> ChartOverlay -> Ant field controls.
// Nothing has run. Canvas2D is selected explicitly; synthetic composition is not native IME evidence.
let server, browser, context, page, url
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No field-control fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/pivot-chart-field-buttons-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1200 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.pivotChartFieldButtonsFixture?.state().model.charts.length === 3)
}, 60_000)
afterEach(async () => {
  try { await page?.evaluate(() => window.pivotChartFieldButtonsFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.pivotChartFieldButtonsFixture.state())
const root = (chartId = ids.owned) => page.locator(`[data-testid="pivot-chart-field-buttons"][data-pivot-chart="${chartId}"]`)
const popup = () => page.getByTestId('pivot-chart-field-filter-popup')
const apply = (kind = 'members') => popup().getByRole('button', { name: kind === 'date' ? '应用透视图日期筛选'
  : kind === 'label' ? '应用透视图标签条件' : '应用透视图字段筛选', exact: true })
const clear = (kind = 'members') => popup().getByRole('button', { name: kind === 'date' ? '清除透视图日期条件'
  : kind === 'label' ? '清除透视图标签条件' : '清除透视图字段筛选', exact: true })
const member = key => popup().getByRole('checkbox', { name: `选择透视图标签${key}`, exact: true })
const allMembers = () => popup().getByRole('checkbox', { name: '全选所有项目', exact: true })
const report = (value, chartId = value.chartId) => value.model.charts.find(chart => chart.id === chartId)?.ownedPivot
  ?? value.model.pivotTables.find(pivot => pivot.id === value.model.charts.find(chart => chart.id === chartId)?.pivotTableId)
const field = (value, name) => report(value).fields.find(candidate => candidate.name === name)
const total = value => value.data.series.flatMap(series => series.values).reduce((sum, value) => sum + (typeof value === 'number' ? value : 0), 0)
const sourceCells = value => value.model.sheets.find(sheet => sheet.id === ids.source).cells
async function mount(options = {}) { await page.evaluate(options => window.pivotChartFieldButtonsFixture.mount(options), options) }
async function open(name = 'Region', area, chartId = ids.owned) {
  await page.evaluate(chartId => window.pivotChartFieldButtonsFixture.selectChart(chartId), chartId)
  const item = await page.evaluate(({ name, area, chartId }) => window.pivotChartFieldButtonsFixture.field(name, area, chartId), { name, area, chartId })
  await root(chartId).locator(`[data-pivot-field="${item.id}"]`).click(); await expect(popup()).toBeVisible()
  if (item.area !== 'value') await expect(popup().getByRole('textbox', { name: '搜索透视图筛选标签', exact: true })).toBeEnabled()
  return item
}
async function waitHistory(count) { await page.waitForFunction(count => window.pivotChartFieldButtonsFixture.state().history.undo === count, count) }
async function chooseCondition(kind, label) {
  await popup().getByRole('combobox', { name: kind === 'date' ? '透视图日期条件' : '透视图标签条件', exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}
async function dateDraft(label = '等于', first = '2026-02-02 00:00:00.000', second) {
  await popup().getByRole('button', { name: '打开透视图日期筛选', exact: true }).click(); await chooseCondition('date', label)
  const inputs = popup().locator('.ant-picker input'); await inputs.nth(0).fill(first); await inputs.nth(0).press('Tab')
  if (second !== undefined) { await inputs.nth(1).fill(second); await inputs.nth(1).press('Tab') }
}
async function labelDraft(label = '等于', first = 'North', second) {
  await popup().getByRole('button', { name: '打开透视图标签条件', exact: true }).click(); await chooseCondition('label', label)
  await popup().getByRole('textbox', { name: '透视图条件标签', exact: true }).fill(first)
  if (second !== undefined) await popup().getByRole('textbox', { name: '透视图条件结束标签', exact: true }).fill(second)
}
async function changeOwner(kind) {
  await page.evaluate(async kind => {
    const fixture = window.pivotChartFieldButtonsFixture
    if (kind === 'source') await fixture.changeSource()
    if (kind === 'history') await fixture.clearHistory()
    if (kind === 'user') await fixture.changeUserWhilePending()
    if (kind === 'host') await fixture.switchSheet()
  }, kind)
}

test.each(['classic', 'grid'])('%s workbench exposes four authored field areas through the actual chart overlay', async variant => {
  await mount({ variant }); const before = await state()
  assert.deepEqual(await root().locator('[data-pivot-button-area]').evaluateAll(nodes => nodes.map(node => node.dataset.pivotButtonArea)), ['filter', 'column', 'row', 'value'])
  for (const [area, name] of [['filter', 'Date'], ['column', 'Segment'], ['row', 'Region'], ['value', 'Amount']]) {
    await expect(root().locator(`[data-pivot-button-area="${area}"]`).getByRole('button', { name: `打开透视图字段${name}`, exact: true })).toBeVisible()
  }
  const positions = await root().locator('[data-pivot-button-area]').evaluateAll(nodes => nodes.map(node => ({ area: node.dataset.pivotButtonArea,
    x: node.getBoundingClientRect().x, y: node.getBoundingClientRect().y })))
  assert.ok(positions.find(value => value.area === 'row').y > positions.find(value => value.area === 'filter').y)
  assert.ok(positions.find(value => value.area === 'column').x > positions.find(value => value.area === 'row').x)
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, 0)
})

test('overflow contains every same-area value instance, including the two already visible buttons', async () => {
  await mount({ overflow: true }); await root().getByRole('button', { name: '更多值字段', exact: true }).click()
  const overflow = page.getByTestId('pivot-chart-field-overflow')
  assert.deepEqual(await overflow.getByRole('button').allTextContents(), ['Amount', 'Amount average', 'Amount count'])
  const keys = await overflow.getByRole('button').evaluateAll(nodes => nodes.map(node => node.dataset.pivotField)); assert.equal(new Set(keys).size, 3)
  await overflow.getByRole('button', { name: '打开透视图字段Amount count', exact: true }).click(); await expect(popup()).toBeVisible()
  await expect(popup().getByRole('textbox', { name: '搜索透视图筛选标签', exact: true })).toHaveCount(0)
  assert.equal((await state()).history.undo, 0)
})

test.each([ids.owned, ids.linked])('%s manual confirmation uses one content command, real totals, history and snapshot reopen', async chartId => {
  await open('Region', 'row', chartId); const before = await state(); await member('South').uncheck()
  assert.deepEqual((await state()).model, before.model); await expect(apply()).toBeEnabled(); await apply().click(); await waitHistory(1)
  const saved = await state(); assert.equal(total(saved), 40); assert.deepEqual(sourceCells(saved), sourceCells(before))
  assert.equal(saved.commands.filter(command => command.id === 'sheet.command.edit-pivot-chart-report').length, 1)
  if (chartId === ids.owned) assert.deepEqual(saved.model.sheets.find(sheet => sheet.id === ids.host).cells, before.model.sheets.find(sheet => sheet.id === ids.host).cells)
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeFocused()
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.undo()); assert.deepEqual((await state()).model, before.model)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.redo()); assert.deepEqual((await state()).model, saved.model)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.reopen()); await open('Region', 'row', chartId)
  await expect(member('North')).toBeChecked(); await expect(member('South')).not.toBeChecked(); assert.equal(total(await state()), 40)
})

test('cancel retains content and history and restores focus to the selected chart workbench', async () => {
  await open(); const before = await state(); await member('South').uncheck(); await popup().getByRole('button', { name: '取消', exact: true }).click()
  await expect(popup()).toHaveCount(0); assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeFocused()
})

test('empty manual selection cannot confirm; all includes search-hidden members', async () => {
  await open(); const before = await state(); await allMembers().uncheck(); await expect(apply()).toBeDisabled()
  await popup().getByRole('textbox', { name: '搜索透视图筛选标签', exact: true }).fill('North'); await allMembers().check()
  await expect(apply()).toBeEnabled(); await apply().click(); assert.equal(total(await state()), 100)
  assert.deepEqual(sourceCells(await state()), sourceCells(before)); assert.equal(field(await state(), 'Region').recordSelection, undefined)
})

test('all clears manual only while the clear action removes manual, caption and value filters', async () => {
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.conditionFilter()); await open(); const initial = await state()
  await allMembers().check(); await apply().click(); await waitHistory(initial.history.undo + 1)
  let current = await state(); assert.ok(field(current, 'Region').labelFilter); assert.ok(field(current, 'Region').valueFilter)
  assert.equal(field(current, 'Region').groupedItems, undefined); assert.equal(total(current), 60)
  await open(); await expect(clear()).toBeEnabled(); const before = await state(); await clear().click(); await waitHistory(before.history.undo + 1)
  current = await state(); for (const key of ['selectedItems', 'groupedItems', 'partialItems', 'recordSelection', 'labelFilter', 'dateFilter', 'valueFilter']) assert.equal(field(current, 'Region')[key], undefined)
  assert.equal(total(current), 100); await page.evaluate(() => window.pivotChartFieldButtonsFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

const partialCases = [[{ partialDate: true }, 'Date', '2026', 100], [{ typedMembers: true }, 'Region', '1', 30]]
test.each(partialCases)('unchanged grouped partial %j survives confirmation and reopen without widening raw membership', async (options, name, key) => {
  await mount(options); await open(name); await expect(popup().locator('.ant-checkbox-indeterminate')).toHaveCount(2)
  await expect(member(key)).not.toBeChecked(); await expect(apply()).toBeEnabled(); await apply().click(); await waitHistory(1)
  assert.equal(total(await state()), 10); await page.evaluate(() => window.pivotChartFieldButtonsFixture.reopen()); await open(name)
  await expect(member(key)).not.toBeChecked(); await expect(popup().getByText(/部分选中 1 项/)).toBeVisible(); assert.equal(total(await state()), 10)
})
test.each(partialCases)('checking partial group %j deliberately includes every raw member represented by that key', async (options, name, key, expected) => {
  await mount(options); await open(name); await member(key).check(); await expect(member(key)).toBeChecked()
  await apply().click(); await waitHistory(1); assert.equal(total(await state()), expected); assert.equal(field(await state(), name).partialItems, undefined)
})

test.each([['透视图字段升序', 'ascending', ['North', 'South']], ['透视图字段降序', 'descending', ['South', 'North']], ['清除透视图字段排序', undefined, ['North', 'South']]])
  ('%s changes the real chart category order and is undoable', async (label, direction, categories) => {
    if (direction === undefined) { await open(); await popup().getByRole('button', { name: '透视图字段降序', exact: true }).click(); await waitHistory(1) }
    await open(); const before = await state(); await popup().getByRole('button', { name: label, exact: true }).click(); await waitHistory(before.history.undo + 1)
    assert.equal(field(await state(), 'Region').sort, direction); assert.deepEqual((await state()).data.categories, categories)
    await page.evaluate(() => window.pivotChartFieldButtonsFixture.undo()); assert.deepEqual((await state()).model, before.model)
  })

test('remove targets only the selected duplicate value instance and retains chart appearance', async () => {
  await mount({ overflow: true }); await root().getByRole('button', { name: '更多值字段', exact: true }).click()
  await page.getByTestId('pivot-chart-field-overflow').getByRole('button', { name: '打开透视图字段Amount average', exact: true }).click()
  const before = await state(); await popup().getByRole('button', { name: '移除透视图字段', exact: true }).click(); await waitHistory(1)
  const current = await state(); assert.deepEqual(report(current).fields.filter(field => field.area === 'value').map(field => field.displayName ?? field.name), ['Amount', 'Amount count'])
  assert.deepEqual(current.model.charts.find(chart => chart.id === ids.owned).options, before.model.charts.find(chart => chart.id === ids.owned).options)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

test.each([ids.owned, ids.linked].flatMap(chartId => ['Date', 'Segment', 'Region', 'Amount'].map(name => [chartId, name])))
  ('%s %s settings route selects the actual report-local card in the existing panel', async (chartId, name) => {
    const item = await open(name, undefined, chartId), before = await state()
    await popup().getByRole('button', { name: '透视图字段设置', exact: true }).click()
    const panel = page.getByTestId('pivot-chart-panel'); await expect(panel).toBeVisible(); await expect(panel.getByRole('tab', { name: '字段', exact: true })).toHaveAttribute('aria-selected', 'true')
    await expect(panel.locator(`[data-pivot-field="${item.id}"]`)).toBeVisible()
    assert.ok(await panel.locator(`[data-pivot-field="${item.id}"]`).evaluate(node => node.contains(document.activeElement)))
    assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  })

test.each([[{ readonly: true }, ids.owned], [{ protectedChart: true }, ids.owned], [{ protectedOutput: true }, ids.linked]])
  ('viewable protection %j retains browsing and disables member, condition and content actions', async (options, chartId) => {
    await mount({ ...options, chartId }); await open('Region', 'row', chartId); const before = await state()
    await popup().getByRole('textbox', { name: '搜索透视图筛选标签', exact: true }).fill('North'); await expect(member('North')).toBeDisabled()
    await expect(apply()).toBeDisabled(); await expect(popup().getByRole('button', { name: '移除透视图字段', exact: true })).toBeDisabled()
    await popup().getByRole('button', { name: '打开透视图标签条件', exact: true }).click(); await expect(popup().getByRole('textbox', { name: '透视图条件标签', exact: true })).toBeDisabled()
    assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, 0)
  })

test.each([[{ hiddenSource: true }, ids.owned], [{ hiddenSource: true }, ids.linked], [{ hiddenOutput: true }, ids.linked]])
  ('hidden source/output %j exposes no field controls', async (options, chartId) => {
    await mount({ ...options, chartId }); await expect(root(chartId)).toHaveCount(0); await expect(popup()).toHaveCount(0); assert.equal((await state()).history.undo, 0)
  })

test('130 members paginate by 80 and searching a hidden member keeps every other selection', async () => {
  await mount({ manyMembers: true }); await open(); await expect(popup().locator('.member-row')).toHaveCount(80)
  await popup().locator('.ant-pagination-item-2').click(); await expect(member('Region 129')).toBeVisible(); await member('Region 129').uncheck()
  const search = popup().getByRole('textbox', { name: '搜索透视图筛选标签', exact: true }); await search.fill('Region 000')
  await expect(popup().locator('.member-row')).toHaveCount(1); await expect(member('Region 000')).toBeChecked(); await search.fill('')
  await expect(popup().locator('.ant-pagination-item-1')).toHaveClass(/ant-pagination-item-active/); await apply().click(); await waitHistory(1)
  assert.equal(total(await state()), 8385); assert.equal(field(await state(), 'Region').groupedItems.length, 129)
})

test.each([[true, true], [true, false], [false, true], [false, false]])('controls [%s,%s] drive actual field visibility independently of expand metadata', async (showFieldButtons, showExpandCollapseButtons) => {
  await open(); await page.evaluate(controls => window.pivotChartFieldButtonsFixture.setControls(controls), { showFieldButtons, showExpandCollapseButtons })
  await expect(popup()).toHaveCount(0); if (showFieldButtons) await expect(root()).toBeVisible(); else await expect(root()).toHaveCount(0)
  assert.deepEqual((await state()).model.charts.find(chart => chart.id === ids.owned).controls, { showFieldButtons, showExpandCollapseButtons })
})

test.each([
  ['dateBetween', '介于', '2026-02-02', '2026-02-20', 50], ['dateEqual', '等于', '2026-02-02', undefined, 20],
  ['dateNewerThan', '晚于', '2026-02-02', undefined, 70], ['dateNewerThanOrEqual', '晚于或等于', '2026-02-02', undefined, 90],
  ['dateNotBetween', '不介于', '2026-02-02', '2026-02-20', 50], ['dateNotEqual', '不等于', '2026-02-02', undefined, 80],
  ['dateOlderThan', '早于', '2026-02-02', undefined, 10], ['dateOlderThanOrEqual', '早于或等于', '2026-02-02', undefined, 30],
])('Ant calendar %s filters real records and writes one typed report transaction', async (operator, label, first, second, expected) => {
  await open('Date'); const before = await state(); await dateDraft(label, `${first} 00:00:00.000`, second && `${second} 00:00:00.000`)
  await expect(apply('date')).toBeEnabled(); assert.deepEqual((await state()).model, before.model); await apply('date').click(); await waitHistory(1)
  assert.equal(total(await state()), expected); assert.equal(field(await state(), 'Date').dateFilter.operator, operator)
  assert.equal((await state()).commands.filter(command => command.id === 'sheet.command.edit-pivot-chart-report').length, 1)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

test.each([
  ['captionBeginsWith', '开头是', 'N', undefined, 40], ['captionBetween', '介于', 'North', 'South', 100],
  ['captionContains', '包含', 'or', undefined, 40], ['captionEndsWith', '结尾是', 'orth', undefined, 40], ['captionEqual', '等于', 'North', undefined, 40],
  ['captionGreaterThan', '大于', 'North', undefined, 60], ['captionGreaterThanOrEqual', '大于或等于', 'North', undefined, 100],
  ['captionLessThan', '小于', 'South', undefined, 40], ['captionLessThanOrEqual', '小于或等于', 'South', undefined, 100],
  ['captionNotBeginsWith', '开头不是', 'N', undefined, 60], ['captionNotBetween', '不介于', 'North', 'South', 0],
  ['captionNotContains', '不包含', 'or', undefined, 60], ['captionNotEndsWith', '结尾不是', 'orth', undefined, 60], ['captionNotEqual', '不等于', 'North', undefined, 60],
])('Ant label %s uses shared caption semantics and updates true chart totals', async (operator, label, first, second, expected) => {
  await open(); await labelDraft(label, first, second); await expect(apply('label')).toBeEnabled(); await apply('label').click(); await waitHistory(1)
  assert.equal(total(await state()), expected); assert.equal(field(await state(), 'Region').labelFilter.operator, operator)
})

test.each([false, true])('calendar date1904=%s and whole-day comparison retain workbook numeric operands', async date1904 => {
  await mount({ date1904, dateSerials: true }); const initial = await state(), serial = sourceCells(initial)['5:5'].value
  await page.evaluate(async serial => { await window.pivotChartFieldButtonsFixture.sourceCell('F4', serial); await window.pivotChartFieldButtonsFixture.sourceCell('F5', serial + .5) }, serial)
  await open('Date'); await dateDraft('等于', '2026-02-20 00:00:00.000'); await apply('date').click()
  assert.equal(total(await state()), 40); assert.equal(field(await state(), 'Date').dateFilter.expected, serial)
  await open('Date'); await popup().getByRole('button', { name: '打开透视图日期筛选', exact: true }).click()
  await popup().getByRole('checkbox', { name: '按整天比较', exact: true }).check(); await apply('date').click()
  assert.equal(total(await state()), 60); assert.equal(field(await state(), 'Date').dateFilter.expected, serial)
})

test('existing Excel day 60 and fractional serial remain exact through Ant compatibility input and snapshot reopen', async () => {
  await mount({ compatibilityDate: true }); await page.evaluate(() => window.pivotChartFieldButtonsFixture.dateFilter({ operator: 'dateEqual', expected: 60.25 }))
  await open('Date'); await popup().getByRole('button', { name: '打开透视图日期筛选', exact: true }).click()
  await expect(popup().getByRole('textbox', { name: '透视图筛选日期', exact: true })).toHaveValue('1900-02-29 06:00:00.000')
  const before = await state(); await apply('date').click(); assert.equal(field(await state(), 'Date').dateFilter.expected, 60.25); assert.equal(total(await state()), 40)
  assert.equal((await state()).history.undo, before.history.undo); await page.evaluate(() => window.pivotChartFieldButtonsFixture.reopen()); await open('Date')
  await popup().getByRole('button', { name: '打开透视图日期筛选', exact: true }).click(); await expect(popup().getByRole('textbox', { name: '透视图筛选日期', exact: true })).toHaveValue('1900-02-29 06:00:00.000')
})

test('default single-category filtering replaces manual state; the real multiple-filter option permits intersection', async () => {
  await open(); await member('South').uncheck(); await apply().click(); await waitHistory(1)
  await open(); await labelDraft('等于', 'South'); await apply('label').click(); await waitHistory(2)
  assert.equal(field(await state(), 'Region').groupedItems, undefined); assert.equal(total(await state()), 60)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.setOptions({ allowMultipleFilters: true })); await open(); await member('South').uncheck(); await apply().click()
  assert.ok(field(await state(), 'Region').labelFilter); assert.equal(total(await state()), 0)
})

test.each(['date', 'label'])('clear %s condition retains other allowed categories and remains undoable', async kind => {
  await page.evaluate(async () => {
    const fixture = window.pivotChartFieldButtonsFixture; await fixture.setOptions({ allowMultipleFilters: true })
    await fixture.dateFilter({ operator: 'dateOlderThanOrEqual', expected: 46073 }); await fixture.labelCondition({ operator: 'captionContains', expected: '2026-' }, undefined, 'Date')
  })
  await open('Date'); await member('2026-01-15').uncheck(); await apply().click(); await open('Date')
  await popup().getByRole('button', { name: kind === 'date' ? '打开透视图日期筛选' : '打开透视图标签条件', exact: true }).click()
  const before = await state(); await clear(kind).click(); await waitHistory(before.history.undo + 1)
  const current = field(await state(), 'Date'); assert.ok(current.groupedItems); assert.equal(kind === 'date' ? current.dateFilter : current.labelFilter, undefined)
  assert.ok(kind === 'date' ? current.labelFilter : current.dateFilter); await page.evaluate(() => window.pivotChartFieldButtonsFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

test.each(['source', 'history', 'user', 'host'])('pending member read discards a changed %s owner', async change => {
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.delayMembers())
  await root().getByRole('button', { name: '打开透视图字段Region', exact: true }).click()
  await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingMembers().length === 1); await changeOwner(change)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.releaseMember()); await expect(popup()).toHaveCount(0)
  assert.equal(field(await state(), 'Region').groupedItems, undefined)
})

test('an older member completion neither replaces a newer field nor clears its loading indicator', async () => {
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.delayMembers()); await root().getByRole('button', { name: '打开透视图字段Region', exact: true }).click()
  await root().getByRole('button', { name: '打开透视图字段Segment', exact: true }).click(); await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingMembers().length === 2)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.releaseMember(0, false)); const item = await page.evaluate(() => window.pivotChartFieldButtonsFixture.field('Segment'))
  await expect(popup()).toHaveAttribute('data-pivot-chart-field', item.id); await expect(popup().getByText('正在读取透视图真实筛选成员…', { exact: true })).toBeVisible()
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.releaseMember()); await expect(member('A')).toBeVisible(); await expect(member('B')).toBeVisible()
})

test.each(['members', 'date', 'label'].flatMap(kind => ['source', 'history', 'user', 'host'].map(change => [kind, change])))
  ('pending %s edit rejects a changed %s context without partial content', async (kind, change) => {
    await open(kind === 'date' ? 'Date' : 'Region')
    if (kind === 'members') await member('South').uncheck()
    if (kind === 'date') await dateDraft()
    if (kind === 'label') await labelDraft()
    await page.evaluate(() => window.pivotChartFieldButtonsFixture.delayEdits()); await apply(kind).click()
    await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingEdits().length === 1); await changeOwner(change)
    const before = await state(); await page.evaluate(() => window.pivotChartFieldButtonsFixture.releaseEdit())
    assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  })

test('an old content finally cannot clear the next field busy state after history generation changes', async () => {
  await open(); await member('South').uncheck(); await page.evaluate(() => window.pivotChartFieldButtonsFixture.delayEdits()); await apply().click()
  await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingEdits().length === 1)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.clearHistory()); await open('Segment'); await member('B').uncheck(); await apply().click()
  await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingEdits().length === 2)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.releaseEdit(0, false)); await expect(apply()).toHaveClass(/ant-btn-loading/)
  await expect(popup().getByRole('button', { name: '取消', exact: true })).toBeDisabled()
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.releaseEdit()); await waitHistory(1); assert.equal(total(await state()), 30)
})

test.each(['classic', 'grid'])('%s theme and synthetic composition preserve input while stale content cannot publish', async variant => {
  await mount({ variant }); await open(); const input = popup().getByRole('textbox', { name: '搜索透视图筛选标签', exact: true })
  await input.focus(); await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await expect(apply()).toBeDisabled(); await expect(popup().getByRole('button', { name: '取消', exact: true })).toBeDisabled()
  const light = await popup().evaluate(node => getComputedStyle(node).color); await page.evaluate(variant => window.pivotChartFieldButtonsFixture.appearance(true, variant), variant)
  assert.notEqual(await popup().evaluate(node => getComputedStyle(node).color), light)
  await input.evaluate(node => { node.value = '中'; node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '中', isComposing: true, inputType: 'insertCompositionText' })) })
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.changeSource()); await expect(input).toHaveValue('中'); await expect(input).toBeEnabled()
  await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中' }))); await expect(popup()).toHaveCount(0)
})

test('picker-handled Escape retains the field popup and outer Escape closes it with focus restoration', async () => {
  await open('Date'); await popup().getByRole('button', { name: '打开透视图日期筛选', exact: true }).click()
  const date = popup().locator('.ant-picker input').first(); await date.click(); await date.press('Escape'); await expect(popup()).toBeVisible()
  await popup().getByRole('button', { name: '成员筛选', exact: true }).click()
  await popup().getByRole('textbox', { name: '搜索透视图筛选标签', exact: true }).press('Escape'); await expect(popup()).toHaveCount(0)
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeFocused()
})

test('disposing the real workbench while a member read waits does not attach the old popup', async () => {
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.delayMembers()); await root().getByRole('button', { name: '打开透视图字段Region', exact: true }).click()
  await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingMembers().length === 1)
  await page.evaluate(() => window.pivotChartFieldButtonsFixture.dispose()); await expect(popup()).toHaveCount(0); await expect(page.locator('#app')).toBeEmpty()
})
