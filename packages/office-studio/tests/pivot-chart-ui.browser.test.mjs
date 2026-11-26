import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// Deferred real Ant/Workbench/Canvas2D/Facade/history drafts; no hook, server or browser has been run.
// Synthetic composition events cover capture ordering only. Windows IME and WebGPU remain unverified.
let server, browser, context, page, url
const errors = []
const ids = { unit: 'pivot-chart-ui-book', source: 'data', host: 'reports', other: 'dashboard',
  primary: 'ui-primary-report', secondary: 'ui-secondary-report', owned: 'ui-owned-chart', linked: 'ui-linked-chart', foreign: 'ui-foreign-chart' }
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/pivot-chart-ui-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1100 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.pivotChartUIFixture?.state().model.charts.length === 3)
}, 60_000)
afterEach(async () => {
  try { await page?.evaluate(() => window.pivotChartUIFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.pivotChartUIFixture.state())
const dialog = () => page.getByRole('dialog', { name: '透视图', exact: true })
const filter = () => page.getByRole('dialog', { name: '标签筛选', exact: true })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const apply = () => dialog().getByRole('button', { name: '应用透视图', exact: true })
const report = (value, id = ids.owned) => value.model.charts.find(chart => chart.id === id)?.ownedPivot
  ?? value.model.pivotTables.find(pivot => pivot.id === value.model.charts.find(chart => chart.id === id)?.pivotTableId)
const field = (value, sourceId, chartId = ids.owned) => report(value, chartId)?.fields.find(field => field.sourceFieldId === sourceId)
async function open(chartId = ids.owned, wait = true) {
  await page.evaluate(chartId => window.pivotChartUIFixture.open(chartId), chartId); await expect(dialog()).toBeVisible()
  if (wait) await expect(dialog().getByRole('textbox', { name: 'Region的显示名称', exact: true })).toBeVisible()
}
async function choose(name, label, target = dialog()) {
  await target.getByRole('combobox', { name, exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}
async function appearance() { await dialog().getByRole('tab', { name: '图表设置', exact: true }).click() }
async function fieldsTab() { await dialog().getByRole('tab', { name: '字段', exact: true }).click() }
async function titleInput() { await appearance(); return dialog().getByRole('textbox', { name: '图表标题', exact: true }) }
async function waitHistory(count) { await page.waitForFunction(count => window.pivotChartUIFixture.state().history.undo === count, count) }
async function close() { await dialog().getByRole('button', { name: '关闭', exact: true }).click(); await expect(dialog()).not.toBeVisible() }

for (const variant of ['classic', 'grid']) test(`${variant} opens range authoring from the real Ribbon without an existing PivotTable`, async () => {
  await page.evaluate(variant => window.pivotChartUIFixture.mount({ existing: false, variant }), variant)
  await page.evaluate(() => window.pivotChartUIFixture.selectSourceRange()); const before = await state()
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  const root = page.getByRole('toolbar', { name: '数据', exact: true }).getByRole('button', { name: '透视图', exact: true })
  await expect(root).toBeEnabled(); await root.click(); await expect(dialog()).toBeVisible()
  await expect(dialog().getByRole('textbox', { name: '透视图来源区域', exact: true })).toHaveValue('D3:H7')
  for (const area of ['筛选', '轴（类别）', '图例（系列）', '值']) await expect(dialog().getByRole('combobox', { name: `添加${area}字段`, exact: true })).toBeVisible()
  await choose('添加轴（类别）字段', 'Region'); await choose('添加值字段', 'Amount')
  const create = dialog().getByRole('button', { name: '创建透视图', exact: true }); await expect(create).toBeEnabled(); await create.click()
  await waitHistory(before.history.undo + 1)
  const after = await state(), chart = after.model.charts[0]
  assert.ok(chart.ownedPivot); assert.equal(chart.pivotTableId, undefined); assert.equal(chart.sheetId, ids.source)
  assert.deepEqual(chart.ownedPivot.fields.map(field => [field.sourceFieldId, field.area]), [['0', 'row'], ['1', 'value']])
  assert.equal(after.model.pivotTables.length, 0); assert.deepEqual(after.model.sheets.map(sheet => sheet.cells), before.model.sheets.map(sheet => sheet.cells))
  assert.equal(after.commands.filter(command => command.id === 'sheet.command.add-pivot-chart').length, 1)
})

test('the selector lists owned and linked charts in this host and changing selection reads their true configurations', async () => {
  const before = await state(); await open()
  await dialog().getByRole('combobox', { name: '选择透视图', exact: true }).click()
  const list = page.locator('.ant-select-dropdown:visible')
  await expect(list.getByText('Owned chart · 独立报告', { exact: true })).toBeVisible()
  await expect(list.getByText('Linked chart · 关联透视表', { exact: true })).toBeVisible()
  await expect(list.getByText('Other sheet chart · 关联透视表', { exact: true })).toHaveCount(0)
  await list.getByText('Linked chart · 关联透视表', { exact: true }).click()
  await expect(dialog().getByRole('combobox', { name: '关联透视表', exact: true })).toBeDisabled()
  await choose('选择透视图', 'Owned chart · 独立报告')
  await expect(dialog().getByRole('textbox', { name: '透视图来源区域', exact: true })).toHaveValue('D3:H7')
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('choosing new after editing the real source range retains the current range instead of reverting to the opening selection', async () => {
  await open(); await choose('选择透视图', '新建透视图')
  const range = dialog().getByRole('textbox', { name: '透视图来源区域', exact: true })
  await range.fill('D3:G7'); await dialog().getByRole('button', { name: '读取来源字段', exact: true }).click()
  await choose('选择透视图', '新建透视图'); await expect(range).toHaveValue('D3:G7')
  assert.equal((await state()).history.undo, 0)
})

test('new linked authoring selects a real PivotTable and creates its chart without duplicating the report', async () => {
  await open(); await choose('选择透视图', '新建透视图'); await choose('透视图来源方式', '关联已有透视表')
  await expect(dialog().getByRole('textbox', { name: 'Amount的显示名称', exact: true })).toHaveValue('Amount')
  const before = await state(), create = dialog().getByRole('button', { name: '创建透视图', exact: true })
  await expect(create).toBeEnabled(); await create.click(); await waitHistory(before.history.undo + 1)
  const after = await state(), created = after.model.charts.find(chart => !before.model.charts.some(previous => previous.id === chart.id))
  assert.equal(created.pivotTableId, ids.primary); assert.equal(created.ownedPivot, undefined)
  assert.equal(after.model.pivotTables.length, before.model.pivotTables.length)
  assert.equal(after.commands.filter(command => command.id === 'sheet.command.add-pivot-chart').length, 1)
})

for (const chartId of [ids.owned, ids.linked]) test(`${chartId} applies field aggregation and chart title together in one command and preserves undo/save/reopen`, async () => {
  await open(chartId); const before = await state()
  await choose('Amount的汇总方式', '平均值')
  await dialog().getByRole('textbox', { name: 'Amount的显示名称', exact: true }).fill('Mean revenue')
  await dialog().getByRole('textbox', { name: 'Amount的数字格式', exact: true }).fill('0.00')
  const title = await titleInput(); await title.fill('Average revenue'); await expect(apply()).toBeEnabled(); await apply().click()
  await waitHistory(before.history.undo + 1); const after = await state()
  assert.equal(field(after, '1', chartId).aggregation, 'average'); assert.equal(field(after, '1', chartId).displayName, 'Mean revenue')
  assert.equal(field(after, '1', chartId).format, '0.00'); assert.equal(after.model.charts.find(chart => chart.id === chartId).title, 'Average revenue')
  assert.equal(after.commands.filter(command => command.id === 'sheet.command.configure-pivot-chart').length, 1)
  assert.deepEqual(after.model.charts.find(chart => chart.id === chartId).options.threeD, before.model.charts.find(chart => chart.id === chartId).options.threeD)
  assert.deepEqual(after.model.sheets.find(sheet => sheet.id === ids.source).cells, before.model.sheets.find(sheet => sheet.id === ids.source).cells)
  await close(); await expect(grid()).toBeFocused()
  await page.evaluate(() => window.pivotChartUIFixture.undo()); assert.equal(field(await state(), '1', chartId).aggregation, field(before, '1', chartId).aggregation)
  await page.evaluate(() => window.pivotChartUIFixture.redo()); assert.equal(field(await state(), '1', chartId).aggregation, 'average')
  await page.evaluate(() => window.pivotChartUIFixture.reopen()); const reopened = await state()
  assert.equal(field(reopened, '1', chartId).aggregation, 'average'); assert.equal(reopened.model.charts.find(chart => chart.id === chartId).title, 'Average revenue')
})

test('four field zones support adding, moving, exact order, removing and report options before a single apply', async () => {
  await open(); const before = await state()
  await choose('添加轴（类别）字段', 'Segment'); await dialog().getByRole('button', { name: '上移Segment', exact: true }).click()
  await choose('Date的字段区域', '图例（系列）'); await choose('添加值字段', 'Score')
  await dialog().getByRole('button', { name: '移除Score', exact: true }).click()
  await choose('值字段放置方向', '按行显示')
  await dialog().getByRole('checkbox', { name: '重复行标签', exact: true }).check()
  await dialog().getByRole('checkbox', { name: '显示行总计', exact: true }).check()
  assert.deepEqual((await state()).model, before.model)
  await expect(apply()).toBeEnabled(); await apply().click(); await waitHistory(before.history.undo + 1)
  const after = await state(), current = report(after)
  assert.deepEqual(current.fields.filter(field => field.area === 'row').sort((a, b) => a.position - b.position).map(field => field.sourceFieldId), ['3', '0'])
  assert.equal(field(after, '2').area, 'column'); assert.equal(current.fields.some(field => field.sourceFieldId === '4'), false)
  assert.equal(current.options.valuePosition, 'row'); assert.equal(current.options.repeatRowLabels, true); assert.equal(current.options.showRowGrandTotal, true)
})

test('label multi-select preserves a true empty filter and clearing it restores source members', async () => {
  await open(); const before = await state()
  await dialog().getByRole('button', { name: '筛选Region', exact: true }).click(); await expect(filter()).toBeVisible()
  await expect(filter().getByRole('combobox', { name: '透视图筛选标签', exact: true })).toBeEnabled()
  await filter().getByRole('checkbox', { name: '显示全部标签', exact: true }).uncheck()
  await filter().getByRole('combobox', { name: '透视图筛选标签', exact: true }).click()
  const dropdown = page.locator('.ant-select-dropdown:visible')
  await dropdown.getByText('North', { exact: true }).click(); await dropdown.getByText('South', { exact: true }).click()
  await page.keyboard.press('Escape'); await filter().getByRole('button', { name: '应用标签筛选', exact: true }).click()
  await expect(filter()).not.toBeVisible(); await expect(apply()).toBeEnabled(); await apply().click(); await waitHistory(before.history.undo + 1)
  assert.deepEqual(field(await state(), '0').selectedItems, [])
  await fieldsTab(); await dialog().getByRole('button', { name: '筛选Region', exact: true }).click()
  await filter().getByRole('checkbox', { name: '显示全部标签', exact: true }).check()
  await filter().getByRole('button', { name: '应用标签筛选', exact: true }).click(); await expect(apply()).toBeEnabled(); await apply().click()
  await waitHistory(before.history.undo + 2); assert.equal(Object.hasOwn(field(await state(), '0'), 'selectedItems'), false)
})

test('unchanged apply reports a no-op and leaves revision, dirty state and history unchanged', async () => {
  await open(); await expect(apply()).toBeEnabled(); await apply().click()
  await expect(dialog().locator('[role="status"]')).toContainText(/更新|没有变化/)
  await page.evaluate(() => window.pivotChartUIFixture.markSaved()); const before = await state()
  await expect(apply()).toBeEnabled(); await apply().click(); await expect(dialog().locator('[role="status"]')).toContainText('设置没有变化')
  const after = await state(); assert.deepEqual(after.model, before.model); assert.equal(after.revision, before.revision)
  assert.equal(after.history.undo, before.history.undo); assert.equal(after.history.dirty, false)
})

test('external edits retain the field draft until reread, and reread synchronizes owner/focus for the next real apply', async () => {
  await open(); const name = dialog().getByRole('textbox', { name: 'Region的显示名称', exact: true })
  await name.fill('Unsaved region'); await page.evaluate(() => window.pivotChartUIFixture.changeSource())
  await expect(name).toHaveValue('Unsaved region'); await expect(apply()).toBeDisabled()
  await expect(dialog().getByRole('alert').first()).toContainText('重新读取')
  await dialog().getByRole('button', { name: '重新读取', exact: true }).click(); await expect(name).toHaveValue('Region')
  const before = await state(); await name.fill('Accepted region'); await expect(apply()).toBeEnabled(); await apply().click(); await waitHistory(before.history.undo + 1)
  assert.equal(field(await state(), '0').displayName, 'Accepted region'); await close(); await expect(grid()).toBeFocused()
})

test('history replacement preserves unsaved text and refuses apply even when the content revision is unchanged', async () => {
  await open(); const before = await state(), name = dialog().getByRole('textbox', { name: 'Region的显示名称', exact: true })
  await name.fill('History draft'); await page.evaluate(() => window.pivotChartUIFixture.clearHistory())
  await expect(name).toHaveValue('History draft'); await expect(apply()).toBeDisabled(); assert.equal((await state()).revision, before.revision)
  await dialog().getByRole('button', { name: '重新读取', exact: true }).click(); await expect(name).toHaveValue('Region')
})

test('synthetic IME capture permits final text after external change and blocks closing, navigation and content publication while composing', async () => {
  await open(); const before = await state(), input = dialog().getByRole('textbox', { name: 'Region的显示名称', exact: true })
  await input.focus(); await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' })))
  await expect(dialog().getByRole('button', { name: '关闭', exact: true })).toBeDisabled(); await expect(apply()).toBeDisabled()
  await expect(dialog().getByRole('combobox', { name: '选择透视图', exact: true })).toBeDisabled()
  await page.evaluate(() => window.pivotChartUIFixture.changeSource()); await expect(input).toBeEnabled()
  await input.evaluate(node => {
    node.value = '区域草稿'; node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '区域草稿', isComposing: true }))
    node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '区域草稿' }))
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '区域草稿' }))
  })
  await expect(input).toHaveValue('区域草稿'); await expect(apply()).toBeDisabled()
  assert.equal(field(await state(), '0').displayName, field(before, '0').displayName)
  assert.equal((await state()).commands.filter(command => command.id === 'sheet.command.configure-pivot-chart').length, 0)
})

test('Escape closes nested labels to the Ant trigger and closes the main panel back to Canvas focus', async () => {
  await open(); const before = await state(), trigger = dialog().getByRole('button', { name: '筛选Region', exact: true })
  await trigger.click(); await expect(filter()).toBeVisible(); await page.keyboard.press('Escape')
  await expect(filter()).not.toBeVisible(); await expect(trigger).toBeFocused()
  await page.keyboard.press('Escape'); await expect(dialog()).not.toBeVisible(); await expect(grid()).toBeFocused()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('a stale delayed member read stops its own loading, retains selected draft labels and cannot publish returned source data', async () => {
  await open(); await page.evaluate(() => window.pivotChartUIFixture.delayMembers())
  await dialog().getByRole('button', { name: '筛选Region', exact: true }).click()
  await page.waitForFunction(() => window.pivotChartUIFixture.pendingMembers().length === 1)
  await page.evaluate(() => window.pivotChartUIFixture.changeSource()); const before = await state()
  await page.evaluate(() => window.pivotChartUIFixture.releaseMember())
  await expect(filter().getByText('正在读取真实来源标签…', { exact: true })).toHaveCount(0)
  await expect(filter().getByRole('alert')).toContainText('过期'); await expect(filter().getByRole('button', { name: '应用标签筛选', exact: true })).toBeDisabled()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('an old member completion cannot populate or clear loading in a newly opened linked-chart filter', async () => {
  await open(); await page.evaluate(() => window.pivotChartUIFixture.delayMembers())
  await dialog().getByRole('button', { name: '筛选Region', exact: true }).click()
  await page.waitForFunction(() => window.pivotChartUIFixture.pendingMembers().length === 1)
  await filter().getByRole('button', { name: '取消', exact: true }).click(); await close(); await open(ids.linked)
  await dialog().getByRole('button', { name: '筛选Region', exact: true }).click()
  await page.waitForFunction(() => window.pivotChartUIFixture.pendingMembers().length === 2)
  const before = await state(); await page.evaluate(() => window.pivotChartUIFixture.releaseMember(0, false))
  await expect(filter().getByText('正在读取真实来源标签…', { exact: true })).toBeVisible()
  await page.evaluate(() => window.pivotChartUIFixture.releaseMember())
  await expect(filter().getByRole('combobox', { name: '透视图筛选标签', exact: true })).toBeEnabled()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

for (const [name, steer] of [
  ['content', () => page.evaluate(() => window.pivotChartUIFixture.changeSource())],
  ['history', () => page.evaluate(() => window.pivotChartUIFixture.clearHistory())],
  ['user', () => page.evaluate(() => window.pivotChartUIFixture.changeUserWhilePending())],
]) test(`the real delayed configure refuses a changed ${name} and retains the newer model`, async () => {
  await open(); const title = await titleInput(); await title.fill('Pending old author')
  await page.evaluate(() => window.pivotChartUIFixture.delayMutation()); await expect(apply()).toBeEnabled(); await apply().click()
  await page.waitForFunction(() => window.pivotChartUIFixture.pendingMutations() === 1); await steer(); const before = await state()
  await page.evaluate(() => window.pivotChartUIFixture.releaseMutation())
  await expect(apply()).not.toHaveClass(/ant-btn-loading/); await expect(title).toHaveValue('Pending old author')
  const after = await state(); assert.deepEqual(after.model, before.model); assert.equal(after.history.undo, before.history.undo)
  assert.equal(after.model.charts.find(chart => chart.id === ids.owned).title, 'Owned chart')
})

test('an old configure completion cannot clear a newer pending panel after worksheet navigation', async () => {
  await open(); const oldTitle = await titleInput(); await oldTitle.fill('Old pending')
  await page.evaluate(() => window.pivotChartUIFixture.delayMutation()); await expect(apply()).toBeEnabled(); await apply().click()
  await page.waitForFunction(() => window.pivotChartUIFixture.pendingMutations() === 1)
  await page.evaluate(() => window.pivotChartUIFixture.switchSheet()); await open(ids.foreign)
  const newTitle = await titleInput(); await newTitle.fill('New pending'); await expect(apply()).toBeEnabled(); await apply().click()
  await page.waitForFunction(() => window.pivotChartUIFixture.pendingMutations() === 2); const before = await state()
  await page.evaluate(() => window.pivotChartUIFixture.releaseMutation('configure', false))
  await expect(apply()).toHaveClass(/ant-btn-loading/); await expect(newTitle).toHaveValue('New pending')
  await page.evaluate(() => window.pivotChartUIFixture.releaseMutation()); await waitHistory(before.history.undo + 1)
  const after = await state(); assert.equal(after.model.charts.find(chart => chart.id === ids.foreign).title, 'New pending')
  assert.equal(after.model.charts.find(chart => chart.id === ids.owned).title, 'Owned chart')
})

test('Studio replacement while members are delayed releases the old panel without mutating the replacement workbook', async () => {
  await open(); await page.evaluate(() => window.pivotChartUIFixture.delayMembers())
  await dialog().getByRole('button', { name: '筛选Region', exact: true }).click()
  await page.waitForFunction(() => window.pivotChartUIFixture.pendingMembers().length === 1)
  await page.evaluate(() => window.pivotChartUIFixture.mount({ id: 'replacement-pivot-chart-ui' }))
  await expect(dialog()).not.toBeVisible(); await expect(filter()).not.toBeVisible()
  const current = await state(); assert.equal(current.model.id, 'replacement-pivot-chart-ui'); assert.equal(current.history.undo, 0)
  assert.equal(current.commands.filter(command => command.id === 'sheet.command.configure-pivot-chart').length, 0)
})

test('theme and Ribbon variant switches preserve field text while Ant modal surfaces and controls follow ConfigProvider tokens', async () => {
  await open(); const before = await state(), input = dialog().getByRole('textbox', { name: 'Region的显示名称', exact: true })
  await input.fill('Theme draft')
  const light = await dialog().locator('[data-testid="pivot-chart-panel"]').evaluate(node => ({ color: getComputedStyle(node).color,
    surface: getComputedStyle(node.closest('.ant-modal-content')).backgroundColor }))
  await page.evaluate(() => window.pivotChartUIFixture.appearance(true, 'grid'))
  const dark = await dialog().locator('[data-testid="pivot-chart-panel"]').evaluate(node => ({ color: getComputedStyle(node).color,
    surface: getComputedStyle(node.closest('.ant-modal-content')).backgroundColor }))
  assert.notEqual(dark.color, light.color); assert.notEqual(dark.surface, light.surface); await expect(input).toHaveValue('Theme draft')
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

for (const [name, options] of [['host', { readonly: true }], ['chart', { protectedChart: true }], ['linked output', { protectedOutput: true }]])
  test(`${name} protection disables panel apply and the real public configure refuses a mutation`, async () => {
    await page.evaluate(options => window.pivotChartUIFixture.mount({ ...options, currentUserId: 'peer' }), options)
    const chartId = name === 'linked output' ? ids.linked : ids.owned; await open(chartId); await expect(apply()).toBeDisabled()
    const before = await state(), rejected = await page.evaluate(async chartId => {
      try { return await window.pivotChartUIFixture.configure(chartId, { patch: { title: 'Forbidden' }, fieldsConfig: { fields: [] } }) }
      catch (error) { return { rejected: true, message: error.message } }
    }, chartId)
    assert.ok(rejected.rejected || rejected.result === false); assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
  })

test('a hidden source range refuses member loading and leaves the field-source add controls unavailable', async () => {
  await page.evaluate(() => window.pivotChartUIFixture.mount({ hiddenSource: true, currentUserId: 'peer' })); await open(ids.owned, false)
  await expect(dialog().getByRole('alert')).toContainText('无法访问'); await expect(apply()).toBeDisabled()
  await expect(dialog().getByRole('combobox', { name: '添加值字段', exact: true })).toBeDisabled()
  await expect(dialog().locator('[data-pivot-field]')).toHaveCount(0)
  assert.equal((await state()).history.undo, 0)
})

for (const chartId of [ids.owned, ids.linked]) test(`${chartId} refresh and removal retain source cells and the independent PivotTable, with removal undo`, async () => {
  await open(chartId); const before = await state()
  await dialog().getByRole('button', { name: '刷新透视图', exact: true }).click()
  await expect(dialog().locator('[role="status"]')).toContainText(/重新读取|已刷新/)
  const refreshed = await state(); await dialog().getByRole('button', { name: '移除透视图', exact: true }).click(); await waitHistory(refreshed.history.undo + 1)
  const after = await state(); assert.equal(after.model.charts.some(chart => chart.id === chartId), false)
  assert.ok(after.model.pivotTables.find(pivot => pivot.id === ids.primary))
  assert.deepEqual(after.model.sheets.find(sheet => sheet.id === ids.source).cells, before.model.sheets.find(sheet => sheet.id === ids.source).cells)
  await close(); await page.evaluate(() => window.pivotChartUIFixture.undo()); assert.ok((await state()).model.charts.find(chart => chart.id === chartId))
})
