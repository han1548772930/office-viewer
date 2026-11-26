import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'
import { showAsUICases, showAsUIIds as ids, showAsUIRule as rule } from './fixtures/pivot-show-data-as-ui-fixture.mjs'

// Deferred real Ant panels, Workbench/Canvas2D, report engine and Facade/history; never executed in this batch.
// Composition events verify event ordering only. Native Windows IME, WebGPU and Office files remain unverified.
let server, browser, context, page, url
const errors = [], surfaces = [['table', 'linked'], ['chart', 'owned'], ['chart', 'linked']]
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/pivot-show-data-as-ui-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1200 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.showAsUIFixture?.state().model.charts.length >= 1)
}, 60_000)
afterEach(async () => {
  try { await page?.evaluate(() => window.showAsUIFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.showAsUIFixture.state())
const mount = options => page.evaluate(options => window.showAsUIFixture.mount(options), options)
const main = kind => page.getByRole('dialog', { name: kind === 'table' ? '数据透视表' : '透视图', exact: true })
const settings = (all = false) => page.getByRole('dialog', { name: all ? '字段设置' : '值显示方式', exact: true })
const shared = () => page.getByTestId('shared-show-as-picker')
const dropdown = () => page.locator('.ant-select-dropdown:visible')
const report = value => value.model.charts[0].ownedPivot ?? value.model.pivotTables.find(pivot => pivot.id === value.model.charts[0].pivotTableId)
const field = (value, id = ids.amount) => report(value).fields.find(field => field.id === id)
const sourceCells = value => value.model.sheets.find(sheet => sheet.id === ids.data).cells
const closeNumber = (actual, expected) => assert.ok(typeof actual === 'number' && Math.abs(actual - expected) < 1e-10, `${actual} ~= ${expected}`)
const baseLabels = { [ids.region]: 'Region', [ids.product]: 'Product', [ids.year]: 'Year', [ids.quarter]: 'Quarter' }
const inputName = kind => kind === 'table' ? '值字段显示名称' : 'Amount的显示名称'
const apply = (kind, all = false) => kind === 'table' ? settings(all).getByRole('button', { name: '确定', exact: true })
  : main(kind).getByRole('button', { name: '应用透视图', exact: true })
const mutationKind = kind => kind === 'table' ? 'tableSet' : 'chartConfigure'
async function choose(name, label, target) {
  await target.getByRole('combobox', { name, exact: true }).click()
  await dropdown().getByText(label, { exact: true }).click()
}
async function openMain(kind) {
  await page.getByRole('button', { name: kind === 'table' ? '打开普通透视表' : '打开透视图', exact: true }).click()
  await expect(main(kind)).toBeVisible()
  if (kind === 'table') {
    await main(kind).getByRole('combobox', { name: '选择数据透视表', exact: true }).click()
    await dropdown().getByText(/^数据透视表 1 ·/).click()
  } else await expect(main(kind).getByRole('combobox', { name: 'Amount的值显示方式', exact: true })).toBeVisible()
}
async function openTableSettings(all = false) {
  await main('table').locator(`.pivot-area-grid [data-field-id="${ids.amount}"]`).getByRole('button', { name: '字段 Amount 操作', exact: true }).click()
  await page.getByRole('menuitem', { name: all ? '字段设置' : '值显示方式', exact: true }).click()
  await expect(settings(all)).toBeVisible()
}
async function openSurface(kind, all = false) { await openMain(kind); if (kind === 'table') await openTableSettings(all) }
async function selectRule(kind, value, label, all = false) {
  const target = kind === 'table' ? settings(all) : main(kind), prefix = kind === 'table' ? '' : 'Amount的'
  await choose(`${prefix}值显示方式`, label, target)
  if (value.baseFieldId) await choose(`${prefix}基准字段`, baseLabels[value.baseFieldId], target)
  if (value.baseItemType) await choose(`${prefix}基准项类型`, value.baseItemType === 'previous' ? '上一项' : '下一项', target)
  else if (value.baseItem) {
    await expect(target.getByRole('combobox', { name: `${prefix}基准项`, exact: true })).toBeEnabled()
    await choose(`${prefix}基准项`, value.baseItem, target)
  }
}
async function commit(kind, count, all = false) {
  await expect(apply(kind, all)).toBeEnabled(); await apply(kind, all).click()
  await page.waitForFunction(count => window.showAsUIFixture.state().history.undo === count, count)
  if (kind === 'table') await expect(settings(all)).not.toBeVisible()
  else await expect(apply(kind)).not.toHaveClass(/ant-btn-loading/)
}
async function closeMain(kind) { await main(kind).getByRole('button', { name: '关闭', exact: true }).click(); await expect(main(kind)).not.toBeVisible() }

test.each(surfaces.flatMap(([kind, ownership]) => showAsUICases.map(([value, label, expected]) => [kind, ownership, value, label, expected])))
  ('%s/%s picker applies %j to actual hierarchy values with one captured command', async (kind, ownership, value, label, expected) => {
    await mount({ ownership, initialRule: rule('runningTotal', ids.region) }); await openSurface(kind)
    const before = await state(); await selectRule(kind, value, label); assert.deepEqual((await state()).model, before.model)
    await commit(kind, before.history.undo + 1); const after = await state()
    closeNumber(after.value, expected); assert.deepEqual(field(after).showDataAs, value.type === 'normal' ? undefined : value)
    assert.deepEqual(sourceCells(after), sourceCells(before)); assert.equal(after.model.pivotTables.length, before.model.pivotTables.length)
    const id = kind === 'table' ? 'sheet.command.set-pivot-show-data-as' : 'sheet.command.configure-pivot-chart'
    const commands = after.commands.filter(command => command.id === id); assert.equal(commands.length, 1)
    assert.equal(commands[0].params.expectedRevision, before.revision); assert.equal(commands[0].params.expectedHistoryGeneration, before.generation)
    assert.equal(commands[0].params.expectedUserId, 'local-user'); assert.equal(commands[0].params.host, 'workbench')
    if (kind === 'chart') assert.ok(commands[0].params.fieldsConfig.fields.find(field => field.id === ids.amount))
  })

test('default Ant picker exposes exactly the fifteen fixed entries in order', async () => {
  await page.evaluate(() => window.showAsUIFixture.sharedPicker({ type: 'normal', baseFieldId: '', baseItem: '' }))
  await shared().getByRole('combobox', { name: '值显示方式', exact: true }).click()
  assert.deepEqual(await dropdown().locator('.ant-select-item-option-content').allTextContents(), showAsUICases.map(([, label]) => label))
  await expect(dropdown().getByText('父总计的百分比', { exact: true })).toHaveCount(0)
})

test('specified-parent UI permits only actual parent fields on both axes', async () => {
  await page.evaluate(() => window.showAsUIFixture.sharedPicker({ type: 'normal', baseFieldId: '', baseItem: '' }))
  await choose('值显示方式', '指定父项汇总的百分比', shared()); await shared().getByRole('combobox', { name: '基准字段', exact: true }).click()
  assert.deepEqual(await dropdown().locator('.ant-select-item-option-content').allTextContents(), ['Region', 'Year'])
  await dropdown().getByText('Year', { exact: true }).click(); assert.equal((await state()).shared.value.baseFieldId, ids.year)
  assert.equal((await state()).shared.valid, true); await expect(shared().getByRole('combobox', { name: '基准项', exact: true })).toHaveCount(0)
})

test.each(surfaces)('%s/%s selects grouped real members and leaves a repeated value instance independent', async (kind, ownership) => {
  await mount({ ownership, grouping: true, repeated: true }); await openSurface(kind); const before = await state()
  const target = kind === 'table' ? settings() : main(kind), prefix = kind === 'table' ? '' : 'Amount的'
  await choose(`${prefix}值显示方式`, '指定项的百分比', target); await choose(`${prefix}基准字段`, 'Product', target)
  await expect(target.getByRole('combobox', { name: `${prefix}基准项`, exact: true })).toBeEnabled()
  await target.getByRole('combobox', { name: `${prefix}基准项`, exact: true }).click()
  assert.deepEqual(await dropdown().locator('.ant-select-item-option-content').allTextContents(), ['Mix', 'C', 'D'])
  await dropdown().getByText('Mix', { exact: true }).click(); await commit(kind, before.history.undo + 1)
  const after = await state(); closeNumber(after.value, 1); assert.equal(field(after).showDataAs.baseFieldId, ids.product)
  assert.equal(field(after, 'amount-second').showDataAs, undefined); assert.deepEqual(sourceCells(after), sourceCells(before))
})

test('duplicate labels and canonical/local identity collisions stay distinct in the real base picker', async () => {
  const model = (await state()).model, pivot = model.pivotTables[0]
  pivot.fields.find(field => field.id === ids.region).id = 'product-data'; pivot.fields.find(field => field.id === ids.year).displayName = 'Region'
  await mount({ model }); await page.evaluate(() => window.showAsUIFixture.sharedPicker({ type: 'runningTotal', baseFieldId: '', baseItem: '' }))
  await shared().getByRole('combobox', { name: '基准字段', exact: true }).click()
  await expect(dropdown().getByText('Region · 行 · product-data', { exact: true })).toBeVisible()
  await expect(dropdown().getByText(`Region · 列 · ${ids.year}`, { exact: true })).toBeVisible()
  await dropdown().getByText('Region · 行 · product-data', { exact: true }).click()
  assert.equal((await state()).shared.value.baseFieldId, 'product-data'); assert.equal((await state()).shared.valid, true)
})

test.each(surfaces.flatMap(([kind, ownership]) => ['differenceFrom', 'percentDifferenceFrom'].flatMap(type => ['previous', 'next'].map(relative => [kind, ownership, type, relative]))))
  ('%s/%s %s supports a real %s relative member without a stale explicit item', async (kind, ownership, type, relative) => {
    await mount({ ownership, initialRule: rule('percentOf', ids.product, 'B') }); await openSurface(kind); const before = await state()
    const value = rule(type, ids.product, '', relative), label = showAsUICases.find(([value]) => value.type === type)[1]
    await selectRule(kind, value, label); const target = kind === 'table' ? settings() : main(kind)
    await expect(target.getByRole('combobox', { name: kind === 'table' ? '基准项' : 'Amount的基准项', exact: true })).toHaveCount(0)
    await commit(kind, before.history.undo + 1); const after = await state(); assert.deepEqual(field(after).showDataAs, value)
    if (relative === 'previous') assert.equal(after.value, '#N/A')
    else closeNumber(after.value, type === 'differenceFrom' ? -20 : -2 / 3)
  })

test.each(surfaces)('%s/%s reset removes authority and accepting normal again remains a no-op', async (kind, ownership) => {
  await mount({ ownership, initialRule: rule('percentOfGrandTotal') }); await openSurface(kind); const before = await state()
  const target = kind === 'table' ? settings() : main(kind)
  await target.getByRole('button', { name: kind === 'table' ? '恢复无计算' : 'Amount的恢复无计算', exact: true }).click()
  await commit(kind, before.history.undo + 1); closeNumber((await state()).value, 10); assert.equal(field(await state()).showDataAs, undefined)
  await closeMain(kind); await openSurface(kind); const normal = await state(); await expect(apply(kind)).toBeEnabled(); await apply(kind).click()
  if (kind === 'table') await expect(settings()).not.toBeVisible()
  else await expect(main(kind).locator('[role="status"]')).toContainText('设置没有变化')
  const unchanged = await state(); assert.deepEqual(unchanged.model, normal.model); assert.deepEqual(unchanged.history, normal.history)
  assert.equal(unchanged.revision, normal.revision); assert.equal(unchanged.commands.length, normal.commands.length)
})

test('ordinary field settings combine rename, aggregation, number format and Show As in one transaction', async () => {
  await openSurface('table', true); const before = await state()
  await settings(true).getByRole('textbox', { name: '值字段显示名称', exact: true }).fill('Average amount')
  await settings(true).getByRole('radio', { name: '平均值', exact: true }).check()
  await settings(true).getByRole('button', { name: '数字格式', exact: true }).click()
  const format = page.getByRole('dialog', { name: '数字格式', exact: true })
  await format.getByRole('textbox', { name: '数字格式代码', exact: true }).fill('0.00'); await format.getByRole('button', { name: '确定', exact: true }).click()
  await selectRule('table', rule('percentOfGrandTotal'), '总计的百分比', true); await commit('table', before.history.undo + 1, true)
  const after = await state(); assert.equal(field(after).displayName, 'Average amount'); assert.equal(field(after).aggregation, 'average')
  assert.equal(field(after).format, '0.00'); closeNumber(after.value, .4)
  assert.equal(after.commands.filter(command => command.id === 'sheet.command.set-pivot-field-setting').length, 1)
  assert.equal(after.commands.filter(command => command.id === 'sheet.command.set-pivot-show-data-as').length, 0)
})

test('an uncommitted ordinary destination saves the value rule locally before one main configure', async () => {
  await openMain('table'); const before = await state()
  await main('table').getByRole('textbox', { name: '透视表起始单元格', exact: true }).fill('B40'); await openTableSettings()
  await selectRule('table', rule('index'), '指数'); await expect(apply('table')).toBeEnabled(); await apply('table').click()
  await expect(settings()).not.toBeVisible(); assert.deepEqual((await state()).model, before.model)
  await expect(main('table').locator('[role="status"]')).toContainText('保存到当前配置')
  await main('table').getByRole('button', { name: '应用字段配置', exact: true }).click()
  await page.waitForFunction(count => window.showAsUIFixture.state().history.undo === count, before.history.undo + 1)
  const after = await state(); assert.equal(report(after).anchor.row, 39); closeNumber(after.value, 20 / 21)
  const commands = after.commands.filter(command => command.id === 'sheet.command.configure-pivot-table')
  assert.equal(commands.length, 1); assert.equal(commands[0].params.hostSheetId, ids.host)
  assert.equal(commands[0].params.expectedRevision, before.revision); assert.deepEqual(sourceCells(after), sourceCells(before))
})

test.each(surfaces)('%s/%s moving a value away and back never revives its prior rule', async (kind, ownership) => {
  await mount({ ownership, initialRule: rule('percentOfGrandTotal') }); await openMain(kind); const before = await state()
  if (kind === 'table') {
    const button = () => main(kind).locator(`.pivot-area-grid [data-field-id="${ids.amount}"]`).getByRole('button', { name: '字段 Amount 操作', exact: true })
    await button().click(); await page.getByRole('menuitem', { name: '移动到筛选', exact: true }).click()
    await button().click(); await page.getByRole('menuitem', { name: '移动到值', exact: true }).click()
    await main(kind).getByRole('button', { name: '应用字段配置', exact: true }).click()
    await page.waitForFunction(count => window.showAsUIFixture.state().history.undo === count, before.history.undo + 1)
  } else {
    await choose('Amount的字段区域', '筛选', main(kind)); await choose('Amount的字段区域', '值', main(kind)); await commit(kind, before.history.undo + 1)
  }
  assert.equal(field(await state()).showDataAs, undefined); closeNumber((await state()).value, 10)
})

test.each(surfaces)('%s/%s real history and JSON save/reopen preserve the applied rule and source cells', async (kind, ownership) => {
  await mount({ ownership }); await openSurface(kind); const before = await state()
  await selectRule(kind, rule('percentOfGrandTotal'), '总计的百分比'); await commit(kind, before.history.undo + 1); await closeMain(kind)
  await expect(page.getByRole('button', { name: kind === 'table' ? '打开普通透视表' : '打开透视图', exact: true })).toBeFocused()
  await page.evaluate(() => window.showAsUIFixture.undo()); closeNumber((await state()).value, 10)
  await page.evaluate(() => window.showAsUIFixture.redo()); closeNumber((await state()).value, .05)
  await page.evaluate(() => window.showAsUIFixture.reopen()); const reopened = await state()
  assert.deepEqual(field(reopened).showDataAs, rule('percentOfGrandTotal')); closeNumber(reopened.value, .05)
  assert.deepEqual(sourceCells(reopened), sourceCells(before))
})

test('late shared members cannot populate or end loading for a newer selected field', async () => {
  await page.evaluate(() => window.showAsUIFixture.sharedPicker({ type: 'percentOf', baseFieldId: 'product-label', baseItem: 'B' }, true))
  await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 1)
  await choose('基准字段', 'Quarter', shared()); await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 2)
  await page.evaluate(() => window.showAsUIFixture.releaseGate('shared-members', 0, false))
  await expect(shared().getByRole('combobox', { name: '基准项', exact: true })).toBeDisabled(); assert.equal((await state()).shared.valid, false)
  await page.evaluate(() => window.showAsUIFixture.releaseGate('shared-members'))
  await expect(shared().getByRole('combobox', { name: '基准项', exact: true })).toBeEnabled(); await choose('基准项', 'Q1', shared())
  assert.deepEqual((await state()).shared.value, rule('percentOf', ids.quarter, 'Q1')); assert.equal((await state()).shared.valid, true)
})

test('a shared context generation re-reads actual members and ignores the old finally', async () => {
  await page.evaluate(() => window.showAsUIFixture.sharedPicker({ type: 'percentOf', baseFieldId: 'product-label', baseItem: 'B' }, true))
  await page.evaluate(() => window.showAsUIFixture.sharedContext()); await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 2)
  await page.evaluate(() => window.showAsUIFixture.releaseGate('shared-members', 0, false))
  await expect(shared().getByRole('combobox', { name: '基准项', exact: true })).toBeDisabled()
  await page.evaluate(() => window.showAsUIFixture.releaseGate('shared-members')); await expect(shared().getByRole('combobox', { name: '基准项', exact: true })).toBeEnabled()
  assert.deepEqual((await state()).shared.reads, [ids.product, ids.product]); assert.equal((await state()).shared.valid, true)
})

test('losing readable access clears visible members and a pending read cannot expose them again', async () => {
  await page.evaluate(() => window.showAsUIFixture.sharedPicker({ type: 'percentOf', baseFieldId: 'product-label', baseItem: 'B' }, true))
  const before = await state(); await page.evaluate(() => window.showAsUIFixture.sharedAccess(false))
  await page.evaluate(() => window.showAsUIFixture.releaseGate('shared-members'))
  await expect(shared().getByRole('combobox', { name: '值显示方式', exact: true })).toBeDisabled()
  await expect(shared().getByRole('combobox', { name: '基准项', exact: true })).toBeDisabled()
  await expect(shared().locator('.ant-select-selection-item').filter({ hasText: /^B$/ })).toHaveCount(0)
  assert.equal((await state()).shared.valid, false); assert.deepEqual((await state()).model, before.model)
})

test('the blank real member is displayed as unavailable while a nonblank member stays selectable', async () => {
  await mount({ blank: true }); await page.evaluate(() => window.showAsUIFixture.sharedPicker({ type: 'percentOf', baseFieldId: 'product-label', baseItem: 'B' }))
  await expect(shared().getByRole('alert').first()).toContainText('空白成员')
  await shared().getByRole('combobox', { name: '基准项', exact: true }).click()
  await expect(dropdown().locator('.ant-select-item-option').filter({ hasText: '（空白）' })).toHaveClass(/disabled/)
  await dropdown().getByText('B', { exact: true }).click(); assert.equal((await state()).shared.valid, true)
})

test.each(surfaces)('%s/%s imported compatibility stays preserved until an explicit reset', async (kind, ownership) => {
  await mount({ ownership, compatibility: true }); await openSurface(kind); const before = await state(), target = kind === 'table' ? settings() : main(kind)
  await expect(target.getByRole('alert').first()).toContainText('保留导入')
  if (kind === 'table') await settings().getByRole('button', { name: '取消', exact: true }).click()
  await closeMain(kind); assert.deepEqual(field(await state()).showDataAs, field(before).showDataAs)
  await openSurface(kind); await (kind === 'table' ? settings() : main(kind)).getByRole('button', { name: kind === 'table' ? '恢复无计算' : 'Amount的恢复无计算', exact: true }).click()
  await commit(kind, before.history.undo + 1); assert.equal(field(await state()).showDataAs, undefined)
})

test.each(surfaces.flatMap(([kind, ownership]) => ['source', 'history', 'user', 'host'].map(change => [kind, ownership, change])))
  ('%s/%s pending real commit rejects a changed %s and preserves the newer workbook', async (kind, ownership, change) => {
    await mount({ ownership }); await openSurface(kind); await selectRule(kind, rule('percentOfGrandTotal'), '总计的百分比')
    await page.evaluate(kind => window.showAsUIFixture.delayMutation(kind), mutationKind(kind))
    await expect(apply(kind)).toBeEnabled(); await apply(kind).click(); await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 1)
    await page.evaluate(async change => {
      if (change === 'source') await window.showAsUIFixture.changeSource()
      else if (change === 'history') await window.showAsUIFixture.clearHistory()
      else if (change === 'user') await window.showAsUIFixture.changeUser()
      else await window.showAsUIFixture.switchSheet()
    }, change)
    const newer = await state(); await page.evaluate(kind => window.showAsUIFixture.releaseGate(kind), mutationKind(kind))
    const after = await state(); assert.deepEqual(after.model, newer.model); assert.deepEqual(after.history, newer.history)
    assert.equal(after.revision, newer.revision); assert.equal(field(after).showDataAs, undefined)
  })

test('ordinary Facade member completion becomes stale after a real source edit', async () => {
  await openSurface('table'); await page.evaluate(() => window.showAsUIFixture.delayMembers())
  await choose('值显示方式', '指定项的百分比', settings()); await choose('基准字段', 'Product', settings())
  await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 1)
  await page.evaluate(() => window.showAsUIFixture.changeSource()); const newer = await state()
  await page.evaluate(() => window.showAsUIFixture.releaseGate('table-members'))
  await expect(apply('table')).toBeDisabled(); await expect(settings().getByRole('combobox', { name: '基准项', exact: true })).toBeDisabled()
  await expect(settings().getByRole('alert').filter({ hasText: '此设置草稿已过期' })).toBeVisible()
  assert.deepEqual((await state()).model, newer.model); assert.deepEqual((await state()).history, newer.history)
})

test('an ordinary old member completion cannot clear loading in a reopened settings session', async () => {
  await openSurface('table'); await page.evaluate(() => window.showAsUIFixture.delayMembers())
  await choose('值显示方式', '指定项的百分比', settings()); await choose('基准字段', 'Product', settings())
  await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 1); await settings().getByRole('button', { name: '取消', exact: true }).click()
  await openTableSettings(); await choose('值显示方式', '指定项的百分比', settings()); await choose('基准字段', 'Quarter', settings())
  await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 2)
  await page.evaluate(() => window.showAsUIFixture.releaseGate('table-members', 0, false))
  await expect(settings().getByRole('combobox', { name: '基准项', exact: true })).toBeDisabled()
  await page.evaluate(() => window.showAsUIFixture.releaseGate('table-members'))
  await expect(settings().getByRole('combobox', { name: '基准项', exact: true })).toBeEnabled(); await choose('基准项', 'Q1', settings())
  await expect(apply('table')).toBeEnabled(); assert.equal(field(await state()).showDataAs, undefined)
})

test.each(surfaces)('%s/%s an old commit finally cannot release a newer pending panel', async (kind, ownership) => {
  await mount({ ownership }); await openSurface(kind); await selectRule(kind, rule('percentOfGrandTotal'), '总计的百分比')
  await page.evaluate(kind => window.showAsUIFixture.delayMutation(kind), mutationKind(kind)); await expect(apply(kind)).toBeEnabled(); await apply(kind).click()
  await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 1)
  await page.evaluate(() => window.showAsUIFixture.changeSource()); await page.evaluate(() => window.showAsUIFixture.closeExternally()); await openSurface(kind)
  await selectRule(kind, rule('index'), '指数'); await expect(apply(kind)).toBeEnabled(); await apply(kind).click()
  await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 2); const newer = await state()
  await page.evaluate(kind => window.showAsUIFixture.releaseGate(kind, 0, false), mutationKind(kind))
  await expect(apply(kind)).toHaveClass(/ant-btn-loading/); assert.deepEqual((await state()).model, newer.model)
  await page.evaluate(kind => window.showAsUIFixture.releaseGate(kind), mutationKind(kind))
  await page.waitForFunction(count => window.showAsUIFixture.state().history.undo === count, newer.history.undo + 1)
  assert.deepEqual(field(await state()).showDataAs, rule('index'))
})

test('Studio replacement with a member gate cannot write into the replacement workbook', async () => {
  await openSurface('table'); await page.evaluate(() => window.showAsUIFixture.delayMembers())
  await choose('值显示方式', '指定项的百分比', settings()); await choose('基准字段', 'Product', settings())
  await page.waitForFunction(() => window.showAsUIFixture.state().pending.length === 1)
  await mount({ ownership: 'owned' }); await expect(settings()).not.toBeVisible()
  const after = await state(); assert.equal(after.history.undo, 0); assert.equal(field(after).showDataAs, undefined); assert.equal(after.commands.length, 0)
})

test.each(['table', 'chart'])('%s synthetic composition preserves the final name and permits one later submission', async kind => {
  await openSurface(kind, kind === 'table'); const before = await state(), all = kind === 'table'
  await selectRule(kind, rule('percentOfGrandTotal'), '总计的百分比', all)
  const target = all ? settings(true) : main(kind), input = target.getByRole('textbox', { name: inputName(kind), exact: true })
  await input.focus(); await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await expect(apply(kind, all)).toBeDisabled(); await expect(target.getByRole('button', { name: all ? '取消' : '关闭', exact: true })).toBeDisabled()
  await input.evaluate(node => {
    node.value = '收入草稿'; node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '收入草稿', inputType: 'insertCompositionText', isComposing: true }))
    node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '收入草稿' }))
    node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '收入草稿', inputType: 'insertText' }))
  })
  await expect(input).toHaveValue('收入草稿'); await commit(kind, before.history.undo + 1, all)
  assert.equal(field(await state()).displayName, '收入草稿'); closeNumber((await state()).value, .05)
})

test.each(['table', 'chart'])('%s synthetic composition keeps late text after an external edit without publishing it', async kind => {
  await openSurface(kind, kind === 'table'); const all = kind === 'table', target = all ? settings(true) : main(kind)
  const input = target.getByRole('textbox', { name: inputName(kind), exact: true }); await input.focus()
  await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await page.evaluate(() => window.showAsUIFixture.changeSource()); const newer = await state(); await expect(input).toBeEnabled()
  await input.evaluate(node => {
    node.value = '保留文字'; node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '保留文字', inputType: 'insertCompositionText', isComposing: true }))
    node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '保留文字' }))
    node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '保留文字', inputType: 'insertText' }))
  })
  await expect(input).toHaveValue('保留文字'); await expect(apply(kind, all)).toBeDisabled(); assert.deepEqual((await state()).model, newer.model)
})

test.each(surfaces.flatMap(([kind, ownership]) => ['classic', 'grid'].map(variant => [kind, ownership, variant])))
  ('%s/%s Ant controls retain text while dark theme and %s layout change', async (kind, ownership, variant) => {
    await mount({ ownership, variant }); await openSurface(kind, kind === 'table'); const before = await state(), all = kind === 'table'
    const target = all ? settings(true) : main(kind), input = target.getByRole('textbox', { name: inputName(kind), exact: true })
    await input.fill('Theme draft'); const style = node => ({ color: getComputedStyle(node).color, surface: getComputedStyle(node.closest('.ant-modal-content')).backgroundColor })
    const light = await target.getByTestId(all ? 'pivot-field-settings' : 'pivot-chart-panel').evaluate(style)
    await page.evaluate(variant => window.showAsUIFixture.appearance(true, variant), variant)
    const dark = await target.getByTestId(all ? 'pivot-field-settings' : 'pivot-chart-panel').evaluate(style)
    assert.notEqual(dark.color, light.color); assert.notEqual(dark.surface, light.surface); await expect(input).toHaveValue('Theme draft')
    assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  })

test('Escape restores the ordinary nested Ant trigger and then the Ant fixture launcher', async () => {
  await openSurface('table'); await page.keyboard.press('Escape'); await expect(settings()).not.toBeVisible()
  await expect(main('table').locator(`.pivot-area-grid [data-field-id="${ids.amount}"]`).getByRole('button', { name: '字段 Amount 操作', exact: true })).toBeFocused()
  await closeMain('table'); await expect(page.getByRole('button', { name: '打开普通透视表', exact: true })).toBeFocused()
})

test('a protected source refuses the ordinary public member read and its value dialog remains closed', async () => {
  await mount({ hiddenSource: true, currentUserId: 'peer' }); await openMain('table')
  await main('table').locator(`.pivot-area-grid [data-field-id="${ids.amount}"]`).getByRole('button', { name: '字段 Amount 操作', exact: true }).click()
  await page.getByRole('menuitem', { name: '值显示方式', exact: true }).click(); await expect(settings()).not.toBeVisible()
  const rejected = await page.evaluate(async () => { try { await window.showAsUIFixture.readMembers(); return false } catch { return true } })
  assert.equal(rejected, true); assert.equal((await state()).history.undo, 0)
})

test('a protected source keeps chart value and member controls unavailable', async () => {
  await mount({ ownership: 'owned', hiddenSource: true, currentUserId: 'peer' })
  await page.getByRole('button', { name: '打开透视图', exact: true }).click(); await expect(main('chart')).toBeVisible()
  await expect(apply('chart')).toBeDisabled(); await expect(main('chart').getByRole('combobox', { name: 'Amount的值显示方式', exact: true })).toHaveCount(0)
  assert.equal((await state()).history.undo, 0)
})

test.each(['table', 'chart'])('%s a protected linked output cannot publish the value draft', async kind => {
  await mount({ protectedOutput: true, currentUserId: 'peer' }); await openMain(kind); const before = await state()
  if (kind === 'table') {
    await main(kind).locator(`.pivot-area-grid [data-field-id="${ids.amount}"]`).getByRole('button', { name: '字段 Amount 操作', exact: true }).click()
    await page.getByRole('menuitem', { name: '值显示方式', exact: true }).click(); await expect(settings()).not.toBeVisible()
  } else await expect(apply(kind)).toBeDisabled()
  const rejected = await page.evaluate(async () => { try { await window.showAsUIFixture.setRule({ type: 'percentOfGrandTotal', baseFieldId: '', baseItem: '' }); return false } catch { return true } })
  assert.equal(rejected, true); assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
})
