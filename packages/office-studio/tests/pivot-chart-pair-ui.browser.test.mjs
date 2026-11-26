import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'
import { pivotChartUIIds as ids } from './fixtures/pivot-chart-ui-fixture.mjs'

// Independent paired-creation and presentation drafts reuse the real existing Ant/Canvas/Facade fixture.
// No server, browser, tests or checks have run. Synthetic composition is not native Windows IME evidence.
let server, browser, context, page, url
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/pivot-chart-ui-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1200 } }); page = await context.newPage()
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
const create = paired => dialog().getByRole('button', { name: paired ? '同时创建透视表和透视图' : '创建透视图', exact: true })
const apply = () => dialog().getByRole('button', { name: '应用透视图', exact: true })
const dropdown = () => page.locator('.ant-select-dropdown:visible')
const report = (value, chartId) => value.model.charts.find(chart => chart.id === chartId)?.ownedPivot
  ?? value.model.pivotTables.find(pivot => pivot.id === value.model.charts.find(chart => chart.id === chartId)?.pivotTableId)
const sourceCells = value => value.model.sheets.find(sheet => sheet.id === ids.source).cells
async function choose(name, label) {
  await dialog().getByRole('combobox', { name, exact: true }).click(); await dropdown().getByText(label, { exact: true }).click()
}
async function appearance() { await dialog().getByRole('tab', { name: '图表设置', exact: true }).click() }
async function fieldsTab() { await dialog().getByRole('tab', { name: '字段', exact: true }).click() }
async function fresh(options = {}) {
  await page.evaluate(options => window.pivotChartUIFixture.mount({ existing: false, ...options }), options)
  await page.evaluate(() => window.pivotChartUIFixture.selectSourceRange())
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  const trigger = page.getByRole('toolbar', { name: '数据', exact: true }).getByRole('button', { name: '透视图', exact: true })
  await expect(trigger).toBeEnabled(); await trigger.click(); await expect(dialog()).toBeVisible()
  await choose('添加轴（类别）字段', 'Region'); await choose('添加值字段', 'Amount')
  for (const label of ['显示行总计', '显示列总计', '显示行小计', '显示列小计']) await dialog().getByRole('checkbox', { name: label, exact: true }).uncheck()
}
async function pair(target = 'Reports', address = 'B9') {
  await choose('透视图创建方式', '同时创建透视表和透视图'); await choose('配对透视表放置工作表', target)
  await dialog().getByRole('textbox', { name: '配对透视表起始单元格', exact: true }).fill(address)
}
async function waitHistory(count) { await page.waitForFunction(count => window.pivotChartUIFixture.state().history.undo === count, count) }
async function setControls(fieldButtons, expandButtons) {
  await appearance()
  for (const [name, enabled] of [['显示字段按钮', fieldButtons], ['显示展开折叠按钮', expandButtons]]) {
    const control = dialog().getByRole('switch', { name, exact: true })
    const current = await control.getAttribute('aria-checked'); if ((current === 'true') !== enabled) await control.click()
    if (enabled) await expect(control).toBeChecked(); else await expect(control).not.toBeChecked()
  }
}

test.each(['classic', 'grid'].flatMap(variant => [false, true].map(paired => [variant, paired])))
  ('%s real Ribbon paired=%s creates one correct owner and one history entry', async (variant, paired) => {
    await fresh({ variant }); const before = await state()
    await dialog().getByRole('combobox', { name: '透视图创建方式', exact: true }).click()
    assert.deepEqual(await dropdown().locator('.ant-select-item-option-content').allTextContents(), ['仅创建透视图', '同时创建透视表和透视图'])
    await dropdown().getByText(paired ? '同时创建透视表和透视图' : '仅创建透视图', { exact: true }).click()
    if (paired) { await choose('配对透视表放置工作表', 'Dashboard'); await dialog().getByRole('textbox', { name: '配对透视表起始单元格', exact: true }).fill('B9') }
    await expect(create(paired)).toBeEnabled(); await create(paired).click(); await waitHistory(before.history.undo + 1)
    const after = await state(), chart = after.model.charts[0]
    assert.equal(chart.sheetId, ids.source); assert.deepEqual(sourceCells(after), sourceCells(before))
    const commands = after.commands.filter(command => command.id === 'sheet.command.add-pivot-chart'); assert.equal(commands.length, 1)
    assert.equal(commands[0].params.host, 'workbench'); assert.equal(commands[0].params.expectedRevision, before.revision)
    assert.equal(commands[0].params.expectedHistoryGeneration, before.generation); assert.equal(commands[0].params.expectedUserId, before.userId)
    if (paired) {
      assert.equal(after.model.pivotTables.length, 1); assert.equal(chart.ownedPivot, undefined); assert.equal(chart.pivotTableId, after.model.pivotTables[0].id)
      assert.deepEqual(after.model.pivotTables[0].anchor, { unitId: ids.unit, subUnitId: ids.other, row: 8, col: 1 })
      assert.deepEqual(commands[0].params.options.pivotTable.anchor, after.model.pivotTables[0].anchor)
      const cells = after.model.sheets.find(sheet => sheet.id === ids.other).cells
      assert.equal(cells['9:2'].value, 40); assert.equal(cells['10:2'].value, 60)
    } else { assert.ok(chart.ownedPivot); assert.equal(after.model.pivotTables.length, 0); assert.equal(commands[0].params.options, undefined) }
    await expect(dialog().getByRole('combobox', { name: '透视图创建方式', exact: true })).toHaveCount(0)
  })

test('new chart authoring from an existing report stays linked and does not offer paired creation', async () => {
  await page.evaluate(() => window.pivotChartUIFixture.open('', { pivotTableId: 'ui-primary-report', sourceKind: 'pivotTable' }))
  await expect(dialog()).toBeVisible(); await expect(dialog().getByRole('combobox', { name: '透视图创建方式', exact: true })).toHaveCount(0)
  const before = await state(); await expect(create(false)).toBeEnabled(); await create(false).click(); await waitHistory(before.history.undo + 1)
  const after = await state(), added = after.model.charts.find(chart => !before.model.charts.some(previous => previous.id === chart.id))
  assert.equal(added.pivotTableId, ids.primary); assert.equal(added.ownedPivot, undefined); assert.equal(after.model.pivotTables.length, before.model.pivotTables.length)
})

test('mode toggles preserve target, fields, author text and presentation without publishing a report', async () => {
  await fresh(); await dialog().getByRole('textbox', { name: 'Amount的显示名称', exact: true }).fill('Revenue'); await pair('Dashboard', 'C12')
  await setControls(false, true); const before = await state()
  await dialog().getByRole('textbox', { name: '图表标题', exact: true }).fill('Draft pair')
  await choose('透视图创建方式', '仅创建透视图'); await expect(dialog().getByRole('textbox', { name: '配对透视表起始单元格', exact: true })).toHaveCount(0)
  await choose('透视图创建方式', '同时创建透视表和透视图')
  await expect(dialog().getByRole('textbox', { name: '配对透视表起始单元格', exact: true })).toHaveValue('C12')
  await expect(dialog().getByRole('textbox', { name: '图表标题', exact: true })).toHaveValue('Draft pair')
  await expect(dialog().getByRole('switch', { name: '显示字段按钮', exact: true })).not.toBeChecked()
  await fieldsTab(); await expect(dialog().getByRole('textbox', { name: 'Amount的显示名称', exact: true })).toHaveValue('Revenue')
  assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
})

test.each(['B9:C10', 'XFE1', 'D3'])('invalid paired anchor %s blocks create and retains the editable target text', async address => {
  await fresh(); await pair(address === 'D3' ? 'Data' : 'Reports', address); const before = await state()
  await expect(create(true)).toBeDisabled(); await expect(dialog().getByRole('textbox', { name: '配对透视表起始单元格', exact: true })).toHaveValue(address)
  await expect(dialog().getByRole('alert').filter({ hasText: /起始位置|单个单元格/ }).first()).toBeVisible()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('full output protection rejects paired publication after the starting cell passed advisory validation', async () => {
  await page.evaluate(() => window.pivotChartUIFixture.mount({ existing: false })); const model = (await state()).model
  model.sheets.find(sheet => sheet.id === ids.host).rangeProtectionRules = [{ id: 'expanded-pair', permissionId: 'expanded-pair', description: '',
    ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', ranges: [{ startRow: 10, endRow: 10, startCol: 2, endCol: 2 }] }]
  await fresh({ model, currentUserId: 'local-user' }); await pair(); const before = await state()
  await expect(create(true)).toBeEnabled(); await create(true).click()
  await expect(dialog().getByRole('alert').filter({ hasText: /权限|保护|查看|编辑/ }).first()).toBeVisible()
  await expect(create(true)).not.toHaveClass(/ant-btn-loading/)
  const after = await state(); assert.deepEqual(after.model, before.model); assert.deepEqual(after.history, before.history)
  assert.equal(after.model.pivotTables.length, 0); assert.equal(after.model.charts.length, 0)
})

test.each([false, true].flatMap(paired => ['source', 'history', 'user', 'host'].map(change => [paired, change])))
  ('pending range creation paired=%s refuses changed %s without publishing either resource', async (paired, change) => {
    await fresh(); if (paired) await pair(); await page.evaluate(() => window.pivotChartUIFixture.delayMutation('insert'))
    await expect(create(paired)).toBeEnabled(); await create(paired).click(); await page.waitForFunction(() => window.pivotChartUIFixture.pendingMutations('insert') === 1)
    await page.evaluate(async change => {
      if (change === 'source') await window.pivotChartUIFixture.changeSource()
      else if (change === 'history') await window.pivotChartUIFixture.clearHistory()
      else if (change === 'user') await window.pivotChartUIFixture.changeUserWhilePending()
      else await window.pivotChartUIFixture.switchSheet()
    }, change)
    const newer = await state(); await page.evaluate(() => window.pivotChartUIFixture.releaseMutation('insert'))
    const after = await state(); assert.deepEqual(after.model, newer.model); assert.deepEqual(after.history, newer.history)
    assert.equal(after.model.pivotTables.length, 0); assert.equal(after.model.charts.length, 0)
  })

test('an old insertion completion cannot release a new pending paired creation after host navigation', async () => {
  await fresh(); await pair(); await page.evaluate(() => window.pivotChartUIFixture.delayMutation('insert'))
  await expect(create(true)).toBeEnabled(); await create(true).click(); await page.waitForFunction(() => window.pivotChartUIFixture.pendingMutations('insert') === 1)
  await page.evaluate(() => window.pivotChartUIFixture.switchSheet())
  await page.evaluate(() => window.pivotChartUIFixture.open('', { sourceKind: 'range', sourceSheetId: 'data', sourceRange: { startRow: 2, endRow: 6, startCol: 3, endCol: 7 } }))
  await expect(dialog()).toBeVisible(); await choose('添加轴（类别）字段', 'Region'); await choose('添加值字段', 'Amount'); await pair('Reports', 'B9')
  await expect(create(true)).toBeEnabled(); await create(true).click(); await page.waitForFunction(() => window.pivotChartUIFixture.pendingMutations('insert') === 2)
  const before = await state(); await page.evaluate(() => window.pivotChartUIFixture.releaseMutation('insert', false))
  await expect(create(true)).toHaveClass(/ant-btn-loading/); assert.deepEqual((await state()).model, before.model)
  await page.evaluate(() => window.pivotChartUIFixture.releaseMutation('insert')); await waitHistory(before.history.undo + 1)
  assert.equal((await state()).model.pivotTables.length, 1); assert.equal((await state()).model.charts.length, 1)
})

test.each([ids.owned, ids.linked].flatMap(chartId => [[false, true], [true, false], [false, false]].map(([a, b]) => [chartId, a, b])))
  ('%s two Ant presentation controls [%s,%s] use one configure without editing report content', async (chartId, a, b) => {
    await page.evaluate(chartId => window.pivotChartUIFixture.open(chartId), chartId); await expect(dialog()).toBeVisible()
    const before = await state(), originalReport = report(before, chartId); await setControls(a, b)
    assert.deepEqual((await state()).model, before.model); await expect(apply()).toBeEnabled(); await apply().click(); await waitHistory(before.history.undo + 1)
    const after = await state(); assert.deepEqual(after.model.charts.find(chart => chart.id === chartId).controls, { showFieldButtons: a, showExpandCollapseButtons: b })
    assert.deepEqual(report(after, chartId), originalReport); assert.deepEqual(sourceCells(after), sourceCells(before))
    assert.equal(after.commands.filter(command => command.id === 'sheet.command.configure-pivot-chart').length, 1)
    assert.equal(after.commands.filter(command => command.id === 'sheet.command.update-pivot-chart-controls').length, 0)
  })

test('paired creation includes presentation choices and restores them through history and snapshot reopen', async () => {
  await fresh(); await pair(); await setControls(false, false); const before = await state()
  await expect(create(true)).toBeEnabled(); await create(true).click(); await waitHistory(before.history.undo + 1)
  const added = (await state()).model.charts[0]; assert.deepEqual(added.controls, { showFieldButtons: false, showExpandCollapseButtons: false })
  await dialog().getByRole('button', { name: '关闭', exact: true }).click(); await expect(dialog()).not.toBeVisible()
  await page.evaluate(() => window.pivotChartUIFixture.undo()); assert.equal((await state()).model.charts.length, 0); assert.equal((await state()).model.pivotTables.length, 0)
  await page.evaluate(() => window.pivotChartUIFixture.redo()); assert.equal((await state()).model.pivotTables.length, 1)
  await page.evaluate(() => window.pivotChartUIFixture.reopen()); await page.evaluate(chartId => window.pivotChartUIFixture.open(chartId), added.id)
  await expect(dialog()).toBeVisible(); await appearance()
  await expect(dialog().getByRole('switch', { name: '显示字段按钮', exact: true })).not.toBeChecked()
  await expect(dialog().getByRole('switch', { name: '显示展开折叠按钮', exact: true })).not.toBeChecked()
})

test.each(['classic', 'grid'])('%s theme and synthetic anchor composition preserve draft while creation/navigation are disabled', async variant => {
  await fresh({ variant }); await pair(); await appearance(); const before = await state()
  const input = dialog().getByRole('textbox', { name: '配对透视表起始单元格', exact: true })
  await input.focus(); await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await expect(create(true)).toBeDisabled(); await expect(dialog().getByRole('combobox', { name: '透视图创建方式', exact: true })).toBeDisabled()
  await expect(dialog().getByRole('switch', { name: '显示字段按钮', exact: true })).toBeDisabled()
  const light = await dialog().getByTestId('pivot-chart-panel').evaluate(node => getComputedStyle(node).color)
  await page.evaluate(variant => window.pivotChartUIFixture.appearance(true, variant), variant)
  const dark = await dialog().getByTestId('pivot-chart-panel').evaluate(node => getComputedStyle(node).color); assert.notEqual(dark, light)
  await input.evaluate(node => {
    node.value = 'C12'; node.dispatchEvent(new InputEvent('input', { bubbles: true, data: 'C12', inputType: 'insertCompositionText', isComposing: true }))
    node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'C12' }))
    node.dispatchEvent(new InputEvent('input', { bubbles: true, data: 'C12', inputType: 'insertText' }))
  })
  await expect(input).toHaveValue('C12'); await expect(create(true)).toBeEnabled(); assert.deepEqual((await state()).model, before.model)
})

test('closing an uncommitted paired draft leaves neither resource and returns to workbench Canvas focus', async () => {
  await fresh(); await pair(); const before = await state(); await dialog().getByRole('button', { name: '关闭', exact: true }).click()
  await expect(dialog()).not.toBeVisible(); await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeFocused()
  assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
})

test('the table destination selector excludes a worksheet the current user cannot edit', async () => {
  await fresh({ readonly: true, currentUserId: 'peer' }); await choose('透视图创建方式', '同时创建透视表和透视图')
  await dialog().getByRole('combobox', { name: '配对透视表放置工作表', exact: true }).click()
  await expect(dropdown().getByText('Reports', { exact: true })).toHaveCount(0); await expect(dropdown().getByText('Dashboard', { exact: true })).toBeVisible()
  await page.keyboard.press('Escape'); assert.equal((await state()).model.pivotTables.length, 0)
})

test('source editing marks a paired draft stale and retains its target until a deliberate reread', async () => {
  await fresh(); await pair('Dashboard', 'C12'); await page.evaluate(() => window.pivotChartUIFixture.changeSource())
  await expect(create(true)).toBeDisabled(); await expect(dialog().getByRole('textbox', { name: '配对透视表起始单元格', exact: true })).toHaveValue('C12')
  await dialog().getByRole('button', { name: '重新读取', exact: true }).click()
  await expect(dialog().getByRole('textbox', { name: '配对透视表起始单元格', exact: true })).toHaveValue('C12')
  await expect(dialog().getByRole('combobox', { name: '透视图创建方式', exact: true })).toBeEnabled(); assert.equal((await state()).model.pivotTables.length, 0)
})
