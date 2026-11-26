import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'
import { pivotChartFieldButtonsIds as ids } from './fixtures/pivot-chart-field-buttons-fixture.mjs'

// Deferred real Ribbon/ChartOverlay/Ant/Canvas contracts. No browser or checks have run.
// Planned backend is Canvas2D. Synthetic composition does not prove native Windows IME behavior.
let server, browser, context, page, base, kind
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No Values fixture port allocated')
  base = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1200 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
})
afterEach(async () => {
  try { await page?.evaluate(() => window.pivotChartFieldButtonsFixture?.dispose() ?? window.pivotTableFieldFilterFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
async function mount(workbench = 'chart', options = {}) {
  kind = workbench; const name = workbench === 'table' ? 'pivot-table-field-filter-workbench.html' : 'pivot-chart-field-buttons-workbench.html'
  await page.goto(base + name); await page.waitForFunction(kind => Boolean(kind === 'table' ? window.pivotTableFieldFilterFixture : window.pivotChartFieldButtonsFixture), kind)
  await page.evaluate(({ kind, options }) => (kind === 'table' ? window.pivotTableFieldFilterFixture : window.pivotChartFieldButtonsFixture).mount({ valuesUi: true, ...options }), { kind, options })
}
const state = () => page.evaluate(kind => (kind === 'table' ? window.pivotTableFieldFilterFixture : window.pivotChartFieldButtonsFixture).state(), kind)
const tablePanel = () => page.getByTestId('pivot-table-panel')
const chartPanel = () => page.getByTestId('pivot-chart-panel')
const root = (chartId = ids.owned) => page.locator(`[data-testid="pivot-chart-field-buttons"][data-pivot-chart="${chartId}"]`)
const valuesCard = () => page.getByTestId('pivot-chart-values-field-button')
const report = value => kind === 'table' ? value.model.pivotTables.find(report => report.id === ids.primary)
  : value.model.charts.find(chart => chart.id === value.chartId)?.ownedPivot ?? value.model.pivotTables.find(report => report.id === ids.primary)
const source = value => value.model.sheets.find(sheet => sheet.id === ids.source).cells
function assertRendered(value, position, index) {
  const matrix = kind === 'table' ? value.data : value.reportMatrix, effective = index < 0 ? 2 : Math.min(index, 2)
  if (position === 'row') { const labels = ['Region', 'Date']; labels.splice(effective, 0, 'Values'); assert.deepEqual(matrix[1].slice(0, 3), labels) }
  else assert.deepEqual([...new Set(matrix[effective].slice(2))].sort(), ['Amount average', 'Amount sum'])
}
async function choose(target, name, label) {
  await target.getByRole('combobox', { name, exact: true }).click(); await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: typeof label === 'string' }).click()
}
const indexLabel = index => index === -1 ? '自动末尾' : index === 0 ? '第 1 个位置（首位）' : index === 2 ? '第 3 个位置（末位）' : '第 2 个位置'
async function openTable() {
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  await page.getByRole('toolbar', { name: '数据', exact: true }).getByRole('button', { name: '数据透视表', exact: true }).click()
  await expect(tablePanel()).toBeVisible(); await choose(tablePanel(), '选择数据透视表', /^数据透视表 1 ·/)
}
async function openChart(chartId = ids.owned) {
  await page.evaluate(chartId => window.pivotChartFieldButtonsFixture.selectChart(chartId), chartId)
  await root(chartId).getByRole('button', { name: '打开透视图字段Amount sum', exact: true }).click()
  await page.getByTestId('pivot-chart-field-filter-popup').getByRole('button', { name: '透视图字段设置', exact: true }).click()
  await expect(chartPanel()).toBeVisible()
}
async function direction(target, subject, position, index) {
  await choose(target, subject === '透视表' ? '值字段排列' : '值字段放置方向', position === 'row' ? '按行显示' : position === 'column' ? '按列显示' : '无 ΣValues 位置')
  if (position !== 'none' && index !== undefined) await choose(target, `${subject} ΣValues 插入位置`, indexLabel(index))
}
async function openValuesMenu(chartId = ids.owned) {
  const direct = root(chartId).getByTestId('pivot-chart-values-field-button')
  if (!await direct.count()) {
    const current = await state(), area = report(current).options.valuePosition
    await root(chartId).getByRole('button', { name: area === 'row' ? '更多轴（类别）字段' : '更多图例（系列）字段', exact: true }).click()
  }
  await valuesCard().getByRole('button', { name: '透视图 ΣValues 位置菜单', exact: true }).click()
  await expect(page.getByRole('menuitem', { name: '放到行区域', exact: true })).toBeVisible()
}
async function menuIndex(index) {
  await page.getByRole('menuitem', { name: 'ΣValues 插入位置', exact: true }).hover()
  await page.getByRole('menuitem', { name: indexLabel(index), exact: true }).click()
}
async function waitHistory(count) { await page.waitForFunction(({ kind, count }) => (kind === 'table' ? window.pivotTableFieldFilterFixture : window.pivotChartFieldButtonsFixture).state().history.undo === count, { kind, count }) }
async function fixtureAction(method, ...args) { return page.evaluate(({ kind, method, args }) => (kind === 'table' ? window.pivotTableFieldFilterFixture : window.pivotChartFieldButtonsFixture)[method](...args), { kind, method, args }) }
async function changeOwner(change) {
  if (change === 'source') await fixtureAction('changeSource')
  if (change === 'history') await fixtureAction('clearHistory')
  if (change === 'user') await fixtureAction('changeUserWhilePending')
  if (change === 'host') await fixtureAction('switchSheet')
}

test.each(['classic', 'grid'].flatMap(variant => ['row', 'column'].flatMap(position => [0, 1, 2, -1].map(index => [variant, position, index]))))
  ('%s ordinary %s/%s uses one real configure for Values, pending destination and history/reopen', async (variant, position, index) => {
    await mount('table', { variant }); await openTable(); const before = await state()
    await tablePanel().getByRole('textbox', { name: '透视表起始单元格', exact: true }).fill('K30'); await direction(tablePanel(), '透视表', position, index)
    assert.deepEqual((await state()).model, before.model); const marker = tablePanel().getByTestId('pivot-table-values-field')
    await expect(marker).toHaveAttribute('data-values-position', position); await expect(marker).toHaveAttribute('data-values-index', String(index))
    await tablePanel().getByRole('button', { name: '应用字段配置', exact: true }).click(); await waitHistory(1)
    const saved = await state(); assertRendered(saved, position, index); assert.deepEqual(source(saved), source(before)); assert.equal(report(saved).anchor.col, 10)
    assert.equal(saved.commands.filter(command => command.id === 'sheet.command.configure-pivot-table').length, 1)
    await fixtureAction('undo'); assert.deepEqual((await state()).model, before.model); await fixtureAction('redo'); assert.deepEqual((await state()).model, saved.model)
    await tablePanel().getByRole('button', { name: '关闭', exact: true }).click(); await fixtureAction('reopen'); await openTable()
    await expect(tablePanel().getByTestId('pivot-table-values-field')).toHaveAttribute('data-values-index', String(index))
  })

test('ordinary None hides the special item, retains measures and saves the column fallback', async () => {
  await mount('table', { valueIndex: 1 }); await openTable(); const before = await state(); await direction(tablePanel(), '透视表', 'none')
  await expect(tablePanel().getByTestId('pivot-table-values-field')).toHaveCount(0)
  await expect(tablePanel().getByRole('combobox', { name: '透视表 ΣValues 插入位置', exact: true })).toBeDisabled()
  await tablePanel().getByRole('button', { name: '应用字段配置', exact: true }).click(); await waitHistory(1)
  const current = await state(); assert.equal(report(current).fields.filter(field => field.area === 'value').length, 2); assert.equal(report(current).options.valuePosition, 'none')
  assert.deepEqual(current.data, before.data); await fixtureAction('undo'); assert.deepEqual((await state()).model, before.model)
})

test.each(['row', 'column'])('ordinary %s raw out-of-range index remains a visible retained option with a clamped marker', async position => {
  await mount('table', { valuePosition: position, valueIndex: 99 }); await openTable()
  await expect(tablePanel().getByTestId('pivot-table-values-field')).toHaveAttribute('data-values-index', '99')
  await expect(tablePanel().getByTestId('pivot-table-values-field')).toHaveAttribute('data-values-effective-index', '2')
  await expect(tablePanel().getByRole('combobox', { name: '透视表 ΣValues 插入位置', exact: true })).toHaveAttribute('aria-expanded', 'false')
  await tablePanel().getByRole('combobox', { name: '透视表 ΣValues 插入位置', exact: true }).click()
  await expect(page.locator('.ant-select-dropdown:visible').getByText('指定位置 100（当前末位）', { exact: true })).toBeVisible(); assert.equal(report(await state()).options.valueIndex, 99)
})

test.each(['row', 'column'])('ordinary ΣValues drag to %s inserts before the real field and remains a draft until configure', async position => {
  await mount('table', { valuePosition: position === 'row' ? 'column' : 'row', valueIndex: 1 }); await openTable(); const before = await state()
  const name = position === 'row' ? 'Region' : 'Segment', field = report(before).fields.find(field => field.name === name)
  await tablePanel().getByTestId('pivot-table-values-field').dragTo(tablePanel().locator(`.pivot-area-grid [data-field-id="${field.id}"]`))
  await expect(tablePanel().getByTestId('pivot-table-values-field')).toHaveAttribute('data-values-position', position)
  await expect(tablePanel().getByTestId('pivot-table-values-field')).toHaveAttribute('data-values-index', '0'); assert.deepEqual((await state()).model, before.model)
  await tablePanel().getByRole('button', { name: '应用字段配置', exact: true }).click(); await waitHistory(1); assertRendered(await state(), position, 0)
})

test.each(['筛选', '值'])('ordinary ΣValues cannot be dropped into the %s area as a fake source field', async area => {
  await mount('table', { valueIndex: 1 }); await openTable(); const before = await state()
  await tablePanel().getByTestId('pivot-table-values-field').dragTo(tablePanel().locator(`.pivot-area-grid section[aria-label="${area}区域"]`))
  await expect(tablePanel().getByTestId('pivot-table-values-field')).toHaveAttribute('data-values-position', 'column'); assert.deepEqual((await state()).model, before.model)
})

test.each(['row', 'column'])('ordinary %s real field move retains the shared Values gap and pending location', async position => {
  await mount('table', { valuePosition: position, valueIndex: 1 }); await openTable()
  await tablePanel().getByRole('textbox', { name: '透视表起始单元格', exact: true }).fill('K30')
  const name = position === 'row' ? 'Region' : 'Segment', current = await state(), field = report(current).fields.find(field => field.name === name)
  await tablePanel().locator(`.pivot-area-grid [data-field-id="${field.id}"]`).getByRole('button', { name: `字段 ${name} 操作`, exact: true }).click()
  await page.getByRole('menuitem', { name: '下移', exact: true }).click(); await expect(tablePanel().getByTestId('pivot-table-values-field')).toHaveAttribute('data-values-index', '0')
  await expect(tablePanel().getByRole('textbox', { name: '透视表起始单元格', exact: true })).toHaveValue('K30'); assert.equal((await state()).history.undo, 0)
})

test.each(['owned', 'linked'].flatMap(ownership => ['row', 'column'].flatMap(position => [0, 1, 2, -1].map(index => [ownership, position, index]))))
  ('%s chart panel %s/%s commits index and direction with one configure and intact chart appearance', async (ownership, position, index) => {
    const chartId = ownership === 'owned' ? ids.owned : ids.linked; await mount('chart', { chartId, valueIndex: 99 }); await openChart(chartId)
    const before = await state(), chart = before.model.charts.find(chart => chart.id === chartId); await direction(chartPanel(), '透视图', position, index)
    assert.deepEqual((await state()).model, before.model); await expect(chartPanel().getByTestId('pivot-chart-values-field')).toHaveAttribute('data-values-index', String(index))
    const apply = chartPanel().getByRole('button', { name: '应用透视图', exact: true }); await expect(apply).toBeEnabled(); await apply.click(); await waitHistory(1)
    const saved = await state(); assert.deepEqual(saved.valuesPosition, { position, index }); assertRendered(saved, position, index)
    const changedChart = saved.model.charts.find(value => value.id === chartId)
    assert.deepEqual(changedChart.options.threeD, chart.options.threeD); assert.deepEqual(changedChart.options.series, chart.options.series)
    assert.equal(changedChart.title, chart.title); assert.equal(changedChart.type, chart.type); assert.deepEqual(source(saved), source(before))
    assert.equal(saved.commands.filter(command => command.id === 'sheet.command.configure-pivot-chart').length, 1)
    await fixtureAction('undo'); assert.deepEqual((await state()).model, before.model)
  })

test.each(['owned', 'linked'])('%s chart panel None preserves all measures and survives snapshot reopen', async ownership => {
  const chartId = ownership === 'owned' ? ids.owned : ids.linked; await mount('chart', { chartId, valueIndex: 1 }); await openChart(chartId)
  const before = await state(); await direction(chartPanel(), '透视图', 'none'); await chartPanel().getByRole('button', { name: '应用透视图', exact: true }).click(); await waitHistory(1)
  const saved = await state(); assert.deepEqual(saved.valuesPosition, { position: 'none', index: 1 }); assert.deepEqual(saved.reportMatrix, before.reportMatrix)
  assert.equal(report(saved).fields.filter(field => field.area === 'value').length, 2); await expect(chartPanel().getByTestId('pivot-chart-values-field')).toHaveCount(0)
  await chartPanel().getByRole('button', { name: '关闭', exact: true }).click(); await fixtureAction('reopen'); await openChart(chartId)
  await expect(chartPanel().getByRole('combobox', { name: '透视图 ΣValues 插入位置', exact: true })).toBeDisabled(); assert.equal(report(await state()).options.valuePosition, 'none')
})

test.each(['owned', 'linked'])('%s single measure retains legacy row preference while the synthetic field stays hidden', async ownership => {
  const chartId = ownership === 'owned' ? ids.owned : ids.linked; await mount('chart', { chartId, multipleValues: false, valuePosition: 'row', valueIndex: 1 })
  await expect(valuesCard()).toHaveCount(0); const current = await state(); assert.deepEqual(current.valuesPosition, { position: 'row', index: 1 })
  assert.equal(current.reportMatrix[1][1], 'Values'); assert.equal(report(current).fields.filter(field => field.area === 'value').length, 1)
})

test.each(['owned', 'linked'].flatMap(ownership => ['row', 'column'].flatMap(position => [0, 1].map(index => [ownership, position, index]))))
  ('%s overlay Values menu %s/%s uses a real owner-bearing report edit and actual axis placement', async (ownership, position, index) => {
    const chartId = ownership === 'owned' ? ids.owned : ids.linked; await mount('chart', { chartId, valuePosition: position, valueIndex: 99 })
    const before = await state(); await openValuesMenu(chartId); await menuIndex(index); await waitHistory(1)
    const current = await state(); assert.deepEqual(current.valuesPosition, { position, index }); assertRendered(current, position, index)
    assert.deepEqual(source(current), source(before)); const command = current.commands.find(command => command.id === 'sheet.command.edit-pivot-chart-report')
    assert.equal(command.params.edit.type, 'updateValuePosition'); assert.equal(command.params.edit.index, index); assert.equal(command.params.expectedUserId, 'local-user')
    await fixtureAction('undo'); assert.deepEqual((await state()).model, before.model)
  })

test.each(['owned', 'linked'].flatMap(ownership => ['row', 'column'].map(position => [ownership, position])))
  ('%s overlay native Values drag to %s submits one guarded position edit', async (ownership, position) => {
    const chartId = ownership === 'owned' ? ids.owned : ids.linked; await mount('chart', { chartId, valuePosition: position === 'row' ? 'column' : 'row', valueIndex: 1 })
    const name = position === 'row' ? 'Region' : 'Segment', before = await state()
    await valuesCard().dragTo(root(chartId).getByRole('button', { name: `打开透视图字段${name}`, exact: true })); await waitHistory(1)
    assert.deepEqual((await state()).valuesPosition, { position, index: 0 }); assertRendered(await state(), position, 0); assert.deepEqual(source(await state()), source(before))
  })

test('Values in the same-area overflow remains a layout item and can be moved to the other axis', async () => {
  await mount('chart', { valueIndex: -1 }); const before = await state(); await openValuesMenu()
  await expect(page.getByTestId('pivot-chart-field-overflow')).toBeVisible()
  await page.getByRole('menuitem', { name: '放到行区域', exact: true }).click(); await waitHistory(1)
  const current = await state(); assert.deepEqual(current.valuesPosition, { position: 'row', index: -1 })
  assert.equal(report(current).fields.length, report(before).fields.length); assert.equal(report(current).fields.filter(field => field.area === 'value').length, 2)
})

test.each([{ readonly: true }, { protectedOutput: true }, { protectedChart: true }])('viewable protection %j keeps Values visible and blocks menu/drag edits', async options => {
  const chartId = options.protectedOutput ? ids.linked : ids.owned; await mount('chart', { ...options, chartId, valueIndex: 0 }); const before = await state()
  await expect(valuesCard().getByRole('button', { name: '透视图 ΣValues 位置菜单', exact: true })).toBeDisabled()
  await expect(valuesCard()).toHaveAttribute('draggable', 'false'); assert.deepEqual((await state()).model, before.model); assert.equal(before.history.undo, 0)
})

test.each([{ hiddenSource: true }, { hiddenOutput: true }])('unreadable %j report cannot expose a synthetic Values button', async options => {
  await mount('chart', { ...options, chartId: ids.linked, valueIndex: 0 }); await expect(valuesCard()).toHaveCount(0); assert.equal((await state()).history.undo, 0)
})

test.each(['owned', 'linked'].flatMap(ownership => ['source', 'history', 'user', 'host'].map(change => [ownership, change])))
  ('pending %s overlay Values rejects changed %s owner without partial report/cells/history', async (ownership, change) => {
    const chartId = ownership === 'owned' ? ids.owned : ids.linked; await mount('chart', { chartId, valueIndex: 0 }); await openValuesMenu(chartId)
    await fixtureAction('delayEdits'); await page.getByRole('menuitem', { name: '放到行区域', exact: true }).click()
    await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingEdits().length === 1); await changeOwner(change); const before = await state()
    await fixtureAction('releaseEdit'); assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  })

test('old overlay Values finally cannot clear a newer busy edit after a history reread', async () => {
  await mount('chart', { valueIndex: 0 }); await openValuesMenu(); await fixtureAction('delayEdits')
  await page.getByRole('menuitem', { name: '放到行区域', exact: true }).click(); await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingEdits().length === 1)
  await fixtureAction('clearHistory'); await openValuesMenu(); await menuIndex(1); await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingEdits().length === 2)
  await fixtureAction('releaseEdit', 0, false); await expect(valuesCard().getByRole('button', { name: '透视图 ΣValues 位置菜单', exact: true })).toBeDisabled()
  await fixtureAction('releaseEdit'); await waitHistory(1); assert.deepEqual((await state()).valuesPosition, { position: 'column', index: 1 })
})

test.each(['classic', 'grid'])('%s Values marker and Ant menu inherit theme tokens without changing content', async variant => {
  await mount('chart', { variant, valueIndex: 0 }); const before = await state(), light = await valuesCard().evaluate(node => getComputedStyle(node).color)
  await fixtureAction('appearance', true, variant); assert.notEqual(await valuesCard().evaluate(node => getComputedStyle(node).color), light)
  await openValuesMenu(); await expect(page.getByRole('menuitem', { name: 'ΣValues 插入位置', exact: true })).toBeVisible(); assert.deepEqual((await state()).model, before.model)
})

test.each(['table', 'chart'])('%s Ant composition blocks Values edits and preserves the stale input until composition ends', async workbench => {
  await mount(workbench, { valueIndex: 1 }); if (workbench === 'table') await openTable(); else await openChart()
  const target = workbench === 'table' ? tablePanel() : chartPanel()
  const input = workbench === 'table' ? target.getByRole('textbox', { name: '透视表起始单元格', exact: true }) : target.locator('.field-card').first().getByRole('textbox').first()
  await input.focus(); await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  const name = workbench === 'table' ? '透视表 ΣValues 插入位置' : '透视图 ΣValues 插入位置'; await expect(target.getByRole('combobox', { name, exact: true })).toBeDisabled()
  await input.evaluate(node => { node.value = '中'; node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '中', isComposing: true, inputType: 'insertCompositionText' })) })
  await fixtureAction('changeSource'); await expect(input).toHaveValue('中'); await expect(input).toBeEnabled()
  await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中' })))
  await expect(target.getByRole('combobox', { name, exact: true })).toBeDisabled()
})

test('overlay Values Escape closes its menu and restores current Canvas focus', async () => {
  await mount('chart', { valueIndex: 0 }); await openValuesMenu()
  await page.getByRole('menuitem', { name: '放到行区域', exact: true }).press('Escape'); await expect(page.locator('.ant-dropdown:visible')).toHaveCount(0)
  await expect(page.locator('canvas:focus')).toHaveCount(1); assert.equal((await state()).history.undo, 0)
})

test('disposing the real workbench removes a pending Values menu and rejects the late command owner', async () => {
  await mount('chart', { valueIndex: 0 }); await openValuesMenu(); await fixtureAction('delayEdits')
  await page.getByRole('menuitem', { name: '放到行区域', exact: true }).click(); await page.waitForFunction(() => window.pivotChartFieldButtonsFixture.pendingEdits().length === 1)
  await fixtureAction('dispose'); await expect(page.locator('#app')).toBeEmpty(); await expect(valuesCard()).toHaveCount(0)
})
