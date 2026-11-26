import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// Deferred real Ant/Workbench/Canvas2D/Facade/date-source/history drafts. Nothing has been started or checked.
let server, browser, context, page, url
const errors = [], ids = { primary: 'timeline-primary', shared: 'timeline-copy', slicer: 'timeline-mixed-slicer', host: 'dashboard', source: 'data', other: 'output' }
const conditionOptions = [['equal', '等于', 80], ['notEqual', '不等于', 210], ['newerThan', '晚于', 200], ['newerThanOrEqual', '晚于或等于', 280],
  ['olderThan', '早于', 10], ['olderThanOrEqual', '早于或等于', 90], ['between', '介于（含两端）', 110], ['notBetween', '不介于（含两端）', 180]]
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No Timeline settings fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/timeline-settings-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1100 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.timelineSettingsFixture?.state().model.slicers.length === 3)
}, 60_000)
afterEach(async () => {
  try { await page?.evaluate(() => window.timelineSettingsFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const action = (method, ...args) => page.evaluate(({ method, args }) => window.timelineSettingsFixture[method](...args), { method, args })
const state = () => action('state'), form = () => page.getByTestId('slicer-settings-form'), filter = () => page.getByRole('dialog', { name: '时间线日期筛选', exact: true })
const toolbar = () => page.getByRole('toolbar', { name: '切片器格式', exact: true })
const info = (value, id = ids.primary) => value.model.slicers.find(control => control.id === id)
const first = value => Math.floor(value.model.sheets.find(sheet => sheet.id === ids.source).cells['3:4'].value)
function retainedDisplay(value) {
  return { sheets: value.model.sheets.map(({ id, cells }) => ({ id, cells })), pivots: value.model.pivotTables, charts: value.model.charts,
    caches: value.model.slicerCaches, sources: value.model.slicers.map(({ id, sourceFieldId, pivotTableId, cacheId }) => ({ id, sourceFieldId, pivotTableId, cacheId })) }
}
async function mount(options = {}) { await action('mount', options) }
async function openSettings(id = ids.primary) {
  await action('select', id); await toolbar().getByRole('button', { name: '时间线设置', exact: true }).click(); await expect(form()).toBeVisible()
}
async function openFilter(id = ids.primary) {
  await openSettings(id); await form().getByRole('button', { name: '日期筛选', exact: true }).click(); await expect(filter()).toBeVisible()
}
async function option(container, name, label) {
  await container.getByRole('combobox', { name, exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}
async function serial(container, fieldLabel, ariaLabel, value) {
  const field = container.locator('.timeline-date-input').filter({ has: page.getByText(fieldLabel, { exact: true }) })
  const toggle = field.getByRole('button', { name: '精确日期值', exact: true })
  if (await toggle.isVisible()) await toggle.click()
  const input = field.getByRole('spinbutton', { name: `${ariaLabel}精确日期值`, exact: true })
  await input.fill(value === undefined ? '' : String(value)); await input.press('Tab')
}
async function number(name, value) { const input = form().getByRole('spinbutton', { name, exact: true }); await input.fill(String(value)); await input.press('Tab') }
async function caption(value) {
  const inherited = form().getByRole('checkbox', { name: '使用名称作为标题', exact: true })
  if (await inherited.isChecked()) await inherited.uncheck()
  await form().getByRole('textbox', { name: '标题', exact: true }).fill(value)
}
async function waitHistory(expected) { await page.waitForFunction(expected => window.timelineSettingsFixture.state().history.undo === expected, expected) }
async function waitPending(kind, count = 1) { await page.waitForFunction(({ kind, count }) => window.timelineSettingsFixture.pending(kind) === count, { kind, count }) }
async function editRange(startDate, endDate, wholeDay = false) {
  await option(filter(), '时间线筛选方式', '日期范围（结束值不含）')
  await serial(filter(), '开始日期（含）', '时间线开始日期', startDate); await serial(filter(), '结束日期（不含）', '时间线结束日期', endDate)
  const checkbox = filter().getByRole('checkbox', { name: '按整日比较', exact: true }); if (wholeDay) await checkbox.check(); else if (await checkbox.isChecked()) await checkbox.uncheck()
}
async function editCondition(operator, label, startDate, endDate) {
  await option(filter(), '时间线筛选方式', '日期条件'); await option(filter(), '时间线日期条件', label)
  const between = operator === 'between' || operator === 'notBetween'
  await serial(filter(), between ? '开始日期（含）' : '比较日期', '时间线条件日期', startDate)
  if (between) await serial(filter(), '结束日期（含）', '时间线条件结束日期', endDate)
}
async function changeOwner(kind) {
  if (kind === 'source') await action('changeSource')
  if (kind === 'history') await action('clearHistory')
  if (kind === 'user') await action('switchUser', 'peer')
  if (kind === 'host') await action('switchSheet', ids.other)
}

test.each([['pivot', 'Output · 透视表 1'], ['ownedPivotChart', 'Dashboard · 透视图 1']])
  ('real %s insertion uses the shared source dialog and formal Timeline Info before Canvas/history publish', async (kind, sourceLabel) => {
    await mount(); const before = await state()
    await page.getByRole('tab', { name: '数据', exact: true }).click()
    await page.getByRole('toolbar', { name: '数据', exact: true }).getByRole('button', { name: '时间线', exact: true }).click()
    const insert = page.getByRole('dialog', { name: '插入时间线', exact: true }); await expect(insert).toBeVisible()
    await option(insert, '切片器来源', sourceLabel); await option(insert, '字段', 'Date')
    assert.deepEqual((await state()).model, before.model)
    await insert.getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
    const saved = await state(), created = saved.model.slicers.find(view => !before.model.slicers.some(prior => prior.id === view.id))
    assert.equal(created.type, 'timeline'); assert.equal(created.sourceFieldId, 'date-source')
    const event = saved.commands.find(command => command.id === 'sheet.command.insert-timeline')
    assert.equal(event.params.info.dataSource.kind, 'pivot'); assert.equal(Object.hasOwn(event.params.info, 'type'), false)
    if (kind === 'ownedPivotChart') assert.deepEqual(event.params.info.dataSource.ownedPivotCharts, [{ subUnitId: ids.host, pivotChartId: 'owned-chart' }])
    await expect(page.locator(`[data-slicer-id="${created.id}"]`).first()).toBeVisible()
    await action('undo'); assert.deepEqual((await state()).model, before.model)
  })

test.each(['classic', 'grid'].flatMap(variant => [ids.primary, ids.shared].map(id => [variant, id])))
  ('%s %s settings edit all ten formal values once while shared filtering and source stay intact', async (variant, id) => {
    await mount({ variant, viewDefaults: true }); await openSettings(id); const before = await state(), day = first(before), otherId = id === ids.primary ? ids.shared : ids.primary
    await form().getByRole('textbox', { name: '时间线名称', exact: true }).fill(`UI calendar ${id}`); await caption('UI calendar caption')
    await option(form(), '时间级别', '年'); await option(form(), '选择级别', '日')
    await serial(form(), '导航日期', '时间线导航日期', day + .123456789)
    for (const name of ['显示标题', '显示选择标签', '显示时间级别', '显示滚动条']) await form().getByRole('checkbox', { name, exact: true }).uncheck()
    await form().getByRole('button', { name: '选择样式', exact: true }).click()
    await page.getByRole('group', { name: '样式图库', exact: true }).getByRole('button', { name: 'Calendar bands', exact: true }).click()
    assert.deepEqual((await state()).model, before.model)
    await form().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
    const saved = await state(), settings = saved.timelines[id].settings
    assert.deepEqual(settings, { name: `UI calendar ${id}`, caption: 'UI calendar caption', level: 0, selectionLevel: 3,
      scrollPosition: day + .123456789, style: 'Calendar bands', showHeader: false, showSelectionLabel: false, showTimeLevel: false, showHorizontalScrollbar: false })
    assert.deepEqual(saved.timelines[otherId].settings, before.timelines[otherId].settings); assert.deepEqual(retainedDisplay(saved), retainedDisplay(before))
    assert.deepEqual(info(saved, id).drawingGeometry, info(before, id).drawingGeometry)
    assert.equal(saved.commands.filter(command => command.id === 'sheet.command.update-timeline-settings').length, 1)
    await action('undo'); assert.deepEqual((await state()).model, before.model); await action('redo'); assert.deepEqual((await state()).timelines[id].settings, settings)
  })

test('mixed Timeline display and projected geometry save one presentation command and undo restores authored affine data', async () => {
  await mount(); await openSettings(); const before = await state(); await caption('Mixed Timeline properties')
  await option(form(), '时间级别', '月'); await form().getByText('大小与属性', { exact: true }).click()
  await number('宽度', 540); await number('高度', 200)
  assert.deepEqual((await state()).model, before.model); await form().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const saved = await state(); assert.equal(saved.timelines[ids.primary].settings.caption, 'Mixed Timeline properties'); assert.equal(saved.timelines[ids.primary].settings.level, 2)
  assert.equal(saved.commands.filter(command => command.id === 'sheet.command.update-slicer-presentation').length, 1)
  assert.equal(saved.commands.filter(command => command.id === 'sheet.command.update-timeline-settings').length, 0)
  assert.deepEqual(retainedDisplay(saved), retainedDisplay(before)); await action('undo'); assert.deepEqual((await state()).model, before.model)
})

test('Timeline display defaults clear seven optional values once, retain name and survive JSON reopen as a no-op draft', async () => {
  await mount({ authoredSettings: true }); await openSettings(); const before = await state()
  await form().getByRole('button', { name: '还原显示默认值', exact: true }).click(); assert.deepEqual((await state()).model, before.model)
  await form().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const saved = await state(), event = saved.commands.find(command => command.id === 'sheet.command.update-timeline-settings')
  assert.deepEqual(saved.timelines[ids.primary].settings, { name: before.timelines[ids.primary].settings.name, level: 2, selectionLevel: 2 })
  assert.deepEqual(event.clearedKeys.sort(), ['caption', 'scrollPosition', 'showHeader', 'showHorizontalScrollbar', 'showSelectionLabel', 'showTimeLevel', 'style'])
  await form().getByRole('button', { name: '取消', exact: true }).click(); await action('reopen'); await openSettings()
  await expect(form().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); assert.deepEqual((await state()).timelines[ids.primary].settings, saved.timelines[ids.primary].settings)
})

test('inherited Timeline caption follows a renamed view before save without altering source labels or the shared view', async () => {
  await mount({ captionInherited: true }); await openSettings(); const before = await state()
  await expect(form().getByRole('checkbox', { name: '使用名称作为标题', exact: true })).toBeChecked()
  await form().getByRole('textbox', { name: '时间线名称', exact: true }).fill('Renamed UI calendar')
  await expect(form().getByRole('textbox', { name: '标题', exact: true })).toHaveValue('Renamed UI calendar')
  await form().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const saved = await state(); assert.equal(saved.timelines[ids.primary].settings.caption, undefined)
  assert.deepEqual(retainedDisplay(saved), retainedDisplay(before)); assert.deepEqual(saved.timelines[ids.shared].settings, before.timelines[ids.shared].settings)
})

test('unknown imported Timeline style stays visible until an explicit clear restores the workbook default', async () => {
  await mount({ unknownStyle: true }); await openSettings(); const before = await state()
  await expect(form().getByRole('button', { name: '选择样式', exact: true })).toContainText('当前样式（Missing imported calendar style）')
  await form().getByRole('button', { name: '选择样式', exact: true }).click()
  await page.getByRole('group', { name: '样式图库', exact: true }).getByRole('button', { name: '文档默认值', exact: true }).click()
  await form().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const saved = await state(); assert.equal(saved.timelines[ids.primary].settings.style, undefined); assert.deepEqual(retainedDisplay(saved), retainedDisplay(before))
})

test.each(['duplicate', 'empty-caption'])('%s Timeline display validation blocks submission and preserves content/history', async kind => {
  await mount(); await openSettings(); const before = await state()
  if (kind === 'duplicate') await form().getByRole('textbox', { name: '时间线名称', exact: true }).fill(` ${before.timelines[ids.shared].settings.name.toUpperCase()} `)
  else await caption('')
  await expect(form().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, 0)
})

test.each([false, true].flatMap(date1904 => [false, true].map(wholeDay => [date1904, wholeDay])))
  ('date1904=%s/wholeDay=%s range retains exact half-open author endpoints and filters all connected reports', async (date1904, wholeDay) => {
    await mount({ date1904 }); await openFilter(); const before = await state(), day = first(before)
    await expect(filter().getByTestId('timeline-inclusive-date-bounds')).toContainText(date1904 ? '1904 日期系统' : '1900 日期系统')
    await editRange(day + .5, day + .75, wholeDay); assert.deepEqual((await state()).model, before.model)
    await filter().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1); await expect(filter()).not.toBeVisible()
    const saved = await state(), range = { startDate: day + .5, endDate: day + .75, ...(wholeDay ? { useWholeDay: true } : {}) }
    assert.deepEqual(saved.timelines[ids.primary].range, range); assert.deepEqual(saved.timelines[ids.shared].range, range)
    assert.deepEqual(saved.timelines[ids.primary].settings, before.timelines[ids.primary].settings); assert.deepEqual(saved.timelines[ids.shared].settings, before.timelines[ids.shared].settings)
    assert.deepEqual(saved.totals, Array(3).fill(wholeDay ? 110 : 80)); assert.equal(saved.commands.filter(command => command.id === 'sheet.command.set-timeline-date-range').length, 1)
    await expect(form().getByTestId('timeline-panel-retained-filter')).toContainText('不含')
    await action('undo'); assert.deepEqual((await state()).model, before.model); await action('redo'); assert.deepEqual((await state()).timelines[ids.primary].range, range)
  })

test.each(conditionOptions)('UI condition %s (%s) retains its one/two operands, uses the shared intent and includes the correct original dates', async (operator, label, total) => {
  await mount(); await openFilter(); const before = await state(), day = first(before), between = operator === 'between' || operator === 'notBetween'
  await editCondition(operator, label, between ? day + .25 : day + .5, day + .75); assert.deepEqual((await state()).model, before.model)
  await filter().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const saved = await state(), condition = { operator, expected: between ? [day + .25, day + .75] : day + .5 }
  assert.deepEqual(saved.timelines[ids.primary].condition, condition); assert.deepEqual(saved.timelines[ids.shared].condition, condition)
  assert.equal(saved.timelines[ids.primary].range, undefined); assert.deepEqual(saved.totals, [total, total, total])
  assert.equal(saved.commands.filter(command => command.id === 'sheet.command.set-timeline-date-condition').length, 1)
  assert.equal(saved.commands.some(command => command.id === 'sheet.command.set-slicer-items'), false)
  await action('undo'); assert.deepEqual((await state()).model, before.model)
})

test.each(['range', 'condition'])('UI clears a retained %s through the canonical undefined range command and undo restores it', async kind => {
  await mount(); const initial = await state(), day = first(initial)
  if (kind === 'range') await action('setRange', ids.primary, { startDate: day + .5, endDate: day + .75 })
  else await action('setCondition', ids.primary, { operator: 'equal', expected: day + .5 })
  await openFilter(); const before = await state(); await option(filter(), '时间线筛选方式', '全部日期')
  await filter().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(before.history.undo + 1)
  const saved = await state(); assert.equal(saved.timelines[ids.primary].range, undefined); assert.equal(saved.timelines[ids.primary].condition, undefined)
  assert.equal(saved.timelines[ids.shared].range, undefined); assert.equal(saved.timelines[ids.shared].condition, undefined)
  assert.deepEqual(saved.commands.filter(command => command.id === 'sheet.command.set-timeline-date-range').at(-1).rangeKeys, [])
  await action('undo'); assert.deepEqual((await state()).model, before.model)
})

test('filling original source calendar days deliberately creates next-day midnight without changing inclusive bound metadata', async () => {
  await mount(); await openFilter(); const before = await state(), day = first(before), bounds = await action('bounds')
  await option(filter(), '时间线筛选方式', '日期范围（结束值不含）')
  await filter().getByRole('button', { name: '填入来源日期的整日范围', exact: true }).click()
  assert.deepEqual(await action('bounds'), bounds); assert.deepEqual((await state()).model, before.model)
  await filter().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  assert.deepEqual((await state()).timelines[ids.primary].range, { startDate: day, endDate: day + 32 })
})

test.each(['noDates', 'hiddenSource'])('%s retains authored filter information while an unreliable source blocks a new range', async kind => {
  await mount(); const initial = await state(), day = first(initial)
  await mount({ [kind]: true, range: { startDate: day + .5, endDate: day + .75 } }); await openFilter(); const before = await state()
  await expect(filter().getByTestId('timeline-retained-date-filter')).toContainText('不含')
  await editRange(day + .25, day + 1); await expect(filter().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  assert.deepEqual((await state()).timelines[ids.primary].range, before.timelines[ids.primary].range); assert.equal((await state()).history.undo, 0)
  await filter().getByRole('button', { name: '取消', exact: true }).click(); assert.deepEqual((await state()).model, before.model)
})

test('Excel 1900-02-29 uses an Ant text bridge that keeps invalid local text without replacing the input or author range', async () => {
  await mount({ compatibilityDay: true }); await openFilter(); const before = await state(), input = filter().getByRole('textbox', { name: '时间线开始日期', exact: true })
  await expect(input).toHaveValue('1900-02-29 06:00:00.000'); await expect(filter().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await input.fill('1900-02-29 invalid'); await expect(input).toHaveValue('1900-02-29 invalid'); await expect(filter().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  assert.deepEqual((await state()).model, before.model)
  await input.fill('1900-02-29 06:00:00.000'); await filter().getByRole('button', { name: '取消', exact: true }).click()
  assert.deepEqual((await state()).timelines[ids.primary].range, { startDate: 60.25, endDate: 60.75 })
})

test('switching Timeline navigation input modes leaves untouched fractional precision intact through another settings edit', async () => {
  await mount({ authoredSettings: true }); await openSettings(); const before = await state(), navigation = before.timelines[ids.primary].settings.scrollPosition
  const field = form().locator('.timeline-date-input').filter({ has: page.getByText('导航日期', { exact: true }) })
  await field.getByRole('button', { name: '精确日期值', exact: true }).click(); await expect(field.getByRole('spinbutton', { name: '时间线导航日期精确日期值', exact: true })).toHaveValue(String(navigation))
  await field.getByRole('button', { name: '日历输入', exact: true }).click(); await caption('Other display edit')
  await form().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const saved = await state(); assert.equal(saved.timelines[ids.primary].settings.scrollPosition, navigation)
  assert.deepEqual(saved.commands.find(command => command.id === 'sheet.command.update-timeline-settings').patchKeys, ['caption'])
})

test('cancelling a date dialog preserves the parent Timeline draft and returns focus to its current Ant trigger', async () => {
  await mount(); await openSettings(); await form().getByRole('textbox', { name: '时间线名称', exact: true }).fill('Pending calendar name')
  const before = await state(), day = first(before), trigger = form().getByRole('button', { name: '日期筛选', exact: true })
  await trigger.click(); await editRange(day + .5, day + .75); await filter().getByRole('button', { name: '取消', exact: true }).click()
  await expect(filter()).not.toBeVisible(); await expect(trigger).toBeFocused()
  await expect(form().getByRole('textbox', { name: '时间线名称', exact: true })).toHaveValue('Pending calendar name')
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, 0)
})

test.each(['settings', 'range', 'condition', 'clear'].flatMap(kind => ['source', 'history', 'user', 'host', 'closed'].map(change => [kind, change])))
  ('pending UI %s rejects changed %s without a late edit or clearing replacement state', async (kind, change) => {
    await mount(); const initial = await state(), day = first(initial)
    if (kind === 'clear') await action('setRange', ids.primary, { startDate: day + .25, endDate: day + 1 })
    if (kind === 'settings') { await openSettings(); await caption('Old pending settings') }
    else { await openFilter(); if (kind === 'range') await editRange(day + .5, day + .75)
      else if (kind === 'condition') await editCondition('equal', '等于', day + .5)
      else await option(filter(), '时间线筛选方式', '全部日期') }
    await action('delay', kind); await (kind === 'settings' ? form() : filter()).getByRole('button', { name: '确定', exact: true }).click(); await waitPending(kind)
    if (change === 'closed') await (kind === 'settings' ? form() : filter()).getByRole('button', { name: '取消', exact: true }).click()
    else await changeOwner(change)
    const before = await state(); await action('release', kind)
    assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  })

test('old bounds result cannot finish or overwrite a newer pending source query', async () => {
  await mount(); await openSettings(); await action('delay', 'bounds'); await form().getByRole('button', { name: '日期筛选', exact: true }).click(); await waitPending('bounds')
  await filter().getByRole('button', { name: '重新读取', exact: true }).click(); await waitPending('bounds', 2)
  await action('release', 'bounds'); await expect(filter().getByText('正在读取来源中的真实日期边界…', { exact: true })).toBeVisible()
  await action('release', 'bounds'); await expect(filter().getByTestId('timeline-inclusive-date-bounds')).toBeVisible(); assert.equal((await state()).history.undo, 0)
})

test('old clear params guard and finally leave a fresh pending clear busy and commit only its captured owner', async () => {
  await mount(); const initial = await state(), day = first(initial); await action('setRange', ids.primary, { startDate: day + .25, endDate: day + 1 }); await openFilter()
  await option(filter(), '时间线筛选方式', '全部日期'); await action('delay', 'clear'); await filter().getByRole('button', { name: '确定', exact: true }).click(); await waitPending('clear')
  await action('clearHistory'); await filter().getByRole('button', { name: '重新读取', exact: true }).click()
  await option(filter(), '时间线筛选方式', '全部日期'); await filter().getByRole('button', { name: '确定', exact: true }).click(); await waitPending('clear', 2)
  await action('release', 'clear'); await expect(filter().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); assert.equal((await state()).history.undo, 0)
  await action('release', 'clear'); await waitHistory(1); await expect(filter()).not.toBeVisible(); assert.equal((await state()).timelines[ids.primary].range, undefined)
})

test.each(['range', 'condition'])('synthetic %s IME composition retains stale Ant input and blocks Enter/Escape until explicit reread', async kind => {
  await mount(); await openFilter(); const before = await state(), day = first(before)
  if (kind === 'range') await editRange(day + .5, day + .75); else await editCondition('equal', '等于', day + .5)
  const input = filter().getByRole('spinbutton', { name: kind === 'range' ? '时间线开始日期精确日期值' : '时间线条件日期精确日期值', exact: true })
  await input.dispatchEvent('compositionstart', { data: '日' }); await action('changeSource'); await input.fill(String(day + .6))
  await input.dispatchEvent('keydown', { key: 'Enter', isComposing: true, keyCode: 229 }); await input.dispatchEvent('keydown', { key: 'Escape', isComposing: true, keyCode: 229 })
  await expect(filter()).toBeVisible(); await expect(input).toHaveValue(String(day + .6)); await expect(filter().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await expect(filter().getByRole('button', { name: '重新读取', exact: true })).toBeDisabled(); await input.dispatchEvent('compositionend', { data: '日期' })
  await expect(input).toHaveValue(String(day + .6)); await filter().getByRole('button', { name: '重新读取', exact: true }).click()
  await expect(filter().getByText('全部日期', { exact: true })).toBeVisible()
  assert.equal((await state()).history.undo, before.history.undo + 1)
})

test.each([[false, 'classic'], [true, 'grid']])('dark=%s/%s uses inherited Ant theme for date controls/popups while retaining the current draft', async (dark, variant) => {
  await mount({ authoredSettings: true, dark, variant }); await openFilter(); const before = await state()
  const originalColor = await filter().getByTestId('timeline-date-filter-form').evaluate(element => getComputedStyle(element).color)
  await action('appearance', !dark, variant); const nextColor = await filter().getByTestId('timeline-date-filter-form').evaluate(element => getComputedStyle(element).color)
  assert.notEqual(nextColor, originalColor); assert.deepEqual((await state()).model, before.model)
  await expect(filter().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await option(filter(), '时间线筛选方式', '日期条件'); await filter().getByRole('combobox', { name: '时间线日期条件', exact: true }).click()
  await expect(page.locator('.ant-select-dropdown:visible')).toBeVisible(); assert.deepEqual((await state()).model, before.model)
})

test('viewable read-only Timeline disables the Ribbon entry and formal API edits without mutating author content', async () => {
  await mount({ readonly: true }); await action('select', ids.primary); const before = await state()
  await expect(toolbar().getByRole('button', { name: '时间线设置', exact: true })).toBeDisabled()
  const refused = await action('updateSettings', ids.primary, { caption: 'Refused read-only caption' })
  assert.equal(refused.accepted, false); assert.deepEqual(refused.model, before.model); assert.equal(refused.history.undo, 0)
  assert.ok(refused.timelines[ids.primary].settings)
})
