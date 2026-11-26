import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// Deferred real Ant/Workbench/Canvas2D/Facade drafts. No browser, server, build or checks have run.
let server, browser, context, page, url
const errors = [], ids = { slicer: 'slicer-a', shared: 'slicer-b', table: 'table-slicer', timeline: 'timeline-a', host: 'reports', other: 'other' }
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No settings fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/slicer-connections-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1100 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.slicerConnectionsFixture?.state().model.slicers.length === 4)
}, 60_000)
afterEach(async () => {
  try { await page?.evaluate(() => window.slicerConnectionsFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const action = (method, ...args) => page.evaluate(({ method, args }) => window.slicerConnectionsFixture[method](...args), { method, args })
const state = () => action('state'), draft = () => page.getByTestId('slicer-settings-form'), dialog = () => page.getByRole('dialog', { name: '排序与无数据项', exact: true })
const toolbar = () => page.getByRole('toolbar', { name: '切片器格式', exact: true })
const info = (value, id = ids.slicer) => value.model.slicers.find(control => control.id === id)
const itemButtons = (id = ids.slicer) => page.locator(`[data-slicer-id="${id}"] [data-slicer-control="item"]`)
function retained(value) {
  return { sheets: value.model.sheets.map(({ id, cells }) => ({ id, cells })), pivots: value.model.pivotTables, charts: value.model.charts, tables: value.model.tables,
    caches: value.model.slicerCaches.map(({ id, source, connections, recordSelection, selectedItems, timelineRange }) => ({ id, source, connections, recordSelection, selectedItems, timelineRange })) }
}
async function mount(options = {}) { await action('mount', { settingsUi: true, ...options }) }
async function openSettings(id = ids.slicer) {
  await action('select', id); await toolbar().getByRole('button', { name: id === ids.timeline ? '时间线设置' : '切片器设置', exact: true }).click()
  await expect(draft()).toBeVisible()
}
async function openCache(id = ids.slicer) { await openSettings(id); await draft().getByRole('button', { name: '排序与无数据项', exact: true }).click(); await expect(dialog()).toBeVisible() }
async function option(name, label) {
  await dialog().getByRole('combobox', { name, exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}
async function number(name, value) { const input = draft().getByRole('spinbutton', { name, exact: true }); await input.fill(String(value)); await input.press('Tab') }
async function caption(value) {
  const inherited = draft().getByRole('checkbox', { name: '使用名称作为标题', exact: true })
  if (await inherited.isChecked()) await inherited.uncheck()
  await draft().getByRole('textbox', { name: '标题', exact: true }).fill(value)
}
async function waitHistory(expected) { await page.waitForFunction(expected => window.slicerConnectionsFixture.state().history.undo === expected, expected) }
async function changeOwner(kind) {
  if (kind === 'source') await action('changeSource')
  if (kind === 'history') await action('clearHistory')
  if (kind === 'user') await action('switchUser', 'peer')
  if (kind === 'host') await action('switchSheet', ids.other)
}

test.each(['classic', 'grid'].flatMap(variant => [ids.slicer, ids.table].map(id => [variant, id])))
  ('%s %s settings save all eight public values once without changing shared filter or drawing', async (variant, id) => {
    await mount({ variant, viewDefaults: true }); await openSettings(id); const before = await state()
    await draft().getByRole('textbox', { name: '切片器名称', exact: true }).fill(`UI ${id}`); await caption('Independent UI caption')
    await number('列数', 2); await number('按钮高度', 30.125); await number('首个显示项', 2)
    await draft().getByRole('checkbox', { name: '显示标题', exact: true }).uncheck()
    await draft().getByRole('checkbox', { name: '锁定位置', exact: true }).check()
    await draft().getByRole('button', { name: '选择样式', exact: true }).click()
    await page.getByRole('group', { name: '样式图库', exact: true }).getByRole('button', { name: '样式 10', exact: true }).click()
    assert.deepEqual((await state()).model, before.model)
    await draft().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(before.history.undo + 1)
    const saved = await state(), settings = saved.settings[id]
    assert.equal(settings.name, `UI ${id}`); assert.equal(settings.caption, 'Independent UI caption'); assert.equal(settings.startItem, 1)
    assert.equal(settings.columnCount, 2); assert.equal(settings.rowHeight, Math.round(30.125 * 9525) / 9525)
    assert.equal(settings.showCaption, false); assert.equal(settings.lockedPosition, true); assert.equal(settings.style, 'SlicerStyleDark2')
    assert.deepEqual(retained(saved), retained(before)); assert.deepEqual(info(saved, id).drawingGeometry, info(before, id).drawingGeometry)
    assert.equal(saved.commands.filter(command => command.id === 'sheet.command.update-slicer-settings').length, 1)
    await expect(itemButtons(id).first()).toHaveText('South')
    await action('undo'); assert.deepEqual((await state()).model, before.model); await action('redo'); assert.deepEqual((await state()).settings[id], settings)
  })

test('inherited caption follows a changed name and can be changed independently without altering source labels', async () => {
  await mount({ captionInherited: true }); await openSettings(); const before = await state()
  await expect(draft().getByRole('checkbox', { name: '使用名称作为标题', exact: true })).toBeChecked()
  await draft().getByRole('textbox', { name: '切片器名称', exact: true }).fill('Renamed UI view')
  await expect(draft().getByRole('textbox', { name: '标题', exact: true })).toHaveValue('Renamed UI view')
  await draft().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const named = await state(); assert.equal(named.settings[ids.slicer].caption, undefined); assert.deepEqual(retained(named), retained(before))
  await caption('Caption override'); await draft().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(2)
  assert.equal((await state()).settings[ids.slicer].name, 'Renamed UI view'); assert.equal((await state()).settings[ids.slicer].caption, 'Caption override')
})

test.each([ids.slicer, ids.table])('%s reset display defaults retains name, clears optional author properties and reopens cleanly', async id => {
  await mount({ authoredSettings: true }); if (id === ids.table) await action('updateSettings', id, { caption: 'Table custom', columnCount: 2, rowHeight: 30,
    startItem: 1, showCaption: false, lockedPosition: false, style: 'SlicerStyleDark2' })
  await openSettings(id); const before = await state(); await draft().getByRole('button', { name: '还原显示默认值', exact: true }).click()
  assert.deepEqual((await state()).model, before.model); await draft().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(before.history.undo + 1)
  const saved = await state(), event = saved.commands.filter(command => command.id === 'sheet.command.update-slicer-settings').at(-1)
  assert.deepEqual(saved.settings[id], { name: before.settings[id].name, rowHeight: 24 })
  assert.deepEqual(event.clearedKeys.sort(), ['caption', 'columnCount', 'lockedPosition', 'showCaption', 'startItem', 'style'])
  await draft().getByRole('button', { name: '取消', exact: true }).click(); await action('reopen'); await openSettings(id)
  await expect(draft().getByRole('checkbox', { name: '使用名称作为标题', exact: true })).toBeChecked()
  await expect(draft().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); assert.deepEqual((await state()).settings[id], saved.settings[id])
})

test.each([ids.slicer, ids.table])('%s mixed display and geometry draft commits once and undo restores authored affine data', async id => {
  await mount(); await openSettings(id); const before = await state(); await caption('Mixed geometry caption')
  await draft().getByText('大小与属性', { exact: true }).click(); await number('宽度', 300); await number('高度', 220)
  assert.deepEqual((await state()).model, before.model); await draft().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const saved = await state(); assert.equal(saved.settings[id].caption, 'Mixed geometry caption')
  assert.equal(saved.commands.filter(command => command.id === 'sheet.command.update-slicer-presentation').length, 1)
  assert.equal(saved.commands.filter(command => command.id === 'sheet.command.update-slicer-settings').length, 0)
  assert.deepEqual(retained(saved), retained(before)); await action('undo'); assert.deepEqual((await state()).model, before.model)
})

test.each(['duplicate', 'empty-caption'])('%s display validation disables submission before any history or source change', async kind => {
  await mount(); await openSettings(); const before = await state()
  if (kind === 'duplicate') await draft().getByRole('textbox', { name: '切片器名称', exact: true }).fill(` ${(before.settings[ids.shared].name).toUpperCase()} `)
  else await caption('')
  await expect(draft().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); assert.deepEqual((await state()).model, before.model)
})

test('row height keeps EMU precision while the one-based first displayed item maps to the raw zero-based offset', async () => {
  await mount({ viewDefaults: true }); await openSettings(); await number('按钮高度', 30.123456); await number('首个显示项', 2)
  await draft().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const current = await state(); assert.equal(current.settings[ids.slicer].rowHeight, Math.round(30.123456 * 9525) / 9525); assert.equal(current.settings[ids.slicer].startItem, 1)
  await expect(itemButtons().first()).toHaveText('South'); await expect(itemButtons(ids.shared).first()).toHaveText('North')
})

test.each([ids.slicer, ids.table])('%s cache dialog uses the real shared setter and sorted Ant items without changing retained filters', async id => {
  await mount(); await openCache(id); const before = await state(); await option('切片器成员排序', '降序'); await option('切片器自定义列表排序', '关闭')
  assert.deepEqual((await state()).model, before.model); await dialog().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const current = await state(); assert.deepEqual(current.cacheOptions[id], { sortOrder: 'descending', customListSort: false })
  assert.deepEqual(retained(current), retained(before)); assert.equal(current.commands.filter(command => command.id === 'sheet.command.patch-slicer-cache-options').length, 1)
  await expect(itemButtons(id).first()).toHaveText('South')
  if (id === ids.slicer) { assert.deepEqual(current.cacheOptions[ids.shared], current.cacheOptions[id]); await expect(itemButtons(ids.shared).first()).toHaveText('South') }
  else assert.deepEqual(current.cacheOptions[ids.slicer], {})
  await action('undo'); assert.deepEqual((await state()).model, before.model)
})

test.each(['none', 'showItemsWithDataAtTop', 'showItemsWithNoData'].flatMap(cross => [false, true].map(hide => [cross, hide])))
  ('March context %s/hide=%s drives typed members and rendered order while retaining selected North', async (cross, hide) => {
    await mount({ marchContext: true, cacheOptions: { hideItemsWithNoData: hide } }); await action('setItems', ids.slicer, ['North']); await openCache()
    const before = await state(), label = cross === 'none' ? '不标记无数据项' : cross === 'showItemsWithDataAtTop' ? '有数据项置顶' : '显示无数据项'
    await option('切片器无数据项显示', label)
    if (cross === 'none') await expect(dialog().getByRole('combobox', { name: '切片器隐藏无数据项', exact: true })).toBeDisabled()
    await dialog().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(before.history.undo + 1)
    const current = await state(), members = await action('members', ids.slicer), names = cross === 'none' ? ['North', 'South'] : hide ? ['South']
      : cross === 'showItemsWithDataAtTop' ? ['South', 'North'] : ['North', 'South']
    assert.deepEqual(members.items.map(item => item.value.value), names); assert.deepEqual(current.displayMembers[ids.slicer].map(item => item.label), names)
    assert.deepEqual(await itemButtons().allTextContents(), names); assert.deepEqual(info(current).selectedItems, ['North']); assert.deepEqual(retained(current), retained(before))
  })

test('cache explicit defaults and reset retain optional presence and survive snapshot reopen', async () => {
  await mount({ cacheOptions: { sortOrder: 'ascending', customListSort: false, crossFilter: 'showItemsWithNoData', hideItemsWithNoData: false } })
  await openCache(); await expect(dialog().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await dialog().getByRole('button', { name: '文档默认值', exact: true }).click(); await dialog().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  const current = await state(); assert.deepEqual(current.cacheOptions[ids.slicer], {})
  const command = current.commands.find(command => command.id === 'sheet.command.patch-slicer-cache-options')
  assert.deepEqual(command.clearedKeys.sort(), ['crossFilter', 'customListSort', 'hideItemsWithNoData', 'sortOrder'])
  await draft().getByRole('button', { name: '取消', exact: true }).click(); await action('reopen'); await openCache()
  await expect(dialog().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); assert.deepEqual((await state()).cacheOptions[ids.shared], {})
})

test('cancelling cache options preserves the parent display draft and restores its current Ant trigger focus', async () => {
  await mount(); await openSettings(); await draft().getByRole('textbox', { name: '切片器名称', exact: true }).fill('Pending name')
  const before = await state(), trigger = draft().getByRole('button', { name: '排序与无数据项', exact: true }); await trigger.click()
  await option('切片器成员排序', '降序'); await dialog().getByRole('button', { name: '取消', exact: true }).click()
  await expect(dialog()).not.toBeVisible(); await expect(trigger).toBeFocused(); await expect(draft().getByRole('textbox', { name: '切片器名称', exact: true })).toHaveValue('Pending name')
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, 0)
})

test('saving shared options keeps an unsaved parent display draft stale until an explicit reread', async () => {
  await mount(); await openSettings(); await draft().getByRole('textbox', { name: '切片器名称', exact: true }).fill('Unsaved name')
  await draft().getByRole('button', { name: '排序与无数据项', exact: true }).click(); await option('切片器成员排序', '降序')
  await dialog().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1)
  await expect(draft().getByRole('textbox', { name: '切片器名称', exact: true })).toHaveValue('Unsaved name'); await expect(draft().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await draft().getByRole('button', { name: '重新读取', exact: true }).click()
  await expect(draft().getByRole('textbox', { name: '切片器名称', exact: true })).toHaveValue((await state()).settings[ids.slicer].name)
  assert.equal((await state()).history.undo, 1)
})

test.each([ids.slicer, ids.table])('%s viewable read-only host shows disabled settings and sorting entry', async id => {
  await mount({ readonly: true }); await openSettings(id); const before = await state()
  await expect(draft().getByRole('textbox', { name: '切片器名称', exact: true })).toBeDisabled()
  await expect(draft().getByRole('button', { name: '排序与无数据项', exact: true })).toBeDisabled()
  await expect(draft().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); assert.deepEqual((await state()).model, before.model)
})

test.each(['settings', 'cacheOptions'].flatMap(kind => ['source', 'history', 'user', 'host'].map(change => [kind, change])))
  ('pending %s rejects changed %s owner at actual dispatch without partial settings/cache/history', async (kind, change) => {
    await mount(); if (kind === 'settings') { await openSettings(); await caption('Late old caption') } else { await openCache(); await option('切片器成员排序', '降序') }
    await action('delayMutation', kind); await (kind === 'settings' ? draft() : dialog()).getByRole('button', { name: '确定', exact: true }).click()
    await page.waitForFunction(kind => window.slicerConnectionsFixture.pendingMutations(kind) === 1, kind)
    await changeOwner(change); const before = await state(); await action('releaseMutation', kind)
    assert.deepEqual((await state()).model, before.model); assert.deepEqual((await state()).history, before.history)
  })

test('an old cache finally cannot clear a newer busy command after a history reread with the same revision', async () => {
  await mount(); await openCache(); await option('切片器成员排序', '降序'); await action('delayMutation', 'cacheOptions')
  await dialog().getByRole('button', { name: '确定', exact: true }).click(); await page.waitForFunction(() => window.slicerConnectionsFixture.pendingMutations('cacheOptions') === 1)
  await action('clearHistory'); await dialog().getByRole('button', { name: '重新读取', exact: true }).click(); await option('切片器成员排序', '降序')
  await dialog().getByRole('button', { name: '确定', exact: true }).click(); await page.waitForFunction(() => window.slicerConnectionsFixture.pendingMutations('cacheOptions') === 2)
  await action('releaseMutation', 'cacheOptions', 0); await expect(dialog().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); assert.equal((await state()).history.undo, 0)
  await action('releaseMutation', 'cacheOptions', 0); await waitHistory(1); await expect(dialog()).not.toBeVisible()
})

test.each(['classic', 'grid'])('%s settings and Ant sorting popup inherit theme tokens without changing content', async variant => {
  await mount({ variant }); await openCache(); const before = await state(), light = await dialog().locator('.slicer-cache-options').evaluate(node => getComputedStyle(node).color)
  await action('appearance', true, variant); assert.notEqual(await dialog().locator('.slicer-cache-options').evaluate(node => getComputedStyle(node).color), light)
  await option('切片器成员排序', '降序'); assert.deepEqual((await state()).model, before.model)
})

test.each(['切片器名称', '标题'])('synthetic Ant composition in %s pins text and blocks Enter/Escape until owner-safe completion', async name => {
  await mount(); await openSettings(); const input = draft().getByRole('textbox', { name, exact: true }); await input.focus()
  await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })))
  await input.evaluate(node => { node.value = '中'; node.dispatchEvent(new InputEvent('input', { bubbles: true, data: '中', isComposing: true, inputType: 'insertCompositionText' })) })
  await expect(draft().getByRole('button', { name: '确定', exact: true })).toBeDisabled(); await input.press('Escape'); await expect(draft()).toBeVisible()
  await action('changeSource'); await expect(input).toHaveValue('中'); await expect(input).toBeEnabled()
  await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中' })))
  await expect(input).toBeDisabled(); await expect(draft().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
})

test('Timeline keeps its independent presentation fields and does not expose Slicer settings or sorting controls', async () => {
  await mount(); await openSettings(ids.timeline); const before = await state()
  await expect(draft().getByRole('combobox', { name: '时间级别', exact: true })).toBeVisible()
  await expect(draft().getByRole('textbox', { name: '切片器名称', exact: true })).toHaveCount(0)
  await expect(draft().getByRole('button', { name: '排序与无数据项', exact: true })).toHaveCount(0); assert.deepEqual((await state()).model, before.model)
})

test.each(['custom', 'unknown'])('%s style gallery uses document resources and retains imported overrides until an explicit default choice', async kind => {
  await mount({ customStyles: true, customOverride: kind === 'custom', unknownStyle: kind === 'unknown' }); await openSettings(); const before = await state()
  await draft().getByRole('button', { name: '选择样式', exact: true }).click()
  await expect(page.getByRole('group', { name: '样式图库', exact: true }).getByRole('button', { name: 'Office fixture style', exact: true })).toBeVisible()
  if (kind === 'unknown') await expect(page.getByRole('group', { name: '样式图库', exact: true }).getByRole('button', { name: '当前样式（Missing imported style）', exact: true })).toHaveAttribute('aria-pressed', 'true')
  assert.deepEqual((await state()).model, before.model)
  await page.getByRole('group', { name: '样式图库', exact: true }).getByRole('button', { name: '文档默认值', exact: true }).click()
  await draft().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(1); assert.equal((await state()).settings[ids.slicer].style, undefined)
})

test.each(['settings', 'cacheOptions'])('disposing the workbench removes pending %s UI and rejects its late owner', async kind => {
  await mount(); if (kind === 'settings') { await openSettings(); await caption('Late disposed caption') } else { await openCache(); await option('切片器成员排序', '降序') }
  await action('delayMutation', kind); await (kind === 'settings' ? draft() : dialog()).getByRole('button', { name: '确定', exact: true }).click()
  await page.waitForFunction(kind => window.slicerConnectionsFixture.pendingMutations(kind) === 1, kind)
  await action('dispose'); await expect(page.locator('#app')).toBeEmpty(); await expect(dialog()).not.toBeVisible(); await expect(draft()).toHaveCount(0)
})
