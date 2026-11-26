import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// Deferred real Ant/Workbench/Canvas2D/Facade/history drafts. No hook, server or browser has been started.
// Windows IME, WebGPU, independent Office files and large-sample performance are separate acceptance work.
let server, browser, context, page, url
const errors = []
const ids = { unit: 'slicer-connections-book', host: 'reports', other: 'other', primary: 'primary-report', secondary: 'secondary-report',
  incompatible: 'other-cache-report', owned: 'owned-chart', linked: 'linked-chart', slicer: 'slicer-a', shared: 'slicer-b', timeline: 'timeline-a', table: 'table-slicer' }
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/slicer-connections-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1000 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.slicerConnectionsFixture?.state().model.slicers.length === 4)
}, 60_000)
afterEach(async () => {
  try { await page?.evaluate(() => window.slicerConnectionsFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.slicerConnectionsFixture.state())
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const formatToolbar = () => page.getByRole('toolbar', { name: '切片器格式', exact: true })
const dialog = () => page.getByRole('dialog', { name: '报表连接', exact: true })
const settings = () => page.locator('.slicer-settings')
const control = (value, id = ids.slicer) => value.model.slicers.find(control => control.id === id)
const reportSelection = (value, id) => (value.model.pivotTables.find(pivot => pivot.id === id) ?? value.model.charts.find(chart => chart.id === id)?.ownedPivot)
  ?.fields.find(field => field.sourceFieldId === '0')?.selectedItems
const primary = () => dialog().getByRole('checkbox', { name: /primary-report/ })
const secondary = () => dialog().getByRole('checkbox', { name: /secondary-report/ })
const owned = () => dialog().getByRole('checkbox', { name: /Owned chart/ })
async function select(id = ids.slicer) {
  await page.evaluate(id => window.slicerConnectionsFixture.select(id), id)
  await expect(page.getByRole('tab', { name: '切片器格式', exact: true })).toHaveAttribute('aria-selected', 'true')
}
async function openConnections(id = ids.slicer, wait = true) {
  await select(id); await formatToolbar().getByRole('button', { name: '报表连接', exact: true }).click()
  await expect(dialog()).toBeVisible(); if (wait) await expect(dialog().getByRole('checkbox')).toHaveCount(3)
}
async function openSettings(id = ids.slicer) {
  await select(id); await formatToolbar().getByRole('button', { name: id === ids.timeline ? '时间线设置' : '切片器设置', exact: true }).click()
  await expect(settings()).toBeVisible()
}
async function waitHistory(expected) { await page.waitForFunction(expected => window.slicerConnectionsFixture.state().history.undo === expected, expected) }

test('the Ant connection dialog lists compatible PivotTables and one owned chart with real checked initial state', async () => {
  const before = await state(); await openConnections()
  await expect(primary()).toBeChecked(); await expect(secondary()).not.toBeChecked(); await expect(owned()).not.toBeChecked()
  await expect(dialog().getByRole('checkbox', { name: /other-cache-report|Linked chart/ })).toHaveCount(0)
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('checked reports publish one command, filter all targets, share view connections and retain undo/save/reopen behavior', async () => {
  await page.evaluate(id => window.slicerConnectionsFixture.setItems(id, ['North']), ids.slicer)
  await page.evaluate(() => window.slicerConnectionsFixture.markSaved())
  const before = await state(); await openConnections(); await secondary().check(); await owned().check()
  await dialog().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(before.history.undo + 1)
  await expect(dialog()).not.toBeVisible(); await expect(grid()).toBeFocused()
  const after = await state()
  assert.deepEqual(after.connections[ids.slicer], after.connections[ids.shared]); assert.equal(after.connections[ids.slicer].pivotTables.length, 2)
  assert.equal(after.connections[ids.slicer].ownedPivotCharts.length, 1)
  for (const id of [ids.primary, ids.secondary, ids.owned]) assert.deepEqual(reportSelection(after, id), ['North'])
  assert.equal(after.commands.filter(command => command.id === 'sheet.command.set-slicer-connections').length, 1); assert.equal(after.history.dirty, true)
  await page.evaluate(() => window.slicerConnectionsFixture.undo()); const undone = await state()
  assert.deepEqual(undone.connections, before.connections); assert.equal(undone.history.dirty, false)
  await page.evaluate(() => window.slicerConnectionsFixture.redo()); assert.deepEqual((await state()).connections, after.connections)
  await page.evaluate(() => window.slicerConnectionsFixture.reopen()); const reopened = await state()
  assert.deepEqual(reopened.connections, after.connections); assert.deepEqual(control(reopened, ids.shared).selectedItems, ['North'])
  assert.deepEqual(control(reopened).drawingGeometry, control(before).drawingGeometry)
})

test('unchecking every report disconnects the shared cache and a later checked report receives retained values', async () => {
  await page.evaluate(id => window.slicerConnectionsFixture.setItems(id, ['North']), ids.slicer)
  const before = await state(); await openConnections(); await primary().uncheck()
  await dialog().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(before.history.undo + 1)
  const disconnected = await state()
  assert.deepEqual(disconnected.connections[ids.slicer], { pivotTables: [], ownedPivotCharts: [] })
  assert.deepEqual(control(disconnected, ids.shared).selectedItems, ['North']); assert.equal(control(disconnected).cacheId, control(before).cacheId)
  await openConnections(); await secondary().check(); await dialog().getByRole('button', { name: '确定', exact: true }).click()
  await waitHistory(before.history.undo + 2); assert.deepEqual(reportSelection(await state(), ids.secondary), ['North'])
})

test('Table controls expose settings and quick styles while hiding report connections in the Ribbon and floating tools', async () => {
  await select(ids.table)
  await expect(formatToolbar().getByRole('button', { name: '快速样式', exact: true })).toBeVisible()
  await expect(formatToolbar().getByRole('button', { name: '切片器设置', exact: true })).toBeVisible()
  await expect(formatToolbar().getByRole('button', { name: '报表连接', exact: true })).toHaveCount(0)
  await expect(page.getByRole('toolbar', { name: '绘图组合工具', exact: true }).getByRole('button', { name: '报表连接', exact: true })).toHaveCount(0)
  await openSettings(ids.table); await expect(settings().getByRole('button', { name: '报表连接', exact: true })).toHaveCount(0)
})

for (const variant of ['classic', 'grid']) test(`${variant} retains the three applicable roots and the upstream style/options layout order`, async () => {
  await page.evaluate(variant => window.slicerConnectionsFixture.appearance(false, variant), variant); await select()
  const toolbar = formatToolbar(), quick = toolbar.getByRole('button', { name: '快速样式', exact: true }), setting = toolbar.getByRole('button', { name: '切片器设置', exact: true }), connect = toolbar.getByRole('button', { name: '报表连接', exact: true })
  await expect(quick).toBeVisible(); await expect(setting).toBeVisible(); await expect(connect).toBeVisible()
  await expect(toolbar.getByRole('button', { name: '时间线设置', exact: true })).toHaveCount(0)
  const q = await quick.boundingBox(), s = await setting.boundingBox(), c = await connect.boundingBox()
  assert.ok(q && s && c); assert.ok(q.x < s.x)
  if (variant === 'grid') { assert.ok(s.y < c.y); assert.ok(q.height > s.height) }
  else { assert.ok(s.x < c.x); assert.ok(Math.abs(s.y - c.y) < 2) }
})

test('selecting a timeline switches the settings root and ending the context restores the previous regular tab', async () => {
  const previous = await page.locator('[role="tab"][aria-selected="true"]').getAttribute('data-tab-id')
  await select(ids.timeline)
  await expect(formatToolbar().getByRole('button', { name: '时间线设置', exact: true })).toBeVisible()
  await expect(formatToolbar().getByRole('button', { name: '切片器设置', exact: true })).toHaveCount(0)
  await page.evaluate(() => window.slicerConnectionsFixture.clearSelection())
  await expect(page.getByRole('tab', { name: '切片器格式', exact: true })).toHaveCount(0)
  await expect(page.locator(`[role="tab"][data-tab-id="${previous}"]`)).toHaveAttribute('aria-selected', 'true')
})

test('quick style tiles commit the real presentation command without changing affine geometry, selection or cache links', async () => {
  await select(); const before = await state()
  await formatToolbar().getByRole('button', { name: '快速样式', exact: true }).click()
  const gallery = page.getByRole('group', { name: '切片器快速样式', exact: true }); await expect(gallery).toBeVisible()
  await gallery.getByRole('button', { name: '样式 2', exact: true }).click(); await waitHistory(before.history.undo + 1)
  const after = await state(); assert.equal(control(after).style, 'SlicerStyleLight2')
  assert.deepEqual(control(after).drawingGeometry, control(before).drawingGeometry); assert.deepEqual(after.connections, before.connections)
  assert.deepEqual(after.selection, before.selection); assert.equal(after.commands.filter(command => command.id === 'sheet.command.update-slicer-presentation').length, 1)
  await page.evaluate(() => window.slicerConnectionsFixture.undo()); assert.equal(control(await state()).style, control(before).style)
})

test('a changed settings draft keeps its text after external edits, requires reread and never overwrites the newer affine author', async () => {
  await openSettings(); const before = await state(), caption = settings().getByRole('textbox', { name: '标题', exact: true })
  await caption.fill('Unsaved draft'); await page.evaluate(() => window.slicerConnectionsFixture.changeCaption())
  await expect(caption).toHaveValue('Unsaved draft'); await expect(settings().getByRole('alert')).toContainText('重新读取')
  await expect(settings().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await settings().getByRole('button', { name: '重新读取', exact: true }).click(); await expect(caption).toHaveValue('Changed externally')
  const baseline = (await state()).history.undo
  await caption.fill('Accepted draft'); await settings().getByRole('button', { name: '确定', exact: true }).click(); await waitHistory(baseline + 1)
  const after = await state(); assert.equal(control(after).caption, 'Accepted draft'); assert.deepEqual(control(after).drawingGeometry, control(before).drawingGeometry)
})

test('Escape closes nested connections and restores the Ant settings trigger, then closes the drawer to the Canvas', async () => {
  await openSettings(); const button = settings().getByRole('button', { name: '报表连接', exact: true }), before = await state()
  await button.click(); await expect(dialog().getByRole('checkbox')).toHaveCount(3)
  await page.keyboard.press('Escape'); await expect(dialog()).not.toBeVisible(); await expect(button).toBeFocused()
  await settings().getByRole('button', { name: '取消', exact: true }).click(); await expect(settings()).not.toBeVisible(); await expect(grid()).toBeFocused()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('Escape from a directly opened connection dialog returns Canvas focus without a content mutation', async () => {
  const before = await state(); await openConnections(); await secondary().check(); await page.keyboard.press('Escape')
  await expect(dialog()).not.toBeVisible(); await expect(grid()).toBeFocused()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('a delayed old panel cannot replace a reopened timeline panel or keep its new controls busy', async () => {
  await page.evaluate(() => window.slicerConnectionsFixture.delayCandidates()); await openConnections(ids.slicer, false)
  await page.waitForFunction(() => window.slicerConnectionsFixture.pendingCandidates() === 1)
  await dialog().getByRole('button', { name: '取消', exact: true }).click(); await expect(dialog()).not.toBeVisible()
  await openConnections(ids.timeline, false); await page.waitForFunction(() => window.slicerConnectionsFixture.pendingCandidates() === 2)
  const before = await state(); await page.evaluate(() => window.slicerConnectionsFixture.releaseCandidates())
  await expect(dialog().getByRole('checkbox')).toHaveCount(3); await expect(dialog().locator('.slicer-muted').first()).toContainText('Date')
  await expect(dialog().getByRole('button', { name: '重新读取', exact: true })).toBeEnabled()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('candidate results from before an external edit remain unavailable until the user rereads the real source', async () => {
  await page.evaluate(() => window.slicerConnectionsFixture.delayCandidates()); await openConnections(ids.slicer, false)
  await page.waitForFunction(() => window.slicerConnectionsFixture.pendingCandidates() === 1)
  await page.evaluate(() => window.slicerConnectionsFixture.changeCaption()); const before = await state()
  await page.evaluate(() => window.slicerConnectionsFixture.releaseCandidates())
  await expect(dialog().getByRole('checkbox')).toHaveCount(0); await expect(dialog().getByRole('alert')).toContainText('重新读取')
  await dialog().getByRole('button', { name: '重新读取', exact: true }).click(); await expect(dialog().getByRole('checkbox')).toHaveCount(3)
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('an old connection completion cannot clear the busy state of a newer pending settings operation', async () => {
  await page.evaluate(() => window.slicerConnectionsFixture.delayMutation('connections')); await openConnections()
  await secondary().check(); await dialog().getByRole('button', { name: '确定', exact: true }).click()
  await page.waitForFunction(() => window.slicerConnectionsFixture.pendingMutations('connections') === 1)
  await dialog().getByRole('button', { name: '取消', exact: true }).click(); await openSettings()
  await page.evaluate(() => window.slicerConnectionsFixture.delayMutation('settings'))
  const caption = settings().getByRole('textbox', { name: '标题', exact: true }), save = settings().getByRole('button', { name: '确定', exact: true })
  const before = await state(); await caption.fill('New pending draft'); await save.click()
  await page.waitForFunction(() => window.slicerConnectionsFixture.pendingMutations('settings') === 1)
  await page.evaluate(() => window.slicerConnectionsFixture.releaseMutation('connections')); await waitHistory(before.history.undo + 1)
  await expect(save).toHaveClass(/ant-btn-loading/); await expect(caption).toHaveValue('New pending draft')
  await page.evaluate(() => window.slicerConnectionsFixture.releaseMutation('settings'))
  await expect(save).not.toHaveClass(/ant-btn-loading/)
  assert.equal(control(await state()).caption, control(before).caption); assert.equal((await state()).history.undo, before.history.undo + 1)
})

test('replacing the Studio while candidates are delayed cannot apply old options or return focus into the new model', async () => {
  await page.evaluate(() => window.slicerConnectionsFixture.delayCandidates()); await openConnections(ids.slicer, false)
  await page.waitForFunction(() => window.slicerConnectionsFixture.pendingCandidates() === 1)
  await page.evaluate(() => window.slicerConnectionsFixture.mount({ id: 'replacement-connections-book' }))
  const current = await state(); assert.equal(current.model.id, 'replacement-connections-book')
  await expect(dialog()).not.toBeVisible(); assert.equal(current.history.undo, 0)
  assert.equal(control(current).unitId, current.model.id); assert.equal(current.commands.filter(command => command.id === 'sheet.command.set-slicer-connections').length, 0)
})

test('history replacement invalidates a checked connection draft even when the content revision is unchanged', async () => {
  await openConnections(); await secondary().check(); const before = await state()
  await page.evaluate(() => window.slicerConnectionsFixture.clearHistory())
  await expect(dialog().getByRole('alert')).toContainText('过期'); await expect(dialog().getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  assert.equal((await state()).revision, before.revision); assert.deepEqual((await state()).connections, before.connections)
  await dialog().getByRole('button', { name: '重新读取', exact: true }).click(); await expect(secondary()).not.toBeChecked()
})

test('the open connection modal and quick-style tiles follow ConfigProvider light/dark tokens across layout switches', async () => {
  await openConnections(); const before = await state()
  const light = await dialog().locator('.slicer-connections').evaluate(node => ({ text: getComputedStyle(node).color, surface: getComputedStyle(node.closest('.ant-modal-content')).backgroundColor }))
  await page.evaluate(() => window.slicerConnectionsFixture.appearance(true, 'grid'))
  const dark = await dialog().locator('.slicer-connections').evaluate(node => ({ text: getComputedStyle(node).color, surface: getComputedStyle(node.closest('.ant-modal-content')).backgroundColor }))
  assert.notEqual(dark.text, light.text); assert.notEqual(dark.surface, light.surface)
  await dialog().getByRole('button', { name: '取消', exact: true }).click()
  await formatToolbar().getByRole('button', { name: '快速样式', exact: true }).click()
  const tile = page.getByRole('group', { name: '切片器快速样式', exact: true }).getByRole('button', { name: '样式 1', exact: true })
  const darkTile = await tile.evaluate(node => getComputedStyle(node).backgroundColor)
  await page.evaluate(() => window.slicerConnectionsFixture.appearance(false, 'grid'))
  assert.notEqual(await tile.evaluate(node => getComputedStyle(node).backgroundColor), darkTile)
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('read-only contextual entries are disabled and the public connection command refuses changes without touching history', async () => {
  await page.evaluate(() => window.slicerConnectionsFixture.mount({ readonly: true, currentUserId: 'peer' })); await select()
  for (const name of ['快速样式', '切片器设置', '报表连接']) await expect(formatToolbar().getByRole('button', { name, exact: true })).toBeDisabled()
  const before = await state(), result = await page.evaluate(({ slicer, secondary, other }) => window.slicerConnectionsFixture.setConnections(slicer,
    { pivotTables: [{ subUnitId: other, pivotTableId: secondary }], ownedPivotCharts: [] }), { slicer: ids.slicer, secondary: ids.secondary, other: ids.other })
  assert.equal(result.result, false); assert.deepEqual(result.model, before.model); assert.equal(result.history.undo, before.history.undo)
})
