import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// Deferred real Workbench/Ant/Canvas2D/history/save-reopen drafts. No server or browser has been started.
// Synthetic composition below does not accept Windows native IME, WebGPU, Office files or performance.
let server, browser, context, page, url
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No Sparkline fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/sheet-sparkline-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1100 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.sparklineFixture?.state().groups.length === 2)
}, 60_000)
afterEach(async () => {
  try { await page?.evaluate(() => window.sparklineFixture?.dispose()); assert.deepEqual(errors, []) } finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const action = (method, ...args) => page.evaluate(({ method, args }) => window.sparklineFixture[method](...args), { method, args })
const state = () => action('state'), form = () => page.getByTestId('sparkline-panel')
const value = (snapshot, address, sheetId = 'data') => {
  const match = /^([A-Z]+)(\d+)$/.exec(address); let column = 0
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64
  return snapshot.model.sheets.find(sheet => sheet.id === sheetId).cells[`${Number(match[2]) - 1}:${column - 1}`]?.sparkline ?? null
}
async function waitHistory(expected) { await page.waitForFunction(expected => window.sparklineFixture.state().history.undo === expected, expected) }
async function option(label, text) {
  await form().getByRole('combobox', { name: label, exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText(text, { exact: true }).click()
}
async function number(label, value) { const input = form().getByRole('spinbutton', { name: label, exact: true }); await input.fill(String(value)); await input.press('Tab') }
async function fields(source, target) {
  await form().getByRole('textbox', { name: '迷你图源数据区域', exact: true }).fill(source)
  await form().getByRole('textbox', { name: '迷你图放置区域', exact: true }).fill(target)
}
async function open(ranges = ['G1']) { await action('openPanel', ranges); await expect(form()).toBeVisible() }
async function apply() { await form().getByRole('button', { name: '应用', exact: true }).click() }
async function advanced() { await form().getByRole('tab', { name: '高级设置', exact: true }).click() }
const typeNames = { line: '折线图', column: '柱形图', winloss: '胜负图', pie: '饼图' }

test.each(['classic', 'grid'].flatMap(variant => Object.keys(typeNames).map(type => [variant, type])))
  ('%s Ribbon inserts real %s sparklines from multiple source areas and changes actual Canvas pixels', async (variant, type) => {
    await action('mount', { withGroups: false, variant }); await action('select', ['I1', 'I3'])
    const before = await state(), pixels = await action('pixels', 'I1')
    await page.getByRole('tab', { name: '插入', exact: true }).click()
    await page.getByRole('toolbar', { name: '插入', exact: true }).getByRole('button', { name: '迷你图', exact: true }).click()
    await expect(form()).toBeVisible(); await option('迷你图类型', typeNames[type])
    await fields("B1:E1; 'Data, O''Neil'!B2:E2", 'I1; I3')
    await expect(form().getByRole('textbox', { name: '迷你图源数据区域', exact: true })).toHaveClass(/ant-input/)
    assert.deepEqual((await state()).model, before.model)
    await apply(); await waitHistory(before.history.undo + 1); await expect(form()).toHaveCount(0)
    const saved = await state(), first = value(saved, 'I1'), second = value(saved, 'I3')
    assert.equal(first.type, type); assert.equal(second.type, type); assert.equal(first.groupId, second.groupId)
    assert.equal(first.sourceRange, 'B1:E1'); assert.equal(second.sourceRange, "'Data, O''Neil'!B2:E2")
    assert.equal(saved.commands.filter(command => command.id === 'sheet.command.set-sparklines').length, 1)
    assert.notEqual((await action('pixels', 'I1')).hash, pixels.hash)
    await action('undo'); assert.deepEqual((await state()).model, before.model)
    await action('redo'); await action('reopen'); assert.deepEqual(value(await state(), 'I3'), second)
  })

test('the UI rectangle placement chooses source rows for a square even with horizontal targets', async () => {
  await action('mount', { withGroups: false }); await open(['I1:J1']); await fields('B1:C2', 'I1:J1')
  await apply(); await waitHistory(1)
  const saved = await state(); assert.equal(value(saved, 'I1').sourceRange, 'B1:C1'); assert.equal(value(saved, 'J1').sourceRange, 'B2:C2')
})

test.each(['cell', 'row', 'column'])('the %s Canvas context menu opens real full-group editing and removal entries', async kind => {
  await action('select', ['G1']); const point = await action('point', 'G1', kind)
  await page.mouse.click(point.x, point.y, { button: 'right' })
  await page.getByRole('menuitem', { name: '迷你图', exact: true }).last().hover()
  await expect(page.getByRole('menuitem', { name: '删除迷你图', exact: true }).last()).toBeVisible()
  await page.getByRole('menuitem', { name: '编辑迷你图', exact: true }).last().click()
  await expect(form()).toBeVisible(); await expect(form().getByRole('combobox', { name: '当前迷你图单元格', exact: true })).toBeVisible()
  assert.equal((await state()).history.undo, 0)
})

test('basic Ant settings apply presets, six markers, line width and gradient to the entire selected group', async () => {
  await open(['G2']); const before = await state(), south = value(before, 'G4')
  await option('迷你图主题', '清新'); await option('迷你图线宽', '4 px')
  await form().getByRole('checkbox', { name: '迷你图渐变', exact: true }).check()
  for (const name of ['高值点', '首值点', '低值点', '末值点', '负值点', '所有标记']) await form().getByRole('checkbox', { name: `迷你图${name}`, exact: true }).check()
  assert.deepEqual((await state()).model, before.model); await apply(); await waitHistory(1)
  const saved = await state()
  for (const address of ['G1', 'G2']) {
    const current = value(saved, address); assert.equal(current.themeType, 'fresh'); assert.equal(current.color, '#2ebb7a')
    assert.equal(current.lineWidth, 4); assert.equal(current.showGradient, true)
    assert.ok(Object.values(current.points).every(point => point.visible && point.color === '#f8be15'))
  }
  assert.deepEqual(value(saved, 'G4'), south); assert.ok(saved.scene.commands.some(command => command.kind === 'polygon' && command.fill.startsWith('rgba(46,187,122,')))
  await action('undo'); assert.deepEqual((await state()).model, before.model)
})

test('advanced Ant controls keep minimum/maximum, blank/text policies, zero axis and hidden cells independent', async () => {
  await open(['G1']); const before = await state(); await advanced()
  await form().getByRole('checkbox', { name: '迷你图显示零轴', exact: true }).check()
  await expect(form().locator('.ant-color-picker-trigger[aria-label="迷你图轴颜色"]')).toBeVisible()
  await form().getByRole('checkbox', { name: '迷你图从右到左', exact: true }).check()
  await option('迷你图下限方式', '整组统一计算'); await option('迷你图上限方式', '自定义'); await number('迷你图自定义上限', 120)
  await form().getByRole('checkbox', { name: '迷你图包含隐藏单元格', exact: true }).check()
  await option('迷你图空单元格处理', '按零值绘制'); await option('迷你图非数字单元格处理', '连接相邻数据点')
  await apply(); await waitHistory(1)
  const saved = await state(), current = value(saved, 'G2')
  assert.equal(current.extremumMin.type, 'group'); assert.deepEqual(current.extremumMax, { type: 'custom', value: 120 })
  assert.equal(current.axisOptions.visible, true); assert.equal(current.axisOptions.reverse, true); assert.equal(current.containHiddenCells, true)
  assert.equal(current.emptyCells, 'zero'); assert.equal(current.nonNumericCells, 'connect'); assert.deepEqual(value(saved, 'G4'), value(before, 'G4'))
})

test('custom extrema retain their values while hidden and reject an empty custom number before publication', async () => {
  await open(); await advanced(); await option('迷你图下限方式', '自定义'); await number('迷你图自定义下限', -12)
  await option('迷你图下限方式', '每个迷你图单独计算'); await expect(form().getByRole('spinbutton', { name: '迷你图自定义下限', exact: true })).toHaveCount(0)
  await option('迷你图下限方式', '自定义'); await expect(form().getByRole('spinbutton', { name: '迷你图自定义下限', exact: true })).toHaveValue('-12')
  await form().getByRole('spinbutton', { name: '迷你图自定义下限', exact: true }).fill(''); await page.keyboard.press('Tab')
  await expect(form().getByRole('button', { name: '应用', exact: true })).toBeDisabled(); assert.equal((await state()).history.undo, 0)
  await number('迷你图自定义下限', -15); await apply(); await waitHistory(1)
  assert.deepEqual(value(await state(), 'G1').extremumMin, { type: 'custom', value: -15 })
})

test('Pie hides irrelevant styling and disables its advanced tab while retaining real deletion controls', async () => {
  await open(); await option('迷你图类型', '饼图')
  await expect(form().getByRole('tab', { name: '高级设置', exact: true })).toHaveAttribute('aria-disabled', 'true')
  await expect(form().getByRole('combobox', { name: '迷你图主题', exact: true })).toHaveCount(0)
  await expect(form().getByRole('checkbox', { name: '迷你图高值点', exact: true })).toHaveCount(0)
  await expect(form().getByRole('button', { name: '删除当前迷你图', exact: true })).toBeDisabled()
  await apply(); await waitHistory(1); await open()
  await expect(form().getByRole('button', { name: '删除当前迷你图', exact: true })).toBeEnabled()
})

test('combine/uncombine keep the panel open, affect only selected members and use independent undo entries', async () => {
  await open(['G2', 'G4']); const before = await state(); await advanced()
  await form().getByRole('button', { name: '合并组', exact: true }).click(); await waitHistory(1); await expect(form()).toBeVisible()
  const combined = await state(); assert.equal(value(combined, 'G2').groupId, value(combined, 'G4').groupId)
  assert.deepEqual(value(combined, 'G1'), value(before, 'G1')); assert.deepEqual(value(combined, 'G5'), value(before, 'G5'))
  await form().getByRole('button', { name: '拆分组', exact: true }).click(); await waitHistory(2); await expect(form()).toBeVisible()
  const split = await state(); assert.notEqual(value(split, 'G2').groupId, value(split, 'G4').groupId)
  assert.equal(split.panelEvents.filter(event => event[0] === 'applied').length, 2)
  await action('undo'); assert.equal(value(await state(), 'G2').groupId, value(combined, 'G2').groupId)
  await action('undo'); assert.deepEqual((await state()).model, before.model)
})

test.each([['item', '删除当前迷你图'], ['group', '删除整组迷你图']])('clear %s commits once, preserves source content and survives undo/save-reopen', async (scope, label) => {
  await open(['G1']); const before = await state(); await form().getByRole('button', { name: label, exact: true }).click(); await waitHistory(1)
  const saved = await state(); assert.equal(value(saved, 'G1'), null)
  if (scope === 'item') assert.deepEqual(value(saved, 'G2'), value(before, 'G2')); else assert.equal(value(saved, 'G2'), null)
  assert.equal(saved.model.sheets[0].cells['0:1'].value, 2); await expect(form()).toHaveCount(0)
  await action('undo'); assert.deepEqual((await state()).model, before.model); await action('redo'); await action('reopen')
  assert.equal(value(await state(), 'G1'), null)
})

test('item source relocation and group settings publish together in one native history transaction', async () => {
  await open(['G2']); const before = await state()
  await option('迷你图数据操作', '修改当前迷你图的数据与位置'); await fields("'Data, O''Neil'!B2:E2", 'I3'); await option('迷你图线宽', '3 px')
  await apply(); await waitHistory(1)
  const saved = await state(); assert.equal(value(saved, 'G2'), null); assert.equal(value(saved, 'G1').lineWidth, 3)
  assert.equal(value(saved, 'I3').groupId, 'north'); assert.equal(value(saved, 'I3').sourceRange, "'Data, O''Neil'!B2:E2")
  const commands = saved.commands.filter(command => command.id === 'sheet.command.set-sparklines'); assert.equal(commands.length, 1)
  assert.equal(commands[0].params.request.kind, 'change-source'); assert.deepEqual(commands[0].params.request.settings, { lineWidth: 3 })
  await action('undo'); assert.deepEqual((await state()).model, before.model); await action('redo'); await action('reopen')
  assert.deepEqual(value(await state(), 'I3'), value(saved, 'I3'))
})

test('group source replacement changes member count while retaining one group ID and clearing old locations', async () => {
  await open(); await option('迷你图数据操作', '修改整组迷你图的数据与位置')
  await fields('B1:E1; B3:E3; B4:E4', 'I1; I3; I5'); await option('迷你图线宽', '4 px'); await apply(); await waitHistory(1)
  const saved = await state(); assert.equal(value(saved, 'G1'), null); assert.equal(value(saved, 'G2'), null)
  assert.ok(['I1', 'I3', 'I5'].every(address => value(saved, address).groupId === 'north' && value(saved, address).lineWidth === 4))
  assert.equal(saved.groups.find(group => group.groupId === 'north').members.length, 3)
  await action('reopen'); await open(['I1']); await expect(form().getByRole('combobox', { name: '当前迷你图单元格', exact: true })).toBeVisible()
})

test('conflicting data placement keeps the Ant draft open and cannot partially update settings or history', async () => {
  await open(); const before = await state(); await option('迷你图数据操作', '修改当前迷你图的数据与位置')
  await fields('B1:E1', 'G4'); await option('迷你图线宽', '4 px'); await apply()
  await expect(form().getByText('迷你图新放置区域与其他现有迷你图冲突。', { exact: true })).toBeVisible()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, 0); await expect(form()).toBeVisible()
})

test('a real false result preserves the panel and produces neither applied nor close notifications', async () => {
  await open(); await option('迷你图数据操作', '修改当前迷你图的数据与位置'); const before = await state()
  await form().getByRole('textbox', { name: '迷你图源数据区域', exact: true }).fill(' B1:E1 '); await apply()
  await expect(form().getByRole('status')).toContainText('当前操作没有产生内容变更。')
  const saved = await state(); assert.deepEqual(saved.model, before.model); assert.equal(saved.history.undo, 0)
  assert.equal(saved.panelEvents.filter(event => event[0] === 'applied' || event[0] === 'open' && event[1] === false).length, 0)
  await expect(form()).toBeVisible()
})

test.each(['revision', 'history', 'host'])('dirty panel drafts remain visible and block publication after %s changes', async kind => {
  await open(); await option('迷你图数据操作', '修改当前迷你图的数据与位置')
  const input = form().getByRole('textbox', { name: '迷你图源数据区域', exact: true }); await input.fill('B2:E2')
  if (kind === 'revision') await action('editCell', 'B1', 777)
  else if (kind === 'history') await action('clearHistory')
  else await action('switchSheet', 'other')
  const steered = await state(); await expect(input).toHaveValue('B2:E2')
  await expect(form().getByText('内容、用户或历史已更改。当前草稿已保留，请重新读取后再应用。', { exact: true })).toBeVisible()
  await expect(form().getByRole('button', { name: '应用', exact: true })).toBeDisabled()
  assert.deepEqual((await state()).model, steered.model)
})

test('explicit reload discards a stale draft and captures the latest real source and history owner', async () => {
  await open(); await option('迷你图数据操作', '修改当前迷你图的数据与位置')
  await form().getByRole('textbox', { name: '迷你图源数据区域', exact: true }).fill('B2:E2'); await action('editCell', 'B1', 888)
  await form().getByRole('button', { name: '重新读取', exact: true }).click()
  await expect(form().getByRole('textbox', { name: '迷你图源数据区域', exact: true })).toHaveValue('B1:E1')
  await fields('B2:E2', 'I1'); await apply(); await waitHistory(2); assert.equal(value(await state(), 'I1').sourceRange, 'B2:E2')
})

test.each(['target', 'studio'])('retained Ant input callbacks cannot edit a replacement %s panel', async kind => {
  await action('mount', { withGroups: false }); await open(['I1'])
  const callback = await action('captureCallback', '迷你图源数据区域')
  if (kind === 'studio') await action('switchUser', 'peer')
  await open(['I3']); await action('callback', callback, 'C1:D1')
  await expect(form().getByRole('textbox', { name: '迷你图源数据区域', exact: true })).toHaveValue('B1:E1')
  await apply(); await waitHistory(1); assert.equal(value(await state(), 'I3').sourceRange, 'B1:E1'); assert.equal(value(await state(), 'I1'), null)
})

test('synthetic IME keeps the final Ant text when a source revision makes the active draft stale', async () => {
  await open(); await option('迷你图数据操作', '修改当前迷你图的数据与位置')
  const input = form().getByRole('textbox', { name: '迷你图源数据区域', exact: true })
  await input.evaluate(node => {
    node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(node, 'B2:E2')
    node.dispatchEvent(new InputEvent('input', { bubbles: true, data: 'B2:E2', inputType: 'insertCompositionText', isComposing: true }))
  })
  await expect(form().getByRole('button', { name: '应用', exact: true })).toBeDisabled(); await action('editCell', 'B1', 777)
  await input.dispatchEvent('compositionend', { bubbles: true, data: 'B2:E2' })
  await expect(input).toHaveValue('B2:E2'); await expect(form().getByRole('button', { name: '应用', exact: true })).toBeDisabled()
  assert.equal((await state()).history.undo, 1)
})

test('an obsolete Modal afterClose callback cannot close a replacement panel or restore the prior input focus', async () => {
  await action('mount', { withGroups: false }); await open(['I1'])
  await form().getByRole('button', { name: '取消', exact: true }).click()
  const callback = await action('captureAfterClose'); await open(['I3'])
  const current = form().getByRole('textbox', { name: '迷你图源数据区域', exact: true }); await current.focus()
  const before = await state(); await action('callback', callback)
  await expect(form()).toBeVisible(); await expect(current).toBeFocused()
  assert.deepEqual((await state()).panelEvents, before.panelEvents); assert.equal((await state()).history.undo, 0)
})

test('a protected unselected member disables full-group edits and real management actions', async () => {
  await action('mount', { protectedMember: true }); await open(['G1'])
  await expect(form().getByText('当前用户没有修改目标区域或整组迷你图的权限。', { exact: true })).toBeVisible()
  await expect(form().getByRole('combobox', { name: '迷你图类型', exact: true })).toBeDisabled()
  await expect(form().getByRole('button', { name: '删除整组迷你图', exact: true })).toBeDisabled(); assert.equal((await state()).history.undo, 0)
})

test('Ant form, popups and Canvas retain editable theme behavior in classic and grid layouts', async () => {
  await open(); const light = await form().evaluate(node => getComputedStyle(node).color)
  assert.ok(await form().locator('.ant-form-item').count()); assert.equal(await form().locator('input:not([class*="ant-"]),textarea:not([class*="ant-"])').count(), 0)
  await action('appearance', true, 'grid'); const dark = await form().evaluate(node => getComputedStyle(node).color); assert.notEqual(dark, light)
  await form().getByRole('combobox', { name: '迷你图主题', exact: true }).click()
  const popup = page.locator('.ant-select-dropdown:visible'), background = await popup.evaluate(node => getComputedStyle(node).backgroundColor)
  assert.notEqual(background, 'rgb(255, 255, 255)'); await popup.getByText('柔和', { exact: true }).click()
  await apply(); await waitHistory(1); assert.equal(value(await state(), 'G1').themeType, 'soft')
  await action('zoom', 125); await action('scroll', 30, 10); assert.ok((await action('pixels', 'G1')).bytes > 0)
})

test.each(['cancel', 'apply'])('the production panel restores the captured Ant name-box focus after %s', async mode => {
  await action('mount', { withGroups: false }); const name = page.getByRole('textbox', { name: '名称框', exact: true })
  await name.focus(); await open(['I1'])
  if (mode === 'cancel') await form().getByRole('button', { name: '取消', exact: true }).click()
  else { await fields('B2:E2', 'I1'); await apply(); await waitHistory(1) }
  await expect(form()).toHaveCount(0); await expect(name).toBeFocused()
  assert.ok((await state()).panelEvents.some(event => event[0] === 'closed' && event[1].sheetId === 'data'))
})
