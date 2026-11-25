import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// P19 drafts only: real ExcelWorkbench/Ant controls/ECharts Canvas2D,
// measured native hits, canonical author history and internal save/reopen.
// No server, browser, tests or checks were started while writing these drafts.
let server, browser, context, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No HTTP port was allocated')
  baseURL = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/chart-object-actions.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0
  context = await browser.newContext({ viewport: { width: 1440, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] })
  page = await context.newPage(); page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await page.waitForFunction(() => window.chartObjectFixture?.state().charts.length === 2)
  await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
}, 60_000)
afterEach(async () => { try { await page?.evaluate(() => window.chartObjectFixture?.dispose()); assert.deepEqual(errors, []) } finally { await context?.close() } })
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.chartObjectFixture.state())
const chart = value => value.charts.find(chart => chart.id === 'native-object')
const hitbox = () => page.locator('.chart-hitbox[data-chart-id="native-object"]')
const menu = () => page.locator('.chart-element-float-menu')
const field = key => menu().locator(`[data-chart-option-key=${JSON.stringify(key)}]`)
const drawer = () => page.locator('.ant-drawer').filter({ has: page.getByText('图表设置', { exact: true }) })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const axisConfig = { xAxis: { title: 'Categories', scaleType: 'linear', line: { visible: true }, tick: { visible: true } },
  yAxis: { title: 'Values', scaleType: 'linear', line: { visible: true }, tick: { visible: true }, gridLine: { visible: true } } }
async function mount(options = {}) {
  await page.evaluate(options => window.chartObjectFixture.mount(options), options)
  await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
}
async function frames() { await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))) }
async function closeMenu() {
  if (await menu().count()) { await menu().getByRole('button', { name: '关闭图表元素工具栏', exact: true }).click(); await expect(menu()).toHaveCount(0) }
}
async function point(selector) { return page.evaluate(selector => window.chartObjectFixture.elementPoint(selector), selector) }
async function select(selector, { shift = false, double = false } = {}) {
  await closeMenu(); const measured = await point(selector)
  if (shift) await page.keyboard.down('Shift')
  try { await page.mouse[double ? 'dblclick' : 'click'](measured.viewport.x, measured.viewport.y) }
  finally { if (shift) await page.keyboard.up('Shift') }
  await expect(menu()).toBeVisible(); await expect(menu()).toHaveAttribute('data-chart-element-kind', shift ? 'series' : selector.kind)
  return measured
}
async function number(key, value) { const input = field(key).getByRole('spinbutton'); await input.fill(String(value)); await input.blur() }
async function choose(key, label) {
  await field(key).getByRole('combobox').click()
  await menu().locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}
async function apply() {
  const button = menu().getByRole('button', { name: '应用图表元素修改', exact: true })
  await expect(button).toBeEnabled(); await button.click()
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingAuthors === 0
    && document.querySelector('.chart-hitbox[data-chart-id="native-object"]')?.getAttribute('aria-busy') === 'false')
  await frames()
}

test('a measured Canvas axis title opens its Ant input and one scoped Apply follows canonical history/save/reopen', async () => {
  await mount({ multiSeries: true, nativeConfig: axisConfig }); const before = await state()
  const measured = await select({ kind: 'axisTitle', axisId: 'xAxis' })
  assert.equal(measured.target.element.axisId, 'xAxis')
  const input = field('xAxis.title').getByRole('textbox'); await expect(input).toBeFocused(); await input.fill('Authored categories')
  assert.deepEqual((await state()).charts, before.charts); await apply()
  const authored = await state(); assert.equal(chart(authored).xAxis.title, 'Authored categories')
  assert.equal(chart(authored).yAxis.title, chart(before).yAxis.title); assert.equal(authored.history.undo, before.history.undo + 1)
  assert.equal(authored.commands.filter(command => command.id === 'sheet.command.update-chart').length, 1)
  await expect(menu().getByRole('button', { name: '应用图表元素修改', exact: true })).toBeDisabled()
  await page.evaluate(() => window.chartObjectFixture.undo()); assert.deepEqual(chart(await state()), chart(before)); await expect(menu()).toHaveCount(0)
  await page.evaluate(() => window.chartObjectFixture.redo()); await page.evaluate(() => window.chartObjectFixture.reopen())
  assert.deepEqual(chart(await state()), chart(authored))
  assert.equal(chart(await state()).officeChartElements, undefined)
})

test('main/secondary axis, gridline and tick forms write only their measured scope and Delete hides an axis without deleting its chart', async () => {
  await mount({ multiSeries: true, nativeConfig: axisConfig }); const before = await state()
  await select({ kind: 'axis', axisId: 'rightYAxis', part: 'label' }); await number('rightYAxis.max', 25); await apply()
  let next = await state(); assert.equal(chart(next).options.rightYAxis.max, 25); assert.equal(chart(next).yAxis.max, undefined)
  assert.equal(next.history.undo, before.history.undo + 1)
  await select({ kind: 'tick', axisId: 'xAxis' }); await number('xAxis.tick.length', 8); await apply()
  next = await state(); assert.equal(chart(next).xAxis.tick.length, 8); assert.equal(chart(next).options.rightYAxis.max, 25)
  await select({ kind: 'gridLine', axisId: 'yAxis' }); await number('yAxis.gridLine.width', 3); await apply()
  next = await state(); assert.equal(chart(next).yAxis.gridLine.width, 3); assert.deepEqual(chart(next).source, chart(before).source)
  await select({ kind: 'axis', axisId: 'xAxis', part: 'label' }); await choose('xAxis.label.rotate', '45°'); await apply()
  assert.equal(chart(await state()).xAxis.label.rotate, 45)
  await hitbox().focus(); await hitbox().press('Delete'); await expect(menu()).toHaveCount(0)
  next = await state(); assert.equal(chart(next).xAxis.label.visible, false); assert.equal(next.charts.length, 2)
  assert.deepEqual(next.model.sheets[0].cells, before.model.sheets[0].cells)
})

test('point edits, scoped reset and Shift-series selection preserve source data, siblings and the actual field binding', async () => {
  await mount({ multiSeries: true, nativeOptions: { series: { 1: { dataPoints: { 0: { size: 13 } } }, 2: { rightAxis: true, name: 'Renamed comparison' } } } })
  const before = await state(); await select({ kind: 'dataPoint', seriesId: 'field:1', dataIndex: 1 })
  await number('point.size', 16); await number('point.fillOpacity', 0.4); await apply()
  let next = await state(); assert.equal(chart(next).options.series[1].dataPoints[1].size, 16)
  assert.equal(chart(next).options.series[1].dataPoints[1].fillOpacity, 0.4)
  assert.deepEqual(chart(next).options.series[1].dataPoints[0], chart(before).options.series[1].dataPoints[0])
  assert.deepEqual(chart(next).options.series[2], chart(before).options.series[2]); assert.equal(next.history.undo, before.history.undo + 1)
  await menu().getByRole('button', { name: '恢复图表元素默认格式', exact: true }).click(); await apply()
  assert.equal(chart(await state()).options.series[1].dataPoints?.[1], undefined)
  await hitbox().focus(); const count = (await state()).history.undo; await hitbox().press('Delete')
  assert.equal((await state()).history.undo, count); assert.equal((await state()).charts.length, 2)
  await select({ kind: 'dataPoint', seriesId: 'field:1', dataIndex: 1 }, { shift: true })
  await expect(menu()).toHaveAttribute('data-chart-element-key', 'series:field:1')
  await number('series.1.point.size', 18); await apply(); next = await state()
  assert.equal(chart(next).options.series[1].point.size, 18); assert.deepEqual(next.model.sheets[0].cells, before.model.sheets[0].cells)
})

test('ordinary labels edit their bound series, while Pie labels use the global pie.label author and disclose that scope', async () => {
  await mount({ multiSeries: true, nativeOptions: { series: { 1: { label: { visible: true, position: 'top', fontSize: 12 } }, 2: { rightAxis: true, name: 'Renamed comparison' } } } })
  let before = await state(); await select({ kind: 'dataLabel', seriesId: 'field:1', dataIndex: 1 })
  await expect(menu()).toContainText('所选系列的数据标签'); await number('series.1.label.fontSize', 18); await apply()
  let next = await state(); assert.equal(chart(next).options.series[1].label.fontSize, 18)
  assert.deepEqual(chart(next).options.series[2], chart(before).options.series[2]); assert.equal(next.history.undo, before.history.undo + 1)
  await mount({ nativeType: 'pie', nativeOptions: { pie: { label: { visible: true, position: 'inside', fontSize: 12 } } } }); before = await state()
  await select({ kind: 'mark', markKind: 'pieStyle' }); await expect(menu()).toContainText('全部扇区的数据标签')
  await number('pie.label.fontSize', 19); await apply(); next = await state()
  assert.equal(chart(next).options.pie.label.fontSize, 19); assert.equal(chart(next).options.series?.[1]?.label, undefined)
  assert.equal(next.history.undo, before.history.undo + 1); await page.evaluate(() => window.chartObjectFixture.reopen())
  assert.deepEqual(chart(await state()), chart(next))
})

test('a normal legend click changes only native view selection, while a double click edits the whole legend through one author command', async () => {
  await mount({ multiSeries: true }); const before = await state(), measured = await point({ kind: 'legend' })
  await page.mouse.click(measured.viewport.x, measured.viewport.y); await frames(); await expect(menu()).toHaveCount(0)
  const view = await page.evaluate(() => window.chartObjectFixture.legendView())
  assert.equal(view.selected[measured.target.element.name], false); assert.deepEqual((await state()).charts, before.charts)
  assert.equal((await state()).history.undo, before.history.undo)
  await select({ kind: 'legend' }, { double: true }); await choose('legend.position', '右侧'); await number('legend.font.fontSize', 15); await apply()
  const authored = await state(); assert.equal(chart(authored).legend.position, 'right'); assert.equal(chart(authored).legend.font.fontSize, 15)
  assert.equal(authored.history.undo, before.history.undo + 1); assert.deepEqual(chart(authored).source, chart(before).source)
})

test('More first commits its dirty point draft, then the canonical focused-panel command opens the real bound series and focus', async () => {
  const before = await state(); await select({ kind: 'dataPoint', seriesId: 'field:1', dataIndex: 1 }); await number('point.size', 11)
  await menu().getByRole('button', { name: '图表元素更多设置', exact: true }).click()
  await expect(menu()).toHaveCount(0); await expect(drawer()).toBeVisible()
  await expect(drawer().locator('[data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', 'field:1')
  await expect(drawer().getByRole('combobox', { name: '选择图表系列', exact: true })).toBeFocused()
  const next = await state(); assert.equal(chart(next).options.series[1].dataPoints[1].size, 11)
  assert.equal(next.history.undo, before.history.undo + 1)
  assert.equal(next.commands.filter(command => command.id === 'sheet.command.chart.open-focused-panel').length, 1)
})

test('IME Enter/Escape never commit mid-composition; cancelling the Ant draft and read-only element actions preserve content', async () => {
  await mount({ nativeConfig: axisConfig }); const before = await state(); await select({ kind: 'axisTitle', axisId: 'xAxis' })
  const input = field('xAxis.title').getByRole('textbox')
  await input.evaluate(node => { node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })); node.value = '输入法草稿'
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '稿', isComposing: true }))
    for (const key of ['Enter', 'Escape']) node.dispatchEvent(new KeyboardEvent('keydown', { key, isComposing: true, bubbles: true, cancelable: true })) })
  await expect(menu()).toBeVisible(); await expect(input).toBeFocused(); assert.deepEqual((await state()).charts, before.charts)
  await input.evaluate(node => { node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '输入法草稿' }))
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '输入法草稿' })) })
  await input.press('Escape'); await expect(menu()).toHaveCount(0); await expect(hitbox()).toBeFocused()
  assert.deepEqual((await state()).charts, before.charts); assert.equal((await state()).history.undo, before.history.undo)
  await mount({ readonly: true, nativeConfig: axisConfig }); const guarded = await state(); await select({ kind: 'axisTitle', axisId: 'xAxis' })
  await expect(field('xAxis.title').getByRole('textbox')).toBeDisabled()
  await expect(menu().getByRole('button', { name: '图表元素更多设置', exact: true })).toBeDisabled()
  await hitbox().focus(); await hitbox().press('Delete'); assert.deepEqual((await state()).charts, guarded.charts)
  assert.equal((await state()).history.undo, guarded.history.undo)
})

test('zoom/theme/font/scroll changes retire drafts, and busy or late canonical completions cannot reopen old sheet/owner sessions', async () => {
  await mount({ nativeConfig: axisConfig }); const before = await state()
  for (const kind of ['zoom', 'theme', 'fonts', 'scroll']) {
    await select({ kind: 'axisTitle', axisId: 'xAxis' }); await field('xAxis.title').getByRole('textbox').fill(`Retired ${kind}`)
    await page.evaluate(async kind => {
      if (kind === 'zoom') await window.chartObjectFixture.zoom(125)
      else if (kind === 'theme') await window.chartObjectFixture.appearance(true, 'grid')
      else if (kind === 'fonts') document.fonts.dispatchEvent(new Event('loadingdone'))
      else await window.chartObjectFixture.scroll(20, 20)
    }, kind)
    await expect(menu()).toHaveCount(0); assert.deepEqual((await state()).charts, before.charts)
  }
  await mount({ nativeConfig: axisConfig }); await select({ kind: 'axisTitle', axisId: 'xAxis' })
  await field('xAxis.title').getByRole('textbox').fill('Committed before leave'); await page.evaluate(() => window.chartObjectFixture.holdAuthor())
  await menu().getByRole('button', { name: '应用图表元素修改', exact: true }).click()
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingAuthors === 1)
  await expect(menu().getByRole('button', { name: '图表元素更多设置', exact: true })).toBeDisabled()
  await hitbox().focus(); await hitbox().press('F2'); await expect(drawer()).toHaveCount(0)
  await page.evaluate(() => window.chartObjectFixture.switchSheet('destination')); await grid().focus()
  await page.evaluate(() => window.chartObjectFixture.finishAuthors()); await frames()
  await expect(menu()).toHaveCount(0); await expect(grid()).toBeFocused(); assert.equal(chart(await state()).xAxis.title, 'Committed before leave')
  await page.evaluate(() => window.chartObjectFixture.switchSheet('source')); await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
  await select({ kind: 'axisTitle', axisId: 'xAxis' }); await field('xAxis.title').getByRole('textbox').fill('Old owner draft')
  await page.evaluate(() => window.chartObjectFixture.reopen()); await grid().focus(); await frames()
  await expect(menu()).toHaveCount(0); await expect(grid()).toBeFocused(); assert.equal(chart(await state()).xAxis.title, 'Committed before leave')
})
