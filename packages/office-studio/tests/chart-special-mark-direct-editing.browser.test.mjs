import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// P20 browser drafts only. Real ExcelWorkbench, Ant forms, ECharts Canvas2D,
// canonical updates and internal save/reopen will run during unified validation.
// No server, browser, build, WASM task or check was started to write this file.
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
const rows = [['Category', 'Original'], ['A', 8], ['B', 5], ['C', 3], ['D', 2]]
const compoundOptions = () => ({ series: { 1: { dataPoints: { 2: { color: '#334455' } } } },
  pie: { aggregateLabel: 'Other', label: { visible: true, position: 'inside' }, secondaryPlot: { split: { type: 'position', count: 2 } } } })
async function mount(options) {
  await page.evaluate(options => window.chartObjectFixture.mount(options), options)
  await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
}
async function frames() { await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))) }
async function closeMenu() {
  if (await menu().count()) { await menu().getByRole('button', { name: '关闭图表元素工具栏', exact: true }).click(); await expect(menu()).toHaveCount(0) }
}
async function select(selector) {
  await closeMenu(); const measured = await page.evaluate(selector => window.chartObjectFixture.elementPoint(selector), { kind: 'mark', ...selector })
  await page.mouse.click(measured.viewport.x, measured.viewport.y)
  await expect(menu()).toBeVisible(); await expect(menu()).toHaveAttribute('data-chart-element-kind', 'mark')
  await expect(menu()).toHaveAttribute('data-chart-mark-kind', selector.markKind)
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

test('measured Pareto bar and cumulative-line marks commit independent styles, preserve point maps and follow canonical history', async () => {
  await mount({ nativeType: 'pareto', rows, nativeOptions: { pareto: {
    barStyle: { dataPoints: { 0: { color: '#112233' } } }, cumulativeLineStyle: { dataPoints: { 1: { size: 18 } } },
  } } }); const before = await state()
  await select({ markKind: 'paretoStyle', markPart: 'bar' }); await number('pareto.barStyle.opacity', 0.4)
  assert.deepEqual((await state()).charts, before.charts); await apply()
  let next = await state(); assert.equal(chart(next).options.pareto.barStyle.opacity, 0.4)
  assert.deepEqual(chart(next).options.pareto.barStyle.dataPoints, chart(before).options.pareto.barStyle.dataPoints)
  assert.deepEqual(chart(next).options.pareto.cumulativeLineStyle, chart(before).options.pareto.cumulativeLineStyle)
  assert.equal(next.history.undo, before.history.undo + 1)
  await select({ markKind: 'paretoStyle', markPart: 'line' }); await number('pareto.cumulativeLineStyle.width', 4); await apply()
  next = await state(); assert.equal(chart(next).options.pareto.cumulativeLineStyle.width, 4)
  assert.equal(chart(next).options.pareto.barStyle.opacity, 0.4); assert.equal(next.history.undo, before.history.undo + 2)
  await menu().getByRole('button', { name: '恢复图表元素默认格式', exact: true }).click(); await apply()
  next = await state(); assert.equal(chart(next).options.pareto.cumulativeLineStyle.width, undefined)
  assert.deepEqual(chart(next).options.pareto.cumulativeLineStyle.dataPoints, chart(before).options.pareto.cumulativeLineStyle.dataPoints)
  assert.deepEqual(next.model.sheets[0].cells, before.model.sheets[0].cells)
  assert.equal(next.commands.filter(command => command.id === 'sheet.command.update-chart').length, 3)
  await page.evaluate(() => window.chartObjectFixture.undo()); assert.equal(chart(await state()).options.pareto.cumulativeLineStyle.width, 4)
  await page.evaluate(() => window.chartObjectFixture.redo()); await page.evaluate(() => window.chartObjectFixture.reopen())
  assert.deepEqual(chart(await state()), chart(next))
}, 60_000)

test('a source-bound Waterfall role edit leaves other roles, source series and data intact and consumes Delete without deleting the chart', async () => {
  await mount({ nativeType: 'waterfall', multiSeries: true, rows: [['Category', 'First', 'Second'], ['A', 10, 4], ['B', -3, -2]],
    nativeOptions: { waterfall: { pointRoles: [], connector: true }, series: {
      1: { waterfallStyles: { positive: { color: '#00aa00' }, negative: { dataPoints: { 1: { color: '#aa0000' } } } } },
      2: { waterfallStyles: { negative: { color: '#222222' } } },
    } } }); const before = await state(), measured = await select({ markKind: 'waterfallStyle', markPart: 'negative', seriesId: 'field:1' })
  assert.equal(measured.target.element.directEdit.binding.sourceName, 'First')
  await number('series.1.waterfallStyles.negative.fillOpacity', 0.4); await apply()
  const next = await state(); assert.equal(chart(next).options.series[1].waterfallStyles.negative.fillOpacity, 0.4)
  assert.deepEqual(chart(next).options.series[1].waterfallStyles.negative.dataPoints, chart(before).options.series[1].waterfallStyles.negative.dataPoints)
  assert.deepEqual(chart(next).options.series[1].waterfallStyles.positive, chart(before).options.series[1].waterfallStyles.positive)
  assert.deepEqual(chart(next).options.series[2], chart(before).options.series[2]); assert.deepEqual(chart(next).options.waterfall.pointRoles, [])
  assert.equal(next.history.undo, before.history.undo + 1)
  await hitbox().focus(); await hitbox().press('Delete'); assert.equal((await state()).history.undo, next.history.undo)
  assert.equal((await state()).charts.length, 2); assert.deepEqual((await state()).model.sheets[0].cells, before.model.sheets[0].cells)
}, 60_000)

test('a named Waterfall mark uses its highest author layer, and dirty More commits before opening the real bound full panel', async () => {
  await mount({ nativeType: 'waterfall', rows: [['Category', 'Original'], ['A', 6], ['B', -2], ['C', 4]], nativeOptions: {
    waterfall: { pointRoles: [], connector: true }, series: { 1: { waterfallStyles: { negative: { color: '#123456' } } } },
    namedSeries: [{ name: 'Original', style: { waterfallStyles: { negative: { color: '#654321' }, positive: { color: '#00aa00' } } } }],
  } }); const before = await state(), measured = await select({ markKind: 'waterfallStyle', markPart: 'negative', seriesId: 'name:"Original"' })
  assert.deepEqual(measured.target.element.directEdit.binding.selector, { name: 'Original' })
  await number('namedSeries."Original".waterfallStyles.negative.fillOpacity', 0.5)
  await menu().getByRole('button', { name: '图表元素更多设置', exact: true }).click()
  await expect(menu()).toHaveCount(0); await expect(drawer()).toBeVisible()
  await expect(drawer().locator('.chart-waterfall-settings [data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', 'name:"Original"')
  await expect(drawer().locator('.chart-waterfall-settings')).toHaveAttribute('data-chart-waterfall-role', 'negative')
  await expect(drawer().locator('[data-chart-option-key=\'namedSeries."Original".waterfallStyles.negative.color\']').getByRole('combobox')).toBeFocused()
  const next = await state(); assert.equal(chart(next).options.namedSeries[0].style.waterfallStyles.negative.fillOpacity, 0.5)
  assert.deepEqual(chart(next).options.series, chart(before).options.series)
  assert.deepEqual(chart(next).options.namedSeries[0].style.waterfallStyles.positive, chart(before).options.namedSeries[0].style.waterfallStyles.positive)
  assert.equal(next.history.undo, before.history.undo + 1)
  assert.equal(next.commands.filter(command => command.id === 'sheet.command.chart.open-focused-panel').length, 1)
  assert.deepEqual(next.commands.find(command => command.id === 'sheet.command.chart.open-focused-panel').params.route.focus,
    { kind: 'waterfallRole', role: 'negative' })
}, 60_000)

test('the measured stacked global subtotal edits only shared subtotal style and never writes a synthetic source point', async () => {
  await mount({ nativeType: 'waterfall', multiSeries: true, rows: [['Category', 'First', 'Second'], ['A', 10, 4], ['B', -3, -2]],
    nativeOptions: { waterfall: { stackType: 'stacked', useSubtotal: true, connector: true },
      series: { 1: { waterfallStyles: { subtotal: { color: '#112233' } } }, 2: { waterfallStyles: { negative: { color: '#223344' } } } } } })
  const before = await state(), measured = await select({ markKind: 'waterfallStyle', markPart: 'subtotal', globalMark: true })
  assert.equal(measured.target.element.directEdit.binding, undefined); await expect(menu()).toContainText('全局合计')
  await number('waterfall.styles.subtotal.fillOpacity', 0.5); await apply()
  const next = await state(); assert.equal(chart(next).options.waterfall.styles.subtotal.fillOpacity, 0.5)
  assert.deepEqual(chart(next).options.series, chart(before).options.series)
  assert.equal(chart(next).options.waterfall.styles.subtotal.dataIndex, undefined)
  assert.equal(chart(next).options.series[4294967295], undefined); assert.equal(next.history.undo, before.history.undo + 1)
  await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(next))
}, 60_000)

test('Pie-of-Pie and Bar-of-Pie aggregate forms author independent opacity, explosion and label through save/reopen', async () => {
  for (const nativeType of ['pieOfPie', 'barOfPie']) {
    await mount({ nativeType, rows, nativeOptions: compoundOptions() }); const before = await state()
    const measured = await select({ markKind: 'pieAggregateStyle' })
    assert.equal(measured.target.element.directEdit.aggregateIndex, 4); assert.equal(measured.target.element.dataIndex, undefined)
    await expect(menu()).toContainText('副图数据'); await field('pie.aggregateLabel').getByRole('textbox').fill('Remaining categories')
    await number('pie.aggregateStyle.fillOpacity', 0.5); await number('pie.aggregateStyle.explosion', 20); await apply()
    const next = await state(); assert.equal(chart(next).options.pie.aggregateStyle.fillOpacity, 0.5); assert.equal(chart(next).options.pie.aggregateStyle.explosion, 20)
    assert.equal(chart(next).options.pie.aggregateLabel, 'Remaining categories'); assert.deepEqual(chart(next).options.series, chart(before).options.series)
    assert.deepEqual(chart(next).options.pie.secondaryPlot, chart(before).options.pie.secondaryPlot); assert.equal(next.history.undo, before.history.undo + 1)
    assert.deepEqual(chart(next).position, chart(before).position); assert.deepEqual(next.model.sheets[0].cells, before.model.sheets[0].cells)
    await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(next))
    await select({ markKind: 'pieAggregateStyle' }); await menu().getByRole('button', { name: '恢复图表元素默认格式', exact: true }).click(); await apply()
    assert.equal(chart(await state()).options.pie.aggregateStyle, undefined); assert.equal(chart(await state()).options.pie.aggregateLabel, 'Remaining categories')
  }
}, 60_000)

test('both composite Pie connectors receive style edits, dirty More opens their real controls, and Delete hides only the connections', async () => {
  await mount({ nativeType: 'pieOfPie', rows, nativeOptions: compoundOptions() }); const before = await state()
  await select({ markKind: 'pieConnectorStyle' }); await number('pie.secondaryPlot.connector.width', 4)
  await choose('pie.secondaryPlot.connector.dashType', '虚线'); await apply()
  const scene = await page.evaluate(() => window.chartObjectFixture.scene()), connections = scene.filter(command => command.chartCanvasElement?.directEdit?.kind === 'pieConnectorStyle')
  assert.ok(connections.length > 2); assert.ok(connections.every(command => command.kind === 'line' && command.width === 4))
  await number('pie.secondaryPlot.connector.opacity', 0.7)
  await menu().getByRole('button', { name: '图表元素更多设置', exact: true }).click()
  await expect(drawer()).toBeVisible(); await expect(drawer().locator('[data-chart-option-key="pie.secondaryPlot.connector.width"]')).toHaveCount(1)
  await expect(drawer().locator('[data-chart-option-key="pie.secondaryPlot.connector.visible"]').getByRole('switch')).toBeFocused()
  assert.equal(chart(await state()).options.pie.secondaryPlot.connector.opacity, 0.7)
  await page.keyboard.press('Escape'); await expect(drawer()).not.toBeVisible()
  await select({ markKind: 'pieConnectorStyle' }); await hitbox().focus(); await hitbox().press('Delete'); await expect(menu()).toHaveCount(0)
  const next = await state(); assert.equal(chart(next).options.pie.secondaryPlot.connector.visible, false); assert.equal(next.charts.length, 2)
  assert.deepEqual(chart(next).options.series, chart(before).options.series); assert.deepEqual(chart(next).options.pie.secondaryPlot.split, chart(before).options.pie.secondaryPlot.split)
  assert.equal(next.history.undo, before.history.undo + 3)
  assert.ok((await page.evaluate(() => window.chartObjectFixture.scene())).every(command => command.chartCanvasElement?.directEdit?.kind !== 'pieConnectorStyle'))
}, 60_000)

test('special-mark drafts respect read-only and IME boundaries, and a pending completion cannot revive a target after type and sheet changes', async () => {
  await mount({ readonly: true, nativeType: 'pieOfPie', rows, nativeOptions: compoundOptions() }); const readonly = await state()
  await select({ markKind: 'pieAggregateStyle' }); await expect(field('pie.aggregateLabel').getByRole('textbox')).toBeDisabled()
  await expect(menu().getByRole('button', { name: '图表元素更多设置', exact: true })).toBeDisabled()
  await hitbox().focus(); await hitbox().press('Delete'); assert.deepEqual((await state()).charts, readonly.charts)
  await mount({ nativeType: 'pieOfPie', rows, nativeOptions: compoundOptions() }); const before = await state()
  await select({ markKind: 'pieAggregateStyle' }); const input = field('pie.aggregateLabel').getByRole('textbox')
  await input.evaluate(node => { node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })); node.value = '合计草稿'
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '稿', isComposing: true }))
    for (const key of ['Enter', 'Escape']) node.dispatchEvent(new KeyboardEvent('keydown', { key, isComposing: true, bubbles: true, cancelable: true })) })
  await expect(menu()).toBeVisible(); assert.deepEqual((await state()).charts, before.charts)
  await input.evaluate(node => { node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '合计草稿' })); node.dispatchEvent(new InputEvent('input', { bubbles: true })) })
  await input.press('Escape'); await expect(menu()).toHaveCount(0); await expect(hitbox()).toBeFocused()
  assert.equal((await state()).history.undo, before.history.undo)
  await select({ markKind: 'pieAggregateStyle' }); await number('pie.aggregateStyle.fillOpacity', 0.4)
  await page.evaluate(() => window.chartObjectFixture.holdAuthor()); await menu().getByRole('button', { name: '应用图表元素修改', exact: true }).click()
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingAuthors === 1)
  await page.evaluate(() => { void window.chartObjectFixture.updateChart({ type: 'pie' }) })
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingAuthors === 2)
  await expect(menu()).toHaveCount(0); assert.equal(chart(await state()).type, 'pie')
  await page.evaluate(() => window.chartObjectFixture.switchSheet('destination')); await grid().focus()
  await page.evaluate(() => window.chartObjectFixture.finishAuthors()); await frames()
  await expect(menu()).toHaveCount(0); await expect(grid()).toBeFocused(); assert.equal(chart(await state()).options.pie.aggregateStyle.fillOpacity, 0.4)
  assert.deepEqual((await state()).model.sheets[0].cells, before.model.sheets[0].cells)
}, 60_000)
