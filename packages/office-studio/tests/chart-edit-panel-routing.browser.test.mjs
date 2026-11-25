import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// P17 deferred drafts: the real ExcelWorkbench, Ant Drawer/Tabs/Form and Canvas.
// No server, browser, tests, build, WASM or benchmark was started this batch.
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
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] })
  page = await context.newPage(); page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await page.waitForFunction(() => window.chartObjectFixture?.state().charts.length === 2)
  await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
}, 60_000)
afterEach(async () => { try { await page?.evaluate(() => window.chartObjectFixture?.dispose()); assert.deepEqual(errors, []) } finally { await context?.close() } })
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.chartObjectFixture.state())
const hitbox = (id = 'native-object') => page.locator(`.chart-hitbox[data-chart-id="${id}"]`)
const drawer = () => page.locator('.ant-drawer').filter({ has: page.getByText('图表设置', { exact: true }) })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const field = key => drawer().locator(`[data-chart-option-key=${JSON.stringify(key)}]`)
const chart = (value, id = 'native-object') => value.charts.find(item => item.id === id)
async function mount(options) { await page.evaluate(options => window.chartObjectFixture.mount(options), options); await expect(hitbox()).toHaveAttribute('data-rendered', 'true') }
async function prepare(id = 'native-object') { await page.evaluate(id => window.chartObjectFixture.prepare(id), id) }
async function route(value, id = 'native-object') {
  assert.equal(await page.evaluate(({ id, value }) => window.chartObjectFixture.openPanel(id, value), { id, value }), true)
  await expect(drawer()).toBeVisible()
}
async function apply() { const button = drawer().getByRole('button', { name: '应用', exact: true }); await expect(button).toBeEnabled(); await button.click() }
async function frames() { await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))) }

test('the data route opens the real data tab/source input; manual tab switches retain an unapplied title draft', async () => {
  const before = await state(); await route({ tab: 'style', titleKey: 'title' })
  await field('title').getByRole('textbox').fill('Unapplied title')
  await route({ tab: 'data' })
  await expect(drawer().getByRole('tab', { name: '数据', exact: true })).toHaveAttribute('aria-selected', 'true')
  await expect(drawer().getByRole('textbox', { name: '图表数据区域', exact: true })).toBeFocused()
  await expect(field('title')).toHaveCount(0)
  await drawer().getByRole('tab', { name: '样式', exact: true }).click()
  await expect(field('title').getByRole('textbox')).toHaveValue('Unapplied title')
  const after = await state(); assert.deepEqual(after.charts, before.charts); assert.equal(after.history.undo, before.history.undo); assert.equal(after.revision, before.revision)
})

test('all five title routes focus their own Ant input and one Apply persists the collected titles through undo/redo/reopen', async () => {
  await mount({ multiSeries: true }); await prepare(); const before = await state()
  for (const [titleKey, optionKey, text] of [
    ['title', 'title', 'Routed title'], ['subtitle', 'subtitle', 'Routed subtitle'], ['xAxisTitle', 'xAxis.title', 'Categories'],
    ['yAxisTitle', 'yAxis.title', 'Primary values'], ['rightYAxisTitle', 'rightYAxis.title', 'Secondary values'],
  ]) {
    await route({ tab: 'style', titleKey }); const input = field(optionKey).getByRole('textbox')
    await expect(input).toBeFocused(); await input.fill(text)
  }
  assert.deepEqual((await state()).charts, before.charts); await apply()
  await page.waitForFunction(() => window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').options.rightYAxis.title === 'Secondary values')
  const authored = await state(); assert.equal(authored.history.undo, before.history.undo + 1)
  assert.equal(chart(authored).title, 'Routed title'); assert.equal(chart(authored).subtitle, 'Routed subtitle'); assert.equal(chart(authored).yAxis.title, 'Primary values')
  await page.evaluate(() => window.chartObjectFixture.undo()); assert.equal(chart(await state()).title, chart(before).title)
  await page.evaluate(() => window.chartObjectFixture.redo()); await page.evaluate(() => window.chartObjectFixture.reopen())
  assert.deepEqual(chart(await state()), chart(authored))
})

test('horizontal, vertical, secondary and gridline routes expand the actual selected axis and do not create history', async () => {
  await mount({ multiSeries: true }); await prepare(); const before = await state()
  for (const [axisId, section] of [['xAxis', 'horizontalAxis'], ['yAxis', 'verticalAxis'], ['rightYAxis', 'rightVerticalAxis']]) {
    await route({ tab: 'style', section, axisId }); await expect(field(`${axisId}.title`).getByRole('textbox')).toBeFocused()
    await route({ tab: 'style', section: 'gridlinesAndTicks', axisId })
    await expect(field(`${axisId}.gridLine.visible`).getByRole('switch')).toBeFocused()
  }
  const after = await state(); assert.deepEqual(after.charts, before.charts); assert.equal(after.revision, before.revision); assert.equal(after.history.undo, before.history.undo)
})

test('line/combination/area routes edit the real shared shape Select and retain series width controls', async () => {
  for (const type of ['line', 'combination', 'area']) {
    await mount({ nativeType: type }); const before = await state(); await route({ tab: 'style', section: 'lineAndArea' })
    const shape = field('area.lineStyle').getByRole('combobox'); await expect(shape).toBeFocused(); await shape.click()
    await page.locator('.ant-select-dropdown:visible').getByText('平滑', { exact: true }).click()
    if (type !== 'area') {
      const input = field('allSeriesStyle.border.width').getByRole('spinbutton'); await input.fill('8'); await input.blur()
    }
    await apply(); await page.waitForFunction(type => {
      const model = window.chartObjectFixture.state().charts.find(item => item.id === 'native-object')
      return model.options.area?.lineStyle === 'smooth' && (type === 'area' || model.options.allSeriesStyle?.border?.width === 8)
    }, type)
    const after = await state(); assert.equal(after.history.undo, before.history.undo + 1); await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
  }
})

test('field and exact source-name routes choose different real series scopes after a display rename', async () => {
  await mount({ multiSeries: true }); await prepare(); await route({ tab: 'style', section: 'series', seriesId: 'field:2' })
  const selector = drawer().getByRole('combobox', { name: '选择图表系列', exact: true })
  await expect(selector).toBeFocused(); await expect(drawer().locator('[data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', 'field:2')
  await field('series.2.name').getByRole('textbox').fill('Routed field label'); await apply()
  await page.waitForFunction(() => window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').options.series[2].name === 'Routed field label')
  await prepare(); await route({ tab: 'style', section: 'series', seriesId: 'name:"Comparison"' })
  await expect(drawer().locator('[data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', 'name:"Comparison"')
  await selector.click(); await page.locator('.ant-select-dropdown:visible').getByText('字段 · Value', { exact: true }).click()
  await expect(drawer().locator('[data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', 'field:1')
  await route({ tab: 'style', section: 'series', seriesId: 'name:"Comparison"' })
  await expect(drawer().locator('[data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', 'name:"Comparison"')
  await expect(selector).toBeFocused(); await field('namedSeries."Comparison".name').getByRole('textbox').fill('Routed name label'); await apply()
  await page.waitForFunction(() => window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').options.namedSeries?.some(entry => entry.name === 'Comparison' && entry.style.name === 'Routed name label'))
  const after = await state(); assert.equal(chart(after).options.series[2].name, 'Routed field label')
  await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(after))
})

test('same-chart route changes keep author drafts and later typing keeps focus until another navigation is requested', async () => {
  const before = await state(); await route({ tab: 'style', titleKey: 'title' }); const title = field('title').getByRole('textbox')
  await title.fill('Draft kept across routes'); await expect(title).toBeFocused(); await frames(); await expect(title).toBeFocused()
  await route({ tab: 'style', section: 'legend' }); await expect(field('legend.visible').getByRole('switch')).toBeFocused()
  await route({ tab: 'style', titleKey: 'title' }); await expect(title).toHaveValue('Draft kept across routes'); await expect(title).toBeFocused()
  await title.fill('Final draft'); await frames(); await expect(title).toBeFocused()
  assert.deepEqual((await state()).charts, before.charts); await apply()
  await page.waitForFunction(() => window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').title === 'Final draft')
  assert.equal((await state()).history.undo, before.history.undo + 1)
})

test('statistics routes focus Add without creating lines; explicit actions collect real trendline/error drafts in one commit', async () => {
  const before = await state(); await route({ tab: 'style', section: 'trendline' })
  await expect(drawer().getByRole('button', { name: '添加趋势线', exact: true })).toBeFocused()
  assert.equal(chart(await state()).options.trendlines, undefined)
  await drawer().getByRole('button', { name: '添加趋势线', exact: true }).click()
  await field('trendlines.0.name').getByRole('textbox').fill('Routed trend')
  await drawer().getByText('趋势线 1', { exact: true }).click(); await expect(field('trendlines.0.name')).toBeHidden()
  await route({ tab: 'style', section: 'trendline' }); await expect(field('trendlines.0.name').getByRole('textbox')).toHaveValue('Routed trend')
  await route({ tab: 'style', section: 'errorBars' }); await expect(drawer().getByRole('button', { name: '添加误差线', exact: true })).toBeFocused()
  await drawer().getByRole('button', { name: '添加误差线', exact: true }).click()
  await field('errorBars.0.value').getByRole('spinbutton').fill('2'); await field('errorBars.0.value').getByRole('spinbutton').blur()
  assert.deepEqual((await state()).charts, before.charts); await apply()
  await page.waitForFunction(() => window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').options.errorBars?.[0]?.value === 2)
  const after = await state(); assert.equal(chart(after).options.trendlines[0].name, 'Routed trend'); assert.equal(after.history.undo, before.history.undo + 1)
})

test('Pareto bar/line and Waterfall series routes expand their genuine existing editors and persist a real Pareto style edit', async () => {
  await mount({ nativeType: 'pareto' }); await prepare(); const before = await state()
  await route({ tab: 'style', section: 'paretoBarSeries' }); await expect(drawer().locator('[data-chart-pareto-group="barStyle"]')).toBeVisible()
  await route({ tab: 'style', section: 'paretoLineSeries' }); await expect(drawer().locator('[data-chart-pareto-group="cumulativeLineStyle"]')).toBeVisible()
  await drawer().getByText('帕累托累计线样式', { exact: true }).click(); await expect(drawer().locator('[data-chart-pareto-group="cumulativeLineStyle"]')).toBeHidden()
  await route({ tab: 'style', section: 'paretoLineSeries' }); await expect(drawer().locator('[data-chart-pareto-group="cumulativeLineStyle"]')).toBeVisible()
  await field('pareto.cumulativeLineStyle.width').getByRole('spinbutton').fill('6'); await field('pareto.cumulativeLineStyle.width').getByRole('spinbutton').blur(); await apply()
  await page.waitForFunction(() => window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').options.pareto?.cumulativeLineStyle?.width === 6)
  assert.equal((await state()).history.undo, before.history.undo + 1)
  await prepare('shared-object'); await route({ tab: 'style', section: 'waterfallSeries', seriesId: 'field:1' }, 'shared-object')
  await expect(drawer().getByRole('combobox', { name: '瀑布图角色样式作用系列', exact: true })).toBeFocused()
  await expect(drawer().locator('[data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', 'field:1')
  await expect(drawer().getByRole('combobox', { name: '瀑布图角色', exact: true })).toBeVisible()
})

test('invalid combinations, absent axes/series and readonly routes fail without opening a fallback editor or changing authors', async () => {
  await prepare(); const before = await state()
  for (const value of [{ tab: 'data', section: 'legend' }, { tab: 'style', section: 'pieStyle' }, { tab: 'style', axisId: 'rightYAxis' },
    { tab: 'style', seriesId: 'field:9' }, { tab: 'style', section: 'legend', privateField: true }]) {
    const reason = await page.evaluate(async value => { try { await window.chartObjectFixture.openPanel('native-object', value); return '' } catch (error) { return error.message } }, value)
    assert.ok(reason); await expect(drawer()).toHaveCount(0)
  }
  assert.deepEqual((await state()).charts, before.charts); assert.equal((await state()).history.undo, before.history.undo)
  await mount({ readonly: true }); const guarded = await state()
  const reason = await page.evaluate(async () => { try { await window.chartObjectFixture.openPanel('native-object', { tab: 'data' }); return '' } catch (error) { return error.message } })
  assert.match(reason, /保护|编辑/); await expect(drawer()).toHaveCount(0); assert.deepEqual((await state()).charts, guarded.charts)
})

test('an API route during IME composition leaves the input mounted and defers the tab/focus change until composition commits', async () => {
  const before = await state(); await route({ tab: 'style', titleKey: 'title' }); const title = field('title').getByRole('textbox')
  await title.evaluate(node => { node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })); node.value = '输入法草稿'
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '稿', isComposing: true })) })
  await route({ tab: 'data' }); await expect(drawer().getByRole('tab', { name: '样式', exact: true })).toHaveAttribute('aria-selected', 'true')
  await expect(title).toBeFocused(); await expect(title).toHaveValue('输入法草稿')
  await title.evaluate(node => { node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '输入法草稿' }))
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '输入法草稿' })) })
  await expect(drawer().getByRole('tab', { name: '数据', exact: true })).toHaveAttribute('aria-selected', 'true')
  await expect(drawer().getByRole('textbox', { name: '图表数据区域', exact: true })).toBeFocused()
  await route({ tab: 'style', titleKey: 'title' }); await expect(field('title').getByRole('textbox')).toHaveValue('输入法草稿')
  assert.deepEqual((await state()).charts, before.charts); assert.equal((await state()).history.undo, before.history.undo)
})

test('queued navigation after a sheet switch or replacement Studio cannot focus the new workbench', async () => {
  await page.evaluate(async () => { const pending = window.chartObjectFixture.openPanel('native-object', { tab: 'style', titleKey: 'subtitle' })
    await window.chartObjectFixture.switchSheet('destination'); await pending })
  await grid().focus(); await frames(); await expect(grid()).toBeFocused(); await expect(drawer()).toHaveCount(0)
  await mount({})
  await page.evaluate(async () => { const pending = window.chartObjectFixture.openPanel('native-object', { tab: 'style', titleKey: 'subtitle' })
    await window.chartObjectFixture.reopen(); await pending })
  await grid().focus(); await frames(); await expect(grid()).toBeFocused(); await expect(drawer()).toHaveCount(0)
})

test('a competing Ant Modal keeps focus during route navigation, and popup theme changes leave authors untouched', async () => {
  const before = await state(); await route({ tab: 'style', titleKey: 'title' })
  const light = await drawer().locator('.ant-drawer-content').evaluate(node => getComputedStyle(node).backgroundColor)
  await page.evaluate(() => window.chartObjectFixture.otherDialog(true)); const input = page.getByRole('textbox', { name: '临时对话框输入', exact: true })
  await input.fill('Modal draft'); await route({ tab: 'style', section: 'legend' }); await frames(); await expect(input).toBeFocused(); await expect(input).toHaveValue('Modal draft')
  await page.evaluate(() => window.chartObjectFixture.appearance(true, 'grid'))
  const dark = await drawer().locator('.ant-drawer-content').evaluate(node => getComputedStyle(node).backgroundColor); assert.notEqual(dark, light)
  const after = await state(); assert.deepEqual(after.charts, before.charts); assert.equal(after.history.undo, before.history.undo)
  await page.evaluate(() => window.chartObjectFixture.otherDialog(false)); await drawer().getByRole('button', { name: '关闭', exact: true }).click()
  await expect(grid()).toBeFocused()
})
