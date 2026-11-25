import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// P18 deferred browser drafts: real ExcelWorkbench, Ant Select/Form/Drawer and Canvas2D.
// Startup code belongs to future validation; no server or browser ran in this batch.
let server, browser, context, page, baseURL
const errors = [], sample = [['Label', 'Value'], ['A', 9], ['B', 2], ['C', 8]]
const style = { autoGradientFill: false, allSeriesStyle: { color: '#123456', border: { width: 7 }, point: { shape: 'none' } } }
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No HTTP port was allocated')
  baseURL = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/chart-object-actions.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 900 } })
  page = await context.newPage(); page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await page.waitForFunction(() => window.chartObjectFixture?.state().charts.length === 2)
}, 60_000)
afterEach(async () => { try { await page?.evaluate(() => window.chartObjectFixture?.dispose()); assert.deepEqual(errors, []) } finally { await context?.close() } })
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.chartObjectFixture.state())
const paint = () => page.evaluate(() => window.chartObjectFixture.renderOption())
const scene = () => page.evaluate(() => window.chartObjectFixture.scene())
const data = () => page.evaluate(() => window.chartObjectFixture.data())
const chart = value => value.charts.find(item => item.id === 'native-object')
const cells = value => value.model.sheets.find(item => item.id === 'source').cells
const hitbox = () => page.locator('.chart-hitbox[data-chart-id="native-object"]')
const drawer = () => page.locator('.ant-drawer').filter({ has: page.getByText('图表设置', { exact: true }) })
const field = key => drawer().locator(`[data-chart-option-key=${JSON.stringify(key)}]`)
const lines = commands => commands.filter(item => item.kind === 'line' && item.width === 7)
const datum = value => value && typeof value === 'object' && !Array.isArray(value) ? value.value : value
async function mount(options = {}) {
  await page.evaluate(options => window.chartObjectFixture.mount(options), { rows: sample, nativeOptions: style, ...options })
  await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
  await page.evaluate(() => window.chartObjectFixture.prepare())
}
async function route(value = { tab: 'style', section: 'lineAndArea' }) {
  assert.equal(await page.evaluate(value => window.chartObjectFixture.openPanel('native-object', value), value), true)
  await expect(drawer()).toBeVisible()
}
async function select(key, label) {
  const control = field(key).getByRole('combobox'); await control.click()
  await page.locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}
async function apply() {
  const button = drawer().getByRole('button', { name: '应用', exact: true }); await expect(button).toBeEnabled(); await button.click()
}
async function waitShape(value) {
  await page.waitForFunction(value => window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').options.area?.lineStyle === value, value)
  await page.waitForFunction(value => window.chartObjectFixture.renderOption().key?.includes('"lineStyle":"' + value + '"'), value)
  await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
}
async function frames() { await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))) }

test('the genuine Line Select authors Smooth once, renders curved commands and preserves source data', async () => {
  await mount(); const before = await state(), originalData = await data()
  await route(); await expect(field('area.lineStyle').getByRole('combobox')).toBeFocused()
  await select('area.lineStyle', '平滑')
  assert.deepEqual((await state()).charts, before.charts); await apply(); await waitShape('smooth')
  await page.waitForFunction(() => window.chartObjectFixture.renderOption().series.some(item => item.type === 'line' && item.smooth === true))
  const after = await state(), option = await paint(), commands = lines(await scene())
  assert.equal(after.history.undo, before.history.undo + 1); assert.equal(option.series[0].smooth, true); assert.equal(option.series[0].step, false)
  assert.equal(commands.length, 24); assert.deepEqual(await data(), originalData); assert.deepEqual(cells(after), cells(before))
  await page.evaluate(() => window.chartObjectFixture.undo()); assert.equal(chart(await state()).options.area?.lineStyle, undefined)
  await page.evaluate(() => window.chartObjectFixture.redo()); await waitShape('smooth')
  await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(after))
})

test('Stacked/percent Area selects Step through Ant and keeps its real stacked series and numeric source', async () => {
  const input = [['Label', 'Lower', 'Upper'], ['A', 4, 2], ['B', 8, 3], ['C', 6, 5]]
  for (const type of ['areaStacked', 'areaPercentStacked']) {
    await mount({ nativeType: type, multiSeries: true, rows: input, nativeOptions: { ...style, series: { 2: { rightAxis: false } } } })
    const before = await state(); await route(); await select('area.lineStyle', '阶梯'); await apply(); await waitShape('step')
    await page.waitForFunction(() => window.chartObjectFixture.renderOption().series.filter(item => item.type === 'line').every(item => item.step === 'middle'))
    const option = await paint(), after = await state(), commands = await scene()
    assert.equal(after.history.undo, before.history.undo + 1); assert.equal(option.series.length, 2)
    assert.ok(option.series.every(item => item.stack && item.step === 'middle' && item.smooth === false))
    assert.ok(commands.some(item => item.kind === 'polygon'))
    for (const line of lines(commands)) assert.ok(line.x1 === line.x2 || line.y1 === line.y2)
    if (type === 'areaPercentStacked') for (let index = 0; index < 3; index++) {
      assert.ok(Math.abs(datum(option.series[0].data[index]) + datum(option.series[1].data[index]) - 100) < 1e-7)
    }
    assert.deepEqual(cells(after), cells(before)); assert.deepEqual(chart(after).source, chart(before).source)
  }
})

test('Combination uses the shared shape Select while keeping its column series and secondary axis untouched', async () => {
  await mount({ nativeType: 'combination', multiSeries: true,
    rows: [['Label', 'Bars', 'Line'], ['A', 4, 9], ['B', 7, 2], ['C', 2, 8]] })
  const before = await state(), oldOption = await paint(); await route(); await select('area.lineStyle', '阶梯'); await apply(); await waitShape('step')
  await page.waitForFunction(() => window.chartObjectFixture.renderOption().series.some(item => item.type === 'line' && item.step === 'middle'))
  const option = await paint(), after = await state()
  assert.deepEqual(option.series[0], oldOption.series[0]); assert.equal(option.series[0].type, 'bar')
  assert.equal(option.series[1].step, 'middle'); assert.equal(option.series[1].yAxisIndex, oldOption.series[1].yAxisIndex)
  assert.deepEqual(chart(after).options.rightYAxis, chart(before).options.rightYAxis)
  assert.equal(after.history.undo, before.history.undo + 1); assert.deepEqual(cells(after), cells(before))
  assert.equal(lines(await scene()).length, 6)
})

test('empty-value Ant choices draw Break/Link/Zero correctly and Scatter exposes only the applicable empty-value control', async () => {
  await mount({ rows: [['Label', 'Value'], ['A', 9], ['B', null], ['C', 2], ['D', 8]], nativeOptions: { ...style, area: { lineStyle: 'smooth' } } })
  const before = await state(); await route()
  for (const [strategy, label, count] of [['break', '断开', 12], ['link', '连接', 24], ['zero', '按零值', 36]]) {
    await select('invalidValueStrategy', label); await apply()
    await page.waitForFunction(strategy => window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').options.invalidValueStrategy === strategy, strategy)
    await frames(); const option = await paint()
    assert.equal(option.series[0].connectNulls, strategy === 'link')
    assert.deepEqual(option.series[0].data.map(datum), strategy === 'zero' ? [9, 0, 2, 8] : [9, null, 2, 8])
    assert.equal(lines(await scene()).length, count); assert.deepEqual(cells(await state()), cells(before))
  }
  await mount({ nativeType: 'scatter', rows: [['X', 'Y'], [1, 6], [2, null], [null, 8], [4, 3]],
    nativeOptions: { autoGradientFill: false, allSeriesStyle: { point: { shape: 'circle' } } } })
  await route(); await expect(field('area.lineStyle')).toHaveCount(0)
  await expect(field('invalidValueStrategy').getByRole('combobox')).toBeFocused()
  await select('invalidValueStrategy', '按零值'); await apply(); await frames()
  const scatter = (await paint()).series[0]
  assert.equal(scatter.type, 'scatter'); assert.equal(scatter.smooth, undefined); assert.equal(scatter.step, undefined)
  assert.deepEqual(scatter.data.map(datum), [[1, 6], [2, 0], [4, 3]])
})

test('3D Area reset clears both legacy and new shape authors and survives actual command undo/redo and JSON reopen', async () => {
  await mount({ nativeType: 'area3D', nativeOptions: { ...style, area: { smooth: true, future: 'retained' } } })
  const before = await state(); await route(); await select('area.lineStyle', '阶梯'); await apply(); await waitShape('step')
  const stepped = await scene(); assert.ok((await paint()).graphicCount > 0)
  await drawer().getByRole('button', { name: '折线与面积线型恢复默认', exact: true }).click(); await apply()
  await page.waitForFunction(() => {
    const area = window.chartObjectFixture.state().charts.find(item => item.id === 'native-object').options.area
    return area.lineStyle === undefined && area.smooth === undefined
  })
  const after = await state(); assert.equal(after.history.undo, before.history.undo + 2)
  assert.equal(chart(after).options.area.future, 'retained'); assert.notDeepEqual(await scene(), stepped)
  await page.evaluate(() => window.chartObjectFixture.undo()); await waitShape('step')
  await page.evaluate(() => window.chartObjectFixture.redo()); await page.evaluate(() => window.chartObjectFixture.reopen())
  assert.deepEqual(chart(await state()), chart(after)); assert.equal(chart(await state()).options.area.smooth, undefined)
})

test('the genuine Select popup follows ConfigProvider themes without authoring and closing restores Canvas focus', async () => {
  await mount(); const before = await state(); await route()
  const control = field('area.lineStyle').getByRole('combobox'); await expect(control).toBeFocused(); await control.click()
  const popup = page.locator('.ant-select-dropdown:visible'); await expect(popup).toBeVisible()
  const light = await popup.evaluate(node => getComputedStyle(node).backgroundColor)
  await page.evaluate(() => window.chartObjectFixture.appearance(true, 'grid')); await frames()
  await expect(popup).toBeVisible(); const dark = await popup.evaluate(node => getComputedStyle(node).backgroundColor)
  assert.notEqual(dark, light); await control.press('Escape'); await expect(control).toBeFocused()
  const after = await state(); assert.deepEqual(after.charts, before.charts); assert.equal(after.history.undo, before.history.undo)
  await drawer().getByRole('button', { name: '关闭', exact: true }).click()
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeFocused()
})

test('IME composition remains mounted during lineAndArea navigation and then one Apply commits text and shape', async () => {
  await mount(); const before = await state(); await route({ tab: 'style', titleKey: 'title' })
  const title = field('title').getByRole('textbox'); await expect(title).toBeFocused()
  await title.evaluate(node => {
    node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })); node.value = '输入法曲线'
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '线', isComposing: true }))
  })
  await route(); await expect(title).toBeFocused(); await expect(title).toHaveValue('输入法曲线')
  assert.deepEqual((await state()).charts, before.charts)
  await title.evaluate(node => {
    node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '输入法曲线' }))
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '输入法曲线' }))
  })
  await expect(field('area.lineStyle').getByRole('combobox')).toBeFocused()
  await select('area.lineStyle', '平滑'); await apply(); await waitShape('smooth')
  const after = await state(); assert.equal(chart(after).title, '输入法曲线'); assert.equal(after.history.undo, before.history.undo + 1)
  assert.deepEqual(cells(after), cells(before)); assert.equal((await paint()).series[0].smooth, true)
})
