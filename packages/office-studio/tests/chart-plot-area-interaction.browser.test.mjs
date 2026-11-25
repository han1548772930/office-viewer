import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// Written only. No Vite server, browser, type check or benchmark ran in this batch.
// The shared fixture uses real ChartOverlay/Ant/ECharts/Canvas2D/Facade. Its
// prepared paint data does not replace the author command's Worker/WASM path.
let server, browser, page, baseURL
const errors = [], manual = { left: 0.2, top: 0.22, width: 0.55, height: 0.5 }
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} ≈ ${expected}`)
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/chart-title-interaction.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; page = await browser.newPage({ viewport: { width: 1440, height: 900 } }); page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL); await page.waitForFunction(() => window.chartTitleInteractionFixture?.state().charts.length === 2)
})
afterEach(async () => { await page?.evaluate(() => window.chartTitleInteractionFixture?.dispose()); await page?.close(); assert.deepEqual(errors, []) })
afterAll(async () => { await browser?.close(); await server?.close() })
const menu = () => page.locator('.chart-plot-area-float-menu')
const hitbox = (id = 'native-title') => page.locator(`.chart-hitbox[data-chart-id="${id}"]`)
const object = (id = 'native-title') => page.locator('.chart-object').filter({ has: hitbox(id) })
const number = label => menu().getByRole('spinbutton', { name: `绘图区${label} (%)`, exact: true })
const state = () => page.evaluate(() => window.chartTitleInteractionFixture.state())
async function mount(options = { plotArea: manual }) { return page.evaluate(options => window.chartTitleInteractionFixture.mount(options), options) }
async function selectPlot(id = 'native-title', entry = false) {
  if (entry) { await hitbox(id).focus(); await hitbox(id).press('Enter'); await object(id).getByRole('button', { name: '编辑图表绘图区', exact: true }).click() }
  else {
    const point = await page.evaluate(id => window.chartTitleInteractionFixture.plotPoint(id), id)
    await page.mouse.click(point.x, point.y)
  }
  await menu().waitFor({ state: 'visible' })
  assert.equal(await hitbox(id).getAttribute('data-plot-area-selected'), 'true')
}
async function waitUpdates(count) {
  await page.waitForFunction(count => {
    const state = window.chartTitleInteractionFixture.state()
    return state.updates === count && !document.querySelector('.chart-hitbox[aria-busy="true"]') && state.charts.every(chart => {
      const author = chart.options.plotArea, target = chart.plotTarget
      return target && target.manual === !!author && (!author || ['left', 'top', 'width', 'height'].every(key => Math.abs(target.area[key] - author[key]) < 1e-9))
    })
  }, count)
  return state()
}
async function apply() { await menu().getByRole('button', { name: '应用修改', exact: true }).click() }

test('plot-edge hover/selection and the contextual Ant entry reveal a real rectangle; four draft coordinates commit once for native and shared charts', async () => {
  const before = await state(), point = await page.evaluate(() => window.chartTitleInteractionFixture.plotPoint('native-title'))
  await page.mouse.move(point.x, point.y)
  await page.waitForFunction(() => window.chartTitleInteractionFixture.state().outlines >= 4)
  assert.equal((await state()).author, before.author)
  for (const [index, id] of ['native-title', 'shared-title'].entries()) {
    await selectPlot(id, index === 1)
    assert.equal(await object(id).locator('[data-plot-area-handle]').count(), 8)
    await menu().getByRole('switch', { name: '绘图区手动布局', exact: true }).click()
    await number('左侧位置').fill('15'); await number('顶部位置').fill('20'); await number('宽度').fill('65'); await number('高度').fill('60')
    assert.equal((await state()).updates, index)
    await apply(); const after = await waitUpdates(index + 1), chart = after.charts.find(chart => chart.id === id)
    assert.deepEqual(chart.options.plotArea, { left: 0.15, top: 0.2, width: 0.65, height: 0.6 })
    near(chart.plotTarget.plot.x, chart.size.width * 0.15); near(chart.plotTarget.plot.width, chart.size.width * 0.65)
    const original = before.charts.find(chart => chart.id === id)
    assert.deepEqual(chart.position, original.position); assert.deepEqual(chart.size, original.size)
    assert.deepEqual(chart.options.titleStyle, original.options.titleStyle); assert.deepEqual(chart.options.subtitleStyle, original.options.subtitleStyle)
    assert.equal(after.fallbackUpdates, 0); assert.equal(after.removals, 0)
  }
  assert.equal((await state()).revision, before.revision + 2)
})

test('a selected plot body previews a move without changing author/history; pointerup commits once and undo/redo/JSON reopening retain the full rectangle', async () => {
  const before = await mount(); await selectPlot()
  const old = await number('左侧位置').elementHandle(); await number('左侧位置').fill('12')
  const point = await page.evaluate(() => window.chartTitleInteractionFixture.plotPoint('native-title', true))
  await page.mouse.move(point.x, point.y); await page.mouse.down(); await page.mouse.move(point.x + 50, point.y + 25, { steps: 5 })
  await menu().waitFor({ state: 'hidden' })
  await page.waitForFunction(() => window.chartTitleInteractionFixture.state().charts[0].plotTarget.area.left > 0.2)
  await old.evaluate(node => node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', code: 'Enter' })))
  await old.dispose()
  const preview = await state()
  assert.equal(preview.author, before.author); assert.equal(preview.updates, 0); assert.equal(preview.revision, before.revision)
  near(preview.charts[0].plotTarget.area.left, manual.left + 50 / before.charts[0].size.width)
  await page.mouse.up(); const after = await waitUpdates(1), chart = after.charts[0]
  near(chart.options.plotArea.left, manual.left + 50 / chart.size.width); near(chart.options.plotArea.top, manual.top + 25 / chart.size.height)
  assert.equal(chart.options.plotArea.width, manual.width); assert.equal(chart.options.plotArea.height, manual.height)
  assert.deepEqual(chart.position, before.charts[0].position); assert.deepEqual(chart.size, before.charts[0].size)
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.undo())).author, before.author)
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.redo())).author, after.author)
  const reopened = await page.evaluate(() => window.chartTitleInteractionFixture.reopen())
  assert.deepEqual(reopened.charts[0].options.plotArea, chart.options.plotArea)
  assert.deepEqual(reopened.charts[0].options.titleStyle, chart.options.titleStyle)
})

test('all eight Ant handles fix opposite edges and publish exactly one full rectangle after the drag', async () => {
  for (const [handle, dx, dy] of [['nw', -20, -15], ['n', 0, -15], ['ne', 20, -15], ['e', 20, 0],
    ['se', 20, 15], ['s', 0, 15], ['sw', -20, 15], ['w', -20, 0]]) {
    const before = await mount(); await selectPlot()
    const button = object().locator(`[data-plot-area-handle="${handle}"]`), bounds = await button.boundingBox()
    assert.ok(bounds); const x = bounds.x + bounds.width / 2, y = bounds.y + bounds.height / 2
    await page.mouse.move(x, y); await page.mouse.down(); await page.mouse.move(x + dx, y + dy, { steps: 4 })
    assert.equal((await state()).author, before.author); assert.equal((await state()).updates, 0)
    await page.mouse.up(); const after = await waitUpdates(1), next = after.charts[0].options.plotArea
    if (handle.includes('w')) near(next.left + next.width, manual.left + manual.width); else near(next.left, manual.left)
    if (handle.includes('n')) near(next.top + next.height, manual.top + manual.height); else near(next.top, manual.top)
    assert.ok(next.width > 0 && next.height > 0 && next.left >= 0 && next.top >= 0 && next.left + next.width <= 1 && next.top + next.height <= 1)
    assert.deepEqual(after.charts[0].size, before.charts[0].size); assert.deepEqual(after.charts[0].position, before.charts[0].position)
    assert.equal(after.revision, before.revision + 1); assert.equal(after.fallbackUpdates, 0)
  }
})

test('invalid rectangles stay local; individual, position, size and entire-layout restoration use the same draft/apply author path', async () => {
  const before = await mount(); await selectPlot()
  await number('左侧位置').fill('90'); await number('宽度').fill('80')
  await menu().getByRole('alert').waitFor({ state: 'visible' })
  assert.equal(await menu().getByRole('button', { name: '应用修改', exact: true }).isDisabled(), true)
  assert.equal((await state()).author, before.author)
  await menu().getByRole('button', { name: '恢复绘图区宽度 (%)自动值', exact: true }).click()
  assert.equal(await menu().getByRole('button', { name: '应用修改', exact: true }).isDisabled(), true)
  await menu().getByRole('button', { name: '恢复绘图区左侧位置 (%)自动值', exact: true }).click()
  assert.equal((await state()).updates, 0); await apply()
  const partial = await waitUpdates(1), baseline = partial.charts[0].plotTarget.automatic
  near(partial.charts[0].options.plotArea.left, baseline.left); near(partial.charts[0].options.plotArea.width, baseline.width)
  assert.equal(partial.charts[0].options.plotArea.top, manual.top); assert.equal(partial.charts[0].options.plotArea.height, manual.height)
  await menu().getByRole('button', { name: '恢复绘图区自动位置', exact: true }).click()
  assert.equal((await state()).updates, 1); await apply(); const positioned = await waitUpdates(2)
  near(positioned.charts[0].options.plotArea.top, baseline.top); assert.equal(positioned.charts[0].options.plotArea.height, manual.height)
  await menu().getByRole('button', { name: '恢复绘图区自动大小', exact: true }).click()
  assert.equal((await state()).updates, 2); await apply(); const sized = await waitUpdates(3)
  near(sized.charts[0].options.plotArea.height, baseline.height)
  await menu().getByRole('button', { name: '恢复绘图区整组自动布局', exact: true }).click()
  assert.equal((await state()).updates, 3); await apply(); const automatic = await waitUpdates(4)
  assert.equal(automatic.charts[0].options.plotArea, undefined); assert.equal(automatic.charts[0].plotTarget.manual, false)
  assert.deepEqual(automatic.charts[0].options.titleStyle, before.charts[0].options.titleStyle)
  assert.deepEqual(automatic.charts[0].options.subtitleStyle, before.charts[0].options.subtitleStyle)
})

test('readonly, permission loss, public veto, view changes and a replaced host prevent old plot drafts or drags from changing author content', async () => {
  const readonly = await mount({ plotArea: manual, readonly: true }); await selectPlot()
  assert.equal(await number('左侧位置').isDisabled(), true)
  assert.equal(await object().locator('[data-plot-area-handle]:enabled').count(), 0)
  assert.equal((await state()).author, readonly.author); assert.equal((await state()).updates, 0)
  const before = await mount(); await selectPlot()
  const point = await page.evaluate(() => window.chartTitleInteractionFixture.plotPoint('native-title', true))
  await page.mouse.move(point.x, point.y); await page.mouse.down(); await page.mouse.move(point.x + 40, point.y + 20)
  await page.evaluate(() => window.chartTitleInteractionFixture.setReadonly(true)); await page.mouse.up()
  assert.equal((await state()).author, before.author); assert.equal((await state()).updates, 0)
  await page.evaluate(() => window.chartTitleInteractionFixture.setReadonly(false)); await selectPlot()
  await number('宽度').fill('60'); await page.evaluate(() => window.chartTitleInteractionFixture.rejectNextUpdate()); await apply()
  await menu().getByRole('alert').waitFor({ state: 'visible' })
  assert.equal((await state()).author, before.author); assert.equal((await state()).revision, before.revision)
  for (const [operation, value] of [['setDark', true], ['setZoom', 125], ['scroll', 20]]) {
    await selectPlot(); await number('宽度').fill('61')
    await page.evaluate(([operation, value]) => window.chartTitleInteractionFixture[operation](value), [operation, value])
    await menu().waitFor({ state: 'hidden' }); assert.equal((await state()).author, before.author)
  }
  await selectPlot(); await page.keyboard.press('Escape'); await menu().waitFor({ state: 'hidden' })
  await selectPlot(); await page.evaluate(() => window.dispatchEvent(new Event('blur'))); await menu().waitFor({ state: 'hidden' })
  await selectPlot(); const old = await number('左侧位置').elementHandle()
  await number('左侧位置').dispatchEvent('compositionstart', { data: '' })
  await page.evaluate(() => window.chartTitleInteractionFixture.release())
  const next = await mount(); await selectPlot('shared-title')
  await old.evaluate(node => {
    node.value = '21'; node.dispatchEvent(new Event('input', { bubbles: true }))
    node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '21' }))
    node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', code: 'Enter' }))
  })
  await old.dispose()
  assert.equal((await state()).author, next.author); assert.equal((await state()).updates, 0)
})

test('plot keyboard edits share history, Delete restores automatic layout, titles keep input priority and native/shared legend rebuilds retain plot geometry', async () => {
  const before = await mount(); await selectPlot()
  await hitbox().focus(); await hitbox().press('ArrowRight'); const moved = await waitUpdates(1)
  near(moved.charts[0].options.plotArea.left, manual.left + 1 / moved.charts[0].size.width)
  await hitbox().press('Shift+ArrowDown'); const lower = await waitUpdates(2)
  near(lower.charts[0].options.plotArea.top, manual.top + 10 / lower.charts[0].size.height)
  await object().locator('[data-plot-area-handle="e"]').focus(); await object().locator('[data-plot-area-handle="e"]').press('ArrowRight')
  const resized = await waitUpdates(3); near(resized.charts[0].options.plotArea.width, manual.width + 1 / resized.charts[0].size.width)
  await hitbox().focus(); await hitbox().press('Delete'); const automatic = await waitUpdates(4)
  assert.equal(automatic.charts[0].options.plotArea, undefined); assert.equal(automatic.removals, 0)
  assert.deepEqual(automatic.charts[0].options.titleStyle, before.charts[0].options.titleStyle)
  await menu().getByRole('switch', { name: '绘图区手动布局', exact: true }).click()
  await number('宽度').fill('62')
  const old = await number('左侧位置').elementHandle()
  const titlePoint = await page.evaluate(() => window.chartTitleInteractionFixture.browserPoint('native-title'))
  await page.mouse.click(titlePoint.x, titlePoint.y, { clickCount: 2 })
  await menu().waitFor({ state: 'hidden' }); await page.locator('.chart-title-float-menu').waitFor({ state: 'visible' })
  await old.evaluate(node => node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', code: 'Enter' }))); await old.dispose()
  assert.equal((await state()).updates, 4)
  await page.locator('.chart-title-float-menu').getByRole('textbox', { name: '编辑图表标题' }).fill('Edited title')
  await page.locator('.chart-title-float-menu').getByRole('button', { name: '应用修改', exact: true }).click(); const titled = await waitUpdates(5)
  assert.equal(titled.charts[0].title, 'Edited title'); assert.equal(titled.charts[0].options.plotArea, undefined)
  await page.locator('.chart-title-float-menu').getByRole('button', { name: '关闭标题工具栏', exact: true }).click()
  for (const id of ['native-title', 'shared-title']) {
    const point = await page.evaluate(id => window.chartTitleInteractionFixture.legendPoint(id), id)
    await page.mouse.click(point.x, point.y); await selectPlot(id)
    const selected = await state(), chart = selected.charts.find(chart => chart.id === id)
    assert.equal(chart.legendView.selected.Main, false)
    assert.equal(selected.updates, 5); assert.equal(selected.revision, titled.revision)
    assert.equal(await page.evaluate(id => 'officeChartPlotArea' in window.chartTitleInteractionFixture.surfaces.get(id).instance.getOption(), id), false)
    assert.ok(chart.plotTarget.bounds.width > 0 && chart.plotTarget.bounds.height > 0)
    await menu().getByRole('button', { name: '关闭绘图区工具栏', exact: true }).click()
  }
})
