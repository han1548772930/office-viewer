import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// Written only. No server, browser, type check, screenshot or benchmark ran.
let server, browser, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
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
const menu = () => page.locator('.chart-title-float-menu')
async function clickTitle(id = 'native-title', kind = 'title', double = false) {
  const point = await page.evaluate(({ id, kind }) => window.chartTitleInteractionFixture.browserPoint(id, kind), { id, kind })
  await page.mouse.click(point.x, point.y, { clickCount: double ? 2 : 1 })
  await menu().waitFor({ state: 'visible' })
  assert.equal(await menu().getAttribute('data-title-kind'), kind)
  return point
}
async function waitUpdates(count) {
  await page.waitForFunction(count => window.chartTitleInteractionFixture.state().updates === count && !document.querySelector('.chart-hitbox[aria-busy="true"]'), count)
  return page.evaluate(() => window.chartTitleInteractionFixture.state())
}

test('real title/subtitle hits select Canvas outlines and Ant fields; text plus a formatting action commit once through the author command', async () => {
  const before = await page.evaluate(() => window.chartTitleInteractionFixture.state()), point = await page.evaluate(() => window.chartTitleInteractionFixture.browserPoint('native-title'))
  await page.mouse.move(point.x, point.y)
  await page.waitForFunction(() => window.chartTitleInteractionFixture.state().outlines > 0)
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).author, before.author)
  await clickTitle('native-title', 'title', true)
  const input = menu().getByRole('textbox', { name: '编辑图表标题' })
  assert.equal(await input.evaluate(node => node === document.activeElement), true)
  assert.deepEqual(await input.evaluate(node => [node.selectionStart, node.selectionEnd]), [0, 4])
  await input.fill('Main changed'); assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).updates, 0)
  await menu().getByRole('button', { name: '标题粗体', exact: true }).click()
  const authored = await waitUpdates(1), chart = authored.charts.find(chart => chart.id === 'native-title')
  assert.equal(chart.title, 'Main changed'); assert.equal(chart.options.titleStyle.bold, false)
  assert.equal(authored.revision, before.revision + 1); assert.equal(authored.fallbackUpdates, 0); assert.equal(authored.removals, 0)
  const size = menu().getByRole('spinbutton', { name: '标题字号', exact: true })
  await size.click()
  await size.press('ControlOrMeta+A'); await size.pressSequentially('24')
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).updates, 1)
  await size.press('Enter'); assert.equal((await waitUpdates(2)).charts[0].options.titleStyle.fontSize, 24)
  await clickTitle('shared-title', 'subtitle')
  assert.equal(await menu().getByRole('combobox', { name: '标题位置', exact: true }).count(), 0)
  assert.equal(await menu().getByRole('combobox', { name: '标题对齐', exact: true }).count(), 0)
  await menu().getByRole('textbox', { name: '编辑图表副标题' }).fill('Sub changed')
  await menu().getByRole('button', { name: '应用修改', exact: true }).click()
  const after = await waitUpdates(3)
  assert.equal(after.charts.find(chart => chart.id === 'shared-title').title, 'Main')
  assert.equal(after.charts.find(chart => chart.id === 'shared-title').subtitle, 'Sub changed')
})

test('IME composition retains draft text until one commit; Escape cancels new draft text and restores the chart focus', async () => {
  const before = await page.evaluate(() => window.chartTitleInteractionFixture.state())
  await clickTitle('native-title', 'title', true)
  const input = menu().getByRole('textbox', { name: '编辑图表标题' })
  await input.dispatchEvent('compositionstart', { data: '' })
  await input.fill('中文标题')
  await input.dispatchEvent('keydown', { key: 'Enter', code: 'Enter', isComposing: true, keyCode: 229 })
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).updates, 0)
  await input.dispatchEvent('compositionend', { data: '中文标题' })
  await input.press('Enter')
  const authored = await waitUpdates(1)
  assert.equal(authored.charts.find(chart => chart.id === 'native-title').title, '中文标题'); assert.equal(authored.revision, before.revision + 1)
  await input.fill('取消这段文字'); await input.press('Escape'); await menu().waitFor({ state: 'hidden' })
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).author, authored.author)
  assert.equal(await page.locator('.chart-hitbox[data-chart-id="native-title"]').evaluate(node => node === document.activeElement), true)
})

test('title dragging previews without author changes, commits once on pointerup and survives undo/redo/JSON reopen with independent automatic restoration', async () => {
  const before = await page.evaluate(() => window.chartTitleInteractionFixture.state()), point = await clickTitle()
  const oldInput = menu().getByRole('textbox', { name: '编辑图表标题' }), oldHandle = await oldInput.elementHandle()
  await oldInput.fill('Draft ended by title drag')
  await page.mouse.move(point.x, point.y); await page.mouse.down(); await page.mouse.move(point.x + 60, point.y + 40, { steps: 6 })
  await oldHandle.evaluate(node => node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', code: 'Enter' })))
  await oldHandle.dispose()
  const preview = await page.evaluate(() => window.chartTitleInteractionFixture.state())
  assert.equal(preview.author, before.author); assert.equal(preview.updates, 0); assert.equal(preview.revision, before.revision)
  await page.mouse.up(); const authored = await waitUpdates(1), chart = authored.charts.find(chart => chart.id === 'native-title')
  assert.ok(chart.options.titleStyle.manualLayout.left > 0); assert.ok(chart.options.titleStyle.manualLayout.top > 0)
  assert.equal(chart.title, 'Main')
  assert.deepEqual(chart.position, before.charts[0].position); assert.deepEqual(chart.size, before.charts[0].size)
  assert.deepEqual(chart.options.subtitleStyle, before.charts[0].options.subtitleStyle)
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.undo())).author, before.author)
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.redo())).author, authored.author)
  const reopened = await page.evaluate(() => window.chartTitleInteractionFixture.reopen())
  assert.deepEqual(reopened.charts[0].options.titleStyle, chart.options.titleStyle)
  await clickTitle(); await menu().getByRole('button', { name: '恢复标题自动位置', exact: true }).click()
  const automatic = await waitUpdates(2)
  assert.equal(automatic.charts[0].options.titleStyle.manualLayout, undefined); assert.equal(automatic.charts[0].title, 'Main')
  assert.equal(automatic.charts[0].options.titleStyle.fontSize, 18); assert.deepEqual(automatic.charts[0].options.subtitleStyle, chart.options.subtitleStyle)
})

test('readonly title selection stays inspectable and author-free; permission loss and a public command veto cancel draft/pending edits', async () => {
  const readonly = await page.evaluate(() => window.chartTitleInteractionFixture.mount({ readonly: true }))
  await clickTitle(); assert.equal(await menu().getByRole('textbox', { name: '编辑图表标题' }).isDisabled(), true)
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).author, readonly.author)
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).updates, 0)
  const before = await page.evaluate(() => window.chartTitleInteractionFixture.mount()), point = await clickTitle()
  await page.mouse.move(point.x, point.y); await page.mouse.down(); await page.mouse.move(point.x + 50, point.y + 30)
  await page.evaluate(() => window.chartTitleInteractionFixture.setReadonly(true)); await page.mouse.up()
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).author, before.author)
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).updates, 0)
  await page.evaluate(() => window.chartTitleInteractionFixture.setReadonly(false)); await clickTitle('native-title', 'title', true)
  await menu().getByRole('textbox', { name: '编辑图表标题' }).fill('Rejected')
  await page.evaluate(() => window.chartTitleInteractionFixture.rejectNextUpdate())
  await menu().getByRole('button', { name: '应用修改', exact: true }).click(); await menu().getByRole('alert').waitFor({ state: 'visible' })
  const rejected = await page.evaluate(() => window.chartTitleInteractionFixture.state())
  assert.equal(rejected.author, before.author); assert.equal(rejected.revision, before.revision); assert.equal(rejected.updates, 0)
})

test('zoom/theme/scroll/Escape/blur and host remount cancel active title state and prevent stale input from editing reused IDs', async () => {
  const before = await page.evaluate(() => window.chartTitleInteractionFixture.state())
  await clickTitle('native-title', 'title', true)
  await menu().getByRole('textbox', { name: '编辑图表标题' }).fill('Old host draft')
  await page.evaluate(() => window.chartTitleInteractionFixture.setDark(true)); await menu().waitFor({ state: 'hidden' })
  await clickTitle(); await page.evaluate(() => window.chartTitleInteractionFixture.setZoom(125)); await menu().waitFor({ state: 'hidden' })
  await page.waitForFunction(() => window.chartTitleInteractionFixture.state().charts.every(chart => chart.width === chart.expectedWidth))
  await clickTitle(); await page.evaluate(() => window.chartTitleInteractionFixture.scroll(20)); await menu().waitFor({ state: 'hidden' })
  await clickTitle(); await page.keyboard.press('Escape'); await menu().waitFor({ state: 'hidden' })
  await clickTitle(); await page.evaluate(() => window.dispatchEvent(new Event('blur'))); await menu().waitFor({ state: 'hidden' })
  assert.equal((await page.evaluate(() => window.chartTitleInteractionFixture.state())).author, before.author)
  await clickTitle('native-title', 'title', true)
  const old = menu().getByRole('textbox', { name: '编辑图表标题' })
  const oldHandle = await old.elementHandle()
  await old.dispatchEvent('compositionstart', { data: '' }); await old.fill('Stale IME draft')
  const released = await page.evaluate(() => window.chartTitleInteractionFixture.release())
  for (const surface of released) { assert.equal(surface.disposed, true); assert.deepEqual(surface.size, [1, 1]) }
  const next = await page.evaluate(() => window.chartTitleInteractionFixture.mount())
  await clickTitle('shared-title', 'title', true)
  await oldHandle.evaluate(node => {
    node.value = 'Late old host text'
    node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'Late old host text' }))
    node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', code: 'Enter' }))
  })
  await oldHandle.dispose()
  assert.equal(await menu().getByRole('textbox', { name: '编辑图表标题' }).inputValue(), 'Main')
  const current = await page.evaluate(() => window.chartTitleInteractionFixture.state())
  assert.equal(next.author, before.author); assert.equal(current.author, next.author); assert.equal(current.updates, 0)
})

test('side/multiline/rotated title hits stay tied to real polygons after shared content and native legend rebuilds', async () => {
  const before = await page.evaluate(() => window.chartTitleInteractionFixture.mount({ position: 'left', rotate: 25, multiline: true }))
  for (const id of ['native-title', 'shared-title']) {
    await clickTitle(id); assert.equal(await menu().getByRole('textbox', { name: '编辑图表标题' }).inputValue(), 'Main\nSecond')
    await menu().getByRole('button', { name: '关闭标题工具栏', exact: true }).click()
    const point = await page.evaluate(id => window.chartTitleInteractionFixture.legendPoint(id), id)
    await page.mouse.click(point.x, point.y)
    await clickTitle(id, 'subtitle'); assert.equal(await menu().getByRole('textbox', { name: '编辑图表副标题' }).inputValue(), 'Sub')
    await menu().getByRole('button', { name: '关闭标题工具栏', exact: true }).click()
  }
  const after = await page.evaluate(() => window.chartTitleInteractionFixture.state())
  assert.equal(after.author, before.author); assert.equal(after.revision, before.revision); assert.equal(after.updates, 0)
  assert.equal(after.charts.every(chart => chart.target.polygons.length > 1 && chart.subtitleTarget.polygons.length > 0), true)
})
