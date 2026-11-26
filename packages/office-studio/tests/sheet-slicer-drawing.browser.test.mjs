import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// Deferred real Workbench/Ant/Canvas2D/native transaction drafts. No fixture, browser or server has been started.
// Windows IME, WebGPU, independent Office files and large-sample performance remain separate acceptance work.
let server, browser, context, page, url
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No fixture port allocated')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/sheet-drawing-workbench.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext({ viewport: { width: 1440, height: 1000 } }); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(url); await page.waitForFunction(() => window.sheetDrawingFixture?.state().scene.objects.length === 3)
  await page.evaluate(() => window.sheetDrawingFixture.mount({ withSlicers: true }))
  await page.waitForFunction(() => window.sheetDrawingFixture.state().scene.objects.length === 5)
  await expect(page.locator('[data-slicer-id="s"] [data-slicer-control="item"]').first()).toBeVisible()
}, 60_000)
afterEach(async () => { try { await page?.evaluate(() => window.sheetDrawingFixture?.dispose()); assert.deepEqual(errors, []) } finally { await context?.close() } })
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.sheetDrawingFixture.state())
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const toolbar = () => page.getByRole('toolbar', { name: '绘图组合工具', exact: true })
const owner = (value, id) => value.model.slicers.find(control => control.id === id)
const geometry = (value, kind, id) => value.scene.objects.find(object => object.ref.kind === kind && object.ref.id === id).geometry
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < .02, `${actual} !== ${expected}`)
const mixed = [{ kind: 'image', id: 'a' }, { kind: 'chart', id: 'c' }, { kind: 'slicer', id: 's' }, { kind: 'timeline', id: 't' }]
async function controlPoint(id, kind, index = 0, ratios) {
  return page.evaluate(({ id, kind, index, ratios }) => window.sheetDrawingFixture.slicerControlPoint(id, kind, index, undefined, ratios), { id, kind, index, ratios })
}
async function clickControl(id, kind, index = 0, modifier) {
  const point = await controlPoint(id, kind, index)
  if (modifier) await page.keyboard.down(modifier)
  try { await page.mouse.click(point.x, point.y) } finally { if (modifier) await page.keyboard.up(modifier) }
}
async function waitHistory(expected) { await page.waitForFunction(expected => window.sheetDrawingFixture.state().history.undo === expected, expected) }
async function menuAction(button, label) {
  await toolbar().getByRole('button', { name: button, exact: true }).click()
  await page.getByRole('menuitem', { name: label, exact: true }).last().click()
}
async function select(refs) { await page.evaluate(refs => window.sheetDrawingFixture.select(refs), refs) }
async function drag(start, end, cancel = false) {
  await page.mouse.move(start.x, start.y); await page.mouse.down(); await page.mouse.move(end.x, end.y, { steps: 8 })
  if (cancel) await page.keyboard.press('Escape')
  await page.mouse.up()
}

test('Ant item pointer Ctrl and Shift apply filters without replacing the shared drawing selection', async () => {
  const before = await state()
  await clickControl('s', 'item', 0); await waitHistory(before.history.undo + 1)
  assert.deepEqual(owner(await state(), 's').selectedItems, ['North'])
  await clickControl('s', 'item', 1, 'Control'); await waitHistory(before.history.undo + 2)
  assert.deepEqual(owner(await state(), 's').selectedItems, ['North', 'South'])
  await clickControl('s', 'item', 3, 'Shift'); await waitHistory(before.history.undo + 3)
  const after = await state()
  assert.deepEqual(owner(after, 's').selectedItems, ['South', 'Region 03', 'Region 04'])
  assert.deepEqual(after.selection, [{ kind: 'slicer', id: 's' }])
  assert.deepEqual(geometry(after, 'slicer', 's'), geometry(before, 'slicer', 's'))
  assert.equal(after.commands.filter(command => command.id === 'sheet.command.set-slicer-items').length, 3)
})

test('outer keys move the slicer while Enter switches to content navigation and Escape restores outer keys', async () => {
  await select([{ kind: 'slicer', id: 's' }]); await grid().focus(); const before = await state()
  await grid().press('Shift+ArrowRight'); await waitHistory(before.history.undo + 1)
  const moved = await state(); near(geometry(moved, 'slicer', 's').matrix.e, geometry(before, 'slicer', 's').matrix.e + 10)
  assert.deepEqual(owner(moved, 's').selectedItems, owner(before, 's').selectedItems)
  await grid().press('Enter'); await grid().press('ArrowDown')
  assert.equal((await state()).history.undo, moved.history.undo)
  await grid().press('Space'); await waitHistory(moved.history.undo + 1)
  const filtered = await state(); assert.deepEqual(owner(filtered, 's').selectedItems, ['South'])
  assert.deepEqual(geometry(filtered, 'slicer', 's'), geometry(moved, 'slicer', 's'))
  await grid().press('Escape'); await grid().press('ArrowRight'); await waitHistory(filtered.history.undo + 1)
  near(geometry(await state(), 'slicer', 's').matrix.e, geometry(filtered, 'slicer', 's').matrix.e + 1)
})

test('Ant multi-select and keyboard item activation share native filter and clear history', async () => {
  await clickControl('s', 'item', 0); const before = await state()
  await clickControl('s', 'multiSelect'); assert.equal((await state()).history.undo, before.history.undo)
  const south = page.locator('[data-slicer-id="s"] [data-slicer-control="item"][data-slicer-index="1"]').first()
  await south.focus(); await south.press('Enter'); await waitHistory(before.history.undo + 1)
  assert.deepEqual(owner(await state(), 's').selectedItems, ['North', 'South'])
  await clickControl('s', 'clear'); await waitHistory(before.history.undo + 2)
  const cleared = await state(); assert.equal(owner(cleared, 's').selectedItems, undefined)
  await page.evaluate(() => window.sheetDrawingFixture.undo()); assert.deepEqual(owner(await state(), 's').selectedItems, ['North', 'South'])
})

test('a rotated timeline level menu commits one native entry and preserves affine geometry across undo and reopen', async () => {
  const beforeTransform = await state(), current = owner(beforeTransform, 't').drawingGeometry
  await page.evaluate(geometry => window.sheetDrawingFixture.presentation('t', { drawingGeometry: { ...geometry,
    matrix: { a: 0, b: 1, c: -1, d: 0, e: 850, f: 80 } } }), current)
  const before = await state(); await clickControl('t', 'level')
  await expect(page.locator('[data-slicer-popup="t"]')).toBeVisible()
  await page.getByRole('menuitem', { name: '年', exact: true }).last().click(); await waitHistory(before.history.undo + 1)
  const after = await state(); assert.equal(owner(after, 't').level, 'year')
  assert.deepEqual(geometry(after, 'timeline', 't'), geometry(before, 'timeline', 't')); await expect(grid()).toBeFocused()
  await page.evaluate(() => window.sheetDrawingFixture.undo()); assert.deepEqual((await state()).model, before.model)
  await page.evaluate(() => window.sheetDrawingFixture.redo()); await page.evaluate(() => window.sheetDrawingFixture.reopen())
  assert.deepEqual((await state()).model, after.model)
})

test('timeline popups inherit theme, close outside the Canvas and restore Canvas focus only for Escape', async () => {
  await clickControl('t', 'level'); const before = await state(), popup = page.locator('[data-slicer-popup="t"]')
  const light = await popup.getByRole('menu').evaluate(node => getComputedStyle(node).backgroundColor)
  const name = page.getByRole('textbox', { name: '名称框', exact: true }); await name.click()
  await expect(popup).toHaveCount(0); await expect(name).toBeFocused()
  await page.evaluate(() => window.sheetDrawingFixture.appearance(true, 'grid')); await clickControl('t', 'level')
  const dark = await page.locator('[data-slicer-popup="t"]').getByRole('menu').evaluate(node => getComputedStyle(node).backgroundColor)
  assert.notEqual(dark, light)
  await page.keyboard.press('Escape'); await expect(page.locator('[data-slicer-popup="t"]')).toHaveCount(0)
  await expect(grid()).toBeFocused(); assert.deepEqual((await state()).model, before.model)
})

test('Ant vertical Slider keeps keyboard focus and commits only after the key is released', async () => {
  const slider = page.getByRole('slider', { name: '滚动切片器', exact: true }), before = await state()
  await slider.focus(); await page.keyboard.down('End')
  await page.waitForFunction(() => Number(document.querySelector('[data-slicer-id="s"] [role="slider"]').getAttribute('aria-valuenow')) > 0)
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
  await page.keyboard.up('End'); await waitHistory(before.history.undo + 1)
  assert.ok(owner(await state(), 's').startItem > 0); await expect(slider).toBeFocused()
  await slider.press('Home'); await waitHistory(before.history.undo + 2)
  assert.equal(owner(await state(), 's').startItem, 0); await expect(slider).toBeFocused()
})

test('leaving a Slider before keyup cancels its preview and preserves the new text-input focus', async () => {
  const slider = page.getByRole('slider', { name: '滚动切片器', exact: true }), before = await state()
  await slider.focus(); await page.keyboard.down('End')
  await page.waitForFunction(() => Number(document.querySelector('[data-slicer-id="s"] [role="slider"]').getAttribute('aria-valuenow')) > 0)
  const name = page.getByRole('textbox', { name: '名称框', exact: true }); await name.click(); await page.keyboard.up('End')
  await expect(slider).toHaveAttribute('aria-valuenow', '0'); await expect(name).toBeFocused()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

test('affine Ant scrollbar pointer preview cancels on Escape and a completed drag commits once', async () => {
  const before = await state(), start = await controlPoint('s', 'thumb'), end = await controlPoint('s', 'scrollbar', 0, { x: .5, y: .85 })
  const scroll = page.locator('[data-slicer-id="s"] [data-slicer-control="scroll"]').first()
  await scroll.dispatchEvent('pointerdown', { button: 2, buttons: 2, pointerId: 51, pointerType: 'mouse', clientX: start.x, clientY: start.y })
  await scroll.dispatchEvent('pointerup', { button: 2, buttons: 0, pointerId: 51, pointerType: 'mouse', clientX: end.x, clientY: end.y })
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
  await page.mouse.move(start.x, start.y); await page.mouse.down(); await page.mouse.move(end.x, end.y, { steps: 6 })
  assert.deepEqual((await state()).model, before.model)
  await page.keyboard.press('Escape'); await page.mouse.up()
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
  await drag(start, end); await waitHistory(before.history.undo + 1)
  const after = await state(); assert.ok(owner(after, 's').startItem > 0)
  assert.deepEqual(geometry(after, 'slicer', 's'), geometry(before, 'slicer', 's'))
  assert.equal(after.commands.filter(command => command.id === 'sheet.command.update-slicer-presentation').length, 1)
})

test('timeline Slider stores a real date position from the multi-month Pivot and retains source bindings', async () => {
  const before = await state(), point = await controlPoint('t', 'scrollbar', 0, { x: .8, y: .5 })
  await page.mouse.click(point.x, point.y); await waitHistory(before.history.undo + 1)
  const after = await state(); assert.ok(owner(after, 't').scrollPosition > 0)
  assert.equal(owner(after, 't').pivotTableId, owner(before, 't').pivotTableId)
  assert.equal(owner(after, 't').sourceFieldId, owner(before, 't').sourceFieldId)
  assert.deepEqual(after.model.pivotTables, before.model.pivotTables)
  await page.evaluate(() => window.sheetDrawingFixture.undo()); assert.deepEqual((await state()).model, before.model)
})

for (const steering of ['sheet', 'user', 'model', 'viewport']) test(`a native item press cannot commit after its ${steering} context changes`, async () => {
  const point = await controlPoint('s', 'item', 0)
  await page.mouse.move(point.x, point.y); await page.mouse.down()
  if (steering === 'sheet') await page.evaluate(() => window.sheetDrawingFixture.switchSheet('other'))
  else if (steering === 'user') await page.evaluate(() => window.sheetDrawingFixture.switchUser('peer'))
  else if (steering === 'model') await page.evaluate(() => window.sheetDrawingFixture.presentation('s', { caption: 'Newer author' }))
  else await page.evaluate(() => window.sheetDrawingFixture.zoom(125))
  const afterSteering = await state(); await page.mouse.up()
  assert.deepEqual((await state()).model, afterSteering.model); assert.equal((await state()).history.undo, afterSteering.history.undo)
  assert.equal((await state()).commands.filter(command => command.id === 'sheet.command.set-slicer-items').length, 0)
})

test('a stale open timeline menu closes when the model changes and cannot replay its old grain command', async () => {
  await clickControl('t', 'level'); await expect(page.locator('[data-slicer-popup="t"]')).toBeVisible()
  await page.evaluate(() => window.sheetDrawingFixture.presentation('t', { caption: 'Newer timeline' }))
  const after = await state(); await expect(page.locator('[data-slicer-popup="t"]')).toHaveCount(0)
  await page.keyboard.press('Enter'); assert.deepEqual((await state()).model, after.model)
})

test('mixed image chart slicer and timeline grouping moves every owner once without filtering its grouped content', async () => {
  await select(mixed); await menuAction('组合或取消组合', '组合')
  const grouped = await state(); assert.equal(grouped.selection[0].kind, 'group')
  const start = await controlPoint('s', 'item', 0); await drag(start, { x: start.x + 45, y: start.y + 25 }); await waitHistory(grouped.history.undo + 1)
  const moved = await state()
  for (const ref of mixed) {
    const old = geometry(grouped, ref.kind, ref.id), next = geometry(moved, ref.kind, ref.id)
    near(next.matrix.e, old.matrix.e + 45); near(next.matrix.f, old.matrix.f + 25)
    for (const key of ['a', 'b', 'c', 'd']) near(next.matrix[key], old.matrix[key])
    assert.equal(next.width, old.width); assert.equal(next.height, old.height)
  }
  assert.deepEqual(moved.model.slicers.map(control => control.selectedItems), grouped.model.slicers.map(control => control.selectedItems))
  assert.deepEqual(moved.model.tables, grouped.model.tables); assert.deepEqual(moved.model.pivotTables, grouped.model.pivotTables)
  const nextStart = await controlPoint('s', 'item', 0); await drag(nextStart, { x: nextStart.x + 20, y: nextStart.y + 10 }, true)
  assert.deepEqual((await state()).model, moved.model)
  await page.evaluate(() => window.sheetDrawingFixture.undo()); assert.deepEqual((await state()).model, grouped.model)
})

test('reflected frozen copies expose one Tab target for each control and the Slider handle at zoom and scroll', async () => {
  const value = owner(await state(), 's').drawingGeometry
  await page.evaluate(value => window.sheetDrawingFixture.presentation('s', { drawingGeometry: { ...value,
    matrix: { a: -1, b: .2, c: .15, d: 1, e: 360, f: 60 } } }), value)
  await page.evaluate(() => window.sheetDrawingFixture.freeze(5, 2)); await page.evaluate(() => window.sheetDrawingFixture.zoom(125))
  await page.evaluate(() => window.sheetDrawingFixture.scroll(30, 20))
  const counts = await page.locator('[data-slicer-id="s"] button[data-slicer-control]').evaluateAll(nodes => {
    const result = {}
    for (const node of nodes) if (node.tabIndex >= 0 && node.getAttribute('aria-hidden') !== 'true') {
      const key = `${node.dataset.slicerControl}:${node.dataset.slicerIndex ?? ''}`; result[key] = (result[key] ?? 0) + 1
    }
    return result
  })
  assert.ok(Object.keys(counts).length); assert.ok(Object.values(counts).every(count => count === 1))
  const sliders = await page.locator('[data-slicer-id="s"] [role="slider"]').evaluateAll(nodes => nodes.filter(node => node.tabIndex >= 0).length)
  assert.equal(sliders, 1)
  const before = await state(); await clickControl('s', 'item', 0); await waitHistory(before.history.undo + 1)
  assert.deepEqual(owner(await state(), 's').selectedItems, ['North'])
})

test('mixed deletion removes native filters in one history while preserving Table and Pivot sources for undo', async () => {
  await select(mixed); await grid().focus(); const before = await state()
  await grid().press('Delete'); await waitHistory(before.history.undo + 1)
  const removed = await state(); assert.deepEqual(removed.model.slicers, [])
  assert.deepEqual(removed.model.tables, before.model.tables); assert.deepEqual(removed.model.pivotTables, before.model.pivotTables)
  await page.evaluate(() => window.sheetDrawingFixture.undo()); assert.deepEqual((await state()).model, before.model)
  assert.deepEqual((await state()).selection, mixed)
})

test('read-only Ant filter controls and shared transform keys leave native content and history unchanged', async () => {
  await page.evaluate(() => window.sheetDrawingFixture.mount({ withSlicers: true, readonly: true }))
  await select([{ kind: 'slicer', id: 's' }]); const before = await state()
  await expect(page.locator('[data-slicer-id="s"] [data-slicer-control="item"]').first()).toBeDisabled()
  await expect(page.locator('[data-slicer-id="t"] [data-slicer-control="level"]').first()).toBeDisabled()
  await grid().focus(); await grid().press('ArrowRight'); await grid().press('Delete')
  assert.deepEqual((await state()).model, before.model); assert.equal((await state()).history.undo, before.history.undo)
})

for (const kind of ['slicer', 'timeline']) test(`${kind} authored quick styles use the live Workbench model and persist one command through undo and reopen`, async () => {
  const source = (await state()).model, id = kind === 'slicer' ? 's' : 't', name = 'Authored ribbon default'
  source[kind === 'slicer' ? 'slicerStyles' : 'timelineStyles'] = {
    [kind === 'slicer' ? 'defaultSlicerStyle' : 'defaultTimelineStyle']: name,
    styles: [{ name, wholeStyle: { bg: { rgb: '#102030' } }, [kind === 'slicer' ? 'slicerStyleElements' : 'timelineStyleElements']: [] }],
  }
  delete source.slicers.find(control => control.id === id).style
  await page.evaluate(model => window.sheetDrawingFixture.mount({ model }), source)
  await select([{ kind, id }])
  const trigger = page.locator('[data-ribbon-item="sheet.menu.slicer-quick-style"]:visible button[aria-haspopup]')
  await expect(trigger).toBeEnabled()
  await trigger.click()
  const gallery = page.locator('.slicer-quick-styles:visible')
  await expect(gallery.getByRole('button', { name, exact: true })).toHaveAttribute('aria-pressed', 'true')
  await expect(gallery.getByRole('button', { name, exact: true }).locator('.slicer-quick-sample')).toHaveCSS('background-color', 'rgb(16, 32, 48)')
  const before = await state(), builtin = kind === 'slicer' ? 'SlicerStyleDark2' : 'TimelineStyleDark2'
  await gallery.getByRole('button', { name: builtin, exact: true }).click()
  await waitHistory(before.history.undo + 1)
  const styled = await state()
  assert.equal(owner(styled, id).style, builtin)
  assert.deepEqual(geometry(styled, kind, id), geometry(before, kind, id))
  assert.equal(styled.commands.filter(command => command.id === 'sheet.command.update-slicer-presentation').length, 1)
  await page.evaluate(() => window.sheetDrawingFixture.undo())
  assert.deepEqual((await state()).model, before.model)
  await page.evaluate(() => window.sheetDrawingFixture.redo())
  await page.evaluate(() => window.sheetDrawingFixture.reopen())
  assert.deepEqual((await state()).model, styled.model)
  await select([{ kind, id }])
  await trigger.click()
  await expect(gallery.getByRole('button', { name: builtin, exact: true })).toHaveAttribute('aria-pressed', 'true')
  const reopened = await state()
  await gallery.getByRole('button', { name: '文档默认值', exact: true }).click()
  await waitHistory(reopened.history.undo + 1)
  assert.equal(owner(await state(), id).style, undefined)
  await trigger.click()
  await expect(gallery.getByRole('button', { name, exact: true })).toHaveAttribute('aria-pressed', 'true')
  await page.keyboard.press('Escape')
  await page.evaluate(() => window.sheetDrawingFixture.undo())
  assert.deepEqual((await state()).model, reopened.model)
})
