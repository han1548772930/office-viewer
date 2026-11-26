import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'
import { drawingAffineFrame } from '../src/domain/drawing-affine.ts'

// Deferred real Workbench/Ant/Canvas2D/ECharts/native history drafts. No server or browser has been launched.
// Synthetic composition guards below do not replace Windows native IME, WebGPU or Office acceptance.
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
  await page.waitForFunction(() => document.querySelector('.chart-hitbox[data-chart-id="c"]')?.dataset.rendered === 'true')
}, 60_000)
afterEach(async () => { try { await page?.evaluate(() => window.sheetDrawingFixture?.dispose()); assert.deepEqual(errors, []) } finally { await context?.close() } })
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.sheetDrawingFixture.state())
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const toolbar = () => page.getByRole('toolbar', { name: '绘图组合工具', exact: true })
const key = (kind, id) => ({ kind, id })
const mixed = [key('image', 'a'), key('chart', 'c')]
const keys = selection => selection.map(ref => `${ref.kind}:${ref.id}`).sort()
const object = (value, kind, id) => value.scene.objects.find(object => object.ref.kind === kind && object.ref.id === id)
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 0.02, `${actual} !== ${expected}`)
async function clickRef(kind, id, modifier, ratios) {
  const point = await page.evaluate(({ kind, id, ratios }) => window.sheetDrawingFixture.point(kind, id, ratios), { kind, id, ratios: ratios ?? { x: 0.03, y: 0.03 } })
  if (modifier) await page.keyboard.down(modifier)
  try { await page.mouse.click(point.x, point.y) } finally { if (modifier) await page.keyboard.up(modifier) }
}
async function menuAction(button, label) {
  await toolbar().getByRole('button', { name: button, exact: true }).click()
  await page.getByRole('menuitem', { name: label, exact: true }).last().click()
}
async function waitSelection(expected) {
  await page.waitForFunction(expected => JSON.stringify(window.sheetDrawingFixture.state().selection.map(ref => `${ref.kind}:${ref.id}`).sort()) === JSON.stringify(expected), keys(expected))
}
async function waitHistory(expected) { await page.waitForFunction(expected => window.sheetDrawingFixture.state().history.undo === expected, expected) }
async function drag(start, dx, dy, cancel = false) {
  await page.mouse.move(start.x, start.y); await page.mouse.down(); await page.mouse.move(start.x + dx, start.y + dy, { steps: 8 })
  if (cancel) await page.keyboard.press('Escape')
  await page.mouse.up()
}

test('Ctrl-click shares native image/chart selection and Ant alignment commits one history entry with undo/redo/reopen', async () => {
  await clickRef('image', 'a'); await clickRef('chart', 'c', 'Control'); await waitSelection(mixed)
  await expect(toolbar()).toBeVisible(); const before = await state()
  await menuAction('对齐绘图对象', '左对齐'); await waitHistory(before.history.undo + 1)
  const aligned = await state()
  near(drawingAffineFrame(object(aligned, 'image', 'a').geometry).left, drawingAffineFrame(object(aligned, 'chart', 'c').geometry).left)
  assert.equal(aligned.commands.filter(command => command.id === 'sheet.command.align-sheet-drawings').length, 1)
  await page.evaluate(() => window.sheetDrawingFixture.undo()); await waitSelection(mixed); assert.deepEqual((await state()).model, before.model)
  await page.evaluate(() => window.sheetDrawingFixture.redo()); assert.deepEqual((await state()).model, aligned.model)
  await page.evaluate(() => window.sheetDrawingFixture.reopen()); assert.deepEqual((await state()).model, aligned.model)
})

test('a grouped chart input bridge moves every native owner once and Escape cancels the next preview', async () => {
  await clickRef('image', 'a'); await clickRef('chart', 'c', 'Control'); await menuAction('组合或取消组合', '组合')
  await page.waitForFunction(() => window.sheetDrawingFixture.state().selection[0]?.kind === 'group')
  const grouped = await state(), start = await page.evaluate(() => window.sheetDrawingFixture.point('chart', 'c', { x: 0.03, y: 0.03 }))
  await drag(start, 70, 35); await waitHistory(grouped.history.undo + 1)
  const moved = await state()
  for (const ref of mixed) {
    const old = object(grouped, ref.kind, ref.id).geometry.matrix, next = object(moved, ref.kind, ref.id).geometry.matrix
    near(next.e - old.e, 70); near(next.f - old.f, 35)
    for (const component of ['a', 'b', 'c', 'd']) near(next[component], old[component])
  }
  const nextStart = await page.evaluate(() => window.sheetDrawingFixture.point('chart', 'c', { x: 0.03, y: 0.03 }))
  await drag(nextStart, 50, 20, true)
  assert.deepEqual((await state()).model, moved.model); assert.equal((await state()).history.undo, moved.history.undo)
})

test('shared Canvas resize and rotation handles retain native author size/source and ungroup without applying another transform', async () => {
  await page.evaluate(selection => window.sheetDrawingFixture.select(selection), mixed); await menuAction('组合或取消组合', '组合')
  await page.waitForFunction(() => window.sheetDrawingFixture.state().selection[0]?.kind === 'group')
  const before = await state(), resize = await page.evaluate(() => window.sheetDrawingFixture.control('se'))
  await drag(resize, 50, 30); await waitHistory(before.history.undo + 1)
  const resized = await state(), rotate = await page.evaluate(() => window.sheetDrawingFixture.control('rotate'))
  await drag(rotate, 35, 10); await waitHistory(resized.history.undo + 1)
  const transformed = await state()
  assert.deepEqual(transformed.model.charts[0].size, before.model.charts[0].size); assert.deepEqual(transformed.model.charts[0].source, before.model.charts[0].source)
  assert.ok(Math.abs(object(transformed, 'chart', 'c').geometry.matrix.b) > 0.001)
  await menuAction('组合或取消组合', '取消组合'); await waitHistory(transformed.history.undo + 1)
  const expanded = await state(); assert.equal(expanded.scene.state.groups.length, 0)
  assert.deepEqual(expanded.scene.objects.map(object => object.geometry), transformed.scene.objects.map(object => object.geometry)); await waitSelection(mixed)
})

test('Canvas mixed paint order controls input even when an overlapping image sits behind the native chart DOM bridge', async () => {
  await page.evaluate(() => window.sheetDrawingFixture.overlapImage())
  const before = await state(), imagePoint = await page.evaluate(() => window.sheetDrawingFixture.point('image', 'a', { x: 0.5, y: 0.5 }))
  await page.mouse.click(imagePoint.x, imagePoint.y); await waitSelection([key('image', 'a')])
  assert.equal((await state()).history.undo, before.history.undo)
  await page.evaluate(() => window.sheetDrawingFixture.arrange([{ kind: 'image', id: 'a' }], 'back'))
  await page.mouse.click(imagePoint.x, imagePoint.y); await waitSelection([key('chart', 'c')])
})

test('a single affine chart focused through Canvas nudges in worksheet pixels while IME and text inputs retain their keys', async () => {
  await page.evaluate(() => window.sheetDrawingFixture.transform([{ kind: 'chart', id: 'c' }], { a: 1, b: 0.15, c: 0.2, d: 1, e: 40, f: 20 }))
  await grid().focus(); const before = await state(); await grid().press('ArrowRight'); await waitHistory(before.history.undo + 1)
  const after = await state(), old = object(before, 'chart', 'c').geometry, next = object(after, 'chart', 'c').geometry
  near(next.matrix.e, old.matrix.e + 1); near(next.matrix.f, old.matrix.f)
  for (const component of ['a', 'b', 'c', 'd']) near(next.matrix[component], old.matrix[component])
  assert.deepEqual(after.model.charts[0].size, before.model.charts[0].size)
  await grid().evaluate(node => node.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', keyCode: 229, isComposing: true, bubbles: true, cancelable: true })))
  assert.deepEqual((await state()).model, after.model)
  const name = page.getByRole('textbox', { name: '名称框', exact: true }); await name.focus(); await name.press('ArrowRight')
  assert.deepEqual((await state()).model, after.model)
  const resize = page.getByRole('button', { name: '调整图表大小', exact: true }); await resize.focus()
  await resize.evaluate(node => node.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', keyCode: 229, isComposing: true, bubbles: true, cancelable: true })))
  await resize.press('Control+ArrowRight'); assert.deepEqual((await state()).model, after.model)
  await resize.press('ArrowRight'); await waitHistory(after.history.undo + 1)
  const resized = await state(); near(resized.model.charts[0].size.width, after.model.charts[0].size.width + 1)
  assert.deepEqual(resized.model.charts[0].source, after.model.charts[0].source)
})

test('frozen chart title and plot resize input use the same author controller at zoom and scroll', async () => {
  await page.evaluate(() => window.sheetDrawingFixture.freeze(12, 4)); await page.evaluate(() => window.sheetDrawingFixture.zoom(125))
  await page.evaluate(() => window.sheetDrawingFixture.scroll(120, 80))
  const titlePoint = await page.evaluate(() => window.sheetDrawingFixture.authorPoint('title', undefined, { frozenX: true, frozenY: true })), before = await state()
  await drag(titlePoint, 12.5, 6.25); await waitHistory(before.history.undo + 1)
  const titled = await state(); assert.deepEqual(titled.model.charts[0].position, before.model.charts[0].position)
  assert.notDeepEqual(titled.model.charts[0].options.titleStyle.manualLayout, before.model.charts[0].options.titleStyle.manualLayout)
  await grid().press('Escape'); await page.evaluate(() => window.sheetDrawingFixture.select([{ kind: 'chart', id: 'c' }]))
  await page.getByRole('button', { name: '编辑图表绘图区', exact: true }).click()
  const nativeHandle = page.getByRole('button', { name: '调整绘图区左上角', exact: true })
  await nativeHandle.evaluate(node => node.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', keyCode: 229, isComposing: true, bubbles: true, cancelable: true })))
  await nativeHandle.press('Control+ArrowDown'); assert.deepEqual((await state()).model, titled.model)
  const handle = await page.evaluate(() => window.sheetDrawingFixture.authorPoint('plot', 'nw', { frozenX: true, frozenY: true }))
  await drag(handle, 12.5, 6.25); await waitHistory(titled.history.undo + 1)
  const plotted = await state(); assert.notDeepEqual(plotted.model.charts[0].options.plotArea, titled.model.charts[0].options.plotArea)
  assert.deepEqual(plotted.model.charts[0].position, titled.model.charts[0].position); assert.deepEqual(plotted.model.charts[0].size, titled.model.charts[0].size)
})

test('worksheet switching restores shared selection and an empty inactive publication stays empty', async () => {
  await page.evaluate(selection => window.sheetDrawingFixture.select(selection), mixed); await waitSelection(mixed)
  const before = await state(); await page.evaluate(() => window.sheetDrawingFixture.switchSheet('other')); await waitSelection([])
  await page.evaluate(() => window.sheetDrawingFixture.switchSheet('data')); await waitSelection(mixed)
  assert.deepEqual((await state()).scene.objects, before.scene.objects)
  await grid().press('Escape'); await waitSelection([])
  await page.evaluate(() => window.sheetDrawingFixture.switchSheet('other')); await page.evaluate(() => window.sheetDrawingFixture.switchSheet('data')); await waitSelection([])
})

test('Ant popup theme and keyboard context focus do not create content history; read-only mixed controls cannot commit', async () => {
  await page.evaluate(selection => window.sheetDrawingFixture.select(selection), mixed); await grid().focus(); const before = await state()
  await grid().press('Shift+F10'); const menu = page.getByRole('menu').last(); await expect(menu).toBeVisible()
  const light = await menu.evaluate(node => getComputedStyle(node).backgroundColor); await page.keyboard.press('Escape')
  await page.evaluate(() => window.sheetDrawingFixture.appearance(true, 'grid')); await grid().focus(); await grid().press('Shift+F10')
  assert.notEqual(await page.getByRole('menu').last().evaluate(node => getComputedStyle(node).backgroundColor), light)
  await page.keyboard.press('Escape'); await expect(grid()).toBeFocused(); assert.deepEqual((await state()).model, before.model)
  await page.evaluate(() => window.sheetDrawingFixture.mount({ readonly: true })); await page.evaluate(selection => window.sheetDrawingFixture.select(selection), mixed)
  const locked = await state(); await expect(toolbar().getByRole('button', { name: '对齐绘图对象', exact: true })).toBeDisabled()
  await grid().focus(); await grid().press('Delete'); await grid().press('ArrowRight'); assert.deepEqual((await state()).model, locked.model)
  await waitSelection(mixed)
})

test('mixed Delete removes only selected owners in one native history and restores the shared selection on undo', async () => {
  await page.evaluate(selection => window.sheetDrawingFixture.select(selection), mixed); await grid().focus(); const before = await state()
  await grid().press('Delete'); await waitHistory(before.history.undo + 1)
  const removed = await state(); assert.deepEqual(removed.model.sheets[0].images.map(image => image.id), ['d']); assert.equal(removed.model.charts.length, 0)
  await waitSelection([]); await page.evaluate(() => window.sheetDrawingFixture.undo()); await waitSelection(mixed)
  assert.deepEqual((await state()).model, before.model)
})

test('mixed Canvas arrow keys retain all movements in one timed undo and redo the final native-owner state', async () => {
  await page.evaluate(() => { Date.now = () => 1000 })
  await page.evaluate(selection => window.sheetDrawingFixture.select(selection), mixed); await grid().focus()
  const before = await state()
  await grid().press('ArrowRight'); await waitHistory(before.history.undo + 1)
  await grid().press('ArrowDown'); await page.waitForFunction(revision => window.sheetDrawingFixture.state().revision === revision, before.revision + 2)
  await grid().press('Shift+ArrowLeft'); await page.waitForFunction(revision => window.sheetDrawingFixture.state().revision === revision, before.revision + 3)
  const after = await state(); assert.equal(after.history.undo, before.history.undo + 1)
  assert.deepEqual(after.commands.slice(before.commands.length).map(command => command.id), Array(3).fill('sheet.command.move-drawing'))
  for (const ref of mixed) {
    const old = object(before, ref.kind, ref.id).geometry, next = object(after, ref.kind, ref.id).geometry
    near(next.matrix.e, old.matrix.e - 9); near(next.matrix.f, old.matrix.f + 1)
  }
  assert.deepEqual(after.model.charts[0].size, before.model.charts[0].size)
  await page.evaluate(() => window.sheetDrawingFixture.undo()); await waitSelection(mixed); assert.deepEqual((await state()).model, before.model)
  await page.evaluate(() => window.sheetDrawingFixture.redo()); assert.deepEqual((await state()).model, after.model)
})

test('a transformed group moves all members once through the same keyboard command at zoom', async () => {
  await page.evaluate(() => { Date.now = () => 1000 })
  await page.evaluate(selection => window.sheetDrawingFixture.select(selection), mixed); await menuAction('组合或取消组合', '组合')
  const grouped = await state(), group = grouped.selection[0]
  await page.evaluate(({ group }) => window.sheetDrawingFixture.transform([group], { a: 1, b: 0.12, c: 0.18, d: 1, e: 40, f: 20 }), { group })
  await page.evaluate(() => window.sheetDrawingFixture.zoom(125)); await grid().focus(); const before = await state()
  await grid().press('Shift+ArrowRight'); await waitHistory(before.history.undo + 1)
  await grid().press('ArrowUp'); await page.waitForFunction(revision => window.sheetDrawingFixture.state().revision === revision, before.revision + 2)
  const after = await state(); assert.equal(after.history.undo, before.history.undo + 1)
  for (const ref of mixed) {
    const old = object(before, ref.kind, ref.id).geometry, next = object(after, ref.kind, ref.id).geometry
    near(next.matrix.e, old.matrix.e + 10); near(next.matrix.f, old.matrix.f - 1)
    for (const name of ['a', 'b', 'c', 'd']) near(next.matrix[name], old.matrix[name])
  }
  assert.deepEqual(after.scene.state.groups.map(group => group.members), before.scene.state.groups.map(group => group.members))
  await page.evaluate(() => window.sheetDrawingFixture.undo()); await waitSelection([group]); assert.deepEqual((await state()).model, before.model)
})

test('a single image uses the shared move command, ignores consumed/composition keys and preserves a saved undo boundary', async () => {
  await page.evaluate(() => { Date.now = () => 1000 })
  await clickRef('image', 'a'); await grid().focus(); const before = await state()
  await grid().press('ArrowRight'); await waitHistory(before.history.undo + 1); const first = await state()
  assert.equal(first.commands.at(-1).id, 'sheet.command.move-drawing')
  await grid().evaluate(node => {
    const consumed = new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true, cancelable: true }); consumed.preventDefault(); node.dispatchEvent(consumed)
    node.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', keyCode: 229, isComposing: true, bubbles: true, cancelable: true }))
  })
  assert.deepEqual((await state()).model, first.model)
  await grid().press('ArrowDown'); await page.waitForFunction(revision => window.sheetDrawingFixture.state().revision === revision, before.revision + 2)
  assert.equal((await state()).history.undo, before.history.undo + 1)
  const saved = await page.evaluate(() => window.sheetDrawingFixture.markSaved()); assert.equal(saved.history.dirty, false)
  await grid().press('ArrowRight'); await waitHistory(before.history.undo + 2)
  assert.equal((await state()).history.dirty, true)
  await page.evaluate(() => window.sheetDrawingFixture.undo())
  assert.deepEqual((await state()).scene.objects, saved.scene.objects); assert.equal((await state()).history.dirty, false)
})
