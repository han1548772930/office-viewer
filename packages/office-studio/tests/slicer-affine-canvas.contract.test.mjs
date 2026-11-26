import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffinePoint, multiplyDrawingAffine } from '../src/domain/drawing-affine.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { SpreadsheetGeometry, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'
import { sheetDrawingPaneLayouts } from '../src/features/sheet-drawings/sheet-drawing-viewport.ts'
import { slicerObjectDrawCommands, slicerObjectLayouts, hitTestSlicerObjects, slicerObjectLocalPoint } from '../src/features/slicer/slicer-renderer.ts'
import { canResumeSlicerCanvasGesture, captureSlicerCanvasGesture, slicerCanvasGestureDelta,
  slicerCanvasGestureGeometry } from '../src/features/slicer/slicer-canvas-gesture.ts'
import { updateSlicerPresentation } from '../src/features/slicer/slicer-engine.ts'
import { slicerDrawingControls, slicerDrawingWorkbook } from './fixtures/slicer-drawing-model-fixture.mjs'

// Deferred geometry/input contracts. These do not verify native IME, Ant popups or either browser backend.
const close = (actual, expected) => { for (const key of Object.keys(expected)) assert.ok(Math.abs(actual[key] - expected[key]) < 1e-8, `${key}: ${actual[key]} / ${expected[key]}`) }
const center = rect => ({ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 })
const world = (layout, point) => drawingAffinePoint(layout.matrix, point.x, point.y)
const viewport = { width: 1000, height: 600, zoom: 1.25, scrollLeft: 30, scrollTop: 20 }
const context = workbook => ({ unitId: workbook.id, sheetId: 'data', revision: 12, userId: 'local-user', viewport })
function layouts(workbook, previews) {
  const geometry = new SpreadsheetGeometry(toExcelWorkbookView(workbook).sheets[0])
  return slicerObjectLayouts(geometry, slicerDrawingControls(workbook), false, 'zh-CN', previews)
}
const matrices = [
  ['quarter rotation', { a: 0, b: 1, c: -1, d: 0, e: 700, f: 80 }],
  ['reflection', { a: -1, b: 0.25, c: 0.2, d: 1, e: 700, f: 80 }],
  ['shear', { a: 1, b: 0.4, c: 0.3, d: 1, e: 150, f: 80 }],
]
for (const [name, matrix] of matrices) for (const id of ['s', 't']) test(`${name} ${id} routes transformed control centers through the inverse matrix`, () => {
  const workbook = slicerDrawingWorkbook(), model = workbook.slicers.find(control => control.id === id)
  updateSlicerPresentation(model, { drawingGeometry: { width: model.drawingGeometry.width, height: model.drawingGeometry.height, matrix } }, workbook)
  const layout = layouts(workbook).find(layout => layout.id === id)
  const target = id === 's' ? layout.items[0].rect : layout.periods[0].rect
  for (const [rect, type] of [[layout.clear, 'clear'], [target, id === 's' ? 'item' : 'period']]) {
    const local = center(rect), transformed = world(layout, local), hit = hitTestSlicerObjects([layout], transformed.x, transformed.y)
    assert.equal(hit.type, type); assert.equal(hit.id, id); close(slicerObjectLocalPoint(layout, transformed), local)
  }
  const outside = world(layout, { x: layout.rect.x - 1, y: layout.rect.y + layout.rect.height / 2 })
  assert.equal(hitTestSlicerObjects([layout], outside.x, outside.y), undefined)
})

test('legacy identity layout preserves its old rectangle coordinates while attaching one common matrix and local clip', () => {
  const workbook = slicerDrawingWorkbook({ affine: false }), layout = layouts(workbook).find(layout => layout.id === 's')
  close(layout.matrix, { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 })
  const local = center(layout.items[0].rect)
  assert.equal(hitTestSlicerObjects([layout], local.x, local.y).type, 'item')
  for (const command of slicerObjectDrawCommands([layout])) {
    assert.deepEqual(command.matrix, layout.matrix); assert.equal(command.clip, undefined)
    assert.deepEqual(command.drawingRef, { kind: 'slicer', id: 's' }); assert.ok(command.localClip.width >= 0 && command.localClip.height >= 0)
  }
})

test('preview geometry is shared by content, matrix, clips and inverse hit testing without modifying authored models', () => {
  const workbook = slicerDrawingWorkbook(), before = structuredClone(workbook), model = workbook.slicers[0]
  const geometry = { ...model.drawingGeometry, matrix: { ...model.drawingGeometry.matrix, e: 350, f: 180 } }
  const preview = layouts(workbook, new Map([['slicer:s', geometry]])).find(layout => layout.id === 's')
  assert.deepEqual(preview.geometry, geometry)
  const point = world(preview, center(preview.items[0].rect))
  assert.equal(hitTestSlicerObjects([preview], point.x, point.y).type, 'item'); assert.deepEqual(workbook, before)
})

test('frozen panes apply scroll and zoom once to every primitive and leave local content clips unchanged', () => {
  const workbook = slicerDrawingWorkbook(); workbook.sheets[0].freeze = { rows: 4, columns: 2 }
  const sheet = toExcelWorkbookView(workbook).sheets[0], grid = new SpreadsheetGeometry(sheet)
  const layout = slicerObjectLayouts(grid, slicerDrawingControls(workbook), false).find(layout => layout.id === 's')
  const commands = slicerObjectDrawCommands([layout]), input = createSpreadsheetScene(grid, viewport, { selectedKey: 'A1', overlayCommands: commands })
  const projected = input.commands.filter(command => command.drawingRef?.kind === 'slicer' && command.drawingRef.id === 's')
  assert.ok(projected.length > commands.length)
  for (const pane of sheetDrawingPaneLayouts(grid, viewport, layout.geometry)) {
    const command = projected.find(command => command.kind === 'rect' && command.x === layout.rect.x && command.y === layout.rect.y
      && command.clip.x === pane.clip.x && command.clip.y === pane.clip.y)
    assert.ok(command); assert.deepEqual(command.localClip, layout.rect)
    close(drawingAffinePoint(command.matrix, layout.rect.x, layout.rect.y), drawingAffinePoint(pane.frame.matrix, 0, 0))
  }
})

test('captured move uses world pixels and returns the same delta consumed by the mixed drawing transaction', () => {
  const workbook = slicerDrawingWorkbook(), layout = layouts(workbook)[0], start = world(layout, center(layout.header))
  const gesture = captureSlicerCanvasGesture(layout, start, context(workbook)), after = slicerCanvasGestureGeometry(gesture, { x: start.x + 17, y: start.y - 8 }, 'move')
  close(after.matrix, { ...layout.geometry.matrix, e: layout.geometry.matrix.e + 17, f: layout.geometry.matrix.f - 8 })
  close(slicerCanvasGestureDelta(layout.geometry, after), { a: 1, b: 0, c: 0, d: 1, e: 17, f: -8 })
})

test('captured resize follows local axes and retains the original extent and affine basis', () => {
  const workbook = slicerDrawingWorkbook(), layout = layouts(workbook).find(layout => layout.id === 's')
  const local = { x: layout.rect.x + layout.rect.width, y: layout.rect.y + layout.rect.height }, start = world(layout, local)
  const gesture = captureSlicerCanvasGesture(layout, start, context(workbook)), after = slicerCanvasGestureGeometry(gesture,
    world(layout, { x: local.x + 36, y: local.y + 16 }), 'resize'), m = layout.geometry.matrix
  close(after.matrix, { a: m.a * 1.2, b: m.b * 1.2, c: m.c * 1.1, d: m.d * 1.1, e: m.e, f: m.f })
  assert.equal(after.width, layout.geometry.width); assert.equal(after.height, layout.geometry.height)
  close(multiplyDrawingAffine(slicerCanvasGestureDelta(layout.geometry, after), layout.geometry.matrix), after.matrix)
})

test('aspect ratio lock constrains local resizing while retaining world translation semantics', () => {
  const workbook = slicerDrawingWorkbook(); workbook.slicers[0].drawing.keepRatio = true
  const layout = layouts(workbook).find(layout => layout.id === 's'), local = center(layout.rect), start = world(layout, local)
  const gesture = captureSlicerCanvasGesture(layout, start, context(workbook)), after = slicerCanvasGestureGeometry(gesture,
    world(layout, { x: local.x + 36, y: local.y + 8 }), 'resize'), m = layout.geometry.matrix
  close(after.matrix, { a: m.a * 1.2, b: m.b * 1.2, c: m.c * 1.2, d: m.d * 1.2, e: m.e, f: m.f })
  const moved = slicerCanvasGestureGeometry(gesture, { x: start.x + 10, y: start.y + 7 }, 'move')
  close(moved.matrix, { e: m.e + 10, f: m.f + 7 })
})

test('captured scrollbar and timeline menu positions preserve local deltas through reflection', () => {
  const workbook = slicerDrawingWorkbook(), model = workbook.slicers.find(control => control.id === 't')
  updateSlicerPresentation(model, { drawingGeometry: { ...model.drawingGeometry, matrix: matrices[1][1] } }, workbook)
  const layout = layouts(workbook).find(layout => layout.id === 't'), local = center(layout.scrollbar ?? layout.level), start = world(layout, local)
  const gesture = captureSlicerCanvasGesture(layout, start, context(workbook)), end = slicerObjectLocalPoint(gesture.layout, world(layout, { x: local.x + 20, y: local.y + 3 }))
  close(end, { x: gesture.localStartPoint.x + 20, y: gesture.localStartPoint.y + 3 })
  const menuPoint = world(layout, center(layout.menu[2]))
  assert.equal(hitTestSlicerObjects([layout], menuPoint.x, menuPoint.y, layout.id).type, 'timelineMenu')
  close(slicerObjectLocalPoint(layout, menuPoint), center(layout.menu[2]))
})

for (const [name, change] of [
  ['unit owner', { unitId: 'reopened-unit' }], ['worksheet', { sheetId: 'other' }], ['revision', { revision: 13 }], ['user', { userId: 'peer' }],
  ['zoom', { viewport: { ...viewport, zoom: 2 } }], ['scroll', { viewport: { ...viewport, scrollTop: 21 } }],
  ['size', { viewport: { ...viewport, width: 999 } }],
]) test(`a stale ${name} rejects a captured control gesture`, () => {
  const workbook = slicerDrawingWorkbook(), layout = layouts(workbook)[0], current = context(workbook)
  const gesture = captureSlicerCanvasGesture(layout, world(layout, center(layout.rect)), current)
  assert.equal(canResumeSlicerCanvasGesture(gesture, layout, current), true)
  assert.equal(canResumeSlicerCanvasGesture(gesture, layout, { ...current, ...change }), false)
})

test('changed model/preview geometry and removed controls reject captured gestures', () => {
  const workbook = slicerDrawingWorkbook(), layout = layouts(workbook)[0], current = context(workbook)
  const gesture = captureSlicerCanvasGesture(layout, world(layout, center(layout.rect)), current)
  layout.model.caption = 'Newer author'
  assert.equal(canResumeSlicerCanvasGesture(gesture, layout, current), false)
  assert.equal(canResumeSlicerCanvasGesture(gesture, undefined, current), false)
  const fresh = layouts(workbook)[0], next = captureSlicerCanvasGesture(fresh, world(fresh, center(fresh.rect)), current)
  const preview = { ...fresh, geometry: { ...fresh.geometry, matrix: { ...fresh.geometry.matrix, e: fresh.geometry.matrix.e + 1 } } }
  assert.equal(canResumeSlicerCanvasGesture(next, preview, current), false)
})

test('locked headers remain selectable body content and do not advertise a native move gesture', () => {
  const workbook = slicerDrawingWorkbook(); workbook.slicers[0].lockedPosition = true
  const layout = layouts(workbook).find(layout => layout.id === 's'), point = world(layout, { x: layout.header.x + 10, y: layout.header.y + 8 })
  assert.equal(hitTestSlicerObjects([layout], point.x, point.y).type, 'body')
  assert.equal(GRID_HEADER_WIDTH, 44); assert.equal(GRID_HEADER_HEIGHT, 26)
})
