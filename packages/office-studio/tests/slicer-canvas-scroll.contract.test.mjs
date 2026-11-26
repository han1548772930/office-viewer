import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffinePoint } from '../src/domain/drawing-affine.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { slicerObjectLayouts, timelineSerialAtOffset } from '../src/features/slicer/slicer-renderer.ts'
import { captureSlicerCanvasGesture, canResumeSlicerCanvasGesture, slicerCanvasScrollOffset } from '../src/features/slicer/slicer-canvas-gesture.ts'
import { updateSlicerPresentation } from '../src/features/slicer/slicer-engine.ts'
import { slicerDrawingControls, slicerDrawingWorkbook } from './fixtures/slicer-drawing-model-fixture.mjs'

// Deferred scrollbar coordinate/preview contracts. Native Slider events, cancellation and focus require the browser draft.
const context = workbook => ({ unitId: workbook.id, sheetId: 'data', revision: 7, userId: 'local-user',
  viewport: { width: 1100, height: 750, zoom: 1.25, scrollLeft: 30, scrollTop: 20 } })
const matrices = [
  ['rotation', { a: 0, b: 1, c: -1, d: 0, e: 700, f: 80 }],
  ['reflection', { a: -1, b: .25, c: .2, d: 1, e: 700, f: 80 }],
  ['shear', { a: 1, b: .4, c: .3, d: 1, e: 150, f: 80 }],
]
function fixture(id, matrix) {
  const workbook = slicerDrawingWorkbook({ scrollable: true }), model = workbook.slicers.find(control => control.id === id)
  updateSlicerPresentation(model, { columnCount: id === 's' ? 2 : 1, anchorType: 'none',
    drawingGeometry: { ...model.drawingGeometry, matrix } }, workbook)
  const grid = new SpreadsheetGeometry(toExcelWorkbookView(workbook).sheets[0]), controls = slicerDrawingControls(workbook)
  return { workbook, grid, controls, layout: slicerObjectLayouts(grid, controls, false).find(layout => layout.id === id) }
}
const world = (layout, point) => drawingAffinePoint(layout.matrix, point.x, point.y)
const center = rect => ({ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 })
function pointOnTrack(layout, fraction) {
  const track = layout.scrollbar
  return world(layout, layout.model.type === 'timeline' ? { x: track.x + track.width * fraction, y: track.y + track.height / 2 }
    : { x: track.x + track.width / 2, y: track.y + track.height * fraction })
}
for (const [name, matrix] of matrices) for (const id of ['s', 't']) test(`${name} ${id} track and thumb scrolling use local axes and native step units`, () => {
  const f = fixture(id, matrix), layout = f.layout, before = structuredClone(f.workbook), step = id === 's' ? 2 : 1
  assert.ok(layout.maxScrollOffset > 0)
  const start = world(layout, center(layout.scrollThumb)), gesture = captureSlicerCanvasGesture(layout, start, context(f.workbook))
  const expected = Math.min(layout.maxScrollOffset, Math.round(layout.maxScrollOffset * .75 / step) * step)
  assert.equal(slicerCanvasScrollOffset(gesture, pointOnTrack(layout, .75), 'track'), expected)
  assert.equal(slicerCanvasScrollOffset(gesture, pointOnTrack(layout, -1), 'track'), 0)
  assert.equal(slicerCanvasScrollOffset(gesture, pointOnTrack(layout, 2), 'track'), layout.maxScrollOffset)
  const local = center(layout.scrollThumb), travel = id === 't' ? layout.scrollbar.width - layout.scrollThumb.width : layout.scrollbar.height - layout.scrollThumb.height
  const end = world(layout, { x: local.x + (id === 't' ? travel * .4 : 0), y: local.y + (id === 's' ? travel * .4 : 0) })
  assert.equal(slicerCanvasScrollOffset(gesture, end, 'thumb'), Math.min(layout.maxScrollOffset, Math.round(layout.maxScrollOffset * .4 / step) * step))
  assert.deepEqual(f.workbook, before)
})

for (const id of ['s', 't']) test(`${id} preview changes its visible window while keeping the captured author and cancel restoration`, () => {
  const f = fixture(id, matrices[2][1]), before = structuredClone(f.workbook), layout = f.layout,
    gesture = captureSlicerCanvasGesture(layout, world(layout, center(layout.scrollThumb)), context(f.workbook)),
    offset = Math.min(8, layout.maxScrollOffset), previews = new Map([[`${layout.model.type}:${id}`, offset]])
  const preview = slicerObjectLayouts(f.grid, f.controls, false, 'zh-CN', undefined, previews).find(layout => layout.id === id)
  assert.equal(preview.scrollOffset, offset); assert.equal(canResumeSlicerCanvasGesture(gesture, preview, context(f.workbook)), true)
  assert.notDeepEqual(id === 's' ? preview.items : preview.periods, id === 's' ? layout.items : layout.periods)
  const restored = slicerObjectLayouts(f.grid, f.controls, false, 'zh-CN', undefined, new Map()).find(layout => layout.id === id)
  assert.equal(restored.scrollOffset, layout.scrollOffset); assert.deepEqual(f.workbook, before)
  assert.equal(canResumeSlicerCanvasGesture(gesture, preview, { ...context(f.workbook), revision: 8 }), false)
})

test('timeline scrollbar units convert back to dates from the real multi-month Pivot source', () => {
  const f = fixture('t', matrices[2][1]), items = f.controls.find(control => control.model.id === 't').items
  const first = timelineSerialAtOffset(items, 'month', false, 0), second = timelineSerialAtOffset(items, 'month', false, 1)
  assert.equal(second - first, 31)
  assert.equal(timelineSerialAtOffset([], 'month', false, 1), undefined)
})

test('a source that fits in one window has no effective scroll movement', () => {
  const workbook = slicerDrawingWorkbook(), grid = new SpreadsheetGeometry(toExcelWorkbookView(workbook).sheets[0])
  const layout = slicerObjectLayouts(grid, slicerDrawingControls(workbook), false).find(layout => layout.id === 's'), start = world(layout, center(layout.rect))
  assert.equal(layout.maxScrollOffset, 0)
  assert.equal(slicerCanvasScrollOffset(captureSlicerCanvasGesture(layout, start, context(workbook)), start, 'thumb'), 0)
})
