import assert from 'node:assert/strict'
import { test } from 'vitest'
import { drawingAffinePoint } from '../src/domain/drawing-affine.ts'
import { normalizeDrawingState, planDrawingGroup, applyDrawingTransaction } from '../src/features/drawing-groups/index.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { slicerObjectLayouts } from '../src/features/slicer/slicer-renderer.ts'
import { updateSlicerPresentation } from '../src/features/slicer/slicer-engine.ts'
import { slicerControlPaneLayouts, slicerControlScreenPolygon, slicerControlTabIndex, slicerControlVisible,
  subtractSlicerControlPolygon } from '../src/features/slicer/slicer-control-overlay.ts'
import { slicerDrawingControls, slicerDrawingWorkbook } from './fixtures/slicer-drawing-model-fixture.mjs'

// Deferred geometry/paint-owner contracts. CSS, Ant, browser focus and either backend are still unverified.
const viewport = { width: 1100, height: 750, zoom: 1.25, scrollLeft: 30, scrollTop: 20 }
const ref = (kind, id) => ({ kind, id })
const rect = (x, y, width, height) => [{ x, y }, { x: x + width, y }, { x: x + width, y: y + height }, { x, y: y + height }]
const area = polygon => Math.abs(polygon.reduce((sum, point, index) => {
  const next = polygon[(index + 1) % polygon.length]; return sum + point.x * next.y - next.x * point.y
}, 0)) / 2
const inside = (polygon, point) => {
  const crosses = polygon.map((first, index) => { const last = polygon[(index + 1) % polygon.length]
    return (last.x - first.x) * (point.y - first.y) - (last.y - first.y) * (point.x - first.x) })
  return crosses.every(value => value >= -1e-7) || crosses.every(value => value <= 1e-7)
}
const near = (a, b) => assert.ok(Math.abs(a - b) < 1e-7, `${a} !== ${b}`)
function fixture({ freeze, id = 's', matrix } = {}) {
  const workbook = slicerDrawingWorkbook(), model = workbook.slicers.find(control => control.id === id)
  if (freeze) workbook.sheets[0].freeze = freeze
  if (matrix) updateSlicerPresentation(model, { anchorType: 'none', drawingGeometry: { ...model.drawingGeometry, matrix } }, workbook)
  const grid = new SpreadsheetGeometry(toExcelWorkbookView(workbook).sheets[0]),
    layout = slicerObjectLayouts(grid, slicerDrawingControls(workbook), false).find(layout => layout.id === id),
    object = { ref: ref(model.type, id), sheetId: 'data', geometry: layout.geometry, available: true, groupable: true }
  return { workbook, grid, layout, object }
}
function scene(objects, order = objects.map(object => object.ref)) {
  const input = { sheetId: 'data', objects }
  return { ...input, state: { ...normalizeDrawingState(input), rootOrder: order } }
}
const matrices = [
  ['rotation', { a: 0, b: 1, c: -1, d: 0, e: 600, f: 80 }],
  ['reflection', { a: -1, b: .25, c: .2, d: 1, e: 700, f: 80 }],
  ['shear', { a: 1, b: .4, c: .3, d: 1, e: 150, f: 80 }],
]
for (const [name, matrix] of matrices) for (const id of ['s', 't']) test(`${name} ${id} Ant control polygons reuse the Canvas pane matrix`, () => {
  const f = fixture({ id, matrix }), panes = slicerControlPaneLayouts(f.grid, viewport, [f.layout], scene([f.object]))
  assert.ok(panes.length)
  for (const pane of panes) {
    assert.equal(slicerControlVisible(pane, f.layout.clear), true)
    const polygon = slicerControlScreenPolygon(pane, f.layout.clear), local = {
      x: f.layout.clear.x - f.layout.rect.x, y: f.layout.clear.y - f.layout.rect.y }
    const point = drawingAffinePoint(pane.pane.frame.matrix, local.x, local.y)
    near(polygon[0].x, point.x); near(polygon[0].y, point.y)
    assert.match(pane.clipPath, /^path\(nonzero, "/)
    assert.equal(slicerControlVisible(pane, { ...f.layout.clear, width: 0 }), false)
  }
})

test('frozen pane copies clip every polygon and expose each logical control to Tab only once', () => {
  const f = fixture({ freeze: { rows: 5, columns: 2 } }), panes = slicerControlPaneLayouts(f.grid, viewport, [f.layout], scene([f.object]))
  assert.ok(panes.length >= 2)
  for (const pane of panes) for (const polygon of pane.polygons) for (const point of polygon) {
    const clip = pane.pane.clip
    assert.ok(point.x >= clip.x - 1e-7 && point.x <= clip.x + clip.width + 1e-7)
    assert.ok(point.y >= clip.y - 1e-7 && point.y <= clip.y + clip.height + 1e-7)
  }
  for (const control of [f.layout.clear, f.layout.multiSelect, ...f.layout.items.map(item => item.rect)]) {
    const visible = panes.filter(pane => slicerControlVisible(pane, control))
    assert.ok(visible.length)
    assert.equal(visible.filter(pane => slicerControlTabIndex(panes, pane, control) === 0).length, 1)
  }
})

for (const reverse of [false, true]) test(`overlapping occluders retain their union even with ${reverse ? 'reversed' : 'normal'} winding`, () => {
  const subject = rect(0, 0, 100, 100), masks = [rect(20, 20, 60, 60), rect(50, 40, 40, 50)]
  if (reverse) masks[1].reverse()
  const before = structuredClone([subject, masks])
  const pieces = masks.reduce((pieces, mask) => pieces.flatMap(piece => subtractSlicerControlPolygon(piece, mask)), [subject])
  near(pieces.reduce((sum, piece) => sum + area(piece), 0), 5600)
  for (const point of [{ x: 35, y: 35 }, { x: 70, y: 60 }, { x: 85, y: 85 }]) assert.equal(pieces.some(piece => inside(piece, point)), false)
  assert.equal(pieces.some(piece => inside(piece, { x: 10, y: 10 })), true)
  assert.deepEqual([subject, masks], before)
})

test('complete leaf paint order decides which drawing owns overlapped controls', () => {
  const f = fixture(), mask = { ...f.object, ref: ref('shape', 'mask'), geometry: structuredClone(f.object.geometry) }
  assert.deepEqual(slicerControlPaneLayouts(f.grid, viewport, [f.layout], scene([f.object, mask])), [])
  const panes = slicerControlPaneLayouts(f.grid, viewport, [f.layout], scene([f.object, mask], [mask.ref, f.object.ref]))
  assert.ok(panes.some(pane => slicerControlVisible(pane, f.layout.items[0].rect)))
})

test('unavailable or non-viewable higher drawings do not hide controls; a non-viewable slicer has no overlay', () => {
  const f = fixture(), mask = { ...f.object, ref: ref('chart', 'mask'), geometry: structuredClone(f.object.geometry) }
  const unavailable = scene([f.object, { ...mask, available: false }])
  assert.ok(slicerControlPaneLayouts(f.grid, viewport, [f.layout], unavailable).length)
  const input = scene([f.object, mask])
  assert.ok(slicerControlPaneLayouts(f.grid, viewport, [f.layout], input, undefined, leaf => leaf.id !== 'mask').length)
  assert.deepEqual(slicerControlPaneLayouts(f.grid, viewport, [f.layout], input, undefined, leaf => leaf.id !== 's'), [])
})

test('higher preview geometry masks controls without changing authored geometry or state', () => {
  const f = fixture(), mask = { ...f.object, ref: ref('image', 'mask'), geometry: { ...f.object.geometry,
    matrix: { ...f.object.geometry.matrix, e: 850, f: 500 } } }, input = scene([f.object, mask]), before = structuredClone(input)
  assert.ok(slicerControlPaneLayouts(f.grid, viewport, [f.layout], input).length)
  const previews = new Map([['image:mask', structuredClone(f.object.geometry)]])
  assert.deepEqual(slicerControlPaneLayouts(f.grid, viewport, [f.layout], input, previews), [])
  assert.deepEqual(input, before)
})

test('nested groups preserve descendant paint ownership rather than treating the group frame as an occluder', () => {
  const f = fixture(), mask = { ...f.object, ref: ref('shape', 'mask') }, other = { ...f.object, ref: ref('chart', 'other'),
    geometry: { ...f.object.geometry, matrix: { ...f.object.geometry.matrix, e: 850, f: 500 } } }
  let input = scene([f.object, mask, other])
  for (const [selection, groupId] of [[[mask.ref, other.ref], 'inner'], [[f.object.ref, ref('group', 'inner')], 'outer']]) {
    const plan = planDrawingGroup(input, { selection, groupId }); assert.equal(plan.status, 'ready')
    input = applyDrawingTransaction(input, plan.transaction).scene
  }
  assert.deepEqual(slicerControlPaneLayouts(f.grid, viewport, [f.layout], input), [])
})

test('an outside or degenerate mask preserves the full subject and disconnected pieces remain positive', () => {
  const subject = rect(0, 0, 100, 100)
  for (const mask of [rect(150, 150, 20, 20), [{ x: 20, y: 20 }, { x: 20, y: 20 }, { x: 20, y: 20 }]]) {
    near(subtractSlicerControlPolygon(subject, mask).reduce((sum, polygon) => sum + area(polygon), 0), 10000)
  }
  assert.equal(subtractSlicerControlPolygon(subject, rect(-10, -10, 120, 120)).length, 0)
})
