import assert from 'node:assert/strict'
import { test } from 'vitest'
import { SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../../src/rendering/spreadsheet/scene.ts'

// Scene contracts are implementation-phase drafts; no renderer or benchmark has been run.
const viewport = { width: 640, height: 280, scrollLeft: 0, scrollTop: 0, zoom: 1 }
const scope = (startRow, startColumn, endRow = startRow, endColumn = startColumn) => ({ startRow, startColumn, endRow, endColumn })
const marker = '#bd29e3'
function sheet(extra = {}) {
  return { id: 'conditional-preview', name: 'Preview', cells: {}, rowCount: 100, columnCount: 100,
    defaultColumnWidth: 80, defaultRowHeight: 24, ...extra }
}
const marks = scene => scene.commands.filter(command => command.kind === 'line' && command.color === marker)
const withoutMarks = scene => scene.commands.filter(command => command.kind !== 'line' || command.color !== marker)
function scene(model, view, ranges, options = {}) {
  return createSpreadsheetScene(new SpreadsheetGeometry(model), view,
    { activeCell: 'A1', selectedCells: [], selectedRanges: [scope(0, 0)], ...options,
      conditionalFormatPreviewRanges: ranges, conditionalFormatPreviewColor: marker })
}

test('multiple rule scopes add transient perimeters and preserve selection, fill handle and source model', () => {
  const model = sheet(), selectedRanges = [scope(0, 0)], before = JSON.stringify(model)
  const ranges = [scope(0, 0, 3, 0), scope(5, 3, 6, 4)]
  const plain = scene(model, viewport, [], { selectedRanges }), preview = scene(model, viewport, ranges, { selectedRanges })
  assert.ok(marks(preview).length > 8); assert.deepEqual(withoutMarks(preview), plain.commands)
  const layout = new SpreadsheetGeometry(model)
  for (const range of ranges) {
    const first = layout.cellRect(range.startRow, range.startColumn, viewport)
    const last = layout.cellRect(range.endRow, range.endColumn, viewport)
    assert.ok(marks(preview).some(line => line.y1 === first.y + .5 && line.y2 === first.y + .5))
    assert.ok(marks(preview).some(line => line.y1 === last.y + last.height + .5 && line.y2 === last.y + last.height + .5))
  }
  assert.deepEqual(selectedRanges, [scope(0, 0)]); assert.equal(JSON.stringify(model), before)
  assert.ok(preview.commands.some(command => command.kind === 'rect' && command.width === 5 && command.height === 5 && command.fill === '#246bfd'))
})

test('scrolling a multi-pane rule clips its perimeter and never introduces frozen-pane seams', () => {
  const model = sheet({ freeze: { rows: 1, columns: 1 } }), view = { ...viewport, scrollLeft: 100, scrollTop: 48 }
  const lines = marks(scene(model, view, [scope(0, 0, 3, 3)])), layout = new SpreadsheetGeometry(model)
  assert.ok(lines.length > 0)
  const seamX = layout.headerWidth + layout.columns.offset(layout.frozenColumns) * view.zoom + .5
  const seamY = layout.headerHeight + layout.rows.offset(layout.frozenRows) * view.zoom + .5
  assert.ok(lines.every(line => !(line.x1 === seamX && line.x2 === seamX) && !(line.y1 === seamY && line.y2 === seamY)))
  for (const line of lines) {
    assert.ok(line.clip && line.x1 >= line.clip.x && line.x2 <= line.clip.x + line.clip.width)
    assert.ok(line.y1 >= line.clip.y && line.y2 <= line.clip.y + line.clip.height)
  }
  const plain = scene(model, view, []), preview = scene(model, view, [scope(0, 0, 3, 3)])
  assert.deepEqual(withoutMarks(preview), plain.commands)
})

test('preview geometry follows editor zoom while theme color remains a transient scene input', () => {
  const model = sheet(), range = scope(1, 1, 3, 2), layout = new SpreadsheetGeometry(model)
  for (const zoom of [.5, 1, 2]) {
    const view = { ...viewport, zoom }, preview = scene(model, view, [range]), lines = marks(preview)
    const first = layout.cellRect(1, 1, view), last = layout.cellRect(3, 2, view)
    assert.ok(lines.length > 0)
    assert.ok(lines.some(line => line.y1 === first.y + .5 && line.y2 === first.y + .5))
    assert.ok(lines.some(line => line.x1 === last.x + last.width + .5 && line.x2 === last.x + last.width + .5))
    const themed = createSpreadsheetScene(layout, view, { activeCell: 'A1', selectedCells: [], selectedRanges: [scope(0, 0)],
      conditionalFormatPreviewRanges: [range], conditionalFormatPreviewColor: '#18c799' })
    assert.equal(themed.commands.filter(command => command.kind === 'line' && command.color === '#18c799').length, lines.length)
    assert.deepEqual(themed.commands.filter(command => command.kind !== 'line' || command.color !== '#18c799'), withoutMarks(preview))
  }
})

test('offscreen, hidden and canceled preview scopes produce no marks and leave content commands intact', () => {
  for (const [model, view, ranges] of [
    [sheet(), viewport, [scope(50, 50)]],
    [sheet({ hiddenRows: [0, 1, 2, 3] }), viewport, [scope(0, 0, 3, 3)]],
    [sheet(), viewport, []],
  ]) {
    const plain = scene(model, view, []), preview = scene(model, view, ranges)
    assert.deepEqual(marks(preview), []); assert.deepEqual(preview.commands, plain.commands)
  }
})

test('a whole-sheet preview generates bounded visible perimeter commands at half, normal and double zoom', () => {
  const model = sheet({ rowCount: 1_048_576, columnCount: 16_384 }), range = scope(0, 0, 1_048_575, 16_383)
  for (const zoom of [.5, 1, 2]) {
    const view = { ...viewport, zoom }, layout = new SpreadsheetGeometry(model), preview = scene(model, view, [range])
    assert.ok(marks(preview).length > 0 && marks(preview).length < (viewport.width + viewport.height) / 7 + 8)
    assert.ok(marks(preview).every(line => line.x1 >= layout.headerWidth && line.y1 >= layout.headerHeight))
    assert.deepEqual(withoutMarks(preview), scene(model, view, []).commands)
  }
})
