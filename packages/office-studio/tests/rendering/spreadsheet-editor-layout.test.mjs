import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createCellEditorLayout, hitCellEditorFillHandle } from '../../src/rendering/spreadsheet/editor-layout.ts'
import { SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'

function geometry(extra = {}) {
  return new SpreadsheetGeometry({
    id: 'sheet', name: 'Sheet', cells: {}, rowCount: 100, columnCount: 20,
    defaultColumnWidth: 80, defaultRowHeight: 24,
    freeze: { rows: 1, columns: 1 }, ...extra,
  })
}
const viewport = { width: 500, height: 300, scrollLeft: 60, scrollTop: 18, zoom: 1.5 }

test('input stays at its scrolled cell coordinates and clips underneath frozen rows/columns', () => {
  const layout = createCellEditorLayout(geometry(), 'B2', viewport)
  assert.deepEqual(layout.rect, { x: 103.5, y: 43.5, width: 121, height: 37 })
  assert.deepEqual(layout.fillHandle, { x: 221, y: 77, width: 6, height: 6 })
  assert.deepEqual(layout.fillHandle, { x: 221, y: 77, width: 6, height: 6 })
  assert.deepEqual(layout.clip, { x: 164, y: 62, width: 60.5, height: 18.5 })
  assert.equal(layout.clipPath, 'inset(18.5px 0px 0px 60.5px)')
})

test('frozen-cell input stays fixed when the unfrozen panes scroll', () => {
  const layout = createCellEditorLayout(geometry(), 'A1', viewport)
  assert.deepEqual(layout.rect, { x: 43.5, y: 25.5, width: 121, height: 37 })
  assert.deepEqual(layout.clip, { x: 44, y: 26, width: 120, height: 36 })
  assert.deepEqual(layout.fillHandle, { x: 161, y: 59, width: 6, height: 6 })
  assert.equal(layout.fillHandleClipPath, 'inset(0px 3px 3px 0px)')
  assert.deepEqual(layout.fillHandle, { x: 161, y: 59, width: 6, height: 6 })
  assert.equal(layout.fillHandleClipPath, 'inset(0px 3px 3px 0px)')
  assert.equal(layout.clipPath, 'inset(0.5px 0.5px 0.5px 0.5px)')
})

test('offscreen inputs are clipped away instead of being moved to a different visible cell', () => {
  const layout = createCellEditorLayout(geometry(), 'B2', { ...viewport, scrollLeft: 300, scrollTop: 200 })
  assert.deepEqual(layout.rect, { x: -136.5, y: -138.5, width: 121, height: 37 })
  assert.equal(layout.clipPath, 'inset(50%)')
})

test('merged input uses the merged anchor and is clipped to the actual viewport', () => {
  const layout = createCellEditorLayout(geometry({ merges: ['C3:L20'] }), 'F8', { ...viewport, scrollLeft: 0, scrollTop: 0, zoom: 1 })
  assert.deepEqual(layout.rect, { x: 203.5, y: 73.5, width: 801, height: 433 })
  assert.deepEqual(layout.clip, { x: 203.5, y: 73.5, width: 296.5, height: 226.5 })
  assert.equal(layout.clipPath, 'inset(0px 504.5px 206.5px 0px)')
})

test('input geometry retains small zoomed cell sizes and rejects addresses outside the model', () => {
  const grid = geometry({ columnWidths: { 1: 24 }, rowHeights: { 1: 16 }, freeze: { rows: 0, columns: 0 } })
  const layout = createCellEditorLayout(grid, 'B2', { ...viewport, scrollLeft: 0, scrollTop: 0, zoom: 0.5 })
  assert.deepEqual(layout.rect, { x: 83.5, y: 37.5, width: 13, height: 9 })
  assert.equal(createCellEditorLayout(grid, 'U101', viewport), undefined)
  assert.equal(createCellEditorLayout(grid, 'not-a-cell', viewport), undefined)
})

test('multiline input grows below its cell and stays inside its scrollable pane', () => {
  const grid = geometry({ freeze: { rows: 0, columns: 0 } })
  const view = { ...viewport, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  const original = createCellEditorLayout(grid, 'B2', view)
  const grown = createCellEditorLayout(grid, 'B2', view, { contentHeight: 85 })
  assert.equal(grown.rect.height, 86)
  assert.deepEqual(grown.fillHandle, original.fillHandle)
  assert.equal(grown.rect.y, original.rect.y)
  const limited = createCellEditorLayout(grid, 'B2', view, { contentHeight: 800 })
  assert.equal(limited.rect.y + limited.rect.height, view.height)
  assert.equal(createCellEditorLayout(grid, 'B2', view, { contentHeight: NaN }).rect.height, original.rect.height)
})

test('a frozen row multiline bridge remains clipped above the frozen boundary', () => {
  const grid = geometry()
  const grown = createCellEditorLayout(grid, 'A1', viewport, { contentHeight: 180 })
  assert.equal(grown.rect.height, 37)
  assert.equal(grown.clip.y + grown.clip.height, 62)
  assert.equal(grown.clipPath, 'inset(0.5px 0.5px 0.5px 0.5px)')
})

test('both halves of an editing fill handle hit the same source cell across zoom levels', () => {
  const grid = geometry({ freeze: { rows: 0, columns: 0 } })
  for (const zoom of [.5, 1, 2]) {
    const view = { ...viewport, scrollLeft: 0, scrollTop: 0, zoom }
    const handle = createCellEditorLayout(grid, 'B2', view).fillHandle
    const source = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }
    assert.deepEqual(hitCellEditorFillHandle(grid, 'B2', view, handle.x + 1, handle.y + 1), source)
    assert.deepEqual(hitCellEditorFillHandle(grid, 'B2', view, handle.x + 4, handle.y + 4), source)
    assert.equal(hitCellEditorFillHandle(grid, 'B2', view, handle.x - 1, handle.y + 1), undefined)
  }
})

test('editing fill hits preserve merged source bounds and reject the portion clipped by a frozen pane', () => {
  const merged = geometry({ merges: ['C3:D4'], freeze: { rows: 0, columns: 0 } })
  const view = { ...viewport, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  const handle = createCellEditorLayout(merged, 'D4', view).fillHandle
  assert.deepEqual(hitCellEditorFillHandle(merged, 'D4', view, handle.x + 2, handle.y + 2),
    { startRow: 2, endRow: 3, startColumn: 2, endColumn: 3 })
  const frozen = geometry(), clip = createCellEditorLayout(frozen, 'A1', viewport).fillHandleClip
  assert.deepEqual(hitCellEditorFillHandle(frozen, 'A1', viewport, clip.x + 1, clip.y + 1),
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })
  assert.equal(hitCellEditorFillHandle(frozen, 'A1', viewport, clip.x + clip.width, clip.y + 1), undefined)
  assert.equal(hitCellEditorFillHandle(frozen, 'A1', viewport, clip.x + 1, clip.y + clip.height), undefined)
})

test('invisible, invalid, and non-finite editing handle coordinates never hit another visible cell', () => {
  const grid = geometry(), offscreen = { ...viewport, scrollLeft: 300, scrollTop: 200 }
  const handle = createCellEditorLayout(grid, 'B2', offscreen).fillHandle
  assert.equal(hitCellEditorFillHandle(grid, 'B2', offscreen, handle.x + 1, handle.y + 1), undefined)
  assert.equal(hitCellEditorFillHandle(grid, 'invalid', viewport, 50, 50), undefined)
  assert.equal(hitCellEditorFillHandle(grid, 'A1', viewport, NaN, 60), undefined)
})
