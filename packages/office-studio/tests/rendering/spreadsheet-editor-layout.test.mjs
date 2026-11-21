import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createCellEditorLayout } from '../../src/rendering/spreadsheet/editor-layout.ts'
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
