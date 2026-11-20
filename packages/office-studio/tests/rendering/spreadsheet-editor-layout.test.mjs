import assert from 'node:assert/strict'
import { test } from 'node:test'
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
  assert.deepEqual(layout.rect, { x: 104, y: 44, width: 120, height: 36 })
  assert.deepEqual(layout.clip, { x: 164, y: 62, width: 60, height: 18 })
  assert.equal(layout.clipPath, 'inset(18px 0px 0px 60px)')
})

test('frozen-cell input stays fixed when the unfrozen panes scroll', () => {
  const layout = createCellEditorLayout(geometry(), 'A1', viewport)
  assert.deepEqual(layout.rect, { x: 44, y: 26, width: 120, height: 36 })
  assert.deepEqual(layout.clip, layout.rect)
  assert.equal(layout.clipPath, 'inset(0px 0px 0px 0px)')
})

test('offscreen inputs are clipped away instead of being moved to a different visible cell', () => {
  const layout = createCellEditorLayout(geometry(), 'B2', { ...viewport, scrollLeft: 300, scrollTop: 200 })
  assert.deepEqual(layout.rect, { x: -136, y: -138, width: 120, height: 36 })
  assert.equal(layout.clipPath, 'inset(50%)')
})

test('merged input uses the merged anchor and is clipped to the actual viewport', () => {
  const layout = createCellEditorLayout(geometry({ merges: ['C3:L20'] }), 'F8', { ...viewport, scrollLeft: 0, scrollTop: 0, zoom: 1 })
  assert.deepEqual(layout.rect, { x: 204, y: 74, width: 800, height: 432 })
  assert.deepEqual(layout.clip, { x: 204, y: 74, width: 296, height: 226 })
  assert.equal(layout.clipPath, 'inset(0px 504px 206px 0px)')
})

test('input geometry retains small zoomed cell sizes and rejects addresses outside the model', () => {
  const grid = geometry({ columnWidths: { 1: 24 }, rowHeights: { 1: 16 }, freeze: { rows: 0, columns: 0 } })
  const layout = createCellEditorLayout(grid, 'B2', { ...viewport, scrollLeft: 0, scrollTop: 0, zoom: 0.5 })
  assert.deepEqual(layout.rect, { x: 84, y: 38, width: 12, height: 8 })
  assert.equal(createCellEditorLayout(grid, 'U101', viewport), undefined)
  assert.equal(createCellEditorLayout(grid, 'not-a-cell', viewport), undefined)
})
