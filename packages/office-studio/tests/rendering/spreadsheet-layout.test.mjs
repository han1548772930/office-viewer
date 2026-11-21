import assert from 'node:assert/strict'
import { test } from 'vitest'
import { AxisIndex, cellAddress, SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'

function createGeometry(merges, columnCount = 100) {
  return new SpreadsheetGeometry({
    id: 'sheet-1',
    name: 'Sheet 1',
    cells: {},
    rowCount: 100,
    columnCount,
    merges,
  })
}

test('merge lookup preserves first-match semantics for overlapping and normalized ranges', () => {
  const geometry = createGeometry([
    'A1:C3',
    'B2:D4',
    'F1:E1',
    'A1:C1',
    'B1:D1',
    'not-a-range',
  ])

  for (let row = 0; row < 7; row += 1) {
    for (let column = 0; column < 8; column += 1) {
      const expected = geometry.merges.find(range => row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn)
      assert.deepEqual(geometry.mergeAt(row, column), expected)
    }
  }
})

test('merge lookup finds the end of a large horizontal merge band and skips gaps', () => {
  const count = 2_048
  const merges = Array.from({ length: count }, (_, index) => {
    const start = index * 3
    return `${cellAddress(4, start)}:${cellAddress(4, start + 1)}`
  })
  const geometry = createGeometry(merges, count * 3)

  assert.deepEqual(geometry.mergeAt(4, (count - 1) * 3 + 1), {
    startRow: 4,
    startColumn: (count - 1) * 3,
    endRow: 4,
    endColumn: (count - 1) * 3 + 1,
  })
  assert.equal(geometry.mergeAt(4, 2), undefined)
  assert.equal(geometry.mergeAt(3, 0), undefined)
})

test('explicit default row and column sizes drive frozen pane layout and hit testing', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'sheet-1',
    name: 'Sheet 1',
    cells: {},
    rowCount: 20,
    columnCount: 20,
    freeze: { rows: 1, columns: 1 },
    defaultRowHeight: 32,
    defaultColumnWidth: 90,
    rowHeights: { 1: 48 },
    columnWidths: { 0: 120 },
  })
  const viewport = { width: 420, height: 220, scrollLeft: 0, scrollTop: 0, zoom: 1 }

  assert.equal(geometry.rows.offset(2), 80)
  assert.equal(geometry.columns.offset(2), 210)
  assert.deepEqual(geometry.cellRect(2, 2, viewport), { x: 254, y: 106, width: 90, height: 32 })
  assert.equal(geometry.hitTest(260, 110, viewport)?.key, 'C3')
  assert.deepEqual(geometry.panes(viewport)[3]?.clip, { x: 44, y: 26, width: 120, height: 32 })
})

test('default first column width is used only when the sheet has no explicit default', () => {
  const workbenchDefaults = createGeometry([], 10)
  const officeDefaults = new SpreadsheetGeometry({
    id: 'sheet-2', name: 'Office sheet', cells: {}, rowCount: 10, columnCount: 10,
    defaultColumnWidth: 80, defaultRowHeight: 24,
  })

  assert.equal(workbenchDefaults.columns.size(0), 152)
  assert.equal(officeDefaults.columns.size(0), 80)
  assert.equal(officeDefaults.rows.size(0), 24)
})

test('hidden final axes and blank space beyond the sheet cannot be selected', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'small-sheet', name: 'Small sheet', cells: {}, rowCount: 3, columnCount: 3,
    defaultColumnWidth: 80, defaultRowHeight: 24,
    hiddenRows: [2], hiddenColumns: [2],
  })
  const viewport = { width: 500, height: 300, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  assert.equal(geometry.hitTest(44 + 159, 26 + 47, viewport)?.key, 'B2')
  assert.equal(geometry.hitTest(44 + 160, 26 + 10, viewport), undefined)
  assert.equal(geometry.hitTest(44 + 10, 26 + 48, viewport), undefined)
  assert.equal(geometry.hitTest(44 + 160, 10, viewport), undefined)
  assert.equal(geometry.hitTest(10, 26 + 48, viewport), undefined)
})

test('fully hidden axes cannot be hit and produce no visible cells', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'hidden-sheet', name: 'Hidden cells', cells: {}, rowCount: 2, columnCount: 2,
    hiddenRows: [0, 1], hiddenColumns: [0, 1],
  })
  const viewport = { width: 500, height: 300, scrollLeft: 0, scrollTop: 0, zoom: 1.5 }
  assert.equal(geometry.hitTest(45, 27, viewport), undefined)
  assert.equal(geometry.hitTest(60, 10, viewport), undefined)
  assert.equal(geometry.hitTest(10, 60, viewport), undefined)
  assert.ok(geometry.panes(viewport).every(pane => pane.rows.length === 0 && pane.columns.length === 0))
})

test('frozen panes preserve hit coordinates after scroll and zoom with hidden axes', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'frozen-sheet', name: 'Frozen sheet', cells: {}, rowCount: 20, columnCount: 20,
    defaultColumnWidth: 80, defaultRowHeight: 24,
    columnWidths: { 0: 100 }, rowHeights: { 0: 30 },
    freeze: { rows: 1, columns: 1 }, hiddenRows: [2], hiddenColumns: [2],
  })
  const viewport = { width: 500, height: 300, scrollLeft: 120, scrollTop: 36, zoom: 1.5 }
  assert.equal(geometry.hitTest(44 + 40, 26 + 20, viewport)?.key, 'A1')
  assert.equal(geometry.hitTest(44 + 150 + 10, 26 + 45 + 10, viewport)?.key, 'D4')
  const rect = geometry.cellRect(3, 3, viewport)
  assert.deepEqual(rect, { x: 194, y: 71, width: 120, height: 36 })
  assert.equal(geometry.hitTest(rect.x + rect.width / 2, rect.y + rect.height / 2, viewport)?.key, 'D4')
})

test('sparse axis indexes ignore malformed hidden indices and normalize custom-size deltas', () => {
  const axis = new AxisIndex(4, -10, { 1: 20 }, [-1, 1.5, Number.NaN, 5])
  assert.equal(axis.offset(1), 0)
  assert.equal(axis.offset(2), 20)
  assert.equal(axis.offset(4), 20)
  assert.deepEqual(axis.visible(0, 100), [1])
})

test('navigation reveals the beginning of an oversized merged cell instead of its far edge', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'wide-sheet', name: 'Wide', cells: {}, rowCount: 100, columnCount: 20,
    defaultColumnWidth: 80, defaultRowHeight: 24,
    merges: ['C3:L20'], freeze: { rows: 1, columns: 1 },
  })
  const viewport = { width: 500, height: 300, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  const scroll = geometry.scrollForCell(2, 2, viewport)
  assert.deepEqual(scroll, { scrollLeft: 80, scrollTop: 24 })
  assert.deepEqual(geometry.cellRect(2, 2, { ...viewport, ...scroll }), { x: 124, y: 50, width: 800, height: 432 })
})

test('navigation scrolls each nonfrozen axis only as far as needed and leaves frozen cells fixed', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'nav-sheet', name: 'Navigation', cells: {}, rowCount: 100, columnCount: 20,
    defaultColumnWidth: 80, defaultRowHeight: 24,
    freeze: { rows: 1, columns: 1 },
  })
  const viewport = { width: 500, height: 300, scrollLeft: 120, scrollTop: 60, zoom: 1.5 }
  assert.deepEqual(geometry.scrollForCell(0, 0, viewport), { scrollLeft: 120, scrollTop: 60 })
  const visibleCell = geometry.hitTest(350, 180, viewport)
  assert.equal(visibleCell.kind, 'cell')
  assert.deepEqual(geometry.scrollForCell(visibleCell.row, visibleCell.column, viewport), { scrollLeft: 120, scrollTop: 60 })
  assert.deepEqual(geometry.scrollForCell(99, 19, viewport), { scrollLeft: 1944, scrollTop: 3326 })
  assert.deepEqual(geometry.scrollForCell(101, 25, viewport), { scrollLeft: 120, scrollTop: 60 })
})

test('visible cell geometry and hit tests agree across fractional zoom, custom dimensions and frozen panes', () => {
  const geometry = new SpreadsheetGeometry({
    id: 'fractional-sheet', name: 'Fractional', cells: {}, rowCount: 20, columnCount: 15,
    defaultColumnWidth: 71.25, defaultRowHeight: 23.5, freeze: { rows: 2, columns: 2 },
    rowHeights: { 3: 44.25, 8: 19.5 }, columnWidths: { 4: 110.5, 10: 21.25 },
    hiddenRows: [1, 6, 19], hiddenColumns: [0, 5, 14],
  })
  let checkedCells = 0
  for (const zoom of [0.5, 0.75, 1, 1.5, 2]) {
    for (const scrollLeft of [0, 13.75, 100, 300]) {
      for (const scrollTop of [0, 17.25, 80, 200]) {
        const viewport = { width: 420, height: 260, scrollLeft, scrollTop, zoom }
        for (const pane of geometry.panes(viewport)) {
          for (const row of pane.rows) {
            for (const column of pane.columns) {
              const rect = geometry.cellRect(row, column, viewport)
              const left = Math.max(rect.x, pane.clip.x), right = Math.min(rect.x + rect.width, pane.clip.x + pane.clip.width)
              const top = Math.max(rect.y, pane.clip.y), bottom = Math.min(rect.y + rect.height, pane.clip.y + pane.clip.height)
              if (left >= right || top >= bottom) continue
              const hit = geometry.hitTest((left + right) / 2, (top + bottom) / 2, viewport)
              assert.deepEqual(hit, { kind: 'cell', row, column, key: cellAddress(row, column) })
              checkedCells += 1
            }
          }
        }
      }
    }
  }
  assert.ok(checkedCells > 1_000)
})
