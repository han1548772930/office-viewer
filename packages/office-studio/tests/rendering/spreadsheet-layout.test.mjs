import assert from 'node:assert/strict'
import { test } from 'node:test'
import { cellAddress, SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'

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
