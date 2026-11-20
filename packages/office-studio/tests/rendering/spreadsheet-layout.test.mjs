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
