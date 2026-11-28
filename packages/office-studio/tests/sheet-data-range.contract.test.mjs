import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createCell } from '../src/domain/model.ts'
import { expandSheetDataRange, prepareSmartSheetFilterRange, sheetCellHasData } from '../src/features/sheet-sort-filter/sheet-data-range.ts'
import { prepareDataPanelSortRange } from '../src/features/sheet-sort-filter/sheet-data-panel-state.ts'

// Implementation-phase drafts, not executed.
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
function sheet(entries = []) {
  const model = createBlankSheet('Data'); model.rowCount = 60; model.columnCount = 20
  for (const [row, col, value] of entries) model.cells[`${row}:${col}`] = createCell(row, col, value)
  return model
}

test('content uses zero, false and public body presence, excluding style-only cells and formulas with empty caches', () => {
  assert.equal(sheetCellHasData(createCell(0, 0, 0)), true)
  assert.equal(sheetCellHasData(createCell(0, 0, false)), true)
  for (const value of ['', null]) assert.equal(sheetCellHasData(createCell(0, 0, value)), false)
  assert.equal(sheetCellHasData({ ...createCell(0, 0, null), formula: '=1', display: '1', richText: [] }), false)
  assert.equal(sheetCellHasData({ ...createCell(0, 0, null), metadata: { publicBody: {} } }), true)
  assert.equal(sheetCellHasData({ ...createCell(0, 0, null), metadata: { publicBody: { body: { dataStream: '\r\n' } } } }), true)
  assert.equal(sheetCellHasData({ ...createCell(0, 0, null), richText: [{ text: '' }] }), true)
})

test('a blank seed may join adjacent or diagonal data and stops before a separate island', () => {
  const model = sheet([[1, 1, 0], [2, 2, false], [3, 2, 'last'], [8, 8, 'island']]), original = structuredClone(model)
  assert.deepEqual(expandSheetDataRange(model, range(2, 1)), range(1, 1, 3, 2))
  assert.deepEqual(expandSheetDataRange(model, range(10, 10)), range(10, 10))
  assert.deepEqual(model, original)
})

test('smart single-cell filters expand in all directions while single-row filters only expand down', () => {
  const model = sheet([[1, 1, 'left'], [1, 2, 'header'], [1, 3, 'right'], [0, 2, 'above'], [2, 2, 0], [3, 2, false]])
  assert.deepEqual(prepareSmartSheetFilterRange(model, range(2, 2)), range(0, 1, 3, 3))
  assert.deepEqual(prepareSmartSheetFilterRange(model, range(1, 2, 1, 3)), range(1, 2, 3, 3))
  assert.deepEqual(prepareSmartSheetFilterRange(model, range(1, 2, 2, 2)), range(1, 2, 2, 2))
  assert.throws(() => prepareSmartSheetFilterRange(model, range(10, 10)), /标题行和数据行/)
})

test('merge interiors use anchor data and an adjacent empty merge does not independently enlarge the region', () => {
  const model = sheet([[1, 1, 'head'], [1, 2, 'interior garbage'], [3, 1, 2], [6, 1, 'separate']])
  model.mergedRanges = [range(1, 1, 2, 2), range(4, 1, 5, 2)]
  assert.deepEqual(expandSheetDataRange(model, range(3, 1)), range(1, 1, 3, 2))
  model.cells['1:1'].value = null
  assert.deepEqual(expandSheetDataRange(model, range(3, 1)), range(3, 1))
})

test('explicit directions, worksheet bounds and malformed coordinates reject without mutating the seed', () => {
  const model = sheet([[0, 0, 'top'], [1, 1, 'diagonal'], [2, 0, 'down']]), source = range(0, 0)
  assert.deepEqual(expandSheetDataRange(model, source, { down: true }), source)
  assert.deepEqual(expandSheetDataRange(model, source), range(0, 0, 2, 1))
  assert.deepEqual(source, range(0, 0))
  for (const directions of [null, [], { down: 'true' }, { northeast: true }]) assert.throws(() => expandSheetDataRange(model, source, directions), TypeError)
  assert.throws(() => expandSheetDataRange(model, range(60, 0)), RangeError)
  assert.throws(() => expandSheetDataRange(model, range(-1, 0)), RangeError)
})

test('custom sort keep/extend choices preserve the original active cell and leave all content untouched', () => {
  const model = sheet([[0, 0, 'label'], [0, 1, 'value'], [1, 0, 'a'], [1, 1, 2], [2, 0, 'b'], [2, 1, 1]])
  const source = range(0, 1, 2, 1), primary = { row: 1, column: 1 }, original = structuredClone(model)
  assert.deepEqual(prepareDataPanelSortRange(model, source, 'keep', primary), source)
  assert.deepEqual(prepareDataPanelSortRange(model, source, 'extend', primary), range(0, 0, 2, 1))
  assert.throws(() => prepareDataPanelSortRange(model, source, 'cancel', primary), TypeError)
  assert.throws(() => prepareDataPanelSortRange(model, source, 'extend', { row: 1, column: 0 }), RangeError)
  assert.throws(() => prepareDataPanelSortRange(model, range(0, 1), 'keep', { row: 0, column: 1 }), RangeError)
  assert.deepEqual(source, range(0, 1, 2, 1)); assert.deepEqual(primary, { row: 1, column: 1 }); assert.deepEqual(model, original)
})
