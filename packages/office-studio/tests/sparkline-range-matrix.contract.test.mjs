import assert from 'node:assert/strict'
import { test } from 'vitest'
import { SparklineRangeMatrix } from '../src/features/sparkline/sparkline-range-matrix.ts'

// Deferred contracts for the public ObjectMatrix-compatible range payload.
test('constructor detaches input while matrix getters edit only the returned matrix', () => {
  const source = { 1: { 2: { value: 3 } } }, matrix = new SparklineRangeMatrix(source)
  matrix.getValue(1, 2).value = 9; matrix.getMatrix()[4] = { 5: { value: 6 } }
  assert.deepEqual(source, { 1: { 2: { value: 3 } } }); assert.equal(matrix.getValue(4, 5).value, 6)
})
test('row-priority iteration is ordered and false stops the complete scan', () => {
  const matrix = new SparklineRangeMatrix({ 4: { 8: 'c' }, 1: { 5: 'b', 2: 'a' } }), values = []
  assert.equal(matrix.forValue((row, col, value) => { values.push([row, col, value]); return value !== 'b' }), matrix)
  assert.deepEqual(values, [[1, 2, 'a'], [1, 5, 'b']])
  const rows = []; matrix.forRow((row, columns) => { rows.push([row, columns]) })
  assert.deepEqual(rows, [[1, [2, 5]], [4, [8]]])
})
test('empty public matrix uses upstream length and distinct empty scope conventions', () => {
  const matrix = new SparklineRangeMatrix()
  assert.equal(matrix.getLength(), 1); assert.equal(matrix.getSizeOf(), 0); assert.equal(matrix.hasValue(), false)
  assert.deepEqual(matrix.getRange(), { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })
  assert.deepEqual(matrix.getStartEndScope(), { startRow: -1, endRow: -1, startColumn: -1, endColumn: -1 })
  assert.deepEqual(matrix.getDataRange(), { startRow: 0, endRow: -1, startColumn: 0, endColumn: 0 })
  assert.deepEqual(matrix.toFullArray(), [[undefined]])
})
test('explicit empty rows retain size extent and sparse array rows', () => {
  const matrix = SparklineRangeMatrix.MakeObjectMatrixSize(8)
  assert.equal(matrix.getLength(), 8); assert.equal(matrix.getSizeOf(), 1); assert.equal(matrix.hasValue(), false)
  assert.deepEqual(matrix.toArray()[7], []); assert.deepEqual(matrix.getRealRowRange(), { startRow: 7, endRow: 7 })
  matrix.reset(); assert.equal(matrix.getSizeOf(), 0)
})
test('row swaps preserve values and move a row to an absent destination', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 0: 'a' }, 3: { 0: 'b' } })
  matrix.swapRow(1, 3); assert.equal(matrix.getValue(1, 0), 'b')
  matrix.swapRow(3, 10); assert.equal(matrix.getValue(10, 0), 'a'); assert.equal(matrix.getRow(3), undefined)
})
test('row and column insertion preserve gaps and row identity', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 1: 'a', 3: 'b' }, 4: { 2: 'c' } }), row = matrix.getRow(1)
  matrix.insertRows(2, 3); matrix.insertColumns(2, 2)
  assert.equal(matrix.getRow(1), row); assert.deepEqual(matrix.getMatrix(), { 1: { 1: 'a', 5: 'b' }, 7: { 4: 'c' } })
})
test('axis removal closes gaps, preserves null and retains empty rows', () => {
  const matrix = new SparklineRangeMatrix({ 0: { 0: 'a', 2: null, 5: 'b' }, 3: { 1: 'c' } })
  matrix.removeColumns(1, 2); matrix.removeRows(1, 2)
  assert.deepEqual(matrix.getMatrix(), { 0: { 0: 'a', 3: 'b' }, 1: {} })
})
test('forward row move uses target as the original insertion boundary', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 0: 'a' }, 3: { 0: 'b' }, 5: { 0: 'c' } })
  matrix.moveRows(1, 2, 6)
  assert.deepEqual(matrix.getMatrix(), { 1: { 0: 'b' }, 3: { 0: 'c' }, 4: { 0: 'a' } })
})
test('backward row move retains holes in the moved block', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 0: 'a' }, 4: { 0: 'b' }, 6: { 0: 'c' } })
  matrix.moveRows(4, 2, 0)
  assert.deepEqual(matrix.getMatrix(), { 0: { 0: 'b' }, 3: { 0: 'a' }, 6: { 0: 'c' } })
})
test('column moves follow the same permutation in every sparse row', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 1: 'a', 3: 'b' }, 4: { 2: null, 5: 'c' } })
  matrix.moveColumns(1, 2, 6)
  assert.deepEqual(matrix.getMatrix(), { 1: { 1: 'b', 4: 'a' }, 4: { 3: 'c', 5: null } })
  matrix.moveColumns(4, 2, 0)
  assert.deepEqual(matrix.getMatrix(), { 1: { 0: 'a', 3: 'b' }, 4: { 1: null, 5: 'c' } })
})
test('overlapping and invalid moves reject before changing either axis', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 1: 'a' }, 9: { 8: 'b' } }), before = matrix.clone()
  assert.throws(() => matrix.moveRows(1, 3, 2), RangeError); assert.throws(() => matrix.moveColumns(1, 3, 2), RangeError)
  assert.throws(() => matrix.insertRows(-1, 2), RangeError); assert.throws(() => matrix.removeColumns(1, 1.5), RangeError)
  assert.deepEqual(matrix.clone(), before)
})
test('million-row sparse shifts retain only stored entries', () => {
  const matrix = new SparklineRangeMatrix({ 1_000_000: { 16_000: 'last' }, 2: { 1: 'first' } })
  matrix.insertRows(500_000, 2); matrix.moveRows(1_000_002, 1, 0); matrix.removeColumns(2, 3)
  assert.deepEqual(matrix.getMatrix(), { 0: { 15_997: 'last' }, 3: { 1: 'first' } })
  assert.equal(matrix.getSizeOf(), 2)
})
test('fragment rebases all cells including explicit undefined and keeps value identity', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 2: { value: 1 }, 3: { value: 2 } }, 2: { 2: { value: 3 } } })
  const fragment = matrix.getFragment(1, 2, 2, 3)
  assert.deepEqual(fragment.getMatrix(), { 0: { 0: { value: 1 }, 1: { value: 2 } }, 1: { 0: { value: 3 }, 1: undefined } })
  assert.equal(fragment.getValue(0, 0), matrix.getValue(1, 2)); assert.equal(Object.hasOwn(fragment.getRow(1), 1), true)
})
test('slice keeps absolute indices and clones truthy values only', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 1: 0, 2: false, 3: null, 4: { value: 1 } } })
  const slice = matrix.getSlice(1, 1, 1, 4)
  assert.deepEqual(slice.getMatrix(), { 1: { 4: { value: 1 } } }); assert.notEqual(slice.getValue(1, 4), matrix.getValue(1, 4))
})
test('row slice reports stored cells and preserves upstream row aliasing', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 2: 'a' }, 4: { 3: 'b', 8: undefined }, 9: {} })
  const { sliceData, cellCount } = matrix.getSliceDataAndCellCountByRows(3, 9)
  assert.equal(cellCount, 2); assert.equal(sliceData.getRow(4), matrix.getRow(4)); assert.deepEqual(sliceData.getRow(9), {})
})
test('sparse full and flat arrays share the same row-priority stored values', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 2: 'a', 3: 'b' }, 2: { 2: 'c' } })
  assert.deepEqual(matrix.toNativeArray(), ['a', 'b', 'c']); assert.equal(0 in matrix.toArray(), false)
  assert.deepEqual(matrix.toFullArray(), [[undefined, undefined, undefined, undefined], [undefined, undefined, 'a', 'b'], [undefined, undefined, 'c', undefined]])
})
test('array data rebases to the first stored row and minimum stored column', () => {
  const matrix = new SparklineRangeMatrix({ 2: {}, 4: { 8: 'a' }, 5: { 6: 'b' } })
  assert.deepEqual(matrix.getArrayData(), { 2: { 2: 'a' }, 3: { 0: 'b' } })
  assert.deepEqual(matrix.getDataRange(), { startRow: 2, endRow: 5, startColumn: 6, endColumn: 8 })
})
test('discrete ranges use fixed row-first adjacency semantics', () => {
  const matrix = new SparklineRangeMatrix({ 0: { 0: 'a', 1: 'b' }, 1: { 1: 'c' }, 4: { 5: 'd' } })
  assert.deepEqual(matrix.getDiscreteRanges(), [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, { startRow: 4, endRow: 4, startColumn: 5, endColumn: 5 }])
})
test('merge updates existing cells only and ignores null values', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 1: 'a', 2: 'b' } })
  matrix.merge(new SparklineRangeMatrix({ 1: { 1: 'c', 2: null, 3: 'new' }, 4: { 1: 'new row' } }))
  assert.deepEqual(matrix.getMatrix(), { 1: { 1: 'c', 2: 'b' } })
})
test('concatRows replaces rows at their original indices and keeps row references', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 1: 'a' } }), other = new SparklineRangeMatrix({ 1: { 2: 'b' }, 4: {} })
  matrix.concatRows(other); assert.equal(matrix.getRow(1), other.getRow(1)); assert.deepEqual(matrix.getMatrix(), { 1: { 2: 'b' }, 4: {} })
})
test('clone drops undefined entries, preserves null, and detaches nested DTO values', () => {
  const matrix = new SparklineRangeMatrix({ 1: { 1: undefined, 2: null, 3: { value: 1 } } }), copy = matrix.clone()
  assert.deepEqual(copy, { 1: { 2: null, 3: { value: 1 } } }); copy[1][3].value = 9
  assert.equal(matrix.getValue(1, 3).value, 1)
})
test('invalid keys and unsafe shifts reject without partially moving columns', () => {
  for (const data of [null, [], { '-1': {} }, { '01': {} }, { 1: [] }, { 1: { NaN: 'a' } }]) assert.throws(() => new SparklineRangeMatrix(data))
  const matrix = new SparklineRangeMatrix({ 0: { 1: 'a' }, 1: { [Number.MAX_SAFE_INTEGER - 1]: 'b' } }), before = matrix.clone()
  assert.throws(() => matrix.insertColumns(1, 2), RangeError); assert.deepEqual(matrix.clone(), before)
})
test('dense outputs reject oversized allocation while the sparse slice remains usable', () => {
  const matrix = new SparklineRangeMatrix({ 2_000: { 2_000: 'a' } })
  assert.throws(() => matrix.toFullArray(), /内存预算/); assert.throws(() => matrix.getFragment(0, 2_000, 0, 2_000), /内存预算/)
  assert.deepEqual(matrix.getSlice(0, 2_000, 0, 2_000).getMatrix(), { 2_000: { 2_000: 'a' } })
})
