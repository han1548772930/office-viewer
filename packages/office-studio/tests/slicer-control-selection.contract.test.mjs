import assert from 'node:assert/strict'
import { test } from 'vitest'
import { slicerControlItemSelection } from '../src/features/slicer/slicer-control-selection.ts'

// Deferred input/source-order contracts, shared by native pointer, Canvas keyboard and Ant buttons.
const items = ['North', 'South', 'West', 'East']
const input = (index, extra = {}) => ({ index, value: items[index], ctrlKey: false, metaKey: false, shiftKey: false, multiSelect: false, ...extra })

test('ordinary click selects one item without changing the source or prior selection', () => {
  const selected = ['West'], before = structuredClone([items, selected])
  assert.deepEqual(slicerControlItemSelection(items, selected, input(1)), ['South'])
  assert.deepEqual(slicerControlItemSelection(items, undefined, input(0)), ['North'])
  assert.deepEqual([items, selected], before)
})
for (const modifier of ['ctrlKey', 'metaKey', 'multiSelect']) test(`${modifier} toggles the same native selection including its implicit all state`, () => {
  assert.deepEqual(slicerControlItemSelection(items, undefined, input(1, { [modifier]: true })), ['North', 'West', 'East'])
  assert.deepEqual(slicerControlItemSelection(items, ['North'], input(1, { [modifier]: true })), ['North', 'South'])
  assert.deepEqual(slicerControlItemSelection(items, ['North', 'South'], input(1, { [modifier]: true })), ['North'])
})
for (const [first, last] of [[0, 2], [2, 0]]) test(`Shift from ${first} to ${last} selects an inclusive source-order interval`, () => {
  assert.deepEqual(slicerControlItemSelection(items, ['East'], input(last, { anchorIndex: first, shiftKey: true, ctrlKey: true })), ['North', 'South', 'West'])
})
test('selecting every item normalizes to null and removing the last selected item retains the empty filter', () => {
  assert.equal(slicerControlItemSelection(items, ['North', 'South', 'West'], input(3, { multiSelect: true })), null)
  assert.equal(slicerControlItemSelection(items, ['East'], input(3, { anchorIndex: 0, shiftKey: true })), null)
  assert.deepEqual(slicerControlItemSelection(items, ['North'], input(0, { ctrlKey: true })), [])
})
test('obsolete selected values and duplicates are reconciled before a current-source toggle', () => {
  assert.deepEqual(slicerControlItemSelection(items, ['North', 'North', 'removed'], input(2, { metaKey: true })), ['North', 'West'])
})
for (const [name, change] of [
  ['changed value', { value: 'Changed' }], ['negative index', { index: -1 }], ['out of range index', { index: items.length }],
  ['fractional index', { index: .5 }], ['negative anchor', { anchorIndex: -1 }],
  ['removed anchor', { anchorIndex: items.length }], ['nonfinite anchor', { anchorIndex: Infinity }],
]) test(`a ${name} cannot apply a stale item selection`, () => {
  assert.throws(() => slicerControlItemSelection(items, ['North'], input(1, change)), /项目|锚点/)
})
