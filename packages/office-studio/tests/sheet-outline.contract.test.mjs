import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  addSheetOutlineGroup,
  createSheetOutlineState,
  getCollapsedOutlineIndexes,
  getOutlineDepth,
  getOutlineLevels,
  removeSheetOutlineGroup,
  setSheetOutlineCollapsed,
  setSheetOutlineSummaryPosition,
} from '../src/features/sheet-outline/sheet-outline.ts'
import { hiddenSheetRows } from '../src/features/sheet-workbench/sheet-row-visibility.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createBlankWorkbook } from '../src/domain/model.ts'

test('outline groups support nested levels, collapse and immutable updates', () => {
  let state = createSheetOutlineState()
  state = addSheetOutlineGroup(state, 'row', 1, 8, 12, 'outer')
  state = addSheetOutlineGroup(state, 'row', 3, 5, 12, 'inner')
  assert.equal(getOutlineDepth({ count: 12, groups: state.groups }, 4), 2)
  assert.deepEqual(getOutlineLevels({ count: 12, groups: state.groups }), [[[1, 8]], [[3, 5]]])
  const collapsed = setSheetOutlineCollapsed(state, 'outer', true)
  assert.deepEqual([...getCollapsedOutlineIndexes({ count: 12, groups: collapsed.groups })], [1, 2, 3, 4, 5, 6, 7, 8])
  assert.equal(state.groups[0].collapsed, false)
  assert.deepEqual([...getCollapsedOutlineIndexes({ count: 12, groups: setSheetOutlineCollapsed(collapsed, 'outer', false).groups })], [])
})

test('summary position is persisted independently of grouped detail ranges', () => {
  const below = addSheetOutlineGroup(createSheetOutlineState(), 'row', 1, 3, 8, 'r1')
  const above = setSheetOutlineSummaryPosition(below, 'row', true)
  assert.equal(below.summaryAbove, false)
  assert.equal(above.summaryAbove, true)
  const folded = setSheetOutlineCollapsed(above, 'r1', true)
  assert.deepEqual([...getCollapsedOutlineIndexes({ count: 8, groups: folded.groups })], [1, 2, 3])
})

test('outline, filter, manual-hidden and zero-height causes combine without overwriting each other', () => {
  const collapsed = setSheetOutlineCollapsed(
    addSheetOutlineGroup(createSheetOutlineState(), 'row', 2, 4, 10, 'group'),
    'group',
    true,
  )
  const outlineHiddenRows = [...getCollapsedOutlineIndexes({ count: 10, groups: collapsed.groups })]
  const hidden = hiddenSheetRows({ rowCount: 10, hiddenRows: [1, 3], filteredRows: [5], outlineHiddenRows, rowHeights: { 7: 0 } })
  assert.deepEqual([...hidden].sort((a, b) => a - b), [1, 2, 3, 4, 5, 7])
  assert.deepEqual(outlineHiddenRows, [2, 3, 4])
})

test('outline state round-trips through the editor view while derived visibility stays separate', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.rowCount = 10
  sheet.columnCount = 8
  sheet.hiddenRows = [1]
  sheet.outline = setSheetOutlineCollapsed(
    addSheetOutlineGroup(createSheetOutlineState(), 'row', 2, 4, 10, 'rows-1'),
    'rows-1',
    true,
  )
  const view = toExcelWorkbookView(workbook)
  assert.deepEqual(view.sheets[0].outlineHiddenRows, [2, 3, 4])
  assert.deepEqual(view.sheets[0].hiddenRows, [1])
  view.sheets[0].outline.groups[0].collapsed = false
  const next = applyExcelWorkbookView(workbook, view)
  assert.deepEqual(next.sheets[0].outline.groups[0].collapsed, false)
  assert.deepEqual(next.sheets[0].hiddenRows, [1])
  assert.deepEqual(workbook.sheets[0].outline.groups[0].collapsed, true)
})

test('outline group boundaries and crossings reject invalid state', () => {
  let state = addSheetOutlineGroup(createSheetOutlineState(), 'column', 0, 2, 6, 'outer')
  assert.throws(() => addSheetOutlineGroup(state, 'column', 2, 4, 6, 'crossing'), /nested or disjoint/)
  assert.throws(() => addSheetOutlineGroup(state, 'row', 0, 0, 6, 'bad'), /at least two/)
  assert.throws(() => addSheetOutlineGroup(state, 'row', 4, 8, 6, 'outside'), /outside its axis/)
  assert.throws(() => setSheetOutlineCollapsed(state, 'missing', true), /not found/)
  state = removeSheetOutlineGroup(state, 'outer')
  assert.deepEqual(state.groups, [])
})
