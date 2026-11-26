import assert from 'node:assert/strict'
import { test } from 'vitest'
import { addSheetOutlineGroup, applySheetOutlineOperation, createSheetOutlineState, getCollapsedOutlineRanges,
  getSheetOutlineHiddenSpans, getSheetOutlineUpdate, outlineGroupDepths, setSheetOutlineCollapsed, setSheetOutlineLevel,
  setSheetOutlineSummaryPosition, validateSheetOutlineState } from '../src/features/sheet-outline/sheet-outline.ts'
import { DimensionOutlineErrorReason as Reason, SheetOutlineError } from '../src/features/sheet-outline/sheet-outline-types.ts'
import { getSheetOutlineLevelControls, getSheetOutlineViewModels } from '../src/features/sheet-outline/sheet-outline-view.ts'

// Drafts for the unified validation stage; not executed during implementation.
const group = (id, axis, start, end, collapsed = false) => ({ id, axis, start, end, collapsed })
const state = (...groups) => ({ ...createSheetOutlineState(), groups })
const rejects = (run, reason) => assert.throws(run, failure => failure instanceof SheetOutlineError && failure.reason === reason)
const protectedRange = ranges => ({ id: 'guard', ownerUserId: 'owner', permissionId: 'guard-permission', editState: 'onlyOwnerCanEdit', allowedUserIds: [], ranges })

test('adjacent grouping retains the original public identity and collapse state', () => {
  const before = state(group('stable', 'row', 2, 4, true))
  const after = addSheetOutlineGroup(before, 'row', 5, 7, 20, 'discarded-new-id')
  assert.deepEqual(after.groups, [group('stable', 'row', 2, 7, true)])
  assert.deepEqual(before.groups, [group('stable', 'row', 2, 4, true)])
})

test('a bridge combines both adjacent groups using the first range identity', () => {
  const before = state(group('left', 'row', 0, 1), group('right', 'row', 4, 6, true), group('column', 'column', 0, 1))
  const after = addSheetOutlineGroup(before, 'row', 2, 3, 20, 'bridge')
  assert.deepEqual(after.groups.filter(item => item.axis === 'row'), [group('left', 'row', 0, 6)])
  assert.equal(after.groups.find(item => item.id === 'column'), before.groups[2])
})

test('equal range layers remain separate and collapse commands synchronize all matching layers', () => {
  const first = addSheetOutlineGroup(createSheetOutlineState(), 'row', 2, 8, 20, 'outer')
  const nested = addSheetOutlineGroup(first, 'row', 2, 8, 20, 'equal-inner')
  assert.equal(outlineGroupDepths(nested.groups).get('equal-inner'), 1)
  const folded = setSheetOutlineCollapsed(nested, 'equal-inner', true)
  assert.deepEqual(folded.groups.map(item => item.collapsed), [true, true])
  assert.deepEqual(setSheetOutlineCollapsed(folded, 'outer', false).groups.map(item => item.collapsed), [false, false])
  assert.deepEqual(nested.groups.map(item => item.collapsed), [false, false])
})

test('row and column nesting have independent eight-layer limits', () => {
  let current = createSheetOutlineState()
  for (const axis of ['row', 'column']) for (let depth = 0; depth < 8; depth += 1) {
    current = addSheetOutlineGroup(current, axis, 10, 19, 50, `${axis}-${depth}`)
  }
  assert.equal(current.groups.length, 16)
  for (const axis of ['row', 'column']) rejects(() => addSheetOutlineGroup(current, axis, 11, 18, 50, `${axis}-9`), Reason.MAX_DEPTH)
  assert.equal(current.groups.length, 16)
})

test('crossings and malformed state produce typed failures before publishing a candidate', () => {
  const before = state(group('existing', 'row', 2, 8))
  rejects(() => addSheetOutlineGroup(before, 'row', 5, 11, 20, 'crossing'), Reason.CROSSING)
  rejects(() => addSheetOutlineGroup(before, 'row', 18, 20, 20, 'outside'), Reason.OUT_OF_BOUNDS)
  rejects(() => validateSheetOutlineState(state({ ...before.groups[0], collapsed: 'false' })), Reason.INVALID_RANGE)
  rejects(() => outlineGroupDepths([{ ...before.groups[0], axis: null }]), Reason.INVALID_RANGE)
  assert.deepEqual(before, state(group('existing', 'row', 2, 8)))
})

test('deterministic level controls show outer, middle and all details through the same model', () => {
  const before = state(group('outer', 'row', 0, 19), group('middle', 'row', 2, 15), group('inner', 'row', 4, 10))
  assert.deepEqual(getSheetOutlineLevelControls(before, 'row').map(item => item.level), [1, 2, 3, 4])
  const second = setSheetOutlineLevel(before, 'row', 2)
  assert.deepEqual(second.groups.map(item => item.collapsed), [false, true, true])
  assert.deepEqual(getSheetOutlineLevelControls(second, 'row').filter(item => item.active).map(item => item.level), [2])
  assert.deepEqual(setSheetOutlineLevel(before, 'row', 1).groups.map(item => item.collapsed), [true, true, true])
  assert.deepEqual(setSheetOutlineLevel(second, 'row', 4).groups.map(item => item.collapsed), [false, false, false])
})

test('full-size collapsed outline visibility stays a compact interval', () => {
  const before = state(group('million-rows', 'row', 0, 1_048_575, true), group('inner', 'row', 100, 999, true))
  assert.deepEqual(getSheetOutlineHiddenSpans(before, 'row', 1_048_576), [{ start: 0, end: 1_048_575, value: true }])
  assert.deepEqual(getSheetOutlineHiddenSpans(before, 'column', 16_384), [])
})

test('raw visibility input rejects invalid extents and truthy non-boolean collapse values', () => {
  rejects(() => getCollapsedOutlineRanges({ count: 0, groups: [] }), Reason.INVALID_RANGE)
  rejects(() => getCollapsedOutlineRanges({ count: 20, groups: [{ ...group('invalid', 'row', 1, 3), collapsed: 1 }] }), Reason.INVALID_RANGE)
  rejects(() => getCollapsedOutlineRanges({ count: 2, groups: [group('outside', 'row', 1, 2, true)] }), Reason.OUT_OF_BOUNDS)
})

test('row and column summary directions change anchors without changing detail membership', () => {
  const before = state(group('row', 'row', 1, 4), group('column', 'column', 1, 4))
  const after = setSheetOutlineSummaryPosition(setSheetOutlineSummaryPosition(before, 'row', true), 'column', true)
  assert.equal(getSheetOutlineViewModels(before, 'row', 10)[0].anchor, 5)
  assert.equal(getSheetOutlineViewModels(after, 'row', 10)[0].anchor, 0)
  assert.equal(getSheetOutlineViewModels(after, 'column', 10)[0].anchor, 0)
  assert.deepEqual(after.groups, before.groups)
})

test('whole-axis groups retain a boundary control when the summary cell is out of bounds', () => {
  const before = state(group('all', 'row', 0, 9, true))
  assert.deepEqual(getSheetOutlineViewModels(before, 'row', 10).map(item => [item.anchor, item.anchorBoundary]), [[null, 10]])
  const above = setSheetOutlineSummaryPosition(before, 'row', true)
  assert.deepEqual(getSheetOutlineViewModels(above, 'row', 10).map(item => [item.anchor, item.anchorBoundary]), [[null, 0]])
})

test('a collapsed ancestor stays effective across an expanded equal-range intermediate layer', () => {
  const before = state(group('folded-outer', 'row', 0, 9, true), group('equal-layer', 'row', 0, 9), group('child', 'row', 2, 4))
  assert.deepEqual(getSheetOutlineViewModels(before, 'row', 20).map(item => item.id), ['folded-outer', 'equal-layer'])
  assert.equal(getSheetOutlineViewModels(before, 'row', 20, true).find(item => item.id === 'child').hiddenByAncestor, true)
})

test('summary settings validate both affected axes and reject the whole change on protected columns', () => {
  const outline = state(group('rows', 'row', 1, 3), group('columns', 'column', 6, 8))
  const sheet = { rowCount: 20, columnCount: 10, outline,
    rangeProtectionRules: [protectedRange([{ startRow: 10, endRow: 10, startCol: 7, endCol: 7 }])] }
  assert.throws(() => getSheetOutlineUpdate(sheet, { kind: 'summary-settings', summaryAbove: true, summaryLeft: true }, 'visitor'), /受保护/)
  assert.equal(outline.summaryAbove, false); assert.equal(outline.summaryLeft, false)
  const allowed = getSheetOutlineUpdate(sheet, { kind: 'summary-settings', summaryAbove: true, summaryLeft: false }, 'visitor')
  assert.equal(allowed.summaryAbove, true); assert.equal(allowed.summaryLeft, false)
})

test('grouping checks a merged adjacent protected group outside the selected new range', () => {
  const outline = state(group('protected-adjacent', 'row', 5, 7))
  const sheet = { rowCount: 20, columnCount: 10, outline,
    rangeProtectionRules: [protectedRange([{ startRow: 7, endRow: 7, startCol: 0, endCol: 9 }])] }
  assert.throws(() => getSheetOutlineUpdate(sheet, { kind: 'group', axis: 'row', start: 2, end: 4, groupId: 'new' }, 'visitor'), /受保护/)
  assert.deepEqual(outline.groups, [group('protected-adjacent', 'row', 5, 7)])
})

test('partial clear, repeated collapse and unchanged settings preserve the no-op boundary', () => {
  const outline = state(group('outer', 'row', 1, 9, true)), sheet = { outline, rowCount: 20, columnCount: 10 }
  assert.equal(applySheetOutlineOperation(outline, 'row', 'remove', 2, 8, 20, ''), outline)
  assert.equal(setSheetOutlineCollapsed(outline, 'outer', true), outline)
  assert.equal(getSheetOutlineUpdate(sheet, { kind: 'clear', axis: 'row', start: 2, end: 8 }, 'user'), undefined)
  assert.equal(getSheetOutlineUpdate(sheet, { kind: 'summary-settings', summaryAbove: false, summaryLeft: false }, 'user'), undefined)
})
