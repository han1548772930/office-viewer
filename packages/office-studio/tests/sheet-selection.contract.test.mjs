import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeSheetSelection, reconcileSheetSelection, resolveVisibleSheetSelection } from '../src/features/sheet-workbench/sheet-selection.ts'
import { nextSheetSelectionCell, planSheetSelectionShortcut, SheetSelectionDirections as Direction } from '../src/features/sheet-workbench/sheet-selection-navigation.ts'
import { normalizeSheetViewState } from '../src/features/sheet-workbench/sheet-workbench.ts'
import { selectionCount } from '../src/rendering/spreadsheet/selection.ts'
import { parseSheetSelectionOperation, parseSheetSelectionShortcut } from '../src/api/sheet-selection-commands.ts'

const range = (startRow, startCol, endRow = startRow, endCol = startCol, rangeType) => ({ startRow, startCol, endRow, endCol,
  ...(rangeType === undefined ? {} : { rangeType }) })
const grid = value => ({ startRow: value.startRow, endRow: value.endRow, startColumn: value.startCol, endColumn: value.endCol,
  ...(value.rangeType === undefined ? {} : { rangeType: value.rangeType }) })
const sheet = (extra = {}) => ({ rowCount: 20, columnCount: 8, defaultRowHeight: 26, defaultColumnWidth: 90, ...extra })
const copy = value => structuredClone(value)

test('selection normalization preserves separate and overlapping rectangles, primary identity and caller ownership', () => {
  const ranges = [range(0, 0, 2, 2), range(1, 1, 3, 3), range(10, 5)], options = { activeRangeIndex: 0, primary: { row: 1, column: 1 } }
  const before = copy({ ranges, options }), state = normalizeSheetSelection(ranges, sheet(), options)
  assert.deepEqual(state, { ranges, activeRangeIndex: 0, primary: { row: 1, column: 1 } })
  state.ranges[0].endRow = 9; state.primary.column = 7
  assert.deepEqual({ ranges, options }, before)
  assert.equal(normalizeSheetSelection(ranges, sheet(), { primary: { row: 1, column: 1 } }).activeRangeIndex, 1)
  assert.equal(normalizeSheetSelection(ranges, sheet()).activeRangeIndex, 2)
})

test('invalid ranges, extents, active indexes and primary coordinates are rejected before a state can be published', () => {
  for (const ranges of [[range(-1, 0)], [range(0, 0, 20, 0)], [range(0, 8)], [range(2, 1, 1, 1)], [range(0.5, 0)],
    [{ ...range(0, 0), rangeType: 4 }], [null], {}]) assert.throws(() => normalizeSheetSelection(ranges, sheet()))
  for (const extent of [sheet({ rowCount: 0 }), sheet({ columnCount: 16_385 }), sheet({ rowCount: Infinity })]) {
    assert.throws(() => normalizeSheetSelection([range(0, 0)], extent), RangeError)
  }
  for (const options of [{ activeRangeIndex: 1 }, { activeRangeIndex: -1 }, { primary: { row: 0, column: 2 } },
    { activeRangeIndex: 0, primary: { row: 2, column: 2 } }, { primary: { row: NaN, column: 0 } }]) {
    assert.throws(() => normalizeSheetSelection([range(0, 0)], sheet(), options), RangeError)
  }
  assert.deepEqual(normalizeSheetSelection([], sheet(), { activeRangeIndex: -1 }), { ranges: [], activeRangeIndex: -1 })
  assert.throws(() => normalizeSheetSelection([], sheet(), { primary: { row: 0, column: 0 } }), RangeError)
})

test('structural reconciliation updates whole-axis intent and clamps the saved primary without merging ranges', () => {
  const before = normalizeSheetSelection([range(15, 0, 18, 7, 1), range(0, 6, 19, 7, 2)], sheet(),
    { activeRangeIndex: 0, primary: { row: 17, column: 6 } })
  const saved = copy(before), after = reconcileSheetSelection(before, sheet({ rowCount: 10, columnCount: 5 }))
  assert.deepEqual(after, { ranges: [range(9, 0, 9, 4, 1), range(0, 4, 9, 4, 2)], activeRangeIndex: 0, primary: { row: 9, column: 4 } })
  assert.deepEqual(before, saved)
})

test('large compact manual hides combine with zero sizes, filtering and outline while the selection stays compact', () => {
  const source = sheet({ rowCount: 1_048_576, columnCount: 16_384,
    rowHiddenSpans: [{ start: 0, end: 900_000, value: true }], columnHiddenSpans: [{ start: 0, end: 16_000, value: true }],
    rowHeights: { 900_001: 0 }, rowAutoHeightSpans: [{ start: 900_001, end: 900_001, value: false }], filteredRows: [900_002],
    outline: { summaryAbove: false, summaryLeft: false, groups: [{ id: 'collapsed', axis: 'row', start: 900_003, end: 900_004, collapsed: true }] } })
  const before = copy(source), state = normalizeSheetSelection([range(0, 0, 1_048_575, 16_383, 3)], source)
  const visible = resolveVisibleSheetSelection(state, source, [900_005])
  assert.deepEqual(visible.ranges, state.ranges)
  assert.deepEqual(visible.primary, { row: 900_006, column: 16_001 })
  assert.deepEqual(source, before)
  assert.equal(source.rowHiddenSpans.length, 1)
  assert.equal(source.hiddenRows, undefined)
})

test('visibility moves the primary to another selected rectangle before replacing only the invisible active rectangle', () => {
  const source = sheet({ hiddenRows: [0, 1, 2] })
  const state = normalizeSheetSelection([range(0, 0, 2, 2), range(10, 5)], source, { activeRangeIndex: 0 })
  assert.deepEqual(resolveVisibleSheetSelection(state, source), { ranges: state.ranges, activeRangeIndex: 1, primary: { row: 10, column: 5 } })
  const hidden = normalizeSheetSelection([range(0, 0, 0, 7, 1), range(2, 0, 2, 7, 1)], source, { activeRangeIndex: 0 })
  assert.deepEqual(resolveVisibleSheetSelection(hidden, source), {
    ranges: [range(3, 0, 3, 7, 1), range(2, 0, 2, 7, 1)], activeRangeIndex: 0, primary: { row: 3, column: 0 },
  })
  const all = sheet({ rowHiddenSpans: [{ start: 0, end: 19, value: true }] })
  assert.deepEqual(resolveVisibleSheetSelection(normalizeSheetSelection([range(0, 0)], all), all), { ranges: [], activeRangeIndex: -1 })
})

test('view restoration retains overlap and active range index; selection counts use the union without enumerating axes', () => {
  const ranges = [range(0, 0, 2, 2), range(1, 1, 3, 3)].map(grid)
  const state = { key: 'B2', anchor: 'A1', ranges, zoom: 125, activeRangeIndex: 0 }, before = copy(state)
  assert.deepEqual(normalizeSheetViewState(sheet(), state), state)
  assert.equal(selectionCount(ranges), 14)
  assert.deepEqual(state, before)
  assert.equal(selectionCount([grid(range(0, 0, 1_048_575, 16_383)), grid(range(1, 1))]), 1_048_576 * 16_384)
  assert.deepEqual(normalizeSheetViewState(sheet(), { ranges: [], activeRangeIndex: -1 }),
    { key: 'A1', anchor: 'A1', ranges: [], activeRangeIndex: -1, zoom: 100 })
})

test('directional primary traversal agrees with an independently enumerated grid for all four directions', () => {
  const source = sheet({ rowCount: 5, columnCount: 5, hiddenRows: [2], hiddenColumns: [2], mergedRanges: [range(0, 0, 1, 1)] })
  const ranges = [range(0, 0, 3, 3), range(4, 4)]
  for (const direction of [Direction.UP, Direction.RIGHT, Direction.DOWN, Direction.LEFT]) {
    const horizontal = direction === Direction.RIGHT || direction === Direction.LEFT
    const reverse = direction === Direction.UP || direction === Direction.LEFT
    const visits = ranges.flatMap((range, index) => {
      const cells = []
      for (let row = range.startRow; row <= range.endRow; row++) for (let column = range.startCol; column <= range.endCol; column++) {
        if (source.hiddenRows.includes(row) || source.hiddenColumns.includes(column)) continue
        const merge = source.mergedRanges.find(merge => row >= merge.startRow && row <= merge.endRow && column >= merge.startCol && column <= merge.endCol)
        if (merge && (row !== merge.startRow || column !== merge.startCol)) continue
        cells.push({ row, column, index, range: merge ?? { startRow: row, endRow: row, startCol: column, endCol: column } })
      }
      return cells.sort((a, b) => horizontal ? a.row - b.row || a.column - b.column : a.column - b.column || a.row - b.row)
    })
    for (let at = 0; at < visits.length; at++) {
      const current = visits[at], expected = visits[(at + (reverse ? -1 : 1) + visits.length) % visits.length]
      const state = normalizeSheetSelection(ranges, source, { activeRangeIndex: current.index, primary: { row: current.row, column: current.column } })
      const next = nextSheetSelectionCell(state, direction, source)
      assert.deepEqual(next, { range: expected.range, primary: { row: expected.row, column: expected.column }, activeRangeIndex: expected.index }, `direction ${direction}, cell ${current.row}:${current.column}`)
    }
  }
})

test('navigation skips a merged interior spanning a million rows and crosses separate selected rectangles', () => {
  const source = sheet({ rowCount: 1_048_576, mergedRanges: [range(0, 0, 1_000_000, 1)] })
  const state = normalizeSheetSelection([range(0, 0, 1_000_001, 1), range(1_048_575, 7)], source)
  state.activeRangeIndex = 0; state.primary = { row: 0, column: 0 }
  assert.deepEqual(nextSheetSelectionCell(state, Direction.RIGHT, source), {
    range: range(1_000_001, 0), primary: { row: 1_000_001, column: 0 }, activeRangeIndex: 0,
  })
  const single = normalizeSheetSelection([range(0, 0), range(10, 5)], sheet(), { activeRangeIndex: 0 })
  assert.deepEqual(nextSheetSelectionCell(single, Direction.LEFT, sheet()), { range: range(10, 5), primary: { row: 10, column: 5 }, activeRangeIndex: 1 })
  assert.equal(nextSheetSelectionCell(normalizeSheetSelection([], source), Direction.RIGHT, source), null)
  assert.throws(() => nextSheetSelectionCell(single, 9, sheet()), TypeError)
})

test('selection command envelopes preserve identity and accept fixed-version nested selection primaries', () => {
  const input = { unitId: 'book', subUnitId: 'sheet', selections: [
    { range: grid(range(0, 0, 2, 2)), primary: { actualRow: 1, actualColumn: 2 }, style: null },
    { range: grid(range(8, 4)), primary: null, style: { color: '#123456' } },
  ] }
  const before = copy(input), request = parseSheetSelectionOperation(input)
  assert.deepEqual(request.options, { activeRangeIndex: 0, primary: { row: 1, column: 2 } })
  assert.deepEqual(request.ranges, input.selections.map(value => ({ ...value.range, unitId: 'book', sheetId: 'sheet' })))
  assert.deepEqual(input, before)
  for (const bad of [undefined, { ...input, unitId: '' }, { ...input, sheetId: 'other' },
    { unitId: 'book', sheetId: 'sheet', ranges: [{ ...grid(range(0, 0)), unitId: 'foreign' }] },
    { unitId: 'book', sheetId: 'sheet', ranges: [grid(range(0.5, 0))] }, { unitId: 'book', sheetId: 'sheet', ranges: {} }]) {
    assert.throws(() => parseSheetSelectionOperation(bad))
  }
})

test('an isolated cell follows sheet boundaries and Enter resumes the first Tab column', () => {
  const source = sheet({ hiddenColumns: [1], hiddenRows: [1] })
  const initial = normalizeSheetSelection([range(0, 0)], source), before = copy(initial)
  const first = planSheetSelectionShortcut(initial, Direction.RIGHT, 'Tab', source)
  assert.deepEqual(first.selection.primary, { row: 0, column: 2 })
  assert.deepEqual(first.tabStart, initial)
  const second = planSheetSelectionShortcut(first.selection, Direction.RIGHT, 'Tab', source, first.tabStart)
  assert.deepEqual(second.selection.primary, { row: 0, column: 3 })
  const enter = planSheetSelectionShortcut(second.selection, Direction.DOWN, 'Enter', source, second.tabStart)
  assert.deepEqual(enter.selection.primary, { row: 2, column: 0 })
  assert.equal(enter.tabStart, undefined)
  assert.deepEqual(initial, before)
  first.tabStart.primary.column = 7
  assert.deepEqual(initial, before)
  const leftEdge = planSheetSelectionShortcut(initial, Direction.LEFT, 'Tab', source)
  assert.deepEqual(leftEdge.selection, initial)
  assert.equal(leftEdge.tabStart, undefined)
  assert.equal(planSheetSelectionShortcut(normalizeSheetSelection([], source), Direction.RIGHT, 'Tab', source), null)
  for (const [direction, key] of [[Direction.DOWN, 'Tab'], [Direction.RIGHT, 'Enter'], [7, 'Enter'], [Direction.RIGHT, 'x']]) {
    assert.throws(() => planSheetSelectionShortcut(initial, direction, key, source), TypeError)
  }
})

test('keyboard navigation retains multiple rectangles and cycles both row and column order after edits', () => {
  const source = sheet(), state = normalizeSheetSelection([range(0, 0, 1, 1), range(5, 5)], source,
    { activeRangeIndex: 0, primary: { row: 0, column: 1 } })
  const before = copy(state)
  const tab = planSheetSelectionShortcut(state, Direction.RIGHT, 'Tab', source)
  assert.deepEqual(tab.selection, { ranges: state.ranges, activeRangeIndex: 0, primary: { row: 1, column: 0 } })
  assert.equal(tab.tabStart, undefined)
  const enter = planSheetSelectionShortcut(tab.selection, Direction.DOWN, 'Enter', source)
  assert.deepEqual(enter.selection.primary, { row: 0, column: 1 })
  const last = normalizeSheetSelection(state.ranges, source, { activeRangeIndex: 0, primary: { row: 1, column: 1 } })
  const other = planSheetSelectionShortcut(last, Direction.DOWN, 'Enter', source)
  assert.deepEqual(other.selection, { ranges: state.ranges, activeRangeIndex: 1, primary: { row: 5, column: 5 } })
  assert.deepEqual(planSheetSelectionShortcut(other.selection, Direction.LEFT, 'Tab', source).selection.primary, { row: 1, column: 1 })
  assert.deepEqual(state, before)
})

test('merged navigation visits a visible representative when its original row and column are hidden', () => {
  const source = sheet({ hiddenRows: [0], hiddenColumns: [0], mergedRanges: [range(0, 0, 2, 2)] })
  const state = normalizeSheetSelection([range(0, 0, 3, 3)], source, { primary: { row: 1, column: 1 } })
  const right = nextSheetSelectionCell(state, Direction.RIGHT, source)
  assert.deepEqual(right.primary, { row: 1, column: 3 })
  assert.deepEqual(nextSheetSelectionCell({ ...state, primary: right.primary }, Direction.LEFT, source),
    { range: range(0, 0, 2, 2), primary: { row: 1, column: 1 }, activeRangeIndex: 0 })
  const isolated = normalizeSheetSelection([range(0, 0, 2, 2)], source, { primary: { row: 1, column: 1 } })
  assert.deepEqual(planSheetSelectionShortcut(isolated, Direction.DOWN, 'Enter', source).selection.primary, { row: 3, column: 1 })
})

test('shortcut commands accept fixed-version keycodes and reject mismatched key and worksheet parameters', () => {
  const target = { unitId: 'book', sheetId: 'sheet' }
  assert.deepEqual(parseSheetSelectionShortcut({ direction: Direction.DOWN, keycode: 13 }, target),
    { ...target, direction: Direction.DOWN, key: 'Enter' })
  assert.deepEqual(parseSheetSelectionShortcut({ unitId: 'second', subUnitId: 'other', direction: Direction.LEFT, key: 'Tab', keycode: 9 }, target),
    { unitId: 'second', sheetId: 'other', direction: Direction.LEFT, key: 'Tab' })
  for (const input of [null, {}, { direction: Direction.RIGHT, key: 'Enter' }, { direction: Direction.RIGHT, key: 'Tab', keycode: 13 },
    { sheetId: 'other', subUnitId: 'sheet', direction: Direction.RIGHT, key: 'Tab' }]) assert.throws(() => parseSheetSelectionShortcut(input, target))
})
