import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createCell } from '../src/domain/model.ts'
import { createSparklineSourceChangeDetector } from '../src/features/sparkline/sparkline-source-changes.ts'
import { readSparklineGroups } from '../src/features/sparkline/sparkline-model.ts'
import { sparklineWorkbook, sparklineJson } from './fixtures/sparkline-workbench-fixture.mjs'

const fixture = () => { const before = sparklineWorkbook(), after = sparklineJson(before), group = readSparklineGroups(after.sheets[0]).find(group => group.groupId === 'north')
  return { before, after, group, changed: () => createSparklineSourceChangeDetector(before, after)('data', group) } }

test('stored source value changes are detected while changes outside every series are ignored', () => {
  const f = fixture(); f.after.sheets[0].cells['2:1'].value = 999; assert.equal(f.changed(), false)
  f.after.sheets[0].cells['0:1'].value = 999; assert.equal(f.changed(), true)
})
test('source insertion deletion and type changes are independently observable', () => {
  const f = fixture(); delete f.after.sheets[0].cells['1:4']; assert.equal(f.changed(), true)
  const g = fixture(); g.after.sheets[0].cells['0:1'].dataType = 'text'; assert.equal(g.changed(), true)
  const h = fixture(); delete h.before.sheets[0].cells['0:1']; h.after.sheets[0].cells['0:1'] = createCell(0, 1, 7); assert.equal(h.changed(), true)
})
test('blank style materialization and unrelated display formatting do not change samples', () => {
  const f = fixture(); delete f.before.sheets[0].cells['0:1']; f.after.sheets[0].cells['0:1'] = createCell(0, 1, null)
  f.after.sheets[0].cells['0:2'].display = 'styled'; assert.equal(f.changed(), false)
})
test('formula cached display and cached type changes are visible even with identical formulas', () => {
  const f = fixture(); Object.assign(f.before.sheets[0].cells['0:1'], { value: null, formula: '=B2', display: '4', cachedDataType: 'number' })
  Object.assign(f.after.sheets[0].cells['0:1'], { value: null, formula: '=B2', display: '5', cachedDataType: 'number' }); assert.equal(f.changed(), true)
  f.after.sheets[0].cells['0:1'].display = '4'; f.after.sheets[0].cells['0:1'].cachedDataType = 'text'; assert.equal(f.changed(), true)
})
test('manual row and column hides are clipped to the source ranges', () => {
  const f = fixture(); f.after.sheets[0].hiddenRows = [50]; f.after.sheets[0].columnHiddenSpans = [{ start: 9, end: 12, value: true }]; assert.equal(f.changed(), false)
  f.after.sheets[0].hiddenRows = [1]; assert.equal(f.changed(), true)
  f.after.sheets[0].hiddenRows = []; f.after.sheets[0].hiddenColumns = [2]; assert.equal(f.changed(), true)
})
test('collapsed outline and filter visibility changes update the same source samples', () => {
  const f = fixture(); f.after.sheets[0].outline = { summaryAbove: false, summaryLeft: false, groups: [{ id: 'r', axis: 'row', start: 0, end: 1, collapsed: true }] }
  assert.equal(f.changed(), true)
  delete f.after.sheets[0].outline; f.after.sheets[0].autoFilter = { range: { startRow: 0, endRow: 3, startCol: 1, endCol: 4 }, columns: [], filteredRows: [1] }
  assert.equal(f.changed(), true)
})
test('different reasons for the same effective hidden span do not generate a data change', () => {
  const f = fixture(); f.before.sheets[0].hiddenRows = [1]
  f.after.sheets[0].outline = { summaryAbove: false, summaryLeft: false, groups: [{ id: 'r', axis: 'row', start: 1, end: 1, collapsed: true }] }
  assert.equal(f.changed(), false)
})
test('containHiddenCells removes visibility-only changes from the sample dependency', () => {
  const f = fixture(); f.group.settings.containHiddenCells = true; f.after.sheets[0].hiddenRows = [1]; assert.equal(f.changed(), false)
  f.after.sheets[0].cells['1:1'].value += 1; assert.equal(f.changed(), true)
})
test('quoted cross-sheet references and deleted sources use workbook ownership', () => {
  const f = fixture(); f.group.sourceRanges = ["'Data, O''Neil'!B1:E1"]
  f.after.sheets[0].cells['0:1'].value += 1; assert.equal(f.changed(), false)
  f.after.sheets[1].cells['0:1'].value += 1; assert.equal(f.changed(), true)
  f.after.sheets.pop(); assert.equal(f.changed(), true)
})
test('a million-row sparse source is checked without expanding the requested rectangle', () => {
  const f = fixture(); f.group.sourceRanges = ['B1:B1048576']; f.after.sheets[0].rowHiddenSpans = [{ start: 900_000, end: 1_048_575, value: true }]
  assert.equal(f.changed(), true)
  f.before.sheets[0].rowHiddenSpans = sparklineJson(f.after.sheets[0].rowHiddenSpans); assert.equal(f.changed(), false)
})
