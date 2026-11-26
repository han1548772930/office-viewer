import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { createSheetOutlineState } from '../src/features/sheet-outline/sheet-outline.ts'
import { cloneSheetOutlineState, DimensionOutlineErrorReason as Reason, SheetOutlineError } from '../src/features/sheet-outline/sheet-outline-types.ts'
import { transformSheetOutlineForAxisEdit as transform } from '../src/features/sheet-outline/sheet-outline-structure.ts'
import { deserializeSheetOutlineResource, deserializeWorkbookOutlineResource, getSheetDimensionOutlines, serializeSheetOutlineResource,
  serializeWorkbookOutlineResource, sheetOutlinePublicAddRange } from '../src/features/sheet-outline/sheet-outline-resources.ts'

// Static drafts only. Office files and the runtime remain unverified.
const group = (id, axis, start, end, collapsed = false) => ({ id, axis, start, end, collapsed })
const state = (...groups) => ({ ...createSheetOutlineState(), groups })
const extent = { rowCount: 20, columnCount: 10 }, scope = { ...extent, unitId: 'target-book', subUnitId: 'target-sheet' }
const copy = value => JSON.parse(JSON.stringify(value))
const rejects = (run, reason) => assert.throws(run, failure => failure instanceof SheetOutlineError && failure.reason === reason)

test('insertions before, inside and after a group follow the shared axis boundaries', () => {
  const before = state(group('stable', 'row', 2, 8)), column = group('column', 'column', 1, 3)
  before.groups.push(column)
  assert.deepEqual(transform(before, { axis: 'row', kind: 'insert', start: 2, count: 2 }, extent).groups[0], group('stable', 'row', 4, 10))
  assert.deepEqual(transform(before, { axis: 'row', kind: 'insert', start: 5, count: 2 }, extent).groups[0], group('stable', 'row', 2, 10))
  assert.equal(transform(before, { axis: 'row', kind: 'insert', start: 9, count: 2 }, extent), before)
  assert.equal(transform(before, { axis: 'row', kind: 'insert', start: 2, count: 2 }, extent).groups[1], column)
})

test('deleting detail axes shrinks or removes a group without changing surviving identity', () => {
  const before = state(group('stable', 'row', 2, 8, true))
  assert.deepEqual(transform(before, { axis: 'row', kind: 'delete', start: 2, count: 2 }, extent).groups, [group('stable', 'row', 2, 6, true)])
  assert.deepEqual(transform(before, { axis: 'row', kind: 'delete', start: 2, count: 7 }, extent).groups, [])
  assert.deepEqual(before.groups, [group('stable', 'row', 2, 8, true)])
})

test('a whole group moves intact and a permutation inside the group remains a no-op', () => {
  const before = state(group('stable', 'row', 2, 7, true))
  assert.deepEqual(transform(before, { axis: 'row', kind: 'move', start: 2, count: 6, destination: 15 }, extent).groups,
    [group('stable', 'row', 9, 14, true)])
  assert.equal(transform(before, { axis: 'row', kind: 'move', start: 3, count: 2, destination: 6 }, extent), before)
})

test('a move that splits an outline rejects every resource update without mutating the source', () => {
  const before = state(group('stable', 'row', 2, 7), group('other-axis', 'column', 1, 3)), saved = copy(before)
  rejects(() => transform(before, { axis: 'row', kind: 'move', start: 3, count: 2, destination: 12 }, extent), Reason.MOVE_SPLITS_OUTLINE)
  assert.deepEqual(before, saved)
})

test('delete-created equal ranges keep the first identity and OR every collapse cause', () => {
  const before = state(group('outer', 'row', 2, 8), group('inner', 'row', 3, 8, true))
  assert.deepEqual(transform(before, { axis: 'row', kind: 'delete', start: 2, count: 1 }, extent).groups, [group('outer', 'row', 2, 7, true)])
  assert.equal(before.groups[0].collapsed, false)
})

test('copy creates independent identities and defensive nested extension data', () => {
  const original = { ...group('source', 'row', 2, 4, true), extension: { office: { labels: ['source'] } } }, before = state(original)
  const after = transform(before, { axis: 'row', kind: 'copy', start: 1, count: 5, destination: 10 }, extent, undefined, () => 'copy')
  assert.deepEqual(after.groups.map(item => [item.id, item.start, item.end, item.collapsed]), [['source', 2, 4, true], ['copy', 11, 13, true]])
  after.groups[1].extension.office.labels.push('changed')
  assert.deepEqual(original.extension.office.labels, ['source'])
})

test('copy does not invent partial groups outside its fully contained source range', () => {
  const before = state(group('partial', 'row', 2, 6))
  const after = transform(before, { axis: 'row', kind: 'copy', start: 1, count: 4, destination: 10 }, extent, undefined, () => 'unused')
  assert.deepEqual(after.groups, before.groups)
})

test('invalid copy identities reject the detached candidate', () => {
  const before = state(group('source', 'row', 2, 4)), saved = copy(before)
  rejects(() => transform(before, { axis: 'row', kind: 'copy', start: 1, count: 5, destination: 10 }, extent, undefined, () => 'source'), Reason.INVALID_RANGE)
  assert.deepEqual(before, saved)
})

test('insertion at the Office limit rejects truncation of an existing group', () => {
  const before = state(group('last-rows', 'row', 1_048_574, 1_048_575))
  rejects(() => transform(before, { axis: 'row', kind: 'insert', start: 0, count: 1 }, { rowCount: 1_048_576, columnCount: 10 }), Reason.OUT_OF_BOUNDS)
  assert.equal(before.groups[0].end, 1_048_575)
})

test('public add converts quantities to inclusive ranges and rejects zero or unsafe counts', () => {
  assert.deepEqual(sheetOutlinePublicAddRange('row', 3, 2, 20), { axis: 'row', start: 3, end: 4 })
  rejects(() => sheetOutlinePublicAddRange('row', 3, 0, 20), Reason.INVALID_RANGE)
  rejects(() => sheetOutlinePublicAddRange('row', 19, 2, 20), Reason.OUT_OF_BOUNDS)
  rejects(() => sheetOutlinePublicAddRange('row', Number.MAX_SAFE_INTEGER, 2, 20), Reason.INVALID_RANGE)
})

test('public snapshots and cloned sheets preserve nested extensions without shared references', () => {
  const before = state({ ...group('source', 'row', 2, 4), extension: { values: [1, { name: 'kept' }] } })
  const snapshot = getSheetDimensionOutlines(before, scope), cloned = cloneSheetOutlineState(before, () => 'clone')
  snapshot[0].extension.values[1].name = 'snapshot edit'; cloned.groups[0].extension.values[0] = 99
  assert.deepEqual(before.groups[0].extension, { values: [1, { name: 'kept' }] })
  assert.equal(snapshot[0].unitId, scope.unitId); assert.equal(snapshot[0].subUnitId, scope.subUnitId)
  assert.equal(cloned.groups[0].id, 'clone')
})

test('resource serialization round-trips summary settings and rebinds imported owners', () => {
  const before = { ...state(group('rows', 'row', 2, 4, true), group('columns', 'column', 1, 3)), summaryAbove: true, summaryLeft: true }
  const resource = serializeSheetOutlineResource(before, { ...scope, unitId: 'old-book', subUnitId: 'old-sheet' })
  const restored = deserializeSheetOutlineResource(resource, scope)
  assert.deepEqual(restored, before)
  assert.deepEqual(getSheetDimensionOutlines(restored, scope).map(item => [item.unitId, item.subUnitId]), [[scope.unitId, scope.subUnitId], [scope.unitId, scope.subUnitId]])
})

test('resource import preserves legacy missing owners and defensively detaches unknown fields', () => {
  const resource = { outlineList: [{ ...group('legacy', 'row', 2, 4), extension: { labels: ['kept'] } }] }
  const restored = deserializeSheetOutlineResource(resource, scope)
  resource.outlineList[0].extension.labels.push('input mutation')
  assert.deepEqual(restored.groups[0].extension.labels, ['kept'])
  rejects(() => deserializeSheetOutlineResource({ outlineList: [{ ...group('bad-owner', 'row', 1, 2), unitId: 3 }] }, scope), Reason.INVALID_RANGE)
})

test('an explicit empty workbook plugin clears all current sheets in one detached result', () => {
  const workbook = createBlankWorkbook(), first = workbook.sheets[0]
  first.outline = state(group('old', 'row', 1, 3)); workbook.sheets.push({ ...copy(first), id: 'second-sheet', name: 'Second' })
  const before = copy(workbook), result = deserializeWorkbookOutlineResource(workbook, '{}')
  assert.deepEqual(Object.keys(result), workbook.sheets.map(sheet => sheet.id))
  assert.deepEqual(Object.values(result), [createSheetOutlineState(), createSheetOutlineState()])
  assert.deepEqual(workbook, before)
})

test('malformed members or unknown sheet IDs reject the whole workbook resource', () => {
  const workbook = createBlankWorkbook(), id = workbook.sheets[0].id, before = copy(workbook)
  rejects(() => deserializeWorkbookOutlineResource(workbook, { unknown: { outlineList: [] } }), Reason.OUT_OF_BOUNDS)
  rejects(() => deserializeWorkbookOutlineResource(workbook, { [id]: { outlineList: [group('good', 'row', 1, 2), { ...group('bad', 'row', 3, 4), collapsed: 1 }] } }), Reason.INVALID_RANGE)
  rejects(() => deserializeWorkbookOutlineResource(workbook, '{invalid'), Reason.INVALID_RANGE)
  assert.deepEqual(workbook, before)
})

test('prototype-like sheet IDs remain ordinary own entries in both resource directions', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.id = '__proto__'; workbook.activeSheetId = sheet.id; sheet.outline = state(group('rows', 'row', 1, 2))
  const resource = serializeWorkbookOutlineResource(workbook)
  assert.equal(Object.hasOwn(resource, '__proto__'), true)
  assert.deepEqual(deserializeWorkbookOutlineResource(workbook, JSON.stringify(resource))['__proto__'], sheet.outline)
  assert.equal(Object.getPrototypeOf(resource), Object.prototype)
})

test('resource import rejects invalid summary schemas and excessive equal-range nesting', () => {
  rejects(() => deserializeSheetOutlineResource({ outlineList: [], officeStudio: { schemaVersion: 2, summaryAbove: true, summaryLeft: false } }, scope), Reason.INVALID_RANGE)
  rejects(() => deserializeSheetOutlineResource({ outlineList: Array.from({ length: 9 }, (_, index) => group(`layer-${index}`, 'row', 1, 5)) }, scope), Reason.MAX_DEPTH)
})
