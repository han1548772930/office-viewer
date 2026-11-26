import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet } from '../src/domain/model.ts'
import { getSheetOutlineHiddenSpans } from '../src/features/sheet-outline/sheet-outline.ts'
import { applySheetOutlineImport, normalizeSheetOutlineImport, reconcileSheetOutlineImportedVisibility } from '../src/features/sheet-outline/sheet-outline-import.ts'

// Deferred DTO/domain contracts. No test runner was started in this batch.
const json = value => JSON.parse(JSON.stringify(value))
const group = (id, axis, start, end, collapsed = false) => ({ id, axis, start, end, collapsed })
const run = (start, end, level, collapsed = false, hidden = null) => ({ start, end, level, collapsed, hidden, attributes: {}, namespaces: {} })
const extent = { rowCount: 30, columnCount: 12 }
function raw(options = {}) {
  return { outline: { summaryAbove: false, summaryLeft: true,
    groups: [group('rows', 'row', 1, 8, true), group('columns', 'column', 2, 4, true)] },
    sourceOutline: { version: 1, manualVisibilityNormalized: options.normalized ?? true,
      native: { summaryAbove: false, summaryLeft: true, zeroHeight: false, rowRuns: [run(1, 8, 1, false, true), run(9, 9, 0, true)],
        columnRuns: [run(1, 1, 0, true), run(2, 4, 1, false, true)], diagnostics: ['NATIVE_HIDDEN_REASONS_AMBIGUOUS: fixture'] },
      projectOutlineExtension: { version: 1, hiddenReasonsExact: options.exact ?? false } } }
}
function sheet() {
  return { ...createBlankSheet(), ...extent, hiddenRows: [4, 15], hiddenColumns: [3, 10],
    rowHiddenSpans: [{ start: 1, end: 19, value: true }], columnHiddenSpans: [{ start: 2, end: 8, value: true }] }
}

test('canonical IDs directions and unknown state/group/source data are detached on import', () => {
  const input = raw({ exact: true })
  input.outline.futureState = { list: [1, 2] }; input.outline.groups[0].futureGroup = { text: 'A & B' }
  input.sourceOutline.futureSource = { flags: [true] }
  const normalized = normalizeSheetOutlineImport(input, extent)
  assert.deepEqual(normalized.outline, input.outline); assert.equal(normalized.hiddenReasonsExact, true)
  normalized.outline.futureState.list.push(3); normalized.outline.groups[0].futureGroup.text = 'changed'; normalized.officeSource.futureSource.flags.push(false)
  assert.deepEqual(input.outline.futureState.list, [1, 2]); assert.equal(input.outline.groups[0].futureGroup.text, 'A & B')
  assert.deepEqual(input.sourceOutline.futureSource.flags, [true])
})

test('native summary dimensions beyond used cells extend the editor while boundary groups remain', () => {
  const input = raw(); input.outline.groups = [group('boundary', 'row', 0, 39, true)]
  input.sourceOutline.native.rowRuns = [run(0, 39, 1, false, true), run(40, 40, 0, true)]
  const result = normalizeSheetOutlineImport(input, { rowCount: 5, columnCount: 5 })
  assert.equal(result.rowCount, 41); assert.equal(result.outline.groups[0].end, 39)
  input.sourceOutline.native.rowRuns.pop()
  assert.equal(normalizeSheetOutlineImport(input, { rowCount: 5, columnCount: 5 }).rowCount, 40)
})

test('all eight native or project layers survive and a ninth layer rejects the candidate', () => {
  const input = raw(); input.outline.groups = Array.from({ length: 8 }, (_, index) => group(`layer-${index}`, 'row', 1, 8))
  input.sourceOutline.native.rowRuns = [run(1, 8, 8)]
  assert.equal(normalizeSheetOutlineImport(input, extent).outline.groups.length, 8)
  input.outline.groups.push(group('ninth', 'row', 1, 8))
  assert.throws(() => normalizeSheetOutlineImport(input, extent), /8|eight|深度|depth/i)
})

test('crossing reversed duplicate-ID and Office-limit ranges reject instead of disappearing', () => {
  const candidates = [[group('a', 'row', 1, 5), group('b', 'row', 3, 8)], [group('a', 'row', 5, 2)],
    [group('a', 'row', 1, 3), group('a', 'column', 1, 3)], [group('a', 'row', 0, 1_048_576)], [group('a', 'column', 0, 16_384)]]
  for (const groups of candidates) { const input = raw(); input.outline.groups = groups; const before = json(input)
    assert.throws(() => normalizeSheetOutlineImport(input, extent)); assert.deepEqual(input, before) }
})

test('malformed axis collapse summary and source types reject the entire DTO', () => {
  for (const edit of [value => { value.outline.groups[0].axis = 'rows' }, value => { value.outline.groups[0].collapsed = 'false' },
    value => { value.outline.summaryAbove = 1 }, value => { value.sourceOutline = [] },
    value => { value.sourceOutline.native.rowRuns[0].level = 9 }]) {
    const input = raw(); edit(input); assert.throws(() => normalizeSheetOutlineImport(input, extent))
  }
})

test('legacy missing IDs are deterministic and parser/public owner fields do not escape into the domain', () => {
  const input = raw(); delete input.outline.groups[0].id
  input.outline.groups[0].unitId = 'old-owner'; input.outline.groups[0].subUnitId = 'old-sheet'
  const first = normalizeSheetOutlineImport(input, extent), second = normalizeSheetOutlineImport(input, extent)
  assert.equal(first.outline.groups[0].id, 'xlsx-outline-legacy-0'); assert.deepEqual(first.outline, second.outline)
  assert.equal(Object.hasOwn(first.outline.groups[0], 'unitId'), false); assert.equal(Object.hasOwn(first.outline.groups[0], 'subUnitId'), false)
})

test('project-exact manual overlap remains hidden after an outline is expanded', () => {
  const target = sheet(), input = raw({ exact: true }), before = json(target)
  const result = applySheetOutlineImport(target, input)
  reconcileSheetOutlineImportedVisibility(target, result, [4, 15])
  assert.deepEqual(target.hiddenRows, before.hiddenRows); assert.deepEqual(target.rowHiddenSpans, before.rowHiddenSpans)
  assert.deepEqual(target.hiddenColumns, before.hiddenColumns); assert.deepEqual(target.columnHiddenSpans, before.columnHiddenSpans)
  target.outline.groups[0].collapsed = false
  assert.deepEqual(getSheetOutlineHiddenSpans(target.outline, 'row', target.rowCount), [])
  assert.ok(target.hiddenRows.includes(4))
})

test('legacy dimension-exact manual state also survives an independently stale outline snapshot', () => {
  const input = raw(), target = sheet()
  input.sourceAxisDimensions = { projectAxisExtension: { version: 1, restored: true } }
  input.sourceOutline.projectOutlineExtension.restored = false
  delete input.sourceOutline.projectOutlineExtension.hiddenReasonsExact
  const before = json(target), result = applySheetOutlineImport(target, input)
  assert.equal(result.hiddenReasonsExact, true); assert.deepEqual(target.hiddenRows, before.hiddenRows)
  reconcileSheetOutlineImportedVisibility(target, result, [4]); assert.deepEqual(target.rowHiddenSpans, before.rowHiddenSpans)
})

test('native inferred collapsed spans remove only their ambiguous manual cause', () => {
  const target = sheet(), result = applySheetOutlineImport(target, raw({ normalized: false }))
  assert.deepEqual(target.hiddenRows, [15]); assert.deepEqual(target.hiddenColumns, [10])
  assert.deepEqual(target.rowHiddenSpans, [{ start: 9, end: 19, value: true }])
  assert.deepEqual(target.columnHiddenSpans, [{ start: 5, end: 8, value: true }])
  assert.deepEqual(getSheetOutlineHiddenSpans(target.outline, 'row', result.rowCount), [{ start: 1, end: 8, value: true }])
})

test('Rust-normalized manual state is not subtracted from a second time', () => {
  const target = sheet(), before = json(target), result = applySheetOutlineImport(target, raw({ normalized: true }))
  assert.equal(result.manualVisibilityNormalized, true)
  assert.deepEqual(target.hiddenRows, before.hiddenRows); assert.deepEqual(target.rowHiddenSpans, before.rowHiddenSpans)
})

test('external filter inference removes failed rows from arrays and splits compressed manual spans', () => {
  const target = sheet(), result = applySheetOutlineImport(target, raw({ normalized: true }))
  reconcileSheetOutlineImportedVisibility(target, result, [4, 15, 17, 18])
  assert.deepEqual(target.hiddenRows, [])
  assert.deepEqual(target.rowHiddenSpans, [{ start: 1, end: 3, value: true }, { start: 5, end: 14, value: true },
    { start: 16, end: 16, value: true }, { start: 19, end: 19, value: true }])
  assert.deepEqual(target.columnHiddenSpans, [{ start: 2, end: 8, value: true }])
})

test('canonical models without native provenance retain authored manual hidden overlaps', () => {
  const target = sheet(), input = { outline: raw().outline }
  const before = json(target), result = applySheetOutlineImport(target, input)
  assert.equal(result.nativeSource, false); assert.equal(result.hiddenReasonsExact, true)
  assert.deepEqual(target.hiddenRows, before.hiddenRows); assert.deepEqual(target.rowHiddenSpans, before.rowHiddenSpans)
})

test('office source metadata has its own detached storage and diagnostic facts remain visible', () => {
  const target = sheet(), input = raw(), result = applySheetOutlineImport(target, input)
  assert.deepEqual(target.metadata.outlineOfficeSource, input.sourceOutline)
  assert.equal(result.diagnostics[0].startsWith('NATIVE_HIDDEN_REASONS_AMBIGUOUS'), true)
  target.metadata.outlineOfficeSource.native.rowRuns[0].end = 2
  assert.equal(input.sourceOutline.native.rowRuns[0].end, 8)
})

test('source-only non-outline files preserve their metadata without manufacturing groups', () => {
  const target = sheet(), input = raw(); delete input.outline
  input.sourceOutline.native.rowRuns = [run(15, 15, 0, false, true)]; input.sourceOutline.native.columnRuns = []
  const result = applySheetOutlineImport(target, input)
  assert.equal(result.outline, undefined); assert.equal(target.outline, undefined)
  assert.equal(target.metadata.outlineOfficeSource.version, 1)
})
