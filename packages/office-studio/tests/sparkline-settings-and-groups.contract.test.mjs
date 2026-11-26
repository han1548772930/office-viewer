import assert from 'node:assert/strict'
import { test } from 'vitest'
import { assertSparklineGroupConsistency, createSparklinePlacementPatches, createSparklineSpec, mergeSparklineSettings,
  planSparklineMutation, readSparklineGroups, resolveSparklineSettings, SPARKLINE_POINT_ORDER, SPARKLINE_THEMES } from '../src/features/sparkline/sparkline-model.ts'
import { parseSparklineRangeList, parseSparklineReference, shiftSparklineReference, sparklineCellAddress } from '../src/features/sparkline/sparkline-reference.ts'
import { installSparklinePatches, sparklineCell, sparklineJson, sparklineWorkbook } from './fixtures/sparkline-workbench-fixture.mjs'

// Deferred meaningful contracts. No test, type check, build, browser, Rust/WASM or benchmark was run for this batch.
const host = () => sparklineWorkbook().sheets[0]
const plan = (sheet, request, options = {}) => planSparklineMutation(sheet, request, { createGroupId: (() => { let next = 0; return () => `new-${++next}` })(), ...options })

test('quoted sheet references preserve escaped apostrophes, absolute axes and coordinate order', () => {
  const reference = parseSparklineReference("='Data, O''Neil'!$b2:D$4")
  assert.equal(reference.sheetName, "Data, O'Neil")
  assert.equal(reference.normalized, "'Data, O''Neil'!$B2:D$4")
  assert.deepEqual(reference.range, { startRow: 1, endRow: 3, startCol: 1, endCol: 3 })
  assert.equal(parseSparklineReference('D4:B2').normalized, 'D4:B2')
  assert.equal(sparklineCellAddress(1_048_575, 16_383), 'XFD1048576')
})

for (const value of ['A0', 'A01', 'XFE1', 'A1048577', 'A1:B2:C3', "'Open!A1", "'A'B'C'!A1", '[book.xlsx]Sheet1!A1', 'A!B!C1', 'A:A', '1:2']) {
  test(`invalid or external sparkline reference is rejected: ${value}`, () => assert.equal(parseSparklineReference(value), undefined))
}

test('range lists split real areas without splitting quoted separators and reject empty/unclosed input', () => {
  assert.deepEqual(parseSparklineRangeList("'Data, O''Neil'!$B$1:$E$1; G1\nH2, I3\r\n"), ["'Data, O''Neil'!$B$1:$E$1", 'G1', 'H2', 'I3'])
  assert.throws(() => parseSparklineRangeList("'Data!A1;B1"), /引号/)
  assert.throws(() => parseSparklineRangeList(' ; ,\n'), /填写/)
})

test('moving relative references retains absolute components and refuses either boundary overflow', () => {
  assert.equal(shiftSparklineReference("'Data, O''Neil'!$B2:D$4", 2, 1), "'Data, O''Neil'!$B4:E$4")
  assert.equal(shiftSparklineReference('$XFD$1048576', 10, 10), '$XFD$1048576')
  assert.throws(() => shiftSparklineReference('A1', -1, 0), /边界/)
  assert.throws(() => shiftSparklineReference('XFD1', 0, 1), /边界/)
  assert.throws(() => shiftSparklineReference('A1', .5, 0), /移动距离/)
})

for (const [source, target, expected] of [
  ['B1:D2', 'G1:G2', ['B1:D1', 'B2:D2']],
  ['B1:C3', 'G1:H1', ['B1:B3', 'C1:C3']],
  ['B1:C2', 'G1:H1', ['B1:C1', 'B2:C2']],
  ['$B$1:$D$2', 'G1:H1', ['$B$1:$D$1', '$B$2:$D$2']],
]) test(`placement matches source dimensions with rows taking square priority: ${source} -> ${target}`, () => {
  const patches = createSparklinePlacementPatches({ sourceRanges: [source], targetRanges: [target], groupId: 'placed' })
  assert.deepEqual(Object.values(patches).map(spec => spec.sourceRange), expected)
  assert.ok(Object.values(patches).every(spec => spec.groupId === 'placed'))
})

test('multiple vector sources pair in area order to disjoint single-cell targets', () => {
  const patches = createSparklinePlacementPatches({ sourceRanges: ['B1:E1', "'Data, O''Neil'!$B$2:$E$2"], targetRanges: ['I4', 'G2'], groupId: 'placed' }, host())
  assert.deepEqual(Object.keys(patches), ['I4', 'G2'])
  assert.equal(patches.G2.sourceRange, "'Data, O''Neil'!$B$2:$E$2")
  assert.equal(createSparklinePlacementPatches({ sourceRanges: ['B1:E1'], targetRanges: ['迷你图!$G$1'], groupId: 'placed' }, host()).G1.sourceRange, 'B1:E1')
})

for (const [sources, targets] of [
  [['B1:D2'], ['G1:H2']], [['B1:E1', 'B2:E2'], ['G1:G2', 'I2']], [['B1:E1', 'B2:E2'], ['G1']],
  [['B1:D2', 'B3:E3'], ['G1', 'G2']], [['B1:E1', 'B2:E2'], ['G1', '$G$1']], [['B1:E1'], ['Other!G1']],
]) test(`invalid placement leaves its host untouched: ${sources.join(';')} -> ${targets.join(';')}`, () => {
  const sheet = host(), before = sparklineJson(sheet)
  assert.throws(() => createSparklinePlacementPatches({ sourceRanges: sources, targetRanges: targets, groupId: 'placed' }, sheet))
  assert.deepEqual(sheet, before)
})

for (const [name, colors] of Object.entries(SPARKLINE_THEMES)) test(`theme ${name} applies series and all six marker colors`, () => {
  const settings = mergeSparklineSettings(resolveSparklineSettings({ points: { highPoint: { visible: true } } }), { themeType: name })
  assert.equal(settings.color, colors[0]); assert.equal(settings.points.highPoint.visible, true)
  assert.ok(SPARKLINE_POINT_ORDER.every(point => settings.points[point].color === colors[1]))
  const customized = mergeSparklineSettings(settings, { points: { firstPoint: { color: '#123456' } } })
  assert.equal(customized.themeType, 'custom'); assert.equal(customized.points.highPoint.color, colors[1])
})

test('partial extremum value edits retain modes, other endpoints and inactive custom values', () => {
  const current = resolveSparklineSettings({ extremumMin: { type: 'group', value: -30 }, extremumMax: { type: 'custom', value: 80 } })
  const next = mergeSparklineSettings(current, { extremumMin: { value: -40 }, axisOptions: { reverse: true }, points: { firstPoint: { visible: true } } })
  assert.deepEqual(next.extremumMin, { type: 'group', value: -40 }); assert.deepEqual(next.extremumMax, current.extremumMax)
  const inactive = mergeSparklineSettings(next, { extremumMax: { type: 'individual' } })
  assert.equal(inactive.extremumMax.value, 80)
  assert.equal(mergeSparklineSettings(inactive, { extremumMax: { type: 'custom' } }).extremumMax.value, 80)
})

for (const patch of [{ points: null }, { points: { highPoint: [] } }, { axisOptions: [] }, { extremumMin: null }, { type: 'area' },
  { containHiddenCells: 1 }, { points: { firstPoint: { visible: 'true' } } }, { nonNumericCells: 'hide' }, { color: 'invalid' },
  { lineWidth: -1 }, { extremumMax: { type: 'custom', value: Infinity } }, { extremumMin: { type: 'custom', value: 5 }, extremumMax: { type: 'custom', value: 5 } }]) {
  test(`malformed settings are rejected before merging: ${JSON.stringify(patch)}`, () => {
    assert.throws(() => resolveSparklineSettings(patch)); assert.throws(() => mergeSparklineSettings(resolveSparklineSettings(), patch))
  })
}

test('legacy Office axes and colors normalize into independently serializable advanced settings', () => {
  const spec = createSparklineSpec({ targetRange: 'G1', sourceRange: '$B$1:$E$1', type: 'column', color: '#123456', negativeColor: '#ee1122',
    axis: 'custom', min: -20, max: 40, extremumMin: { value: -10 }, axisOptions: { visible: true } })
  assert.deepEqual(spec.extremumMin, { type: 'custom', value: -10 }); assert.equal(spec.points.negativePoint.visible, true)
  assert.equal(spec.min, -10); assert.equal(spec.max, 40); assert.equal(spec.sourceRange, '$B$1:$E$1')
})

test('reading one selected member returns full membership without a second persistent group registry', () => {
  const sheet = host(), groups = readSparklineGroups(sheet, ['G2'])
  assert.equal(groups.length, 1); assert.equal(groups[0].groupId, 'north')
  assert.deepEqual(groups[0].selectedAddresses, ['G2']); assert.deepEqual(groups[0].targetRanges, ['G1', 'G2'])
  assert.equal(groups[0].consistent, true); assert.equal(sheet.sparklineGroups, undefined)
})

test('setting one member updates its full group and preserves other groups and source ranges', () => {
  const sheet = host(), before = sparklineJson(sheet), result = plan(sheet, { kind: 'settings', ranges: ['G2'], settings: { lineWidth: 4, points: { firstPoint: { visible: true } } } })
  assert.deepEqual(Object.keys(result.patches), ['G1', 'G2']); assert.equal(result.changed, true)
  assert.deepEqual(sheet, before)
  installSparklinePatches(sheet, result.patches)
  for (const address of ['G1', 'G2']) assert.equal(sparklineCell(sheet, address).sparkline.lineWidth, 4)
  assert.deepEqual(sparklineCell(sheet, 'G4').sparkline, sparklineCell(before, 'G4').sparkline)
  assert.equal(sparklineCell(sheet, 'G2').sparkline.sourceRange, 'B2:E2'); assertSparklineGroupConsistency(sheet)
})

test('combine extracts selected members from mixed groups and leaves each unselected member in its original group', () => {
  const sheet = host(), before = sparklineJson(sheet), result = plan(sheet, { kind: 'combine', ranges: ['G2', 'G4'], groupId: 'combined' })
  assert.deepEqual(Object.keys(result.patches), ['G2', 'G4']); installSparklinePatches(sheet, result.patches)
  assert.equal(sparklineCell(sheet, 'G2').sparkline.groupId, 'combined'); assert.equal(sparklineCell(sheet, 'G4').sparkline.type, 'line')
  assert.equal(sparklineCell(sheet, 'G1').sparkline.groupId, 'north'); assert.equal(sparklineCell(sheet, 'G5').sparkline.groupId, 'south')
  assert.deepEqual(sparklineCell(sheet, 'G1').sparkline, sparklineCell(before, 'G1').sparkline); assertSparklineGroupConsistency(sheet)
})

test('uncombine assigns unique groups to selected members, retains sources and makes singleton repetitions no-ops', () => {
  const sheet = host(), result = plan(sheet, { kind: 'uncombine', ranges: ['G1:G2'] })
  assert.notEqual(result.patches.G1.groupId, result.patches.G2.groupId)
  assert.equal(result.patches.G1.sourceRange, 'B1:E1'); installSparklinePatches(sheet, result.patches)
  assert.equal(plan(sheet, { kind: 'uncombine', ranges: ['G1:G2'] }).changed, false)
  assert.equal(readSparklineGroups(sheet).length, 3)
})

for (const scope of ['item', 'group']) test(`clear ${scope} produces the exact footprint and preserves source cell content`, () => {
  const sheet = host(), result = plan(sheet, { kind: 'clear', scope, ranges: ['G2'] })
  assert.deepEqual(Object.keys(result.patches), scope === 'item' ? ['G2'] : ['G1', 'G2'])
  assert.ok(Object.values(result.patches).every(spec => spec === null)); installSparklinePatches(sheet, result.patches)
  assert.equal(sparklineCell(sheet, 'B1').value, 2); assert.equal(sparklineCell(sheet, 'G4').sparkline.groupId, 'south')
})

test('item source and location plus full-group settings form one read-only atomic plan', () => {
  const workbook = sparklineWorkbook(), sheet = workbook.sheets[0], before = sparklineJson(sheet)
  const result = plan(sheet, { kind: 'change-source', ranges: ['G2'], primary: 'G2', resetType: 'item', sourceRanges: ["'Data, O''Neil'!B2:E2"],
    targetRanges: ['I3'], settings: { lineWidth: 3, extremumMax: { type: 'group' } } }, { sheets: workbook.sheets })
  assert.deepEqual(sheet, before); assert.deepEqual(Object.keys(result.patches).sort(), ['G1', 'G2', 'I3'])
  assert.equal(result.patches.G2, null); assert.equal(result.patches.I3.groupId, 'north'); assert.equal(result.patches.G1.lineWidth, 3)
  assert.equal(result.patches.I3.sourceRange, "'Data, O''Neil'!B2:E2")
  installSparklinePatches(sheet, result.patches); assertSparklineGroupConsistency(sheet)
})

test('group source replacement can change member count and clears every displaced location', () => {
  const sheet = host(), result = plan(sheet, { kind: 'change-source', resetType: 'group', ranges: ['G2'], sourceRanges: ['B1:E3'], targetRanges: ['I1:I3'] })
  assert.deepEqual(Object.keys(result.patches).sort(), ['G1', 'G2', 'I1', 'I2', 'I3'])
  assert.equal(result.patches.G1, null); assert.equal(result.patches.I3.sourceRange, 'B3:E3')
  assert.ok(['I1', 'I2', 'I3'].every(address => result.patches[address].groupId === 'north'))
  installSparklinePatches(sheet, result.patches); assert.equal(readSparklineGroups(sheet, ['I1'])[0].members.length, 3)
})

for (const request of [
  { kind: 'change-source', resetType: 'item', ranges: ['G1'], sourceRanges: ['B1:E1'], targetRanges: ['G4'] },
  { kind: 'change-source', resetType: 'item', ranges: ['G1'], sourceRanges: ['B1:E1', 'B2:E2'], targetRanges: ['I1'] },
  { kind: 'change-source', resetType: 'group', ranges: ['G1', 'G4'], sourceRanges: ['B1:E2'], targetRanges: ['I1:I2'] },
  { kind: 'change-source', resetType: 'item', ranges: ['G1'], primary: 'G2', sourceRanges: ['B1:E1'], targetRanges: ['I1'] },
]) test(`invalid source transaction is all-or-nothing: ${JSON.stringify(request)}`, () => {
  const sheet = host(), before = sparklineJson(sheet); assert.throws(() => plan(sheet, request)); assert.deepEqual(sheet, before)
})

test('new sources require a live workbook sheet while deleting a member with a lost source is allowed', () => {
  const sheet = host(), before = sparklineJson(sheet)
  assert.throws(() => plan(sheet, { kind: 'insert', sourceRanges: ['Missing!B1:E1'], targetRanges: ['I1'] }), /不存在/)
  assert.deepEqual(sheet, before); sparklineCell(sheet, 'G1').sparkline.sourceRange = 'Missing!B1:E1'
  assert.deepEqual(plan(sheet, { kind: 'clear', ranges: ['G1'], scope: 'item' }).patches, { G1: null })
  assert.throws(() => plan(sheet, { kind: 'settings', ranges: ['G1'], settings: { lineWidth: 4 } }), /不存在/)
})

test('equivalent normalized settings produce no patches and inconsistent groups cannot publish unrelated updates', () => {
  const sheet = host(); assert.equal(plan(sheet, { kind: 'settings', ranges: ['G1'], settings: {} }).changed, false)
  sparklineCell(sheet, 'G2').sparkline.color = '#999999'
  assert.equal(readSparklineGroups(sheet, ['G1'])[0].consistent, false)
  assert.throws(() => plan(sheet, { kind: 'settings', ranges: ['G4'], settings: { showRadius: true } }), /不一致/)
  const repaired = plan(sheet, { kind: 'settings', ranges: ['G1'], settings: { color: '#246fce' } })
  installSparklinePatches(sheet, repaired.patches); assertSparklineGroupConsistency(sheet)
})

for (const groupId of ['', ' north', 'north', 'ungrouped:G1']) test(`new group IDs reject empty, reserved and colliding values: ${groupId}`, () => {
  assert.throws(() => plan(host(), { kind: 'insert', sourceRanges: ['B1:E1'], targetRanges: ['I1'], groupId }))
})
