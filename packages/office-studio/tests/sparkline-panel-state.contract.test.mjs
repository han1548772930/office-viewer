import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { parseSparklineReference } from '../src/features/sparkline/sparkline-reference.ts'
import { sparklineCell, sparklineJson, sparklineWorkbook } from './fixtures/sparkline-workbench-fixture.mjs'

// Deferred draft/owner/full-footprint contracts against the real Facade and content history.
// The Vite SSR loader below has not been started and none of these drafts has run.
let server, api, panel, model
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, panel, model] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/sparkline/sparkline-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/sparkline/sparkline-model.ts'),
  ])
})
afterAll(async () => { await server?.close() })
function setup(value = sparklineWorkbook()) {
  const studio = new api.OfficeStudio({ workbook: value, currentUserId: 'local-user' }), book = api.FOffice.newAPI(studio).getActiveWorkbook()
  const target = { workbookId: book.getId(), sheetId: 'data' }, sheet = book.getSheetBySheetId('data')
  return { studio, book, sheet, target, snapshot: (ranges = ['G1']) => panel.readSparklinePanelState(studio, target, ranges),
    state: () => sparklineJson(book.save()), history: () => studio.editor.history.units[book.getId()]?.undo.length ?? 0, dispose: () => studio.dispose() }
}
const rule = (address, viewState = 'othersCanView') => ({ id: 'peer-rule', permissionId: 'peer-rule', ownerUserId: 'peer',
  allowedUserIds: [], editState: 'onlyMe', viewState, ranges: [parseSparklineReference(address).range], description: 'Deferred permission fixture' })

test('panel reads detached full groups and creates item/group source drafts from authoritative cell specs', () => {
  const f = setup()
  try {
    const snapshot = f.snapshot(['G2']), item = panel.createSparklinePanelDraft(snapshot, undefined, 'G2', 'item')
    assert.equal(item.sourceText, 'B2:E2'); assert.equal(item.targetText, 'G2'); assert.equal(item.primary, 'G2')
    const group = panel.createSparklinePanelDraft(snapshot, 'north', 'G2', 'group')
    assert.equal(group.sourceText, 'B1:E1; B2:E2'); assert.equal(group.targetText, 'G1; G2')
    snapshot.groups[0].settings.color = '#000000'; assert.equal(f.sheet.getSparkline('G1').color, '#246fce')
    assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('new insertion drafts preserve disjoint targets and produce one multi-source insertion request', () => {
  const f = setup(sparklineWorkbook({ withGroups: false }))
  try {
    const snapshot = f.snapshot(['I1', 'I3']), original = panel.createSparklinePanelDraft(snapshot), draft = sparklineJson(original)
    draft.sourceText = "B1:E1; 'Data, O''Neil'!B2:E2"; draft.settings.type = 'pie'
    const request = panel.buildSparklinePanelRequest(draft, original, snapshot.ranges, false)
    assert.equal(request.kind, 'insert'); assert.deepEqual(request.targetRanges, ['I1', 'I3'])
    assert.deepEqual(request.sourceRanges, ['B1:E1', "'Data, O''Neil'!B2:E2"]); assert.equal(request.settings.type, 'pie')
    panel.assertSparklinePanelMutation(snapshot.owner, request); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('nested dirty patches include only changed axis/marker/extremum fields', () => {
  const original = model.resolveSparklineSettings({ extremumMin: { type: 'group', value: -20 }, extremumMax: { type: 'individual', value: 50 } })
  const current = panel.editSparklineSettingsDraft(original, { extremumMin: { value: -25 }, axisOptions: { reverse: true }, points: { firstPoint: { visible: true } } })
  assert.deepEqual(panel.buildSparklineSettingsPatch(current, original), {
    axisOptions: { reverse: true }, extremumMin: { value: -25 }, points: { firstPoint: { visible: true } },
  })
  assert.equal(current.extremumMin.type, 'group'); assert.deepEqual(original.extremumMin, { type: 'group', value: -20 })
})

test('editing one nested field across mixed groups retains each untouched mode, color and visibility', () => {
  const f = setup()
  try {
    const snapshot = f.snapshot(['G1', 'G4']), original = panel.createSparklinePanelDraft(snapshot), draft = sparklineJson(original)
    const south = f.sheet.getSparkline('G4')
    draft.settings = panel.editSparklineSettingsDraft(draft.settings, { axisOptions: { reverse: false }, extremumMin: { value: -30 }, points: { firstPoint: { visible: true } } })
    const request = panel.buildSparklinePanelRequest(draft, original, snapshot.ranges, true)
    assert.deepEqual(request.settings.extremumMin, { value: -30 }); assert.equal(request.settings.axisOptions, undefined)
    panel.assertSparklinePanelMutation(snapshot.owner, request)
    assert.equal(f.sheet.applySparklineMutation(request, panel.sparklinePanelOwnerContext(snapshot.owner)), true)
    const after = f.sheet.getSparkline('G4')
    assert.equal(after.extremumMin.type, south.extremumMin.type); assert.equal(after.extremumMin.value, -30)
    assert.equal(after.axisOptions.reverse, south.axisOptions.reverse); assert.equal(after.points.lowPoint.color, south.points.lowPoint.color)
    assert.equal(after.color, south.color); assert.equal(f.history(), 1)
  } finally { f.dispose() }
})

test('switching extrema away from custom and back retains inactive numeric values in the draft', () => {
  const original = model.resolveSparklineSettings({ extremumMin: { type: 'custom', value: -12 }, extremumMax: { type: 'custom', value: 90 } })
  const inactive = panel.editSparklineSettingsDraft(original, { extremumMin: { type: 'individual' }, extremumMax: { type: 'group' } })
  assert.equal(inactive.extremumMin.value, -12); assert.equal(inactive.extremumMax.value, 90)
  const restored = panel.editSparklineSettingsDraft(inactive, { extremumMin: { type: 'custom' }, extremumMax: { type: 'custom' } })
  assert.deepEqual(panel.buildSparklineSettingsPatch(restored, original), {})
})

test('an empty custom numeric draft remains editable but cannot be published until finite and ordered', () => {
  const original = model.resolveSparklineSettings({ extremumMin: { type: 'custom', value: -5 }, extremumMax: { type: 'custom', value: 5 } })
  const incomplete = panel.editSparklineSettingsDraft(original, { extremumMin: { value: undefined } })
  assert.equal(incomplete.extremumMin.value, undefined); assert.throws(() => panel.buildSparklineSettingsPatch(incomplete, original), /有限数值/)
  const invalid = panel.editSparklineSettingsDraft(incomplete, { extremumMin: { value: 6 } })
  assert.throws(() => panel.buildSparklineSettingsPatch(invalid, original), /递增/)
  const restored = panel.editSparklineSettingsDraft(invalid, { extremumMin: { value: -6 } })
  assert.deepEqual(panel.buildSparklineSettingsPatch(restored, original), { extremumMin: { value: -6 } })
})

test('theme presets recolor all six markers while direct edits select custom without replacing untouched values', () => {
  const original = model.resolveSparklineSettings({ points: { highPoint: { visible: true, color: '#cc2244' } } })
  const themed = panel.editSparklineSettingsDraft(original, { themeType: 'fresh' })
  assert.equal(themed.color, '#2ebb7a'); assert.equal(themed.points.highPoint.color, '#f8be15'); assert.equal(themed.points.highPoint.visible, true)
  const custom = panel.editSparklineSettingsDraft(themed, { points: { highPoint: { color: '#556677' } } })
  assert.equal(custom.themeType, 'custom'); assert.equal(custom.points.lowPoint.color, '#f8be15')
})

test('settings plus item data relocation are emitted as one request and one history transaction', () => {
  const f = setup(), before = f.state()
  try {
    const snapshot = f.snapshot(['G2']), original = panel.createSparklinePanelDraft(snapshot, undefined, 'G2', 'item'), draft = sparklineJson(original)
    draft.sourceText = "'Data, O''Neil'!B2:E2"; draft.targetText = 'I3'; draft.settings.lineWidth = 4
    const request = panel.buildSparklinePanelRequest(draft, original, snapshot.ranges, true)
    assert.equal(request.kind, 'change-source'); assert.equal(request.resetType, 'item'); assert.deepEqual(request.settings, { lineWidth: 4 })
    panel.assertSparklinePanelMutation(snapshot.owner, request)
    assert.equal(f.sheet.applySparklineMutation(request, panel.sparklinePanelOwnerContext(snapshot.owner)), true)
    assert.equal(f.history(), 1); assert.equal(f.sheet.getSparkline('G2'), null); assert.equal(f.sheet.getSparkline('G1').lineWidth, 4)
    assert.equal(f.sheet.getSparkline('I3').groupId, 'north'); assert.equal(panel.sparklinePanelOwnerCurrent(snapshot.owner, true), true)
    assert.equal(panel.sparklinePanelOwnerCurrent(snapshot.owner), false)
    f.book.undo(); assert.deepEqual(f.state(), before); f.book.redo(); assert.equal(f.sheet.getSparkline('I3').lineWidth, 4)
  } finally { f.dispose() }
})

test('equivalent source spelling returns false without committing history or replacing the current owner', () => {
  const f = setup()
  try {
    const snapshot = f.snapshot(), original = panel.createSparklinePanelDraft(snapshot, undefined, 'G1', 'item'), draft = sparklineJson(original)
    draft.sourceText = ' B1:E1 '
    assert.equal(panel.sparklinePanelDraftDirty(draft, original), true)
    const request = panel.buildSparklinePanelRequest(draft, original, snapshot.ranges, true)
    assert.equal(f.sheet.applySparklineMutation(request, panel.sparklinePanelOwnerContext(snapshot.owner)), false)
    assert.equal(f.history(), 0); assert.equal(panel.sparklinePanelOwnerCurrent(snapshot.owner), true)
  } finally { f.dispose() }
})

for (const kind of ['revision', 'history', 'user', 'host', 'active-kind', 'disposed']) test(`captured panel owners reject ${kind} changes before publication`, () => {
  const f = setup(), snapshot = f.snapshot()
  try {
    if (kind === 'revision') f.sheet.getRange('B1').setValue(999)
    else if (kind === 'history') f.studio.editor.history.clear(f.target.workbookId)
    else if (kind === 'user') Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
    else if (kind === 'host') f.book.setActiveSheet('other')
    else if (kind === 'active-kind') f.studio.setActiveKind('document')
    else f.studio.dispose()
    assert.equal(panel.sparklinePanelOwnerCurrent(snapshot.owner), false)
    assert.throws(() => panel.sparklinePanelOwnerContext(snapshot.owner), /更改/)
    assert.throws(() => panel.assertSparklinePanelMutation(snapshot.owner, { kind: 'settings', ranges: ['G1'], settings: { lineWidth: 4 } }))
  } finally { f.dispose() }
})

test('owner capture is tied to the active workbook host and cannot silently use an inactive sheet', () => {
  const f = setup()
  try {
    assert.throws(() => panel.captureSparklinePanelOwner(f.studio, { ...f.target, sheetId: 'other' }), /工作台/)
    assert.equal(panel.sparklinePanelOwnerCurrent(panel.captureSparklinePanelOwner(f.studio, f.target)), true)
  } finally { f.dispose() }
})

test('full-group edit permission blocks a selected writable member when an unselected member is protected', () => {
  const value = sparklineWorkbook(); value.sheets[0].rangeProtectionRules = [rule('G2')]
  const f = setup(value)
  try {
    const snapshot = f.snapshot(['G1']), group = snapshot.selectedGroups[0]
    assert.equal(panel.canEditSparklinePanelRanges(snapshot.owner, ['G1']), true)
    assert.equal(panel.canEditSparklinePanelRanges(snapshot.owner, group.targetRanges), false)
    assert.throws(() => panel.assertSparklinePanelMutation(snapshot.owner, { kind: 'settings', ranges: ['G1'], settings: { lineWidth: 4 } }), /受保护/)
    assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('an unreadable unselected group member prevents exposing a partial group in the panel snapshot', () => {
  const value = sparklineWorkbook(); value.sheets[0].rangeProtectionRules = [rule('G2', 'noOneElseCanView')]
  const f = setup(value)
  try { assert.throws(() => f.snapshot(['G1']), /权限/); assert.equal(f.history(), 0) } finally { f.dispose() }
})

for (const kind of ['source-sheet', 'source-range', 'new-target']) test(`the panel rejects a protected ${kind} footprint before its transaction`, () => {
  const value = sparklineWorkbook(), source = value.sheets[1]
  sparklineCell(value.sheets[0], 'G1').sparkline.sourceRange = "'Data, O''Neil'!B2:E2"
  if (kind === 'source-sheet') source.worksheetProtection = { id: 'source', permissionId: 'source', unitType: 'worksheet', ownerUserId: 'peer',
    editState: 'onlyMe', allowedUserIds: [], viewState: 'noOneElseCanView' }
  if (kind === 'source-range') source.rangeProtectionRules = [rule('B2:E2', 'noOneElseCanView')]
  if (kind === 'new-target') value.sheets[0].rangeProtectionRules = [rule('I3')]
  const f = setup(value)
  try {
    const snapshot = f.snapshot(), before = f.state()
    const request = kind === 'new-target' ? { kind: 'change-source', ranges: ['G1'], resetType: 'item', sourceRanges: ['B1:E1'], targetRanges: ['I3'] }
      : { kind: 'settings', ranges: ['G1'], settings: { lineWidth: 4 } }
    assert.throws(() => panel.assertSparklinePanelMutation(snapshot.owner, request), /权限|受保护/)
    assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('group source drafts preserve multiple ranges and settings in one replacement request', () => {
  const f = setup()
  try {
    const snapshot = f.snapshot(), original = panel.createSparklinePanelDraft(snapshot, 'north', 'G1', 'group'), draft = sparklineJson(original)
    draft.sourceText = 'B1:E1; B3:E3; B4:E4'; draft.targetText = 'I1; I3; I5'; draft.settings.showGradient = true
    const request = panel.buildSparklinePanelRequest(draft, original, snapshot.ranges, true)
    assert.equal(request.resetType, 'group'); assert.deepEqual(request.targetRanges, ['I1', 'I3', 'I5']); assert.deepEqual(request.settings, { showGradient: true })
    panel.assertSparklinePanelMutation(snapshot.owner, request); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})
