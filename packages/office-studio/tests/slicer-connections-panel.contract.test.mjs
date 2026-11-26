import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { connectionFixtureIds as ids, connectionReportSelection, slicerConnectionsWorkbook } from './fixtures/slicer-connections-fixture.mjs'

// Deferred model/Facade/panel contracts. No test or Vite hook below has been run.
let server, api, panel, connections, owners, drawings, ribbon, schema
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, panel, connections, owners, drawings, ribbon, schema] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-connections-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-panel-owner.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/sheet-drawings/sheet-drawing-host.ts'),
    server.ssrLoadModule('/packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts'),
    server.ssrLoadModule('/packages/office-studio/src/platform/ribbon/ribbon-schema.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const model = f => f.studio.getWorkbookModel(f.unitId)
const history = f => f.studio.editor.history.units[f.unitId]?.undo.length ?? 0
const target = f => ({ workbookId: f.unitId, sheetId: ids.host })
const code = expected => error => error?.code === expected
const option = (kind = 'pivotTable', subUnitId = ids.host, id = ids.primary) => ({ kind, subUnitId, id, name: id, sheetName: subUnitId, connected: false })
function setup(workbook = slicerConnectionsWorkbook(), currentUserId = 'local-user') {
  const studio = new api.OfficeStudio({ workbook, currentUserId }), facade = api.FOffice.newAPI(studio), book = facade.getActiveWorkbook()
  return { studio, facade, book, unitId: book.getId(), dispose: () => studio.dispose() }
}
function rootItems(tab) { return tab.groups.flatMap(group => group.items) }
function ribbonState(control) {
  return { canUndo: false, canRedo: false, hasSheet: true, canMerge: false, canUnmerge: false, canDeleteRow: true, canDeleteColumn: true,
    chartAvailable: true, dataValidationAvailable: true, canSort: true, fontFamily: 'Arial', fontSize: 11, numberFormat: 'General',
    gridlines: true, frozen: false, zoom: 100, fullscreen: false, slicerControl: control }
}

test('connection keys keep report kinds and separator-containing worksheet/report IDs distinct', () => {
  const values = [option('pivotTable', 'a:b', 'c'), option('pivotTable', 'a', 'b:c'), option('ownedPivotChart', 'a:b', 'c')]
  assert.equal(new Set(values.map(connections.slicerConnectionKey)).size, 3)
  assert.deepEqual(values.map(value => JSON.parse(connections.slicerConnectionKey(value))), values.map(({ kind, subUnitId, id }) => [kind, subUnitId, id]))
  assert.deepEqual(connections.slicerConnectionKeys({ pivotTables: [{ subUnitId: 'a:b', pivotTableId: 'c' }],
    ownedPivotCharts: [{ subUnitId: 'a:b', pivotChartId: 'c' }] }), [connections.slicerConnectionKey(values[0]), connections.slicerConnectionKey(values[2])])
})

test('selection replaces both collections, deduplicates keys and preserves different report kinds with the same ID', () => {
  const pivot = option(), owned = option('ownedPivotChart'), options = [pivot, owned]
  const selected = connections.selectedSlicerConnections(options, [connections.slicerConnectionKey(owned), connections.slicerConnectionKey(pivot), connections.slicerConnectionKey(owned)])
  assert.deepEqual(selected, { pivotTables: [{ subUnitId: ids.host, pivotTableId: ids.primary }], ownedPivotCharts: [{ subUnitId: ids.host, pivotChartId: ids.primary }] })
  assert.deepEqual(options, [pivot, owned])
})

test('empty checked selection produces a real empty replacement', () => {
  assert.deepEqual(connections.selectedSlicerConnections([option()], []), { pivotTables: [], ownedPivotCharts: [] })
})

test('unknown selected keys cannot be silently discarded while replacing connections', () => {
  assert.throws(() => connections.selectedSlicerConnections([option()], [connections.slicerConnectionKey(option('pivotTable', ids.host, 'missing'))]), code('STALE_CONNECTION_SELECTION'))
})

test('nonstring selected values are rejected before constructing command parameters', () => {
  assert.throws(() => connections.selectedSlicerConnections([option()], [12]), code('INVALID_CONNECTION_SELECTION'))
})

for (const [name, invalid] of [
  ['unknown kind', option('chart')], ['missing worksheet', option('pivotTable', '')], ['missing report', option('pivotTable', ids.host, '')],
]) test(`a candidate with ${name} is rejected instead of becoming an empty successful replacement`, () => {
  assert.throws(() => connections.selectedSlicerConnections([invalid], []), code('INVALID_CONNECTION_TARGET'))
})

test('insertion source keys distinguish Table, PivotTable and independent PivotChart IDs', () => {
  const keys = ['table', 'pivot', 'ownedPivotChart'].map(kind => panel.slicerPanelSourceKey({ kind, id: 'same:report' }))
  assert.equal(new Set(keys).size, 3)
})

test('reading connections returns compatible reports, the owned chart and checked initial state without creating history', async () => {
  const f = setup()
  try {
    const before = json(model(f)), baseline = history(f), snapshot = await connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer)
    assert.deepEqual(snapshot.options.map(({ kind, subUnitId, id, connected }) => ({ kind, subUnitId, id, connected })), [
      { kind: 'pivotTable', subUnitId: ids.host, id: ids.primary, connected: true },
      { kind: 'pivotTable', subUnitId: ids.other, id: ids.secondary, connected: false },
      { kind: 'ownedPivotChart', subUnitId: ids.host, id: ids.owned, connected: false },
    ])
    assert.equal(snapshot.cacheId, f.book.getSlicer(ids.shared).getCacheId()); assert.ok(connections.slicerConnectionsPanelCurrent(snapshot))
    snapshot.connections.pivotTables[0].pivotTableId = 'draft-only'
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('Table slicers refuse report connection panels while retaining their local filter source', async () => {
  const f = setup()
  try {
    const before = json(model(f))
    await assert.rejects(connections.readPanelSlicerConnections(f.studio, target(f), ids.table), code('UNSUPPORTED_CONNECTIONS'))
    assert.equal(f.book.getSlicer(ids.table).getConnections(), undefined); assert.deepEqual(json(model(f)), before); assert.equal(history(f), 0)
  } finally { f.dispose() }
})

test('an inaccessible already connected report rejects the read instead of unchecking it', async () => {
  const f = setup(), original = api.FSheetSlicer.prototype.getConnectionCandidates
  try {
    const before = json(model(f))
    api.FSheetSlicer.prototype.getConnectionCandidates = async function() { return (await original.call(this)).filter(value => value.id !== ids.primary) }
    await assert.rejects(connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer), code('UNAVAILABLE_CONNECTED_REPORT'))
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), 0)
  } finally { api.FSheetSlicer.prototype.getConnectionCandidates = original; f.dispose() }
})

test('saving the panel applies current selection to every target and every view of its shared cache in one history entry', async () => {
  const f = setup(); let reopened
  try {
    await f.book.getSlicer(ids.slicer).setItems(['North'])
    const before = json(f.book.save()), baseline = history(f), snapshot = await connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer)
    const keys = snapshot.options.map(connections.slicerConnectionKey)
    assert.equal(await connections.applyPanelSlicerConnections(snapshot, keys), true); assert.equal(history(f), baseline + 1)
    const after = json(f.book.save()), first = f.book.getSlicer(ids.slicer).getConnections()
    assert.deepEqual(first, f.book.getSlicer(ids.shared).getConnections()); assert.equal(first.pivotTables.length, 2); assert.equal(first.ownedPivotCharts.length, 1)
    for (const reportId of [ids.primary, ids.secondary, ids.owned]) assert.deepEqual(connectionReportSelection(model(f), reportId), ['North'])
    assert.ok(connections.slicerConnectionsPanelCurrent(snapshot, true)); assert.equal(connections.slicerConnectionsPanelCurrent(snapshot), false)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), after)
    reopened = new api.OfficeStudio({ workbook: after }); const book = api.FOffice.newAPI(reopened).getActiveWorkbook()
    assert.deepEqual(book.getSlicer(ids.slicer).getConnections(), first); assert.deepEqual(book.getSlicer(ids.shared).getModel().selectedItems, ['North'])
  } finally { reopened?.dispose(); f.dispose() }
})

test('disconnecting every report retains source and chosen values, and reattachment applies the retained selection', async () => {
  const f = setup()
  try {
    await f.book.getSlicer(ids.slicer).setItems(['North'])
    const originalSource = json(f.book.getSlicer(ids.slicer).getModel()), snapshot = await connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer)
    const baseline = history(f)
    assert.equal(await connections.applyPanelSlicerConnections(snapshot, []), true)
    assert.deepEqual(f.book.getSlicer(ids.shared).getConnections(), { pivotTables: [], ownedPivotCharts: [] })
    assert.deepEqual(f.book.getSlicer(ids.shared).getModel().selectedItems, ['North']); assert.deepEqual(f.book.getSlicer(ids.slicer).getItemLabels(), ['North', 'South'])
    const current = f.book.getSlicer(ids.slicer).getModel()
    assert.equal(current.cacheId, originalSource.cacheId); assert.equal(current.sourceFieldId, originalSource.sourceFieldId)
    const next = await connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer)
    assert.equal(await connections.applyPanelSlicerConnections(next, [connections.slicerConnectionKey(next.options.find(option => option.id === ids.secondary))]), true)
    assert.deepEqual(connectionReportSelection(model(f), ids.secondary), ['North']); assert.equal(history(f), baseline + 2)
  } finally { f.dispose() }
})

test('saving unchanged selected reports is a no-op in the model and history', async () => {
  const f = setup()
  try {
    const before = json(model(f)), baseline = history(f), snapshot = await connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer)
    assert.equal(await connections.applyPanelSlicerConnections(snapshot, connections.slicerConnectionKeys(snapshot.connections)), true)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('an incompatible same-range report injected into a draft is refused atomically by the content command', async () => {
  const f = setup()
  try {
    const snapshot = await connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer), invalid = option('pivotTable', ids.host, ids.incompatible)
    snapshot.options.push(invalid); const before = json(model(f)), baseline = history(f)
    assert.equal(await connections.applyPanelSlicerConnections(snapshot, [connections.slicerConnectionKey(invalid)]), false)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('a read-only panel cannot publish connection changes or alter its shared cache', async () => {
  const f = setup(slicerConnectionsWorkbook({ readonly: true }), 'peer')
  try {
    const snapshot = await connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer), before = json(model(f))
    assert.equal(panel.canConfigurePanelSlicer(f.studio, target(f), ids.slicer), false)
    assert.equal(await connections.applyPanelSlicerConnections(snapshot, snapshot.options.map(connections.slicerConnectionKey)), false)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), 0)
  } finally { f.dispose() }
})

for (const [name, steer] of [
  ['active worksheet', f => f.book.setActiveSheet(ids.other)],
  ['user identity', f => Object.defineProperty(f.studio, 'currentUserId', { value: 'peer', configurable: true })],
  ['history generation', f => f.studio.editor.history.clear(f.unitId)],
  ['content revision', f => f.book.getSlicer(ids.slicer).updatePresentation({ caption: 'Newer author' })],
]) test(`a captured connection panel refuses publication after its ${name} changes`, async () => {
  const f = setup()
  try {
    const snapshot = await connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer)
    await steer(f); const before = json(model(f)), baseline = history(f)
    assert.equal(connections.slicerConnectionsPanelCurrent(snapshot), false)
    await assert.rejects(connections.applyPanelSlicerConnections(snapshot, snapshot.options.map(connections.slicerConnectionKey)), code('STALE_CONNECTION_SOURCE'))
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

for (const [name, steer] of [
  ['worksheet switch', f => f.book.setActiveSheet(ids.other)],
  ['user switch', f => Object.defineProperty(f.studio, 'currentUserId', { value: 'peer', configurable: true })],
  ['history replacement', f => f.studio.editor.history.clear(f.unitId)],
  ['newer content', f => f.book.getSlicer(ids.slicer).updatePresentation({ caption: 'Newer content' })],
]) test(`candidate completion after ${name} cannot become a current panel snapshot`, async () => {
  const f = setup(), original = api.FSheetSlicer.prototype.getConnectionCandidates
  let release
  try {
    api.FSheetSlicer.prototype.getConnectionCandidates = async function() { await new Promise(resolve => { release = resolve }); return original.call(this) }
    const pending = connections.readPanelSlicerConnections(f.studio, target(f), ids.slicer)
    await Promise.resolve(); assert.equal(typeof release, 'function'); await steer(f)
    const before = json(model(f)), baseline = history(f); release()
    await assert.rejects(pending, code('STALE_CONNECTION_SOURCE')); assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
  } finally { release?.(); api.FSheetSlicer.prototype.getConnectionCandidates = original; f.dispose() }
})

test('independent chart insertion uses the canonical source field instead of its report-local field instance', async () => {
  const f = setup()
  try {
    const state = panel.readSlicerPanelState(f.studio, target(f)), source = state.pivots.find(source => source.kind === 'ownedPivotChart')
    assert.ok(source.fields.find(field => field.id === '0' && field.available)); assert.ok(source.fields.find(field => field.id === '1').controls.includes('timeline'))
    assert.equal(source.fields.some(field => field.id === 'owned-field-0'), false)
    const baseline = history(f), owner = owners.captureSlicerPanelOwner(f.studio, target(f))
    const id = await panel.insertPanelSlicer(f.studio, target(f), { ownedPivotChartId: ids.owned, sourceFieldId: '0', type: 'slicer' }, owners.slicerPanelOwnerContext(owner))
    const inserted = f.book.getSlicer(id).getModel()
    assert.equal(inserted.ownedPivotChartId, ids.owned); assert.equal(inserted.sourceFieldId, '0'); assert.equal(inserted.caption, 'Region')
    assert.deepEqual(f.book.getSlicer(id).getConnections(), { pivotTables: [], ownedPivotCharts: [{ subUnitId: ids.host, pivotChartId: ids.owned }] })
    assert.equal(history(f), baseline + 1)
  } finally { f.dispose() }
})

test('unchanged projected layout is omitted from a metadata patch and preserves full affine author geometry', async () => {
  const f = setup()
  try {
    const original = panel.readSlicerPanelState(f.studio, target(f)).controls.find(control => control.info.id === ids.slicer).info
    const draft = json(original); draft.caption = 'Presentation only'; draft.style = 'SlicerStyleLight2'
    const patch = panel.buildPanelSlicerPresentationPatch(draft, original), geometry = json(original.drawingGeometry)
    assert.deepEqual(patch, { caption: 'Presentation only', style: 'SlicerStyleLight2' })
    await panel.applyPanelSlicerPresentation(f.studio, target(f), ids.slicer, patch)
    assert.deepEqual(f.book.getSlicer(ids.slicer).getModel().drawingGeometry, geometry)
    assert.equal(history(f), 1)
  } finally { f.dispose() }
})

test('changing layout or anchor properties emits only the changed authored properties', () => {
  const original = slicerConnectionsWorkbook().slicers[0], draft = json(original)
  assert.deepEqual(panel.buildPanelSlicerPresentationPatch(draft, original), {})
  draft.layout.width += 10; draft.drawing.anchorType = 'none'; draft.drawing.keepRatio = true
  assert.deepEqual(panel.buildPanelSlicerPresentationPatch(draft, original), { layout: draft.layout, anchorType: 'none', keepRatio: true })
})

test('panel arrangement crosses a chart sibling in shared root order and remains one undoable content action', async () => {
  const workbook = slicerConnectionsWorkbook(), host = workbook.sheets.find(sheet => sheet.id === ids.host)
  workbook.slicers = workbook.slicers.filter(control => control.id === ids.slicer); workbook.charts = workbook.charts.filter(chart => chart.id === ids.owned)
  host.drawingState = { version: 1, revision: 0, groups: [], rootOrder: [{ kind: 'chart', id: ids.owned }, { kind: 'slicer', id: ids.slicer }] }
  const f = setup(workbook)
  try {
    const before = json(f.book.save())
    assert.deepEqual(panel.readPanelSlicerArrangement(f.studio, target(f), ids.slicer), { forward: false, front: false, backward: true, back: true })
    await panel.applyPanelSlicerPresentation(f.studio, target(f), ids.slicer, { arrange: 'backward' })
    assert.deepEqual(drawings.createSheetDrawingScene(model(f), ids.host).state.rootOrder, [{ kind: 'slicer', id: ids.slicer }, { kind: 'chart', id: ids.owned }])
    assert.equal(history(f), 1); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test('a protected chart sibling disables and rejects the otherwise editable slicer arrangement', async () => {
  const workbook = slicerConnectionsWorkbook(), host = workbook.sheets.find(sheet => sheet.id === ids.host)
  workbook.slicers = workbook.slicers.filter(control => control.id === ids.slicer); workbook.charts = workbook.charts.filter(chart => chart.id === ids.owned)
  host.drawingState = { version: 1, revision: 0, groups: [], rootOrder: [{ kind: 'chart', id: ids.owned }, { kind: 'slicer', id: ids.slicer }] }
  host.rangeProtectionRules = [{ id: 'chart-lock', permissionId: 'chart-lock', ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '',
    ranges: drawings.sheetDrawingRefAffectedRanges(workbook, ids.host, [{ kind: 'chart', id: ids.owned }]) }]
  const f = setup(workbook)
  try {
    const before = json(model(f)); assert.equal(panel.canConfigurePanelSlicer(f.studio, target(f), ids.slicer), true)
    assert.deepEqual(panel.readPanelSlicerArrangement(f.studio, target(f), ids.slicer), { forward: false, front: false, backward: false, back: false })
    await assert.rejects(panel.applyPanelSlicerPresentation(f.studio, target(f), ids.slicer, { arrange: 'backward' }), /权限|保护|permission|protected/i)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), 0)
  } finally { f.dispose() }
})

test('the slicer format inventory has four contextual roots in two upstream groups and preserves all regular root declarations', () => {
  const control = { id: ids.slicer, type: 'slicer', supportsConnections: true, canConfigure: true, revision: 7 }
  const before = ribbon.createSheetRibbon(ribbonState()), after = ribbon.createSheetRibbon(ribbonState(control)), format = after.find(tab => tab.id === 'slicer-format')
  assert.deepEqual(after.filter(tab => !tab.contextual).map(tab => [tab.id, rootItems(tab).map(item => item.id)]), before.filter(tab => !tab.contextual).map(tab => [tab.id, rootItems(tab).map(item => item.id)]))
  assert.deepEqual(format.groups.map(group => [group.id, group.order]), [['slicer-format.styles', 0], ['slicer-format.options', 1]])
  assert.deepEqual(rootItems(format).map(item => item.id), ['sheet.menu.slicer-quick-style', 'sheet.menu.slicer-settings', 'sheet.menu.timeline-settings', 'sheet.menu.slicer-connections'])
  assert.deepEqual(rootItems(format).map(item => item.gridLayout), [{ row: 1, column: 1, rowSpan: 2, showLabel: true },
    { row: 1, column: 1, showLabel: true }, { row: 1, column: 1, showLabel: true }, { row: 2, column: 1, showLabel: true }])
  const quick = rootItems(format)[0]
  assert.equal(quick.popupCommandId, 'sheet.command.update-slicer-presentation'); assert.equal(quick.popupProps.expectedRevision, 7); assert.equal(quick.popupProps.slicerId, ids.slicer)
})

for (const [type, supportsConnections, expected] of [
  ['slicer', true, ['sheet.menu.slicer-quick-style', 'sheet.menu.slicer-settings', 'sheet.menu.slicer-connections']],
  ['timeline', true, ['sheet.menu.slicer-quick-style', 'sheet.menu.timeline-settings', 'sheet.menu.slicer-connections']],
  ['slicer', false, ['sheet.menu.slicer-quick-style', 'sheet.menu.slicer-settings']],
]) test(`${type} with report connections ${supportsConnections} resolves the correct visible contextual entries`, () => {
  const tabs = schema.visibleRibbonTabs(ribbon.createSheetRibbon(ribbonState({ id: ids.slicer, type, supportsConnections, canConfigure: true, revision: 0 })))
  assert.deepEqual(rootItems(tabs.find(tab => tab.id === 'slicer-format')).map(item => item.id), expected)
})

test('contextual roots remain visible but cannot invoke commands when the selected control is not editable', () => {
  const tabs = schema.visibleRibbonTabs(ribbon.createSheetRibbon(ribbonState({ id: ids.slicer, type: 'slicer', supportsConnections: true, canConfigure: false, revision: 0 })))
  for (const item of rootItems(tabs.find(tab => tab.id === 'slicer-format'))) {
    assert.equal(item.disabled, true); assert.equal(schema.ribbonInvocation(item), undefined)
  }
})
