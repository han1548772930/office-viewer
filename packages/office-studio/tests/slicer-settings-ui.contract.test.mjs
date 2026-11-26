import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { slicerSettingsUiWorkbook } from './fixtures/slicer-settings-ui-fixture.mjs'
import { connectionFixtureIds as ids } from './fixtures/slicer-connections-fixture.mjs'

// Deferred real Facade/model/command/history/controller contracts. No checks have run.
let server, api, settings, cache, panel, hooks, rendering
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, settings, cache, panel, hooks, rendering] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-settings-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-cache-options-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/use-slicer-cache-options.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-renderer.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function setup(options = {}) {
  const studio = new api.OfficeStudio({ workbook: slicerSettingsUiWorkbook(options), currentUserId: 'local-user' })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), target = { workbookId: book.getId(), sheetId: ids.host }, commands = []
  studio.events.addEvent('CommandExecuted', event => commands.push(event))
  const model = () => studio.getWorkbookModel(book.getId()), control = (id = ids.slicer) => book.getSlicerControl(id)
  const state = () => json(book.save()), history = () => studio.editor.history.units[book.getId()]?.undo.length ?? 0
  const snapshot = (id = ids.slicer) => settings.readPanelSlicerSettings(studio, target, id)
  const draft = (id = ids.slicer) => settings.createSlicerPanelEditDraft(panel.readSlicerPanelState(studio, target).controls.find(value => value.info.id === id).info, snapshot(id).settings)
  return { studio, book, target, commands, model, control, state, history, snapshot, draft, dispose: () => studio.dispose() }
}
function retained(f) {
  const model = f.model()
  return json({ sheets: model.sheets.map(sheet => ({ id: sheet.id, cells: sheet.cells })), pivots: model.pivotTables, charts: model.charts, tables: model.tables,
    caches: model.slicerCaches?.map(({ id, source, connections, recordSelection, selectedItems, timelineRange }) => ({ id, source, connections, recordSelection, selectedItems, timelineRange })),
    selections: [ids.slicer, ids.shared, ids.table].map(id => f.control(id).getSelection()) })
}
function gate(method) {
  const original = api.FSheetSlicer.prototype[method], pending = []
  api.FSheetSlicer.prototype[method] = async function(...args) { await new Promise(resolve => pending.push(resolve)); return original.apply(this, args) }
  return { pending, release(index = 0) { pending.splice(index, 1)[0]?.() }, restore() { api.FSheetSlicer.prototype[method] = original; pending.splice(0).forEach(resolve => resolve()) } }
}
function controller(f, id = ids.slicer) {
  const calls = [], scope = effectScope(), options = shallowReactive({ studio: f.studio, ...f.target, slicerId: id, open: true,
    changed: id => calls.push(['changed', id]), close: owner => { calls.push(['close', owner]); options.open = false } })
  const controls = scope.run(() => hooks.useSlicerCacheOptions(options))
  return { calls, options, controls, stop: () => scope.stop() }
}
async function changeOwner(f, kind) {
  if (kind === 'source') f.book.getSheetBySheetId(ids.source).getRange('C2').setValue(11)
  if (kind === 'history') f.studio.editor.history.clear(f.target.workbookId)
  if (kind === 'user') Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
  if (kind === 'host') f.book.setActiveSheet(ids.other)
  await nextTick()
}

test.each([ids.slicer, ids.shared, ids.table])('%s draft uses real settings and preserves optional defaults without a content edit', id => {
  const f = setup({ viewDefaults: true })
  try {
    const before = f.state(), snapshot = f.snapshot(id), draft = f.draft(id)
    assert.equal(draft.caption, f.control(id).getModel().caption ?? f.control(id).getModel().name ?? '')
    assert.equal(draft.captionFromName, true); assert.equal(draft.rowHeight, 24); assert.equal(draft.startItem, undefined)
    assert.deepEqual(settings.buildPanelSlicerSettingsPatch(draft, f.draft(id)), {}); assert.deepEqual(settings.buildPanelSlicerGeometryPatch(draft, f.draft(id)), {})
    snapshot.settings.name = 'Detached only'; assert.notEqual(f.control(id).getSettings().name, snapshot.settings.name)
    assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['caption', 'style', 'startItem', 'columnCount', 'showCaption', 'lockedPosition'])('clearing optional %s preserves own undefined and does not flatten inherited values', key => {
  const f = setup({ authoredSettings: true })
  try {
    const original = f.draft(), current = f.draft()
    if (key === 'caption') current.captionFromName = true; else delete current[key]
    const patch = settings.buildPanelSlicerSettingsPatch(current, original)
    assert.deepEqual(Object.keys(patch), [key]); assert.equal(Object.hasOwn(patch, key), true); assert.equal(patch[key], undefined)
    assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each([ids.slicer, ids.table])('%s full display settings call the public setter once and preserve source, filter, shared views and geometry', async id => {
  const f = setup({ filtered: true }), snapshot = f.snapshot(id), before = f.state(), stable = retained(f), other = json(f.control(ids.shared).getModel())
  try {
    const patch = { name: `Named ${id}`, caption: 'Independent caption', rowHeight: 30.123456, columnCount: 2, startItem: 1,
      showCaption: false, lockedPosition: false, style: 'SlicerStyleDark2' }
    assert.equal(await settings.applyPanelSlicerSettings(snapshot, patch), true)
    const saved = f.control(id).getSettings()
    assert.equal(saved.rowHeight, Math.round(patch.rowHeight * 9525) / 9525)
    for (const key of ['name', 'caption', 'columnCount', 'startItem', 'showCaption', 'lockedPosition', 'style']) assert.deepEqual(saved[key], patch[key])
    assert.equal(f.commands.filter(command => command.id === 'sheet.command.update-slicer-settings').length, 1); assert.equal(f.history(), 1)
    assert.deepEqual(retained(f), stable); assert.deepEqual(f.control(id).getModel().drawingGeometry, before.slicers.find(control => control.id === id).drawingGeometry)
    if (id !== ids.shared) assert.deepEqual(f.control(ids.shared).getModel(), other)
    f.book.undo(); assert.deepEqual(f.state(), before); f.book.redo(); assert.deepEqual(f.control(id).getSettings(), saved)
  } finally { f.dispose() }
})

test.each([ids.slicer, ids.table])('%s mixed view and projected geometry use one validated presentation transaction', async id => {
  const f = setup(), before = f.state(), stable = retained(f), original = f.draft(id), current = f.draft(id)
  try {
    current.name = `Mixed ${id}`; current.caption = 'Mixed caption'; current.captionFromName = false; current.startItem = 1
    current.layout.width += 40; current.layout.height += 20; current.drawing.anchorType = 'none'; current.drawing.keepRatio = true
    const view = settings.buildPanelSlicerSettingsPatch(current, original), geometry = settings.buildPanelSlicerGeometryPatch(current, original)
    assert.deepEqual(Object.keys(geometry).sort(), ['anchorType', 'keepRatio', 'layout'])
    assert.equal(await settings.applyPanelSlicerSettings(f.snapshot(id), view, geometry), true)
    assert.equal(f.history(), 1); assert.equal(f.commands.filter(command => command.id === 'sheet.command.update-slicer-presentation').length, 1)
    assert.equal(f.commands.filter(command => command.id === 'sheet.command.update-slicer-settings').length, 0)
    assert.equal(f.control(id).getSettings().name, current.name); assert.equal(f.control(id).getModel().drawing.anchorType, 'none')
    assert.deepEqual(retained(f), stable); f.book.undo(); assert.deepEqual(f.state(), before)
  } finally { f.dispose() }
})

test('caption inheritance follows a renamed view without renaming its canonical source or cache', async () => {
  const f = setup({ captionInherited: true }), original = f.draft(), current = f.draft(), stable = retained(f)
  try {
    current.name = 'Renamed view'; const patch = settings.buildPanelSlicerSettingsPatch(current, original)
    assert.deepEqual(patch, { name: 'Renamed view' }); assert.equal(await settings.applyPanelSlicerSettings(f.snapshot(), patch), true)
    assert.equal(f.control().getSettings().caption, undefined); assert.equal(f.control().getModel().name, 'Renamed view')
    assert.deepEqual(retained(f), stable)
  } finally { f.dispose() }
})

test('reset display defaults clears authored optional keys in one command and survives JSON reopen', async () => {
  const f = setup({ authoredSettings: true }), original = f.draft(), current = f.draft(), before = f.state()
  let reopened
  try {
    current.captionFromName = true; current.rowHeight = 24
    for (const key of ['style', 'startItem', 'columnCount', 'showCaption', 'lockedPosition']) delete current[key]
    const patch = settings.buildPanelSlicerSettingsPatch(current, original)
    assert.equal(await settings.applyPanelSlicerSettings(f.snapshot(), patch), true); assert.equal(f.history(), 1)
    const event = f.commands.find(command => command.id === 'sheet.command.update-slicer-settings')
    for (const key of ['caption', 'style', 'startItem', 'columnCount', 'showCaption', 'lockedPosition']) assert.equal(Object.hasOwn(event.params.patch, key), true)
    const saved = f.control().getSettings(); assert.deepEqual(saved, { name: 'Region view', rowHeight: 24 })
    reopened = new api.OfficeStudio({ workbook: json(f.book.save()), currentUserId: 'local-user' })
    assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getSlicer(ids.slicer).getSettings(), saved)
    f.book.undo(); assert.deepEqual(f.state(), before)
  } finally { reopened?.dispose(); f.dispose() }
})

test.each([ids.shared, ids.timeline, ids.table])('duplicate name of %s refuses before a mixed layout edit or history', async otherId => {
  const f = setup(), current = f.draft(), original = f.draft(), before = f.state()
  try {
    current.name = ` ${f.control(otherId).getModel().name.toUpperCase()} `
    assert.match(settings.slicerSettingsPanelDraftError(f.snapshot(), current, original), /同名/)
    await assert.rejects(settings.applyPanelSlicerSettings(f.snapshot(), { name: current.name }, { layout: { ...current.layout, width: current.layout.width + 10 } }), /同名/)
    assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each([{ rowHeight: 0 }, { rowHeight: undefined }, { rowHeight: Number.POSITIVE_INFINITY }, { columnCount: 0 },
  { columnCount: 1.5 }, { startItem: -1 }, { startItem: 4294967296 }, { caption: '' }])('invalid display draft %j cannot dispatch a content command', async patch => {
  const f = setup(), before = f.state()
  try {
    await assert.rejects(settings.applyPanelSlicerSettings(f.snapshot(), patch)); assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
    assert.equal(f.commands.filter(command => command.id === 'sheet.command.update-slicer-settings').length, 0)
  } finally { f.dispose() }
})

test.each(['custom', 'unknown'])('%s document style is retained by the settings draft and uses the shared document palette', async kind => {
  const f = setup({ customStyles: true, customOverride: kind === 'custom', unknownStyle: kind === 'unknown' })
  try {
    const original = f.draft(), current = f.draft(), before = f.control().getSettings().style
    assert.deepEqual(settings.buildPanelSlicerSettingsPatch(current, original), {})
    current.caption = 'Style retained'; current.captionFromName = false
    assert.equal(await settings.applyPanelSlicerSettings(f.snapshot(), settings.buildPanelSlicerSettingsPatch(current, original)), true)
    assert.equal(f.control().getSettings().style, before)
    if (kind === 'unknown') assert.deepEqual(rendering.slicerPalette(before, false, undefined, f.model()), rendering.slicerPalette(undefined, false, undefined, f.model()))
    else assert.equal(rendering.slicerPalette(before, false, undefined, f.model()).header.toUpperCase(), '#103F77')
  } finally { f.dispose() }
})

test('cache Select sentinels preserve optional author presence, explicit false and unrelated keys', () => {
  const original = { sortOrder: 'ascending', customListSort: true, crossFilter: 'showItemsWithNoData', hideItemsWithNoData: true }
  let current = cache.setPanelSlicerCacheOption(original, 'customListSort', 'false')
  assert.deepEqual(cache.buildPanelSlicerCacheOptionsPatch(current, original), { customListSort: false })
  current = cache.setPanelSlicerCacheOption(current, 'sortOrder', cache.SLICER_CACHE_OPTIONS_DEFAULT)
  const patch = cache.buildPanelSlicerCacheOptionsPatch(current, original)
  assert.equal(Object.hasOwn(patch, 'sortOrder'), true); assert.equal(patch.sortOrder, undefined); assert.equal(patch.customListSort, false)
  assert.deepEqual(original, { sortOrder: 'ascending', customListSort: true, crossFilter: 'showItemsWithNoData', hideItemsWithNoData: true })
  assert.throws(() => cache.setPanelSlicerCacheOption(current, 'crossFilter', 'unknown'))
})

test.each([ids.slicer, ids.table])('%s cache options publish shared state and member order while preserving all filters and placement', async id => {
  const f = setup({ filtered: true }), before = f.state(), stable = retained(f), snapshot = cache.readPanelSlicerCacheOptions(f.studio, f.target, id)
  try {
    assert.equal(await cache.applyPanelSlicerCacheOptions(snapshot, { sortOrder: 'descending', customListSort: false }), true)
    const options = f.control(id).getCacheOptions(), window = await f.control(id).getItems({ count: 20 })
    assert.deepEqual(options, { sortOrder: 'descending', customListSort: false }); assert.deepEqual(window.items.map(item => item.value.value), ['South', 'North'])
    if (id === ids.slicer) assert.deepEqual(f.control(ids.shared).getCacheOptions(), options)
    else assert.deepEqual(f.control(ids.slicer).getCacheOptions(), {})
    assert.equal(f.history(), 1); assert.deepEqual(retained(f), stable)
    assert.deepEqual(f.control(id).getModel().drawingGeometry, before.slicers.find(control => control.id === id).drawingGeometry)
    f.book.undo(); assert.deepEqual(f.state(), before); f.book.redo(); assert.deepEqual(f.control(id).getCacheOptions(), options)
  } finally { f.dispose() }
})

test.each(['none', 'showItemsWithDataAtTop', 'showItemsWithNoData'].flatMap(crossFilter => [false, true].map(hide => [crossFilter, hide])))
  ('real March context %s / hide=%s changes member projection without changing retained North selection', async (crossFilter, hide) => {
    const f = setup({ marchContext: true }), stable = retained(f)
    try {
      await f.control().setItems(['North']); const selection = f.control().getSelection(), beforeOptions = retained(f)
      assert.equal(await cache.applyPanelSlicerCacheOptions(cache.readPanelSlicerCacheOptions(f.studio, f.target, ids.slicer),
        { crossFilter, hideItemsWithNoData: hide }), true)
      const members = await f.control().getItems({ count: 20 })
      assert.deepEqual(members.items.map(item => item.value.value), crossFilter === 'none' ? ['North', 'South'] : hide ? ['South']
        : crossFilter === 'showItemsWithDataAtTop' ? ['South', 'North'] : ['North', 'South'])
      assert.deepEqual(members.items.map(item => item.hasData), crossFilter === 'none' ? [true, true] : hide ? [true]
        : crossFilter === 'showItemsWithDataAtTop' ? [true, false] : [false, true])
      assert.deepEqual(f.control().getSelection(), selection); assert.deepEqual(retained(f), beforeOptions); assert.deepEqual(stable.caches.map(cache => cache.source), beforeOptions.caches.map(cache => cache.source))
    } finally { f.dispose() }
  })

test('cache reset clears own optional keys once, no-op rereads preserve history, and author defaults survive JSON reopen', async () => {
  const f = setup({ cacheOptions: { sortOrder: 'descending', customListSort: true, crossFilter: 'showItemsWithDataAtTop', hideItemsWithNoData: true } })
  let reopened
  try {
    const initial = cache.readPanelSlicerCacheOptions(f.studio, f.target, ids.slicer), patch = cache.buildPanelSlicerCacheOptionsPatch({}, initial.options)
    for (const key of Object.keys(initial.options)) { assert.equal(Object.hasOwn(patch, key), true); assert.equal(patch[key], undefined) }
    assert.equal(await cache.applyPanelSlicerCacheOptions(initial, patch), true); assert.deepEqual(f.control().getCacheOptions(), {}); assert.equal(f.history(), 1)
    assert.equal(await cache.applyPanelSlicerCacheOptions(cache.readPanelSlicerCacheOptions(f.studio, f.target, ids.slicer), {}), true); assert.equal(f.history(), 1)
    reopened = new api.OfficeStudio({ workbook: json(f.book.save()), currentUserId: 'local-user' })
    assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getSlicer(ids.slicer).getCacheOptions(), {})
  } finally { reopened?.dispose(); f.dispose() }
})

test.each(['settings', 'cache'])('%s read of Timeline is unavailable and cannot masquerade as Slicer options', kind => {
  const f = setup()
  try {
    assert.throws(() => kind === 'settings' ? f.snapshot(ids.timeline) : cache.readPanelSlicerCacheOptions(f.studio, f.target, ids.timeline))
    assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['settings', 'cache'])('%s viewable read-only host permits read and refuses edits without content or history', async kind => {
  const f = setup({ readonly: true }), before = f.state()
  try {
    const accepted = kind === 'settings' ? await settings.applyPanelSlicerSettings(f.snapshot(), { name: 'Refused' })
      : await cache.applyPanelSlicerCacheOptions(cache.readPanelSlicerCacheOptions(f.studio, f.target, ids.slicer), { sortOrder: 'descending' })
    assert.equal(accepted, false); assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['settings', 'cache'].flatMap(kind => ['source', 'history', 'user', 'host', 'closed', 'disposed'].map(change => [kind, change])))
  ('pending %s rejects changed %s at real dispatch before any partial edit', async (kind, change) => {
    const f = setup(), delayed = gate(kind === 'settings' ? 'updateSettings' : 'setCacheOptions'); let active = true
    try {
      const pending = kind === 'settings' ? settings.applyPanelSlicerSettings(f.snapshot(), { name: 'Late old view' }, {}, () => active)
        : cache.applyPanelSlicerCacheOptions(cache.readPanelSlicerCacheOptions(f.studio, f.target, ids.slicer), { sortOrder: 'descending' }, () => active)
      const rejection = assert.rejects(pending)
      if (change === 'closed') active = false; else if (change === 'disposed') f.dispose(); else await changeOwner(f, change)
      const before = change === 'disposed' ? undefined : f.state(), history = change === 'disposed' ? undefined : f.history()
      delayed.release(); await rejection
      if (before) { assert.deepEqual(f.state(), before); assert.equal(f.history(), history) }
    } finally { delayed.restore(); if (change !== 'disposed') f.dispose() }
  })

test.each(['settings', 'cache'].flatMap(kind => ['history', 'user'].map(change => [kind, change])))
  ('%s owner guard rechecks %s changed by a later BeforeCommandExecute listener', async (kind, change) => {
    const f = setup(), commandId = kind === 'settings' ? 'sheet.command.update-slicer-settings' : 'sheet.command.patch-slicer-cache-options'
    let listener
    try {
      const pending = kind === 'settings' ? settings.applyPanelSlicerSettings(f.snapshot(), { caption: 'Old owner' })
        : cache.applyPanelSlicerCacheOptions(cache.readPanelSlicerCacheOptions(f.studio, f.target, ids.slicer), { sortOrder: 'descending' })
      listener = f.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === commandId) {
        if (change === 'history') f.studio.editor.history.clear(f.target.workbookId)
        else Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
      } })
      const before = f.state(); await assert.rejects(pending); assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
    } finally { listener?.dispose(); f.dispose() }
  })

test('cache controller accepts its own revision once and closes only after the real command publishes', async () => {
  const f = setup(), c = controller(f)
  try {
    assert.equal(c.controls.dirty.value, false); c.controls.set('sortOrder', 'descending'); assert.equal(c.controls.dirty.value, true)
    assert.equal(await c.controls.apply(), true); assert.equal(f.history(), 1)
    assert.equal(c.calls.filter(call => call[0] === 'changed').length, 1); assert.equal(c.calls.filter(call => call[0] === 'close').length, 1)
    assert.deepEqual(f.control().getCacheOptions(), { sortOrder: 'descending' })
  } finally { c.stop(); f.dispose() }
})

test.each(['history', 'external cache command'])('an old cache guard and finally preserve a newer pending edit after %s replacement', async change => {
  const f = setup(), c = controller(f), delayed = gate('setCacheOptions')
  try {
    c.controls.set('sortOrder', 'descending'); const old = c.controls.apply(); assert.equal(delayed.pending.length, 1)
    if (change === 'history') f.studio.editor.history.clear(f.target.workbookId)
    else await f.studio.executeCommand('sheet.command.patch-slicer-cache-options', { unitId: f.target.workbookId, slicerId: ids.slicer,
      patch: { sortOrder: 'ascending' }, expectedRevision: c.controls.snapshot.value.owner.revision, host: 'workbench' })
    await nextTick(); assert.equal(c.controls.busy.value, false); assert.equal(c.controls.stale.value, true)
    const baseline = f.history()
    c.controls.read(); c.controls.set('sortOrder', 'descending'); const fresh = c.controls.apply(); assert.equal(delayed.pending.length, 2)
    delayed.release(0); assert.equal(await old, false); assert.equal(c.controls.busy.value, true); assert.equal(c.controls.stale.value, false); assert.equal(f.history(), baseline)
    delayed.release(0); assert.equal(await fresh, true); assert.equal(f.history(), baseline + 1); assert.deepEqual(f.control().getCacheOptions(), { sortOrder: 'descending' })
  } finally { delayed.restore(); c.stop(); f.dispose() }
})

test.each(['closed', 'disposed'])('cache controller %s revokes a pending command without changed callbacks', async kind => {
  const f = setup(), c = controller(f), delayed = gate('setCacheOptions'), before = f.state()
  try {
    c.controls.set('sortOrder', 'descending'); const pending = c.controls.apply()
    if (kind === 'closed') c.controls.close(); else c.stop()
    delayed.release(); assert.equal(await pending, false); assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
    assert.equal(c.calls.filter(call => call[0] === 'changed').length, 0)
  } finally { delayed.restore(); c.stop(); f.dispose() }
})

test('cache reset preserves optional defaults until apply and read-only controls cannot mutate the draft', async () => {
  const f = setup({ readonly: true, cacheOptions: { customListSort: true } }), c = controller(f), before = f.state()
  try {
    assert.equal(c.controls.canConfigure.value, false); c.controls.set('customListSort', 'false'); c.controls.reset()
    assert.deepEqual(c.controls.draft.value, { customListSort: true }); assert.equal(await c.controls.apply(), false); assert.deepEqual(f.state(), before)
  } finally { c.stop(); f.dispose() }
})
