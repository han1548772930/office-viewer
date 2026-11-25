import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { createBlankSheet } from '../src/domain/model.ts'
import { createSheetRangeProtectionRule } from '../src/features/sheet-protection/range-protection.ts'
import { drawingFrame, drawingRef as ref, sheetDrawingWorkbook } from './fixtures/sheet-drawing-model-fixture.mjs'

// Deferred integration drafts. These hooks have not been invoked: no Vite/dev, runner, browser, Office or WASM was started.
let server, api, commands, drawings, broker, nativeHost, chartCommands
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  commands = await server.ssrLoadModule('/packages/office-studio/src/api/sheet-drawing-commands.ts')
  drawings = await server.ssrLoadModule('/packages/office-studio/src/features/drawing-groups/index.ts')
  broker = await server.ssrLoadModule('/packages/office-studio/src/api/sheet-drawing-selection.ts')
  nativeHost = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-drawings/sheet-drawing-host.ts')
  chartCommands = await server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-object-commands.ts')
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const keys = refs => refs.map(item => item.kind + ':' + item.id)
const history = (studio, id) => studio.editor.history.units[id]?.undo.length ?? 0
function setup(model = sheetDrawingWorkbook()) {
  const studio = new api.OfficeStudio({ workbook: model }), facade = api.FOffice.newAPI(studio), workbook = facade.getActiveWorkbook()
  return { studio, facade, workbook, id: workbook.getId(), sheetId: 'data', ids: drawings.SHEET_DRAWING_GROUP_COMMAND_IDS }
}
function selectionOwner(studio, id, initial = [ref('image', 'a'), ref('chart', 'c')]) {
  let selection = { sheetId: 'data', refs: initial }
  const binding = broker.bindSheetDrawingSelection(studio, id, () => selection, value => { selection = json(value) })
  return { binding, read: () => selection, set: refs => { selection = { sheetId: 'data', refs } } }
}
const scene = (studio, id) => nativeHost.createSheetDrawingScene(studio.getWorkbookModel(id), 'data')

test('registered grouping and affine transformation each commit one history entry and restore content/selection on undo/redo and JSON reopening', async () => {
  const { studio, workbook, id, ids } = setup(), owner = selectionOwner(studio, id)
  let reopened
  try {
    const original = json(workbook.save()), count = history(studio, id), revision = studio.getWorkbookRevision(id)
    assert.equal(await studio.executeCommand(ids.group, { unitId: id, sheetId: 'data', groupId: 'g' }), true)
    assert.equal(history(studio, id), count + 1); assert.equal(studio.getWorkbookRevision(id), revision + 1)
    assert.deepEqual(owner.read().refs, [ref('group', 'g')]); assert.deepEqual(workbook.save().sheets[0].images, original.sheets[0].images)
    const grouped = json(workbook.save())
    workbook.undo(); assert.deepEqual(json(workbook.save()), original); assert.deepEqual(owner.read().refs, [ref('image', 'a'), ref('chart', 'c')])
    workbook.redo(); assert.deepEqual(json(workbook.save()), grouped); assert.deepEqual(owner.read().refs, [ref('group', 'g')])
    assert.equal(await studio.executeCommand(ids.transform, { unitId: id, sheetId: 'data', groupId: 'g', frame: drawingFrame(100, 200, 140, 120) }), true)
    assert.equal(history(studio, id), count + 2)
    const saved = json(workbook.save()), expected = scene(studio, id)
    assert.deepEqual(drawings.drawingGeometryBounds(expected.objects.find(item => item.ref.kind === 'image' && item.ref.id === 'a').geometry), { left: 100, top: 200, width: 40, height: 30 })
    workbook.undo(); assert.deepEqual(json(workbook.save()), grouped); assert.deepEqual(owner.read().refs, [ref('group', 'g')])
    workbook.redo(); assert.deepEqual(json(workbook.save()), saved)
    reopened = new api.OfficeStudio({ workbook: saved })
    assert.deepEqual(json(nativeHost.createSheetDrawingScene(reopened.getWorkbookModel(id), 'data')), json(expected))
    assert.deepEqual(reopened.getWorkbookModel(id).charts[0].source, original.charts[0].source)
    assert.deepEqual(reopened.getWorkbookModel(id).sheets[0].images[0].resource, original.sheets[0].images[0].resource)
  } finally { owner.binding.dispose(); reopened?.dispose(); studio.dispose() }
})

test('nested ungrouping preserves transformed leaf geometry and author payload, restores group selection on undo and children on redo', async () => {
  const { studio, workbook, id, ids } = setup(), owner = selectionOwner(studio, id)
  try {
    await studio.executeCommand(ids.group, { unitId: id, sheetId: 'data', selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'inner' })
    await studio.executeCommand(ids.group, { unitId: id, sheetId: 'data', selection: [ref('group', 'inner'), ref('shape', 'b')], groupId: 'outer' })
    await studio.executeCommand(ids.transform, { unitId: id, sheetId: 'data', groupId: 'outer', delta: { a: 1, b: 0.5, c: 0.25, d: 1, e: 3, f: -2 } })
    const transformed = json(workbook.save()), count = history(studio, id), leafGeometry = scene(studio, id).objects.map(item => item.geometry)
    assert.equal(await studio.executeCommand(ids.ungroup, { unitId: id, sheetId: 'data', groupId: 'outer' }), true)
    assert.equal(history(studio, id), count + 1); assert.deepEqual(scene(studio, id).objects.map(item => item.geometry), leafGeometry)
    assert.equal(scene(studio, id).state.groups.length, 1)
    const childSelection = scene(studio, id).state.rootOrder.filter(item => item.id === 'inner' || item.id === 'b')
    assert.deepEqual(owner.read().refs, childSelection)
    assert.deepEqual(workbook.save().sheets[0].images, transformed.sheets[0].images)
    assert.deepEqual(workbook.save().sheets[0].shapes, transformed.sheets[0].shapes)
    const expanded = json(workbook.save())
    workbook.undo(); assert.deepEqual(json(workbook.save()), transformed); assert.deepEqual(owner.read().refs, [ref('group', 'outer')])
    workbook.redo(); assert.deepEqual(json(workbook.save()), expanded); assert.deepEqual(owner.read().refs, childSelection)
  } finally { owner.binding.dispose(); studio.dispose() }
})

test('public parent/children batches infer the target from the first child, group every batch in one history and select all new sibling groups', async () => {
  const { studio, facade, workbook, id, ids } = setup(), emitted = []
  let event
  try {
    const other = facade.createWorkbook(sheetDrawingWorkbook({ id: 'other-drawing-book' })), otherBefore = json(other.save()), count = history(studio, id)
    event = studio.events.addEvent('CommandExecuted', value => emitted.push(value))
    const child = (kind, drawingId) => ({ unitId: id, subUnitId: 'data', drawingId, drawingType: { image: 0, shape: 1, chart: 2 }[kind] })
    const batches = [{ parent: { drawingId: 'first', drawingType: 6 }, children: [child('image', 'a'), child('chart', 'c')] },
      { parent: { drawingId: 'second', drawingType: 6 }, children: [child('shape', 'b'), child('image', 'd')] }]
    assert.equal(await studio.executeCommand(ids.group, batches), true)
    assert.equal(history(studio, id), count + 1); assert.deepEqual(new Set(scene(studio, id).state.groups.map(item => item.id)), new Set(['first', 'second']))
    assert.deepEqual(broker.getSheetDrawingSelection(studio, id, 'data'), [ref('group', 'first'), ref('group', 'second')])
    assert.deepEqual(json(other.save()), otherBefore)
    assert.equal(emitted.find(value => value.id === ids.group)?.unitId, id)
    workbook.undo(); assert.deepEqual(new Set(keys(broker.getSheetDrawingSelection(studio, id, 'data'))), new Set(['image:a', 'chart:c', 'shape:b', 'image:d']))
    workbook.redo(); assert.deepEqual(broker.getSheetDrawingSelection(studio, id, 'data'), [ref('group', 'first'), ref('group', 'second')])
  } finally { event?.dispose(); studio.dispose() }
})

test('public singleton batches with a transform remain no-ops; malformed flips, parent types and conflicting frame/delta reject without history', async () => {
  const { studio, workbook, id, ids } = setup()
  try {
    const before = json(workbook.save()), count = history(studio, id)
    const child = { unitId: id, subUnitId: 'data', drawingId: 'a', drawingType: 0 }
    assert.equal(await studio.executeCommand(ids.group, [{ parent: { drawingId: 'single', transform: { left: 44, top: 26, width: 20, height: 10 } }, children: [child] }]), false)
    await assert.rejects(studio.executeCommand(ids.group, [{ parent: { drawingId: 'bad', transform: { left: 44, top: 26, width: 20, height: 10, flipX: 'false' } }, children: [child] }]), /布尔/)
    await assert.rejects(studio.executeCommand(ids.group, [{ parent: { drawingId: 'bad', drawingType: 0 }, children: [child] }]), /父项/)
    await assert.rejects(studio.executeCommand(ids.group, { unitId: id, selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'bad',
      frame: drawingFrame(10, 20, 70, 40), delta: { a: 1, b: 0, c: 0, d: 1, e: 2, f: 0 } }), /只能/)
    assert.equal(await studio.executeCommand(ids.group, []), false)
    assert.deepEqual(json(workbook.save()), before); assert.equal(history(studio, id), count)
  } finally { studio.dispose() }
})

test('canonical public drawing references with numeric kinds infer their host and apply a requested group frame in the same history entry', async () => {
  const { studio, workbook, id, ids } = setup()
  try {
    const count = history(studio, id), frame = drawingFrame(100, 200, 140, 120)
    assert.equal(commands.sheetDrawingCommandInput({ drawings: [{ unitId: id, subUnitId: 'data', drawingId: 'a', drawingType: 0 }] }).unitId, id)
    assert.equal(await studio.executeCommand(ids.group, { groupId: 'g', frame,
      drawings: [{ unitId: id, subUnitId: 'data', drawingId: 'a', drawingType: 0 }, { unitId: id, subUnitId: 'data', drawingId: 'c', drawingType: 2 }] }), true)
    assert.equal(history(studio, id), count + 1); assert.deepEqual(scene(studio, id).state.groups[0].frame, frame)
    assert.deepEqual(drawings.drawingGeometryBounds(scene(studio, id).objects.find(item => item.ref.id === 'c').geometry), { left: 180, top: 260, width: 60, height: 60 })
    assert.deepEqual(broker.getSheetDrawingSelection(studio, id, 'data'), [ref('group', 'g')])
    const after = json(workbook.save()); workbook.undo(); assert.equal(scene(studio, id).state.groups.length, 0); workbook.redo(); assert.deepEqual(json(workbook.save()), after)
  } finally { studio.dispose() }
})

test('the existing arrangement registration delegates to mixed typed ordering with stable numeric public actions and one undoable commit', async () => {
  const model = sheetDrawingWorkbook()
  model.sheets[0].drawingState = { version: 1, revision: 0, groups: [], rootOrder: [ref('image', 'a'), ref('shape', 'b'), ref('chart', 'c'), ref('image', 'd')] }
  const { studio, workbook, id, ids } = setup(model)
  try {
    const original = json(workbook.save()), count = history(studio, id)
    assert.equal(await studio.executeCommand(ids.arrange, { unitId: id, subUnitId: 'data', selection: [ref('chart', 'c'), ref('image', 'a')], arrangeType: 2 }), true)
    assert.deepEqual(keys(scene(studio, id).state.rootOrder), ['shape:b', 'image:d', 'image:a', 'chart:c'])
    assert.equal(history(studio, id), count + 1)
    assert.equal(await studio.executeCommand(ids.arrange, { unitId: id, subUnitId: 'data', selection: [ref('chart', 'c')], arrangeType: 2 }), false)
    const after = json(workbook.save()); workbook.undo(); assert.deepEqual(json(workbook.save()), original); workbook.redo(); assert.deepEqual(json(workbook.save()), after)
  } finally { studio.dispose() }
})

test('stale content/history/host/selection, Abort and denied old/new footprint permissions leave content, history and selection untouched', async () => {
  const { studio, workbook, id, ids } = setup(), owner = selectionOwner(studio, id)
  const permission = studio.assertCanEditRange
  try {
    const before = json(workbook.save()), count = history(studio, id), revision = studio.getWorkbookRevision(id), generation = studio.editor.history.captureSavepoint(id).generation
    const base = { unitId: id, sheetId: 'data', groupId: 'g', selection: [ref('image', 'a'), ref('chart', 'c')] }
    const canceled = new AbortController(); canceled.abort()
    for (const extra of [{ expectedRevision: revision + 1 }, { expectedHistoryGeneration: generation + 1 }, { expectedSelection: [ref('image', 'd')] },
      { host: { unitId: 'different', sheetId: 'data' } }, { signal: canceled.signal }, { subUnitId: 'other-sheet' }]) {
      await assert.rejects(studio.executeCommand(ids.group, { ...base, ...extra }))
    }
    studio.assertCanEditRange = () => { throw new Error('blocked object footprint') }
    await assert.rejects(studio.executeCommand(ids.group, base), /blocked object footprint/)
    assert.deepEqual(json(workbook.save()), before); assert.equal(history(studio, id), count); assert.equal(studio.getWorkbookRevision(id), revision)
    assert.deepEqual(owner.read().refs, [ref('image', 'a'), ref('chart', 'c')])
  } finally { studio.assertCanEditRange = permission; owner.binding.dispose(); studio.dispose() }
})

test('mixed Delete honors slicer veto/events, keeps its source table/media owner and restores the full group and selection through real undo/redo', async () => {
  const { studio, workbook, id, ids } = setup(sheetDrawingWorkbook({ slicer: true })), owner = selectionOwner(studio, id), removed = []
  let veto, events
  try {
    await studio.executeCommand(ids.group, { unitId: id, sheetId: 'data', selection: [ref('image', 'a'), ref('shape', 'b'), ref('chart', 'c')], groupId: 'g' })
    owner.set([ref('group', 'g'), ref('slicer', 's')])
    const before = json(workbook.save()), count = history(studio, id)
    veto = studio.events.addEvent('BeforeSlicerChange', event => { assert.equal(event.commandId, 'sheet.command.remove-slicer'); event.cancel = true })
    events = studio.events.addEvent('SlicerRemoved', event => removed.push(event))
    const input = { unitId: id, sheetId: 'data', selection: owner.read().refs, expectedSelection: owner.read().refs }
    await assert.rejects(studio.executeCommand(ids.remove, input), /取消/)
    assert.deepEqual(json(workbook.save()), before); assert.equal(history(studio, id), count); assert.equal(removed.length, 0)
    veto.dispose()
    assert.equal(await studio.executeCommand(ids.remove, input), true)
    assert.equal(history(studio, id), count + 1); assert.deepEqual(removed, [{ unitId: id, hostSheetId: 'data', slicerId: 's' }])
    assert.deepEqual(workbook.save().tables, before.tables); assert.deepEqual(workbook.save().sheets[0].images, [before.sheets[0].images.find(image => image.id === 'd')])
    assert.equal(workbook.save().slicers.length, 0); assert.equal(workbook.save().charts.length, 0); assert.deepEqual(owner.read().refs, [])
    const deleted = json(workbook.save()); workbook.undo(); assert.deepEqual(json(workbook.save()), before); assert.deepEqual(owner.read().refs, input.selection)
    workbook.redo(); assert.deepEqual(json(workbook.save()), deleted); assert.deepEqual(owner.read().refs, [])
  } finally { veto?.dispose(); events?.dispose(); owner.binding.dispose(); studio.dispose() }
})

test('primary chart deletion and a separately protected chart source cannot be bypassed through group commands', async () => {
  const { studio, workbook, id, ids } = setup()
  try {
    await studio.executeCommand(ids.group, { unitId: id, sheetId: 'data', selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'g' })
    studio.getWorkbookModel(id).sheets[0].metadata = { [chartCommands.CHART_SHEET_METADATA_KEY]: { version: 1, primaryChartId: 'c' } }
    const before = json(workbook.save()), count = history(studio, id)
    await assert.rejects(studio.executeCommand(ids.remove, { unitId: id, sheetId: 'data', groupId: 'g' }), /主图/)
    assert.deepEqual(json(workbook.save()), before); assert.equal(history(studio, id), count)
  } finally { studio.dispose() }
  const model = sheetDrawingWorkbook(), source = createBlankSheet('Source')
  source.id = 'source'; source.cells = json(model.sheets[0].cells)
  source.rangeProtectionRules = [createSheetRangeProtectionRule({ id: 'source-read-rule', permissionId: 'source-read', editState: 'onlyMe', viewState: 'noOneElseCanView' },
    [{ startRow: 0, startCol: 0, endRow: 2, endCol: 1 }], 'other-owner')]
  model.sheets.push(source); model.charts[0].source.sheetId = source.id
  const protectedHost = setup(model)
  try {
    const before = json(protectedHost.workbook.save()), count = history(protectedHost.studio, protectedHost.id)
    await assert.rejects(protectedHost.studio.executeCommand(protectedHost.ids.group, { unitId: protectedHost.id, sheetId: 'data', selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'blocked' }), /图表源区域/)
    assert.deepEqual(json(protectedHost.workbook.save()), before); assert.equal(history(protectedHost.studio, protectedHost.id), count)
  } finally { protectedHost.studio.dispose() }
})

test('selection caches isolate API instances/workbooks/worksheets and disposal of a replaced owner cannot clear its new owner', () => {
  const first = setup(), second = setup(), writes = []
  let original, replacement
  try {
    original = broker.bindSheetDrawingSelection(first.studio, first.id, () => ({ sheetId: 'data', refs: [ref('image', 'a')] }), value => writes.push(['old', value]))
    broker.setSheetDrawingSelection(first.studio, first.id, 'other-sheet', [ref('shape', 'other')])
    assert.deepEqual(broker.getSheetDrawingSelection(first.studio, first.id, 'other-sheet'), [ref('shape', 'other')]); assert.equal(writes.length, 0)
    broker.setSheetDrawingSelection(second.studio, second.id, 'data', [ref('image', 'd')])
    assert.deepEqual(broker.getSheetDrawingSelection(second.studio, second.id, 'data'), [ref('image', 'd')])
    assert.deepEqual(broker.getSheetDrawingSelection(first.studio, first.id, 'data'), [ref('image', 'a')])
    replacement = broker.bindSheetDrawingSelection(first.studio, first.id, () => ({ sheetId: 'data', refs: [ref('chart', 'c')] }), value => writes.push(['new', value]))
    original.dispose(); broker.setSheetDrawingSelection(first.studio, first.id, 'data', [ref('chart', 'c')])
    assert.equal(writes.length, 1); assert.equal(writes[0][0], 'new')
    const other = first.facade.createWorkbook(sheetDrawingWorkbook({ id: 'inactive-selection' }))
    broker.setSheetDrawingSelection(first.studio, first.id, 'data', [ref('image', 'a')]); assert.equal(writes.length, 1)
    assert.deepEqual(broker.getSheetDrawingSelection(first.studio, first.id, 'data'), [ref('image', 'a')])
    broker.setSheetDrawingSelection(first.studio, other.getId(), 'data', [ref('image', 'd')])
    assert.deepEqual(broker.getSheetDrawingSelection(first.studio, other.getId(), 'data'), [ref('image', 'd')])
    broker.clearSheetDrawingSelections(first.studio, first.id)
    assert.deepEqual(broker.getSheetDrawingSelection(first.studio, first.id, 'data'), [])
    assert.deepEqual(broker.getSheetDrawingSelection(second.studio, second.id, 'data'), [ref('image', 'd')])
  } finally { original?.dispose(); replacement?.dispose(); first.studio.dispose(); second.studio.dispose() }
})

test('chart mask pruning during group Delete is part of the same content snapshot and real undo/redo restores both registry and owners', async () => {
  const model = sheetDrawingWorkbook(), resource = { id: 'mask', sourceType: 'url', source: 'https://example.invalid/mask.png', mimeType: 'image/png', sizeBytes: 8 }
  model.charts[0].type = 'wordCloud'; model.charts[0].options.wordCloud = { maskImageId: 'mask' }
  model.metadata = { keep: 'other metadata', chartImageResources: { mask: resource } }
  const { studio, workbook, id, ids } = setup(model)
  try {
    await studio.executeCommand(ids.group, { unitId: id, sheetId: 'data', selection: [ref('image', 'a'), ref('chart', 'c')], groupId: 'g' })
    const before = json(workbook.save()), count = history(studio, id)
    assert.equal(await studio.executeCommand(ids.remove, { unitId: id, sheetId: 'data', groupId: 'g' }), true)
    assert.equal(history(studio, id), count + 1); assert.equal(workbook.save().metadata.chartImageResources, undefined)
    assert.equal(workbook.save().metadata.keep, 'other metadata')
    const after = json(workbook.save()); workbook.undo(); assert.deepEqual(json(workbook.save()), before)
    workbook.redo(); assert.deepEqual(json(workbook.save()), after)
  } finally { studio.dispose() }
})
