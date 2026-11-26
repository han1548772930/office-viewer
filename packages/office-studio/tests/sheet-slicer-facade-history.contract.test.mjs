import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { drawingAffineBounds } from '../src/domain/drawing-affine.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../src/rendering/spreadsheet/layout.ts'
import { sheetClipboardObjectRange } from '../src/features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { slicerDrawingAuthor, slicerDrawingRefs, slicerDrawingWorkbook } from './fixtures/slicer-drawing-model-fixture.mjs'

// Deferred registry/Facade/history integration. No server hook or test below has been executed.
let server, api, host, broker
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, host, broker] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/sheet-drawings/sheet-drawing-host.ts'),
    server.ssrLoadModule('/packages/office-studio/src/api/sheet-drawing-selection.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const history = f => f.studio.editor.history.units[f.unitId]?.undo.length ?? 0
const model = f => f.studio.getWorkbookModel(f.unitId)
const control = (f, id = 's') => model(f).slicers.find(control => control.id === id)
const leaf = (f, id = 's') => ({ kind: control(f, id).type, id })
const delta = { a: 1, b: .15, c: .2, d: 1, e: 50, f: 30 }
const translate = { a: 1, b: 0, c: 0, d: 1, e: 600, f: 400 }
function setup(value = slicerDrawingWorkbook({ filtered: true })) {
  const studio = new api.OfficeStudio({ workbook: value }), facade = api.FOffice.newAPI(studio), book = facade.getActiveWorkbook()
  return { studio, facade, book, sheet: book.getActiveSheet(), unitId: book.getId(), dispose: () => studio.dispose() }
}
function presentation(f, movement = translate, id = 's') {
  const geometry = control(f, id).drawingGeometry, matrix = geometry.matrix
  return { drawingGeometry: { ...geometry, matrix: { ...matrix, e: matrix.e + movement.e, f: matrix.f + movement.f } } }
}
async function change(f, mode, movement = translate) {
  return mode === 'native' ? f.book.getSlicer('s').updatePresentation(presentation(f, movement))
    : f.sheet.transformDrawings([leaf(f)], { delta: movement })
}

test('five native leaf kinds form one group, preserve filtering sources and reopen after an affine history transaction', () => {
  const f = setup(); let reopened
  try {
    const before = json(f.book.save()), author = before.slicers.map(slicerDrawingAuthor), baseline = history(f)
    f.sheet.groupDrawings(slicerDrawingRefs, 'mixed')
    f.sheet.transformDrawings([{ kind: 'group', id: 'mixed' }], { delta })
    assert.equal(history(f), baseline + 2)
    const saved = json(f.book.save()), scene = host.createSheetDrawingScene(model(f), 'data')
    assert.deepEqual(saved.slicers.map(slicerDrawingAuthor), author)
    assert.deepEqual(saved.tables, before.tables); assert.deepEqual(saved.pivotTables, before.pivotTables)
    assert.deepEqual(saved.charts[0].source, before.charts[0].source)
    assert.deepEqual(scene.state.groups[0].members, slicerDrawingRefs)
    assert.notDeepEqual(saved.slicers[0].drawingGeometry.matrix, before.slicers[0].drawingGeometry.matrix)
    f.book.undo(); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
    f.book.redo(); f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
    reopened = new api.OfficeStudio({ workbook: saved })
    const reopenedBook = api.FOffice.newAPI(reopened).getActiveWorkbook()
    assert.deepEqual(json(host.createSheetDrawingScene(reopened.getWorkbookModel(f.unitId), 'data')), json(scene))
    assert.deepEqual(reopenedBook.getSlicer('s').getInfo().selectedItems, ['North'])
    assert.deepEqual(reopenedBook.getSlicer('t').getInfo().timelineRange, saved.slicers[1].timelineRange)
  } finally { reopened?.dispose(); f.dispose() }
})

test('native affine presentation changes one content entry and its expected revision is the workbook revision', async () => {
  const f = setup()
  try {
    const before = json(f.book.save()), revision = f.studio.getWorkbookRevision(f.unitId), baseline = history(f), events = []
    f.studio.events.addEvent('SlicerChanged', event => events.push({ ...event }))
    const slicer = f.book.getSlicer('s'), patch = presentation(f)
    assert.equal(await slicer.updatePresentation(patch, { expectedRevision: revision }), slicer)
    assert.equal(history(f), baseline + 1); assert.deepEqual(slicer.getInfo().drawingGeometry, patch.drawingGeometry)
    assert.deepEqual(slicerDrawingAuthor(slicer.getInfo()), slicerDrawingAuthor(before.slicers[0]))
    assert.deepEqual(events, [{ unitId: f.unitId, hostSheetId: 'data', slicerId: 's', revision: before.slicers[0].revision + 1 }])
    const after = json(f.book.save())
    f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), after)
    await assert.rejects(slicer.updatePresentation(presentation(f), { expectedRevision: revision }), /版本|内容|历史/)
    assert.equal(history(f), baseline + 1); assert.deepEqual(json(f.book.save()), after)
  } finally { f.dispose() }
})

for (const [name, expectedRevision] of [['negative', -1], ['fractional', 1.5], ['NaN', NaN], ['infinite', Infinity]]) {
  test(`native presentation rejects a ${name} expected revision before emitting cancellable events`, async () => {
    const f = setup()
    try {
      const before = json(model(f)), baseline = history(f), events = []
      f.studio.events.addEvent('BeforeSlicerChange', event => events.push(event))
      await assert.rejects(f.book.getSlicer('s').updatePresentation(presentation(f), { expectedRevision }), TypeError)
      assert.deepEqual(events, []); assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
    } finally { f.dispose() }
  })
}

test('native no-op presentation keeps slicer revision, model snapshot and content history unchanged', async () => {
  const f = setup()
  try {
    const before = json(model(f)), baseline = history(f), events = []
    f.studio.events.addEvent('SlicerChanged', event => events.push(event))
    await f.book.getSlicer('s').updatePresentation({ drawingGeometry: json(control(f).drawingGeometry), caption: control(f).caption })
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline); assert.deepEqual(events, [])
  } finally { f.dispose() }
})

for (const mode of ['native', 'shared']) test(`${mode} BeforeSlicerChange cancellation rejects the entire detached candidate`, async () => {
  const f = setup()
  try {
    const before = json(model(f)), baseline = history(f), changed = []
    f.studio.events.addEvent('BeforeSlicerChange', event => { if (event.slicerId === 's') event.cancel = true })
    f.studio.events.addEvent('SlicerChanged', event => changed.push(event))
    if (mode === 'shared') assert.throws(() => f.sheet.transformDrawings(slicerDrawingRefs, { delta }), /取消/)
    else await assert.rejects(change(f, mode), /cancel|取消/)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline); assert.deepEqual(changed, [])
  } finally { f.dispose() }
})

for (const mode of ['native', 'shared']) test(`${mode} reentry during BeforeSlicerChange preserves the inner change and rejects stale geometry`, async () => {
  const f = setup()
  try {
    const before = json(model(f)), baseline = history(f); let once = false
    f.studio.events.addEvent('BeforeSlicerChange', event => {
      if (event.slicerId !== 's' || once) return
      once = true
      f.studio.executeCommandSync('sheet.command.update-slicer-presentation', { unitId: f.unitId, slicerId: 's', patch: { caption: 'Inner author' } })
    })
    await assert.rejects(change(f, mode), /更改|stale|内容|模型/)
    assert.equal(control(f).caption, 'Inner author'); assert.deepEqual(json(control(f).drawingGeometry), before.slicers[0].drawingGeometry)
    assert.deepEqual(json(control(f, 't')), before.slicers[1]); assert.equal(history(f), baseline + 1)
  } finally { f.dispose() }
})

for (const mode of ['native', 'shared']) for (const [name, steer] of [
  ['user replacement', f => Object.defineProperty(f.studio, 'currentUserId', { value: 'peer' })],
  ['history replacement', f => f.studio.editor.history.clear(f.unitId)],
  ['in-place model edit', f => { control(f).caption = 'External author' }],
]) test(`${mode} rejects captured geometry after ${name} in a cancellable observer`, async () => {
  const f = setup()
  try {
    const geometry = json(control(f).drawingGeometry)
    f.studio.events.addEvent('BeforeSlicerChange', event => { if (event.slicerId === 's') steer(f) })
    await assert.rejects(change(f, mode), /更改|内容|历史|用户|模型/)
    assert.deepEqual(json(control(f).drawingGeometry), geometry); assert.equal(history(f), 0)
  } finally { f.dispose() }
})

for (const mode of ['native', 'shared']) for (const side of ['before', 'after']) for (const action of ['edit', 'view']) {
  test(`${mode} validates ${action} on the ${side} complete affine footprint beyond the anchor cell`, async () => {
    const value = slicerDrawingWorkbook(), s = value.slicers[0], geometry = s.drawingGeometry,
      matrix = side === 'after' ? { ...geometry.matrix, e: geometry.matrix.e + translate.e, f: geometry.matrix.f + translate.f } : geometry.matrix,
      bounds = drawingAffineBounds({ ...geometry, matrix }), footprint = sheetClipboardObjectRange(value, 'data', {
        left: bounds.left + GRID_HEADER_WIDTH, top: bounds.top + GRID_HEADER_HEIGHT, width: bounds.width, height: bounds.height })
    value.sheets[0].rangeProtectionRules = [{ id: 'corner', permissionId: 'corner', ownerUserId: 'peer', allowedUserIds: [],
      editState: 'onlyMe', viewState: action === 'view' ? 'noOneElseCanView' : 'othersCanView', description: '',
      ranges: [{ startRow: footprint.endRow, endRow: footprint.endRow, startCol: footprint.endCol, endCol: footprint.endCol }] }]
    const f = setup(value)
    try {
      const before = json(model(f)), baseline = history(f)
      await assert.rejects(change(f, mode), action === 'view' ? /查看/ : /权限|保护/)
      assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
    } finally { f.dispose() }
  })
}

test('shared settled events retain each committed slicer revision when the first listener edits another member', () => {
  const f = setup()
  try {
    const before = json(model(f)), events = []; let once = false
    f.studio.events.addEvent('SlicerChanged', event => {
      if (event.slicerId === 's' && !once) {
        once = true
        f.studio.executeCommandSync('sheet.command.update-slicer-presentation', { unitId: f.unitId, slicerId: 't', patch: { caption: 'Nested timeline' } })
      }
      event.revision = 9999
    })
    f.studio.events.addEvent('SlicerChanged', event => events.push({ ...event }))
    f.sheet.transformDrawings([leaf(f), leaf(f, 't')], { delta })
    const timelineEvents = events.filter(event => event.slicerId === 't')
    assert.deepEqual(timelineEvents.map(event => event.revision), [before.slicers[1].revision + 2, before.slicers[1].revision + 1])
    assert.equal(events.find(event => event.slicerId === 's').revision, before.slicers[0].revision + 1)
    assert.equal(control(f, 't').caption, 'Nested timeline'); assert.equal(history(f), 2)
  } finally { f.dispose() }
})

test('native settled events isolate payload edits and preserve the outer receipt after observer reentry', async () => {
  const f = setup()
  try {
    const revision = control(f).revision, events = []; let once = false
    f.studio.events.addEvent('SlicerChanged', event => {
      if (event.slicerId === 's' && !once) {
        once = true
        f.studio.executeCommandSync('sheet.command.update-slicer-presentation', { unitId: f.unitId, slicerId: 's', patch: { caption: 'Nested slicer' } })
      }
      event.revision = 9999
    })
    f.studio.events.addEvent('SlicerChanged', event => events.push({ ...event }))
    await change(f, 'native')
    assert.deepEqual(events.map(event => event.revision), [revision + 2, revision + 1])
    assert.equal(control(f).caption, 'Nested slicer'); assert.equal(history(f), 2)
  } finally { f.dispose() }
})

for (const mode of ['native', 'shared']) test(`${mode} keeps an accepted edit when a settled observer throws`, async () => {
  const f = setup()
  try {
    const revision = control(f).revision, errors = []
    f.studio.events.addEvent('SlicerChanged', () => { throw new Error('Observer failure') })
    f.studio.events.addEvent('EventListenerError', event => errors.push(event))
    await change(f, mode)
    assert.equal(control(f).revision, revision + 1); assert.equal(history(f), 1)
    assert.equal(errors.length, 1); assert.equal(errors[0].eventName, 'SlicerChanged')
  } finally { f.dispose() }
})

test('native arrangement and shared grouping use the same root order and restore typed selections through history', () => {
  const f = setup()
  try {
    broker.setSheetDrawingSelection(f.studio, f.unitId, 'data', [leaf(f), leaf(f, 't')])
    f.studio.executeCommandSync('sheet.command.update-slicer-presentation', { unitId: f.unitId, slicerId: 's', patch: { arrange: 'back' } })
    const ordered = json(model(f).sheets[0].drawingState.rootOrder)
    assert.deepEqual(ordered[0], leaf(f))
    f.sheet.groupDrawings([leaf(f), leaf(f, 't')], 'controls')
    assert.deepEqual(broker.getSheetDrawingSelection(f.studio, f.unitId, 'data'), [{ kind: 'group', id: 'controls' }])
    f.book.undo(); assert.deepEqual(model(f).sheets[0].drawingState.rootOrder, ordered)
    assert.deepEqual(broker.getSheetDrawingSelection(f.studio, f.unitId, 'data'), [leaf(f), leaf(f, 't')])
    f.book.redo(); assert.deepEqual(broker.getSheetDrawingSelection(f.studio, f.unitId, 'data'), [{ kind: 'group', id: 'controls' }])
  } finally { f.dispose() }
})

test('batch removal emits both native removal events and restores the original source bindings on undo', () => {
  const f = setup()
  try {
    const before = json(f.book.save()), removed = []
    f.studio.events.addEvent('SlicerRemoved', event => removed.push({ ...event }))
    f.sheet.groupDrawings(slicerDrawingRefs, 'mixed'); f.sheet.removeDrawings([{ kind: 'group', id: 'mixed' }])
    assert.deepEqual(removed.map(event => event.slicerId), ['s', 't']); assert.deepEqual(model(f).slicers, [])
    assert.deepEqual(json(model(f).tables), before.tables); assert.deepEqual(json(model(f).pivotTables), before.pivotTables)
    f.book.undo(); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})


const nativeContextMutations = [
  ['items', 's', (control, context) => control.setItems(['South'], context)],
  ['timeline range', 't', (control, context) => control.setTimelineRange({ start: '2026-01-01', end: '2026-01-31', level: 'month' }, context)],
  ['presentation', 's', (control, context) => control.updatePresentation({ caption: 'Context checked' }, context)],
  ['clear', 's', (control, context) => control.clearFilter(context)],
  ['remove', 's', (control, context) => control.remove(context)],
]
for (const [name, id, mutate] of nativeContextMutations) test(`${name} rejects a stale workbook revision before notification or mutation`, async () => {
  const f = setup()
  try {
    const revision = f.studio.getWorkbookRevision(f.unitId)
    await f.book.getSlicer('s').updatePresentation({ caption: 'A newer author' })
    const before = json(model(f)), baseline = history(f), events = []
    f.studio.events.addEvent('BeforeSlicerChange', event => events.push(event))
    await assert.rejects(mutate(f.book.getSlicer(id), { expectedRevision: revision }), /??|??|??/)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline); assert.deepEqual(events, [])
  } finally { f.dispose() }
})

for (const [name, id, mutate] of nativeContextMutations) test(`${name} workbench context rejects an inactive bound workbook`, async () => {
  const f = setup()
  try {
    const bound = f.book.getSlicer(id), other = f.facade.createWorkbook(slicerDrawingWorkbook({ id: 'other-owner' }))
    const before = json(model(f)), otherBefore = json(f.studio.getWorkbookModel(other.getId())), baseline = history(f), events = []
    f.studio.events.addEvent('BeforeSlicerChange', event => events.push(event))
    await assert.rejects(mutate(bound, { host: 'workbench', expectedRevision: f.studio.getWorkbookRevision(f.unitId) }), /???|??/)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline)
    assert.deepEqual(json(f.studio.getWorkbookModel(other.getId())), otherBefore); assert.deepEqual(events, [])
  } finally { f.dispose() }
})

for (const [name, id, mutate] of nativeContextMutations) test(`${name} default Facade context remains bound to its inactive unit`, async () => {
  const f = setup()
  try {
    const bound = f.book.getSlicer(id), other = f.facade.createWorkbook(slicerDrawingWorkbook({ id: 'other-owner' }))
    const before = json(model(f)), otherBefore = json(f.studio.getWorkbookModel(other.getId())), baseline = history(f)
    await mutate(bound, { expectedRevision: f.studio.getWorkbookRevision(f.unitId) })
    assert.notDeepEqual(json(model(f)), before); assert.equal(history(f), baseline + 1)
    assert.equal(f.studio.activeWorkbookId, other.getId()); assert.deepEqual(json(f.studio.getWorkbookModel(other.getId())), otherBefore)
    f.book.undo(); assert.deepEqual(json(model(f)), before)
  } finally { f.dispose() }
})

for (const [name, id, mutate] of nativeContextMutations) test(`${name} workbench context rejects an inactive host worksheet`, async () => {
  const f = setup()
  try {
    f.book.setActiveSheet('dates')
    const before = json(model(f)), baseline = history(f), events = []
    f.studio.events.addEvent('BeforeSlicerChange', event => events.push(event))
    await assert.rejects(mutate(f.book.getSlicer(id), { host: 'workbench', expectedRevision: f.studio.getWorkbookRevision(f.unitId) }), /???|??/)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline); assert.deepEqual(events, [])
  } finally { f.dispose() }
})

test('workbench ownership is checked again when BeforeSlicerChange switches the active workbook', async () => {
  const f = setup()
  try {
    const before = json(model(f)), baseline = history(f), settled = []
    f.studio.events.addEvent('BeforeSlicerChange', event => {
      if (event.unitId === f.unitId) f.facade.createWorkbook(slicerDrawingWorkbook({ id: 'reentrant-owner' }))
    })
    f.studio.events.addEvent('SlicerChanged', event => settled.push(event))
    await assert.rejects(f.book.getSlicer('s').updatePresentation({ caption: 'Stale owner' }, {
      host: 'workbench', expectedRevision: f.studio.getWorkbookRevision(f.unitId),
    }), /???|??|??|??/)
    assert.equal(f.studio.activeWorkbookId, 'reentrant-owner'); assert.deepEqual(json(model(f)), before)
    assert.equal(history(f), baseline); assert.deepEqual(settled, [])
  } finally { f.dispose() }
})

test('an unknown native interaction host is rejected before cancellable observers run', async () => {
  const f = setup()
  try {
    const before = json(model(f)), baseline = history(f), events = []
    f.studio.events.addEvent('BeforeSlicerChange', event => events.push(event))
    await assert.rejects(f.book.getSlicer('s').updatePresentation({ caption: 'Invalid host' }, { host: 'other-host' }), TypeError)
    assert.deepEqual(json(model(f)), before); assert.equal(history(f), baseline); assert.deepEqual(events, [])
  } finally { f.dispose() }
})
