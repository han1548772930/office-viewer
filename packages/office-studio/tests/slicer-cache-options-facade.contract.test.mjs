import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Real command/history/cancellation drafts only. This setup and every test remain unexecuted.
let server, api, fixture
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture] = await Promise.all([server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-filter-fixture.mjs')])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value)), kinds = ['pivot', 'table']
const authored = book => { const snapshot = json(book.save()); delete snapshot.modified; return snapshot }
function setup(t, kind) {
  const workbook = kind === 'table' ? fixture.slicerTableRecordSelectionWorkbook() : fixture.slicerRecordSelectionWorkbook()
  const studio = new api.OfficeStudio({ workbook }); t.onTestFinished(() => studio.dispose())
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), unitId = book.getId(), id = workbook.slicers[0].id
  const control = book.getSlicerById(id), shared = book.getSlicerById(workbook.slicers[1].id)
  return { studio, book, control, shared, unitId, model: () => studio.getWorkbookModel(unitId),
    history: () => studio.editor.history.units[unitId]?.undo.length ?? 0,
    cache: () => studio.getWorkbookModel(unitId).slicerCaches.find(cache => cache.id === control.getCacheId()) }
}

for (const kind of kinds) for (const action of ['author', 'clear', 'mixed']) {
  test(`${kind} canceled ${action} cache options publish no cache/view revisions, content history or notifications`, async t => {
    const f = setup(t, kind); await f.control.setRecordSelection({ mode: 'include', values: [1] })
    await f.control.setCacheOptions({ sortOrder: 'ascending', customListSort: true })
    const before = authored(f.book), baseline = f.history(), notifications = []
    const cancel = f.studio.events.addEvent('BeforeSlicerChange', event => { event.cancel = true })
    const changed = f.studio.events.addEvent('SlicerChanged', event => notifications.push(event.slicerId))
    t.onTestFinished(() => cancel.dispose()); t.onTestFinished(() => changed.dispose())
    const patch = action === 'author' ? { sortOrder: 'descending', crossFilter: 'showItemsWithDataAtTop', hideItemsWithNoData: true }
      : action === 'clear' ? { sortOrder: undefined, customListSort: undefined }
        : { sortOrder: undefined, customListSort: false, crossFilter: 'none', hideItemsWithNoData: true }
    await assert.rejects(f.control.setCacheOptions(patch), /canceled|取消/i)
    assert.deepEqual(authored(f.book), before); assert.equal(f.history(), baseline); assert.deepEqual(notifications, [])
    assert.deepEqual(f.shared.getCacheOptions(), { sortOrder: 'ascending', customListSort: true })
  })
}

for (const kind of kinds) test(`${kind} accepted options change commits once, notifies shared views and preserves all filter authorities`, async t => {
  const f = setup(t, kind); await f.control.setRecordSelection({ mode: 'include', values: [1] })
  const before = authored(f.book), baseline = f.history(), cacheRevision = f.cache().revision, notifications = [], chartEvents = []
  const changed = f.studio.events.addEvent('SlicerChanged', event => notifications.push(event.slicerId))
  const chartChanged = f.studio.events.addEvent('PivotChartChanged', event => chartEvents.push(event.pivotChartId))
  t.onTestFinished(() => changed.dispose()); t.onTestFinished(() => chartChanged.dispose())
  const options = { sortOrder: 'descending', customListSort: false, crossFilter: 'showItemsWithNoData', hideItemsWithNoData: true }
  assert.equal(await f.control.setCacheOptions(options), true); assert.equal(f.history(), baseline + 1)
  assert.equal(f.cache().revision, cacheRevision + 1)
  assert.deepEqual(new Set(notifications), new Set([f.control.getId(), f.shared.getId()])); assert.deepEqual(chartEvents, [])
  assert.deepEqual(f.shared.getRecordSelection(), { mode: 'include', values: [1] })
  const after = authored(f.book)
  assert.deepEqual(after.pivotTables, before.pivotTables); assert.deepEqual(after.charts, before.charts); assert.deepEqual(after.tables, before.tables)
  const snapshot = f.shared.getCacheOptions(); snapshot.sortOrder = 'ascending'; assert.equal(f.control.getCacheOptions().sortOrder, 'descending')
  assert.equal(await f.shared.setCacheOptions({ hideItemsWithNoData: true, crossFilter: 'showItemsWithNoData', customListSort: false, sortOrder: 'descending' }), true)
  assert.equal(f.history(), baseline + 1); assert.deepEqual(authored(f.book), after)
  f.book.undo(); assert.deepEqual(authored(f.book), before)
  f.book.redo(); assert.deepEqual(authored(f.book), after)
  const reopened = new api.OfficeStudio({ workbook: f.book.save() }); t.onTestFinished(() => reopened.dispose())
  const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getSlicerById(f.control.getId())
  assert.deepEqual(restored.getCacheOptions(), options); assert.deepEqual(restored.getRecordSelection(), { mode: 'include', values: [1] })
})

for (const kind of kinds) test(`${kind} own undefined options survive command presence handling and undo/redo as optional author state`, async t => {
  const f = setup(t, kind)
  await f.control.setCacheOptions({ sortOrder: 'descending', customListSort: true, crossFilter: 'none', hideItemsWithNoData: false })
  const before = authored(f.book), baseline = f.history()
  assert.equal(await f.shared.setCacheOptions({ sortOrder: undefined, customListSort: undefined }), true)
  assert.equal(f.history(), baseline + 1)
  assert.deepEqual(f.control.getCacheOptions(), { crossFilter: 'none', hideItemsWithNoData: false })
  const after = authored(f.book)
  f.book.undo(); assert.deepEqual(authored(f.book), before)
  f.book.redo(); assert.deepEqual(authored(f.book), after)
  assert.equal(await f.control.setCacheOptions({ sortOrder: undefined, customListSort: undefined }), true)
  assert.equal(f.history(), baseline + 1); assert.deepEqual(authored(f.book), after)
})

for (const kind of kinds) test(`${kind} a canceled options edit leaves an in-flight typed window valid and correctly ordered`, async t => {
  const f = setup(t, kind); await f.control.setCacheOptions({ sortOrder: 'ascending' })
  const expected = await f.control.getItems({ start: 0, count: 100 }), before = authored(f.book), baseline = f.history()
  const pending = f.shared.getItems({ start: 0, count: 100 })
  const cancel = f.studio.events.addEvent('BeforeSlicerChange', event => { event.cancel = true }); t.onTestFinished(() => cancel.dispose())
  await assert.rejects(f.control.setCacheOptions({ sortOrder: 'descending' }), /canceled|取消/i)
  assert.deepEqual(await pending, expected); assert.deepEqual(authored(f.book), before); assert.equal(f.history(), baseline)
})

for (const kind of kinds) test(`${kind} a throwing Before cache-options listener aborts without any shared revision or history`, async t => {
  const f = setup(t, kind), before = authored(f.book), baseline = f.history()
  const listener = f.studio.events.addEvent('BeforeSlicerChange', () => { throw new Error('options listener failed') })
  t.onTestFinished(() => listener.dispose())
  await assert.rejects(f.control.setCacheOptions({ customListSort: true }), /options listener failed/)
  assert.deepEqual(authored(f.book), before); assert.equal(f.history(), baseline); assert.deepEqual(f.shared.getCacheOptions(), {})
})
