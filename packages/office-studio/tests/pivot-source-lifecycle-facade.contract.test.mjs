import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { resolve } from 'node:path'

// Real command/history/save/reopen drafts. Deferred: this file and its Vite setup have not been run.
let server, api, fixture, lifecycle
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture, lifecycle] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-source-lifecycle-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-source-lifecycle.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function setup(t, options = {}) {
  const studio = new api.OfficeStudio({ workbook: fixture.pivotSourceLifecycleWorkbook(options) })
  t.onTestFinished(() => studio.dispose())
  const facade = api.FOffice.newAPI(studio), book = facade.getActiveWorkbook(), unitId = book.getId()
  return { studio, api: facade, book, unitId, model: () => studio.getWorkbookModel(unitId),
    history: () => studio.editor.history.units[unitId]?.undo.length ?? 0,
    command: (id, input) => facade.executeCommand(id, { unitId, subUnitId: 'data', ...input }) }
}
const sum = chart => chart.getData().series[0].values.reduce((total, value) => total + Number(value ?? 0), 0)

for (const [command, input, total] of [
  ['sheet.command.insert-col', { start: 5, count: 1 }, 100],
  ['sheet.command.remove-col', { start: 3, count: 1 }, 100],
  ['sheet.command.copy-rows', { start: 3, count: 1, destinationIndex: 4 }, 110],
]) test(`${command} commits cache, owned report, linked output, undo/redo and saved data as one content history`, async t => {
  const f = setup(t), before = json(f.book.save()), baseline = f.history()
  assert.equal(await f.command(command, input), true); assert.equal(f.history(), baseline + 1)
  assert.equal(f.model().pivotCaches.find(cache => cache.id === 'source-cache').revision, 1)
  assert.ok(fixture.allLifecycleReports(f.model()).every(report => report.revision === 1))
  for (const id of ['linked-a', 'linked-b', 'owned-chart']) assert.equal(sum(f.book.getPivotChartById(id)), total)
  const changed = json(f.book.save())
  f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  f.book.redo(); assert.deepEqual(json(f.book.save()), changed)
  const reopened = new api.OfficeStudio({ workbook: changed }); t.onTestFinished(() => reopened.dispose())
  const reopenedBook = api.FOffice.newAPI(reopened).getActiveWorkbook()
  assert.deepEqual(json(reopenedBook.save()), changed)
  assert.equal(sum(reopenedBook.getPivotChartById('owned-chart')), total)
})

test('the registered copy-sheet command forks a self-contained source once and undoes every copied resource together', async t => {
  const f = setup(t, { selfContained: true }), before = json(f.book.save()), baseline = f.history()
  assert.equal(await f.command('sheet.command.copy-sheet', { targetSheetId: 'copy', targetSheetName: 'Source copy' }), true)
  assert.equal(f.history(), baseline + 1)
  const copied = f.model().pivotTables.filter(report => report.anchor.subUnitId === 'copy')
  const owned = f.model().charts.find(chart => chart.sheetId === 'copy' && chart.ownedPivot).ownedPivot
  assert.equal(copied.length, 2); assert.notEqual(copied[0].pivotCacheId, 'source-cache'); assert.equal(owned.pivotCacheId, copied[0].pivotCacheId)
  assert.equal(owned.source.subUnitId, 'copy'); assert.equal(owned.axisOutput, undefined)
  assert.ok(copied.every(report => report.collapseInfo['region-local'].North === true))
  const changed = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  f.book.redo(); assert.deepEqual(json(f.book.save()), changed)
})

test('source rename through the canonical command preserves hierarchy values and reopens cache identities', async t => {
  const f = setup(t), before = json(f.book.save()), baseline = f.history(), events = []
  const listener = f.api.addEvent(f.api.Event.SheetNameChanged, event => events.push([event.sheetId, event.oldName, event.newName]))
  t.onTestFinished(() => listener.dispose())
  assert.equal(await f.command('sheet.command.set-worksheet-name', { name: 'Source 2026' }), true)
  assert.deepEqual(events, [['data', 'Data', 'Source 2026']]); assert.equal(f.history(), baseline + 1)
  assert.ok(fixture.allLifecycleReports(f.model()).every(report => report.source.subUnitId === 'data' && report.source.sheetName === 'Source 2026'))
  assert.ok(f.model().charts.every(chart => chart.revision === 1)); assert.equal(sum(f.book.getPivotChartById('owned-chart')), 100)
  const changed = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  f.book.redo(); assert.deepEqual(json(f.book.save()), changed)
  const reopened = new api.OfficeStudio({ workbook: changed }); t.onTestFinished(() => reopened.dispose())
  assert.equal(reopened.getWorkbookModel(f.unitId).pivotCaches[0].source.subUnitId, 'data')
  assert.equal(reopened.getWorkbookModel(f.unitId).pivotCaches[0].source.sheetName, 'Source 2026')
})

for (const command of ['sheet.command.remove-sheet', 'sheet.command.delete-sheet']) test(`${command} retains unavailable source identity and history restores the complete live records`, async t => {
  const f = setup(t), before = json(f.book.save()), baseline = f.history()
  assert.equal(await f.command(command, {}), true); assert.equal(f.history(), baseline + 1)
  assert.equal(f.model().charts.length, 3); assert.equal(f.model().pivotCaches[0].source.subUnitId, 'data')
  assert.ok(fixture.allLifecycleReports(f.model()).every(report => lifecycle.getPivotSourceAvailability(f.model(), report) === 'worksheet-missing' && report.axisOutput === undefined))
  const changed = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  assert.equal(sum(f.book.getPivotChartById('owned-chart')), 100)
  f.book.redo(); assert.deepEqual(json(f.book.save()), changed)
  const reopened = new api.OfficeStudio({ workbook: changed }); t.onTestFinished(() => reopened.dispose())
  const model = reopened.getWorkbookModel(f.unitId)
  assert.ok(fixture.allLifecycleReports(model).every(report => lifecycle.getPivotSourceAvailability(model, report) === 'worksheet-missing'))
})

test('a same-name replacement created through the shared worksheet command cannot revive the deleted source reference', async t => {
  const f = setup(t)
  await f.command('sheet.command.remove-sheet', {})
  assert.equal(await f.command('sheet.command.insert-sheet', { name: 'Data', sheet: { id: 'replacement-data' } }), true)
  assert.ok(fixture.allLifecycleReports(f.model()).every(report => report.source.subUnitId === 'data'))
  assert.equal(f.model().pivotCaches[0].source.subUnitId, 'data')
  assert.ok(fixture.allLifecycleReports(f.model()).every(report => lifecycle.getPivotSourceAvailability(f.model(), report) === 'worksheet-missing'))
})

test('canceled source rename leaves staged cache/report changes outside the document history', async t => {
  const f = setup(t), before = json(f.book.save()), baseline = f.history()
  const listener = f.api.addEvent(f.api.Event.BeforeSheetNameChange, event => { event.cancel = true })
  t.onTestFinished(() => listener.dispose())
  assert.equal(await f.command('sheet.command.set-worksheet-name', { name: 'Canceled source' }), false)
  assert.deepEqual(json(f.book.save()), before); assert.equal(f.history(), baseline)
})

test('reentrant source edits invalidate an obsolete rename candidate and keep only the real structure transaction', async t => {
  const f = setup(t), baseline = f.history()
  const listener = f.api.addEvent(f.api.Event.BeforeSheetNameChange, () => {
    f.api.executeCommandSync('sheet.command.insert-col', { unitId: f.unitId, subUnitId: 'data', start: 5, count: 1 })
  })
  t.onTestFinished(() => listener.dispose())
  await assert.rejects(f.command('sheet.command.set-worksheet-name', { name: 'Obsolete source' }), /上下文已更改/)
  assert.equal(f.history(), baseline + 1); assert.equal(f.model().sheets.find(sheet => sheet.id === 'data').name, 'Data')
  assert.equal(f.model().pivotTables[0].fields.find(field => field.id === 'amount-local').sourceColumn, 6)
  assert.equal(sum(f.book.getPivotChartById('owned-chart')), 100)
})

test('no-op source moves and stale structure commands do not add phantom source/cache history', async t => {
  const f = setup(t), before = json(f.book.save()), baseline = f.history(), revision = f.studio.getWorkbookRevision(f.unitId)
  assert.equal(await f.command('sheet.command.move-rows', { start: 3, count: 1, destinationIndex: 3 }), false)
  assert.deepEqual(json(f.book.save()), before); assert.equal(f.history(), baseline)
  await f.command('sheet.command.insert-col', { start: 5, count: 1 })
  const changed = json(f.book.save())
  await assert.rejects(f.command('sheet.command.remove-col', { start: 3, count: 1, expectedRevision: revision }), /更改|版本/)
  assert.deepEqual(json(f.book.save()), changed); assert.equal(f.history(), baseline + 1)
})
