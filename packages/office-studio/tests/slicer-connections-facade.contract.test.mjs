import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Deferred real registry/Facade/history integration drafts; none have been run.
let server, api, fixture
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture] = await Promise.all([server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/slicer-connections-fixture.mjs')])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const history = f => f.studio.editor.history.units[f.unitId]?.undo.length ?? 0
const model = f => f.studio.getWorkbookModel(f.unitId)
const targets = () => ({ pivotTables: ['primary', 'secondary'].map(key => ({ subUnitId: key === 'primary' ? fixture.connectionFixtureIds.host : fixture.connectionFixtureIds.other,
  pivotTableId: fixture.connectionFixtureIds[key] })), ownedPivotCharts: [{ subUnitId: fixture.connectionFixtureIds.host, pivotChartId: fixture.connectionFixtureIds.owned }] })
function setup(options) {
  const studio = new api.OfficeStudio({ workbook: fixture.slicerConnectionsWorkbook(options) }), facade = api.FOffice.newAPI(studio), book = facade.getActiveWorkbook()
  return { studio, facade, book, control: book.getSlicer(fixture.connectionFixtureIds.slicer), shared: book.getSlicer(fixture.connectionFixtureIds.shared), unitId: book.getId(), dispose: () => studio.dispose() }
}

test('Facade replaces shared connections in one history entry, and undo/redo/save/reopen retain source and author geometry', async () => {
  const f = setup(); let reopened
  try {
    const before = json(f.book.save()), baseline = history(f)
    assert.equal(await f.control.setConnections(targets()), true)
    assert.equal(history(f), baseline + 1); assert.deepEqual(f.shared.getConnections(), targets())
    const after = json(f.book.save())
    assert.deepEqual(after.slicers.map(view => view.drawingGeometry), before.slicers.map(view => view.drawingGeometry))
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
    f.book.redo(); assert.deepEqual(json(f.book.save()), after)
    reopened = new api.OfficeStudio({ workbook: after })
    assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getSlicer(fixture.connectionFixtureIds.slicer).getConnections(), targets())
  } finally { reopened?.dispose(); f.dispose() }
})

test('a satisfied replacement is accepted without content revision, history or settled change events', async () => {
  const f = setup()
  try {
    const snapshot = json(model(f)), revision = f.studio.getWorkbookRevision(f.unitId), baseline = history(f), changes = []
    f.studio.events.addEvent('SlicerChanged', event => changes.push(event))
    assert.equal(await f.control.setConnections(f.control.getConnections()), true)
    assert.deepEqual(json(model(f)), snapshot); assert.equal(f.studio.getWorkbookRevision(f.unitId), revision); assert.equal(history(f), baseline); assert.deepEqual(changes, [])
  } finally { f.dispose() }
})

test('ordinary incompatible targets and Table controls return false without a partial accepted connection', async () => {
  const f = setup()
  try {
    const snapshot = json(model(f)), baseline = history(f), requested = targets()
    requested.pivotTables.push({ subUnitId: fixture.connectionFixtureIds.host, pivotTableId: fixture.connectionFixtureIds.incompatible })
    assert.equal(await f.control.setConnections(requested), false)
    assert.equal(await f.book.getSlicer(fixture.connectionFixtureIds.table).setConnections(targets()), false)
    assert.deepEqual(json(model(f)), snapshot); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('read-only connection edits return false before any cancellable slicer event', async () => {
  const f = setup({ readonly: true })
  try {
    const snapshot = json(model(f)), events = [], baseline = history(f)
    f.studio.events.addEvent('BeforeSlicerChange', event => events.push(event))
    assert.equal(await f.control.setConnections(targets()), false)
    assert.deepEqual(events, []); assert.deepEqual(json(model(f)), snapshot); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('an ordinary connected-report output conflict returns false and leaves every report and history unchanged', async () => {
  const f = setup()
  try {
    await f.studio.executeCommand('sheet.command.set-pivot-label-filter', { unitId: f.unitId, pivotTableId: fixture.connectionFixtureIds.secondary,
      fieldId: '0', selectedItems: ['North'] })
    f.book.getSheetBySheetId(fixture.connectionFixtureIds.host).getRange('I20').setValue('occupied')
    const snapshot = json(model(f)), baseline = history(f)
    assert.equal(await f.control.setConnections(targets()), false)
    assert.deepEqual(json(model(f)), snapshot); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

for (const kind of ['slicer', 'command']) test(`unexpected ${kind} listener failures reject without being converted to an ordinary refusal`, async () => {
  const f = setup()
  try {
    const snapshot = json(model(f)), baseline = history(f), failure = new Error('unexpected-listener')
    f.studio.events.addEvent(kind === 'slicer' ? 'BeforeSlicerChange' : 'BeforeCommandExecute', () => { throw failure })
    await assert.rejects(f.control.setConnections(targets()), error => error === failure)
    assert.deepEqual(json(model(f)), snapshot); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('cancellation rejects the entire candidate before any report or shared view changes', async () => {
  const f = setup()
  try {
    const snapshot = json(model(f)), baseline = history(f)
    f.studio.events.addEvent('BeforeSlicerChange', event => { event.cancel = true })
    await assert.rejects(f.control.setConnections(targets()), /canceled|取消/i)
    assert.deepEqual(json(model(f)), snapshot); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('settled notifications cover every shared view and connected table and isolate observer payload mutation', async () => {
  const f = setup()
  try {
    await f.control.setConnections(targets())
    const controls = [], pivots = [], failures = []
    f.studio.events.addEvent('SlicerChanged', event => { event.slicerId = 'corrupted'; throw new Error('settled-observer') })
    f.studio.events.addEvent('SlicerChanged', event => controls.push({ ...event }))
    f.studio.events.addEvent('PivotTableRendered', event => pivots.push(json(event)))
    f.studio.events.addEvent('EventListenerError', event => failures.push(event))
    await f.control.setItems(['North'])
    assert.deepEqual(new Set(controls.map(event => event.slicerId)), new Set([fixture.connectionFixtureIds.slicer, fixture.connectionFixtureIds.shared]))
    assert.deepEqual(new Set(pivots.map(event => event.pivotTableId)), new Set([fixture.connectionFixtureIds.primary, fixture.connectionFixtureIds.secondary]))
    assert.ok(failures.length >= 2)
    for (const key of ['primary', 'secondary', 'owned']) assert.deepEqual(fixture.connectionReportSelection(model(f), fixture.connectionFixtureIds[key]), ['North'])
  } finally { f.dispose() }
})

test('a stale expected content revision rejects replacement and retains history and geometry', async () => {
  const f = setup()
  try {
    const revision = f.studio.getWorkbookRevision(f.unitId)
    f.book.setName('Changed elsewhere')
    const snapshot = json(model(f)), baseline = history(f)
    await assert.rejects(f.control.setConnections(targets(), { expectedRevision: revision, host: 'workbench' }), /更改|版本/)
    assert.deepEqual(json(model(f)), snapshot); assert.equal(history(f), baseline)
  } finally { f.dispose() }
})

test('a bound Facade edits its inactive workbook while workbench context refuses the changed host', async () => {
  const f = setup()
  try {
    const other = f.studio.createWorkbook({ name: 'Other workbook' }, true)
    assert.equal(await f.control.setConnections(targets()), true)
    const snapshot = json(model(f)), baseline = history(f)
    await assert.rejects(f.control.setConnections({ pivotTables: [] }, { expectedRevision: f.studio.getWorkbookRevision(f.unitId), host: 'workbench' }), /宿主/)
    assert.deepEqual(json(model(f)), snapshot); assert.equal(history(f), baseline)
    assert.equal(f.studio.getWorkbookModel(other.id).name, 'Other workbook')
  } finally { f.dispose() }
})

test('an external connected Pivot field command updates the shared cache and every connected report in one history entry', async () => {
  const f = setup()
  try {
    await f.control.setConnections(targets())
    const before = json(f.book.save()), baseline = history(f)
    await f.studio.executeCommand('sheet.command.set-pivot-label-filter', { unitId: f.unitId, pivotTableId: fixture.connectionFixtureIds.secondary, fieldId: '0', selectedItems: ['South'] })
    assert.equal(history(f), baseline + 1); assert.deepEqual(f.control.getInfo().selectedItems, ['South']); assert.deepEqual(f.shared.getInfo().selectedItems, ['South'])
    for (const key of ['primary', 'secondary', 'owned']) assert.deepEqual(fixture.connectionReportSelection(model(f), fixture.connectionFixtureIds[key]), ['South'])
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test('deleting the original report retains surviving connections and live filtering; undo restores every report', async () => {
  const f = setup()
  try {
    await f.control.setConnections(targets())
    const before = json(f.book.save())
    await f.studio.executeCommand('sheet.command.remove-pivot-table', { unitId: f.unitId, pivotTableId: fixture.connectionFixtureIds.primary })
    assert.equal(f.control.getConnections().pivotTables.length, 1)
    assert.equal(f.control.getConnections().pivotTables[0].pivotTableId, fixture.connectionFixtureIds.secondary)
    await f.control.setItems(['North'])
    assert.deepEqual(fixture.connectionReportSelection(model(f), fixture.connectionFixtureIds.owned), ['North'])
    f.book.undo(); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})
