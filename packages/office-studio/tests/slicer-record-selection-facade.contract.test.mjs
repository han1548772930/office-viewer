import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Real command/history/cancellation/save/reopen drafts. This Vite setup and every test remain unexecuted.
let server, api, fixture, nativeFilters
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture, nativeFilters] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-filter-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const authored = book => { const snapshot = json(book.save()); delete snapshot.modified; return snapshot }
function setup(t, kind = 'pivot', condition = false) {
  const workbook = kind === 'table' ? fixture.slicerTableRecordSelectionWorkbook() : fixture.slicerRecordSelectionWorkbook()
  if (condition) {
    const table = workbook.tables.find(table => table.id === 'typed-table'), source = workbook.sheets.find(sheet => sheet.id === 'data')
    const sheet = { ...source, autoFilter: undefined }
    nativeFilters.setAutoFilter(sheet, table.range)
    nativeFilters.setFilterCriteria(sheet, 6, { kind: 'condition', operator: 'equals', value: 1 })
    table.autoFilter = sheet.autoFilter
  }
  const studio = new api.OfficeStudio({ workbook }); t.onTestFinished(() => studio.dispose())
  const facade = api.FOffice.newAPI(studio), book = facade.getActiveWorkbook(), unitId = book.getId()
  const id = kind === 'table' ? 'typed-table-slicer' : 'typed-slicer'
  return { studio, facade, book, unitId, control: book.getSlicerById(id), shared: book.getSlicerById(`${id}-copy`),
    history: () => studio.editor.history.units[unitId]?.undo.length ?? 0, model: () => studio.getWorkbookModel(unitId) }
}
const sum = (workbook, report) => api.calculatePivotReport(workbook, report).values.reduce((total, row) =>
  total + (typeof row.at(-1) === 'number' ? row.at(-1) : 0), 0)

for (const kind of ['pivot', 'table']) test(`${kind} typed selection and partial toggle each commit once and survive undo/redo/reopen`, async t => {
  const f = setup(t, kind), before = authored(f.book), baseline = f.history()
  assert.equal(await f.control.setRecordSelection({ mode: 'include', values: [1, true, null] }), true)
  assert.equal(f.history(), baseline + 1)
  assert.deepEqual(f.control.getRecordSelection(), { mode: 'include', values: [1, true, null] })
  assert.deepEqual(f.shared.getItemState(), { checkedItems: [], partialItems: ['1', 'true', ''] })
  if (kind === 'pivot') for (const report of fixture.filterReports(f.model())) assert.equal(sum(f.model(), report), 90)
  const first = authored(f.book)
  assert.equal(await f.control.setItemStates(['beta'], ['1', 'true', '']), true); assert.equal(f.history(), baseline + 2)
  assert.deepEqual(f.control.getRecordSelection(), { mode: 'include', values: [1, true, null, 'beta'] })
  const after = authored(f.book)
  f.book.undo(); assert.deepEqual(authored(f.book), first)
  f.book.undo(); assert.deepEqual(authored(f.book), before)
  f.book.redo(); f.book.redo(); assert.deepEqual(authored(f.book), after)
  const reopened = new api.OfficeStudio({ workbook: f.book.save() }); t.onTestFinished(() => reopened.dispose())
  const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getSlicerById(f.control.getId())
  assert.deepEqual(restored.getRecordSelection(), { mode: 'include', values: [1, true, null, 'beta'] })
  assert.deepEqual(restored.getItemState(), { checkedItems: ['beta'], partialItems: ['1', 'true', ''] })
})

for (const kind of ['pivot', 'table']) test(`${kind} typed raw reorder is accepted without content history or model revisions`, async t => {
  const f = setup(t, kind)
  await f.control.setRecordSelection({ mode: 'include', values: [1, true, null] })
  const before = authored(f.book), baseline = f.history(), snapshot = f.control.getRecordSelection()
  snapshot.values.push('1'); assert.deepEqual(f.control.getRecordSelection(), { mode: 'include', values: [1, true, null] })
  assert.equal(await f.control.setRecordSelection({ mode: 'include', values: [null, 1, true, 1] }), true)
  assert.deepEqual(authored(f.book), before); assert.equal(f.history(), baseline)
})

for (const kind of ['pivot', 'table']) for (const action of ['record', 'partial', 'all']) {
  test(`${kind} canceled ${action} typed edit publishes no partial cache, native filter, report or history`, async t => {
    const f = setup(t, kind); await f.control.setRecordSelection({ mode: 'include', values: [1] })
    const before = authored(f.book), baseline = f.history()
    const listener = f.studio.events.addEvent('BeforeSlicerChange', event => { event.cancel = true }); t.onTestFinished(() => listener.dispose())
    const pending = action === 'record' ? f.control.setRecordSelection({ mode: 'include', values: ['1'] })
      : action === 'partial' ? f.control.setItemStates(['beta'], ['1']) : f.control.setRecordSelection({ mode: 'all' })
    await assert.rejects(pending, /canceled|取消/i)
    assert.deepEqual(authored(f.book), before); assert.equal(f.history(), baseline)
  })
}

test('typed Pivot filtering notifies both linked/owned charts and all shared views after one commit', async t => {
  const f = setup(t), charts = [], controls = []
  const chartListener = f.studio.events.addEvent('PivotChartChanged', event => charts.push(event.pivotChartId))
  const slicerListener = f.studio.events.addEvent('SlicerChanged', event => controls.push(event.slicerId))
  t.onTestFinished(() => chartListener.dispose()); t.onTestFinished(() => slicerListener.dispose())
  const baseline = f.history()
  assert.equal(await f.control.setRecordSelection({ mode: 'include', values: [1] }), true)
  assert.equal(f.history(), baseline + 1)
  assert.deepEqual(new Set(charts), new Set(['linked-chart', 'owned-chart']))
  assert.deepEqual(new Set(controls), new Set(['typed-slicer', 'typed-slicer-copy']))
  for (const report of fixture.filterReports(f.model())) assert.equal(sum(f.model(), report), 10)
})

test('zero-connection raw identity is saved and reconnect replaces targets without widening its caption', async t => {
  const f = setup(t); await f.control.setRecordSelection({ mode: 'include', values: [1] })
  await f.control.setConnections({ pivotTables: [] })
  await f.control.setRecordSelection({ mode: 'include', values: ['1'] })
  assert.deepEqual(f.control.getRecordSelection(), { mode: 'include', values: ['1'] })
  const saved = f.book.save(), reopened = new api.OfficeStudio({ workbook: saved }); t.onTestFinished(() => reopened.dispose())
  const restoredBook = api.FOffice.newAPI(reopened).getActiveWorkbook(), restored = restoredBook.getSlicerById('typed-slicer')
  assert.equal(await restored.setConnections({ pivotTables: [{ subUnitId: 'output', pivotTableId: 'secondary-report' }],
    ownedPivotCharts: [{ subUnitId: 'dashboard', pivotChartId: 'owned-chart' }] }), true)
  const model = reopened.getWorkbookModel(f.unitId)
  assert.equal(sum(model, model.pivotTables[0]), 10); assert.equal(sum(model, model.pivotTables[1]), 20)
  assert.equal(sum(model, model.charts.find(chart => chart.ownedPivot).ownedPivot), 20)
})

test('native condition partial UI reads remain available and a Canvas edit explicitly replaces that column', async t => {
  const f = setup(t, 'table', true), before = authored(f.book), baseline = f.history()
  assert.throws(() => f.control.getRecordSelection(), /原生条件筛选/)
  assert.deepEqual(f.control.getItemState(), { checkedItems: [], partialItems: ['1'] })
  assert.deepEqual(f.control.getInfo().selectedItems, ['1']); assert.deepEqual(authored(f.book), before)
  assert.equal(await f.control.setItemStates(['beta'], ['1']), true); assert.equal(f.history(), baseline + 1)
  assert.deepEqual(f.control.getRecordSelection(), { mode: 'include', values: [1, 'beta'] })
  f.book.undo(); assert.deepEqual(authored(f.book), before)
})
