import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Real registry/Facade/report/history drafts. All execution is deferred.
let server, api, fixture, selection
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture, selection] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/slicer-connections-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-control-selection.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const total = chart => (chart.getData().series[0]?.values ?? []).reduce((sum, value) => sum + value, 0)
function authored(book) { const value = json(book.saveModel()); delete value.modified; return value }
function setup() {
  const workbook = fixture.slicerConnectionsWorkbook({ affine: false }), ids = fixture.connectionFixtureIds
  for (const [index, value] of [1, '1', true, 'true'].entries()) {
    workbook.sheets.find(sheet => sheet.id === ids.source).cells[`${index + 1}:0`] = api.createCell(index + 1, 0, value)
    workbook.sheets.find(sheet => sheet.id === ids.host).cells[`${index + 1}:8`] = api.createCell(index + 1, 8, value)
  }
  for (const report of workbook.pivotTables) api.refreshPivotTable(workbook, report)
  api.refreshPivotTable(workbook, workbook.charts.find(chart => chart.id === ids.owned).ownedPivot, 'data-only')
  const studio = new api.OfficeStudio({ workbook }), book = api.FOffice.newAPI(studio).getActiveWorkbook()
  return { studio, book, ids, unitId: book.getId(), control: book.getSlicer(ids.slicer), shared: book.getSlicer(ids.shared),
    model: () => studio.getWorkbookModel(book.getId()), history: () => studio.editor.history.units[book.getId()]?.undo.length ?? 0,
    dispose: () => studio.dispose() }
}
async function connect(f) {
  assert.equal(await f.control.setConnections({ pivotTables: [
    { subUnitId: f.ids.host, pivotTableId: f.ids.primary }, { subUnitId: f.ids.other, pivotTableId: f.ids.secondary },
  ], ownedPivotCharts: [{ subUnitId: f.ids.host, pivotChartId: f.ids.owned }] }), true)
}

test('typed selection shares partial state across reports and views in one undo entry, save and reopen', async () => {
  const f = setup(); let reopened
  try {
    await connect(f)
    const before = authored(f.book), source = json(f.model().sheets[0].cells), baseline = f.history(), events = []
    f.studio.events.addEvent('PivotChartChanged', event => events.push({ ...event }))
    assert.equal(await f.control.setRecordSelection({ mode: 'include', values: [1, true] }), true)
    assert.equal(f.history(), baseline + 1)
    assert.deepEqual(f.control.getItemState(), { checkedItems: [], partialItems: ['1', 'true'] })
    assert.deepEqual(f.shared.getItemState(), f.control.getItemState())
    for (const id of [f.ids.owned, f.ids.linked]) assert.equal(total(f.book.getPivotChartById(id)), 40)
    assert.deepEqual(new Set(events.map(event => event.pivotChartId)), new Set([f.ids.owned, f.ids.linked]))
    assert.deepEqual(f.model().sheets[0].cells, source)
    const after = authored(f.book); f.book.undo(); assert.deepEqual(authored(f.book), before)
    f.book.redo(); assert.deepEqual(authored(f.book), after)
    reopened = new api.OfficeStudio({ workbook: f.book.saveModel() })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getSlicer(f.ids.slicer)
    assert.deepEqual(restored.getRecordSelection(), { mode: 'include', values: [1, true] })
    assert.deepEqual(restored.getItemState(), { checkedItems: [], partialItems: ['1', 'true'] })
  } finally { reopened?.dispose(); f.dispose() }
})

test('the Canvas/Ant toggle planner preserves untouched raw identities and clears the implicit full selection', async () => {
  const f = setup()
  try {
    await connect(f); await f.control.setRecordSelection({ mode: 'include', values: [1, true] })
    const toggle = (index, state) => selection.slicerControlItemStates(f.control.getItemLabels(), state,
      { index, value: f.control.getItemLabels()[index], ctrlKey: true, metaKey: false, shiftKey: false, multiSelect: false })
    const state = toggle(1, f.control.getItemState())
    assert.equal(await f.control.setItemStates(state.checkedItems, state.partialItems), true)
    assert.deepEqual(f.control.getItemState(), { checkedItems: ['true'], partialItems: ['1'] })
    assert.equal(total(f.book.getPivotChartById(f.ids.owned)), 80)
    assert.ok(f.control.getRecordSelection().values.some(value => value === 1))
    assert.ok(!f.control.getRecordSelection().values.some(value => value === '1'))
    const all = toggle(0, f.control.getItemState())
    assert.equal(await f.control.setItemStates(all.checkedItems, all.partialItems), true)
    assert.equal(f.control.getRecordSelection(), undefined); assert.equal(f.control.getSelectedItems(), undefined)
    assert.equal(total(f.book.getPivotChartById(f.ids.owned)), 100)
  } finally { f.dispose() }
})

test.each(['include', 'exclude'])('%s selection and display state getters are defensive and keep absent typed conditions', async mode => {
  const f = setup()
  try {
    assert.equal(await f.control.setRecordSelection({ mode, values: [1, 'absent'] }), true)
    const records = f.control.getRecordSelection(), state = f.control.getItemState(), labels = f.control.getSelectedItems()
    records.values.push('corrupted'); state.partialItems.push('corrupted'); labels.push('corrupted')
    assert.deepEqual(f.control.getRecordSelection(), { mode, values: [1, 'absent'] })
    assert.ok(!f.control.getItemState().partialItems.includes('corrupted')); assert.ok(!f.control.getSelectedItems().includes('corrupted'))
  } finally { f.dispose() }
})

test('Before veto, stale revision and stale host reject item transactions without changing reports or history', async () => {
  const f = setup()
  try {
    const before = authored(f.book), baseline = f.history()
    const veto = f.studio.events.addEvent('BeforeSlicerChange', event => { event.cancel = true })
    await assert.rejects(f.control.setItemStates(['1'], []), /canceled|取消/i); veto.dispose()
    await assert.rejects(f.control.setItemStates(['1'], [], { expectedRevision: -1 }), /版本/)
    f.book.setActiveSheet(f.ids.other)
    const changed = authored(f.book), revision = f.studio.getWorkbookRevision(f.unitId)
    await assert.rejects(f.control.setItemStates(['1'], [], { expectedRevision: revision, host: 'workbench' }), /宿主/)
    assert.deepEqual(authored(f.book), changed); assert.equal(f.history(), baseline)
    assert.deepEqual(changed.slicerCaches, before.slicerCaches)
  } finally { f.dispose() }
})

test('read and mutation gates enforce every connected report and source range permission', async () => {
  const f = setup()
  try {
    await connect(f)
    f.model().sheets.find(sheet => sheet.id === f.ids.other).worksheetProtection = { id: 'locked-report', ownerUserId: 'peer',
      allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
    const before = authored(f.book), baseline = f.history()
    assert.equal(await f.control.setItemStates(['1'], []), false)
    assert.deepEqual(authored(f.book), before); assert.equal(f.history(), baseline)
    f.model().sheets[0].worksheetProtection = { id: 'hidden-source', ownerUserId: 'peer', allowedUserIds: [],
      editState: 'onlyMe', viewState: 'noOneElseCanView' }
    for (const getter of ['getItems', 'getSelectedItems', 'getRecordSelection', 'getItemState']) assert.throws(() => f.control[getter](), /权限|查看/)
  } finally { f.dispose() }
})
