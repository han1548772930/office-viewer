import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Real registry/Facade/selection/history drafts; execution remains deferred.
let server, api, fixture, selection, menu, pivot
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture, selection, menu, pivot] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-axis-state-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/src/api/pivot-chart-axis-selection.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-axis-menu.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-engine.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function authorSnapshot(book) { const snapshot = json(book.getSnapshot()); delete snapshot.modified; return snapshot }
function setup(ownership = 'owned') {
  const studio = new api.OfficeStudio({ workbook: fixture.pivotAxisWorkbook({ ownership }) })
  const office = api.FOffice.newAPI(studio), book = office.getActiveWorkbook(), chart = book.getPivotChartById('axis-chart-a')
  return { studio, office, book, chart, unitId: book.getId(), model: () => studio.getWorkbookModel(book.getId()),
    history: () => studio.editor.history.units[book.getId()]?.undo.length ?? 0, dispose: () => studio.dispose() }
}
function capture(f, owner = {}, fieldId = 'region-label', item = 'North') {
  const model = f.model(), chart = model.charts.find(chart => chart.id === f.chart.getId()), report = pivot.getPivotChartReport(model, chart)
  const input = { unitId: f.unitId, subUnitId: 'axis-host', pivotChartId: chart.id, fieldId, item,
    collapsed: f.chart.getCollapse(fieldId, item), fields: f.chart.getAxisState().fields, host: 'workbench',
    expectedRevision: f.studio.getWorkbookRevision(f.unitId), expectedUserId: f.studio.currentUserId,
    expectedHistoryGeneration: f.studio.editor.history.captureSavepoint(f.unitId).generation,
    chartSnapshot: JSON.stringify(chart), reportSnapshot: JSON.stringify(report) }
  selection.setPivotChartAxisSelection(f.studio, owner, input)
  return { owner, input }
}

test.each(['owned', 'linked'])('%s Facade collapse and category-level commands share one content history and saved report state', async ownership => {
  const f = setup(ownership); let reopened
  try {
    const baseline = f.history(), before = authorSnapshot(f.book), source = json(f.model().sheets[0].cells)
    assert.equal(await f.chart.setCollapse('region-label', true, 'North'), true)
    assert.equal(f.history(), baseline + 1); assert.equal(f.chart.getCollapse('region-label', 'North'), true)
    assert.equal(f.chart.getData().categories.length, 5); assert.deepEqual(f.model().sheets[0].cells, source)
    const saved = authorSnapshot(f.book)
    f.book.undo(); assert.deepEqual(authorSnapshot(f.book), before)
    f.book.redo(); assert.deepEqual(authorSnapshot(f.book), saved)
    reopened = new api.OfficeStudio({ workbook: saved })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById(f.chart.getId())
    assert.equal(restored.getCollapse('region-label', 'North'), true)
    assert.equal(await restored.setCategoryLevel('product-label'), true)
    assert.equal(restored.getCollapse('region-label'), false); assert.equal(restored.getCollapse('product-label'), true)
    assert.equal(await restored.stepCategoryLevel('expand'), true); assert.equal(restored.getCollapse('product-label'), false)
  } finally { reopened?.dispose(); f.dispose() }
})

test.each(['owned', 'linked'])('%s no-parameter item aliases use the current frame selection and never invent an active member', async ownership => {
  const f = setup(ownership)
  try {
    const baseline = f.history(), commands = api.PIVOT_HIERARCHY_COMMANDS
    assert.equal(await f.studio.executeCommand(commands.chartCollapseItem, undefined), false)
    assert.equal(f.history(), baseline)
    capture(f)
    assert.equal(await f.studio.executeCommand(commands.chartCollapseItem, undefined), true)
    assert.equal(f.chart.getCollapse('region-label', 'North'), true)
    assert.equal(selection.getPivotChartAxisSelection(f.studio), undefined)
    assert.equal(await f.studio.executeCommand(commands.chartExpandItem, undefined), false)
    capture(f)
    assert.equal(await f.studio.executeCommand(commands.chartExpandItem, undefined), true)
    assert.equal(f.chart.getCollapse('region-label', 'North'), false)
    assert.equal(f.history(), baseline + 2)
  } finally { f.dispose() }
})

test('selection strips runtime owners, returns defensive control DTOs and prevents a stale owner from clearing a replacement', () => {
  const f = setup()
  try {
    const first = capture(f), circular = {}; circular.self = circular
    selection.setPivotChartAxisSelection(f.studio, first.owner, { ...first.input, owner: { studio: f.studio, circular }, target: circular })
    const read = selection.getPivotChartAxisSelection(f.studio)
    assert.equal(read.owner, undefined); assert.equal(read.target, undefined)
    read.fields[0].name = 'Changed'; assert.equal(selection.getPivotChartAxisSelection(f.studio).fields[0].name, 'Region')
    const replacement = capture(f)
    selection.clearPivotChartAxisSelection(f.studio, first.owner)
    assert.ok(selection.getPivotChartAxisSelection(f.studio))
    selection.clearPivotChartAxisSelection(f.studio, replacement.owner)
    assert.equal(selection.getPivotChartAxisSelection(f.studio), undefined)
  } finally { f.dispose() }
})

test.each(['source-edit', 'switch-sheet', 'switch-unit', 'user', 'history-generation', 'chart-snapshot', 'report-snapshot'])
  ('%s invalidates a no-parameter axis target before any content command', async boundary => {
    const f = setup()
    try {
      capture(f)
      if (boundary === 'source-edit') f.book.getSheetBySheetId('axis-data').getRange('I3').setValue(11)
      if (boundary === 'switch-sheet') f.book.setActiveSheet('axis-dashboard')
      if (boundary === 'switch-unit') f.office.createWorkbook({ name: 'Other' })
      if (boundary === 'user') Object.defineProperty(f.studio, 'currentUserId', { value: 'peer', configurable: true })
      if (boundary === 'history-generation') f.studio.editor.history.clear(f.unitId)
      if (boundary === 'chart-snapshot') f.model().charts[0].title = 'Replaced without revision'
      if (boundary === 'report-snapshot') f.model().charts[0].ownedPivot.options.repeatRowLabels = false
      const before = json(f.model()), baseline = f.history()
      assert.equal(selection.getPivotChartAxisSelection(f.studio), undefined)
      assert.equal(await f.studio.executeCommand(api.PIVOT_HIERARCHY_COMMANDS.chartCollapseItem, undefined), false)
      assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
    } finally { f.dispose() }
  })

test('axis menus preserve the four fixed entries followed by every ordered row field and the source expand/collapse wording', () => {
  const f = setup()
  try {
    const { input } = capture(f), entries = menu.pivotChartAxisMenu(input)
    assert.deepEqual(entries.map(entry => entry.key), ['expand-item', 'collapse-item', 'expand-field', 'collapse-field',
      'category-level:region-label', 'category-level:product-label', 'category-level:item-label'])
    assert.deepEqual(entries.slice(4).map(entry => entry.label), ['折叠到“Region”', '展开到“Product”', '展开到“Item”'])
    assert.equal(entries[0].disabled, true); assert.equal(entries[1].disabled, false)
    assert.equal(entries[0].params, undefined)
    assert.ok(entries.slice(2).every(entry => entry.params.subUnitId === 'axis-host'))
    assert.ok(menu.pivotChartAxisMenu(input, true).every(entry => entry.disabled))
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s Before veto and conflicting host aliases leave report/chart output and history unchanged', async ownership => {
  const f = setup(ownership)
  try {
    const before = json(f.model()), baseline = f.history(), target = { unitId: f.unitId, subUnitId: 'axis-host', pivotChartId: f.chart.getId(),
      fieldId: 'region-label', collapse: true, item: 'North' }
    await assert.rejects(() => f.studio.executeCommand(api.PIVOT_HIERARCHY_COMMANDS.chartSetCollapse, { ...target, sheetId: 'axis-dashboard' }), /宿主/)
    f.studio.events.addEvent('BeforePivotChartChange', event => { event.cancel = true })
    await assert.rejects(() => f.studio.executeCommand(api.PIVOT_HIERARCHY_COMMANDS.chartSetCollapse, target), /取消/)
    assert.deepEqual(json(f.model()), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test('ordinary Pivot hierarchy accepts an explicit workbench host separate from its output sheet', async () => {
  const f = setup('linked')
  try {
    f.book.setActiveSheet('axis-dashboard')
    const baseline = f.history(), report = f.book.getPivotTables()[0]
    assert.equal(await report.setCollapse('region-label', true, 'North', { host: 'workbench', hostSheetId: 'axis-dashboard' }), report)
    assert.equal(f.history(), baseline + 1); assert.equal(report.getCollapse('region-label', 'North'), true)
    assert.equal(f.book.getPivotChartById('axis-chart-a').getData().categories.length, 5)
  } finally { f.dispose() }
})
