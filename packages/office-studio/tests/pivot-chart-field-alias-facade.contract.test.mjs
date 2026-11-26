import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Real registry, Facade, output, history and reopen drafts. Execution remains deferred.
let server, api, fixture
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-chart-engine-fixture.mjs'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function authored(book) { const snapshot = json(book.getSnapshot()); delete snapshot.modified; return snapshot }
function setup(ownership = 'owned') {
  const workbook = fixture.pivotChartEngineWorkbook({ ownership })
  for (let row = 3; row <= 6; row++) workbook.sheets[0].cells[`${row}:5`].style.numberFormat = 'yyyy-mm-dd'
  const fields = [...fixture.pivotChartFields(), { id: 'date-local', sourceFieldId: 'date-source', sourceColumn: 5,
    name: 'Date', area: 'filter', position: 0 }]
  api.applyPivotChartEdit(workbook, workbook.charts[0], { type: 'setFieldsConfig', config: { fields, options: json(fixture.pivotChartReportOptions) } })
  const studio = new api.OfficeStudio({ workbook }), office = api.FOffice.newAPI(studio), book = office.getActiveWorkbook()
  return { studio, office, book, unitId: book.getId(), chart: book.getPivotChartById('chart-a'),
    history: () => studio.editor.history.units[book.getId()]?.undo.length ?? 0,
    model: () => studio.getWorkbookModel(book.getId()), dispose: () => studio.dispose() }
}
const target = f => ({ unitId: f.unitId, subUnitId: 'reports', pivotChartId: f.chart.getId() })
const context = f => ({ expectedRevision: f.studio.getWorkbookRevision(f.unitId),
  expectedHistoryGeneration: f.studio.editor.history.captureSavepoint(f.unitId).generation,
  expectedUserId: f.studio.currentUserId, host: 'workbench' })
const total = chart => chart.getData().series[0].values.reduce((sum, value) => sum + value, 0)
const fieldActions = [
  ['add', () => ({ dataFieldId: 'segment-source', fieldArea: 'filter', index: 0 }), chart => {
    const ids = chart.getFieldIdsByArea('filter'); assert.equal(ids.length, 2)
    assert.equal(chart.getFieldSetting(ids[0]).sourceFieldId, 'segment-source')
  }],
  ['move', () => ({ fieldId: 'region-local', fieldArea: 'column', index: 0 }), chart => {
    assert.deepEqual(chart.getFieldIdsByArea('row'), []); assert.deepEqual(chart.getFieldIdsByArea('column'), ['region-local'])
  }],
  ['remove', () => ({ fieldIds: ['region-local'] }), chart => assert.equal(chart.getFieldSetting('region-local'), undefined)],
  ['setOptions', () => ({ options: { showRowGrandTotal: true } }), chart => assert.equal(chart.getOptions().showRowGrandTotal, true)],
  ['setFieldsConfig', () => { const fields = fixture.pivotChartFields(); fields[1].aggregation = 'count'
    return { pivotTableConfig: { fields, options: json(fixture.pivotChartReportOptions) } }
  }, chart => assert.deepEqual(chart.getData().series[0].values, [2, 1, 1])],
  ['setFieldSetting', () => ({ fieldId: 'region-local', setting: { displayName: 'Territory' } }), chart => assert.equal(chart.getFieldSetting('region-local').displayName, 'Territory')],
  ['setFilter', () => ({ fieldId: 'region-local', items: ['North'], partialItems: [] }), chart => {
    assert.deepEqual(chart.getData().categories, ['North']); assert.equal(total(chart), 40)
    assert.deepEqual(chart.getManualFilterItemState('region-local'), { checkedItems: ['North'], partialItems: [] })
  }],
  ['setSort', () => ({ fieldId: 'region-local', info: { type: 'descending' } }), chart => assert.deepEqual(chart.getData().categories, ['West', 'South', 'North'])],
  ['setValueFilter', () => ({ fieldId: 'region-local', valueFilterInfo: { operator: 'valueGreaterThan', expected: 30, valueFieldId: 'amount-local' } }), chart => {
    assert.deepEqual(new Set(chart.getData().categories), new Set(['North', 'West'])); assert.equal(total(chart), 80)
  }],
  ['setDateGroup', () => ({ fieldId: 'date-local', dateType: 'year' }), chart => {
    assert.equal(chart.getFieldSetting('date-local').dateGroupType, 'year'); assert.deepEqual(chart.getFieldItems('date-local'), ['2024'])
  }],
]

test.each(['owned', 'linked'].flatMap(ownership => fieldActions.map(([action, input, verify]) => [ownership, action, input, verify])))
  ('%s fixed %s alias preserves command identity, changes the real report and reopens from one history entry', async (ownership, action, input, verify) => {
    const f = setup(ownership); let reopened
    try {
      const before = authored(f.book), source = json(f.model().sheets[0].cells), baseline = f.history(), ids = [], beforeIds = []
      const commandId = api.PIVOT_CHART_FIELD_COMMANDS[action]
      f.studio.events.addEvent('CommandExecuted', event => ids.push(event.id))
      f.studio.events.addEvent('BeforePivotChartChange', event => beforeIds.push(event.commandId))
      assert.equal(await f.studio.executeCommand(commandId, { ...target(f), ...input() }), true)
      verify(f.chart); assert.equal(f.history(), baseline + 1)
      assert.ok(ids.includes(commandId)); assert.deepEqual(beforeIds, [commandId]); assert.deepEqual(f.model().sheets[0].cells, source)
      if (ownership === 'linked') verify(f.book.getPivotChartById('chart-b'))
      const after = authored(f.book)
      f.book.undo(); assert.deepEqual(authored(f.book), before)
      f.book.redo(); assert.deepEqual(authored(f.book), after)
      const saved = f.book.save(); assert.equal(saved.modified, false)
      reopened = new api.OfficeStudio({ workbook: saved })
      verify(api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById('chart-a'))
    } finally { reopened?.dispose(); f.dispose() }
  })

test.each(['owned', 'linked'])('%s Facade routes field and report operations through the fixed command aliases', async ownership => {
  const f = setup(ownership)
  try {
    const ids = [], events = f.studio.events.addEvent('CommandExecuted', event => ids.push(event.id))
    await f.chart.renameField('region-local', 'Territory')
    await f.chart.setOptions({ repeatRowLabels: true })
    await f.chart.setLabelManualFilter('region-local', ['North'])
    await f.chart.setLabelSort('region-local', { type: 'descending' })
    await f.chart.clearLabelSort('region-local')
    assert.deepEqual(ids, [api.PIVOT_CHART_FIELD_COMMANDS.setFieldSetting, api.PIVOT_CHART_FIELD_COMMANDS.setOptions,
      api.PIVOT_CHART_FIELD_COMMANDS.setFilter, api.PIVOT_CHART_FIELD_COMMANDS.setSort, api.PIVOT_CHART_FIELD_COMMANDS.setSort])
    assert.equal(f.chart.getFieldSetting('region-local').displayName, 'Territory'); assert.equal(total(f.chart), 40)
    events.dispose()
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s partial control updates keep the other persisted switch and survive undo and reopen', async ownership => {
  const f = setup(ownership); let reopened
  try {
    const baseline = f.history()
    assert.equal(await f.chart.setControls({ showExpandCollapseButtons: false }), true)
    assert.equal(await f.studio.executeCommand(api.PIVOT_CHART_COMMANDS.setControls,
      { ...target(f), controls: { showFieldButtons: false } }), true)
    assert.deepEqual(f.chart.getControls(), { showFieldButtons: false, showExpandCollapseButtons: false })
    assert.equal(f.history(), baseline + 2)
    const copy = f.chart.getControls(); copy.showFieldButtons = true
    assert.equal(f.chart.getControls().showFieldButtons, false)
    f.book.undo(); assert.deepEqual(f.chart.getControls(), { showFieldButtons: true, showExpandCollapseButtons: false })
    f.book.redo(); assert.deepEqual(f.chart.getControls(), { showFieldButtons: false, showExpandCollapseButtons: false })
    reopened = new api.OfficeStudio({ workbook: f.book.save() })
    assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById('chart-a').getControls(), f.chart.getControls())
  } finally { reopened?.dispose(); f.dispose() }
})

test.each(['owned', 'linked'])('%s invalid identities, source fields and controls never create an undo entry', async ownership => {
  const f = setup(ownership)
  try {
    const before = authored(f.book), baseline = f.history()
    await assert.rejects(() => f.studio.executeCommand(api.PIVOT_CHART_FIELD_COMMANDS.setFilter,
      { ...target(f), sheetId: 'dashboard', fieldId: 'region-local', items: ['North'] }), /宿主/)
    await assert.rejects(() => f.chart.addField('missing-source', 'row'), /field|字段/i)
    await assert.rejects(() => f.chart.setControls({ showFieldButtons: 1 }), /controls|控件|boolean|布尔/i)
    assert.deepEqual(authored(f.book), before); assert.equal(f.history(), baseline)
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s Before veto and stale history context cancel aliases before report, chart and output publication', async ownership => {
  const f = setup(ownership)
  try {
    const before = authored(f.book), captured = context(f)
    const veto = f.studio.events.addEvent('BeforePivotChartChange', event => { event.cancel = true })
    await assert.rejects(() => f.chart.setLabelManualFilter('region-local', ['North'], false, captured), /取消/)
    veto.dispose(); f.studio.editor.history.clear(f.unitId)
    await assert.rejects(() => f.chart.setLabelManualFilter('region-local', ['North'], false, captured), /历史|更改/)
    assert.deepEqual(authored(f.book), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})
