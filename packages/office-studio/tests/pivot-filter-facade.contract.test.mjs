import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Public enum adapters and cross-module report edits; all execution remains deferred.
let server, api, fixture, engine
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixture, engine] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/pivot-chart-engine-fixture.mjs'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-engine.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function authored(book) { const snapshot = json(book.getSnapshot()); delete snapshot.modified; return snapshot }
function setup(ownership = 'owned', date1904 = false) {
  const workbook = fixture.pivotChartEngineWorkbook({ ownership }), data = workbook.sheets[0]
  workbook.date1904 = date1904
  ;[45292.25, 45323.5, 45352.75, 45383].forEach((serial, index) => {
    const cell = data.cells[`${index + 3}:5`]; cell.value = serial; cell.style.numberFormat = 'yyyy-mm-dd hh:mm'
  })
  api.applyPivotChartEdit(workbook, workbook.charts[0], { type: 'setFieldsConfig', config: {
    fields: [...fixture.pivotChartFields(), { id: 'date-local', sourceFieldId: 'date-source', sourceColumn: 5, name: 'Date', area: 'filter', position: 0 },
      { id: 'region-filter', sourceFieldId: 'region-source', sourceColumn: 3, name: 'Region', area: 'filter', position: 1 }],
    options: json(fixture.pivotChartReportOptions),
  } })
  const studio = new api.OfficeStudio({ workbook }), book = api.FOffice.newAPI(studio).getActiveWorkbook()
  return { studio, book, unitId: book.getId(), chart: book.getPivotChartById('chart-a'),
    model: () => studio.getWorkbookModel(book.getId()), history: () => studio.editor.history.units[book.getId()]?.undo.length ?? 0,
    dispose: () => studio.dispose() }
}
const total = chart => (chart.getData().series[0]?.values ?? []).reduce((sum, value) => sum + value, 0)
const dateRules = [[16, [45292.25, 45352.75], 60], [17, 45323.5, 20], [18, 45323.5, 70], [19, 45323.5, 90],
  [20, [45323.5, 45352.75], 50], [21, 45323.5, 80], [22, 45323.5, 10], [23, 45323.5, 30]]
const labelRules = [[1, 'N', 40], [2, ['North', 'South'], 60], [3, 'o', 60], [4, 'th', 60], [5, 'South', 20],
  [6, 'South', 40], [7, 'South', 60], [8, 'South', 40], [9, 'South', 60], [10, 'N', 60],
  [11, ['North', 'South'], 40], [12, 'o', 40], [13, 'th', 40], [14, 'South', 80]]

test.each(['owned', 'linked'].flatMap(ownership => dateRules.map(rule => [ownership, ...rule])))
  ('%s fixed date operator %s filters the real source serials and preserves the author DTO on reopen', async (ownership, operator, expected, amount) => {
    const f = setup(ownership); let reopened
    try {
      const before = authored(f.book), source = json(f.model().sheets[0].cells), baseline = f.history()
      assert.deepEqual(f.chart.getFilterCapabilities('date-local'), { supportsDateFilter: true, date1904: false })
      assert.equal(f.chart.getFilterCapabilities('amount-local').supportsDateFilter, false)
      assert.equal(await f.chart.setDateFilter('date-local', { type: 1, operator, expected }), true)
      assert.equal(total(f.chart), amount); assert.equal(f.history(), baseline + 1); assert.deepEqual(f.model().sheets[0].cells, source)
      const filter = f.chart.getDateFilter('date-local'); assert.equal(filter.useWholeDay, false)
      const caller = json(filter); if (Array.isArray(caller.expected)) caller.expected[0] = -1
      else caller.expected = -1
      assert.deepEqual(f.chart.getDateFilter('date-local'), filter)
      if (ownership === 'linked') assert.equal(total(f.book.getPivotChartById('chart-b')), amount)
      const after = authored(f.book); f.book.undo(); assert.deepEqual(authored(f.book), before)
      f.book.redo(); assert.deepEqual(authored(f.book), after)
      reopened = new api.OfficeStudio({ workbook: f.book.save() })
      const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById('chart-a')
      assert.deepEqual(restored.getDateFilter('date-local'), filter); assert.equal(total(restored), amount)
    } finally { reopened?.dispose(); f.dispose() }
  })

test.each(['owned', 'linked'].flatMap(ownership => labelRules.map(rule => [ownership, ...rule])))
  ('%s fixed caption operator %s filters the canonical source through the same report transaction', async (ownership, operator, expected, amount) => {
    const f = setup(ownership)
    try {
      const baseline = f.history()
      assert.equal(await f.chart.setLabelConditionFilter('region-filter', { type: 1, operator, expected }), true)
      assert.equal(total(f.chart), amount); assert.equal(f.history(), baseline + 1)
      assert.deepEqual(f.chart.getLabelConditionFilter('region-filter').expected, expected)
      if (ownership === 'linked') assert.equal(total(f.book.getPivotChartById('chart-b')), amount)
      await f.chart.clearLabelConditionFilter('region-filter'); assert.equal(total(f.chart), 100)
    } finally { f.dispose() }
  })

test.each([false, true])('whole-day comparison preserves workbook date1904=%s and differs from exact serial equality', async date1904 => {
  const f = setup('owned', date1904)
  try {
    await f.chart.setDateFilter('date-local', { operator: 'dateEqual', expected: 45323.1 })
    assert.equal(total(f.chart), 0)
    await f.chart.setDateFilter('date-local', { operator: 'dateEqual', expected: 45323.1, useWholeDay: true })
    assert.equal(total(f.chart), 20); assert.equal(f.chart.getFilterCapabilities('date-local').date1904, date1904)
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s exclusive filters replace other categories on every report-local instance of one source', async ownership => {
  const f = setup(ownership)
  try {
    await f.chart.setLabelManualFilter('region-local', ['North'])
    await f.chart.setLabelConditionFilter('region-filter', { operator: 'captionEqual', expected: 'South' })
    assert.deepEqual(f.chart.getManualFilterItemState('region-local').checkedItems, ['North', 'South', 'West'])
    assert.equal(total(f.chart), 20)
    await f.chart.setValueFilter('region-local', { operator: 'valueGreaterThan', expected: 30, valueFieldId: 'amount-local' })
    assert.equal(f.chart.getLabelConditionFilter('region-filter'), undefined); assert.equal(total(f.chart), 80)
    assert.equal(f.chart.getReportOptions().allowMultipleFilters ?? false, false)
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s disabling multiple filters clears all report categories atomically when any source conflicts', async ownership => {
  const f = setup(ownership)
  try {
    await f.chart.setReportOptions({ allowMultipleFilters: true })
    await f.chart.setLabelManualFilter('region-local', ['North'])
    await f.chart.setLabelConditionFilter('region-filter', { operator: 'captionNotEqual', expected: 'South' })
    await f.chart.setValueFilter('region-local', { operator: 'valueGreaterThan', expected: 35, valueFieldId: 'amount-local' })
    await f.chart.setDateFilter('date-local', { operator: 'dateBetween', expected: [45290, 45390] })
    assert.equal(total(f.chart), 40)
    const before = authored(f.book), baseline = f.history()
    await f.chart.setReportOptions({ allowMultipleFilters: false })
    assert.equal(f.history(), baseline + 1); assert.equal(total(f.chart), 100)
    assert.equal(f.chart.getDateFilter('date-local'), undefined); assert.equal(f.chart.getLabelConditionFilter('region-filter'), undefined)
    assert.equal(f.chart.getValueFilter('region-local'), undefined)
    assert.deepEqual(f.chart.getManualFilterItemState('region-local'), { checkedItems: ['North', 'South', 'West'], partialItems: [] })
    f.book.undo(); assert.deepEqual(authored(f.book), before); assert.equal(total(f.chart), 40)
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s isAll removes only the member category and clearAll ignores incoming items while preserving unrelated date conditions', async ownership => {
  const f = setup(ownership)
  try {
    await f.chart.setReportOptions({ allowMultipleFilters: true })
    await f.chart.setLabelManualFilter('region-local', ['North'])
    await f.chart.setLabelConditionFilter('region-filter', { operator: 'captionNotEqual', expected: 'South' })
    await f.chart.setValueFilter('region-local', { operator: 'valueGreaterThan', expected: 30, valueFieldId: 'amount-local' })
    await f.chart.setDateFilter('date-local', { operator: 'dateBetween', expected: [45290, 45390] })
    await f.chart.setLabelManualFilter('region-local', [], true)
    assert.equal(total(f.chart), 80); assert.ok(f.chart.getLabelConditionFilter('region-filter')); assert.ok(f.chart.getValueFilter('region-local'))
    await f.studio.executeCommand(api.PIVOT_CHART_FIELD_COMMANDS.setFilter, { unitId: f.unitId, subUnitId: 'reports', pivotChartId: 'chart-a',
      fieldId: 'region-local', items: ['not-a-member'], partialItems: ['also-not-a-member'], clearAllFilters: true })
    assert.equal(total(f.chart), 100); assert.ok(f.chart.getDateFilter('date-local'))
    assert.equal(f.chart.getLabelConditionFilter('region-filter'), undefined); assert.equal(f.chart.getValueFilter('region-local'), undefined)
  } finally { f.dispose() }
})

test('ordinary report Facade uses fixed field aliases with an independent workbench host and restores date conditions', async () => {
  const f = setup('linked'); let reopened
  try {
    f.book.setActiveSheet('dashboard')
    const pivot = f.book.getPivotTables()[0], ids = [], payloads = [], baseline = f.history()
    f.studio.events.addEvent('CommandExecuted', event => ids.push(event.id))
    f.studio.events.addEvent('BeforePivotTableChange', event => payloads.push(event))
    await pivot.setDateFilter('date-local', { operator: 17, expected: 45323.5 }, { host: 'workbench', hostSheetId: 'dashboard' })
    assert.equal(total(f.chart), 20); assert.equal(f.history(), baseline + 1)
    assert.deepEqual(ids, [api.PIVOT_REPORT_FILTER_COMMANDS.setDateFilter]); assert.equal(payloads[0].subUnitId, 'reports')
    reopened = new api.OfficeStudio({ workbook: f.book.save() })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotTables()[0]
    assert.deepEqual(restored.getDateFilter('date-local'), pivot.getDateFilter('date-local'))
    await pivot.clearDateFilter('date-local', { host: 'workbench', hostSheetId: 'dashboard' }); assert.equal(total(f.chart), 100)
    await pivot.setLabelManualFilter('region-local', ['North'], false, { host: 'workbench', hostSheetId: 'dashboard' })
    assert.equal(total(f.chart), 40); assert.equal(ids.at(-1), api.PIVOT_REPORT_FILTER_COMMANDS.setFilter)
  } finally { reopened?.dispose(); f.dispose() }
})

test.each(['operator', 'pair-order', 'nonfinite', 'wrong-type', 'value-field', 'permission'])
  ('%s date boundary cannot change reports, output, source cells or history', async boundary => {
    const f = setup()
    try {
      if (boundary === 'permission') f.model().sheets[1].worksheetProtection = { id: 'locked-host', permissionId: 'host-permission',
        unitType: 'worksheet', ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
      const before = authored(f.book), baseline = f.history()
      const fieldId = boundary === 'value-field' ? 'amount-local' : 'date-local'
      const input = boundary === 'operator' ? { operator: 24, expected: 45323.5 }
        : boundary === 'pair-order' ? { operator: 16, expected: [45352, 45323] }
        : boundary === 'nonfinite' ? { operator: 17, expected: Infinity }
        : boundary === 'wrong-type' ? { type: 0, operator: 17, expected: 45323.5 }
        : { operator: 17, expected: 45323.5 }
      await assert.rejects(() => f.chart.setDateFilter(fieldId, input))
      assert.deepEqual(authored(f.book), before); assert.equal(f.history(), baseline)
    } finally { f.dispose() }
  })

test('manual filter arrays expose defensive state and retain a partial typed group through the fixed chart alias', async () => {
  const f = setup()
  try {
    // Same display key with number and text is deliberately created in the real sparse source.
    const model = f.model(), source = model.sheets[0], report = model.charts[0].ownedPivot
    source.cells['3:3'] = api.createCell(3, 3, 1); source.cells['4:3'] = api.createCell(4, 3, '1')
    source.cells['5:3'] = api.createCell(5, 3, 'Other'); source.cells['6:3'] = api.createCell(6, 3, 'Other')
    engine.setPivotRecordSelection(model, report, 'region-local', { mode: 'include', values: [1] }, 'deferred')
    const initial = f.chart.getManualFilterItemState('region-local')
    assert.deepEqual(initial, { checkedItems: [], partialItems: ['1'] })
    initial.partialItems.push('Other'); assert.deepEqual(f.chart.getManualFilterItemState('region-local').partialItems, ['1'])
    await f.studio.executeCommand(api.PIVOT_CHART_FIELD_COMMANDS.setFilter, { unitId: f.unitId, subUnitId: 'reports', pivotChartId: 'chart-a',
      fieldId: 'region-local', items: ['Other'], partialItems: ['1'] })
    assert.equal(total(f.chart), 80); assert.deepEqual(f.chart.getManualFilterItemState('region-local'), { checkedItems: ['Other'], partialItems: ['1'] })
    const selection = f.chart.getFieldSetting('region-local').recordSelection
    assert.ok(selection.values.some(value => value === 1 && typeof value === 'number'))
    assert.ok(!selection.values.some(value => value === '1'))
  } finally { f.dispose() }
})
