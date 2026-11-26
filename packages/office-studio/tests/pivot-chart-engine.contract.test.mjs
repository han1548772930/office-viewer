import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createCell } from '../src/domain/model.ts'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { calculatePivotReport, getPivotSourceFieldItems, pivotDrillDown, pivotOutputValues, refreshPivotTable,
  renderPivotTable, setPivotField, setPivotFieldSetting, setPivotLabelFilter, setPivotManualGroups, setPivotNumberGroup,
  setPivotOptions, setPivotSort, setPivotValueFilter, PivotOutputConflictError } from '../src/features/pivot/services/pivot-engine.ts'
import { applyPivotChartEdit, createPivotChartReport, getPivotChartFieldItems, getPivotChartReport,
  getPivotChartSourceFields, PivotChartEditError } from '../src/features/pivot/services/pivot-chart-engine.ts'
import { pivotChartEngineWorkbook, pivotChartFields, pivotChartReportOptions, pivotChartSourceInput } from './fixtures/pivot-chart-engine-fixture.mjs'

// Contract drafts only: do not execute before the project's unified verification phase.
const clone = value => structuredClone(value)
const cells = workbook => clone(workbook.sheets.map(sheet => ({ id: sheet.id, cells: sheet.cells })))
const first = workbook => workbook.charts[0]
const effective = workbook => getPivotChartReport(workbook, first(workbook))
const matrix = workbook => calculatePivotReport(workbook, effective(workbook)).values
const edit = (workbook, value) => applyPivotChartEdit(workbook, first(workbook), value)
const configure = (workbook, fields, options = pivotChartReportOptions) => edit(workbook, { type: 'setFieldsConfig', config: { fields, options } })
const dimension = (sourceFieldId, sourceColumn, name) => ({ id: `${sourceFieldId}-local`, sourceFieldId, sourceColumn, name, area: 'row', position: 0 })
const invalidEdit = error => error instanceof PivotChartEditError && error.code === 'invalid-edit'
const connectedEdit = error => error instanceof PivotChartEditError && error.code === 'field-connected'

test('empty owned reports use the actual host without registering a pivot, creating a sheet, or reserving cells', () => {
  const workbook = pivotChartEngineWorkbook({ ownership: 'none' }), before = cells(workbook), sheetIds = workbook.sheets.map(sheet => sheet.id)
  workbook.sheets[1].cells['0:0'] = createCell(0, 0, 'Occupied host cell')
  const occupied = cells(workbook), report = createPivotChartReport(workbook, clone(pivotChartSourceInput), 'reports', 'new-owned')
  assert.deepEqual(report.fields, []); assert.equal(report.revision, 0); assert.equal(report.pivotCacheId, 'engine-cache')
  assert.deepEqual(report.anchor, { unitId: workbook.id, subUnitId: 'reports', row: 0, col: 0 })
  assert.deepEqual(workbook.sheets.map(sheet => sheet.id), sheetIds); assert.deepEqual(workbook.pivotTables, [])
  assert.deepEqual(cells(workbook), occupied); assert.notDeepEqual(before, occupied)
  for (const key of ['outputRange', 'generatedCellKeys', 'drillDownMap']) assert.equal(Object.hasOwn(report, key), false)
  assert.deepEqual(calculatePivotReport(workbook, report).values, [])
})

test('owned source metadata and unconfigured raw members remain readable before any placement', () => {
  const workbook = pivotChartEngineWorkbook({ configured: false }), before = clone(workbook)
  assert.deepEqual(getPivotChartSourceFields(workbook, first(workbook)).map(field => field.fieldId),
    ['region-source', 'amount-source', 'date-source', 'segment-source', 'score-source'])
  assert.deepEqual(getPivotChartFieldItems(workbook, first(workbook), 'region-source'), ['North', 'South', 'West'])
  assert.deepEqual(getPivotChartFieldItems(workbook, first(workbook), 'segment-source'), ['A', 'B'])
  assert.deepEqual(workbook, before)
})

test('field members resolve configured local IDs and retain the shared raw-value and blank-key semantics', () => {
  const workbook = pivotChartEngineWorkbook(), source = workbook.sheets[0]
  assert.deepEqual(getPivotChartFieldItems(workbook, first(workbook), 'region-local'), ['North', 'South', 'West'])
  for (const [index, value] of [null, false, '', false].entries()) source.cells[`${index + 3}:7`] = createCell(index + 3, 7, value)
  assert.deepEqual(getPivotSourceFieldItems(workbook, effective(workbook), 'score-source'), [null, false, ''])
  assert.deepEqual(getPivotChartFieldItems(workbook, first(workbook), 'score-source'), ['', 'false'])
  const before = clone(workbook)
  assert.throws(() => getPivotChartFieldItems(workbook, first(workbook), 'missing-field'), invalidEdit); assert.deepEqual(workbook, before)
})

test('one full field configuration produces real filtered aggregation and one final revision without worksheet output', () => {
  const workbook = pivotChartEngineWorkbook({ configured: false }), before = cells(workbook), report = effective(workbook), revision = report.revision
  const fields = pivotChartFields(); fields[0].selectedItems = ['North']; fields[0].sort = 'descending'; fields[1].format = '0.00'
  assert.equal(configure(workbook, fields), true)
  assert.equal(report.revision, revision + 1); assert.equal(first(workbook).revision, 1)
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['North', 40]])
  assert.equal(calculatePivotReport(workbook, report).numberFormats['1:1'], '0.00'); assert.deepEqual(cells(workbook), before)
  assert.equal(Object.hasOwn(report, 'outputRange'), false)
})

test('owned data and ordinary worksheet output use the same grouping and aggregation result', () => {
  const owned = pivotChartEngineWorkbook(), linked = pivotChartEngineWorkbook({ ownership: 'linked' })
  assert.deepEqual(matrix(owned), pivotOutputValues(linked, effective(linked)))
  const fields = pivotChartFields(); fields[0].manualGroups = [{ name: 'Outer', members: ['North', 'West'] }]
  configure(owned, fields); configure(linked, fields)
  assert.deepEqual(matrix(owned), pivotOutputValues(linked, effective(linked)))
  assert.deepEqual(matrix(owned), [['Region', 'Amount'], ['Outer', 80], ['South', 20]])
})

test.each([
  ['sum', 100], ['average', 25], ['count', 4], ['countNums', 4], ['min', 10], ['max', 40], ['product', 240000],
  ['stdDev', Math.sqrt(500 / 3)], ['stdDevp', Math.sqrt(125)], ['var', 500 / 3], ['varp', 125],
])('shared %s aggregation remains available to owned report values', (aggregation, expected) => {
  const workbook = pivotChartEngineWorkbook(), fields = pivotChartFields()
  fields[0].manualGroups = [{ name: 'All', members: ['North', 'South', 'West'] }]; fields[1].aggregation = aggregation
  const before = cells(workbook); configure(workbook, fields)
  assert.deepEqual(matrix(workbook)[0], ['Region', 'Amount']); assert.equal(matrix(workbook)[1][0], 'All')
  assert.ok(Math.abs(matrix(workbook)[1][1] - expected) < 1e-10); assert.deepEqual(cells(workbook), before)
})

test('distinct local instances of one value source calculate independently and require explicit disambiguation', () => {
  const workbook = pivotChartEngineWorkbook(), fields = pivotChartFields()
  fields[1].displayName = 'Total'; fields.push({ ...clone(fields[1]), id: 'amount-average', position: 1, aggregation: 'average', displayName: 'Average' })
  configure(workbook, fields)
  assert.deepEqual(matrix(workbook), [['Region', 'Total', 'Average'], ['North', 40, 20], ['South', 20, 20], ['West', 40, 40]])
  const before = clone(workbook)
  assert.throws(() => edit(workbook, { type: 'setFieldSetting', fieldId: 'amount-source', setting: { subtotalType: 'max' } }), invalidEdit)
  assert.deepEqual(workbook, before)
  const revision = effective(workbook).revision
  edit(workbook, { type: 'moveField', fieldId: 'amount-average', fieldArea: 'value', index: 0 })
  assert.equal(effective(workbook).revision, revision + 1)
  assert.deepEqual(matrix(workbook), [['Region', 'Average', 'Total'], ['North', 20, 40], ['South', 20, 20], ['West', 40, 40]])
  assert.equal(effective(workbook).fields.find(field => field.id === 'amount-local').position, 1)
})

test('canonical IDs take precedence over an unrelated local-ID collision without introducing duplicate instances', () => {
  const workbook = pivotChartEngineWorkbook(), fields = pivotChartFields(); fields[1].id = 'segment-source'; configure(workbook, fields)
  edit(workbook, { type: 'addField', dataFieldId: 'segment-source', fieldArea: 'row', index: 1 })
  const report = effective(workbook), segment = report.fields.find(field => field.sourceFieldId === 'segment-source')
  assert.notEqual(segment.id, 'segment-source'); assert.equal(new Set(report.fields.map(field => field.id)).size, report.fields.length)
  assert.deepEqual(getPivotChartFieldItems(workbook, first(workbook), 'segment-source'), ['A', 'B'])
  edit(workbook, { type: 'removeFields', fieldIds: ['segment-source'] })
  assert.equal(report.fields.some(field => field.sourceFieldId === 'segment-source'), false)
  assert.equal(report.fields.find(field => field.id === 'segment-source').sourceFieldId, 'amount-source')
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['North', 40], ['South', 20], ['West', 40]])
})

test('numeric source indices resolve stable cache IDs and support initially empty authoring', () => {
  const workbook = pivotChartEngineWorkbook({ configured: false }), before = cells(workbook)
  edit(workbook, { type: 'addField', dataFieldId: 0, fieldArea: 'row', index: 0 })
  assert.deepEqual(matrix(workbook), []); edit(workbook, { type: 'addField', dataFieldId: 1, fieldArea: 'value', index: 0 })
  assert.equal(effective(workbook).fields[0].sourceFieldId, 'region-source')
  assert.equal(effective(workbook).fields[1].sourceFieldId, 'amount-source'); assert.deepEqual(cells(workbook), before)
  assert.equal(matrix(workbook)[1][1], 40)
})

test('local instance group keys also preserve ordinary worksheet drill-down to the actual source rows', () => {
  const workbook = pivotChartEngineWorkbook({ ownership: 'linked' }), report = effective(workbook)
  assert.deepEqual(pivotDrillDown(workbook, report, 1, 1), { headers: ['Region', 'Amount', 'Date', 'Segment', 'Score'],
    rows: [['North', 10, 45292, 'A', 5], ['North', 30, 45352, 'B', 25]] })
})

test('numeric grouping, raw-member filtering, and clearing share the existing report engine', () => {
  const workbook = pivotChartEngineWorkbook(), value = pivotChartFields()[1]
  configure(workbook, [dimension('score-source', 7, 'Score'), value])
  edit(workbook, { type: 'setNumberGroup', fieldId: 'score-source', group: { start: 0, end: 40, interval: 20 } })
  assert.deepEqual(matrix(workbook), [['Score', 'Amount'], ['0-20', 30], ['20-40', 70]])
  edit(workbook, { type: 'setFilter', fieldId: 'score-source', items: ['5'] })
  assert.deepEqual(matrix(workbook), [['Score', 'Amount'], ['0-20', 10]])
  assert.deepEqual(getPivotChartFieldItems(workbook, first(workbook), 'score-source'), ['5', '15', '25', '35'])
  edit(workbook, { type: 'setFilter', fieldId: 'score-source', items: null }); edit(workbook, { type: 'setNumberGroup', fieldId: 'score-source', group: null })
  assert.deepEqual(matrix(workbook), [['Score', 'Amount'], [5, 10], [15, 20], [25, 30], [35, 40]])
})

test('date mode and date buckets use actual 1900-system values, with explicit Excel-like mode and clearing', () => {
  const workbook = pivotChartEngineWorkbook(); configure(workbook, [dimension('date-source', 5, 'Date'), pivotChartFields()[1]])
  const before = clone(workbook)
  assert.throws(() => edit(workbook, { type: 'setDateGroup', fieldId: 'date-source', dateType: 'month', excelLike: true }), invalidEdit)
  assert.deepEqual(workbook, before)
  edit(workbook, { type: 'setDateGroupingMode', mode: 'excelLike' })
  edit(workbook, { type: 'setDateGroup', fieldId: 'date-source', dateType: 'month', excelLike: true })
  assert.deepEqual(matrix(workbook), [['Date', 'Amount'], ['2024-01', 10], ['2024-02', 20], ['2024-03', 30], ['2024-04', 40]])
  edit(workbook, { type: 'setDateGroup', fieldId: 'date-source', dateType: 'year' })
  assert.deepEqual(matrix(workbook), [['Date', 'Amount'], ['2024', 100]])
  edit(workbook, { type: 'setDateGroup', fieldId: 'date-source', dateType: null })
  assert.equal(matrix(workbook)[1][0], 45292)
})

test('manual grouping and label sorting apply to real source members and clear without orphaned settings', () => {
  const workbook = pivotChartEngineWorkbook()
  edit(workbook, { type: 'setManualGroups', fieldId: 'region-source', groups: [{ name: 'Outer', members: ['North', 'West'] }] })
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['Outer', 80], ['South', 20]])
  edit(workbook, { type: 'setManualGroups', fieldId: 'region-source', groups: [] })
  edit(workbook, { type: 'setSort', fieldId: 'region-source', direction: 'descending' })
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['West', 40], ['South', 20], ['North', 40]])
  edit(workbook, { type: 'setSort', fieldId: 'region-source', direction: null })
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['North', 40], ['South', 20], ['West', 40]])
})

test('value filtering references canonical values, shares aggregation, and clearAllFilters restores source groups', () => {
  const workbook = pivotChartEngineWorkbook()
  edit(workbook, { type: 'setValueFilter', fieldId: 'region-source', filter: { operator: 'valueGreaterThan', expected: 30, valueFieldId: 'amount-source' } })
  assert.equal(effective(workbook).fields[0].valueFilter.valueFieldId, 'amount-local')
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['North', 40], ['West', 40]])
  edit(workbook, { type: 'setFilter', fieldId: 'region-source', items: ['North'] })
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['North', 40]])
  edit(workbook, { type: 'setFilter', fieldId: 'region-source', items: [], isAll: true, clearAllFilters: true })
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['North', 40], ['South', 20], ['West', 40]])
})

test('multiple values can move to rows while keeping each local aggregation and leaving source cells intact', () => {
  const workbook = pivotChartEngineWorkbook(), fields = pivotChartFields(); fields[1].displayName = 'Total'
  fields.push({ ...clone(fields[1]), id: 'amount-average', position: 1, aggregation: 'average', displayName: 'Average' })
  configure(workbook, fields); const before = cells(workbook)
  edit(workbook, { type: 'updateValuePosition', position: 'row' })
  assert.deepEqual(matrix(workbook).slice(0, 3), [['Region', 'Values', 'Value'], ['North', 'Total', 40], [null, 'Average', 20]])
  assert.deepEqual(cells(workbook), before)
})

test('equivalent options, settings, label sets, and reordered field arrays are content no-ops', () => {
  const workbook = pivotChartEngineWorkbook(); edit(workbook, { type: 'setFilter', fieldId: 'region-source', items: ['North', 'West'] })
  const before = clone(workbook), config = { fields: clone(effective(workbook).fields).reverse(), options: clone(effective(workbook).options) }
  config.fields.find(field => field.area === 'row').selectedItems.reverse()
  for (const value of [{ type: 'setFieldsConfig', config }, { type: 'setOptions', options: { valuePosition: 'column' } },
    { type: 'setFieldSetting', fieldId: 'amount-local', setting: { subtotalType: 'sum' } },
    { type: 'setFilter', fieldId: 'region-source', items: ['West', 'North'] },
    { type: 'moveField', fieldId: 'region-local', fieldArea: 'row', index: 0 }, { type: 'removeFields', fieldIds: ['missing'] }]) {
    assert.equal(edit(workbook, value), false); assert.deepEqual(workbook, before)
  }
})

test('explicit refresh recalculates current source values and publishes one revision', () => {
  const workbook = pivotChartEngineWorkbook(), report = effective(workbook), revision = report.revision, chartRevision = first(workbook).revision
  workbook.sheets[0].cells['3:4'].value = 15
  const before = cells(workbook); assert.equal(edit(workbook, { type: 'refresh' }), true)
  assert.equal(report.revision, revision + 1); assert.equal(first(workbook).revision, chartRevision + 1)
  assert.equal(matrix(workbook)[1][1], 45); assert.deepEqual(cells(workbook), before)
})

test('linked field edits update one owner and every linked source mirror while preserving per-chart visuals', () => {
  const workbook = pivotChartEngineWorkbook({ ownership: 'linked' }), report = effective(workbook), revision = report.revision
  workbook.charts[1].title = 'Independent second chart'; const chartRevisions = workbook.charts.map(chart => chart.revision), dataBefore = clone(workbook.sheets[0].cells)
  edit(workbook, { type: 'addField', dataFieldId: 'segment-source', fieldArea: 'column', index: 0 })
  assert.equal(report.revision, revision + 1); assert.equal(workbook.pivotTables.length, 1)
  workbook.charts.forEach((chart, index) => {
    assert.equal(getPivotChartReport(workbook, chart), report); assert.deepEqual(chart.source.range, report.outputRange)
    assert.equal(chart.source.sheetId, report.anchor.subUnitId); assert.equal(chart.revision, chartRevisions[index] + 1)
  })
  assert.equal(workbook.charts[1].title, 'Independent second chart'); assert.deepEqual(workbook.sheets[0].cells, dataBefore)
  assert.deepEqual(pivotOutputValues(workbook, report), matrix(workbook))
})

test('linked output conflicts preserve previous generated cells, field config, chart revisions, and source mirrors', () => {
  const workbook = pivotChartEngineWorkbook({ ownership: 'linked' }); workbook.sheets[1].cells['8:3'] = createCell(8, 3, 'Occupied expansion')
  const before = clone(workbook)
  assert.throws(() => edit(workbook, { type: 'addField', dataFieldId: 'segment-source', fieldArea: 'column', index: 0 }), PivotOutputConflictError)
  assert.deepEqual(workbook, before)
})

test.each(['owned', 'linked'])('%s connected label removal, reset, and conversion to values refuse atomically', ownership => {
  const workbook = pivotChartEngineWorkbook({ ownership, connected: true }), before = clone(workbook)
  for (const value of [{ type: 'removeFields', fieldIds: ['region-source'] }, { type: 'resetFields', area: 'row' },
    { type: 'moveField', fieldId: 'region-local', fieldArea: 'value', index: 1 }]) {
    assert.throws(() => edit(workbook, value), connectedEdit); assert.deepEqual(workbook, before)
  }
})

test('retained connections keep protecting a field when a sibling is stale or its control view is absent', () => {
  for (const absentView of [false, true]) {
    const workbook = pivotChartEngineWorkbook({ connected: true, staleSibling: true }); if (absentView) workbook.slicers = []
    const before = clone(workbook)
    assert.throws(() => edit(workbook, { type: 'resetFields', area: 'row' }), connectedEdit); assert.deepEqual(workbook, before)
  }
})

test('value fields used by surviving filters cannot be removed or converted, while resetting the entire report is valid', () => {
  const workbook = pivotChartEngineWorkbook()
  edit(workbook, { type: 'setValueFilter', fieldId: 'region-source', filter: { operator: 'valueGreaterThan', expected: 0, valueFieldId: 'amount-source' } })
  const before = clone(workbook)
  for (const value of [{ type: 'removeFields', fieldIds: ['amount-local'] }, { type: 'resetFields', area: 'value' },
    { type: 'moveField', fieldId: 'amount-local', fieldArea: 'row', index: 1 }]) {
    assert.throws(() => edit(workbook, value), invalidEdit); assert.deepEqual(workbook, before)
  }
  const revision = effective(workbook).revision; assert.equal(edit(workbook, { type: 'resetFields' }), true)
  assert.equal(effective(workbook).revision, revision + 1); assert.deepEqual(effective(workbook).fields, []); assert.deepEqual(matrix(workbook), [])
})

test('changing local identity while keeping a connected source label preserves the connection', () => {
  const workbook = pivotChartEngineWorkbook({ connected: true }), fields = pivotChartFields(), connections = clone(workbook.slicerCaches)
  fields[0].id = 'new-region-instance'; assert.equal(configure(workbook, fields), true)
  assert.equal(effective(workbook).fields[0].sourceFieldId, 'region-source'); assert.deepEqual(workbook.slicerCaches, connections)
  assert.deepEqual(matrix(workbook), [['Region', 'Amount'], ['North', 40], ['South', 20], ['West', 40]])
})

test.each([
  { type: 'addField', dataFieldId: -1, fieldArea: 'row' }, { type: 'moveField', fieldId: 'region-source', fieldArea: 'row', index: null },
  { type: 'setOptions', options: { valuePosition: '' } }, { type: 'updateValuePosition', position: 'column', index: -2 },
  { type: 'setFilter', fieldId: 'region-source', items: ['North'], partialItems: ['West'] },
  { type: 'setFieldSetting', fieldId: 'amount-local', setting: { showDataAs: { type: 'percentOfGrandTotal' } } },
  { type: 'setDateFilter', fieldId: 'region-source', filterInfo: {} }, { type: 'setCollapse', fieldId: 'region-source', collapse: true },
  { type: 'setCategoryLevel', fieldId: 'region-source' }, { type: 'setNumberGroup', fieldId: 'region-source', group: { start: 0, end: 10, interval: 0 } },
  { type: 'setManualGroups', fieldId: 'region-source', groups: [{ name: 'Bad', members: ['North', { unknown: true }] }] },
])('invalid or unsupported edit %# keeps report, cells, caches and charts unchanged', value => {
  const workbook = pivotChartEngineWorkbook(), before = clone(workbook)
  assert.throws(() => edit(workbook, value), invalidEdit); assert.deepEqual(workbook, before)
})

test('invalid full configurations reject duplicate IDs, source mismatches, positions and unsupported field semantics', () => {
  const workbook = pivotChartEngineWorkbook(), before = clone(workbook)
  const cases = [fields => { fields[1].id = fields[0].id }, fields => { fields[0].sourceColumn = 4 },
    fields => { fields[0].position = 2 }, fields => { fields[1].aggregation = 'median' },
    fields => { fields[0].showDataAs = { type: 'runningTotal' } }, fields => { fields[0].manualGroups = []; fields[0].dateGroupType = 'month' }]
  for (const mutate of cases) {
    const fields = pivotChartFields(); mutate(fields)
    assert.throws(() => configure(workbook, fields), invalidEdit); assert.deepEqual(workbook, before)
  }
})

test('shared data-only setters keep all worksheet cells intact through settings, filtering, grouping and placement', () => {
  const workbook = pivotChartEngineWorkbook(), report = effective(workbook), before = cells(workbook)
  const actions = [() => setPivotFieldSetting(workbook, report, 'amount-local', { subtotalType: 'average' }, 'data-only'),
    () => setPivotSort(workbook, report, 'region-local', 'descending', 'data-only'),
    () => setPivotLabelFilter(workbook, report, 'region-local', ['North'], 'data-only'),
    () => setPivotValueFilter(workbook, report, 'region-local', { operator: 'valueGreaterThan', expected: 0, valueFieldId: 'amount-local' }, 'data-only'),
    () => setPivotManualGroups(workbook, report, 'region-local', [{ name: 'Outer', members: ['North', 'West'] }], 'data-only'),
    () => setPivotNumberGroup(workbook, report, 'region-local', { start: 0, end: 40, interval: 20 }, 'data-only'),
    () => setPivotOptions(workbook, report, { repeatRowLabels: true }, 'data-only'),
    () => setPivotField(workbook, report, 'segment-source', 'filter', 0, 'data-only')]
  for (const action of actions) { action(); assert.deepEqual(cells(workbook), before); assert.equal(Object.hasOwn(report, 'generatedCellKeys'), false) }
})

test('invalid render and refresh modes fail before writing output or incrementing revisions', () => {
  const workbook = pivotChartEngineWorkbook(), report = effective(workbook), before = clone(workbook)
  assert.throws(() => refreshPivotTable(workbook, report, 'unknown'), TypeError)
  assert.throws(() => renderPivotTable(workbook, report, 'deferred'), TypeError); assert.deepEqual(workbook, before)
})

test('unavailable linked reports, sources and hosts report errors without replacing ownership', () => {
  for (const missing of ['report', 'source', 'host']) {
    const workbook = pivotChartEngineWorkbook({ ownership: 'linked' })
    if (missing === 'report') workbook.pivotTables = []
    else workbook.sheets = workbook.sheets.filter(sheet => sheet.id !== (missing === 'source' ? 'data' : 'reports'))
    const before = clone(workbook)
    assert.throws(() => edit(workbook, { type: 'setOptions', options: { repeatRowLabels: true } }), error => error instanceof PivotChartEditError && error.code === 'invalid-resource')
    assert.deepEqual(workbook, before); assert.equal(first(workbook).ownedPivot, undefined)
  }
})

test('duplicate report IDs, invalid sources and missing hosts refuse owned creation without mutating caches', () => {
  const workbook = pivotChartEngineWorkbook(), before = clone(workbook)
  assert.throws(() => createPivotChartReport(workbook, clone(pivotChartSourceInput), 'reports', 'owned-report'), PivotChartEditError)
  assert.throws(() => createPivotChartReport(workbook, { ...clone(pivotChartSourceInput), range: { ...pivotChartSourceInput.range, endRow: 200 } }, 'reports', 'bad-range'), RangeError)
  assert.throws(() => createPivotChartReport(workbook, clone(pivotChartSourceInput), 'missing', 'missing-host'))
  assert.deepEqual(workbook, before)
})

test('canonical snapshot reopening retains owned fields and cache identity without inventing output cells or reports', () => {
  const workbook = pivotChartEngineWorkbook(), fields = pivotChartFields(); fields[0].selectedItems = ['North']; fields[1].format = '0.00'; configure(workbook, fields)
  const saved = JSON.parse(JSON.stringify(workbook)), reopened = workbookFromInput(saved)
  assert.deepEqual(reopened.charts[0].ownedPivot.fields, saved.charts[0].ownedPivot.fields)
  assert.equal(reopened.charts[0].ownedPivot.pivotCacheId, 'engine-cache'); assert.deepEqual(reopened.pivotTables, [])
  assert.deepEqual(cells(reopened), cells(workbook)); assert.deepEqual(matrix(reopened), [['Region', 'Amount'], ['North', 40]])
  reopened.charts[0].ownedPivot.fields[0].selectedItems.push('West'); assert.deepEqual(workbook.charts[0].ownedPivot.fields[0].selectedItems, ['North'])
})
