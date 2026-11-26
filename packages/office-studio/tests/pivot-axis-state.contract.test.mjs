import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { createCell } from '../src/domain/model.ts'
import { calculatePivotReport, getPivotAxisState, getPivotCollapse, getPivotCollapseItems, pivotDrillDown, pivotOutputValues,
  refreshPivotTable, setPivotCategoryLevel, setPivotCollapse, setPivotCollapseInfo, setPivotShowDataAs, stepPivotCategoryLevel,
  PivotOutputConflictError } from '../src/features/pivot/services/pivot-engine.ts'
import { applyPivotChartEdit, getPivotChartAxisState, getPivotChartCollapseItems, getPivotChartReport, PivotChartEditError } from '../src/features/pivot/services/pivot-chart-engine.ts'
import { PIVOT_AXIS_RECORDS, pivotAxisFields, pivotAxisWorkbook } from './fixtures/pivot-axis-state-fixture.mjs'

// Deferred contracts only. Tests, type/build, Rust/WASM, browsers and benchmarks have not been run.
const clone = value => structuredClone(value)
const chart = workbook => workbook.charts[0]
const report = workbook => getPivotChartReport(workbook, chart(workbook))
const result = workbook => calculatePivotReport(workbook, report(workbook))
const edit = (workbook, value) => applyPivotChartEdit(workbook, chart(workbook), value)
const collapse = (workbook, fieldId, value = true, item) => edit(workbook, { type: 'setCollapse', fieldId, collapse: value, ...(item === undefined ? {} : { item }) })
const sameLabels = (actual, expected) => actual.length === expected.length && actual.every((value, index) => Object.is(value, expected[index]))
const dataAxis = axis => axis.filter(entry => entry.kind === 'detail' || entry.kind === 'collapsed')
const at = (workbook, rows, columns = [2024, 'Q1', 'Retail'], valueId = 'amount-value') => {
  const output = result(workbook), valueRows = report(workbook).options.valuePosition === 'row'
  const row = dataAxis(output.rowAxis).find(entry => sameLabels(entry.labels, rows) && (!valueRows || entry.valueFieldId === valueId))
  const column = dataAxis(output.columnAxis).find(entry => sameLabels(entry.labels, columns) && (valueRows || entry.valueFieldId === valueId))
  assert.ok(row && column, 'real output coordinate exists'); return output.values[row.index][column.index]
}
const close = (actual, expected) => typeof expected === 'number' ? assert.ok(Math.abs(actual - expected) < 1e-9, `${actual} ~= ${expected}`) : assert.equal(actual, expected)
const sheetCells = workbook => clone(workbook.sheets.map(sheet => ({ id: sheet.id, cells: sheet.cells })))
const invalid = error => error instanceof PivotChartEditError && error.code === 'invalid-edit'
const show = (type, baseFieldId = '', baseItem = '') => ({ type, baseFieldId, baseItem })

test('expanded output exposes original scalar identity, scopes and absolute matrix coordinates', () => {
  const workbook = pivotAxisWorkbook(), output = result(workbook)
  assert.equal(output.headerRows, 4); assert.equal(output.rowAxis.length, 8); assert.equal(output.columnAxis.length, 4)
  assert.equal(output.rowAxis[0].index, 4); assert.equal(output.columnAxis[0].index, 3)
  assert.deepEqual(output.rowAxis[0].labels, ['North', 'A', 'a1'])
  assert.deepEqual(output.rowAxis[0].fieldFilters, { 'region-label': 'North', 'product-label': 'A', 'item-label': 'a1' })
  assert.deepEqual(output.columnAxis[0].fieldMembers, { 'year-label': [2024], 'quarter-label': ['Q1'], 'channel-label': ['Retail'] })
  assert.equal(output.rowAxis[0].kind, 'detail'); assert.equal(output.rowAxis[0].fieldId, 'item-label'); close(at(workbook, ['North', 'A', 'a1']), 10)
})

test.each(['owned', 'linked'])('%s whole-field collapse replaces descendants with original parent aggregates', ownership => {
  const workbook = pivotAxisWorkbook({ ownership }), before = clone(workbook.sheets[0].cells), revision = report(workbook).revision
  assert.equal(collapse(workbook, 'region-data'), true); assert.equal(report(workbook).revision, revision + 1)
  const output = result(workbook); assert.deepEqual(dataAxis(output.rowAxis).map(entry => entry.labels), [['North'], ['South']])
  assert.equal(output.rowAxis.every(entry => entry.kind === 'collapsed' && entry.fieldId === 'region-label' && entry.level === 0), true)
  close(at(workbook, ['North']), 100); close(at(workbook, ['South']), 100)
  assert.deepEqual(workbook.sheets[0].cells, before)
  if (ownership === 'owned') assert.equal(report(workbook).axisOutput, undefined)
  else assert.deepEqual(report(workbook).axisOutput, { rowAxis: output.rowAxis, columnAxis: output.columnAxis })
})

test('one member collapses only its real branch while other parents keep their children', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'region-label', true, 'North')
  assert.deepEqual(dataAxis(result(workbook).rowAxis).map(entry => entry.labels), [['North'], ['South', 'C', 'c1'], ['South', 'C', 'c2'], ['South', 'D', 'd1'], ['South', 'D', 'd2']])
  close(at(workbook, ['North']), 100); close(at(workbook, ['South', 'C', 'c1']), 20)
  assert.equal(getPivotCollapse(report(workbook), 'region-label', 'North'), true)
  assert.equal(getPivotCollapse(report(workbook), 'region-label', 'South'), false)
})

test('expanding one member of a fully collapsed field retains every other cached member state', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'region-label'); collapse(workbook, 'region-label', false, 'North')
  assert.equal(getPivotCollapse(report(workbook), 'region-label', 'North'), false); assert.equal(getPivotCollapse(report(workbook), 'region-label', 'South'), true)
  assert.equal(result(workbook).rowAxis.filter(entry => entry.kind === 'detail').length, 4)
  close(at(workbook, ['North', 'A', 'a1']), 10); close(at(workbook, ['South']), 100)
  assert.equal(getPivotChartAxisState(workbook, chart(workbook)).mixed, true)
})

test('whole-field expansion restores descendants without changing source records or field instances', () => {
  const workbook = pivotAxisWorkbook(), original = result(workbook), fields = clone(report(workbook).fields), before = sheetCells(workbook)
  collapse(workbook, 'region-label'); collapse(workbook, 'region-label', false)
  assert.deepEqual(result(workbook), original); assert.deepEqual(report(workbook).fields, fields); assert.deepEqual(sheetCells(workbook), before)
  assert.equal(report(workbook).collapseInfo, undefined)
})

test('new source members distinguish whole-field state from a materialized cached-member expansion', () => {
  const whole = pivotAxisWorkbook(), materialized = pivotAxisWorkbook()
  collapse(whole, 'region-label'); collapse(materialized, 'region-label'); collapse(materialized, 'region-label', false, 'North')
  for (const workbook of [whole, materialized]) {
    const data = workbook.sheets[0]
    for (const cell of Object.values(data.cells)) if (cell.col === 2 && data.cells[`${cell.row}:4`]?.value === 'a2') cell.value = 'East'
    refreshPivotTable(workbook, report(workbook), 'data-only')
  }
  assert.equal(result(whole).rowAxis.some(entry => sameLabels(entry.labels, ['East']) && entry.kind === 'collapsed'), true)
  assert.equal(result(materialized).rowAxis.some(entry => sameLabels(entry.labels, ['East', 'A', 'a2']) && entry.kind === 'detail'), true)
  assert.equal(getPivotCollapse(report(materialized), 'region-label', 'South'), true)
})

test('field-member collapse applies at each ancestor containing that grouped member and retains each scope', () => {
  const records = clone(PIVOT_AXIS_RECORDS); records.forEach(row => { if (row[1] === 'C') row[1] = 'A' })
  const workbook = pivotAxisWorkbook({ records }); collapse(workbook, 'product-label', true, 'A')
  close(at(workbook, ['North', 'A']), 30); close(at(workbook, ['South', 'A']), 30)
  assert.equal(result(workbook).rowAxis.filter(entry => entry.kind === 'collapsed').length, 2)
  assert.deepEqual(result(workbook).rowAxis.find(entry => sameLabels(entry.labels, ['South', 'A'])).fieldFilters,
    { 'region-label': 'South', 'product-label': 'A' })
})

test('nested descendant states survive an ancestor collapse and become effective after expansion', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'product-label', true, 'A'); collapse(workbook, 'region-label', true, 'North')
  close(at(workbook, ['North']), 100); collapse(workbook, 'region-label', false, 'North')
  close(at(workbook, ['North', 'A']), 30); close(at(workbook, ['North', 'B', 'b1']), 30)
  assert.equal(getPivotCollapse(report(workbook), 'product-label', 'A'), true)
})

test('shared row and column collapse use the same original scope and preserve full header dimensions', () => {
  const workbook = pivotAxisWorkbook(), current = report(workbook)
  setPivotCollapse(workbook, current, 'region-label', true, 'North', 'data-only')
  setPivotCollapse(workbook, current, 'year-label', true, undefined, 'data-only')
  const output = result(workbook); close(at(workbook, ['North'], [2024]), 500)
  assert.equal(output.headerRows, 4); assert.equal(output.values[0][3], 2024); assert.equal(output.values[1][3], null); assert.equal(output.values[2][3], null)
  assert.equal(output.columnAxis[0].kind, 'collapsed'); assert.deepEqual(output.columnAxis[0].fieldFilters, { 'year-label': 2024 })
})

test('column member collapse coexists with row members and keeps uncapped column peers', () => {
  const workbook = pivotAxisWorkbook(), current = report(workbook)
  collapse(workbook, 'product-label', true, 'A'); setPivotCollapse(workbook, current, 'quarter-label', true, 'Q1', 'data-only')
  close(at(workbook, ['North', 'A'], [2024, 'Q1']), 45); close(at(workbook, ['North', 'A'], [2024, 'Q2', 'Retail']), 60)
  assert.deepEqual(dataAxis(result(workbook).columnAxis).map(entry => entry.labels), [[2024, 'Q1'], [2024, 'Q2', 'Retail'], [2024, 'Q2', 'Web']])
})

test.each(['column', 'row'])('multiple measures and %s values retain instance identity at collapsed intersections', valuePosition => {
  const workbook = pivotAxisWorkbook({ valuePosition, multipleValues: true }), current = report(workbook)
  collapse(workbook, 'region-label', true, 'North'); setPivotCollapse(workbook, current, 'year-label', true, undefined, 'data-only')
  close(at(workbook, ['North'], [2024]), 500); close(at(workbook, ['North'], [2024], 'amount-average'), 31.25)
  const output = result(workbook), measureAxis = valuePosition === 'row' ? output.rowAxis : output.columnAxis
  assert.equal(measureAxis.some(entry => entry.valueFieldId === 'amount-value'), true)
  assert.equal(measureAxis.some(entry => entry.valueFieldId === 'amount-average'), true)
})

test.each([['sum', 100], ['average', 25], ['count', 4], ['countNums', 4], ['max', 40], ['min', 10], ['product', 240000],
  ['stdDev', Math.sqrt(500 / 3)], ['stdDevp', Math.sqrt(125)], ['var', 500 / 3], ['varp', 125]])
('%s collapsed scopes use the original measure aggregation', (aggregation, expected) => {
  const workbook = pivotAxisWorkbook({ aggregation }); collapse(workbook, 'region-label', true, 'North'); close(at(workbook, ['North']), expected)
})

test.each([[show('percentOfGrandTotal'), 0.1], [show('percentOfRowTotal'), 0.2], [show('percentOfColumnTotal'), 0.5],
  [show('percentOfParentRowTotal'), 0.5], [show('percentOfParentColumnTotal'), 2 / 3], [show('index'), 1],
  [show('percentOfParent', 'year-label'), 0.2], [show('runningTotal', 'region-label'), 100],
  [show('percentRunningTotal', 'region-label'), 0.5], [show('rankAscending', 'region-label'), 1], [show('rankDescending', 'region-label'), 1],
  [show('percentOf', 'product-label', 'B'), null], [show('rankAscending', 'product-label'), null]])
('Show As %# is derived after collapsing the original scope, including hidden-member policies', (rule, expected) => {
  const workbook = pivotAxisWorkbook(); setPivotShowDataAs(workbook, report(workbook), 'amount-value', rule, 'data-only')
  collapse(workbook, 'region-label', true, 'North'); close(at(workbook, ['North']), expected)
})

test('collapsed parents are real data entries even when all subtotal outputs are hidden', () => {
  const workbook = pivotAxisWorkbook({ totals: false }); collapse(workbook, 'product-label')
  const output = result(workbook); assert.equal(output.rowAxis.length, 4); assert.equal(output.rowAxis.every(entry => entry.kind === 'collapsed'), true)
  assert.equal(output.values.some(row => row.some(value => typeof value === 'string' && value.endsWith(' Total'))), false)
  close(at(workbook, ['North', 'B']), 70)
})

test('visible subtotals remain distinct and a collapsed node never gets a duplicate subtotal', () => {
  const workbook = pivotAxisWorkbook({ totals: true }); collapse(workbook, 'product-label', true, 'A')
  const output = result(workbook), node = output.rowAxis.filter(entry => sameLabels(entry.labels, ['North', 'A']))
  assert.equal(node.length, 1); assert.equal(node[0].kind, 'collapsed')
  const north = output.rowAxis.find(entry => entry.kind === 'subtotal' && sameLabels(entry.labels, ['North']))
  const q1 = output.columnAxis.find(entry => entry.kind === 'detail' && sameLabels(entry.labels, [2024, 'Q1', 'Retail']))
  close(output.values[north.index][q1.index], 100)
  assert.equal(output.rowAxis.some(entry => entry.kind === 'grand-total'), true)
})

test('sorting and repeat-label options change presentation while metadata keeps complete original identities', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'region-label', true, 'North')
  edit(workbook, { type: 'setSort', fieldId: 'region-label', direction: 'descending' })
  edit(workbook, { type: 'setOptions', options: { repeatRowLabels: false } })
  const output = result(workbook); assert.deepEqual(output.rowAxis[0].labels, ['South', 'C', 'c1'])
  assert.deepEqual(output.rowAxis[1].labels, ['South', 'C', 'c2']); assert.equal(output.values[output.rowAxis[1].index][0], null)
  close(at(workbook, ['North']), 100)
})

test('equivalent field/member state changes and untouched levels do not publish or increment revisions', () => {
  const workbook = pivotAxisWorkbook({ ownership: 'linked' }), initial = clone(workbook)
  assert.equal(collapse(workbook, 'region-label', false), false); assert.equal(collapse(workbook, 'region-label', false, 'North'), false)
  assert.equal(edit(workbook, { type: 'setCategoryLevel', fieldId: 'item-label' }), false); assert.deepEqual(workbook, initial)
  collapse(workbook, 'region-label', true, 'North'); const before = clone(workbook)
  assert.equal(collapse(workbook, 'region-data', true, 'North'), false); assert.deepEqual(workbook, before)
  const read = getPivotChartAxisState(workbook, chart(workbook)); read.fields[0].name = 'Changed'
  assert.equal(getPivotChartAxisState(workbook, chart(workbook)).fields[0].name, 'Region')
})

test('restored explicit expansion markers are equivalent to the default state for no-op edits', () => {
  const workbook = pivotAxisWorkbook(), current = report(workbook)
  current.collapseInfo = { 'region-label': false, 'product-label': { A: false } }; const before = clone(workbook), revision = current.revision
  setPivotCollapse(workbook, current, 'region-label', false, undefined, 'data-only')
  assert.equal(current.revision, revision); assert.equal(collapse(workbook, 'region-label', false), false)
  assert.equal(edit(workbook, { type: 'setFieldsConfig', config: { fields: clone(current.fields), collapseInfo: null } }), false)
  assert.deepEqual(workbook, before)
})

test.each([['region-label', 2, 0], ['product-label', 4, 1], ['item-label', 8, 2]])
('category level %s sets actual row depth and preserves column state', (fieldId, rows, level) => {
  const workbook = pivotAxisWorkbook(), current = report(workbook)
  setPivotCollapse(workbook, current, 'year-label', true, undefined, 'data-only')
  edit(workbook, { type: 'setCategoryLevel', fieldId }); const output = result(workbook), state = getPivotChartAxisState(workbook, chart(workbook))
  assert.equal(output.rowAxis.length, rows); assert.equal(state.categoryLevelId, fieldId); assert.equal(state.categoryLevelIndex, level); assert.equal(state.mixed, false)
  assert.equal(getPivotCollapse(current, 'year-label'), true); assert.equal(output.columnAxis[0].kind, 'collapsed')
})

test('setting a complete category level clears member overrides and leaves deeper source data recoverable', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'region-label', true, 'North')
  edit(workbook, { type: 'setCategoryLevel', fieldId: 'product-data' }); close(at(workbook, ['North', 'A']), 30); close(at(workbook, ['South', 'D']), 70)
  assert.equal(getPivotCollapse(report(workbook), 'region-label'), false); assert.equal(getPivotCollapse(report(workbook), 'product-label'), true)
  edit(workbook, { type: 'setCategoryLevel', fieldId: 'item-label' }); close(at(workbook, ['North', 'A', 'a1']), 10)
})

test('category step chooses the fixed first collapsed or last noncollapsed parent and stops at the boundary', () => {
  const workbook = pivotAxisWorkbook(), current = report(workbook), revision = current.revision
  assert.equal(edit(workbook, { type: 'stepCategoryLevel', direction: 'collapse' }), true); assert.equal(getPivotCollapse(current, 'product-label'), true)
  assert.equal(edit(workbook, { type: 'stepCategoryLevel', direction: 'collapse' }), true); assert.equal(getPivotCollapse(current, 'region-label'), true)
  assert.equal(edit(workbook, { type: 'stepCategoryLevel', direction: 'collapse' }), false)
  assert.equal(edit(workbook, { type: 'stepCategoryLevel', direction: 'expand' }), true); assert.equal(getPivotCollapse(current, 'region-label'), false)
  assert.equal(edit(workbook, { type: 'stepCategoryLevel', direction: 'expand' }), true); assert.equal(getPivotCollapse(current, 'product-label'), false)
  assert.equal(edit(workbook, { type: 'stepCategoryLevel', direction: 'expand' }), false); assert.equal(current.revision, revision + 4)
})

test('a single row level is a valid display depth and has no fold or step capability', () => {
  const workbook = pivotAxisWorkbook(), fields = pivotAxisFields().filter(field => field.id !== 'product-label' && field.id !== 'item-label')
  edit(workbook, { type: 'setFieldsConfig', config: { fields } }); const before = clone(workbook)
  assert.equal(edit(workbook, { type: 'setCategoryLevel', fieldId: 'region-label' }), false)
  assert.equal(edit(workbook, { type: 'stepCategoryLevel', direction: 'expand' }), false)
  assert.equal(edit(workbook, { type: 'stepCategoryLevel', direction: 'collapse' }), false)
  assert.throws(() => collapse(workbook, 'region-label'), invalid); assert.deepEqual(workbook, before)
  const state = getPivotChartAxisState(workbook, chart(workbook)); assert.equal(state.canExpand, false); assert.equal(state.canCollapse, false)
  close(at(workbook, ['North']), 100)
})

test.each(['worksheet', 'data-only', 'deferred'])('shared %s strategy retains identical hierarchy computation', mode => {
  const workbook = pivotAxisWorkbook({ ownership: mode === 'worksheet' ? 'linked' : 'owned' }), current = report(workbook), before = sheetCells(workbook), revision = current.revision
  setPivotCategoryLevel(workbook, current, 'region-label', mode); assert.equal(current.revision, revision + 1); close(at(workbook, ['North']), 100)
  if (mode === 'worksheet') { assert.notDeepEqual(sheetCells(workbook), before); assert.deepEqual(pivotOutputValues(workbook, current), result(workbook).values) }
  else { assert.deepEqual(sheetCells(workbook), before); assert.equal(current.axisOutput, undefined) }
  assert.equal(stepPivotCategoryLevel(workbook, current, 'collapse', mode), false)
})

test('invalid refresh strategies fail before mutating any hierarchy state, including a boundary no-op', () => {
  const workbook = pivotAxisWorkbook(), current = report(workbook), before = clone(workbook)
  for (const operation of [() => setPivotCollapse(workbook, current, 'region-label', true, undefined, 'unknown'),
    () => setPivotCollapseInfo(workbook, current, { 'region-label': true }, 'unknown'),
    () => setPivotCategoryLevel(workbook, current, 'region-label', 'unknown'),
    () => stepPivotCategoryLevel(workbook, current, 'expand', 'unknown')]) {
    assert.throws(operation); assert.deepEqual(workbook, before)
  }
})

test('worksheet retained axis metadata belongs to the same published report and linked charts revise once', () => {
  const workbook = pivotAxisWorkbook({ ownership: 'linked' }), current = report(workbook), revisions = workbook.charts.map(chart => chart.revision)
  collapse(workbook, 'product-label', true, 'A'); const output = result(workbook)
  assert.deepEqual(current.axisOutput, { rowAxis: output.rowAxis, columnAxis: output.columnAxis })
  workbook.charts.forEach((chart, index) => { assert.equal(chart.revision, revisions[index] + 1); assert.deepEqual(chart.source.range, current.outputRange) })
  assert.equal(workbook.charts[1].title, 'Other visual')
})

test('collapsed drilldown uses the untouched underlying source scope rather than transformed display values', () => {
  const workbook = pivotAxisWorkbook({ ownership: 'linked' }); collapse(workbook, 'product-label', true, 'A')
  const output = result(workbook), row = output.rowAxis.find(entry => sameLabels(entry.labels, ['North', 'A'])), column = output.columnAxis[0]
  const detail = pivotDrillDown(workbook, report(workbook), row.index, column.index)
  assert.deepEqual(detail.rows, [['North', 'A', 'a1', 2024, 'Q1', 'Retail', 10], ['North', 'A', 'a2', 2024, 'Q1', 'Retail', 20]])
})

test('manual grouped collapse queries actual grouped keys and keeps raw source members for drilldown', () => {
  const workbook = pivotAxisWorkbook({ ownership: 'linked' })
  edit(workbook, { type: 'setManualGroups', fieldId: 'product-label', groups: [{ name: 'Mix', members: ['A', 'B'] }] })
  assert.deepEqual(getPivotChartCollapseItems(workbook, chart(workbook), 'product-data'), ['Mix', 'C', 'D'])
  collapse(workbook, 'product-label', true, 'Mix'); close(at(workbook, ['North', 'Mix']), 100)
  const output = result(workbook), row = output.rowAxis.find(entry => sameLabels(entry.labels, ['North', 'Mix']))
  assert.deepEqual(row.fieldMembers['product-label'], ['A', 'B'])
  assert.equal(pivotDrillDown(workbook, report(workbook), row.index, output.columnAxis[0].index).rows.length, 4)
})

test('numeric grouped members can collapse real buckets without substituting raw labels', () => {
  const records = clone(PIVOT_AXIS_RECORDS); records.forEach(row => { row[0] = row[0] === 'North' ? 10 : 30 })
  const workbook = pivotAxisWorkbook({ records }); edit(workbook, { type: 'setNumberGroup', fieldId: 'region-label', group: { start: 0, end: 50, interval: 25 } })
  assert.deepEqual(getPivotCollapseItems(workbook, report(workbook), 'region-label'), ['0-25', '25-50'])
  collapse(workbook, 'region-label', true, '0-25'); close(at(workbook, ['0-25']), 100)
  assert.deepEqual(result(workbook).rowAxis[0].fieldMembers['region-label'], [10])
})

test('date grouped column collapse retains the real date interval and protects active grouped identities', () => {
  const records = clone(PIVOT_AXIS_RECORDS); records.forEach(row => { row[3] = row[4] === 'Q1' ? 45292 : 45383 })
  const workbook = pivotAxisWorkbook({ records }), current = report(workbook)
  edit(workbook, { type: 'setDateGroup', fieldId: 'year-label', dateType: 'month' })
  assert.deepEqual(getPivotCollapseItems(workbook, current, 'year-label'), ['2024-01', '2024-04'])
  setPivotCollapse(workbook, current, 'year-label', true, '2024-01', 'data-only'); close(at(workbook, ['North', 'A', 'a1'], ['2024-01']), 15)
  const before = clone(workbook); assert.throws(() => edit(workbook, { type: 'setDateGroup', fieldId: 'year-label', dateType: 'year' }), invalid); assert.deepEqual(workbook, before)
})

test('grouping that detaches an active member rejects, while one complete config can remap it atomically', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'product-label', true, 'A'); const before = clone(workbook)
  assert.throws(() => edit(workbook, { type: 'setManualGroups', fieldId: 'product-label', groups: [{ name: 'Mix', members: ['A', 'B'] }] }), invalid)
  assert.deepEqual(workbook, before)
  const fields = clone(report(workbook).fields); fields[1].manualGroups = [{ name: 'Mix', members: ['A', 'B'] }]
  const revision = report(workbook).revision
  edit(workbook, { type: 'setFieldsConfig', config: { fields, collapseInfo: { 'product-label': { Mix: true } } } })
  assert.equal(report(workbook).revision, revision + 1); close(at(workbook, ['North', 'Mix']), 100)
})

test('whole-field collapse survives grouping because it has no individual member dependency', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'product-label')
  edit(workbook, { type: 'setManualGroups', fieldId: 'product-label', groups: [{ name: 'Mix', members: ['A', 'B'] }] })
  assert.equal(getPivotCollapse(report(workbook), 'product-label'), true); close(at(workbook, ['North', 'Mix']), 100)
})

test('removal retires now-leaf hierarchy state and reset clears every attached collapse resource', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'product-label', true, 'A'); collapse(workbook, 'region-label')
  edit(workbook, { type: 'removeFields', fieldIds: ['item-data'] })
  assert.equal(report(workbook).collapseInfo['product-label'], undefined); assert.equal(getPivotCollapse(report(workbook), 'region-label'), true)
  assert.equal(getPivotChartAxisState(workbook, chart(workbook)).fields[1].canCollapse, false)
  const before = clone(workbook); assert.throws(() => collapse(workbook, 'product-label'), invalid); assert.deepEqual(workbook, before)
  edit(workbook, { type: 'resetFields' }); assert.equal(report(workbook).collapseInfo, undefined)
  assert.deepEqual(result(workbook).rowAxis, []); assert.deepEqual(result(workbook).columnAxis, [])
})

test('moving a parent to the other axis retains identity when it has descendants and moving to filter retires it', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'region-label')
  edit(workbook, { type: 'moveField', fieldId: 'region-label', fieldArea: 'column', index: 0 })
  assert.equal(getPivotCollapse(report(workbook), 'region-label'), true)
  assert.equal(result(workbook).columnAxis.every(entry => entry.kind === 'collapsed' && entry.fieldId === 'region-label'), true)
  edit(workbook, { type: 'moveField', fieldId: 'region-label', fieldArea: 'filter', index: 0 })
  assert.equal(report(workbook).collapseInfo, undefined)
})

test('canonical command targets and stored report-local IDs remain separate under a source-ID collision', () => {
  const workbook = pivotAxisWorkbook(), fields = pivotAxisFields(); fields[0].id = 'product-data'
  edit(workbook, { type: 'setFieldsConfig', config: { fields, collapseInfo: { 'product-data': true } } })
  assert.equal(getPivotAxisState(workbook, report(workbook)).fields[0].collapsed, true); close(at(workbook, ['North']), 100)
  collapse(workbook, 'product-data'); assert.equal(getPivotCollapse(report(workbook), 'product-label'), true)
  assert.equal(getPivotCollapse(report(workbook), 'product-data'), true)
})

test('ambiguous repeated source targets require a report instance ID', () => {
  const workbook = pivotAxisWorkbook(), fields = pivotAxisFields()
  fields[2] = { ...clone(fields[1]), id: 'product-second', position: 2 }
  edit(workbook, { type: 'setFieldsConfig', config: { fields } }); const before = clone(workbook)
  assert.throws(() => collapse(workbook, 'product-data'), invalid); assert.deepEqual(workbook, before)
  assert.equal(collapse(workbook, 'product-label', true, 'A'), true)
})

test.each([
  { type: 'setCollapse', fieldId: 'item-label', collapse: true }, { type: 'setCollapse', fieldId: 'amount-value', collapse: true },
  { type: 'setCollapse', fieldId: 'year-label', collapse: true }, { type: 'setCollapse', fieldId: 'missing', collapse: true },
  { type: 'setCollapse', fieldId: 'region-label', collapse: 1 }, { type: 'setCollapse', fieldId: 'region-label', collapse: true, item: 1 },
  { type: 'setCollapse', fieldId: 'region-label', collapse: true, item: 'missing' }, { type: 'setCollapse', fieldId: 'region-label' },
  { type: 'setCollapse', fieldId: 'region-label', collapse: true, unknown: true }, { type: 'setCategoryLevel', fieldId: 'year-label' },
  { type: 'setCategoryLevel', fieldId: 'missing' }, { type: 'stepCategoryLevel', direction: 'sideways' },
  { type: 'setFieldsConfig', config: { fields: pivotAxisFields(), collapseInfo: { 'item-label': true } } },
  { type: 'setFieldsConfig', config: { fields: pivotAxisFields(), collapseInfo: { 'region-label': 1 } } },
  { type: 'setFieldsConfig', config: { fields: pivotAxisFields(), collapseInfo: { 'region-label': { North: 'true' } } } },
  { type: 'setFieldsConfig', config: { fields: pivotAxisFields(), collapseInfo: { missing: true } } },
])('invalid hierarchy DTO %# refuses before publishing content or output', value => {
  const workbook = pivotAxisWorkbook(), before = clone(workbook); assert.throws(() => edit(workbook, value), invalid); assert.deepEqual(workbook, before)
})

test('restored unknown or leaf collapse state is diagnosed instead of silently ignored by calculation', () => {
  const workbook = pivotAxisWorkbook(), current = report(workbook)
  current.collapseInfo = { missing: true }; assert.throws(() => result(workbook))
  current.collapseInfo = { 'item-label': true }; assert.throws(() => result(workbook))
  current.collapseInfo = { 'region-label': { North: 1 } }; assert.throws(() => result(workbook))
})

test('scalar key collisions reject member edits while whole fields preserve both typed branches', () => {
  const records = clone(PIVOT_AXIS_RECORDS); records.forEach(row => { row[0] = row[0] === 'North' ? 1 : '1' })
  const workbook = pivotAxisWorkbook({ records }), before = clone(workbook)
  assert.throws(() => collapse(workbook, 'region-label', true, '1'), invalid); assert.deepEqual(workbook, before)
  collapse(workbook, 'region-label'); close(at(workbook, [1]), 100); close(at(workbook, ['1']), 100)
})

test('prototype-like field and member keys remain ordinary own entries in collapse and aggregate scopes', () => {
  const records = clone(PIVOT_AXIS_RECORDS); records.forEach(row => { if (row[0] === 'North') row[0] = '__proto__' })
  const workbook = pivotAxisWorkbook({ records }), fields = pivotAxisFields(); fields[0].id = '__proto__'
  edit(workbook, { type: 'setFieldsConfig', config: { fields } }); collapse(workbook, '__proto__', true, '__proto__')
  assert.equal(Object.hasOwn(report(workbook).collapseInfo, '__proto__'), true)
  assert.equal(Object.hasOwn(report(workbook).collapseInfo.__proto__, '__proto__'), true)
  close(at(workbook, ['__proto__']), 100)
  const entry = result(workbook).rowAxis.find(entry => sameLabels(entry.labels, ['__proto__']))
  assert.equal(Object.hasOwn(entry.fieldFilters, '__proto__'), true); assert.equal(entry.fieldFilters.__proto__, '__proto__')
})

test('filtered or changed source members keep dormant collapse identity and recover when the member returns', () => {
  const workbook = pivotAxisWorkbook(); collapse(workbook, 'region-label', true, 'North')
  edit(workbook, { type: 'setFilter', fieldId: 'region-label', items: ['South'] })
  assert.equal(result(workbook).rowAxis.every(entry => entry.labels[0] === 'South' && entry.kind === 'detail'), true)
  assert.equal(getPivotCollapse(report(workbook), 'region-label', 'North'), true)
  edit(workbook, { type: 'setFilter', fieldId: 'region-label', items: null }); close(at(workbook, ['North']), 100)
  for (const cell of Object.values(workbook.sheets[0].cells)) if (cell.col === 2 && cell.value === 'North') cell.value = 'East'
  refreshPivotTable(workbook, report(workbook), 'data-only'); assert.equal(result(workbook).rowAxis.some(entry => entry.labels[0] === 'East' && entry.kind === 'detail'), true)
  for (const cell of Object.values(workbook.sheets[0].cells)) if (cell.col === 2 && cell.value === 'East') cell.value = 'North'
  refreshPivotTable(workbook, report(workbook), 'data-only'); close(at(workbook, ['North']), 100)
})

test('linked expansion conflict preserves collapsed output, retained metadata and all chart mirrors', () => {
  const workbook = pivotAxisWorkbook({ ownership: 'linked' }); collapse(workbook, 'region-label')
  workbook.sheets[1].cells['26:5'] = createCell(26, 5, 'Expansion conflict'); const before = clone(workbook)
  assert.throws(() => collapse(workbook, 'region-label', false), PivotOutputConflictError); assert.deepEqual(workbook, before)
})

test('full collapse state normalization ignores member key order and explicit reset restores the source hierarchy', () => {
  const workbook = pivotAxisWorkbook(), current = report(workbook)
  setPivotCollapseInfo(workbook, current, { 'region-label': { South: true, North: true } }, 'data-only')
  const revision = current.revision; setPivotCollapseInfo(workbook, current, { 'region-label': { North: true, South: true } }, 'data-only')
  assert.equal(current.revision, revision); setPivotCollapseInfo(workbook, current, null, 'data-only')
  assert.equal(result(workbook).rowAxis.length, 8); assert.equal(current.collapseInfo, undefined)
})

test('snapshot reopening keeps independent persistent member state and derived worksheet metadata without aliases', () => {
  const workbook = pivotAxisWorkbook({ ownership: 'linked' }); collapse(workbook, 'region-label', true, 'North')
  const reopened = workbookFromInput(JSON.parse(JSON.stringify(workbook)))
  close(at(reopened, ['North']), 100); assert.deepEqual(report(reopened).axisOutput, report(workbook).axisOutput)
  assert.equal(getPivotCollapse(report(reopened), 'region-label', 'North'), true)
  collapse(reopened, 'region-label', false, 'North'); assert.equal(getPivotCollapse(report(workbook), 'region-label', 'North'), true)
})
