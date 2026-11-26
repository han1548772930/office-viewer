import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { createCell } from '../src/domain/model.ts'
import { calculatePivotReport, getPivotShowDataAs, getPivotShowDataAsBaseItems, pivotDrillDown, pivotOutputValues, refreshPivotTable,
  setPivotFieldSetting, setPivotShowDataAs, PivotOutputConflictError } from '../src/features/pivot/services/pivot-engine.ts'
import { applyPivotChartEdit, getPivotChartConfiguredField, getPivotChartFieldItems, getPivotChartReport,
  getPivotChartShowDataAsBaseItems, PivotChartEditError } from '../src/features/pivot/services/pivot-chart-engine.ts'
import { getPivotShowDataAsBaseFieldIds, normalizePivotShowDataAs, PIVOT_SHOW_DATA_AS_RULES } from '../src/features/pivot/services/pivot-show-data-as.ts'
import { PIVOT_SHOW_DATA_AS_RECORDS, pivotShowDataAsFields, pivotShowDataAsWorkbook } from './fixtures/pivot-show-data-as-fixture.mjs'

// Deferred contracts only. No tests, type/build, browser, WASM or benchmarks have been run.
const clone = value => structuredClone(value)
const chart = workbook => workbook.charts[0]
const report = workbook => getPivotChartReport(workbook, chart(workbook))
const calculation = workbook => calculatePivotReport(workbook, report(workbook))
const matrix = workbook => calculation(workbook).values
const cells = workbook => clone(workbook.sheets.map(sheet => ({ id: sheet.id, cells: sheet.cells })))
const rule = (type, baseFieldId = '', baseItem = '', baseItemType) => ({ type, baseFieldId, baseItem, ...(baseItemType ? { baseItemType } : {}) })
const set = (workbook, value, fieldId = 'amount-value') => applyPivotChartEdit(workbook, chart(workbook), { type: 'setShowDataAs', fieldId, showDataAs: value })
const edit = (workbook, value) => applyPivotChartEdit(workbook, chart(workbook), value)
const valueAt = (workbook, labels = ['North', 'A'], column = [2024, 'Q1'], valueName = 'Amount') => {
  const output = matrix(workbook), row = output.slice(3).find(row => labels.every((label, index) => Object.is(row[index], label)))
  const index = output[0].findIndex((label, index) => index >= 2 && Object.is(label, column[0]) && Object.is(output[1][index], column[1]) && output[2][index] === valueName)
  assert.ok(row && index >= 0, 'real report coordinate exists'); return row[index]
}
const close = (actual, expected) => typeof expected === 'number' ? assert.ok(Math.abs(actual - expected) < 1e-10, `${actual} ~= ${expected}`) : assert.equal(actual, expected)
const invalid = error => error instanceof PivotChartEditError && error.code === 'invalid-edit'

test('fixed Show Values As inventory is fifteen rules, plus one separately marked historical compatibility rule', () => {
  assert.equal(PIVOT_SHOW_DATA_AS_RULES.filter(rule => !rule.compatibilityOnly).length, 15)
  assert.equal(PIVOT_SHOW_DATA_AS_RULES.filter(rule => rule.compatibilityOnly).length, 1)
  assert.equal(new Set(PIVOT_SHOW_DATA_AS_RULES.map(rule => rule.type)).size, 16)
})

test.each([
  [rule('normal'), 10], [rule('percentOfGrandTotal'), 0.05], [rule('percentOfRowTotal'), 1 / 3],
  [rule('percentOfColumnTotal'), 1 / 7], [rule('percentOfParentRowTotal'), 1 / 4], [rule('percentOfParentColumnTotal'), 1 / 3],
  [rule('index'), 20 / 21], [rule('percentOfParent', 'region-label'), 1 / 4], [rule('percentOf', 'product-label', 'B'), 1 / 3],
  [rule('differenceFrom', 'product-label', 'B'), -20], [rule('percentDifferenceFrom', 'product-label', 'B'), -2 / 3],
  [rule('runningTotal', 'product-label'), 10], [rule('percentRunningTotal', 'product-label'), 1 / 4],
  [rule('rankAscending', 'product-label'), 1], [rule('rankDescending', 'product-label'), 2],
  [rule('percentOfParentGrandTotal'), 1 / 10],
])('mode %# derives a real cell from original scoped aggregates', (value, expected) => {
  const workbook = pivotShowDataAsWorkbook(), before = cells(workbook), originalRevision = report(workbook).revision
  const changed = set(workbook, value)
  assert.equal(changed, value.type !== 'normal'); assert.equal(report(workbook).revision, originalRevision + (changed ? 1 : 0))
  close(valueAt(workbook), expected); assert.deepEqual(cells(workbook), before)
})

test('percentages and index use true report denominators when all subtotal and grand-total outputs are hidden', () => {
  const workbook = pivotShowDataAsWorkbook({ totals: false })
  set(workbook, rule('percentOfGrandTotal')); close(valueAt(workbook), 0.05)
  set(workbook, rule('percentOfParentRowTotal')); close(valueAt(workbook), 0.25)
  set(workbook, rule('percentOfParentColumnTotal')); close(valueAt(workbook), 1 / 3)
  set(workbook, rule('index')); close(valueAt(workbook), 20 / 21)
  assert.equal(matrix(workbook).some(row => row[0] === 'Grand total'), false)
})

test('parent ratios use current subtotal depth and shared raw aggregates rather than sums of displayed ratios', () => {
  const workbook = pivotShowDataAsWorkbook()
  set(workbook, rule('percentOfParentRowTotal'))
  close(valueAt(workbook, ['North Total', null]), 40 / 70)
  close(valueAt(workbook, ['Grand total', null]), 1)
  set(workbook, rule('percentOfParentColumnTotal'))
  close(valueAt(workbook, ['North', 'A'], ['2024 Total', null]), 1)
  close(valueAt(workbook, ['North', 'A'], ['Grand total', null]), 1)
})

test('specified parents are non-leaf fields on either hierarchy and remove only their own descendants', () => {
  const workbook = pivotShowDataAsWorkbook()
  assert.deepEqual(getPivotShowDataAsBaseFieldIds(report(workbook), 'percentOfParent'), ['region-label', 'year-label'])
  set(workbook, rule('percentOfParent', 'region-data')); assert.equal(getPivotShowDataAs(report(workbook), 'amount-value').baseFieldId, 'region-label')
  close(valueAt(workbook), 1 / 4); close(valueAt(workbook, ['North Total', null]), 1)
  set(workbook, rule('percentOfParent', 'year-label')); close(valueAt(workbook), 1 / 3)
  const before = clone(workbook); assert.throws(() => set(workbook, rule('percentOfParent', 'product-label')), invalid); assert.deepEqual(workbook, before)
})

test('running values reset at real row parents and follow current label-sort order', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('runningTotal', 'product-label'))
  close(valueAt(workbook, ['North', 'A']), 10); close(valueAt(workbook, ['North', 'B']), 40)
  close(valueAt(workbook, ['South', 'C']), 20); close(valueAt(workbook, ['South', 'D']), 30)
  edit(workbook, { type: 'setSort', fieldId: 'product-label', direction: 'descending' })
  close(valueAt(workbook, ['North', 'B']), 30); close(valueAt(workbook, ['North', 'A']), 40)
  set(workbook, rule('percentRunningTotal', 'product-label'))
  close(valueAt(workbook, ['North', 'B']), 3 / 4); close(valueAt(workbook, ['North', 'A']), 1)
  close(valueAt(workbook, ['Grand total', null]), 1)
})

test('running values and relative differences also use column base fields and their parent contexts', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('runningTotal', 'quarter-label'))
  close(valueAt(workbook), 10); close(valueAt(workbook, ['North', 'A'], [2024, 'Q2']), 30)
  set(workbook, rule('percentRunningTotal', 'quarter-label')); close(valueAt(workbook), 1 / 3)
  set(workbook, rule('differenceFrom', 'quarter-label', '', 'previous'))
  close(valueAt(workbook), '#N/A'); close(valueAt(workbook, ['North', 'A'], [2024, 'Q2']), 10)
})

test('concrete and relative reference items keep baseline cells blank and return explicit unavailable-edge values', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('differenceFrom', 'product-label', 'A'))
  close(valueAt(workbook), null); close(valueAt(workbook, ['North', 'B']), 20); close(valueAt(workbook, ['South', 'C']), '#N/A')
  set(workbook, rule('percentDifferenceFrom', 'product-label', 'A'))
  close(valueAt(workbook), null); close(valueAt(workbook, ['North', 'B']), 2)
  set(workbook, rule('differenceFrom', 'product-label', '', 'previous'))
  close(valueAt(workbook), '#N/A'); close(valueAt(workbook, ['North', 'B']), 20); close(valueAt(workbook, ['South', 'D']), -10)
  set(workbook, rule('differenceFrom', 'product-label', '', 'next'))
  close(valueAt(workbook), -20); close(valueAt(workbook, ['North', 'B']), '#N/A')
  close(valueAt(workbook, ['Grand total', null]), null)
})

test('competition ranking preserves ties and never ranks subtotal scopes without the selected member', () => {
  const records = clone(PIVOT_SHOW_DATA_AS_RECORDS); records[2][4] = 10; records.push(['North', 'E', 2024, 'Q1', 40])
  const workbook = pivotShowDataAsWorkbook({ records }); set(workbook, rule('rankAscending', 'product-label'))
  close(valueAt(workbook, ['North', 'A']), 1); close(valueAt(workbook, ['North', 'B']), 1); close(valueAt(workbook, ['North', 'E']), 3)
  close(valueAt(workbook, ['North Total', null]), null)
  set(workbook, rule('rankDescending', 'product-label'))
  close(valueAt(workbook, ['North', 'E']), 1); close(valueAt(workbook, ['North', 'A']), 2); close(valueAt(workbook, ['North', 'B']), 2)
})

test('zero reference and parent denominators return DIV/0 while negative source values keep their actual signs', () => {
  const records = clone(PIVOT_SHOW_DATA_AS_RECORDS); records[0][4] = 0; records[2][4] = 0
  const workbook = pivotShowDataAsWorkbook({ records }); set(workbook, rule('percentOfParentRowTotal'))
  close(valueAt(workbook), '#DIV/0!')
  set(workbook, rule('percentOf', 'product-label', 'A')); close(valueAt(workbook, ['North', 'B']), '#DIV/0!')
  set(workbook, rule('percentDifferenceFrom', 'product-label', 'A')); close(valueAt(workbook, ['North', 'B']), '#DIV/0!')
  const negative = clone(PIVOT_SHOW_DATA_AS_RECORDS); negative[0][4] = -10; negative[2][4] = -30
  const negativeWorkbook = pivotShowDataAsWorkbook({ records: negative }); set(negativeWorkbook, rule('percentOfParentRowTotal'))
  close(valueAt(negativeWorkbook), 1 / 4)
})

test('finite source operands whose difference overflows produce an explicit numeric error', () => {
  const records = clone(PIVOT_SHOW_DATA_AS_RECORDS); records[0][4] = Number.MAX_VALUE; records[2][4] = -Number.MAX_VALUE
  const workbook = pivotShowDataAsWorkbook({ records }); set(workbook, rule('differenceFrom', 'product-label', 'B'))
  close(valueAt(workbook), '#NUM!')
  set(workbook, rule('percentDifferenceFrom', 'product-label', 'B')); close(valueAt(workbook), '#NUM!')
})

test('non-additive aggregation denominators are recalculated by the original engine for each measure', () => {
  const workbook = pivotShowDataAsWorkbook(), fields = pivotShowDataAsFields()
  fields.push({ ...clone(fields[4]), id: 'amount-average', position: 1, aggregation: 'average', displayName: 'Average', showDataAs: rule('percentOfGrandTotal') })
  edit(workbook, { type: 'setFieldsConfig', config: { fields } })
  close(valueAt(workbook), 10); close(valueAt(workbook, ['North', 'A'], [2024, 'Q1'], 'Average'), 10 / 25)
  assert.equal(report(workbook).fields.find(field => field.id === 'amount-value').showDataAs, undefined)
  set(workbook, rule('percentOfRowTotal'), 'amount-average')
  close(valueAt(workbook, ['North', 'A'], [2024, 'Q1'], 'Average'), 10 / 15)
})

test('full fields config can set aggregation, format and show rules once for all independently owned value instances', () => {
  const workbook = pivotShowDataAsWorkbook(), before = cells(workbook), revision = report(workbook).revision, fields = pivotShowDataAsFields()
  fields[4].showDataAs = rule('percentOfGrandTotal'); fields[4].format = '0.0%'
  fields.push({ ...clone(fields[4]), id: 'amount-rank', position: 1, displayName: 'Rank', aggregation: 'average', format: '0', showDataAs: rule('rankDescending', 'product-label') })
  edit(workbook, { type: 'setFieldsConfig', config: { fields } })
  assert.equal(report(workbook).revision, revision + 1); close(valueAt(workbook), 0.05)
  close(valueAt(workbook, ['North', 'A'], [2024, 'Q1'], 'Rank'), 2); assert.deepEqual(cells(workbook), before)
})

test('field settings combine a real show rule, aggregation and number format in one report revision', () => {
  const workbook = pivotShowDataAsWorkbook(), revision = report(workbook).revision
  edit(workbook, { type: 'setFieldSetting', fieldId: 'amount-data', setting: {
    subtotalType: 'average', format: '0.000%', showDataAs: rule('percentOfGrandTotal'),
  } })
  assert.equal(report(workbook).revision, revision + 1); close(valueAt(workbook), 0.4)
  assert.equal(calculation(workbook).numberFormats['3:2'], '0.000%')
})

test('percentage and rank output defaults respect an explicit authored format and reset removes the derived default', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('percentOfGrandTotal'))
  assert.equal(calculation(workbook).numberFormats['3:2'], '0.00%')
  set(workbook, rule('rankAscending', 'product-label')); assert.equal(calculation(workbook).numberFormats['3:2'], '0')
  set(workbook, null); assert.equal(calculation(workbook).numberFormats['3:2'], undefined); close(valueAt(workbook), 10)
})

test('same rule, canonical base ID, reordered DTO keys and explicit normal/reset are no-ops after normalization', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('percentOf', 'product-label', 'A'))
  const before = clone(workbook)
  assert.equal(set(workbook, { baseItem: 'A', baseFieldId: 'product-data', type: 'percentOf', baseItemType: 'item' }), false)
  assert.deepEqual(workbook, before)
  const sharedRevision = report(workbook).revision
  setPivotShowDataAs(workbook, report(workbook), 'amount-value', { baseItem: 'A', baseFieldId: 'product-label', type: 'percentOf' }, 'data-only')
  assert.equal(report(workbook).revision, sharedRevision)
  set(workbook, null); const normal = clone(workbook)
  assert.equal(set(workbook, rule('normal')), false); assert.equal(set(workbook, null), false); assert.deepEqual(workbook, normal)
})

test.each(['worksheet', 'data-only', 'deferred'])('shared %s mode uses one field setter and one selected refresh strategy', mode => {
  const workbook = pivotShowDataAsWorkbook({ ownership: mode === 'worksheet' ? 'linked' : 'owned' }), current = report(workbook), before = cells(workbook), revision = current.revision
  setPivotShowDataAs(workbook, current, 'amount-value', rule('percentOfGrandTotal'), mode)
  assert.equal(current.revision, revision + 1)
  if (mode === 'worksheet') { assert.notDeepEqual(cells(workbook), before); close(pivotOutputValues(workbook, current)[3][2], 0.05) }
  else { assert.deepEqual(cells(workbook), before); assert.equal(Object.hasOwn(current, 'outputRange'), false) }
  close(valueAt(workbook), 0.05)
})

test('owned and linked chart edits calculate identical show values and update every linked revision once', () => {
  const owned = pivotShowDataAsWorkbook(), linked = pivotShowDataAsWorkbook({ ownership: 'linked' }), before = linked.charts.map(chart => chart.revision)
  const value = rule('percentOfParentRowTotal'); set(owned, value); set(linked, value)
  assert.deepEqual(matrix(owned), pivotOutputValues(linked, report(linked)))
  linked.charts.forEach((current, index) => { assert.equal(current.revision, before[index] + 1); assert.deepEqual(current.source.range, report(linked).outputRange) })
  assert.equal(linked.charts[1].title, 'Second visual')
})

test('moving values to rows preserves show rules and original drill-down source membership', () => {
  const workbook = pivotShowDataAsWorkbook({ ownership: 'linked' }); set(workbook, rule('percentOfGrandTotal'))
  const detail = pivotDrillDown(workbook, report(workbook), 3, 2)
  assert.deepEqual(detail.rows, [['North', 'A', 2024, 'Q1', 10]])
  edit(workbook, { type: 'updateValuePosition', position: 'row' })
  const output = matrix(workbook)
  assert.equal(output[2][3], 0.05); assert.equal(output[2][0], 'North'); assert.equal(output[2][1], 'A'); assert.equal(output[2][2], 'Amount')
  assert.equal(getPivotShowDataAs(report(workbook), 'amount-value').type, 'percentOfGrandTotal')
})

test('base-item query returns actual grouped labels and remains distinct from raw label-filter members', () => {
  const workbook = pivotShowDataAsWorkbook()
  edit(workbook, { type: 'setManualGroups', fieldId: 'product-label', groups: [{ name: 'Mix', members: ['A', 'B'] }] })
  assert.deepEqual(getPivotShowDataAsBaseItems(workbook, report(workbook), 'product-label'), ['Mix', 'C', 'D'])
  assert.deepEqual(getPivotChartShowDataAsBaseItems(workbook, chart(workbook), 'product-data'), ['Mix', 'C', 'D'])
  assert.deepEqual(getPivotChartFieldItems(workbook, chart(workbook), 'product-data'), ['A', 'B', 'C', 'D'])
  set(workbook, rule('percentOf', 'product-label', 'Mix'))
  const before = clone(workbook); assert.throws(() => set(workbook, rule('percentOf', 'product-label', 'A')), invalid); assert.deepEqual(workbook, before)
})

test('date grouped members can be selected as a real reference item without using raw serial labels', () => {
  const records = clone(PIVOT_SHOW_DATA_AS_RECORDS); records.forEach(row => { row[3] = row[3] === 'Q1' ? 45292 : 45383 })
  const workbook = pivotShowDataAsWorkbook({ records })
  edit(workbook, { type: 'setDateGroup', fieldId: 'quarter-label', dateType: 'month' })
  assert.deepEqual(getPivotShowDataAsBaseItems(workbook, report(workbook), 'quarter-label'), ['2024-01', '2024-04'])
  set(workbook, rule('differenceFrom', 'quarter-label', '2024-01'))
  close(valueAt(workbook, ['North', 'A'], [2024, '2024-01']), null)
  close(valueAt(workbook, ['North', 'A'], [2024, '2024-04']), 10)
})

test('numeric grouped reference labels resolve real bucket aggregates', () => {
  const records = clone(PIVOT_SHOW_DATA_AS_RECORDS), products = { A: 10, B: 20, C: 30, D: 40 }; records.forEach(row => { row[1] = products[row[1]] })
  const workbook = pivotShowDataAsWorkbook({ records })
  edit(workbook, { type: 'setNumberGroup', fieldId: 'product-label', group: { start: 0, end: 50, interval: 25 } })
  assert.deepEqual(getPivotShowDataAsBaseItems(workbook, report(workbook), 'product-label'), ['0-25', '25-50'])
  set(workbook, rule('percentOf', 'product-label', '0-25'))
  close(valueAt(workbook, ['North', '0-25']), 1); close(valueAt(workbook, ['South', '25-50']), '#N/A')
})

test('stored report-local base IDs remain stable when another axis has a colliding canonical source ID', () => {
  const workbook = pivotShowDataAsWorkbook(), fields = pivotShowDataAsFields(); fields[1].id = 'quarter-data'
  fields[4].showDataAs = rule('runningTotal', 'quarter-data')
  edit(workbook, { type: 'setFieldsConfig', config: { fields } })
  assert.equal(getPivotShowDataAs(report(workbook), 'amount-value').baseFieldId, 'quarter-data')
  close(valueAt(workbook, ['North', 'B']), 40)
  set(workbook, rule('runningTotal', 'quarter-data')); close(valueAt(workbook, ['North', 'B']), 40)
})

test('ambiguous raw scalar members refuse a concrete base key instead of silently selecting one type', () => {
  const records = clone(PIVOT_SHOW_DATA_AS_RECORDS)
  records[0][1] = records[1][1] = 1; records[2][1] = records[3][1] = '1'
  const workbook = pivotShowDataAsWorkbook({ records }), before = clone(workbook)
  assert.throws(() => set(workbook, rule('percentOf', 'product-label', '1')), invalid); assert.deepEqual(workbook, before)
})

test('source filtering recalculates denominators and refresh retains unavailable reference diagnostics', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('percentOfGrandTotal'))
  edit(workbook, { type: 'setFilter', fieldId: 'region-label', items: ['North'] }); close(valueAt(workbook), 0.1)
  set(workbook, rule('percentOf', 'product-label', 'B'))
  edit(workbook, { type: 'setFilter', fieldId: 'product-label', items: ['A'] }); close(valueAt(workbook), '#N/A')
  const revision = report(workbook).revision; refreshPivotTable(workbook, report(workbook), 'data-only')
  assert.equal(report(workbook).revision, revision + 1); close(valueAt(workbook), '#N/A')
  assert.equal(getPivotShowDataAs(report(workbook), 'amount-value').baseItem, 'B')
})

test('surviving show rules prevent removal or conversion of their base fields, but a coordinated full reset is valid', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('runningTotal', 'product-label')); const before = clone(workbook)
  for (const value of [{ type: 'removeFields', fieldIds: ['product-label'] }, { type: 'resetFields', area: 'row' },
    { type: 'moveField', fieldId: 'product-label', fieldArea: 'filter', index: 0 }]) {
    assert.throws(() => edit(workbook, value), invalid); assert.deepEqual(workbook, before)
  }
  assert.equal(edit(workbook, { type: 'resetFields' }), true); assert.deepEqual(report(workbook).fields, []); assert.deepEqual(matrix(workbook), [])
})

test('full config can remove a base and reset its surviving measure rule in the same single revision', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('runningTotal', 'product-label'))
  const fields = clone(report(workbook).fields).filter(field => field.id !== 'product-label'); delete fields.find(field => field.id === 'amount-value').showDataAs
  const revision = report(workbook).revision; edit(workbook, { type: 'setFieldsConfig', config: { fields } })
  assert.equal(report(workbook).revision, revision + 1); assert.equal(getPivotShowDataAs(report(workbook), 'amount-value'), undefined)
})

test('reads and writes share target canonical-ID priority and reject an ambiguous repeated value source', () => {
  const workbook = pivotShowDataAsWorkbook(), fields = pivotShowDataAsFields(); fields[4].id = 'product-data'
  edit(workbook, { type: 'setFieldsConfig', config: { fields } })
  assert.equal(getPivotChartConfiguredField(workbook, chart(workbook), 'product-data').id, 'product-label')
  assert.equal(getPivotChartConfiguredField(workbook, chart(workbook), 'not-configured'), undefined)
  fields.push({ ...clone(fields[4]), id: 'second-value', position: 1 }); edit(workbook, { type: 'setFieldsConfig', config: { fields } })
  assert.throws(() => getPivotChartConfiguredField(workbook, chart(workbook), 'amount-data'), invalid)
})

test.each([
  { type: 'unsupported', baseFieldId: '', baseItem: '' }, { type: 4, baseFieldId: '', baseItem: '' },
  { type: 'percentOfGrandTotal' }, { ...rule('percentOfGrandTotal'), unknown: true },
  rule('percentOfGrandTotal', 'product-label'), rule('runningTotal'), rule('runningTotal', 'missing'), rule('runningTotal', 'amount-value'),
  rule('runningTotal', 'product-label', 'A'), rule('percentOf', 'product-label'), rule('percentOf', 'product-label', 'missing'),
  rule('percentOf', 'product-label', '', 'previous'), rule('differenceFrom', 'product-label', 'A', 'previous'),
  rule('differenceFrom', 'product-label', '', 'unknown'), rule('percentOfParent', 'quarter-label'),
])('invalid show rule %# refuses atomically instead of retaining an unused config', value => {
  const workbook = pivotShowDataAsWorkbook(), before = clone(workbook)
  assert.throws(() => set(workbook, value), invalid); assert.deepEqual(workbook, before)
})

test('missing or non-value targets and contradictory full-config base fields refuse without publication', () => {
  const workbook = pivotShowDataAsWorkbook(), before = clone(workbook)
  assert.throws(() => set(workbook, rule('percentOfGrandTotal'), 'region-label'), invalid)
  assert.throws(() => set(workbook, rule('percentOfGrandTotal'), 'missing'), invalid)
  const fields = pivotShowDataAsFields(); fields[4].showDataAs = rule('runningTotal', 'missing')
  assert.throws(() => edit(workbook, { type: 'setFieldsConfig', config: { fields } }), invalid); assert.deepEqual(workbook, before)
})

test('normalizer accepts only meaningful base parameters and resolves member sources without mutating the report', () => {
  const workbook = pivotShowDataAsWorkbook(), current = report(workbook), before = clone(current)
  assert.deepEqual(normalizePivotShowDataAs(current, rule('runningTotal', 'product-data')), rule('runningTotal', 'product-label'))
  assert.equal(normalizePivotShowDataAs(current, rule('normal')), undefined)
  assert.throws(() => normalizePivotShowDataAs(current, rule('percentOf', 'product-label', '1'), () => [1, '1']))
  assert.deepEqual(current, before)
})

test('invalid shared setter parameters fail before altering a combined aggregation and format edit', () => {
  const workbook = pivotShowDataAsWorkbook(), current = report(workbook), before = clone(workbook)
  assert.throws(() => setPivotFieldSetting(workbook, current, 'amount-value', { subtotalType: 'average', format: '0.0%', showDataAs: rule('percentOf', 'product-label', 'missing') }, 'data-only'))
  assert.deepEqual(workbook, before)
})

test('linked output conflict rolls back a combined show-rule/field expansion and every source mirror', () => {
  const workbook = pivotShowDataAsWorkbook({ ownership: 'linked', totals: false }), current = report(workbook)
  workbook.sheets[1].cells['12:5'] = createCell(12, 5, 'Expansion conflict')
  const before = clone(workbook), fields = clone(current.fields); fields[4].showDataAs = rule('percentOfGrandTotal')
  fields.push({ ...clone(fields[4]), id: 'amount-second', position: 1, displayName: 'Second value' })
  assert.throws(() => edit(workbook, { type: 'setFieldsConfig', config: { fields } }), PivotOutputConflictError)
  assert.deepEqual(workbook, before)
})

test('snapshot reopening retains independent show rules without source or getter aliases', () => {
  const workbook = pivotShowDataAsWorkbook(); set(workbook, rule('percentDifferenceFrom', 'product-label', 'B'))
  const saved = JSON.parse(JSON.stringify(workbook)), reopened = workbookFromInput(saved), read = getPivotShowDataAs(report(reopened), 'amount-value')
  assert.deepEqual(read, rule('percentDifferenceFrom', 'product-label', 'B')); close(valueAt(reopened), -2 / 3)
  read.baseItem = 'A'; assert.equal(getPivotShowDataAs(report(reopened), 'amount-value').baseItem, 'B')
  assert.deepEqual(cells(reopened), cells(workbook)); assert.deepEqual(reopened.pivotTables, [])
})
