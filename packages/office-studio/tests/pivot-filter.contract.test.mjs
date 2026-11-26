import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { calculatePivotReport, clearPivotFieldFilters, getPivotDateFilter, getPivotFieldFilterCapabilities,
  getPivotLabelConditionFilter, getPivotManualFilterItemState, getPivotSourceRecordSelection, getPivotSourceSelectedItems,
  pivotDrillDown, setPivotDateFilter, setPivotLabelConditionFilter, setPivotLabelFilter, setPivotManualGroups,
  setPivotOptions, setPivotRecordSelection, setPivotValueFilter } from '../src/features/pivot/services/pivot-engine.ts'
import { applyPivotChartEdit } from '../src/features/pivot/services/pivot-chart-engine.ts'
import { compilePivotDateFilter, compilePivotLabelConditionFilter, getPivotFieldFilterCategories, hasPivotMultipleFilters,
  normalizePivotDateFilter, normalizePivotRecordSelection, PIVOT_DATE_FILTER_OPERATORS,
  PIVOT_LABEL_CONDITION_OPERATORS } from '../src/features/pivot/services/pivot-filter-predicate.ts'
import { pivotFilterWorkbook } from './fixtures/pivot-filter-fixture.mjs'

// Deferred contracts: source readings and edits only; no tests/type/build/browser/WASM/benchmarks executed.
const clone = value => structuredClone(value)
const ordinary = workbook => workbook.pivotTables[0]
const owned = workbook => workbook.charts.find(chart => chart.ownedPivot).ownedPivot
const field = (report, id) => report.fields.find(field => field.id === id)
const total = (workbook, report) => calculatePivotReport(workbook, report).values.reduce((sum, row) =>
  sum + (typeof row.at(-1) === 'number' ? row.at(-1) : 0), 0)
const mode = ownership => ownership === 'owned' ? 'data-only' : 'worksheet'
const dateRules = [
  ['dateBetween', first => [first, first + 1], 100], ['dateEqual', first => first, 30],
  ['dateNewerThan', first => first, 180], ['dateNewerThanOrEqual', first => first, 210],
  ['dateNotBetween', first => [first, first + 1], 110], ['dateNotEqual', first => first, 180],
  ['dateOlderThan', first => first + 1, 30], ['dateOlderThanOrEqual', first => first + 1, 100],
]
const captionRules = [
  ['captionBeginsWith', 'N', 90], ['captionBetween', ['North', 'South'], 200], ['captionContains', 'o', 230],
  ['captionEndsWith', 'th', 230], ['captionEqual', 'North', 90], ['captionGreaterThan', 'South', 80],
  ['captionGreaterThanOrEqual', 'South', 190], ['captionLessThan', 'South', 90], ['captionLessThanOrEqual', 'South', 200],
  ['captionNotBeginsWith', 'N', 190], ['captionNotBetween', ['North', 'South'], 80], ['captionNotContains', 'o', 50],
  ['captionNotEndsWith', 'th', 50], ['captionNotEqual', 'North', 190],
]

test('fixed operator inventory is eight date and fourteen caption conditions, independent of entry counts', () => {
  assert.deepEqual(PIVOT_DATE_FILTER_OPERATORS, dateRules.map(([operator]) => operator))
  assert.deepEqual(PIVOT_LABEL_CONDITION_OPERATORS, captionRules.map(([operator]) => operator))
})

for (const date1904 of [false, true]) for (const ownership of ['ordinary', 'owned']) for (const [operator, expected, amount] of dateRules) {
  test(`${ownership}/${date1904 ? '1904' : '1900'}/${operator} filters actual formatted serial source records`, () => {
    const workbook = pivotFilterWorkbook({ date1904 }), report = ownership === 'owned' ? owned(workbook) : ordinary(workbook)
    const source = clone(workbook.sheets[0].cells), first = Math.floor(source['3:4'].value)
    setPivotDateFilter(workbook, report, 'date-local', { operator, expected: expected(first), useWholeDay: true }, mode(ownership))
    assert.equal(total(workbook, report), amount)
    assert.deepEqual(workbook.sheets[0].cells, source)
    assert.equal(getPivotDateFilter(report, 'date-local').useWholeDay, true)
    if (ownership === 'owned') assert.equal(report.generatedCellKeys, undefined)
  })
}

for (const ownership of ['ordinary', 'owned']) for (const [operator, expected, amount] of captionRules) {
  test(`${ownership}/${operator} uses case-sensitive ordinal grouped captions in real aggregation`, () => {
    const workbook = pivotFilterWorkbook(), report = ownership === 'owned' ? owned(workbook) : ordinary(workbook)
    setPivotLabelConditionFilter(workbook, report, 'region-local', { operator, expected }, mode(ownership))
    assert.equal(total(workbook, report), amount)
    assert.deepEqual(getPivotLabelConditionFilter(report, 'region-local'), { operator, expected })
  })
}

test('date comparison defaults to exact serial time and whole-day mode preserves Excel compatibility day 60', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook), first = workbook.sheets[0].cells['3:4'].value
  setPivotDateFilter(workbook, report, 'date-local', { operator: 'dateEqual', expected: first })
  assert.equal(total(workbook, report), 10); assert.equal(getPivotDateFilter(report, 'date-local').useWholeDay, false)
  setPivotDateFilter(workbook, report, 'date-local', { operator: 'dateEqual', expected: first, useWholeDay: true })
  assert.equal(total(workbook, report), 30)
  assert.equal(compilePivotDateFilter({ operator: 'dateBetween', expected: [60, 60], useWholeDay: true })(60.75), true)
  assert.equal(compilePivotDateFilter({ operator: 'dateEqual', expected: 60 })(61), false)
  assert.equal(compilePivotDateFilter({ operator: 'dateNotEqual', expected: 60 })(Number.NaN), false)
})

test('caption operators treat wildcard characters literally and preserve ordinal case ordering', () => {
  const starts = compilePivotLabelConditionFilter({ operator: 'captionBeginsWith', expected: 'N*' })
  assert.equal(starts('North'), false); assert.equal(starts('N*orth'), true)
  assert.equal(compilePivotLabelConditionFilter({ operator: 'captionEqual', expected: 'North' })('north'), false)
  assert.equal(compilePivotLabelConditionFilter({ operator: 'captionGreaterThan', expected: 'West' })('north'), true)
})

for (const invalid of [
  { operator: 'dateTomorrow', expected: 1 }, { operator: 'dateEqual', expected: Number.NaN },
  { operator: 'dateBetween', expected: [2, 1] }, { operator: 'dateEqual', expected: [1, 2] },
  { operator: 'dateBetween', expected: [1, 2, 3] }, { operator: 'dateEqual', expected: 1, useWholeDay: 'yes' },
]) test(`invalid date operand ${JSON.stringify(invalid)} leaves report and output unchanged`, () => {
  const workbook = pivotFilterWorkbook(), before = clone(workbook)
  assert.throws(() => setPivotDateFilter(workbook, ordinary(workbook), 'date-local', invalid), /Invalid pivot|finite and ordered/)
  assert.deepEqual(workbook, before)
})

test('date capability uses real number formatting and ISO adapters, rather than the field name', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook)
  assert.deepEqual(getPivotFieldFilterCapabilities(workbook, report, 'date-local'), { supportsDateFilter: true, date1904: false })
  assert.equal(getPivotFieldFilterCapabilities(workbook, report, 'number-local').supportsDateFilter, false)
  assert.equal(getPivotFieldFilterCapabilities(workbook, report, 'amount-local').supportsDateFilter, false)
  const before = clone(workbook)
  assert.throws(() => setPivotDateFilter(workbook, report, 'number-local', { operator: 'dateEqual', expected: 1 }), /date source field/)
  assert.deepEqual(workbook, before)
  const first = Math.floor(workbook.sheets[0].cells['3:4'].value)
  workbook.sheets[0].cells['3:4'].value = '2026-01-01'; workbook.sheets[0].cells['3:4'].style.numberFormat = 'General'
  setPivotDateFilter(workbook, report, 'date-local', { operator: 'dateEqual', expected: first })
  assert.equal(total(workbook, report), 10)
})

test('aggregation and drill-down share member, caption and date predicates', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook), first = Math.floor(workbook.sheets[0].cells['3:4'].value)
  setPivotOptions(workbook, report, { allowMultipleFilters: true })
  setPivotLabelFilter(workbook, report, 'region-local', ['North'])
  setPivotLabelConditionFilter(workbook, report, 'region-local', { operator: 'captionBeginsWith', expected: 'N' })
  setPivotDateFilter(workbook, report, 'date-local', { operator: 'dateEqual', expected: first, useWholeDay: true })
  assert.equal(total(workbook, report), 30)
  const calculation = calculatePivotReport(workbook, report), rows = []
  for (let row = calculation.headerRows; row < calculation.values.length; row++) rows.push(...pivotDrillDown(workbook, report, row, 2).rows)
  assert.equal(rows.reduce((sum, row) => sum + row[2], 0), 30)
  assert.ok(rows.every(row => row[0] === 'North' && Math.floor(row[1]) === first))
})

test('default exclusive categories replace member, condition and value filters only for the same canonical source', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook)
  setPivotLabelFilter(workbook, report, 'mixed-local', ['1'])
  setPivotLabelFilter(workbook, report, 'region-local', ['North'])
  setPivotLabelConditionFilter(workbook, report, 'region-local', { operator: 'captionEqual', expected: 'North' })
  assert.equal(field(report, 'region-local').selectedItems, undefined)
  assert.deepEqual([...getPivotFieldFilterCategories(field(report, 'region-local'))], ['condition'])
  setPivotValueFilter(workbook, report, 'region-local', { operator: 'valueGreaterThan', expected: 25, valueFieldId: 'amount-local' })
  assert.equal(field(report, 'region-local').labelFilter, undefined)
  assert.deepEqual([...getPivotFieldFilterCategories(field(report, 'region-local'))], ['value'])
  assert.deepEqual(field(report, 'mixed-local').selectedItems, ['1']); assert.equal(total(workbook, report), 30)
})

test('enabled multiple categories intersect before scoped value aggregation', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook)
  setPivotOptions(workbook, report, { allowMultipleFilters: true })
  setPivotLabelFilter(workbook, report, 'region-local', ['North', 'South'])
  setPivotLabelConditionFilter(workbook, report, 'region-local', { operator: 'captionBeginsWith', expected: 'N' })
  setPivotValueFilter(workbook, report, 'region-local', { operator: 'valueGreaterThan', expected: 80, valueFieldId: 'amount-local' })
  assert.deepEqual([...getPivotFieldFilterCategories(field(report, 'region-local'))], ['member', 'condition', 'value'])
  assert.equal(hasPivotMultipleFilters(report), true); assert.equal(total(workbook, report), 90)
})

test('date and caption share the condition category while a date-only reset preserves other categories', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook), first = workbook.sheets[0].cells['3:4'].value
  setPivotOptions(workbook, report, { allowMultipleFilters: true })
  setPivotLabelFilter(workbook, report, 'date-local', [String(first)])
  setPivotValueFilter(workbook, report, 'date-local', { operator: 'valueGreaterThan', expected: 5, valueFieldId: 'amount-local' })
  setPivotDateFilter(workbook, report, 'date-local', { operator: 'dateEqual', expected: first })
  setPivotLabelConditionFilter(workbook, report, 'date-local', { operator: 'captionEqual', expected: String(first) })
  assert.equal(field(report, 'date-local').dateFilter, undefined)
  const before = clone(workbook); setPivotDateFilter(workbook, report, 'date-local', null); assert.deepEqual(workbook, before)
  setPivotDateFilter(workbook, report, 'date-local', { operator: 'dateEqual', expected: first })
  assert.equal(field(report, 'date-local').labelFilter, undefined)
  setPivotDateFilter(workbook, report, 'date-local', null)
  assert.deepEqual([...getPivotFieldFilterCategories(field(report, 'date-local'))], ['member', 'value'])
  assert.equal(total(workbook, report), 10)
})

test('explicit false clears the entire report when one source has multiple categories, including hidden repeated instances', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook)
  report.fields.push({ id: 'region-repeat', sourceFieldId: 'region-source', sourceColumn: 3, name: 'Region', area: 'filter', position: 2,
    labelFilter: { operator: 'captionEqual', expected: 'North' } })
  field(report, 'region-local').selectedItems = ['North']
  field(report, 'mixed-local').recordSelection = { mode: 'include', values: [1] }
  assert.equal(hasPivotMultipleFilters(report), true)
  setPivotOptions(workbook, report, { allowMultipleFilters: false })
  assert.ok(report.fields.every(candidate => getPivotFieldFilterCategories(candidate).size === 0))
  assert.equal(report.options.allowMultipleFilters, undefined); assert.equal(total(workbook, report), 280)
})

test('different sources with one category each survive an explicit false option and retain the default no-op', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook), first = Math.floor(workbook.sheets[0].cells['3:4'].value)
  setPivotLabelFilter(workbook, report, 'region-local', ['North'])
  setPivotDateFilter(workbook, report, 'date-local', { operator: 'dateEqual', expected: first, useWholeDay: true })
  assert.equal(hasPivotMultipleFilters(report), false)
  const before = clone(workbook); setPivotOptions(workbook, report, { allowMultipleFilters: false })
  assert.deepEqual(workbook, before); assert.equal(total(workbook, report), 30)
})

test('clear-all removes every filter category on canonical source copies and leaves unrelated raw members', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook)
  setPivotOptions(workbook, report, { allowMultipleFilters: true })
  setPivotRecordSelection(workbook, report, 'mixed-local', { mode: 'include', values: [1] })
  setPivotLabelFilter(workbook, report, 'region-local', ['North'])
  report.fields.push({ id: 'region-repeat', sourceFieldId: 'region-source', sourceColumn: 3, name: 'Region', area: 'filter', position: 2,
    labelFilter: { operator: 'captionEqual', expected: 'North' } })
  clearPivotFieldFilters(workbook, report, 'region-repeat')
  for (const candidate of report.fields.filter(candidate => candidate.sourceFieldId === 'region-source')) assert.equal(getPivotFieldFilterCategories(candidate).size, 0)
  assert.deepEqual(getPivotSourceRecordSelection(workbook, report, 'mixed-local'), { mode: 'include', values: [1] })
  assert.equal(total(workbook, report), 10)
})

test('caption conditions match displayed manual groups while raw member queries retain source types', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook)
  setPivotManualGroups(workbook, report, 'mixed-local', [{ name: 'Duplicate', members: [1, '1'] }])
  setPivotLabelConditionFilter(workbook, report, 'mixed-local', { operator: 'captionEqual', expected: 'Duplicate' })
  assert.equal(total(workbook, report), 30)
  setPivotLabelFilter(workbook, report, 'mixed-local', ['Duplicate'], 'worksheet', { keyMode: 'grouped' })
  assert.deepEqual(getPivotManualFilterItemState(workbook, report, 'mixed-local'), { checkedItems: ['Duplicate'], partialItems: [] })
  assert.deepEqual(getPivotSourceSelectedItems(workbook, report, 'mixed-local'), ['1'])
  assert.deepEqual(getPivotSourceRecordSelection(workbook, report, 'mixed-local'), { mode: 'include', values: [1, '1'] })
})

for (const [raw, key, amount] of [[1, '1', 10], [true, 'true', 30], [null, '', 50]]) {
  test(`partial ${key} preserves ${typeof raw} membership and reports an actual partial checkbox state`, () => {
    const workbook = pivotFilterWorkbook(), report = ordinary(workbook)
    setPivotRecordSelection(workbook, report, 'mixed-local', { mode: 'include', values: [raw] })
    setPivotLabelFilter(workbook, report, 'mixed-local', [], 'worksheet', { keyMode: 'grouped', partialItems: [key] })
    assert.deepEqual(getPivotManualFilterItemState(workbook, report, 'mixed-local'), { checkedItems: [], partialItems: [key] })
    assert.deepEqual(getPivotSourceRecordSelection(workbook, report, 'mixed-local'), { mode: 'include', values: [raw] })
    assert.equal(total(workbook, report), amount)
  })
}

test('invalid grouped and partial checklist data changes no report or output', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook), before = clone(workbook)
  for (const [items, options] of [[['1'], { keyMode: 'grouped', partialItems: ['1'] }],
    [[], { partialItems: ['1'] }], [[], { keyMode: 'grouped', partialItems: ['missing'] }],
    [null, { keyMode: 'grouped', partialItems: ['1'] }]]) {
    assert.throws(() => setPivotLabelFilter(workbook, report, 'mixed-local', items, 'worksheet', options), /partial|Partial|Invalid Pivot|both checked/)
    assert.deepEqual(workbook, before)
  }
})

test('normalized filter getter snapshots are defensive and setting the same condition remains a no-op', () => {
  const workbook = pivotFilterWorkbook(), report = ordinary(workbook), first = Math.floor(workbook.sheets[0].cells['3:4'].value)
  const filter = { operator: 'dateBetween', expected: [first, first + 1] }
  setPivotDateFilter(workbook, report, 'date-local', filter)
  filter.expected[0] = -1
  const snapshot = getPivotDateFilter(report, 'date-local'); snapshot.expected[1] = 99999
  const before = clone(workbook)
  setPivotDateFilter(workbook, report, 'date-local', { operator: 'dateBetween', expected: [first, first + 1], useWholeDay: false })
  assert.deepEqual(workbook, before)
  assert.equal(normalizePivotDateFilter({ operator: 'dateEqual', expected: first }).useWholeDay, false)
  assert.deepEqual(normalizePivotRecordSelection({ mode: 'include', values: [1, '1', 1] }), { mode: 'include', values: [1, '1'] })
})

for (const flag of [undefined, true, false]) {
  test(`full owned configuration ${String(flag)} preserves snapshot categories or explicitly clears the report`, () => {
    const workbook = pivotFilterWorkbook(), chart = workbook.charts.find(chart => chart.ownedPivot), fields = clone(owned(workbook).fields)
    fields.find(candidate => candidate.id === 'region-local').selectedItems = ['North']
    fields.find(candidate => candidate.id === 'region-local').labelFilter = { operator: 'captionEqual', expected: 'North' }
    fields.find(candidate => candidate.id === 'mixed-local').recordSelection = { mode: 'include', values: [1] }
    applyPivotChartEdit(workbook, chart, { type: 'setFieldsConfig', config: { fields,
      ...(flag === undefined ? {} : { options: { allowMultipleFilters: flag } }) } })
    assert.equal(hasPivotMultipleFilters(chart.ownedPivot), flag !== false)
    assert.equal(total(workbook, chart.ownedPivot), flag === false ? 280 : 10)
    const reopened = workbookFromInput(JSON.parse(JSON.stringify(workbook)))
    assert.equal(total(reopened, owned(reopened)), flag === false ? 280 : 10)
    assert.equal(hasPivotMultipleFilters(owned(reopened)), flag !== false)
  })
}
