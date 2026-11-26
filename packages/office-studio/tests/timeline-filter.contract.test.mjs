import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { createCell } from '../src/domain/model.ts'
import { getPivotTimelineDateRange, renderPivotTable, setPivotDateFilter, setPivotRecordSelection,
  setPivotTimelineDateRange } from '../src/features/pivot/services/pivot-engine.ts'
import { setSlicerConnections, setSlicerItems, synchronizeSlicerSelections } from '../src/features/slicer/slicer-engine.ts'
import { readTimelineDateCondition, readTimelineDateRange, readTimelineFilterState, setTimelineDateCondition,
  setTimelineDateRange, updateTimelineCacheState } from '../src/features/slicer/timeline-filter.ts'
import { timelineDateConditionToPivot } from '../src/features/slicer/timeline-date-filter.ts'
import { readTimelineSettings } from '../src/features/slicer/timeline-settings.ts'
import { filterReports, timelineAllConnections, timelineCache, timelineFilterWorkbook, timelineFirst,
  timelineReportField, timelineReportTotals, timelineViews } from './fixtures/timeline-filter-fixture.mjs'

// Source/report/cache integration drafts only. All tests remain unexecuted, including JSON reopening.
const clone = value => structuredClone(value), modes = [undefined, false, true]
const setup = options => { const workbook = timelineFilterWorkbook(options); return { workbook, view: timelineViews(workbook)[0], shared: timelineViews(workbook)[1], first: timelineFirst(workbook) } }
const rangeFor = f => ({ startDate: f.first + .25, endDate: f.first + .75 })
const allTotals = (workbook, expected) => assert.deepEqual(timelineReportTotals(workbook), filterReports(workbook).map(() => expected))
const withMode = (input, useWholeDay) => ({ ...input, ...(useWholeDay === undefined ? {} : { useWholeDay }) })
const conditions = [
  { operator: 'between', pair: true, exact: 110, days: 110 }, { operator: 'notBetween', pair: true, exact: 180, days: 180 },
  { operator: 'equal', exact: 10, days: 110 }, { operator: 'notEqual', exact: 280, days: 180 },
  { operator: 'newerThan', exact: 280, days: 180 }, { operator: 'newerThanOrEqual', exact: 290, days: 290 },
  { operator: 'olderThan', exact: 0, days: 0 }, { operator: 'olderThanOrEqual', exact: 10, days: 110 },
]

for (const date1904 of [false, true]) for (const grouped of [false, true]) test(`Date${date1904 ? '1904' : '1900'}/grouped=${grouped} range filters original typed dates across every report`, () => {
  const f = setup({ date1904, grouped }), before = clone(f.workbook), input = rangeFor(f), settings = timelineViews(f.workbook).map(readTimelineSettings)
  assert.equal(setTimelineDateRange(f.workbook, f.view, input), true); allTotals(f.workbook, 90)
  assert.deepEqual(readTimelineDateRange(f.workbook, f.shared), input); assert.equal(readTimelineDateCondition(f.workbook, f.shared), undefined)
  for (const report of filterReports(f.workbook)) {
    const field = timelineReportField(report)
    assert.deepEqual(getPivotTimelineDateRange(report, field.id), input)
    assert.equal(field.recordSelection, undefined); assert.equal(field.selectedItems, undefined); assert.equal(field.dateFilter, undefined)
    assert.equal(report.revision, filterReports(before).find(prior => prior.id === report.id).revision + 1)
  }
  assert.deepEqual(timelineViews(f.workbook).map(readTimelineSettings), settings)
  assert.deepEqual(f.workbook.sheets.find(sheet => sheet.id === 'data'), before.sheets.find(sheet => sheet.id === 'data'))
  const cache = timelineCache(f.workbook)
  assert.equal(cache.revision, timelineCache(before).revision + 1)
  assert.equal(cache.recordSelection, undefined); assert.equal(cache.selectedItems, undefined); assert.equal(cache.timelineRange, undefined)
  for (const view of timelineViews(f.workbook)) assert.equal(view.revision, before.slicers.find(prior => prior.id === view.id).revision + 1)
})

for (const date1904 of [false, true]) for (const useWholeDay of modes) test(`Date${date1904 ? '1904' : '1900'} mode=${String(useWholeDay)} retains authored bounds and exact end exclusion`, () => {
  const f = setup({ date1904 }), input = withMode(rangeFor(f), useWholeDay)
  setTimelineDateRange(f.workbook, f.view, input); allTotals(f.workbook, useWholeDay === true ? 110 : 90)
  assert.deepEqual(readTimelineDateRange(f.workbook, f.shared), input)
  assert.equal(Object.hasOwn(readTimelineDateRange(f.workbook, f.shared), 'useWholeDay'), useWholeDay !== undefined)
})

for (const date1904 of [false, true]) for (const item of conditions) for (const useWholeDay of modes) test(`Date${date1904 ? '1904' : '1900'} ${item.operator}/${String(useWholeDay)} remains a fixed predicate`, () => {
  const f = setup({ date1904 }), input = withMode({ operator: item.operator, expected: item.pair ? [f.first + .25, f.first + .75] : f.first + .25 }, useWholeDay)
  setTimelineDateCondition(f.workbook, f.view, input); allTotals(f.workbook, useWholeDay === true ? item.days : item.exact)
  assert.equal(readTimelineDateRange(f.workbook, f.shared), undefined); assert.deepEqual(readTimelineDateCondition(f.workbook, f.shared), input)
  for (const report of filterReports(f.workbook)) {
    assert.deepEqual(timelineReportField(report).dateFilter, timelineDateConditionToPivot(input))
    assert.deepEqual(timelineReportField(report).timelineDateCondition, input)
    assert.equal(timelineReportField(report).timelineDateRange, undefined)
  }
})

test('range and fixed condition replace one another; clearing restores all source rows without changing independent view settings', () => {
  const f = setup(), settings = timelineViews(f.workbook).map(readTimelineSettings)
  setTimelineDateRange(f.workbook, f.view, rangeFor(f)); allTotals(f.workbook, 90)
  setTimelineDateCondition(f.workbook, f.shared, { operator: 'notBetween', expected: [f.first + .25, f.first + .75] }); allTotals(f.workbook, 180)
  assert.equal(timelineCache(f.workbook).timelineDateRange, undefined)
  setTimelineDateRange(f.workbook, f.view, rangeFor(f)); assert.equal(timelineCache(f.workbook).timelineDateCondition, undefined)
  assert.equal(setTimelineDateRange(f.workbook, f.shared, undefined), true); allTotals(f.workbook, 6360)
  assert.deepEqual(readTimelineFilterState(timelineCache(f.workbook)), {})
  assert.deepEqual(timelineViews(f.workbook).map(readTimelineSettings), settings)
})

test('new source dates inside a retained range participate after refresh without snapshot regeneration', () => {
  const f = setup(), input = rangeFor(f); setTimelineDateRange(f.workbook, f.view, input)
  const source = f.workbook.sheets.find(sheet => sheet.id === 'data'), values = ['North', f.first + .6, 7, 1, 12]
  values.forEach((value, offset) => {
    const cell = createCell(14, offset + 3, value); if (offset === 1) cell.style.numberFormat = 'yyyy-mm-dd hh:mm:ss'
    source.cells[`14:${offset + 3}`] = cell
  })
  for (const report of filterReports(f.workbook)) report.source.range.endRow = 14
  for (const cache of f.workbook.pivotCaches) cache.source.range.endRow = 14
  for (const report of f.workbook.pivotTables) renderPivotTable(f.workbook, report)
  allTotals(f.workbook, 97); assert.deepEqual(readTimelineDateRange(f.workbook, f.shared), input)
  assert.equal(timelineCache(f.workbook).recordSelection, undefined)
})

test('removing a source date type hint excludes its colliding ordinary number while cached Date remains eligible', () => {
  const f = setup(); setTimelineDateRange(f.workbook, f.view, rangeFor(f))
  f.workbook.sheets.find(sheet => sheet.id === 'data').cells['3:4'].style.numberFormat = 'General'
  for (const report of f.workbook.pivotTables) renderPivotTable(f.workbook, report)
  allTotals(f.workbook, 80)
})

for (const kind of ['range', 'condition']) test(`retained ${kind} getters remain defensive after original source loss`, () => {
  const f = setup(), condition = { operator: 'between', expected: [f.first + .25, f.first + .75] }
  if (kind === 'range') setTimelineDateRange(f.workbook, f.view, rangeFor(f))
  else setTimelineDateCondition(f.workbook, f.view, condition)
  f.workbook.sheets = f.workbook.sheets.filter(sheet => sheet.id !== 'data'); const before = clone(f.workbook)
  if (kind === 'range') {
    const retained = readTimelineDateRange(f.workbook, f.shared); assert.deepEqual(retained, rangeFor(f)); retained.startDate = 0
    assert.deepEqual(readTimelineDateRange(f.workbook, f.view), rangeFor(f)); assert.equal(readTimelineDateCondition(f.workbook, f.view), undefined)
  } else {
    const retained = readTimelineDateCondition(f.workbook, f.shared); assert.deepEqual(retained, condition); retained.expected[0] = 0
    assert.deepEqual(readTimelineDateCondition(f.workbook, f.view), condition); assert.equal(readTimelineDateRange(f.workbook, f.view), undefined)
  }
  assert.deepEqual(f.workbook, before)
  assert.throws(() => setTimelineDateRange(f.workbook, f.view, undefined), error => error.code === 'source-missing')
  assert.deepEqual(f.workbook, before)
})

for (const missing of ['view', 'cache', 'host', 'unit']) test(`missing ${missing} resources return undefined from retained getters without recreation`, () => {
  const f = setup(); setTimelineDateRange(f.workbook, f.view, rangeFor(f))
  if (missing === 'view') f.workbook.slicers = f.workbook.slicers.filter(view => view.id !== f.view.id)
  if (missing === 'cache') f.workbook.slicerCaches = f.workbook.slicerCaches.filter(cache => cache.id !== f.view.cacheId)
  if (missing === 'host') f.workbook.sheets = f.workbook.sheets.filter(sheet => sheet.id !== f.view.hostSheetId)
  if (missing === 'unit') f.view.unitId = 'other-unit'
  const before = clone(f.workbook)
  assert.equal(readTimelineDateRange(f.workbook, f.view), undefined); assert.equal(readTimelineDateCondition(f.workbook, f.view), undefined)
  assert.throws(() => setTimelineDateRange(f.workbook, f.view, rangeFor(f)), error => error.code === 'invalid-resource')
  assert.deepEqual(f.workbook, before)
})

for (const invalid of [null, { startDate: 1, endDate: 1 }, { startDate: 2, endDate: 1 },
  { startDate: NaN, endDate: 2 }, { startDate: 1, endDate: 2, level: 'day' }]) test(`invalid live range ${JSON.stringify(invalid)} cannot modify reports or caches`, () => {
  const f = setup(), before = clone(f.workbook)
  assert.throws(() => setTimelineDateRange(f.workbook, f.view, invalid)); assert.deepEqual(f.workbook, before)
})

for (const invalid of [{ operator: 'after', expected: 1 }, { operator: 'between', expected: [2, 1] },
  { operator: 'equal', expected: [1, 2] }]) test(`invalid live condition ${JSON.stringify(invalid)} cannot modify reports or caches`, () => {
  const f = setup(), before = clone(f.workbook)
  assert.throws(() => setTimelineDateCondition(f.workbook, f.view, invalid)); assert.deepEqual(f.workbook, before)
})

test('identical authored range and repeated clearing are no-ops for reports, views and output cells', () => {
  const f = setup(); assert.equal(setTimelineDateRange(f.workbook, f.view, undefined), false)
  setTimelineDateRange(f.workbook, f.view, rangeFor(f)); const before = clone(f.workbook)
  assert.equal(setTimelineDateRange(f.workbook, f.shared, rangeFor(f)), false); assert.deepEqual(f.workbook, before)
  setTimelineDateRange(f.workbook, f.view, undefined); const cleared = clone(f.workbook)
  assert.equal(setTimelineDateCondition(f.workbook, f.view, undefined), false); assert.deepEqual(f.workbook, cleared)
})

test('zero-connection shared cache retains author filtering without changing detached reports, then reconnects its canonical predicate', () => {
  const f = setup(); setTimelineDateRange(f.workbook, f.view, rangeFor(f))
  setSlicerConnections(f.workbook, f.view, { pivotTables: [] }); const detached = clone(f.workbook.pivotTables), owned = clone(f.workbook.charts)
  const condition = { operator: 'notBetween', expected: [f.first + .25, f.first + .75] }
  setTimelineDateCondition(f.workbook, f.shared, condition)
  assert.deepEqual(f.workbook.pivotTables, detached); assert.deepEqual(f.workbook.charts, owned)
  assert.deepEqual(readTimelineDateCondition(f.workbook, f.view), condition)
  setSlicerConnections(f.workbook, f.view, timelineAllConnections(f.workbook)); allTotals(f.workbook, 180)
  for (const report of filterReports(f.workbook)) assert.deepEqual(timelineReportField(report).dateFilter, timelineDateConditionToPivot(condition))
})

test('ordinary shared Slicer synchronization preserves a canonical Timeline range and its independent levels', () => {
  const f = setup(); setTimelineDateRange(f.workbook, f.view, rangeFor(f)); const before = clone(f.workbook)
  setSlicerItems(f.workbook, f.workbook.slicers.find(view => view.type === 'slicer'), ['1'])
  synchronizeSlicerSelections(f.workbook, before)
  assert.deepEqual(readTimelineDateRange(f.workbook, f.shared), rangeFor(f))
  assert.deepEqual(timelineViews(f.workbook).map(readTimelineSettings), timelineViews(before).map(readTimelineSettings))
  assert.equal(timelineCache(f.workbook).recordSelection, undefined)
})

test('external Pivot range, fixed date condition and reset synchronize through the original source identity', () => {
  const f = setup(), report = f.workbook.pivotTables[1], field = timelineReportField(report), input = rangeFor(f)
  let before = clone(f.workbook); setPivotTimelineDateRange(f.workbook, report, field.id, input); synchronizeSlicerSelections(f.workbook, before)
  assert.deepEqual(readTimelineDateRange(f.workbook, f.view), input); allTotals(f.workbook, 90)
  before = clone(f.workbook); const condition = { operator: 'notBetween', expected: [f.first + .25, f.first + .75] }
  setPivotDateFilter(f.workbook, report, field.id, timelineDateConditionToPivot(condition)); synchronizeSlicerSelections(f.workbook, before)
  assert.deepEqual(readTimelineDateCondition(f.workbook, f.shared), condition); allTotals(f.workbook, 180)
  before = clone(f.workbook); setPivotDateFilter(f.workbook, report, field.id, null); synchronizeSlicerSelections(f.workbook, before)
  assert.deepEqual(readTimelineFilterState(timelineCache(f.workbook)), {}); allTotals(f.workbook, 6360)
})

for (const allowMultipleFilters of [false, true]) test(`legacy member migration releases its old include snapshot with multiple=${allowMultipleFilters}`, () => {
  const f = setup(), cache = timelineCache(f.workbook), selected = { mode: 'include', values: [{ type: 'date', value: f.first + .25 }] }
  for (const report of filterReports(f.workbook)) {
    report.options.allowMultipleFilters = allowMultipleFilters
    setPivotRecordSelection(f.workbook, report, timelineReportField(report).id, selected, report.generatedCellKeys === undefined ? 'data-only' : 'worksheet')
  }
  cache.recordSelection = clone(selected); cache.selectedItems = [String(f.first + .25)]
  cache.timelineRange = { start: f.first + .25, end: f.first + .25, level: 'day' }
  for (const view of timelineViews(f.workbook)) { view.selectedItems = [...cache.selectedItems]; view.timelineRange = clone(cache.timelineRange) }
  setTimelineDateRange(f.workbook, f.view, rangeFor(f)); allTotals(f.workbook, 90)
  for (const report of filterReports(f.workbook)) assert.equal(timelineReportField(report).recordSelection, undefined)
  assert.equal(cache.recordSelection, undefined); assert.equal(cache.timelineRange, undefined)
})

for (const allowMultipleFilters of [false, true]) test(`nonlegacy hand-authored member filters obey the report multiple=${allowMultipleFilters} policy`, () => {
  const f = setup(), selected = { mode: 'include', values: [{ type: 'date', value: f.first + .25 }] }
  for (const report of filterReports(f.workbook)) {
    report.options.allowMultipleFilters = allowMultipleFilters
    setPivotRecordSelection(f.workbook, report, timelineReportField(report).id, selected, report.generatedCellKeys === undefined ? 'data-only' : 'worksheet')
  }
  setTimelineDateRange(f.workbook, f.view, rangeFor(f)); allTotals(f.workbook, allowMultipleFilters ? 10 : 90)
  for (const report of filterReports(f.workbook)) assert.deepEqual(timelineReportField(report).recordSelection, allowMultipleFilters ? selected : undefined)
})

test('clear removes date rules on repeated source dimensions and keeps an unrelated caption rule', () => {
  const f = setup(), report = f.workbook.pivotTables[0], field = timelineReportField(report)
  report.options.allowMultipleFilters = true
  report.fields.push({ ...clone(field), id: 'date-repeat', area: 'filter', position: 2, timelineDateRange: rangeFor(f) })
  report.fields.find(candidate => candidate.sourceFieldId === 'region-source').labelFilter = { operator: 'captionEqual', expected: 'North' }
  updateTimelineCacheState(f.workbook, f.view, { range: rangeFor(f) })
  setTimelineDateRange(f.workbook, f.view, undefined)
  for (const candidate of report.fields.filter(candidate => candidate.sourceFieldId === field.sourceFieldId)) {
    assert.equal(candidate.timelineDateRange, undefined); assert.equal(candidate.dateFilter, undefined)
  }
  assert.deepEqual(report.fields.find(candidate => candidate.sourceFieldId === 'region-source').labelFilter, { operator: 'captionEqual', expected: 'North' })
})

for (const conflict of ['user-cell', 'new-report-overlap', 'owned-calculation']) test(`${conflict} failure leaves every report, output map, cache and shared view unchanged`, () => {
  const f = setup(), secondary = f.workbook.pivotTables[1]
  if (conflict === 'user-cell') f.workbook.sheets.find(sheet => sheet.id === secondary.anchor.subUnitId)
    .cells[`${secondary.anchor.row}:${secondary.anchor.col}`] = createCell(secondary.anchor.row, secondary.anchor.col, 'Keep author cell')
  if (conflict === 'new-report-overlap') secondary.anchor.row = f.workbook.pivotTables[0].anchor.row
  if (conflict === 'owned-calculation') f.workbook.charts.find(chart => chart.ownedPivot).ownedPivot.fields
    .find(field => field.area === 'value').showDataAs = { type: 'not-a-mode' }
  const before = clone(f.workbook)
  assert.throws(() => setTimelineDateRange(f.workbook, f.view, rangeFor(f)), error => conflict === 'owned-calculation' || error.code === 'invalid-target')
  assert.deepEqual(f.workbook, before)
})

for (const date1904 of [false, true]) for (const kind of ['range', 'condition']) test(`Date${date1904 ? '1904' : '1900'} ${kind} JSON reopening retains source rules and shared settings`, () => {
  const f = setup({ date1904 }), condition = { operator: 'notBetween', expected: [f.first + .25, f.first + .75], useWholeDay: false }
  if (kind === 'range') setTimelineDateRange(f.workbook, f.view, rangeFor(f))
  else setTimelineDateCondition(f.workbook, f.view, condition)
  const expected = clone(readTimelineFilterState(timelineCache(f.workbook))), settings = timelineViews(f.workbook).map(readTimelineSettings)
  const reopened = workbookFromInput(JSON.parse(JSON.stringify(f.workbook))), restored = timelineViews(reopened)[0]
  assert.deepEqual(readTimelineFilterState(timelineCache(reopened)), expected)
  assert.deepEqual(timelineViews(reopened).map(readTimelineSettings), settings)
  assert.deepEqual(readTimelineDateRange(reopened, restored), expected.range); assert.deepEqual(readTimelineDateCondition(reopened, restored), expected.condition)
  allTotals(reopened, kind === 'range' ? 90 : 180)
})
