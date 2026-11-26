import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { createCell } from '../src/domain/model.ts'
import { calculatePivotReport, getPivotSourceRecordSelection, renderPivotTable, setPivotLabelFilter,
  setPivotManualGroups, setPivotRecordSelection } from '../src/features/pivot/services/pivot-engine.ts'
import { getPivotChartReport } from '../src/features/pivot/services/pivot-chart-engine.ts'
import { createSlicer, readSlicerItemState, readSlicerRecordSelection, readSlicerSelection, setSlicerConnections,
  setSlicerItems, setSlicerItemStates, setSlicerRecordSelection, synchronizeSlicerSelections } from '../src/features/slicer/slicer-engine.ts'
import { getSlicerCache, getSlicerConnectionReports, normalizeSlicerCaches } from '../src/features/slicer/slicer-connections.ts'
import { applySlicerReportFilter } from '../src/features/slicer/slicer-report-filter.ts'
import { sameSlicerRecordSelection } from '../src/features/slicer/slicer-record-selection.ts'
import { filterReports, pivotFilterWorkbook, slicerRecordSelectionWorkbook } from './fixtures/pivot-filter-fixture.mjs'

// Deferred model/contracts only: no test, type, build, browser, WASM or benchmark was executed.
const clone = value => structuredClone(value)
const view = workbook => workbook.slicers.find(control => control.id === 'typed-slicer')
const mixedField = report => report.fields.find(field => field.sourceFieldId === 'mixed-source')
const total = (workbook, report) => calculatePivotReport(workbook, report).values.reduce((sum, row) =>
  sum + (typeof row.at(-1) === 'number' ? row.at(-1) : 0), 0)
const choose = (workbook, selection) => {
  const before = clone(workbook), report = workbook.pivotTables[0]
  setPivotRecordSelection(workbook, report, mixedField(report).id, selection)
  synchronizeSlicerSelections(workbook, before)
}
const addRawMember = (workbook, raw) => {
  const data = workbook.sheets.find(sheet => sheet.id === 'data'), row = 10
  ;['Extra', data.cells['3:4'].value, 25, raw, 8].forEach((value, index) => { data.cells[`${row}:${index + 3}`] = createCell(row, index + 3, value) })
  data.cells['10:4'].style.numberFormat = 'yyyy-mm-dd hh:mm:ss'
  for (const report of filterReports(workbook)) report.source.range.endRow = row
  for (const cache of workbook.pivotCaches) cache.source.range.endRow = row
}

test('raw identity distinguishes number/text, boolean/text and null/empty while order and duplicate values are no-ops', () => {
  for (const [first, second] of [[1, '1'], [true, 'true'], [null, '']]) {
    assert.equal(sameSlicerRecordSelection({ mode: 'include', values: [first] }, { mode: 'include', values: [second] }), false)
  }
  assert.equal(sameSlicerRecordSelection({ mode: 'include', values: [1, '1', null, true, 1] },
    { mode: 'include', values: [true, null, '1', 1] }), true)
  assert.equal(sameSlicerRecordSelection({ mode: 'include', values: [] }, undefined), false)
  assert.equal(sameSlicerRecordSelection({ mode: 'all' }, undefined), true)
  assert.equal(sameSlicerRecordSelection({ mode: 'exclude', values: [] }, undefined), true)
})

for (const [raw, key, expected] of [[1, '1', 10], [true, 'true', 30], [null, '', 50]]) {
  test(`typed ${typeof raw}:${String(raw)} fans out without selecting the colliding caption`, () => {
    const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
    choose(workbook, { mode: 'include', values: [raw] })
    for (const report of filterReports(workbook)) {
      assert.equal(total(workbook, report), expected)
      assert.deepEqual(getPivotSourceRecordSelection(workbook, report, mixedField(report).id), { mode: 'include', values: [raw] })
    }
    assert.equal(getPivotChartReport(workbook, workbook.charts.find(chart => chart.id === 'linked-chart')), workbook.pivotTables[0])
    assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'include', values: [raw] })
    assert.deepEqual(readSlicerSelection(workbook, control), [key])
    for (const shared of workbook.slicers) assert.deepEqual(readSlicerItemState(workbook, shared), { checkedItems: [], partialItems: [key] })
    assert.notEqual(mixedField(workbook.pivotTables[1]).id, mixedField(workbook.pivotTables[0]).id)
  })
}

test('a full raw caption checkbox selects both scalar types and clear removes the shared raw selection', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
  choose(workbook, { mode: 'include', values: [1] })
  setSlicerItems(workbook, control, ['1'])
  for (const report of filterReports(workbook)) assert.equal(total(workbook, report), 30)
  assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'include', values: [1, '1'] })
  assert.deepEqual(readSlicerItemState(workbook, control), { checkedItems: ['1'], partialItems: [] })
  setSlicerItems(workbook, control, null)
  assert.equal(readSlicerRecordSelection(workbook, control), undefined)
  for (const report of filterReports(workbook)) {
    assert.equal(mixedField(report).recordSelection, undefined); assert.equal(total(workbook, report), 280)
  }
})

test('toggling another display key preserves number, boolean and blank partial membership', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
  choose(workbook, { mode: 'include', values: [1, true, null] })
  setSlicerItemStates(workbook, control, ['beta'], ['1', 'true', ''])
  assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'include', values: [1, true, null, 'beta'] })
  assert.deepEqual(readSlicerItemState(workbook, control), { checkedItems: ['beta'], partialItems: ['1', 'true', ''] })
  for (const report of filterReports(workbook)) assert.equal(total(workbook, report), 160)
  const before = clone(workbook); setSlicerItemStates(workbook, control, ['beta'], ['1', 'true', '']); assert.deepEqual(workbook, before)
})

test('partial item edits retain exclude semantics and explicitly retained unavailable raw members', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
  setSlicerRecordSelection(workbook, control, { mode: 'exclude', values: [1, 'future'] })
  setSlicerItemStates(workbook, control, ['true', ''], ['1'])
  assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'exclude', values: [1, 'beta', 'future'] })
  for (const report of filterReports(workbook)) assert.equal(total(workbook, report), 200)
})

test('Canvas checked-all clears the member rule while a direct include-all retains explicit raw membership', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
  setSlicerRecordSelection(workbook, control, { mode: 'include', values: [1, '1', true, 'true', null, '', 'beta'] })
  assert.equal(readSlicerRecordSelection(workbook, control).mode, 'include')
  setSlicerItemStates(workbook, control, ['1', 'true', '', 'beta'], [])
  assert.equal(readSlicerRecordSelection(workbook, control), undefined)
  const before = clone(workbook); addRawMember(workbook, false); synchronizeSlicerSelections(workbook, before)
  for (const report of filterReports(workbook)) assert.equal(total(workbook, report), 305)
})

for (const [checked, partial] of [[['1'], ['1']], [['missing'], []], [[], ['missing']]]) {
  test(`invalid raw item-state ${JSON.stringify([checked, partial])} changes no cache or report`, () => {
    const workbook = slicerRecordSelectionWorkbook(), before = clone(workbook)
    assert.throws(() => setSlicerItemStates(workbook, view(workbook), checked, partial), /both checked|unknown item/)
    assert.deepEqual(workbook, before)
  })
}

test('grouped partial rules project typed raw members and never propagate the group caption as source identity', () => {
  const workbook = slicerRecordSelectionWorkbook(), report = workbook.pivotTables[0], control = view(workbook)
  choose(workbook, { mode: 'include', values: [1] })
  const before = clone(workbook)
  setPivotManualGroups(workbook, report, mixedField(report).id, [{ name: 'Duplicate', members: [1, '1'] }])
  setPivotLabelFilter(workbook, report, mixedField(report).id, [], 'worksheet', { keyMode: 'grouped', partialItems: ['Duplicate'] })
  synchronizeSlicerSelections(workbook, before)
  assert.deepEqual(mixedField(report).partialItems, ['Duplicate'])
  assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'include', values: [1] })
  for (const candidate of filterReports(workbook)) assert.equal(total(workbook, candidate), 10)
  const partial = clone(workbook)
  setPivotLabelFilter(workbook, report, mixedField(report).id, ['Duplicate'], 'worksheet', { keyMode: 'grouped' })
  synchronizeSlicerSelections(workbook, partial)
  for (const candidate of filterReports(workbook)) assert.equal(total(workbook, candidate), 30)
  assert.deepEqual(readSlicerSelection(workbook, control), ['1'])
  assert.equal(readSlicerSelection(workbook, control).includes('Duplicate'), false)
})

test('detached reports retain typed membership and reconnect uses current cache raw identity', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
  choose(workbook, { mode: 'include', values: [1] })
  setSlicerConnections(workbook, control, { pivotTables: [] })
  setSlicerItems(workbook, control, ['true'])
  for (const report of filterReports(workbook)) assert.deepEqual(getPivotSourceRecordSelection(workbook, report, mixedField(report).id), { mode: 'include', values: [1] })
  setSlicerConnections(workbook, control, { pivotTables: [{ subUnitId: 'output', pivotTableId: 'secondary-report' }],
    ownedPivotCharts: [{ subUnitId: 'dashboard', pivotChartId: 'owned-chart' }] })
  assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'include', values: [true, 'true'] })
  assert.equal(total(workbook, workbook.pivotTables[1]), 70)
  assert.equal(total(workbook, workbook.charts.find(chart => chart.ownedPivot).ownedPivot), 70)
  assert.equal(total(workbook, workbook.pivotTables[0]), 10)
})

test('typed cache snapshots reopen defensively and preserve three scalar collisions', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook), selected = { mode: 'include', values: [1, true, null] }
  choose(workbook, selected)
  const snapshot = readSlicerRecordSelection(workbook, control); snapshot.values.push('1')
  assert.deepEqual(readSlicerRecordSelection(workbook, control), selected)
  const reopened = workbookFromInput(JSON.parse(JSON.stringify(workbook)))
  assert.deepEqual(normalizeSlicerCaches(reopened), [])
  assert.deepEqual(readSlicerRecordSelection(reopened, view(reopened)), selected)
  for (const report of filterReports(reopened)) assert.equal(total(reopened, report), 90)
  assert.deepEqual(readSlicerItemState(reopened, view(reopened)), { checkedItems: [], partialItems: ['1', 'true', ''] })
})

test('new controls seed a typed cache from the existing report without converting partial members to captions', () => {
  const workbook = pivotFilterWorkbook(), report = workbook.pivotTables[0]
  setPivotRecordSelection(workbook, report, mixedField(report).id, { mode: 'include', values: [1] })
  const control = createSlicer(workbook, { id: 'new-control', unitId: workbook.id, hostSheetId: 'dashboard',
    pivotTableId: report.id, sourceFieldId: 'mixed-source', type: 'slicer' })
  assert.deepEqual(getSlicerCache(workbook, control).recordSelection, { mode: 'include', values: [1] })
  assert.deepEqual(readSlicerItemState(workbook, control), { checkedItems: [], partialItems: ['1'] })
})

test('different canonical source fields retain separate typed cache filters even with equal raw values', () => {
  const workbook = slicerRecordSelectionWorkbook(), mixed = view(workbook)
  const number = createSlicer(workbook, { id: 'number-slicer', unitId: workbook.id, hostSheetId: 'dashboard',
    pivotTableId: 'ordinary-report', sourceFieldId: 'number-source', type: 'slicer' })
  setSlicerRecordSelection(workbook, mixed, { mode: 'include', values: [1] })
  assert.equal(readSlicerRecordSelection(workbook, number), undefined)
  setSlicerRecordSelection(workbook, number, { mode: 'include', values: [2] })
  assert.deepEqual(readSlicerRecordSelection(workbook, mixed), { mode: 'include', values: [1] })
  assert.notEqual(getSlicerCache(workbook, mixed).id, getSlicerCache(workbook, number).id)
  assert.equal(total(workbook, workbook.pivotTables[0]), 0)
  assert.equal(total(workbook, workbook.pivotTables[1]), 10)
})

for (const [mode, beforeTotal, afterTotal] of [['include', 10, 0], ['exclude', 270, 280]]) {
  test(`${mode} retains a pruned raw identity and applies it when the member returns`, () => {
    const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
    choose(workbook, { mode, values: [1] })
    for (const report of filterReports(workbook)) assert.equal(total(workbook, report), beforeTotal)
    const before = clone(workbook); workbook.sheets.find(sheet => sheet.id === 'data').cells['3:6'].value = 'replacement'
    synchronizeSlicerSelections(workbook, before)
    assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode, values: [1] })
    for (const report of filterReports(workbook)) assert.equal(total(workbook, report), afterTotal)
    workbook.sheets.find(sheet => sheet.id === 'data').cells['3:6'].value = 1
    for (const report of filterReports(workbook)) assert.equal(total(workbook, report), beforeTotal)
  })
}

for (const [selection, expected] of [[{ mode: 'include', values: [1] }, 10], [{ mode: 'exclude', values: [1] }, 295], [{ mode: 'all' }, 305]]) {
  test(`${selection.mode} applies to a new typed source member across all connected reports`, () => {
    const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
    choose(workbook, selection)
    const before = clone(workbook); addRawMember(workbook, false); synchronizeSlicerSelections(workbook, before)
    for (const report of filterReports(workbook)) assert.equal(total(workbook, report), expected)
    assert.deepEqual(readSlicerRecordSelection(workbook, control), selection.mode === 'all' ? undefined : selection)
  })
}

test('include-empty remains an empty shared report while all and exclude-empty clear it', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
  choose(workbook, { mode: 'include', values: [] })
  assert.deepEqual(getSlicerCache(workbook, control).recordSelection, { mode: 'include', values: [] })
  for (const report of filterReports(workbook)) assert.equal(total(workbook, report), 0)
  choose(workbook, { mode: 'exclude', values: [] })
  for (const report of filterReports(workbook)) assert.equal(total(workbook, report), 280)
  assert.equal(readSlicerRecordSelection(workbook, control), undefined)
})

test('unchanged typed fan-out, shared selection and connections do not revise any model', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
  choose(workbook, { mode: 'include', values: [1, '1'] })
  const before = clone(workbook), reports = getSlicerConnectionReports(workbook, control)
  assert.equal(applySlicerReportFilter(workbook, reports, { mode: 'include', values: ['1', 1, 1] }), false)
  assert.equal(setSlicerConnections(workbook, control, clone(getSlicerCache(workbook, control).connections)), true)
  synchronizeSlicerSelections(workbook, before); assert.deepEqual(workbook, before)
})

for (const allowMultipleFilters of [false, true]) {
  test(`same raw member selection obeys source category exclusivity when multiple=${allowMultipleFilters}`, () => {
    const workbook = slicerRecordSelectionWorkbook(), control = view(workbook), report = workbook.pivotTables[0]
    choose(workbook, { mode: 'include', values: [1] })
    report.options.allowMultipleFilters = allowMultipleFilters
    mixedField(report).labelFilter = { operator: 'captionEqual', expected: '1' }
    setSlicerRecordSelection(workbook, control, { mode: 'include', values: [1] })
    assert.equal(mixedField(report).labelFilter !== undefined, allowMultipleFilters)
    assert.equal(total(workbook, report), 10)
  })
}

test('equal string mirrors with conflicting raw scalar choices refuse the shared transaction', () => {
  const workbook = slicerRecordSelectionWorkbook(), before = clone(workbook)
  const [ordinary, secondary] = workbook.pivotTables
  setPivotRecordSelection(workbook, ordinary, mixedField(ordinary).id, { mode: 'include', values: [1] })
  setPivotRecordSelection(workbook, secondary, mixedField(secondary).id, { mode: 'include', values: ['1'] })
  const staged = clone(workbook), cache = clone(getSlicerCache(workbook, view(workbook)))
  assert.throws(() => synchronizeSlicerSelections(workbook, before), /不同的报告筛选/)
  assert.deepEqual(workbook, staged); assert.deepEqual(getSlicerCache(workbook, view(workbook)), cache)
})

test('an output conflict stages no partial typed report selection or cache update', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook)
  choose(workbook, { mode: 'include', values: [1] })
  const secondary = workbook.pivotTables[1], output = workbook.sheets.find(sheet => sheet.id === 'output')
  output.cells[`${secondary.anchor.row + 7}:${secondary.anchor.col + 2}`] = createCell(secondary.anchor.row + 7, secondary.anchor.col + 2, 'occupied')
  const before = clone(workbook)
  assert.throws(() => setSlicerItems(workbook, control, null), /overlap|overwrite|conflict|occupied/i)
  assert.deepEqual(workbook, before)
})

test('normalization resolves observed legacy raw types and retains unobservable legacy selections', () => {
  const workbook = slicerRecordSelectionWorkbook(), control = view(workbook), cache = getSlicerCache(workbook, control)
  cache.selectedItems = ['1']; delete cache.recordSelection
  assert.deepEqual(normalizeSlicerCaches(workbook), [])
  assert.deepEqual(cache.recordSelection, { mode: 'include', values: [1, '1'] })
  cache.selectedItems = ['missing member']; delete cache.recordSelection
  assert.deepEqual(normalizeSlicerCaches(workbook), [])
  assert.equal(cache.recordSelection, undefined); assert.deepEqual(cache.selectedItems, ['missing member'])
  assert.throws(() => readSlicerRecordSelection(workbook, control), /原始类型/)
  const retained = clone(cache); workbook.sheets = workbook.sheets.filter(sheet => sheet.id !== 'data')
  assert.ok(normalizeSlicerCaches(workbook).some(diagnostic => diagnostic.code === 'source-missing'))
  assert.deepEqual(cache, retained)
})
