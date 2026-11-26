import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { getRecordValueKey } from '../src/domain/record-value.ts'
import { setAutoFilter, setFilterCriteria } from '../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { setPivotLabelFilter } from '../src/features/pivot/services/pivot-engine.ts'
import { getSlicerCache } from '../src/features/slicer/slicer-connections.ts'
import { createSlicer, readSlicerRecordSelection, setSlicerConnections, setSlicerRecordSelection } from '../src/features/slicer/slicer-engine.ts'
import { applySlicerCacheItemOptions, effectiveSlicerCacheOptions, normalizeSlicerCacheOptionsPatch,
  patchSlicerCacheOptions, readSlicerCacheOptions } from '../src/features/slicer/slicer-cache-options.ts'
import { SlicerMemberQueryService } from '../src/features/slicer/slicer-member-query.ts'
import { filterReports, slicerRecordSelectionWorkbook, slicerTableRecordSelectionWorkbook } from './fixtures/pivot-filter-fixture.mjs'
import { createSlicerMemberTestClient } from './fixtures/slicer-member-compute-fixture.mjs'

// Source, native integration and shared-option drafts only. Every test remains unexecuted.
const clone = value => structuredClone(value), kinds = ['pivot', 'table']
const setup = kind => {
  const workbook = kind === 'table' ? slicerTableRecordSelectionWorkbook() : slicerRecordSelectionWorkbook()
  const slicer = workbook.slicers[0], shared = workbook.slicers[1]
  return { workbook, slicer, shared, cache: () => getSlicerCache(workbook, slicer) }
}
const members = values => values.map((value, index) => ({ key: getRecordValueKey(value), value, hasData: index % 2 === 0 }))
const scalar = value => value.type === 'blank' ? null : value.value
async function query(f) {
  const service = new SlicerMemberQueryService({ computeClient: createSlicerMemberTestClient() })
  try { return await service.getItems(f.workbook, f.slicer, { start: 0, count: 100 }, { unitId: f.workbook.id, baseRevision: 0, assertCurrent() {} }) }
  finally { service.dispose() }
}
function northOnly(f) {
  if (getSlicerCache(f.workbook, f.slicer).source.kind === 'table') {
    const table = f.workbook.tables.find(table => table.id === 'typed-table'), source = f.workbook.sheets.find(sheet => sheet.id === 'data')
    const sheet = { ...source, autoFilter: table.autoFilter ? clone(table.autoFilter) : undefined }
    if (!sheet.autoFilter) setAutoFilter(sheet, table.range)
    setFilterCriteria(sheet, 3, { kind: 'condition', operator: 'equals', value: 'North' }); table.autoFilter = sheet.autoFilter
  } else for (const report of filterReports(f.workbook)) {
    const region = report.fields.find(field => field.sourceFieldId === 'region-source')
    setPivotLabelFilter(f.workbook, report, region.id, ['North'], report.generatedCellKeys === undefined ? 'data-only' : 'worksheet')
  }
}

for (const kind of kinds) test(`${kind} cache options are shared defensive authored state and preserve all filter authorities`, () => {
  const f = setup(kind); setSlicerRecordSelection(f.workbook, f.slicer, { mode: 'include', values: [1] })
  const before = clone(f.workbook), selection = readSlicerRecordSelection(f.workbook, f.slicer), source = clone(f.cache().source), links = clone(f.cache().connections)
  const patch = { sortOrder: 'descending', customListSort: true, crossFilter: 'showItemsWithDataAtTop', hideItemsWithNoData: false }
  assert.equal(patchSlicerCacheOptions(f.workbook, f.slicer, patch), true)
  assert.deepEqual(readSlicerCacheOptions(f.workbook, f.shared), patch)
  assert.equal(f.cache().revision, before.slicerCaches.find(cache => cache.id === f.cache().id).revision + 1)
  for (const view of f.workbook.slicers) assert.equal(view.revision, before.slicers.find(prior => prior.id === view.id).revision + 1)
  assert.deepEqual(f.cache().source, source); assert.deepEqual(f.cache().connections, links)
  assert.deepEqual(readSlicerRecordSelection(f.workbook, f.slicer), selection)
  assert.deepEqual(f.workbook.pivotTables, before.pivotTables); assert.deepEqual(f.workbook.charts, before.charts)
  assert.deepEqual(f.workbook.tables, before.tables)
  const snapshot = readSlicerCacheOptions(f.workbook, f.shared); snapshot.sortOrder = 'ascending'
  assert.equal(readSlicerCacheOptions(f.workbook, f.slicer).sortOrder, 'descending')
  const after = clone(f.workbook); assert.equal(patchSlicerCacheOptions(f.workbook, f.shared, patch), false); assert.deepEqual(f.workbook, after)
})

for (const [key, value] of [['sortOrder', 'ascending'], ['customListSort', false], ['crossFilter', 'none'], ['hideItemsWithNoData', false]]) {
  test(`presence-aware ${key} authors explicit defaults, clears own undefined and preserves absent properties`, () => {
    const f = setup('pivot'); assert.deepEqual(readSlicerCacheOptions(f.workbook, f.slicer), {})
    assert.equal(patchSlicerCacheOptions(f.workbook, f.slicer, { [key]: value }), true)
    assert.deepEqual(readSlicerCacheOptions(f.workbook, f.shared), { [key]: value })
    const before = clone(f.workbook); assert.equal(patchSlicerCacheOptions(f.workbook, f.shared, {}), false); assert.deepEqual(f.workbook, before)
    assert.equal(patchSlicerCacheOptions(f.workbook, f.slicer, { [key]: undefined }), true)
    assert.deepEqual(readSlicerCacheOptions(f.workbook, f.shared), {}); assert.equal(f.cache().options, undefined)
    const after = clone(f.workbook); assert.equal(patchSlicerCacheOptions(f.workbook, f.slicer, { [key]: undefined }), false)
    assert.deepEqual(f.workbook, after)
  })
}

for (const invalid of [null, [], { sortOrder: 'natural' }, { customListSort: 'yes' }, { crossFilter: 'disabled' },
  { hideItemsWithNoData: 1 }, { source: {} }, { recordSelection: { mode: 'all' } }, { connections: [] }]) {
  test(`invalid cache patch ${JSON.stringify(invalid)} cannot partially mutate source, cache, views or reports`, () => {
    const f = setup('pivot'), before = clone(f.workbook)
    assert.throws(() => patchSlicerCacheOptions(f.workbook, f.slicer, invalid), /共享缓存选项/)
    assert.throws(() => normalizeSlicerCacheOptionsPatch(invalid), /共享缓存选项/)
    assert.deepEqual(f.workbook, before)
  })
}

for (const sortOrder of ['ascending', 'descending']) test(`${sortOrder} uses typed blank/number/date/string/boolean/error identity ordering`, () => {
  const input = members([{ type: 'string', value: '1' }, { type: 'error', value: '#N/A' }, { type: 'boolean', value: true },
    { type: 'date', value: 1 }, { type: 'blank' }, { type: 'number', value: 1 }]), before = clone(input)
  const types = ['blank', 'number', 'date', 'string', 'boolean', 'error']
  assert.deepEqual(applySlicerCacheItemOptions(input, { sortOrder }, 'en-US').map(item => item.value.type), sortOrder === 'ascending' ? types : types.reverse())
  assert.deepEqual(input, before)
})

for (const customListSort of [false, true]) for (const sortOrder of ['ascending', 'descending']) {
  test(`localized calendar custom=${customListSort}/${sortOrder} sorts month and weekday aliases before ordinary text`, () => {
    const input = members(['Sunday', 'March', 'Tuesday', 'February', 'Mon', 'Jan', 'Other'].map(value => ({ type: 'string', value })))
    const expected = customListSort ? ['Jan', 'February', 'March', 'Sunday', 'Mon', 'Tuesday', 'Other']
      : ['February', 'Jan', 'March', 'Mon', 'Other', 'Sunday', 'Tuesday']
    assert.deepEqual(applySlicerCacheItemOptions(input, { customListSort, sortOrder }, 'en-US').map(item => scalar(item.value)), sortOrder === 'ascending' ? expected : expected.reverse())
  })
}

for (const crossFilter of ['none', 'showItemsWithDataAtTop', 'showItemsWithNoData']) for (const hideItemsWithNoData of [false, true]) {
  test(`cross-filter ${crossFilter}/hide=${hideItemsWithNoData} partitions or hides only real no-data items`, () => {
    const input = members(['C', 'A', 'D', 'B'].map(value => ({ type: 'string', value })))
    const expected = crossFilter === 'none' ? ['A', 'B', 'C', 'D'] : hideItemsWithNoData ? ['C', 'D']
      : crossFilter === 'showItemsWithDataAtTop' ? ['C', 'D', 'A', 'B'] : ['A', 'B', 'C', 'D']
    const projected = applySlicerCacheItemOptions(input, { crossFilter, hideItemsWithNoData }, 'en-US')
    assert.deepEqual(projected.map(item => scalar(item.value)), expected)
    if (crossFilter === 'none') assert.ok(projected.every(item => item.hasData))
    assert.equal(input.filter(item => item.hasData).length, 2)
  })
}

for (const kind of kinds) for (const crossFilter of ['none', 'showItemsWithDataAtTop', 'showItemsWithNoData']) for (const hideItemsWithNoData of [false, true]) {
  test(`${kind}/${crossFilter}/hide=${hideItemsWithNoData} applies complete native/report availability before bounded slicing`, async () => {
    const f = setup(kind); northOnly(f)
    patchSlicerCacheOptions(f.workbook, f.slicer, { sortOrder: 'ascending', crossFilter, hideItemsWithNoData })
    const before = clone(f.workbook), window = await query(f), available = window.items.filter(item => item.hasData)
    assert.equal(window.total, crossFilter !== 'none' && hideItemsWithNoData ? 3 : 7)
    if (crossFilter === 'none') assert.equal(available.length, 7)
    else assert.deepEqual(available.map(item => scalar(item.value)), [1, '', '1'])
    if (crossFilter === 'showItemsWithDataAtTop' && !hideItemsWithNoData) assert.ok(window.items.slice(0, 3).every(item => item.hasData))
    assert.deepEqual(f.workbook, before)
    const service = new SlicerMemberQueryService({ computeClient: createSlicerMemberTestClient() })
    try {
      const page = await service.getItems(f.workbook, f.shared, { start: 1, count: 2 }, { unitId: f.workbook.id, baseRevision: 0, assertCurrent() {} })
      assert.equal(page.total, window.total); assert.deepEqual(page.items, window.items.slice(1, 3))
    } finally { service.dispose() }
  })
}

for (const kind of kinds) test(`${kind} own typed member selection does not make other available raw identities disappear`, async () => {
  const f = setup(kind); northOnly(f); setSlicerRecordSelection(f.workbook, f.slicer, { mode: 'include', values: [1] })
  patchSlicerCacheOptions(f.workbook, f.slicer, { crossFilter: 'showItemsWithNoData', hideItemsWithNoData: true })
  const selection = clone(readSlicerRecordSelection(f.workbook, f.slicer)), window = await query(f)
  assert.equal(window.total, 3); assert.deepEqual(window.items.map(item => scalar(item.value)), [1, '', '1'])
  assert.deepEqual(readSlicerRecordSelection(f.workbook, f.slicer), selection)
})

for (const kind of kinds) test(`${kind} JSON reopening retains authored cache options and exact shared member ordering`, async () => {
  const f = setup(kind); northOnly(f)
  const options = { sortOrder: 'descending', customListSort: false, crossFilter: 'showItemsWithDataAtTop', hideItemsWithNoData: false }
  patchSlicerCacheOptions(f.workbook, f.slicer, options)
  const expected = await query(f), reopened = workbookFromInput(JSON.parse(JSON.stringify(f.workbook)))
  const restored = { workbook: reopened, slicer: reopened.slicers[0], shared: reopened.slicers[1] }
  assert.deepEqual(readSlicerCacheOptions(reopened, restored.shared), options)
  assert.deepEqual(await query(restored), expected)
})

test('a zero-connection cache preserves both typed selection and shared options without touching detached reports', () => {
  const f = setup('pivot'); setSlicerRecordSelection(f.workbook, f.slicer, { mode: 'include', values: [1] })
  setSlicerConnections(f.workbook, f.slicer, { pivotTables: [] }); const before = clone(f.workbook.pivotTables)
  patchSlicerCacheOptions(f.workbook, f.shared, { crossFilter: 'none', sortOrder: 'descending' })
  assert.deepEqual(readSlicerRecordSelection(f.workbook, f.slicer), { mode: 'include', values: [1] })
  assert.deepEqual(f.workbook.pivotTables, before); assert.deepEqual(f.cache().connections, { pivotTables: [] })
})

test('options remain independent between canonical source caches sharing the same report', () => {
  const f = setup('pivot'), report = f.workbook.pivotTables[0]
  const separate = createSlicer(f.workbook, { id: 'region-options-view', unitId: f.workbook.id, hostSheetId: 'dashboard',
    pivotTableId: report.id, sourceFieldId: 'region-source', type: 'slicer' })
  patchSlicerCacheOptions(f.workbook, f.slicer, { sortOrder: 'descending', hideItemsWithNoData: true })
  assert.deepEqual(readSlicerCacheOptions(f.workbook, separate), {})
  assert.notEqual(getSlicerCache(f.workbook, separate).id, f.cache().id)
})

test('timeline controls refuse Slicer cache sorting patches before any resource changes', () => {
  const f = setup('pivot'), report = f.workbook.pivotTables[0]
  const timeline = createSlicer(f.workbook, { id: 'timeline-options-view', unitId: f.workbook.id, hostSheetId: 'dashboard',
    pivotTableId: report.id, sourceFieldId: 'date-source', type: 'timeline' }), before = clone(f.workbook)
  assert.throws(() => patchSlicerCacheOptions(f.workbook, timeline, { sortOrder: 'descending' }), /时间线/)
  assert.deepEqual(f.workbook, before)
})

test('calendar aliases with equal ranks keep source order and optional defaults remain authored separately', () => {
  const input = members(['January', 'Jan', 'February', 'Feb', 'Other'].map(value => ({ type: 'string', value })))
  assert.deepEqual(applySlicerCacheItemOptions(input, { customListSort: true }, 'en-US').map(item => scalar(item.value)), ['January', 'Jan', 'February', 'Feb', 'Other'])
  assert.deepEqual(effectiveSlicerCacheOptions({}), { sortOrder: 'ascending', customListSort: false,
    crossFilter: 'showItemsWithNoData', hideItemsWithNoData: false })
})

test('projected raw records are defensive and no-data hiding never alters retained selection inputs', () => {
  const input = members([{ type: 'date', value: 45000 }, { type: 'error', value: '#N/A' }]), before = clone(input)
  const projected = applySlicerCacheItemOptions(input, { hideItemsWithNoData: true })
  projected[0].value.value = 1; projected[0].hasData = false
  assert.deepEqual(input, before); assert.equal(projected.length, 1)
})
