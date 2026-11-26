import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { normalizeWorkbookAutoFilter } from '../src/parser/workbookAutoFilter.ts'
import { setAutoFilter, setFilterCriteria } from '../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { getSlicerCache, getSlicerTableNativeFilter, normalizeSlicerCaches } from '../src/features/slicer/slicer-connections.ts'
import { readSlicerItemState, readSlicerRecordSelection, readSlicerSelection, setSlicerItemStates,
  setSlicerRecordSelection, synchronizeSlicerSelections } from '../src/features/slicer/slicer-engine.ts'
import { slicerTableRecordSelectionWorkbook } from './fixtures/pivot-filter-fixture.mjs'

// Native Table values and integration drafts only; all execution remains deferred.
const clone = value => structuredClone(value)
const table = workbook => workbook.tables.find(table => table.id === 'typed-table')
const view = workbook => workbook.slicers.find(control => control.id === 'typed-table-slicer')
const visible = workbook => Array.from({ length: 7 }, (_, index) => index + 3).filter(row => !table(workbook).autoFilter.filteredRows.includes(row))
const source = workbook => workbook.sheets.find(sheet => sheet.id === 'data')
const native = (workbook, criteria, column = 6) => {
  const target = table(workbook), sheet = { ...source(workbook), autoFilter: target.autoFilter ? clone(target.autoFilter) : undefined }
  if (!sheet.autoFilter) setAutoFilter(sheet, clone(target.range))
  setFilterCriteria(sheet, column, criteria); target.autoFilter = sheet.autoFilter
}

for (const [raw, key, row] of [[1, '1', 3], [true, 'true', 5], [null, '', 7]]) {
  test(`native Table include ${typeof raw}:${String(raw)} preserves collision types and partial views`, () => {
    const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook)
    setSlicerRecordSelection(workbook, control, { mode: 'include', values: [raw] })
    assert.deepEqual(visible(workbook), [row])
    assert.deepEqual(getSlicerTableNativeFilter(workbook, control), { kind: 'values', selectedItems: [raw] })
    for (const shared of workbook.slicers) {
      assert.deepEqual(readSlicerRecordSelection(workbook, shared), { mode: 'include', values: [raw] })
      assert.deepEqual(readSlicerItemState(workbook, shared), { checkedItems: [], partialItems: [key] })
    }
  })
}

test('Table exclude uses the native typed predicate and preserves a deleted raw member as a condition', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook), selection = { mode: 'exclude', values: [1, 'future'] }
  setSlicerRecordSelection(workbook, control, selection)
  assert.deepEqual(visible(workbook), [4, 5, 6, 7, 8, 9])
  assert.equal(table(workbook).autoFilter.columns[6].selectionMode, 'exclude')
  source(workbook).cells['3:6'].value = 'replacement'
  setSlicerRecordSelection(workbook, control, selection)
  assert.deepEqual(visible(workbook), [3, 4, 5, 6, 7, 8, 9])
  assert.deepEqual(readSlicerRecordSelection(workbook, control), selection)
  source(workbook).cells['3:6'].value = 1
  setSlicerRecordSelection(workbook, control, selection)
  assert.deepEqual(visible(workbook), [4, 5, 6, 7, 8, 9])
})

test('Table include retains unavailable typed values while current raw captions are only a mirror', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook), selection = { mode: 'include', values: [1, 'future'] }
  setSlicerRecordSelection(workbook, control, selection)
  source(workbook).cells['3:6'].value = 'replacement'; setSlicerRecordSelection(workbook, control, selection)
  assert.deepEqual(readSlicerRecordSelection(workbook, control), selection)
  assert.deepEqual(readSlicerSelection(workbook, control), []); assert.deepEqual(visible(workbook), [])
  source(workbook).cells['9:6'].value = 'future'; setSlicerRecordSelection(workbook, control, selection)
  assert.deepEqual(visible(workbook), [9]); assert.deepEqual(readSlicerSelection(workbook, control), ['future'])
})

test('Table display toggles preserve raw partial members from the native filter rather than the caption cache', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook)
  native(workbook, { kind: 'values', selectedItems: [1, true, null] })
  getSlicerCache(workbook, control).selectedItems = ['1', 'true', '']
  setSlicerItemStates(workbook, control, ['beta'], ['1', 'true', ''])
  assert.deepEqual(table(workbook).autoFilter.columns[6].selectedItems, [1, true, null, 'beta'])
  assert.deepEqual(visible(workbook), [3, 5, 7, 9])
  const before = clone(workbook); setSlicerItemStates(workbook, control, ['beta'], ['1', 'true', '']); assert.deepEqual(workbook, before)
})

test('external native values update the typed cache even when their string mirrors are equal', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook)
  setSlicerRecordSelection(workbook, control, { mode: 'include', values: [1] })
  const before = clone(workbook); native(workbook, { kind: 'values', selectedItems: ['1'] }); synchronizeSlicerSelections(workbook, before)
  assert.deepEqual(getSlicerCache(workbook, control).recordSelection, { mode: 'include', values: ['1'] })
  assert.deepEqual(readSlicerSelection(workbook, control), ['1']); assert.deepEqual(visible(workbook), [4])
  assert.ok(workbook.slicers.every(shared => shared.revision > before.slicers.find(prior => prior.id === shared.id).revision))
})

test('native condition restrictions survive synchronization and a typed setter replaces only its own column', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook)
  setSlicerRecordSelection(workbook, control, { mode: 'include', values: [1] })
  native(workbook, { kind: 'condition', operator: 'greaterThan', value: 25 }, 5)
  native(workbook, { kind: 'condition', operator: 'contains', value: 'true' }, 6)
  const before = clone(workbook); synchronizeSlicerSelections(workbook, before)
  assert.deepEqual(workbook, before)
  assert.throws(() => readSlicerRecordSelection(workbook, control), /原生条件筛选/)
  assert.deepEqual(readSlicerItemState(workbook, control), { checkedItems: ['true'], partialItems: [] })
  assert.deepEqual(readSlicerSelection(workbook, control), ['true'])
  setSlicerRecordSelection(workbook, control, { mode: 'include', values: [true] })
  assert.deepEqual(table(workbook).autoFilter.columns[5], { kind: 'condition', operator: 'greaterThan', value: 25 })
  assert.deepEqual(visible(workbook), [5])
  setSlicerRecordSelection(workbook, control, { mode: 'all' })
  assert.equal(table(workbook).autoFilter.columns[6], undefined)
  assert.deepEqual(visible(workbook), [5, 6, 7, 8, 9])
})

test('a native condition partial item survives another Canvas toggle before the explicit typed replacement', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook)
  native(workbook, { kind: 'condition', operator: 'equals', value: 1 })
  const before = clone(workbook)
  assert.deepEqual(readSlicerItemState(workbook, control), { checkedItems: [], partialItems: ['1'] })
  assert.throws(() => readSlicerRecordSelection(workbook, control), /原生条件筛选/)
  assert.deepEqual(workbook, before)
  setSlicerItemStates(workbook, control, ['beta'], ['1'])
  assert.deepEqual(getSlicerTableNativeFilter(workbook, control), { kind: 'values', selectedItems: [1, 'beta'] })
  assert.deepEqual(visible(workbook), [3, 9])
})

test('Table include-empty, exclude-empty and checked-all have distinct native effects', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook)
  setSlicerRecordSelection(workbook, control, { mode: 'include', values: [] })
  assert.deepEqual(visible(workbook), []); assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'include', values: [] })
  setSlicerRecordSelection(workbook, control, { mode: 'exclude', values: [] })
  assert.equal(table(workbook).autoFilter.columns[6], undefined); assert.deepEqual(visible(workbook), [3, 4, 5, 6, 7, 8, 9])
  setSlicerRecordSelection(workbook, control, { mode: 'include', values: [1, '1', true, 'true', null, '', 'beta'] })
  assert.equal(readSlicerRecordSelection(workbook, control).mode, 'include')
  setSlicerItemStates(workbook, control, ['1', 'true', '', 'beta'], [])
  assert.equal(table(workbook).autoFilter.columns[6], undefined); assert.equal(readSlicerRecordSelection(workbook, control), undefined)
})

test('Table JSON reopening and the parser value-filter boundary preserve native exclude and scalar types', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook), selection = { mode: 'exclude', values: [1, true, null, 'future'] }
  setSlicerRecordSelection(workbook, control, selection)
  const normalized = normalizeWorkbookAutoFilter(JSON.parse(JSON.stringify(table(workbook).autoFilter)))
  assert.deepEqual(normalized.columns[6], { kind: 'values', selectedItems: [1, true, null, 'future'], selectionMode: 'exclude' })
  const reopened = workbookFromInput(JSON.parse(JSON.stringify(workbook)))
  assert.deepEqual(normalizeSlicerCaches(reopened), [])
  assert.deepEqual(readSlicerRecordSelection(reopened, view(reopened)), selection)
  assert.deepEqual(visible(reopened), [4, 6, 8, 9])
})

test('Table readers ignore stale caption cache values and their native snapshots are defensive', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook)
  native(workbook, { kind: 'values', selectedItems: [true] })
  getSlicerCache(workbook, control).recordSelection = { mode: 'include', values: ['true'] }
  const state = getSlicerTableNativeFilter(workbook, control); state.selectedItems.push('true')
  assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'include', values: [true] })
  assert.deepEqual(readSlicerItemState(workbook, control), { checkedItems: [], partialItems: ['true'] })
})

test('legacy native values remain include without a selection mode and reject invalid modes before mutation', () => {
  const workbook = slicerTableRecordSelectionWorkbook(), control = view(workbook)
  native(workbook, { kind: 'values', selectedItems: [1] })
  assert.deepEqual(visible(workbook), [3]); assert.deepEqual(readSlicerRecordSelection(workbook, control), { mode: 'include', values: [1] })
  const before = clone(workbook)
  assert.throws(() => setSlicerRecordSelection(workbook, control, { mode: 'bad', values: [1] }), /Invalid pivot/)
  assert.throws(() => native(workbook, { kind: 'values', selectedItems: [1], selectionMode: 'bad' }), /selection mode/)
  assert.deepEqual(workbook, before)
})
