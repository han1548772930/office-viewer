import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createCell } from '../src/domain/model.ts'
import { getRecordValueKey, normalizeRecordSelection, normalizeRecordValue } from '../src/domain/record-value.ts'
import { compilePivotRecordSelection } from '../src/features/pivot/services/pivot-filter-predicate.ts'
import { sheetCellRecordValue } from '../src/features/sheet-workbench/sheet-record-value.ts'
import { compileSheetCellFilterCriteria } from '../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { normalizeWorkbookAutoFilter } from '../src/parser/workbookAutoFilter.ts'
import { planSlicerRecordItemStates, slicerRecordSelectionItemState } from '../src/features/slicer/slicer-record-selection.ts'

// Drafts for shared raw identity and actual filter consumers. No test has been executed.
test('RecordValue keeps blank/string/number/date/boolean/error keys distinct and normalizes negative zero', () => {
  const values = [{ type: 'blank' }, { type: 'string', value: '' }, { type: 'string', value: '1' },
    { type: 'number', value: 1 }, { type: 'date', value: 1 }, { type: 'boolean', value: true },
    { type: 'string', value: '#N/A' }, { type: 'error', value: '#N/A' }]
  assert.equal(new Set(values.map(getRecordValueKey)).size, values.length)
  assert.equal(getRecordValueKey({ type: 'number', value: -0 }), '["number",0]')
  assert.equal(getRecordValueKey({ type: 'blank' }), '["blank"]')
  for (const input of [null, [], { type: 'blank', value: undefined }, { type: 'number', value: NaN },
    { type: 'date', value: Infinity }, { type: 'boolean', value: 1 }, { type: 'number', value: '1' }, { type: 'error', value: null }]) {
    assert.throws(() => normalizeRecordValue(input))
  }
})

test('RecordSelection validates real modes, owns returned members, and keeps distinct date/error conditions', () => {
  const input = { mode: 'include', values: [{ type: 'number', value: 1 }, { type: 'date', value: 1 },
    { type: 'number', value: 1 }, { type: 'error', value: '#N/A' }] }, result = normalizeRecordSelection(input)
  assert.equal(result.values.length, 3)
  input.values[0].value = 20; assert.equal(result.values[0].value, 1)
  for (const invalid of [{ mode: { toString() { return 'exclude' } }, values: [] }, { mode: 'all', values: [] },
    { mode: 'include' }, { mode: 'exclude', values: [], unknown: true }, { mode: 'include', values: [1] }]) {
    assert.throws(() => normalizeRecordSelection(invalid))
  }
})

test('typed native value filters distinguish a date serial and error from same-value number and string cells', () => {
  const numeric = createCell(1, 0, 1), date = createCell(2, 0, 1), text = createCell(3, 0, '#N/A'), error = createCell(4, 0, '#N/A')
  date.style.numberFormat = 'yyyy-mm-dd'; error.dataType = 'e'; error.metadata = { error: '#N/A' }
  const criteria = { kind: 'values', selectedItems: [], recordDate1904: false,
    recordSelection: { mode: 'include', values: [{ type: 'date', value: 1 }, { type: 'error', value: '#N/A' }] } }
  const normalized = normalizeWorkbookAutoFilter({ range: { startRow: 0, endRow: 4, startCol: 0, endCol: 0 },
    columns: { 0: criteria }, filteredRows: [] }).columns[0], matches = compileSheetCellFilterCriteria(normalized)
  assert.equal(matches(date), true); assert.equal(matches(numeric), false)
  assert.equal(matches(error), true); assert.equal(matches(text), false)
  const legacy = compilePivotRecordSelection({ mode: 'include', values: [1] })
  assert.equal(legacy(numeric.value, sheetCellRecordValue(numeric)), true)
  assert.equal(legacy(date.value, sheetCellRecordValue(date)), true)
  assert.deepEqual(normalized.recordSelection, criteria.recordSelection)
})

test('formula cached date/error types and 1904 serials do not read localized display labels', () => {
  const date = createCell(1, 0, 1), error = createCell(2, 0, '#DIV/0!')
  date.formula = '=1'; date.cachedDataType = 'n'; date.style.numberFormat = 'yyyy-mm-dd'; date.display = '1904年1月2日'
  error.formula = '=1/0'; error.cachedDataType = 'e'; error.display = '错误'; error.metadata = { error: '#DIV/0!' }
  assert.deepEqual(sheetCellRecordValue(date, true), { type: 'date', value: 1 })
  assert.deepEqual(sheetCellRecordValue(error), { type: 'error', value: '#DIV/0!' })
})

test('partial caption toggles retain typed date/error members while checked captions select their other identities', () => {
  const raw = [1, 1, '#N/A', '#N/A', 'other'], typed = [{ type: 'number', value: 1 }, { type: 'date', value: 1 },
    { type: 'string', value: '#N/A' }, { type: 'error', value: '#N/A' }, { type: 'string', value: 'other' }]
  const current = { mode: 'include', values: [typed[1], typed[3]] }
  assert.deepEqual(slicerRecordSelectionItemState(current, raw, typed), { checkedItems: [], partialItems: ['1', '#N/A'] })
  const next = planSlicerRecordItemStates(current, raw, ['other'], ['1', '#N/A'], typed)
  assert.deepEqual(next, { mode: 'include', values: [typed[1], typed[3], typed[4]] })
  assert.deepEqual(slicerRecordSelectionItemState(next, raw, typed), { checkedItems: ['other'], partialItems: ['1', '#N/A'] })
})
