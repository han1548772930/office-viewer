import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createCell } from '../src/domain/model.ts'
import { compileSheetCellFilterCriteria, setAutoFilter, setFilterCriteria, getFilteredRows } from '../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { sheetFilterConditionDraft, sheetFilterConditionFromDraft, sheetCompoundFilterFromDrafts } from '../src/features/sheet-sort-filter/sheet-filter-condition-draft.ts'
import { matchesSheetFilterWildcard } from '../src/features/sheet-sort-filter/sheet-filter-wildcard.ts'

// Implementation-phase drafts; runtime verification is deferred.
const condition = (operator, value, extra = {}) => ({ kind: 'condition', operator, ...(value === undefined ? {} : { value }), ...extra })
function setup() {
  const sheet = createBlankSheet('Compound'); sheet.rowCount = 20; sheet.columnCount = 5
  for (const [row, value] of [[0, 'Value'], [1, 5], [2, 15], [3, 25], [4, null], [5, 'Alpha'], [6, 'Beta'], [7, '#N/A']]) sheet.cells[`${row}:0`] = createCell(row, 0, value)
  sheet.cells['7:0'].dataType = 'e'
  setAutoFilter(sheet, { startRow: 0, endRow: 7, startCol: 0, endCol: 0 })
  return sheet
}

test('AND and OR share the existing scalar predicates, preserving blank/error behavior and leaving source cells unchanged', () => {
  const sheet = setup(), cells = structuredClone(sheet.cells)
  setFilterCriteria(sheet, 0, { kind: 'compound', join: 'and', conditions: [condition('greaterThan', 10), condition('lessThan', 20)] })
  assert.deepEqual(getFilteredRows(sheet), [1, 3, 4, 5, 6, 7])
  setFilterCriteria(sheet, 0, { kind: 'compound', join: 'or', conditions: [condition('isBlank'), condition('isError')] })
  assert.deepEqual(getFilteredRows(sheet), [1, 2, 3, 5, 6])
  assert.deepEqual(sheet.cells, cells)
})

test('compound criteria capture detached conditions and retain per-condition case and two-boundary settings', () => {
  const first = condition('contains', 'Al', { caseSensitive: true }), second = condition('between', 5, { secondValue: 15 })
  const criteria = { kind: 'compound', join: 'or', conditions: [first, second] }, matches = compileSheetCellFilterCriteria(criteria)
  first.value = 'Beta'; second.secondValue = -1
  assert.equal(matches(createCell(0, 0, 'Alpha')), true)
  assert.equal(matches(createCell(0, 0, 'alpha')), false)
  assert.equal(matches(createCell(0, 0, 10)), true)
  assert.equal(matches(createCell(0, 0, 25)), false)
})

test('invalid compound cardinality, nested rules, joins and scalar settings fail before changing prior filter state', () => {
  const sheet = setup(), before = structuredClone(sheet.autoFilter)
  for (const rule of [
    { kind: 'compound', join: 'xor', conditions: [condition('isBlank')] },
    { kind: 'compound', join: 'and', conditions: [] },
    { kind: 'compound', join: 'or', conditions: [condition('isBlank'), condition('isBlank'), condition('isBlank')] },
    { kind: 'compound', join: 'and', conditions: [{ kind: 'color', cellFillColors: [] }] },
    { kind: 'compound', join: 'and', conditions: [{ kind: 'compound', join: 'or', conditions: [condition('isBlank')] }] },
    { kind: 'compound', join: 'and', conditions: [condition('equals', Number.NaN)] },
    { kind: 'compound', join: 'and', conditions: [condition('equals', 1, { caseSensitive: 'yes' })] },
  ]) {
    assert.throws(() => setFilterCriteria(sheet, 0, rule), TypeError)
    assert.deepEqual(sheet.autoFilter, before)
  }
})

test('the custom form maps None rows to clear, a single rule to the existing condition, and two rules to one compound model', () => {
  const none = sheetFilterConditionDraft(), one = sheetFilterConditionDraft(condition('greaterThan', 10))
  const two = sheetFilterConditionDraft(condition('lessThanOrEqual', 20, { caseSensitive: true }))
  assert.equal(sheetCompoundFilterFromDrafts('or', [none, none]), undefined)
  assert.deepEqual(sheetCompoundFilterFromDrafts('and', [one, none]), { kind: 'condition', operator: 'greaterThan', value: 10, caseSensitive: false })
  assert.deepEqual(sheetCompoundFilterFromDrafts('or', [none, two]), { kind: 'condition', operator: 'lessThanOrEqual', value: 20, caseSensitive: true })
  const compound = sheetCompoundFilterFromDrafts('and', [one, two])
  assert.equal(compound.kind, 'compound'); assert.equal(compound.join, 'and'); assert.equal(compound.conditions.length, 2)
  assert.throws(() => sheetCompoundFilterFromDrafts('and', [{ ...one, value: 'NaN' }, two]), /有效数值/)
  assert.throws(() => sheetCompoundFilterFromDrafts('and', [one]), TypeError)
})

test('condition form hydration retains comparison type, two boundaries and independent case settings', () => {
  const source = condition('between', 1.5, { secondValue: 3.5, caseSensitive: true })
  assert.deepEqual(sheetFilterConditionFromDraft(sheetFilterConditionDraft(source)), source)
  assert.deepEqual(sheetFilterConditionFromDraft(sheetFilterConditionDraft(condition('equals', false))),
    { kind: 'condition', operator: 'equals', value: false, caseSensitive: false })
  assert.deepEqual(sheetFilterConditionFromDraft(sheetFilterConditionDraft(condition('isNotBlank'))),
    { kind: 'condition', operator: 'isNotBlank', caseSensitive: false })
})

test('text conditions and compound children share fixed * and ? wildcard matching with literal regexp syntax and case control', () => {
  assert.equal(matchesSheetFilterWildcard('ALPHA', 'a?ph*'), true)
  assert.equal(matchesSheetFilterWildcard('ALPHA', 'a?ph*', true), false)
  assert.equal(matchesSheetFilterWildcard('a.b[1]', 'a.b[1]'), true)
  assert.equal(matchesSheetFilterWildcard('axb1', 'a.b[1]'), false)
  assert.equal(matchesSheetFilterWildcard('~abc', '~*'), true)
  assert.equal(matchesSheetFilterWildcard('abc', '~*'), false)
  const match = compileSheetCellFilterCriteria({ kind: 'compound', join: 'and', conditions: [condition('beginsWith', 'Al?'), condition('doesNotContain', '?et')] })
  assert.equal(match(createCell(1, 0, 'Alpha')), true)
  assert.equal(match(createCell(1, 0, 'Alphabet')), false)
  assert.equal(match(createCell(1, 0, 'Beta')), false)
})
