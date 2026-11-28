import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeWorkbookAutoFilter } from '../src/parser/workbookAutoFilter.ts'
import { remapChartSourceDateFilter } from '../src/features/charts/chart-source-import-dates.ts'
import { assertSheetClipboardWorkbookResources } from '../src/features/sheet-workbench/sheet-clipboard-resource-validation.ts'

// Deferred drafts: parsing, resource transfer and calendar rebasing use the same
// complete condition model. Runtime checks remain in the unified verification.
const range = { startRow: 0, startCol: 0, endRow: 8, endCol: 0 }
const condition = (operator, value, extra = {}) => ({ kind: 'condition', operator, value, ...extra })
const decode = criterion => normalizeWorkbookAutoFilter({ range, columns: { 0: criterion }, filteredRows: [3] }).columns[0]
const compound = (join = 'and') => ({ kind: 'compound', join, conditions: [
  condition('greaterThan', 45000.25), condition('notBetween', 46000, { secondValue: 47000.5, caseSensitive: true }),
] })

test('the parser preserves AND/OR, one or two conditions, explicit null and independent operands', () => {
  for (const join of ['and', 'or']) {
    const input = compound(join), output = decode(input)
    assert.deepEqual(output, input)
    output.conditions[1].secondValue = 0
    assert.equal(input.conditions[1].secondValue, 47000.5)
    assert.deepEqual(decode({ kind: 'compound', join, conditions: [condition('equals', null)] }).conditions[0].value, null)
  }
})

test('invalid joins, counts, nesting, operand types and missing range operands fail at the parser boundary', () => {
  for (const input of [
    { ...compound(), join: 'xor' }, { ...compound(), conditions: [] },
    { ...compound(), conditions: [...compound().conditions, condition('equals', 1)] },
    { ...compound(), conditions: [compound()] }, { ...compound(), conditions: [{ kind: 'values', selectedItems: [1] }] },
    { ...compound(), conditions: [condition('between', 1)] },
    { ...compound(), conditions: [condition('equals', {})] },
  ]) assert.throws(() => decode(input))
})

test('cross-workbook calendar rebasing updates both compound predicates and retains non-date and color rules', () => {
  const input = compound('or'), output = remapChartSourceDateFilter(input, -1462, true, true)
  assert.deepEqual(output.conditions.map(item => [item.value, item.secondValue]), [[43538.25, undefined], [44538, 45538.5]])
  assert.equal(output.join, 'or'); assert.equal(output.conditions[1].caseSensitive, true)
  assert.equal(input.conditions[0].value, 45000.25)
  assert.equal(remapChartSourceDateFilter(input, -1462, true, false), input)
  const color = { kind: 'color', cellFillColors: [null, '#ff0000'] }
  assert.equal(remapChartSourceDateFilter(color, -1462, true, true), color)
})

function resources(criterion) {
  return { tables: [{ id: 'table', name: 'Table1', sheetId: 'sheet', range,
    columns: [{ id: 'column', name: 'Header' }], showHeaderRow: true, showTotalsRow: false, styleName: 'TableStyleMedium2',
    autoFilter: { range, columns: { 0: criterion }, filteredRows: [3] } }] }
}

test('table clipboard validation retains compound and color criteria and rejects malformed compound members', () => {
  assert.doesNotThrow(() => assertSheetClipboardWorkbookResources(resources(compound())))
  assert.doesNotThrow(() => assertSheetClipboardWorkbookResources(resources({ kind: 'color', cellFillColors: [null, '#ff0000'] })))
  const merged = resources(compound())
  merged.tables[0].autoFilter.filteredRows.push(9)
  assert.doesNotThrow(() => assertSheetClipboardWorkbookResources(merged))
  for (const filteredRows of [[3, 3], [0], [1_048_576]]) {
    const invalid = resources(compound()); invalid.tables[0].autoFilter.filteredRows = filteredRows
    assert.throws(() => assertSheetClipboardWorkbookResources(invalid))
  }
  for (const criterion of [{ ...compound(), join: 'xor' }, { ...compound(), conditions: [compound()] },
    { ...compound(), conditions: [condition('between', 1)] }, { kind: 'color', cellTextColors: [null] }]) {
    assert.throws(() => assertSheetClipboardWorkbookResources(resources(criterion)))
  }
})
