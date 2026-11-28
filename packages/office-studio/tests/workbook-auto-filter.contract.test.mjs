import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeWorkbookAutoFilter } from '../src/parser/workbookAutoFilter.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { clearFilter } from '../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'

const range = { startRow: 0, startCol: 1, endRow: 9, endCol: 1 }
const filter = () => ({ range: { ...range }, columns: { 1: { kind: 'values', selectedItems: [1, '1', null, ''] } }, filteredRows: [2, 3, 5] })
function workbook(raw) {
  return normalizeWorkbook({ fileName: 'filtered.xlsx', format: 'xlsx', parserVersion: 'draft', payload: { data: { sheets: [raw] } } })
}

test('parser filter boundary preserves scalar types, explicit null, and independent cloned state', () => {
  const input = filter(), normalized = normalizeWorkbookAutoFilter(input)
  assert.deepEqual(normalized, input)
  normalized.columns[1].selectedItems.push(true)
  assert.deepEqual(input.columns[1].selectedItems, [1, '1', null, ''])
  const condition = normalizeWorkbookAutoFilter({ range, columns: { 1: { kind: 'condition', operator: 'equals', value: null, caseSensitive: true } }, filteredRows: [] })
  assert.equal(Object.hasOwn(condition.columns[1], 'value'), true)
  assert.equal(condition.columns[1].value, null)
})

test('invalid filter columns, non-scalar values, missing operands and duplicate data rows fail at the boundary', () => {
  for (const input of [
    { ...filter(), columns: { bad: { kind: 'values', selectedItems: [] } } },
    { ...filter(), columns: { 0: { kind: 'values', selectedItems: [] } } },
    { ...filter(), columns: { 1: { kind: 'values', selectedItems: [{}] } } },
    { ...filter(), columns: { 1: { kind: 'condition', operator: 'equals' } } },
    { ...filter(), columns: { 1: { kind: 'condition', operator: 'between', value: 1 } } },
    { ...filter(), filteredRows: [0] }, { ...filter(), filteredRows: [2, 2] },
  ]) assert.throws(() => normalizeWorkbookAutoFilter(input))
})

test('exact typed filter snapshots keep manual/filter overlap, zero axes and logical worksheet counts', () => {
  const model = workbook({ name: 'Filter', axisUnit: 'cssPixel', rowCount: 20, columnCount: 4, hiddenRows: [3], rowHeights: [{ index: 5, size: 0 }],
    autoFilter: filter(), sourceAutoFilter: { name: 'autoFilter' }, sourceAxisDimensions: { projectAxisExtension: { restored: true }, projectAutoFilterExtension: { restored: true } }, cells: [] })
  const sheet = model.sheets[0]
  assert.deepEqual(sheet.hiddenRows, [3])
  assert.deepEqual(sheet.autoFilter.filteredRows, [2, 3, 5])
  const view = toExcelWorkbookView(model)
  assert.equal(view.sheets[0].rowCount, 20)
  assert.equal(view.sheets[0].columnCount, 4)
  assert.deepEqual(view.sheets[0].filteredRows, [2, 3, 5])
  const reapplied = applyExcelWorkbookView(model, view).sheets[0]
  assert.deepEqual(reapplied.hiddenRows, [3])
  assert.deepEqual(reapplied.autoFilter.columns[1].selectedItems, [1, '1', null, ''])
  view.sheets[0].autoFilter.columns[1].selectedItems.push(true)
  assert.deepEqual(sheet.autoFilter.columns[1].selectedItems, [1, '1', null, ''])
  view.sheets[0].autoFilter = undefined
  assert.equal(applyExcelWorkbookView(model, view).sheets[0].autoFilter, undefined)
  clearFilter(sheet)
  assert.deepEqual(sheet.hiddenRows, [3])
  assert.deepEqual(sheet.autoFilter.filteredRows, [])
  assert.equal(sheet.rowHeights[5], 0)
})

test('standard filter imports infer failing rows without turning them into permanent manual hides', () => {
  const model = workbook({ name: 'Standard', axisUnit: 'cssPixel', hiddenRows: [2, 3], autoFilter: { range, columns: { 1: { kind: 'values', selectedItems: [1] } }, filteredRows: [2, 3] },
    sourceAutoFilter: { name: 'autoFilter' }, cells: [{ row: 1, col: 1, value: 1 }, { row: 2, col: 1, value: 2 }, { row: 3, col: 1, value: 1 }] })
  const sheet = model.sheets[0]
  assert.deepEqual(sheet.hiddenRows, [3])
  assert.equal(sheet.autoFilter.filteredRows.includes(2), true)
  clearFilter(sheet)
  assert.deepEqual(sheet.hiddenRows, [3])
  assert.deepEqual(sheet.autoFilter.filteredRows, [])
})

test('external color-filter hide reasons are reconciled only after cross-sheet conditional references exist', () => {
  const filterRange = { startRow: 0, endRow: 2, startCol: 0, endCol: 0 }
  const model = normalizeWorkbook({ fileName: 'conditional-filter.xlsx', format: 'xlsx', parserVersion: 'draft', payload: { data: {
    sheets: [{ name: 'Data', hiddenRows: [1, 2], sourceAutoFilter: { name: 'autoFilter' },
      autoFilter: { range: filterRange, columns: { 0: { kind: 'color', cellFillColors: ['#ff0000'] } }, filteredRows: [1] },
      cells: [{ row: 0, col: 0, value: 'Header' }, { row: 1, col: 0, value: 1 }, { row: 2, col: 0, value: 2 }],
      conditionalFormats: [{ id: 'linked-color', type: 'expression', operator: 'equal', formula1: '=A1>Limit!$A$1', ranges: [filterRange],
        style: { fill: '#ff0000' }, priority: 1, stopIfTrue: false }] },
    { name: 'Limit', cells: [{ row: 0, col: 0, value: 1 }] }],
  } } })
  const sheet = model.sheets[0]
  assert.deepEqual(sheet.autoFilter.filteredRows, [1])
  assert.deepEqual(sheet.hiddenRows, [2])
  clearFilter(sheet, undefined, model)
  assert.deepEqual(sheet.autoFilter.filteredRows, []); assert.deepEqual(sheet.hiddenRows, [2])
})

test('parser cache accepts a complete merged group past ref.endRow while retaining title, duplicate and absolute bounds checks', () => {
  const input = { range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 }, columns: {}, filteredRows: [1, 2, 3, 4] }
  assert.deepEqual(normalizeWorkbookAutoFilter(input).filteredRows, [1, 2, 3, 4])
  for (const filteredRows of [[0], [2, 2], [1_048_576]]) assert.throws(() => normalizeWorkbookAutoFilter({ ...input, filteredRows }))
})
