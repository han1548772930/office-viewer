import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { createSheetTable } from '../src/features/sheet-tables/sheet-tables.ts'

function sample() {
  const workbook = createBlankWorkbook()
  workbook.id = 'sheet-tables-contract'
  return workbook
}

test('table headers use stored numeric and Boolean values with the existing blank and duplicate-name rules', () => {
  const workbook = sample()
  const sheet = workbook.sheets[0]
  for (const [column, value] of [0, 12.5, false, true, ' TRUE ', '0', null].entries()) {
    sheet.cells[`0:${column}`] = createCell(0, column, value)
    sheet.cells[`0:${column}`].display = 'formatted display'
  }
  const original = cloneModel(workbook)
  const result = createSheetTable(workbook, { id: 'scalar-table', name: 'ScalarTable', sheetId: sheet.id,
    range: { startRow: 0, endRow: 2, startCol: 0, endCol: 7 } })
  assert.deepEqual(result.tables[0].columns.map(column => column.name), ['0', '12.5', 'FALSE', 'TRUE', 'TRUE_2', '0_2', 'Column7', 'Column8'])
  assert.deepEqual(workbook, original)
  assert.deepEqual(result.sheets, original.sheets)
})

test('declared Boolean and cached Boolean cells use TRUE/FALSE while forced strings retain their stored text', () => {
  const workbook = sample()
  const sheet = workbook.sheets[0]
  const cells = [
    { value: 0, dataType: 'boolean' },
    { value: 1, dataType: 'b' },
    { value: 'false', dataType: 'text' },
    { value: '0', dataType: 'force_string' },
    { value: 0, formula: '=FALSE()', dataType: 'formula', cachedDataType: 'bool' },
  ]
  for (const [column, cell] of cells.entries()) sheet.cells[`0:${column}`] = { ...createCell(0, column, cell.value), ...cell }
  const result = createSheetTable(workbook, { name: 'TypedHeaders', sheetId: sheet.id,
    range: { startRow: 0, endRow: 1, startCol: 0, endCol: 4 } })
  assert.deepEqual(result.tables[0].columns.map(column => column.name), ['FALSE', 'TRUE', 'false_2', '0', 'FALSE_3'])
  assert.equal(sheet.cells['0:4'].formula, '=FALSE()')
  assert.equal(sheet.cells['0:3'].dataType, 'force_string')
})

test('rich-text headers take their complete run text without modifying source runs or cached cell values', () => {
  const workbook = sample()
  const sheet = workbook.sheets[0]
  sheet.cells['0:0'] = { ...createCell(0, 0, 'stale'), richText: [
    { text: ' 部', fontFamily: 'Aptos', metadata: { vendor: true } },
    { text: '门 😀 ', hyperlink: 'https://example.test/departments' },
  ] }
  sheet.cells['0:1'] = { ...createCell(0, 1, 123), richText: [{ text: '部门 😀' }] }
  sheet.cells['0:2'] = { ...createCell(0, 2, 'cached header'), richText: [] }
  const original = cloneModel(workbook)
  const result = createSheetTable(workbook, { name: 'RichHeaders', sheetId: sheet.id,
    range: { startRow: 0, endRow: 1, startCol: 0, endCol: 2 } })
  assert.deepEqual(result.tables[0].columns.map(column => column.name), ['部门 😀', '部门 😀_2', 'Column3'])
  assert.deepEqual(workbook, original)
  assert.deepEqual(result.sheets[0].cells, original.sheets[0].cells)
})

test('table creation rejects contained, containing, edge and partial merge intersections without publishing a candidate', () => {
  const range = { startRow: 2, endRow: 4, startCol: 2, endCol: 4 }
  for (const merged of [
    { startRow: 3, endRow: 3, startCol: 3, endCol: 4 },
    { startRow: 0, endRow: 6, startCol: 0, endCol: 6 },
    { startRow: 0, endRow: 2, startCol: 1, endCol: 2 },
    { startRow: 4, endRow: 6, startCol: 4, endCol: 6 },
  ]) {
    const workbook = sample()
    workbook.sheets[0].mergedRanges = [merged]
    const original = cloneModel(workbook)
    assert.throws(() => createSheetTable(workbook, { name: 'RejectedTable', sheetId: workbook.sheets[0].id, range }), /intersect merged cells/)
    assert.deepEqual(workbook, original)
  }
})

test('table creation accepts immediately adjacent merges on all four sides', () => {
  const range = { startRow: 2, endRow: 4, startCol: 2, endCol: 4 }
  const workbook = sample()
  workbook.sheets[0].mergedRanges = [
    { startRow: 0, endRow: 1, startCol: 2, endCol: 4 },
    { startRow: 5, endRow: 6, startCol: 2, endCol: 4 },
    { startRow: 2, endRow: 4, startCol: 0, endCol: 1 },
    { startRow: 2, endRow: 4, startCol: 5, endCol: 6 },
  ]
  const result = createSheetTable(workbook, { name: 'AdjacentTable', sheetId: workbook.sheets[0].id, range })
  assert.deepEqual(result.tables[0].range, range)
  assert.deepEqual(result.sheets[0].mergedRanges, workbook.sheets[0].mergedRanges)
})

test('native snapshot reopen preserves generated headers and the existing strict naming contract', () => {
  const workbook = sample()
  const sheet = workbook.sheets[0]
  sheet.cells['0:0'] = createCell(0, 0, false)
  sheet.cells['0:1'] = createCell(0, 1, 0)
  const result = createSheetTable(workbook, { id: 'saved-table', name: ' SavedHeaders ', sheetId: sheet.id,
    range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 } })
  const reopened = workbookFromInput(JSON.parse(JSON.stringify(result)))
  assert.deepEqual(reopened.tables, result.tables)
  assert.equal(reopened.tables[0].name, 'SavedHeaders')
  assert.deepEqual(reopened.tables[0].columns.map(column => column.name), ['FALSE', '0'])
  const before = cloneModel(result)
  assert.throws(() => createSheetTable(result, { name: 'savedheaders', sheetId: sheet.id,
    range: { startRow: 3, endRow: 4, startCol: 0, endCol: 1 } }), /unique/)
  assert.throws(() => createSheetTable(result, { name: 'bad name', sheetId: sheet.id,
    range: { startRow: 3, endRow: 4, startCol: 0, endCol: 1 } }), /Table name/)
  assert.deepEqual(result, before)
})
