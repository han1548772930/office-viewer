import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createCell } from '../src/domain/model.ts'
import { readDataPanelSelectedValueKeys, readDataPanelValues } from '../src/features/sheet-sort-filter/sheet-data-panel-state.ts'
import { buildSheetFilterValueTree, searchSheetFilterValueTree, sheetFilterValueTreeSelection,
  toggleSheetFilterValueTreeNode } from '../src/features/sheet-sort-filter/sheet-filter-value-tree.ts'
import { setAutoFilter, setFilterCriteria } from '../src/features/sheet-sort-filter/sort-filter-engine.ts'

// Implementation-phase drafts; runtime verification is deferred.
const range = { startRow: 0, endRow: 6, startCol: 0, endCol: 1 }
function setup() {
  const sheet = createBlankSheet('Dates'); sheet.rowCount = 40; sheet.columnCount = 10
  for (const [row, value] of [[0, 'Date'], [1, 45292], [2, 45292.5], [3, 45293], [4, 45323], [5, '2024'], [6, null]]) {
    sheet.cells[`${row}:0`] = createCell(row, 0, value)
    sheet.cells[`${row}:1`] = createCell(row, 1, row === 0 ? 'Region' : row % 2 ? 'keep' : 'drop')
    if (row > 0 && row < 5) sheet.cells[`${row}:0`].style.numberFormat = 'yyyy-mm-dd hh:mm'
  }
  const workbook = { kind: 'workbook', id: 'date-tree', name: 'Dates', sheets: [sheet], activeSheetId: sheet.id, date1904: false, modified: false }
  return { sheet, workbook }
}
function candidates() { const h = setup(); return { ...h, values: readDataPanelValues(h.sheet, range, 0, h.workbook) } }

test('calendar trees group years, chronologically ordered months and days while retaining distinct original serials and scalar keys', () => {
  const { values } = candidates(), tree = buildSheetFilterValueTree(values)
  assert.equal(tree[0].key, 'date:2024'); assert.equal(tree[0].count, 4)
  assert.deepEqual(tree[0].children.map(month => [month.key, month.title, month.count]), [['date:2024-1', '一月', 3], ['date:2024-2', '二月', 1]])
  assert.deepEqual(tree[0].children[0].children.map(day => [day.key, day.count]), [['date:2024-1-1', 2], ['date:2024-1-2', 1]])
  assert.deepEqual(tree[0].children[0].children[0].valueKeys.map(key => values.find(value => value.key === key).value), [45292, 45292.5])
  assert.ok(tree.some(node => node.key === 'value:["string","2024"]'))
  assert.ok(tree.some(node => node.key === 'value:["object",null]'))
  assert.equal(new Set(tree.map(node => node.key)).size, tree.length)
})

test('year, month and day toggles preserve hidden scalar selections and expose partial day selection for two times', () => {
  const { values } = candidates(), tree = buildSheetFilterValueTree(values), initial = [values.find(value => value.value === '2024').key]
  const day = tree[0].children[0].children[0], firstTime = day.valueKeys[0]
  const partial = sheetFilterValueTreeSelection(tree, [...initial, firstTime])
  assert.ok(partial.halfChecked.includes(day.key)); assert.ok(partial.halfChecked.includes('date:2024'))
  const checked = toggleSheetFilterValueTreeNode(tree, [...initial, firstTime], day.key)
  assert.ok(day.valueKeys.every(key => checked.includes(key))); assert.ok(checked.includes(initial[0]))
  const month = toggleSheetFilterValueTreeNode(tree, checked, 'date:2024-1')
  assert.equal(month.length, 4)
  const cleared = toggleSheetFilterValueTreeNode(tree, month, 'date:2024-1')
  assert.deepEqual(cleared, initial)
  assert.throws(() => toggleSheetFilterValueTreeNode(tree, [], 'missing'), RangeError)
})

test('search uses OR across keywords, retains a matched year subtree and prunes unrelated month siblings', () => {
  const { values } = candidates(), tree = buildSheetFilterValueTree(values)
  assert.equal(searchSheetFilterValueTree(tree, '2024', values)[0], tree[0])
  const month = searchSheetFilterValueTree(tree, '二月', values)
  assert.equal(month.length, 1); assert.equal(month[0].children.length, 1); assert.equal(month[0].count, 1)
  assert.equal(month[0].children[0].key, 'date:2024-2')
  const multiple = searchSheetFilterValueTree(tree, '二月 空白', values)
  assert.equal(multiple.length, 2)
  const raw = searchSheetFilterValueTree(tree, '2024-01-02', values)
  assert.deepEqual(raw[0].valueKeys, values.filter(value => value.value === 45293).map(value => value.key))
  assert.deepEqual(searchSheetFilterValueTree(tree, 'not-present', values), [])
})

test('date-only and datetime grouping uses inherited effective formats while time, duration, text and rich content stay scalar', () => {
  const { sheet, workbook } = setup()
  sheet.cells['1:0'].authoredStyle = {}; workbook.defaultStyle = { numberFormat: 'yyyy-mm-dd' }
  sheet.cells['2:0'].style.numberFormat = 'h:mm:ss'
  sheet.cells['3:0'].style.numberFormat = '[h]:mm:ss'
  sheet.cells['4:0'].richText = [{ text: '2024-02-01' }]
  const values = readDataPanelValues(sheet, range, 0, workbook)
  assert.deepEqual(values.find(item => item.value === 45292).calendar, { year: 2024, month: 1, day: 1 })
  for (const value of [45292.5, 45293, 45323, '2024', null]) assert.equal(values.find(item => item.value === value).calendar, undefined)
})

test('1900 fictitious leap day and the 1904 epoch are decoded by the shared calendar without rewriting stored serials', () => {
  const { sheet, workbook } = setup()
  sheet.cells['1:0'].value = 60; sheet.cells['2:0'].value = 0; sheet.cells['3:0'].value = -1
  let values = readDataPanelValues(sheet, range, 0, workbook)
  assert.deepEqual(values.find(item => item.value === 60).calendar, { year: 1900, month: 2, day: 29 })
  assert.equal(values.find(item => item.value === 0).calendar, undefined)
  assert.equal(values.find(item => item.value === -1).calendar, undefined)
  workbook.date1904 = true
  values = readDataPanelValues(sheet, range, 0, workbook)
  assert.deepEqual(values.find(item => item.value === 0).calendar, { year: 1904, month: 1, day: 1 })
  assert.equal(sheet.cells['2:0'].value, 0)
})

test('merged date counts and candidate groups exclude other columns while ignoring the current column filter', () => {
  const { sheet, workbook } = setup()
  sheet.mergedRanges = [{ startRow: 1, endRow: 3, startCol: 0, endCol: 0 }]
  setAutoFilter(sheet, range, workbook)
  setFilterCriteria(sheet, 1, { kind: 'values', selectedItems: ['keep'] }, workbook)
  setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: [0] }, workbook)
  const values = readDataPanelValues(sheet, range, 0, workbook), tree = buildSheetFilterValueTree(values)
  assert.equal(values.find(item => item.value === 45292).count, 2)
  assert.deepEqual(tree[0].valueKeys, values.filter(item => item.value === 45292).map(item => item.key))
  assert.equal(tree[0].count, 2)
})

test('value-tree hydration follows include/exclude predicates instead of marking an excluded day checked', () => {
  const { sheet, workbook, values } = candidates()
  const include = { kind: 'values', selectedItems: [45292, 45292.5] }, exclude = { ...include, selectionMode: 'exclude' }
  const included = readDataPanelSelectedValueKeys(sheet, range, 0, values, include, workbook)
  const excluded = readDataPanelSelectedValueKeys(sheet, range, 0, values, exclude, workbook)
  assert.deepEqual(included, values.filter(item => item.value === 45292 || item.value === 45292.5).map(item => item.key))
  assert.deepEqual(excluded, values.filter(item => item.value !== 45292 && item.value !== 45292.5).map(item => item.key))
  const tree = buildSheetFilterValueTree(values)
  assert.ok(sheetFilterValueTreeSelection(tree, included).checked.includes('date:2024-1-1'))
  assert.ok(!sheetFilterValueTreeSelection(tree, excluded).checked.includes('date:2024-1-1'))
})
