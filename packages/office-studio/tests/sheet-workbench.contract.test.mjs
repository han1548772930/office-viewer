import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  calculateSelectionStatistics, createUniqueSheetName, formatSelectionStatistic,
  normalizeSheetViewState, parseSheetAddress, selectSelectionStatistic, validateSheetName, visibleSheetMenuEntries,
} from '../src/features/sheet-workbench/sheet-workbench.ts'

const bounds = { rowCount: 1_048_576, columnCount: 16_384 }
const range = (startRow, startColumn, endRow = startRow, endColumn = startColumn) => ({ startRow, startColumn, endRow, endColumn })
const sheet = (cells, overrides = {}) => ({ id: 's1', name: '数据', cells, rowCount: 100, columnCount: 26, ...overrides })

test('name box parses absolute and reversed A1 references without losing the logical range', () => {
  assert.deepEqual(parseSheetAddress(' $c$9:$a$2 ', bounds), {
    range: range(1, 0, 8, 2), address: 'A2:C9', anchor: 'A2',
  })
  assert.deepEqual(parseSheetAddress('xfd1048576', bounds), {
    range: range(1_048_575, 16_383), address: 'XFD1048576', anchor: 'XFD1048576',
  })
  assert.equal(parseSheetAddress('A1:A1', bounds)?.address, 'A1')
})

test('whole rows and columns stay as compact bounded ranges', () => {
  assert.deepEqual(parseSheetAddress('C:A', bounds), {
    range: range(0, 0, 1_048_575, 2), address: 'A:C', anchor: 'A1',
  })
  assert.deepEqual(parseSheetAddress('$9:$2', bounds), {
    range: range(1, 0, 8, 16_383), address: '2:9', anchor: 'A2',
  })
})

test('malformed, cross-sheet, named and out-of-bounds addresses never jump to an invented cell', () => {
  for (const input of ['', 'A0', 'A01', 'A-1', 'A1junk', 'XFE1', 'A1048577', 'A9007199254740993',
    'AAAAAAAAAAAAAA1', 'A1:B', 'A1 B2', 'A1,B2', '数据!A1', "'数据'!A1", 'Sales', 'R1C1', '0:2']) {
    assert.equal(parseSheetAddress(input, bounds), undefined, input)
  }
  assert.equal(parseSheetAddress('B2', { rowCount: 1, columnCount: 2 }), undefined)
  assert.equal(parseSheetAddress('A1', { rowCount: 0, columnCount: 2 }), undefined)
  assert.equal(parseSheetAddress('A1', { rowCount: Infinity, columnCount: 2 }), undefined)
})

test('selection statistics preserve raw scalar types, formula caches, errors and sparse blanks', () => {
  const statistics = calculateSelectionStatistics(sheet({
    A1: { value: '1,200.00', rawValue: 1200, displayValue: '1,200.00' },
    B1: { value: '123', rawValue: '123' },
    C1: { value: 'TRUE', rawValue: true },
    D1: { value: '', formula: '=A1/4', cachedValue: '300', rawValue: 300 },
    E1: { value: '', formula: '=1/0', cachedValue: '#DIV/0!' },
    F1: { value: '', rawValue: null, bold: true },
    G1: { value: '' },
    H1: { value: '', formula: '=UNRESOLVED(A1)' },
    I1: { value: '', formula: '=""', cachedValue: '' },
    J1: { value: ' 2.5e1 ' },
    K1: { value: '0.25', rawValue: 0.25, displayValue: '25%' },
    L1: { value: '#VALUE!' },
    M1: { value: 'FALSE', rawValue: false },
    N1: { value: '', rawValue: '' },
  }), [range(0, 0, 0, 25)])
  assert.equal(statistics.count, 4)
  assert.equal(statistics.countA, 11)
  assert.equal(statistics.sum, 1525.25)
  assert.equal(statistics.min, 0.25)
  assert.equal(statistics.max, 1200)
  assert.equal(statistics.average, 381.3125)
  assert.deepEqual(statistics.items.map(item => item.id), ['MAX', 'MIN', 'SUM', 'COUNTA', 'COUNT', 'AVERAGE'])
})

test('overlapping ranges count each populated cell once and exclude hidden or filtered rows', () => {
  const source = sheet({
    A1: { value: '1' }, B1: { value: '2' }, A2: { value: '100' },
    A3: { value: '1000' }, A4: { value: '4' }, B4: { value: '5' },
  }, { hiddenRows: [1], hiddenColumns: [1] })
  const snapshot = structuredClone(source)
  const result = calculateSelectionStatistics(source, [range(0, 0, 3, 1), range(0, 0, 3, 0)], { filteredRows: [2] })
  assert.equal(result.sum, 12)
  assert.equal(result.count, 4)
  assert.equal(result.countA, 4)
  assert.deepEqual(source, snapshot)
})

test('sparse whole-sheet statistics do not materialize the full selected range', () => {
  const result = calculateSelectionStatistics(sheet({
    A1: { value: '2' }, XFD1048576: { value: '3' },
  }, bounds), [range(0, 0, 1_048_575, 16_383), range(0, 0, 1_048_575, 16_383)])
  assert.equal(result.sum, 5)
  assert.equal(result.countA, 2)
})

test('footer hides singleton or blank statistics and falls back to count for text-only selections', () => {
  const single = calculateSelectionStatistics(sheet({ A1: { value: '9' } }), [range(0, 0)])
  assert.equal(single.sum, 9)
  assert.deepEqual(single.items, [])
  assert.equal(selectSelectionStatistic(single), undefined)
  const empty = calculateSelectionStatistics(sheet({}), [range(0, 0, 5, 5)])
  assert.equal(empty.min, null)
  assert.equal(empty.average, null)
  assert.deepEqual(empty.items, [])
  const text = calculateSelectionStatistics(sheet({ A1: { value: '苹果' }, B1: { value: '桃子' } }), [range(0, 0, 0, 1)])
  assert.deepEqual(text.items, [{ id: 'COUNTA', label: '计数', value: 2 }])
  assert.equal(selectSelectionStatistic(text, 'AVERAGE')?.id, 'COUNTA')
  const mixed = calculateSelectionStatistics(sheet({ A1: { value: '9' }, B1: { value: '桃子' } }), [range(0, 0, 0, 1)])
  assert.equal(selectSelectionStatistic(mixed)?.id, 'SUM')
  assert.equal(selectSelectionStatistic(mixed, 'COUNT')?.value, 1)
})

test('statistics normalize reversed and clipped selections and ignore invalid coordinates', () => {
  const source = sheet({ A1: { value: '2' }, B1: { value: '3' }, Z100: { value: '100' } })
  assert.equal(calculateSelectionStatistics(source, [range(1, 1, -3, -3)]).sum, 5)
  assert.equal(calculateSelectionStatistics(source, [range(NaN, 0, 10, 1)]).countA, 0)
  assert.equal(calculateSelectionStatistics(source, [range(200, 200, 500, 500)]).countA, 0)
})

test('sparse statistics reject malformed range keys and deduplicate case aliases', () => {
  const source = sheet({ A1: { value: '2' }, a1: { value: '2' }, B1: { value: '3' },
    'C1:C1': { value: '100' }, 'D:D': { value: '100' }, A0: { value: '100' },
    AA1: { value: '100' }, A101: { value: '100' },
  })
  const statistics = calculateSelectionStatistics(source, [range(0, 0, 99, 25)])
  assert.equal(statistics.sum, 5)
  assert.equal(statistics.countA, 2)
})

test('footer formats the selected item, leaving numeric counts unformatted by the cell pattern', () => {
  assert.equal(formatSelectionStatistic({ id: 'SUM', label: '求和', value: 15.5 }, { numberFormat: '0.00' }), '求和: 15.50')
  assert.equal(formatSelectionStatistic({ id: 'AVERAGE', label: '平均值', value: 0.125 }, { numberFormat: '0.0%' }), '平均值: 12.5%')
  assert.equal(formatSelectionStatistic({ id: 'COUNT', label: '数值计数', value: 2 }, { numberFormat: '0.00%' }), '数值计数: 2')
  assert.equal(formatSelectionStatistic({ id: 'MIN', label: '最小值', value: -0 }), '最小值: 0')
  assert.equal(formatSelectionStatistic({ id: 'SUM', label: '求和', value: 1e20 }), '求和: 100,000,000,000,000,000,000')
})

test('sheet names reject invalid, overlong and case-insensitively duplicated names', () => {
  const names = [{ id: '1', name: 'Sales' }, { id: '2', name: '数据' }]
  for (const name of ['', '   ', 'a'.repeat(32), "'名字", "名字'", 'a:b', 'a\\b', 'a/b', 'a?b', 'a*b', 'a[b', 'a]b', 'a\nb']) {
    assert.equal(validateSheetName(name, names).valid, false, name)
  }
  assert.equal(validateSheetName('sales', names).code, 'duplicate')
  assert.deepEqual(validateSheetName('sales', names, '1'), { valid: true })
  assert.deepEqual(validateSheetName("销售 '2026' 计划", names), { valid: true })
  assert.deepEqual(validateSheetName('a'.repeat(31), names), { valid: true })
})

test('new sheet naming is deterministic, case-insensitive and valid at the name-length boundary', () => {
  const names = [{ id: '1', name: '工作表1' }, { id: '2', name: '工作表3' }]
  assert.equal(createUniqueSheetName(names), '工作表2')
  assert.equal(createUniqueSheetName([{ id: '1', name: 'SALES1' }], 'Sales'), 'Sales2')
  const long = createUniqueSheetName([], 'a'.repeat(31))
  assert.equal(long.length, 31)
  assert.equal(validateSheetName(long, []).valid, true)
  assert.equal(createUniqueSheetName([], 'bad/name'), '工作表1')
})

test('all-sheets menu shows normal hidden entries while veryHidden stays undiscoverable', () => {
  const entries = visibleSheetMenuEntries([
    { id: 'a', name: '正常' },
    { id: 'b', name: '隐藏', visible: false },
    { id: 'c', name: '很隐藏', hiddenState: 'veryHidden' },
    { id: 'd', name: '很隐藏2', hiddenState: 2 },
    { id: 'e', name: '恢复', visible: false, hiddenState: 0 },
    { id: 'f', name: '隐藏2', visible: true, hiddenState: 'hidden' },
  ], 'a')
  assert.deepEqual(entries, [
    { id: 'a', name: '正常', hidden: false, selected: true },
    { id: 'b', name: '隐藏', hidden: true, selected: false },
    { id: 'e', name: '恢复', hidden: false, selected: false },
    { id: 'f', name: '隐藏2', hidden: true, selected: false },
  ])
})

test('view restoration preserves a valid active cell, range anchor and zoom without mutating the session', () => {
  const state = { key: 'b3', anchor: '$D$5', ranges: [range(2, 1, 4, 3)], zoom: 125 }
  const snapshot = structuredClone(state)
  assert.deepEqual(normalizeSheetViewState({ rowCount: 20, columnCount: 10 }, state), {
    key: 'B3', anchor: 'D5', ranges: [range(2, 1, 4, 3)], zoom: 125,
  })
  assert.deepEqual(state, snapshot)
})

test('view restoration clamps axis deletions and removes overlapping selections', () => {
  const restored = normalizeSheetViewState({ rowCount: 3, columnCount: 4 }, {
    key: 'Z100', anchor: 'Y90', ranges: [range(1, 1, 99, 25), range(1, 1, 2, 2)], zoom: 900,
  })
  assert.deepEqual(restored, { key: 'D3', anchor: 'D3', ranges: [range(1, 1, 2, 3)], zoom: 400 })
})

test('view restoration repairs invalid sessions and keeps active cell and anchor inside the selection', () => {
  assert.deepEqual(normalizeSheetViewState({ rowCount: 3, columnCount: 4 }), {
    key: 'A1', anchor: 'A1', ranges: [range(0, 0)], zoom: 100,
  })
  assert.deepEqual(normalizeSheetViewState({ rowCount: 3, columnCount: 4 }, {
    key: 'A0', anchor: 'D3', ranges: [range(1, 1)], zoom: NaN,
  }), { key: 'B2', anchor: 'B2', ranges: [range(1, 1)], zoom: 100 })
  assert.deepEqual(normalizeSheetViewState({ rowCount: 0, columnCount: 0 }, { zoom: 1 }), {
    key: 'A1', anchor: 'A1', ranges: [], zoom: 10,
  })
})
