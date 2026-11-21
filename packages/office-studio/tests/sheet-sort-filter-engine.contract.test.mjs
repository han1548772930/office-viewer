import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let api
let engine

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  engine = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts')
})

afterAll(async () => { await server?.close() })

function makeSheet(rows, columns = 3) {
  const sheet = api.createBlankSheet('Rows')
  sheet.rowCount = rows
  sheet.columnCount = columns
  return sheet
}

function setRows(sheet, matrix) {
  matrix.forEach((row, rowIndex) => row.forEach((value, colIndex) => {
    sheet.cells[`${rowIndex}:${colIndex}`] = api.createCell(rowIndex, colIndex, value)
  }))
  sheet.rowCount = Math.max(sheet.rowCount ?? 0, matrix.length)
  sheet.columnCount = Math.max(sheet.columnCount ?? 0, ...matrix.map((row) => row.length))
}

const fullRange = (endRow, endCol) => ({ startRow: 0, startCol: 0, endRow, endCol })

test('value filters compose by AND, keep original data, and clear without altering manual hidden rows', () => {
  const sheet = makeSheet(6)
  setRows(sheet, [
    ['Region', 'Sales', 'Quarter'],
    ['East', 10, 'Q1'],
    ['West', 20, 'Q1'],
    ['East', 30, 'Q2'],
    ['East', 5, 'Q2'],
    ['West', 40, 'Q2'],
  ])
  sheet.hiddenRows = [5]
  engine.setAutoFilter(sheet, fullRange(5, 2))
  engine.setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: ['East'] })
  engine.setFilterCriteria(sheet, 1, { kind: 'condition', operator: 'greaterThanOrEqual', value: 10 })

  assert.deepEqual(engine.getFilteredRows(sheet), [2, 4, 5])
  assert.equal(sheet.cells['2:1'].value, 20)
  assert.deepEqual(sheet.hiddenRows, [5])
  assert.deepEqual(sheet.autoFilter.columns[0].selectedItems, ['East'])

  engine.clearFilter(sheet, 1)
  assert.deepEqual(engine.getFilteredRows(sheet), [2, 5])
  engine.clearFilter(sheet)
  assert.deepEqual(engine.getFilteredRows(sheet), [])
  assert.deepEqual(sheet.hiddenRows, [5])
})

test('filter conditions handle text, blanks, dates as numeric serials, and errors', () => {
  const sheet = makeSheet(7, 2)
  setRows(sheet, [
    ['Item', 'Value'],
    ['Alpha', 1],
    ['Beta', 2],
    ['alphabet', 3],
    ['', 4],
    ['Date', 45292],
    ['Error', '#N/A'],
  ])
  sheet.cells['6:1'].dataType = 'e'
  engine.setAutoFilter(sheet, fullRange(6, 1))
  engine.setFilterCriteria(sheet, 0, { kind: 'condition', operator: 'contains', value: 'ALP' })
  assert.deepEqual(engine.getFilteredRows(sheet), [2, 4, 5, 6])

  engine.setFilterCriteria(sheet, 0, { kind: 'condition', operator: 'isBlank' })
  assert.deepEqual(engine.getFilteredRows(sheet), [1, 2, 3, 5, 6])
  engine.setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: ['Date'] })
  engine.setFilterCriteria(sheet, 1, { kind: 'condition', operator: 'between', value: 45_000, secondValue: 46_000 })
  assert.deepEqual(engine.getFilteredRows(sheet), [1, 2, 3, 4, 6])
  engine.setFilterCriteria(sheet, 1, { kind: 'condition', operator: 'isError' })
  // Column filters compose: the Date row is not an error, and the Error row
  // is still excluded by the first column until that criterion is cleared.
  assert.deepEqual(engine.getFilteredRows(sheet), [1, 2, 3, 4, 5, 6])
  engine.clearFilter(sheet, 0)
  assert.deepEqual(engine.getFilteredRows(sheet), [1, 2, 3, 4, 5])
})

test('invalid filter criteria and ranges fail before altering the previous model', () => {
  const sheet = makeSheet(3, 2)
  setRows(sheet, [['A', 'B'], [1, 2], [3, 4]])
  engine.setAutoFilter(sheet, fullRange(2, 1))
  engine.setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: [1] })
  const before = structuredClone(sheet.autoFilter)

  assert.throws(() => engine.setFilterCriteria(sheet, 1, { kind: 'condition', operator: 'between', value: 1 }), /two comparison values/)
  assert.deepEqual(sheet.autoFilter, before)
  assert.throws(() => engine.setFilterCriteria(sheet, 2, { kind: 'values', selectedItems: [] }), /inside the auto-filter range/)
  assert.deepEqual(sheet.autoFilter, before)
  assert.throws(() => engine.setAutoFilter(sheet, { startRow: 0, startCol: 0, endRow: 4, endCol: 1 }), /worksheet bounds/)
  assert.deepEqual(sheet.autoFilter, before)
})

test('sort is stable, multi-key, header aware, and moves complete sparse cell payloads', () => {
  const sheet = makeSheet(5, 4)
  setRows(sheet, [
    ['Team', 'Name', 'Score', 'Note'],
    ['B', 'Lin', 9, 'row-one'],
    ['A', 'Xu', 4, 'row-two'],
    ['A', 'An', 4, 'row-three'],
    ['B', 'He', 2, 'row-four'],
  ])
  sheet.cells['2:3'].style.fill = '#ffcc00'
  sheet.cells['2:3'].metadata = { preserved: true }
  sheet.cells['3:3'].formula = '=C4*$B$1'
  sheet.cells['2:3'].richText = [{ text: 'row-two', bold: true }]
  sheet.cells['2:3'].hyperlink = 'https://example.com/row-two'
  sheet.cells['4:2'].display = '2.00'
  const xuPayload = structuredClone(sheet.cells['2:3'])
  const anPayload = structuredClone(sheet.cells['3:3'])
  const unchangedHeader = structuredClone(sheet.cells['0:0'])

  engine.sortRange(sheet, fullRange(4, 3), [
    { column: 0, direction: 'ascending' },
    { column: 2, direction: 'ascending' },
    { column: 1, direction: 'ascending' },
  ], undefined, { hasHeader: true })

  assert.equal(sheet.cells['0:0'].value, 'Team')
  assert.deepEqual([sheet.cells['1:1'].value, sheet.cells['2:1'].value, sheet.cells['3:1'].value, sheet.cells['4:1'].value], ['An', 'Xu', 'He', 'Lin'])
  assert.deepEqual(sheet.cells['0:0'], unchangedHeader)
  assert.equal(sheet.cells['2:3'].style.fill, '#ffcc00')
  assert.deepEqual(sheet.cells['2:3'].metadata, { preserved: true })
  assert.deepEqual(sheet.cells['2:3'], xuPayload)
  assert.deepEqual(sheet.cells['1:3'], { ...anPayload, row: 1, id: '1:3' })
  assert.equal(sheet.cells['1:3'].formula, '=C4*$B$1')
  assert.equal(sheet.cells['2:3'].id, '2:3')
  assert.equal(sheet.cells['3:2'].value, 2)
  assert.equal(sheet.cells['3:2'].display, '2.00')

  const translated = makeSheet(3, 2)
  setRows(translated, [['Value', 'Formula'], [2, 22], [1, 11]])
  translated.cells['1:1'].formula = '=A2+$B$1'
  translated.cells['2:1'].formula = '=A3+$B$1'
  engine.sortRange(translated, fullRange(2, 1), [{ column: 0, direction: 'ascending' }], undefined, { hasHeader: true }, (formula, rowDelta, colDelta) => `${formula} (${rowDelta},${colDelta})`)
  assert.equal(translated.cells['1:1'].formula, '=A3+$B$1 (-1,0)')
  assert.equal(translated.cells['2:1'].formula, '=A2+$B$1 (1,0)')
})

test('sort follows Excel type order, keeps blanks last in both directions, and preserves stable ties', () => {
  const sheet = makeSheet(8, 2)
  setRows(sheet, [
    ['Key', 'Payload'],
    [true, 'boolean'],
    ['Z', 'text-z'],
    [2, 'number-2'],
    ['', 'blank-1'],
    [null, 'blank-2'],
    ['#N/A', 'error'],
    [1, 'number-1'],
  ])
  sheet.cells['6:0'].dataType = 'e'
  sheet.cells['6:0'].value = '#N/A'
  sheet.cells['6:0'].metadata = { error: '#N/A' }
  engine.sortRange(sheet, fullRange(7, 1), [{ column: 0, direction: 'ascending' }], undefined, { hasHeader: true })
  assert.deepEqual([1, 2, 3, 4, 5, 6, 7].map((row) => sheet.cells[`${row}:1`]?.value), ['number-1', 'number-2', 'text-z', 'boolean', 'error', 'blank-1', 'blank-2'])

  engine.sortRange(sheet, fullRange(7, 1), [{ column: 0, direction: 'descending' }], undefined, { hasHeader: true })
  assert.deepEqual([1, 2, 3, 4, 5, 6, 7].map((row) => sheet.cells[`${row}:1`]?.value), ['error', 'boolean', 'text-z', 'number-2', 'number-1', 'blank-1', 'blank-2'])
})

test('descending multi-key sorts keep blanks last per key and retain numeric and textual types', () => {
  const sheet = makeSheet(8, 3)
  setRows(sheet, [
    ['Group', 'Score', 'Payload'],
    ['A', null, 'blank-a'],
    ['A', 2, 'number-first'],
    ['A', '2', 'numeric-text'],
    ['A', 2, 'number-second'],
    ['A', '', 'blank-b'],
    [null, 999, 'blank-group'],
    ['B', 1, 'group-b'],
  ])
  sheet.cells['2:1'].display = '200%'
  const order = engine.getSortedRowOrder(sheet, fullRange(7, 2), [
    { column: 0, direction: 'descending' },
    { column: 1, direction: 'descending' },
  ], { hasHeader: true })
  assert.deepEqual(order, [6, 2, 1, 3, 0, 4, 5])
  engine.sortRange(sheet, fullRange(7, 2), [
    { column: 0, direction: 'descending' },
    { column: 1, direction: 'descending' },
  ], order, { hasHeader: true })
  assert.deepEqual([1, 2, 3, 4, 5, 6, 7].map((row) => sheet.cells[`${row}:2`]?.value), [
    'group-b', 'numeric-text', 'number-first', 'number-second', 'blank-a', 'blank-b', 'blank-group',
  ])
  assert.equal(sheet.cells['2:1'].value, '2')
  assert.equal(sheet.cells['3:1'].value, 2)
  assert.equal(sheet.cells['3:1'].display, '200%')
})

test('formula translation failure leaves every cell in the original range intact', () => {
  const sheet = makeSheet(4, 3)
  setRows(sheet, [['Key', 'Formula', 'Outside'], [3, 30, 'untouched'], [1, 10], [2, 20]])
  for (let row = 1; row <= 3; row += 1) sheet.cells[`${row}:1`].formula = `=A${row + 1}*10`
  const before = structuredClone(sheet)
  let translations = 0
  assert.throws(() => engine.sortRange(sheet, fullRange(3, 1), [{ column: 0, direction: 'ascending' }], undefined, { hasHeader: true }, () => {
    if (++translations === 2) throw new Error('Reference translation failed')
    return '=prepared-but-not-committed'
  }), /Reference translation failed/)
  assert.equal(translations, 2)
  assert.deepEqual(sheet, before)
})

test('manual and filtered rows keep their positions; Worker offsets expand to a full permutation', () => {
  const sheet = makeSheet(5, 2)
  setRows(sheet, [['Item', 'Number'], ['first', 30], ['hidden-by-filter', 20], ['hidden-manually', 40], ['last', 10]])
  sheet.hiddenRows = [3]
  engine.setAutoFilter(sheet, fullRange(4, 1))
  engine.setFilterCriteria(sheet, 0, { kind: 'values', selectedItems: ['first', 'hidden-manually', 'last'] })
  assert.deepEqual(engine.getFilteredRows(sheet), [2])
  const numeric = engine.getNumericSortColumns(sheet, fullRange(4, 1), [{ column: 1, direction: 'ascending' }], { hasHeader: true })
  assert.deepEqual(numeric.destinationOffsets, [0, 3])
  assert.deepEqual(numeric.columns, [[30, 10]])
  assert.deepEqual(engine.expandVisibleSortOrder(4, numeric.destinationOffsets, [1, 0]), [3, 1, 2, 0])
  const order = engine.getSortedRowOrder(sheet, fullRange(4, 1), [{ column: 1, direction: 'ascending' }], { hasHeader: true })
  assert.deepEqual(order, [3, 1, 2, 0])
  engine.sortRange(sheet, fullRange(4, 1), [{ column: 1, direction: 'ascending' }], order, { hasHeader: true })
  assert.equal(sheet.cells['1:0'].value, 'last')
  assert.equal(sheet.cells['2:0'].value, 'hidden-by-filter')
  assert.equal(sheet.cells['3:0'].value, 'hidden-manually')
  assert.equal(sheet.cells['4:0'].value, 'first')
  assert.deepEqual(sheet.hiddenRows, [3])
  assert.throws(() => engine.expandVisibleSortOrder(4, [0, 3], [1, 1]), /permutation/)
})

test('invalid sort permutations and merged-cell intersections cannot partially mutate a sheet', () => {
  const sheet = makeSheet(3, 2)
  setRows(sheet, [['Key', 'Payload'], [2, 'two'], [1, 'one']])
  const before = structuredClone(sheet.cells)
  assert.throws(() => engine.sortRange(sheet, fullRange(2, 1), [{ column: 0, direction: 'ascending' }], [0, 0]), /permutation/)
  assert.deepEqual(sheet.cells, before)
  sheet.mergedRanges.push({ startRow: 1, startCol: 0, endRow: 1, endCol: 1 })
  assert.throws(() => engine.sortRange(sheet, fullRange(2, 1), [{ column: 0, direction: 'ascending' }]), /merged/)
  assert.deepEqual(sheet.cells, before)
})
