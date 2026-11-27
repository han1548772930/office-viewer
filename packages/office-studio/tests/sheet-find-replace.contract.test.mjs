import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetReplacePatches, findSheetMatches, nextSheetMatchIndex } from '../src/features/sheet-find-replace/sheet-find-replace.ts'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'

const sheet = (id, cells, overrides = {}) => ({ id, name: id, cells, rowCount: 20, columnCount: 10, ...overrides })
const query = (overrides = {}) => ({
  text: 'alpha', replacement: 'beta', findBy: 'value', direction: 'row', scope: 'sheet', caseSensitive: false, wholeCell: false,
  ...overrides,
})

test('find supports value and formula modes, whole-cell matching, hidden sheets, and workbook scope', () => {
  const sheets = [
    sheet('active', {
      A1: { value: 'Alpha beta', rawValue: 'Alpha beta' },
      B1: { value: '', formula: '=ALPHA+1', cachedValue: 'Alpha result', rawValue: 'Alpha result' },
      C1: { value: 'alphabet', rawValue: 'alphabet' },
    }),
    sheet('hidden', { A1: { value: 'alpha' } }, { visible: false }),
    sheet('other', { A1: { value: 'alpha' }, A2: { value: 'alpha' } }, { filteredRows: [1] }),
  ]

  assert.deepEqual(findSheetMatches(sheets, 'active', query()).map(match => [match.key, match.replaceable]), [
    ['A1', true], ['B1', false], ['C1', true],
  ])
  assert.deepEqual(findSheetMatches(sheets, 'active', query({ wholeCell: true })).map(match => match.key), [])
  assert.deepEqual(findSheetMatches(sheets, 'active', query({ findBy: 'formula', text: 'alpha' })).map(match => [match.key, match.replaceable]), [['B1', true]])
  assert.deepEqual(findSheetMatches(sheets, 'active', query({ scope: 'workbook' })).map(match => [match.sheetId, match.key]), [
    ['active', 'A1'], ['active', 'B1'], ['active', 'C1'], ['other', 'A1'],
  ])
})

test('find direction and selected-range scope produce stable match order', () => {
  const source = sheet('active', {
    A1: { value: 'alpha' }, B1: { value: 'alpha' }, A2: { value: 'alpha' }, B2: { value: 'alpha' }, C3: { value: 'alpha' },
  })
  const selected = [{ startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 }]
  assert.deepEqual(findSheetMatches([source], 'active', query(), selected).map(match => match.key), ['A2', 'B2'])
  const oneCellSelection = [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }]
  assert.deepEqual(findSheetMatches([source], 'active', query(), oneCellSelection).map(match => match.key), ['A1', 'B1', 'A2', 'B2', 'C3'])
  assert.deepEqual(findSheetMatches([source], 'active', query(), oneCellSelection, true).map(match => match.key), ['A1'])
  assert.deepEqual(findSheetMatches([source], 'active', query(), [], true), [])
  assert.deepEqual(findSheetMatches([source], 'active', query({ direction: 'column' })).map(match => match.key), ['A1', 'A2', 'B1', 'B2', 'C3'])
  const matches = findSheetMatches([source], 'active', query())
  assert.equal(nextSheetMatchIndex(matches, 'active', 'B1', 1), 2)
  assert.equal(nextSheetMatchIndex(matches, 'active', 'A1', -1), matches.length - 1)
  assert.equal(nextSheetMatchIndex(matches, 'missing', 'A1', -1), matches.length - 1)
})

test('replacement preserves formulas as formulas and does not replace cached formula results in value mode', () => {
  const source = sheet('active', {
    A1: { value: 'Alpha alpha', rawValue: 'Alpha alpha' },
    B1: { value: '', formula: '=ALPHA+1', cachedValue: 'Alpha result', rawValue: 'Alpha result' },
  })
  const matches = findSheetMatches([source], 'active', query())
  assert.deepEqual(createSheetReplacePatches([source], matches, query(), true, -1).map(patch => [patch.key, patch.cell.value, patch.cell.rawValue]), [
    ['A1', 'beta beta', 'beta beta'],
  ])

  const numberCell = sheet('active', { A1: { value: '12', rawValue: 12 } })
  const numberQuery = query({ text: '2', replacement: '3' })
  const numberMatch = findSheetMatches([numberCell], 'active', numberQuery)
  assert.equal(createSheetReplacePatches([numberCell], numberMatch, numberQuery, false, 0)[0]?.cell.rawValue, '13')

  const formulaQuery = query({ findBy: 'formula', text: 'alpha', replacement: 'gamma' })
  const formulaMatches = findSheetMatches([source], 'active', formulaQuery)
  const patches = createSheetReplacePatches([source], formulaMatches, formulaQuery, true, -1)
  assert.deepEqual(patches.map(patch => [patch.key, patch.cell.formula, patch.cell.cachedValue]), [['B1', '=gamma+1', undefined]])
})

test('single replacement from an unmatched cell follows column-priority search order', () => {
  const sheets = [sheet('active', { A2: { value: 'alpha' }, B2: { value: 'alpha' } })]
  const columnQuery = query({ direction: 'column' })
  const matches = findSheetMatches(sheets, 'active', columnQuery)
  assert.deepEqual(matches.map(match => match.key), ['A2', 'B2'])
  const targetIndex = nextSheetMatchIndex(matches, 'active', 'B1', 1, columnQuery.direction, sheets.map(sheet => sheet.id))
  assert.deepEqual(createSheetReplacePatches(sheets, matches, columnQuery, false, targetIndex).map(patch => [patch.key, patch.cell.value]), [
    ['B2', 'beta'],
  ])
})

test('single replacement from an unmatched sheet follows the complete workbook sheet order', () => {
  const sheets = [
    sheet('before', { A1: { value: 'alpha' } }),
    sheet('empty', {}),
    sheet('after', { A1: { value: 'alpha' } }),
  ]
  const workbookQuery = query({ scope: 'workbook' })
  const matches = findSheetMatches(sheets, 'empty', workbookQuery)
  assert.deepEqual(matches.map(match => match.sheetId), ['before', 'after'])
  const targetIndex = nextSheetMatchIndex(matches, 'empty', 'A1', 1, workbookQuery.direction, sheets.map(sheet => sheet.id))
  assert.deepEqual(createSheetReplacePatches(sheets, matches, workbookQuery, false, targetIndex).map(patch => [patch.sheetId, patch.key, patch.cell.value]), [
    ['after', 'A1', 'beta'],
  ])
})

test('single replacement targets only the selected replaceable match and whole-cell replacement replaces all content', () => {
  const source = sheet('active', { A1: { value: 'alpha alpha' }, B1: { value: 'alpha' } })
  const matches = findSheetMatches([source], 'active', query())
  assert.deepEqual(createSheetReplacePatches([source], matches, query(), false, 1).map(patch => patch.key), ['B1'])
  const fullCell = query({ wholeCell: true, text: 'alpha alpha', replacement: 'replaced' })
  const exactMatches = findSheetMatches([source], 'active', fullCell)
  assert.equal(createSheetReplacePatches([source], exactMatches, fullCell, false, 0)[0]?.cell.value, 'replaced')
})

test('view adapter preserves the explicit text type of a numeric replacement', () => {
  const workbook = createBlankWorkbook()
  const view = toExcelWorkbookView(workbook)
  const sheet = view.sheets[0]
  assert.ok(sheet)
  sheet.cells.A1 = { value: '13', rawValue: '13' }
  const updated = applyExcelWorkbookView(workbook, view)
  assert.equal(updated.sheets[0]?.cells['0:0']?.value, '13')
  assert.equal(typeof updated.sheets[0]?.cells['0:0']?.value, 'string')
})
