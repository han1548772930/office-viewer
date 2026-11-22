import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetCellEditSession, getSheetCellEditConflict, resolveSheetDate1904 } from '../src/features/sheet-workbench/sheet-cell-edit-session.ts'

const number = (value = .125, pattern = '0.00%') => ({ value: String(value), rawValue: value, dataType: 'number', rawNumberFormat: pattern })
const context = (cell = number(), overrides = {}) => ({ documentId: 'book-a', sheetId: 'sheet-a', key: 'A1', date1904: false, cell, ...overrides })
const start = state => createSheetCellEditSession(state, { region: 'de-DE', currentYear: 2026 })

test('date system inheritance respects explicit false overrides and the default 1900 system', () => {
  assert.equal(resolveSheetDate1904(undefined, {}), false)
  assert.equal(resolveSheetDate1904({}, { date1904: true }), true)
  assert.equal(resolveSheetDate1904({ date1904: false }, { date1904: true }), false)
  assert.equal(resolveSheetDate1904({ date1904: true }, { date1904: false }), true)
})

test('the editor captures its owner, regional settings and source before external cell mutation', () => {
  const cell = number()
  const state = context(cell)
  const session = start(state)
  cell.rawValue = .5
  cell.value = '.5'
  assert.equal(session.region, 'de-DE')
  assert.equal(session.currentYear, 2026)
  assert.equal(session.source.value, .125)
  assert.equal(getSheetCellEditConflict(session, state), 'content')
})

test('a draft cannot commit into a different document, sheet, cell or date system', () => {
  const state = context()
  const session = start(state)
  assert.equal(getSheetCellEditConflict(session, { ...state, key: 'a1' }), undefined)
  for (const [changes, expected] of [
    [{ documentId: 'book-b' }, 'document'], [{ documentId: undefined }, 'document'],
    [{ sheetId: 'sheet-b' }, 'sheet'], [{ key: 'B1' }, 'cell'], [{ date1904: true }, 'date-system'],
  ]) assert.equal(getSheetCellEditConflict(session, { ...state, ...changes }), expected)
})

test('source value, explicit text edits and value-type changes conflict even when a stale display matches', () => {
  const state = context()
  const session = start(state)
  for (const cell of [
    { ...state.cell, rawValue: .5 }, { ...state.cell, value: '50%' },
    { ...state.cell, dataType: 'text' }, { ...state.cell, formula: '=A2/100' }, undefined,
  ]) assert.equal(getSheetCellEditConflict(session, { ...state, cell }), 'content')
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { ...state.cell, dataType: 'n' } }), undefined)
  const text = context({ value: '0125', rawValue: '0125', dataType: 'text' })
  assert.equal(getSheetCellEditConflict(start(text), { ...text, cell: { ...text.cell, dataType: 'force_string' } }), 'content')
})

test('number-format edits conflict but unrelated visual changes and presentation refreshes remain safe', () => {
  const state = context()
  const session = start(state)
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { ...state.cell, rawNumberFormat: '0.000%' } }), 'number-format')
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { ...state.cell, color: '#ff0000', background: '#ffffff', bold: true,
    fontFamily: 'Arial', fontSize: 14, displayValue: '12,50 %', cachedValue: 'irrelevant' } }), undefined)
})

test('formula recalculation can update cached values and types without invalidating a changed formula draft', () => {
  const state = context({ ...number(12.5, '0.00'), formula: '=SUM(B1:B2)', cachedValue: '12.50', cachedDataType: 'number' })
  const session = start(state)
  const recalculated = { ...state.cell, rawValue: null, value: '', dataType: 'error', cachedDataType: 'error', cachedValue: '#DIV/0!', displayValue: '#DIV/0!' }
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: recalculated }), undefined)
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { ...recalculated, formula: '=SUM(C1:C2)' } }), 'content')
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { ...recalculated, rawNumberFormat: '0%' } }), 'number-format')
})

test('absent blank cells and explicitly styled blanks share a baseline while empty text remains distinct', () => {
  const state = { ...context(), cell: undefined }
  const session = start(state)
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { value: '', color: '#ff0000' } }), undefined)
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { value: '', rawValue: null, dataType: 'blank' } }), undefined)
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { value: '', rawValue: '', dataType: 'text' } }), 'content')
  assert.equal(getSheetCellEditConflict(session, { ...state, cell: { value: '', rawNumberFormat: '@' } }), 'number-format')
})
