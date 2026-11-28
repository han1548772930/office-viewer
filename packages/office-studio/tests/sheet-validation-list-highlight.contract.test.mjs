import assert from 'node:assert/strict'
import { afterEach, test, vi } from 'vitest'
import { sheetValidationListHighlightRange } from '../src/features/sheet-validation/validation-list-highlight.ts'
import { filterSheetValidationListOptions } from '../src/features/sheet-validation/validation-list-query.ts'

// Presentation/Unicode drafts; no test runner or real browser has been run.
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals() })
test.each([
  ['Alpha ALPHA', 'alpha', 0, 5], ['before 汉字 after', '汉', 7, 8], ['İX', 'x', 1, 2], ['İX', 'i', 0, 1],
  ['A\u0301B', '\u0301', 0, 2], ['x👩‍💻y', '💻', 1, 6], ['x👩‍💻y', '\udc69', 1, 6],
  ['x\ud800y', '\ud800', 1, 2], ['x\ny', '\n', 1, 2], ['ΟΣΑ ΟΣ', 'Σ', 1, 2],
])('the first match in %j for %j keeps original spelling and whole grapheme boundaries', (label, query, start, end) => {
  const range = sheetValidationListHighlightRange(label, query)
  assert.deepEqual(range, { start, end })
  assert.equal(label.slice(0, range.start) + label.slice(range.start, range.end) + label.slice(range.end), label)
  assert.equal(filterSheetValidationListOptions([{ value: label, label }], query).length, 1)
})

test.each([['Alpha', ''], ['Alpha', 'missing'], ['', 'a']])('empty or missing %j/%j queries do not invent a highlighted match', (label, query) => {
  assert.equal(sheetValidationListHighlightRange(label, query), undefined)
})

test('Turkish contextual contraction and Lithuanian expansion map a folded match back to its original glyphs', () => {
  const lower = String.prototype.toLocaleLowerCase
  const spy = vi.spyOn(String.prototype, 'toLocaleLowerCase').mockImplementation(function () { return lower.call(this, 'tr-TR') })
  assert.deepEqual(sheetValidationListHighlightRange('I\u0307X', 'x'), { start: 2, end: 3 })
  assert.deepEqual(sheetValidationListHighlightRange('I\u0307X', 'i'), { start: 0, end: 2 })
  assert.deepEqual(sheetValidationListHighlightRange('I IX', 'ı'), { start: 0, end: 1 })
  spy.mockImplementation(function () { return lower.call(this, 'lt-LT') })
  assert.deepEqual(sheetValidationListHighlightRange('I\u0301X', 'x'), { start: 2, end: 3 })
  assert.deepEqual(sheetValidationListHighlightRange('I\u0301X', '\u0307'), { start: 0, end: 2 })
})

test('the existing non-Segmenter fallback preserves surrogate pairs and original candidate spelling', () => {
  vi.stubGlobal('Intl', Object.create(Intl, { Segmenter: { value: undefined, configurable: true } }))
  assert.deepEqual(sheetValidationListHighlightRange('x😀y', '\ude00'), { start: 1, end: 3 })
  assert.deepEqual(sheetValidationListHighlightRange('İX', 'x'), { start: 1, end: 2 })
})
