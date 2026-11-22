import assert from 'node:assert/strict'
import { test } from 'vitest'
import { replaceSheetCellInputSelection, sheetCellInputActionForKey, sheetCellNeedsMultilineEditor } from '../src/features/sheet-workbench/sheet-cell-edit-input.ts'
import { editSheetCellText, sheetCellInputText } from '../src/features/sheet-workbench/sheet-cell-edit.ts'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'

test('a line break replaces exactly the selected UTF-16 text and returns the next native caret', () => {
  const replacement = replaceSheetCellInputSelection('甲👩🏽‍💻乙丙', '\n', { start: 1, end: 8 })
  assert.deepEqual(replacement, { value: '甲\n乙丙', selection: { start: 2, end: 2 } })
  assert.deepEqual(replaceSheetCellInputSelection('a\nb', '\n', { start: 2, end: 2 }), {
    value: 'a\n\nb', selection: { start: 3, end: 3 },
  })
})

test('editor paste normalizes Office line endings without splitting one cell into a grid', () => {
  const replacement = replaceSheetCellInputSelection('head tail', 'a\tb\r\nc\rd\n', { start: 5, end: 9 })
  assert.equal(replacement.value, 'head a\tb\nc\nd\n')
  assert.deepEqual(replacement.selection, { start: 13, end: 13 })
  assert.deepEqual(replaceSheetCellInputSelection('value', '', { start: 1, end: 4 }), {
    value: 've', selection: { start: 1, end: 1 },
  })
})

test('native selection bounds are clamped and reversed selections retain the same replacement meaning', () => {
  assert.deepEqual(replaceSheetCellInputSelection('abcd', '中', { start: 3.8, end: 1.2 }), {
    value: 'a中d', selection: { start: 2, end: 2 },
  })
  assert.equal(replaceSheetCellInputSelection('abcd', '\n', { start: -2, end: 100 }).value, '\n')
  assert.equal(replaceSheetCellInputSelection('abcd', '\n', { start: NaN, end: Infinity }).value, 'abcd\n')
  assert.equal(replaceSheetCellInputSelection('abcd', '\n').value, 'abcd\n')
})

test('Alt+Enter, commit in place and directional commits have distinct actions', () => {
  for (const [event, action] of [
    [{ key: 'Enter', altKey: true }, 'line-break'], [{ key: 'Enter', altKey: true, shiftKey: true }, 'line-break'],
    [{ key: 'Enter', ctrlKey: true }, 'commit'], [{ key: 'Enter', metaKey: true }, 'commit'],
    [{ key: 'Enter' }, 'next-row'], [{ key: 'Enter', shiftKey: true }, 'previous-row'],
    [{ key: 'Tab' }, 'next-column'], [{ key: 'Tab', shiftKey: true }, 'previous-column'],
    [{ key: 'Escape' }, 'cancel'],
  ]) assert.equal(sheetCellInputActionForKey(event), action)
  for (const event of [{ key: 'Tab', altKey: true }, { key: 'Tab', ctrlKey: true },
    { key: 'Enter', ctrlKey: true, altKey: true }, { key: 'ArrowDown' }, { key: 'z', ctrlKey: true }]) {
    assert.equal(sheetCellInputActionForKey(event), 'native')
  }
})

test('local and native IME states suppress grid actions without blocking native composition', () => {
  for (const key of ['Enter', 'Tab', 'Escape']) {
    assert.equal(sheetCellInputActionForKey({ key, altKey: true }, true), 'native')
    assert.equal(sheetCellInputActionForKey({ key, isComposing: true }), 'native')
    assert.equal(sheetCellInputActionForKey({ key, keyCode: 229 }), 'native')
  }
})

test('the multiline bridge follows explicit wrap strategy, legacy wrap and retained session state', () => {
  assert.equal(sheetCellNeedsMultilineEditor({ value: '', wrap: true }, ''), true)
  assert.equal(sheetCellNeedsMultilineEditor({ value: '', wrap: true, wrapStrategy: 'clip' }, ''), false)
  assert.equal(sheetCellNeedsMultilineEditor({ value: '', wrapStrategy: 'wrap' }, ''), true)
  assert.equal(sheetCellNeedsMultilineEditor(undefined, 'a\r\nb'), true)
  assert.equal(sheetCellNeedsMultilineEditor(undefined, 'a\rb'), true)
  assert.equal(sheetCellNeedsMultilineEditor(undefined, 'a', true), true)
  assert.equal(sheetCellNeedsMultilineEditor(undefined, 'a'), false)
})

test('multiline text keeps its raw value, formula distinction and format across model/view save and reopen', () => {
  const workbook = createBlankWorkbook()
  workbook.sheets[0].cells['0:0'] = createCell(0, 0, 'original')
  const view = toExcelWorkbookView(workbook)
  const draft = replaceSheetCellInputSelection('12.5%', '\n3', { start: 5, end: 5 }).value
  view.sheets[0].cells.A1 = editSheetCellText({ ...view.sheets[0].cells.A1, wrapStrategy: 'wrap', wrap: true }, draft)
  const model = applyExcelWorkbookView(workbook, view)
  assert.equal(model.sheets[0].cells['0:0'].value, '12.5%\n3')
  assert.equal(model.sheets[0].cells['0:0'].dataType, 'text')
  assert.equal(model.sheets[0].cells['0:0'].style.wrapStrategy, 'wrap')
  const reopened = toExcelWorkbookView(JSON.parse(JSON.stringify(model)))
  assert.equal(sheetCellInputText(reopened.sheets[0].cells.A1), draft)
  const formula = editSheetCellText(undefined, '=SUM(\nA1:B1\n)')
  assert.equal(formula.formula, '=SUM(\nA1:B1\n)')
  assert.equal(sheetCellInputText(formula), formula.formula)
})
