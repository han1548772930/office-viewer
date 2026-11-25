import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { toExcelCellView } from '../src/services/excelViewAdapter.ts'
import { createSheetCellEditSession } from '../src/features/sheet-workbench/sheet-cell-edit-session.ts'
import { sheetCellInputText } from '../src/features/sheet-workbench/sheet-cell-edit.ts'
import { sheetCellFontValues } from '../src/features/sheet-workbench/sheet-cell-rich-text-style.ts'
import { formatSheetCellRichTextRange, readSheetCellInlineFormats, replaySheetCellInlineFormats, SheetCellInlineFormatDraft } from '../src/features/sheet-workbench/sheet-cell-inline-format.ts'
import { planSheetCellInlineFontFormat, planSheetCellInlineStyle } from '../src/features/sheet-workbench/sheet-cell-inline-format-command.ts'
import { planSheetCellInput } from '../src/features/sheet-workbench/sheet-cell-input-command.ts'

// Draft: model, operation and input-commit contracts; not executed.
const defaults = { region: 'en-US', currentYear: 2026, userId: 'local-user' }
const trace = (source, operations = [], sessionId = 'inline-draft') => ({ version: 1, sessionId, source, operations })
const format = (at, input, start, end, field, value) => ({ at, input, start, end, field, value })
function sample(text = 'abcd') {
  const book = createBlankWorkbook(); book.id = 'inline-book'; book.calculationMode = 'manual'
  const sheet = book.sheets[0]; sheet.id = 'inline-sheet'; book.activeSheetId = sheet.id; sheet.rowCount = 20; sheet.columnCount = 8
  const cell = createCell(0, 0, text); cell.metadata = { vendor: { keep: true } }
  sheet.cells['0:0'] = cell
  return book
}
function editInput(book, operations = [], input = sheetCellInputText(toExcelCellView(book.sheets[0].cells['0:0']))) {
  const cell = toExcelCellView(book.sheets[0].cells['0:0'])
  const session = createSheetCellEditSession({ documentId: book.id, sheetId: book.sheets[0].id, key: 'A1',
    date1904: book.date1904 ?? false, cell }, defaults)
  return { unitId: book.id, sheetId: book.sheets[0].id, row: 0, column: 0, input, session,
    editTrace: trace(sheetCellInputText(cell), operations, session.sessionId) }
}

test('partial ranges split source IDs once and retain links, unknown fields and caller ownership', () => {
  const runs = [{ id: 'source', text: 'abcdef', font: { italic: true, vendor: 7 }, hyperlink: 'https://example.test', vendor: { keep: 1 } }]
  const before = structuredClone(runs), result = formatSheetCellRichTextRange(runs, 'abcdef', 1, 4, 'bold', true)
  assert.deepEqual(result.map(run => run.text), ['a', 'bcd', 'ef'])
  assert.deepEqual(result.map(run => run.id), ['source', undefined, undefined])
  assert.equal(result[1].font.bold, true); assert.equal(result[0].font.bold, undefined); assert.equal(result[2].font.bold, undefined)
  assert.ok(result.every(run => run.hyperlink === 'https://example.test' && run.font.italic && run.font.vendor === 7))
  result[1].vendor.keep = 9; assert.deepEqual(runs, before)
})

test('reset affects only selected authored font fields without removing empty extension markers', () => {
  const runs = [{ id: 'one', text: 'ab', color: '#ff0000', font: { color: '#ff0000', italic: true } },
    { id: 'marker', text: '', vendor: true }, { id: 'two', text: 'cd', font: { color: '#00ff00' } }]
  const result = formatSheetCellRichTextRange(runs, 'abcd', 0, 2, 'color', null)
  assert.equal(result[0].color, undefined); assert.deepEqual(result[0].font, { italic: true })
  assert.deepEqual(result[1], runs[1]); assert.deepEqual(result[2], runs[2])
})

test('untrusted style traces reject stale text, bad ordering, invalid values and split surrogate offsets', () => {
  const source = 'A🙂B', journal = trace(source, [{ start: 3, end: 3, text: 'X', deletedText: '' }])
  for (const edit of [format(0, source, 2, 3, 'bold', true), format(0, 'wrong', 0, 1, 'bold', true),
    format(2, source, 0, 1, 'bold', true), format(0, source, 0, 1, 'size', Infinity),
    format(0, source, 0, 1, 'bold', 'yes'), format(0, source, 0, 1, 'fill', '#ff0000')]) {
    assert.throws(() => readSheetCellInlineFormats([edit], journal))
  }
  assert.throws(() => readSheetCellInlineFormats([format(1, 'A🙂XB', 0, 1, 'bold', true), format(0, source, 0, 1, 'italic', true)], journal))
  assert.equal(formatSheetCellRichTextRange([{ text: source }], source, 2, 3, 'bold', true), undefined)
  assert.equal(readSheetCellInlineFormats([format(0, source, 1, 3, 'bold', true)], journal).length, 1)
})

test('range formatting and later typing replay in chronological source coordinates', () => {
  const cell = { value: 'abcd', rawValue: 'abcd', richText: [{ id: 'one', text: 'ab' }, { id: 'two', text: 'cd', font: { italic: true } }] }
  const journal = trace('abcd', [{ start: 2, end: 2, text: 'X', deletedText: '', origin: 'beforeinput' }])
  const edits = [format(0, 'abcd', 0, 2, 'bold', true), format(1, 'abXcd', 2, 3, 'color', '#ff0000')]
  const replay = replaySheetCellInlineFormats(cell, journal, readSheetCellInlineFormats(edits, journal), { numberFormat: 'General' })
  assert.equal(replay.cell.value, 'abXcd')
  assert.deepEqual(sheetCellFontValues(replay.cell, 'bold', { start: 0, end: 3 }), [true, true])
  assert.deepEqual(sheetCellFontValues(replay.cell, 'color', { start: 2, end: 3 }), ['#ff0000'])
  assert.deepEqual(sheetCellFontValues(replay.cell, 'italic', { start: 3, end: 5 }), [true])
  assert.deepEqual(cell.richText, [{ id: 'one', text: 'ab' }, { id: 'two', text: 'cd', font: { italic: true } }])
})

test('caret formatting changes typing preference and styles only text subsequently inserted there', () => {
  const cell = { value: 'abcd', rawValue: 'abcd', richText: [{ id: 'left', text: 'ab', font: { italic: true } }, { id: 'right', text: 'cd' }] }
  const edits = [format(0, 'abcd', 2, 2, 'family', 'Arial')]
  const initial = replaySheetCellInlineFormats(cell, trace('abcd'), edits, { numberFormat: 'General' })
  assert.deepEqual(initial.cell.richText, cell.richText); assert.equal(initial.typing.font.family, 'Arial')
  const replay = replaySheetCellInlineFormats(cell, trace('abcd', [{ start: 2, end: 2, text: 'X', deletedText: '' }]), edits, { numberFormat: 'General' })
  assert.deepEqual(sheetCellFontValues(replay.cell, 'family', { start: 2, end: 3 }), ['Arial'])
  assert.deepEqual(sheetCellFontValues(replay.cell, 'italic', { start: 2, end: 3 }), [true])
  assert.deepEqual(sheetCellFontValues(replay.cell, 'family', { start: 3, end: 5 }), ['Aptos'])
  assert.equal(replay.typing.position, 3)
})

test('editor apostrophe escapes and literal apostrophes under text format map to different raw intervals', () => {
  const escaped = { value: 'abc', rawValue: 'abc', dataType: 'force_string', richText: [{ id: 'source', text: 'abc' }] }
  const replay = replaySheetCellInlineFormats(escaped, trace("'abc"), [format(0, "'abc", 2, 3, 'bold', true)], { numberFormat: 'General' })
  assert.equal(replay.cell.value, 'abc'); assert.equal(replay.cell.richText.map(run => run.text).join(''), 'abc')
  assert.deepEqual(sheetCellFontValues(replay.cell, 'bold', { start: 1, end: 2 }), [true])
  const literal = { value: "'abc", rawValue: "'abc", richText: [{ id: 'literal', text: "'abc" }] }
  const text = replaySheetCellInlineFormats(literal, trace("'abc"), [format(0, "'abc", 0, 1, 'italic', true)], { numberFormat: '@' })
  assert.equal(text.cell.richText[0].text, "'"); assert.equal(text.cell.richText[0].font.italic, true)
})

test('draft style undo and redo follow native journal prefixes and drop abandoned future branches', () => {
  const draft = new SheetCellInlineFormatDraft(), initial = trace('abcd')
  const first = format(0, 'abcd', 0, 2, 'bold', true); draft.append(initial, first)
  assert.equal(draft.undo(initial), true); assert.deepEqual(draft.formats(initial), []); assert.equal(draft.canRedo(initial), true)
  assert.equal(draft.redo(initial), true)
  const edited = trace('abcd', [{ start: 2, end: 2, text: 'X', deletedText: '' }])
  draft.append(edited, format(1, 'abXcd', 2, 3, 'italic', true))
  assert.equal(draft.formats(edited).length, 2)
  draft.reconcile(initial); assert.deepEqual(draft.formats(initial), [first])
  const branch = trace('abcd', [{ start: 2, end: 2, text: 'Y', deletedText: '' }])
  draft.reconcile(branch); assert.deepEqual(draft.formats(branch), [first]); assert.equal(draft.canRedo(branch), false)
  assert.equal(draft.canUndo(branch), false); assert.equal(draft.canUndo(initial), true)
})

test('editing operation formats a numeric-looking text cell without parsing its value or mutating the workbook', () => {
  const book = sample('123'), before = structuredClone(book), input = editInput(book)
  const operation = planSheetCellInlineFontFormat(book, { ...input, commandId: 'sheet.command.toggle-bold', selection: { start: 1, end: 2 } }, defaults)
  assert.equal(operation.accepted, true); assert.equal(operation.edit.at, 0); assert.deepEqual(book, before)
  const commit = planSheetCellInput(book, { ...input, inlineFormats: operation.formats }, defaults)
  assert.equal(commit.result.accepted, true); assert.equal(commit.result.changed, true)
  assert.equal(commit.cell.value, '123'); assert.equal(commit.cell.dataType, 'text'); assert.equal(commit.cell.style.numberFormat, 'General')
  assert.equal(commit.cell.richText.map(run => run.text).join(''), '123'); assert.deepEqual(book, before)
})

test('partial font operations reject real numbers and formulas and retain a stale source draft', () => {
  for (const value of [123, true]) {
    const book = sample(value), input = editInput(book)
    assert.equal(planSheetCellInlineFontFormat(book, { ...input, commandId: 'sheet.command.toggle-bold', selection: { start: 0, end: 1 } }, defaults).accepted, false)
    assert.throws(() => planSheetCellInlineStyle(book, { unitId: book.id, sheetId: book.sheets[0].id, row: 0, column: 0, start: 0, end: 1, style: { bold: true } }, defaults))
  }
  const book = sample('abcd'), input = editInput(book), before = structuredClone(book)
  const stale = { ...input, input: 'wrong', editTrace: { ...input.editTrace, source: 'wrong' } }
  assert.equal(planSheetCellInlineFontFormat(book, { ...stale, commandId: 'sheet.command.toggle-bold', selection: { start: 0, end: 1 } }, defaults).accepted, false)
  const formula = { ...input, input: '=1', editTrace: { ...input.editTrace, operations: [{ start: 0, end: 4, text: '=1', deletedText: 'abcd' }] } }
  assert.equal(planSheetCellInlineFontFormat(book, { ...formula, commandId: 'sheet.command.toggle-bold', selection: { start: 0, end: 2 } }, defaults).accepted, false)
  assert.deepEqual(book, before)
})

test('direct styles preserve raw text, forced-string types, display, number format and metadata', () => {
  for (const value of ['123', "'abc"]) {
    const book = sample(value), cell = book.sheets[0].cells['0:0']
    cell.dataType = 'force_string'; cell.style.numberFormat = 'General'; cell.display = 'preserved display'
    const before = structuredClone(cell)
    const plan = planSheetCellInlineStyle(book, { unitId: book.id, sheetId: book.sheets[0].id, row: 0, column: 0, start: 0, end: 1, style: { italic: true, color: '#aabbcc' } }, defaults)
    assert.equal(plan.result.changed, true)
    assert.deepEqual({ ...plan.cell, richText: before.richText }, before)
    assert.equal(plan.cell.richText.map(run => run.text).join(''), value)
    assert.deepEqual(cell, before)
  }
})

test('format-only input accepts existing invalid data while actual content changes still receive stop validation', () => {
  const book = sample('abcd'), sheet = book.sheets[0]
  sheet.dataValidations = [{ id: 'length', type: 'textLength', operator: 'lessThanOrEqual', formula1: '2', allowBlank: false,
    errorStyle: 'stop', errorTitle: '长度', error: '最多两个字符', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }] }]
  const input = editInput(book), formats = [format(0, 'abcd', 0, 1, 'bold', true)]
  assert.equal(planSheetCellInput(book, { ...input, inlineFormats: formats }, defaults).result.accepted, true)
  const changed = editInput(book, [{ start: 4, end: 4, text: 'X', deletedText: '' }], 'abcdX')
  const before = structuredClone(book), rejected = planSheetCellInput(book, { ...changed, inlineFormats: formats }, defaults)
  assert.equal(rejected.result.accepted, false); assert.equal(rejected.result.reason, 'validation'); assert.deepEqual(book, before)
})
