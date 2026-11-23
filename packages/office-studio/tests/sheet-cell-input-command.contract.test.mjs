import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { applyExcelCellView, applyExcelWorkbookView, toExcelCellView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createSheetCellEditSession } from '../src/features/sheet-workbench/sheet-cell-edit-session.ts'
import { sheetCellContentKey, sheetCellInputText } from '../src/features/sheet-workbench/sheet-cell-edit.ts'
import { planSheetCellInput, sheetCellInputCommandInput } from '../src/features/sheet-workbench/sheet-cell-input-command.ts'
import { editSheetCellRichText } from '../src/features/sheet-workbench/sheet-cell-rich-text-edit.ts'

const copy = value => JSON.parse(JSON.stringify(value))
const defaults = { region: 'en-US', currentYear: 2026, userId: 'local-user' }
function model(value = null, pattern = 'General') {
  const book = createBlankWorkbook(); book.id = 'input-book'; book.calculationMode = 'manual'
  book.sheets[0].id = 'input-sheet'; book.sheets[0].rowCount = 20; book.sheets[0].columnCount = 8; book.activeSheetId = 'input-sheet'
  if (value !== null) {
    const cell = createCell(0, 0, value); cell.style.numberFormat = pattern
    book.sheets[0].cells['0:0'] = cell
  }
  return book
}
const input = (text, extra = {}) => ({ unitId: 'input-book', sheetId: 'input-sheet', row: 0, column: 0, input: text, ...extra })
const session = (book, extra = {}) => createSheetCellEditSession({ documentId: book.id, sheetId: book.sheets[0].id,
  key: 'A1', date1904: book.date1904, cell: book.sheets[0].cells['0:0'] ? toExcelCellView(book.sheets[0].cells['0:0']) : undefined }, { ...defaults, ...extra })

test('cell input rejects malformed coordinates, conflicting aliases and invalid interpretation settings before planning', () => {
  const parsed = sheetCellInputCommandInput({ ...input('12'), sheetId: undefined, subUnitId: 'input-sheet', column: undefined, col: 2 })
  assert.equal(parsed.sheetId, 'input-sheet'); assert.equal(parsed.column, 2)
  for (const changes of [{ row: -1 }, { row: 1.5 }, { row: Infinity }, { row: 1_048_576 }, { column: 16_384 },
    { input: 12 }, { unitId: '' }, { sheetId: [] }, { subUnitId: 'other' }, { col: 1 }, { date1904: 1 },
    { region: 'invalid_@' }, { currentYear: 0 }, { currentYear: 2026.5 }, { session: { source: {} } }]) {
    assert.throws(() => sheetCellInputCommandInput(input('12', changes)))
  }
  const start = session(model(12))
  assert.throws(() => sheetCellInputCommandInput(input('13', { session: start, region: 'de-DE' })), /会话不一致/)
  assert.throws(() => sheetCellInputCommandInput(input('13', { session: { ...start, source: { ...start.source, value: NaN } } })))
  for (const key of ['A0', 'XFE1', 'A1048577', 'A1:B2', ' A1 ', 'AAAAAAAAAAAAAAAAAAAAAAAA1']) {
    assert.throws(() => sheetCellInputCommandInput(input('13', { session: { ...start, key } })))
  }
  assert.equal(sheetCellInputCommandInput(input('13', { session: { ...start, key: 'a1' } })).session.key, 'A1')
})

test('a shared input plan parses percentages, forced text, explicit text formats and formulas without mutating the model', () => {
  for (const [book, text, expected] of [
    [model(), '12.5%', { value: .125, formula: null, type: 'number' }],
    [model(), "'00123", { value: '00123', formula: null, type: 'force_string' }],
    [model('old', '@'), '=1+2', { value: '=1+2', formula: null, type: 'text' }],
    [model(99), '=1+2', { value: null, formula: '=1+2', type: 'formula' }],
  ]) {
    const before = copy(book), plan = planSheetCellInput(book, input(text), defaults)
    assert.equal(plan.result.accepted, true); assert.equal(plan.result.changed, true)
    assert.equal(plan.cell.value, expected.value); assert.equal(plan.cell.formula, expected.formula); assert.equal(plan.cell.dataType, expected.type)
    assert.equal(plan.cell.cachedDataType, null); assert.deepEqual(book, before)
    if (text.endsWith('%')) assert.match(plan.cell.style.numberFormat, /%/)
  }
})

test('the session preserves its regional interpretation while current visual styles and extensions remain authoritative', () => {
  const book = model(.125, '0.00%'), start = session(book, { region: 'de-DE' })
  const cell = book.sheets[0].cells['0:0']; cell.style.font.color = '#ff0000'; cell.style.border = { top: { style: 'thin' } }
  cell.metadata = { vendor: { id: 'preserved' } }
  const before = copy(book), plan = planSheetCellInput(book, input('25,5%', { session: start }), { ...defaults, region: 'en-US' })
  assert.equal(plan.result.accepted, true); assert.equal(plan.cell.value, .255); assert.equal(plan.cell.style.font.color, '#ff0000')
  plan.cell.style.border.top.style = 'double'; plan.cell.metadata.vendor.id = 'caller-mutation'
  assert.deepEqual(book, before)
})

test('untouched editor text preserves unsupported formula caches and empty sparse cells without a history-worthy edit', () => {
  const book = model(17); const cell = book.sheets[0].cells['0:0']
  cell.formula = '=UNSUPPORTED()'; cell.dataType = 'formula'; cell.cachedDataType = 'number'; cell.display = '17.00'
  const start = session(book); cell.value = 18; cell.display = '18.00'
  const plan = planSheetCellInput(book, input('=UNSUPPORTED()', { session: start }), defaults)
  assert.equal(plan.result.accepted, true); assert.equal(plan.result.changed, false); assert.equal(plan.cell, undefined)
  const blank = model(), before = copy(blank)
  assert.equal(planSheetCellInput(blank, input(''), defaults).result.changed, false)
  assert.deepEqual(blank, before)
})

test('document, sheet, cell, date-system, number-format and source-content conflicts preserve the input and source model', () => {
  const original = model(12, '0.00'), start = session(original)
  for (const [change, expected] of [
    [book => { book.id = 'other' }, 'document'],
    [book => { book.date1904 = true }, 'date-system'],
    [book => { book.sheets[0].cells['0:0'].style.numberFormat = '0%' }, 'number-format'],
    [book => { book.sheets[0].cells['0:0'].value = 13 }, 'content'],
  ]) {
    const book = copy(original); change(book); const before = copy(book)
    const plan = planSheetCellInput(book, input('20', { unitId: book.id, session: start }), defaults)
    assert.equal(plan.result.accepted, false); assert.equal(plan.result.conflict, expected); assert.deepEqual(book, before)
  }
  assert.equal(planSheetCellInput(original, input('20', { session: { ...start, sheetId: 'other' } }), defaults).result.conflict, 'sheet')
  assert.equal(planSheetCellInput(original, input('20', { column: 1, session: start }), defaults).result.conflict, 'cell')
})

test('merged interior input targets its content owner and protection checks the complete merge', () => {
  const book = model('owner'), sheet = book.sheets[0]
  sheet.mergedRanges = [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]; sheet.hiddenRows = [0]
  const plan = planSheetCellInput(book, input('new', { row: 1, column: 1 }), defaults)
  assert.deepEqual(plan.position, { row: 0, column: 0 }); assert.equal(plan.cell.id, '0:0'); assert.equal(plan.cell.value, 'new')
  assert.equal(sheetCellContentKey({ merges: ['A1:B2'] }, 'B2'), 'A1')
  sheet.rangeProtectionRules = [{ id: 'locked', permissionId: 'locked-permission', ownerUserId: 'remote', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'othersCanView', ranges: [{ startRow: 1, endRow: 1, startCol: 1, endCol: 1 }], description: '' }]
  const before = copy(book)
  assert.throws(() => planSheetCellInput(book, input('new', { row: 1, column: 1 }), defaults), /受保护/)
  assert.deepEqual(book, before)
})

test('stop validation refuses a candidate atomically while warnings retain typed values and real issue details', () => {
  const book = model(5), sheet = book.sheets[0]
  sheet.dataValidations = [{ id: 'integers', type: 'whole', operator: 'between', formula1: '1', formula2: '10', allowBlank: false,
    errorStyle: 'stop', errorTitle: '范围', error: '1 到 10', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }] }]
  const before = copy(book), rejected = planSheetCellInput(book, input('=20'), defaults)
  assert.equal(rejected.result.accepted, false); assert.equal(rejected.result.reason, 'validation'); assert.equal(rejected.result.issues[0].ruleId, 'integers')
  assert.deepEqual(book, before)
  sheet.dataValidations[0].errorStyle = 'warning'
  const warning = planSheetCellInput(book, input('20'), defaults)
  assert.equal(warning.result.accepted, true); assert.equal(warning.cell.value, 20); assert.equal(warning.result.issues[0].severity, 'warning')
  warning.result.issues[0].message = 'mutated'
  assert.equal(sheet.dataValidations[0].error, '1 到 10')
})

test('F2 text edits retain unaffected run formatting, direct overwrite clears rich runs and mismatched source runs are refused', () => {
  const book = model('abcdef'), cell = book.sheets[0].cells['0:0']
  cell.richText = [{ text: 'ab', font: { bold: true } }, { text: 'cd', font: { italic: true } }, { text: 'ef', vendor: { id: 7 } }]
  const start = session(book), before = copy(book), plan = planSheetCellInput(book, input('abXdef', { session: start }), defaults)
  assert.equal(plan.cell.richText.map(run => run.text).join(''), 'abXdef')
  assert.equal(plan.cell.richText.find(run => run.text === 'Xd').font.italic, true)
  assert.deepEqual(plan.cell.richText.find(run => run.text === 'ef').vendor, { id: 7 }); assert.deepEqual(book, before)
  assert.equal(planSheetCellInput(book, input('abcdef', { session: session(book, { mode: 'overwrite' }) }), defaults).cell.richText, undefined)
  cell.richText[0].text = 'invalid'
  assert.equal(planSheetCellInput(book, input('changed', { session: start }), defaults).result.reason, 'rich-text-source')
})

test('run replacement respects code-point boundaries and retains empty source entities and caller ownership', () => {
  const runs = [{ text: 'A', font: { bold: true } }, { text: '🙂', font: { italic: true } }, { text: 'B' }, { text: '', vendor: { empty: true } }]
  const before = copy(runs), result = editSheetCellRichText(runs, 'A🙂B', 'A👩🏽‍💻B')
  assert.equal(result.map(run => run.text).join(''), 'A👩🏽‍💻B')
  assert.deepEqual(result.find(run => run.text === '').vendor, { empty: true })
  result[0].font.bold = false; assert.deepEqual(runs, before)
  assert.equal(editSheetCellRichText([{ text: 1 }], '1', '2'), undefined)
})

test('text edits keep interior empty markers in source order and do not duplicate explicit run identifiers', () => {
  const runs = [{ id: 'leading', text: '', vendor: { kind: 'leading' } },
    { id: 'one', text: 'ab', font: { bold: true }, vendor: { id: 7 } },
    { id: 'middle', text: '', vendor: { kind: 'boundary' } },
    { id: 'two', text: 'cd', font: { italic: true } },
    { id: 'trailing', text: '', vendor: { kind: 'trailing' } }]
  const before = copy(runs), result = editSheetCellRichText(runs, 'abcd', 'aXcd')
  assert.deepEqual(result.map(run => [run.id, run.text]), [['leading', ''], ['one', 'aX'], ['middle', ''], ['two', 'cd'], ['trailing', '']])
  assert.equal(new Set(result.map(run => run.id)).size, result.length)
  assert.deepEqual(result[1].vendor, { id: 7 }); assert.deepEqual(runs, before)
  const insert = editSheetCellRichText(runs, 'abcd', 'abYcd')
  assert.deepEqual(insert.map(run => [run.id, run.text]), [['leading', ''], ['one', 'abY'], ['middle', ''], ['two', 'cd'], ['trailing', '']])
  const inside = editSheetCellRichText([{ id: 'one', text: 'abcdef', font: { italic: true } }], 'abcdef', 'abXYef')
  assert.deepEqual(inside, [{ id: 'one', text: 'abXYef', font: { italic: true } }])
})

test('single-cell and complete-view mapping share raw values, formula caches, styles and preserved resource semantics', () => {
  const book = model(12, '0.00'), previous = book.sheets[0].cells['0:0']
  previous.hyperlink = 'https://example.test'; previous.metadata = { vendor: { payload: 7 } }
  const view = toExcelWorkbookView(book), edited = { ...view.sheets[0].cells.A1, value: '15', rawValue: 15, displayValue: '15.00' }
  view.sheets[0].cells.A1 = edited
  const single = applyExcelCellView({ row: 0, col: 0 }, edited, previous, { date1904: false, previousDate1904: false, numberFormatLocale: 'en-US' })
  assert.deepEqual(single, applyExcelWorkbookView(book, view, 'en-US').sheets[0].cells['0:0'])
  assert.equal(sheetCellInputText(toExcelCellView(single)), '15')
})

function repeatedModel(pattern = 'General', forced = false) {
  const book = model('ababab', pattern), cell = book.sheets[0].cells['0:0']
  if (forced) cell.dataType = 'force_string'
  cell.richText = [{ id: 'one', text: 'ab', font: { bold: true }, vendor: { keep: 1 } },
    { id: 'marker', text: '', vendor: { marker: true } }, { id: 'two', text: 'ab', font: { italic: true } },
    { id: 'three', text: 'ab', hyperlink: 'https://example.test', vendor: { keep: 3 } }]
  return book
}
const rangeTrace = (start, source, operations) => ({ version: 1, sessionId: start.sessionId, source, operations })
const splice = (start, end, text, deletedText) => ({ start, end, text, deletedText, origin: 'beforeinput' })

test('the shared planner uses repeated-text source ranges and honors same-final-text run changes', () => {
  const book = repeatedModel(), start = session(book), before = copy(book)
  const editTrace = rangeTrace(start, 'ababab', [splice(2, 4, '', 'ab'), splice(2, 2, 'ab', '')])
  const plan = planSheetCellInput(book, input('ababab', { session: start, editTrace }), defaults)
  assert.equal(plan.result.accepted, true); assert.equal(plan.result.changed, true); assert.equal(plan.cell.value, 'ababab')
  assert.deepEqual(plan.cell.richText.map(run => [run.id, run.text]), [['one', 'abab'], ['marker', ''], ['three', 'ab']])
  assert.equal(plan.cell.richText[0].font.bold, true); assert.equal(plan.cell.richText[2].hyperlink, 'https://example.test')
  assert.deepEqual(book, before)
})

test('trace validation rejects wrong sessions, reordered edits and stale final text before any candidate mutation', () => {
  const book = repeatedModel(), start = session(book), before = copy(book)
  const editTrace = rangeTrace(start, 'ababab', [splice(2, 2, 'X', ''), splice(3, 5, 'Y', 'ab')])
  for (const changes of [
    { editTrace: { ...editTrace, sessionId: 'old-session' } }, { session: undefined, editTrace },
    { editTrace: { ...editTrace, operations: [...editTrace.operations].reverse() } },
    { input: 'wrong', editTrace }, { editTrace: { ...editTrace, operations: [splice(1, 2, 'X', 'stale')] } },
    { session: { ...start, sessionId: '' }, editTrace },
  ]) assert.throws(() => planSheetCellInput(book, input('abXYab', { session: start, ...changes }), defaults))
  assert.deepEqual(book, before)
})

test('a self-consistent trace against another source is refused while trace and rich-source objects stay caller-owned', () => {
  const book = repeatedModel(), start = session(book), before = copy(book)
  const wrong = rangeTrace(start, 'other', [splice(0, 5, 'new', 'other')])
  const rejected = planSheetCellInput(book, input('new', { session: start, editTrace: wrong }), defaults)
  assert.equal(rejected.result.accepted, false); assert.equal(rejected.result.reason, 'edit-trace')
  const editTrace = rangeTrace(start, 'ababab', [splice(2, 2, 'X', '')])
  const parsed = sheetCellInputCommandInput(input('abXabab', { session: start, editTrace }))
  parsed.session.source.richText[0].font.bold = false; parsed.editTrace.operations[0].text = 'caller'
  assert.equal(start.source.richText[0].font.bold, true); assert.equal(editTrace.operations[0].text, 'X')
  assert.deepEqual(book, before)
})

test('authoritative same-text rich mutations conflict and supplied session runs cannot override them', () => {
  const book = repeatedModel(), start = session(book)
  book.sheets[0].cells['0:0'].richText[1].vendor.marker = false
  const before = copy(book), editTrace = rangeTrace(start, 'ababab', [splice(2, 2, 'X', '')])
  const plan = planSheetCellInput(book, input('abXabab', { session: start, editTrace }), defaults)
  assert.equal(plan.result.reason, 'conflict'); assert.equal(plan.result.conflict, 'rich-text')
  assert.deepEqual(book, before)
  const next = session(book); next.source.richText[0].font.bold = false
  assert.equal(planSheetCellInput(book, input('abXabab', { session: next, editTrace: rangeTrace(next, 'ababab', editTrace.operations) }), defaults).result.conflict, 'rich-text')
  assert.deepEqual(book, before)
})

test('escaped forced text and explicit text formats project editor offsets to their raw rich text', () => {
  const forced = repeatedModel('General', true), forcedStart = session(forced)
  const forcedPlan = planSheetCellInput(forced, input("'abXabab", { session: forcedStart,
    editTrace: rangeTrace(forcedStart, "'ababab", [splice(3, 3, 'X', '')]) }), defaults)
  assert.equal(forcedPlan.cell.value, 'abXabab'); assert.equal(forcedPlan.cell.dataType, 'force_string')
  assert.deepEqual(forcedPlan.cell.richText.map(run => [run.id, run.text]), [['one', 'abX'], ['marker', ''], ['two', 'ab'], ['three', 'ab']])
  const plain = repeatedModel('@'), plainStart = session(plain)
  const plainPlan = planSheetCellInput(plain, input("'ababab", { session: plainStart,
    editTrace: rangeTrace(plainStart, 'ababab', [splice(0, 0, "'", '')]) }), defaults)
  assert.equal(plainPlan.cell.value, "'ababab"); assert.equal(plainPlan.cell.richText[0].text, "'ab")
})

test('explicit source paths still clear rich text for formula/numeric conversion and direct overwrite', () => {
  for (const [text, expected] of [['=1+2', 'formula'], ['12', 'number']]) {
    const book = repeatedModel(), start = session(book)
    const plan = planSheetCellInput(book, input(text, { session: start,
      editTrace: rangeTrace(start, 'ababab', [splice(0, 6, text, 'ababab')]) }), defaults)
    assert.equal(plan.result.accepted, true); assert.equal(plan.cell.dataType, expected); assert.equal(plan.cell.richText, undefined)
  }
  const book = repeatedModel(), start = session(book, { mode: 'overwrite' })
  const plan = planSheetCellInput(book, input('ababab', { session: start, editTrace: rangeTrace(start, 'ababab', [splice(0, 6, 'ababab', 'ababab')]) }), defaults)
  assert.equal(plan.result.changed, true); assert.equal(plan.cell.richText, undefined)
})

test('a trace returning to its source without active operations preserves no-op caches and sparse cells', () => {
  const book = repeatedModel(), start = session(book), before = copy(book)
  const plan = planSheetCellInput(book, input('ababab', { session: start, editTrace: rangeTrace(start, 'ababab', []) }), defaults)
  assert.equal(plan.result.changed, false); assert.equal(plan.cell, undefined); assert.deepEqual(book, before)
  const blank = model(), empty = session(blank)
  assert.equal(planSheetCellInput(blank, input('', { session: empty, editTrace: rangeTrace(empty, '', []) }), defaults).result.changed, false)
  assert.deepEqual(blank.sheets[0].cells, {})
})

test('legacy API sessions remain usable without claiming an exact repeated-text input trajectory', () => {
  const book = repeatedModel(), current = session(book), { sessionId: _id, ...legacy } = current
  const { richText: _rich, ...source } = legacy.source; legacy.source = source
  const plan = planSheetCellInput(book, input('abXabab', { session: legacy }), defaults)
  assert.equal(plan.result.accepted, true); assert.equal(plan.cell.richText.map(run => run.text).join(''), 'abXabab')
  assert.equal(sheetCellInputCommandInput(input('abXabab', { session: legacy })).session.sessionId, undefined)
})
