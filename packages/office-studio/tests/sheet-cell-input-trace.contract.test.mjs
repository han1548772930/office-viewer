import assert from 'node:assert/strict'
import { test } from 'vitest'
import { SheetCellInputEditRecorder, applySheetCellTextEdit, projectSheetCellInputEditTrace,
  readSheetCellInputEditTrace, replaySheetCellInputEditTrace } from '../src/features/sheet-workbench/sheet-cell-input-trace.ts'
import { editSheetCellRichText } from '../src/features/sheet-workbench/sheet-cell-rich-text-edit.ts'

const copy = value => JSON.parse(JSON.stringify(value))
const runs = () => [
  { id: 'one', text: 'ab', font: { bold: true }, vendor: { identity: 1 } },
  { id: 'marker', text: '', vendor: { keep: true } },
  { id: 'two', text: 'ab', font: { italic: true } },
  { id: 'three', text: 'ab', hyperlink: 'https://example.test', vendor: { identity: 3 } },
]
const operation = (start, end, text, deletedText) => ({ start, end, text, deletedText })
const trace = (source, operations) => ({ version: 1, sessionId: 'session-a', source, operations })
const native = (recorder, phase, value, start, end, extra = {}) => recorder.handle({ sessionId: recorder.sessionId,
  hostId: 'cell-host', host: 'cell', phase, value, start, end, ...extra })
function recorder(source = 'ababab') {
  const value = new SheetCellInputEditRecorder('session-a', source)
  assert.equal(native(value, 'bind', source, source.length, source.length), true)
  return value
}
function input(value, start, end, after, inputType = 'insertText', data = null) {
  assert.equal(native(value, 'beforeinput', value.value, start, end, { inputType, data }), true)
  const caret = start + after.length - value.value.length + end - start
  assert.equal(native(value, 'input', after, caret, caret, { inputType, data }), true)
}

test('repeated insertion and deletion use the actual source interval and preserve unaffected run identities', () => {
  const source = runs(), before = copy(source)
  const inserted = editSheetCellRichText(source, 'ababab', 'abababab', [operation(2, 2, 'ab', '')])
  assert.deepEqual(inserted.map(run => [run.id, run.text]), [['one', 'abab'], ['marker', ''], ['two', 'ab'], ['three', 'ab']])
  assert.equal(inserted[0].font.bold, true); assert.equal(inserted[3].hyperlink, 'https://example.test')
  const deleted = editSheetCellRichText(source, 'ababab', 'abab', [operation(2, 4, '', 'ab')])
  assert.deepEqual(deleted.map(run => [run.id, run.text]), [['one', 'ab'], ['marker', ''], ['three', 'ab']])
  inserted[0].vendor.identity = 99; assert.deepEqual(source, before)
})

test('several native edits ending in the original text retain their actual changed run affinity', () => {
  const value = recorder()
  input(value, 2, 4, 'abab', 'deleteByCut')
  input(value, 2, 2, 'ababab', 'insertText', 'ab')
  assert.equal(value.trace.operations.length, 2)
  const result = editSheetCellRichText(runs(), 'ababab', value.value, value.trace.operations)
  assert.deepEqual(result.map(run => [run.id, run.text]), [['one', 'abab'], ['marker', ''], ['three', 'ab']])
  assert.equal(result[0].font.bold, true)
})

test('collapsed native backward and forward deletions use caret direction and actual deleted length', () => {
  for (const [inputType, caret] of [['deleteContentBackward', 4], ['deleteContentForward', 2], ['deleteWordBackward', 4]]) {
    const value = recorder()
    assert.equal(native(value, 'beforeinput', value.value, caret, caret, { inputType }), true)
    assert.equal(native(value, 'input', 'abab', 2, 2, { inputType }), true)
    assert.deepEqual(value.trace.operations.map(({ start, end, deletedText }) => [start, end, deletedText]), [[2, 4, 'ab']])
  }
})

test('selected same-text replacement is recorded while duplicate value-only input is ignored', () => {
  const value = recorder()
  input(value, 2, 4, 'ababab', 'insertFromPaste', 'ab')
  assert.deepEqual(value.trace.operations.map(({ start, end, text }) => [start, end, text]), [[2, 4, 'ab']])
  assert.equal(native(value, 'input', 'ababab', 4, 4, { inputType: 'insertText' }), true)
  assert.equal(value.trace.operations.length, 1)
})

test('one IME group keeps its original range across candidate updates and a duplicate trailing input', () => {
  const value = recorder()
  assert.equal(native(value, 'compositionstart', 'ababab', 2, 4), true)
  assert.equal(native(value, 'input', 'ab中ab', 3, 3, { inputType: 'insertCompositionText', composing: true }), true)
  assert.equal(native(value, 'input', 'ab中文ab', 4, 4, { inputType: 'insertCompositionText', composing: true }), true)
  assert.equal(native(value, 'compositionend', 'ab中文ab', 4, 4, { data: '中文' }), true)
  assert.equal(value.composing, false)
  assert.equal(native(value, 'input', 'ab中文ab', 4, 4, { inputType: 'insertFromComposition', data: '中文' }), true)
  assert.deepEqual(value.trace.operations.map(({ start, end, deletedText, text, origin }) => [start, end, deletedText, text, origin]),
    [[2, 4, 'ab', '中文', 'composition']])
  const result = editSheetCellRichText(runs(), 'ababab', value.value, value.trace.operations)
  assert.equal(result.find(run => run.id === 'two').font.italic, true)
  assert.equal(result.find(run => run.id === 'three').hyperlink, 'https://example.test')
})

test('compositionend preceding the final native input revises its group instead of inserting a second copy', () => {
  const value = recorder()
  native(value, 'compositionstart', 'ababab', 2, 4)
  native(value, 'input', 'ab中ab', 3, 3, { inputType: 'insertCompositionText' })
  native(value, 'compositionend', 'ab中ab', 3, 3, { data: '中文' })
  native(value, 'input', 'ab中文ab', 4, 4, { inputType: 'insertFromComposition', data: '中文' })
  assert.equal(value.value, 'ab中文ab'); assert.equal(value.trace.operations.length, 1)
  assert.equal(value.trace.operations[0].text, '中文')
})

test('a canceled composition restores the preceding range history and a later edit starts a new group', () => {
  const value = recorder()
  input(value, 2, 2, 'abXabab', 'insertText', 'X')
  const before = copy(value.trace)
  native(value, 'compositionstart', 'abXabab', 3, 5)
  native(value, 'input', 'abX中ab', 4, 4, { inputType: 'insertCompositionText' })
  native(value, 'compositionend', 'abXabab', 3, 5, { data: '' })
  assert.deepEqual(value.trace, before); assert.equal(value.value, 'abXabab')
  input(value, 3, 5, 'abXYab', 'insertText', 'Y')
  assert.equal(value.trace.operations.length, 2)
})

test('native undo and redo recover the recorded rich path, including browser-coalesced edits and redo branching', () => {
  const value = recorder()
  input(value, 2, 4, 'abab', 'deleteByCut')
  input(value, 2, 2, 'abXab', 'insertText', 'X')
  const committed = copy(value.trace)
  native(value, 'input', 'ababab', 2, 4, { inputType: 'historyUndo' })
  assert.equal(value.trace.operations.length, 0)
  assert.deepEqual(editSheetCellRichText(runs(), 'ababab', value.value, value.trace.operations), runs())
  native(value, 'input', 'abXab', 3, 3, { inputType: 'historyRedo' })
  assert.deepEqual(value.trace, committed)
  native(value, 'input', 'abab', 2, 2, { inputType: 'historyUndo' })
  input(value, 2, 2, 'abYab', 'insertText', 'Y')
  assert.equal(native(value, 'input', 'abXab', 3, 3, { inputType: 'historyRedo' }), false)
  assert.equal(value.value, 'abYab'); assert.equal(value.trace.operations[1].text, 'Y')
})

test('cell and formula hosts share one draft while stale sessions, unbound hosts and foreign composition events are refused', () => {
  const value = recorder()
  input(value, 2, 2, 'abXabab', 'insertText', 'X')
  const formula = (phase, text, start, end) => value.handle({ sessionId: value.sessionId, hostId: 'formula-host', host: 'formula', phase, value: text, start, end })
  assert.equal(formula('bind', 'ababab', 2, 2), false)
  assert.equal(formula('bind', 'abXabab', 3, 3), true)
  assert.equal(formula('compositionstart', 'abXabab', 3, 5), true)
  assert.equal(native(value, 'input', 'abXwrongab', 8, 8, { inputType: 'insertCompositionText' }), false)
  const before = copy(value.trace)
  assert.equal(value.handle({ sessionId: 'old-session', hostId: 'cell-host', host: 'cell', phase: 'replace', value: value.value,
    start: 0, end: 2, text: 'old' }), false)
  assert.equal(value.handle({ sessionId: value.sessionId, hostId: 'never-bound', host: 'cell', phase: 'replace', value: value.value,
    start: 0, end: 2, text: 'old' }), false)
  assert.deepEqual(value.trace, before); assert.equal(value.compositionHost, 'formula')
})

test('native spellcheck replacement can use its resulting caret and known data without a global prefix search', () => {
  const value = recorder('word word')
  native(value, 'beforeinput', 'word word', 7, 7, { inputType: 'insertReplacementText', data: 'term' })
  native(value, 'input', 'word term', 9, 9, { inputType: 'insertReplacementText', data: 'term' })
  assert.deepEqual(value.trace.operations.map(({ start, end, text, origin }) => [start, end, text, origin]), [[5, 9, 'term', 'input-caret']])
})

test('unlocated compatibility input is explicitly labeled rather than claiming a beforeinput range', () => {
  const value = recorder('ababab')
  native(value, 'selection', 'ababab', 6, 6)
  native(value, 'input', 'abZab', 3, 3, { inputType: 'insertReplacementText' })
  assert.equal(value.trace.operations[0].origin, 'compatibility')
  assert.equal(replaySheetCellInputEditTrace(value.trace), 'abZab')
})

test('programmatic replacements preserve source ranges and operation snapshots cannot mutate the active draft', () => {
  const value = recorder('ababab')
  native(value, 'replace', value.value, 2, 4, { text: '\nnext', inputType: 'insertFromPaste' })
  const snapshot = value.trace
  assert.equal(snapshot.operations[0].origin, 'programmatic')
  snapshot.operations[0].text = 'caller change'; snapshot.operations.push(operation(0, 0, 'x', ''))
  assert.equal(value.value, 'ab\nnextab'); assert.equal(value.trace.operations.length, 1)
})

test('raw projection handles a retained escape, its deletion, force-text insertion and literal apostrophes', () => {
  const cases = [
    ["'abab", 'abab', 'abXab', [operation(3, 3, 'X', '')], false],
    ["'abab", 'abab', 'abab', [operation(0, 1, '', "'")], false],
    ['abab', 'abab', 'abab', [operation(0, 0, "'", '')], false],
    ["'abab", 'abab', "'abab", [operation(0, 0, "'", '')], false],
    ["'abab", "'abab", "'abXab", [operation(3, 3, 'X', '')], true],
  ]
  for (const [editor, before, after, operations, textFormat] of cases) {
    const projected = projectSheetCellInputEditTrace(trace(editor, operations), before, after, textFormat)
    assert.ok(projected)
    assert.equal(projected.reduce((text, edit) => applySheetCellTextEdit(text, edit), before), after)
  }
})

test('UTF-16 source checks protect supplementary characters, including projected editor escape offsets', () => {
  const source = [{ id: 'a', text: 'A' }, { id: 'emoji', text: '🙂', font: { italic: true } }, { id: 'b', text: 'B' }]
  const edit = operation(1, 3, '👩🏽‍💻', '🙂')
  const result = editSheetCellRichText(source, 'A🙂B', 'A👩🏽‍💻B', [edit])
  assert.equal(result.find(run => run.id === 'emoji').font.italic, true)
  assert.equal(editSheetCellRichText(source, 'A🙂B', 'A?B', [operation(2, 3, '?', '\ude42')]), undefined)
  assert.throws(() => readSheetCellInputEditTrace(trace('A🙂B', [operation(1, 2, '?', '\ud83d')])))
  const projected = projectSheetCellInputEditTrace(trace("'A🙂B", [operation(2, 4, '中', '🙂')]), 'A🙂B', 'A中B', false)
  assert.deepEqual(projected.map(({ start, end, text }) => [start, end, text]), [[1, 3, '中']])
})

test('empty source markers remain separate when all text is deleted and retyped using the former caret style', () => {
  const source = [{ id: 'leading', text: '', vendor: { leading: true } },
    { id: 'text', text: 'ab', font: { italic: true }, hyperlink: 'https://example.test' },
    { id: 'trailing', text: '', vendor: { trailing: true } }]
  const result = editSheetCellRichText(source, 'ab', '中', [operation(0, 2, '', 'ab'), operation(0, 0, '中', '')])
  assert.equal(result.find(run => run.id === 'leading').text, ''); assert.equal(result.find(run => run.id === 'trailing').text, '')
  assert.equal(result.find(run => run.id === 'text').font.italic, true)
  assert.equal(result.find(run => run.id === 'text').hyperlink, 'https://example.test')
  assert.equal(new Set(result.map(run => run.id)).size, result.length)
})

test('the trace runtime boundary rejects invalid versions, non-text payloads, stale deleted text and operation reordering', () => {
  const valid = trace('abab', [operation(2, 2, 'X', ''), operation(3, 5, 'Y', 'ab')])
  assert.equal(replaySheetCellInputEditTrace(readSheetCellInputEditTrace(valid)), 'abXY')
  for (const candidate of [null, [], { ...valid, version: 2 }, { ...valid, sessionId: '' }, { ...valid, source: 1 },
    { ...valid, operations: null }, trace('ab', [operation(-1, 0, 'X', '')]), trace('ab', [operation(0, 3, 'X', 'ab')]),
    trace('ab', [operation(2, 1, 'X', '')]), trace('ab', [operation(.5, 1, 'X', 'a')]),
    trace('ab', [operation(0, 1, 4, 'a')]), trace('ab', [operation(0, 1, 'X', 'stale')]),
    trace('ab', [{ ...operation(0, 1, 'X', 'a'), origin: { toString: () => 'beforeinput' } }]),
    trace(valid.source, [...valid.operations].reverse())]) assert.throws(() => readSheetCellInputEditTrace(candidate))
  const captured = readSheetCellInputEditTrace(valid); valid.operations[0].text = 'caller'
  assert.equal(captured.operations[0].text, 'X')
})

test('a missing compositionstart can still use the first composing beforeinput source selection', () => {
  const value = recorder()
  native(value, 'beforeinput', value.value, 2, 4, { inputType: 'insertCompositionText', composing: true })
  native(value, 'input', 'ab中ab', 3, 3, { inputType: 'insertCompositionText', composing: true })
  native(value, 'compositionend', 'ab中ab', 3, 3, { data: '中' })
  assert.equal(value.composing, false)
  assert.deepEqual(value.trace.operations.map(({ start, end, text, origin }) => [start, end, text, origin]), [[2, 4, '中', 'composition']])
})

test('the first IME beforeinput may refine a reconversion interval captured as a collapsed start caret', () => {
  const value = recorder()
  native(value, 'compositionstart', value.value, 4, 4)
  native(value, 'beforeinput', value.value, 2, 4, { inputType: 'insertCompositionText', composing: true })
  native(value, 'input', 'ab中ab', 3, 3, { inputType: 'insertCompositionText', composing: true })
  native(value, 'compositionend', 'ab中ab', 3, 3, { data: '中' })
  assert.deepEqual(value.trace.operations.map(({ start, end, deletedText }) => [start, end, deletedText]), [[2, 4, 'ab']])
})

test('a late native undo from the former input host cannot discard a newer formula draft', () => {
  const value = recorder()
  input(value, 2, 2, 'abXabab', 'insertText', 'X')
  const base = { sessionId: value.sessionId, hostId: 'formula-host', host: 'formula' }
  value.handle({ ...base, phase: 'bind', value: value.value, start: 3, end: 5 })
  value.handle({ ...base, phase: 'beforeinput', value: value.value, start: 3, end: 5, inputType: 'insertText', data: 'Y' })
  value.handle({ ...base, phase: 'input', value: 'abXYab', start: 4, end: 4, inputType: 'insertText', data: 'Y' })
  const before = copy(value.trace)
  assert.equal(native(value, 'input', 'ababab', 2, 4, { inputType: 'historyUndo' }), false)
  assert.equal(value.value, 'abXYab'); assert.deepEqual(value.trace, before)
})
