import assert from 'node:assert/strict'
import { test } from 'vitest'
import { layoutSheetCellRichText } from '../src/features/sheet-workbench/sheet-cell-rich-text-layout.ts'
import { layoutSheetStackedText } from '../src/features/sheet-workbench/sheet-cell-text-layout.ts'
import { calculateSheetAxisAutoFit } from '../src/features/sheet-workbench/sheet-auto-fit.ts'
import { createSheetCellTextCommands } from '../src/rendering/spreadsheet/cell-text.ts'
import { createSheetRichTextEditorScene, hitSheetRichTextEditor, hitSheetRichTextEditorPosition, moveSheetRichTextEditorCaret, sheetRichTextEditorCaret } from '../src/rendering/spreadsheet/rich-text-editor.ts'

// Draft only. Deterministic measurements verify geometry contracts, not browser font shaping.
const size = font => Number(/([\d.]+)px/u.exec(font)[1])
const measure = (text, font) => [...text].length * size(font) / 2
const rect = { x: 0, y: 0, width: 160, height: 90 }
const appearance = { background: '#ffffff', color: '#ddeeff', accent: '#112233', selectionColor: '#ccddff', measure }
const mixed = () => ({ value: 'aa\nbb', rawValue: 'aa\nbb', fontSize: 11, vertical: 'top',
  richText: [{ text: 'aa\n', font: { family: 'Arial', size: 9 } }, { text: 'bb', font: { family: 'Georgia', size: 18 }, color: '#ff0000' }] })

test('mixed font lines use actual run widths, variable heights and original UTF-16 paragraph offsets', () => {
  const cell = mixed(), layout = layoutSheetCellRichText(cell, cell.value, { width: 150, height: Infinity, wrap: false, measure })
  assert.deepEqual(layout.lines.map(line => [line.start, line.end]), [[0, 2], [3, 5]])
  assert.deepEqual(layout.lines.map(line => line.width), [12, 24])
  assert.ok(Math.abs(layout.lines[0].lineHeight - 14.4) < 1e-8)
  assert.ok(Math.abs(layout.lines[1].top - 14.4) < 1e-8)
  assert.ok(Math.abs(layout.height - 43.2) < 1e-8)
})

test('identical text with different run fonts breaks by source ranges instead of a string lookup', () => {
  const cell = { value: 'ab ab', rawValue: 'ab ab', richText: [{ text: 'ab ', fontSize: 9 }, { text: 'ab', fontSize: 18 }] }
  const layout = layoutSheetCellRichText(cell, cell.value, { width: 25, height: Infinity, wrap: true, measure })
  assert.deepEqual(layout.lines.map(line => [line.text, line.start, line.end]), [['ab ', 0, 3], ['ab', 3, 5]])
  assert.deepEqual(layout.lines.map(line => line.width), [18, 24])
})

test('an imported run boundary inside a grapheme uses its leading character font without drawing split marks', () => {
  const cell = { value: 'e\u0301x', rawValue: 'e\u0301x', richText: [{ text: 'e', fontSize: 9 }, { text: '\u0301x', fontSize: 24 }] }
  const layout = layoutSheetCellRichText(cell, cell.value, { width: 200, height: Infinity, wrap: false, measure })
  assert.deepEqual(layout.lines[0].fragments.map(fragment => [fragment.text, fragment.start, fragment.end, fragment.fontSize]),
    [['e\u0301', 0, 2, 12], ['x', 2, 3, 32]])
})

test('backend-neutral text commands retain individual fonts, colors and hyperlink decorations', () => {
  const cell = { value: 'abcd', rawValue: 'abcd', fontSize: 11, color: '#556677',
    richText: [{ text: 'ab', fontFamily: 'Arial', fontSize: 9, strike: true },
      { text: 'cd', font: { family: 'Georgia', size: 18 }, hyperlink: 'https://example.test' }] }
  const commands = createSheetCellTextCommands({ cell, presentation: { text: cell.value }, rect, clip: rect, zoom: 1, measure })
  const text = commands.filter(command => command.kind === 'text')
  assert.deepEqual(text.map(command => [command.text, command.color, size(command.font)]), [['ab', '#556677', 12], ['cd', '#0563c1', 24]])
  assert.match(text[0].font, /Arial/u); assert.match(text[1].font, /Georgia/u)
  assert.equal(commands.filter(command => command.kind === 'line').length, 2)
})

test('rotation and shrink-to-fit scale all run fonts together and preserve their relative sizes', () => {
  const cell = { ...mixed(), value: 'aabb', rawValue: 'aabb', rotation: 45, shrinkToFit: true,
    richText: [{ text: 'aa', fontSize: 9 }, { text: 'bb', fontSize: 18 }] }
  const commands = createSheetCellTextCommands({ cell, presentation: { text: cell.value }, rect: { ...rect, width: 40 }, clip: rect, zoom: 1, measure })
  const text = commands.filter(command => command.kind === 'text')
  assert.equal(text.length, 2); assert.ok(text.every(command => command.rotation === 45))
  assert.ok(size(text[1].font) < 24); assert.ok(Math.abs(size(text[1].font) / size(text[0].font) - 2) < 1e-8)
})

test('stacked rich text uses each glyph font height and rejects nonfinite optional glyph measurements', () => {
  const cell = { value: 'ab', rawValue: 'ab', richText: [{ text: 'a', fontSize: 9 }, { text: 'b', fontSize: 18 }] }
  const layout = layoutSheetCellRichText(cell, cell.value, { width: 100, height: Infinity, wrap: false, stacked: true, measure })
  assert.equal(layout.columns[0].height, 36); assert.deepEqual(layout.columns[0].glyphs.map(glyph => glyph.top), [0, 12])
  const invalid = layoutSheetStackedText('ab', { height: 20, lineHeight: 10, wrap: true, measure: text => text.length * 5,
    glyphHeight: start => start === 0 ? NaN : Infinity })
  assert.equal(invalid.columns[0].height, 20); assert.equal(invalid.height, 20)
})

test('auto-fit derives dimensions from the same rich text layout and formula/number displays keep their separate path', () => {
  const cell = mixed(), sheet = { id: 'sheet', name: 'Sheet', cells: { A1: cell }, rowCount: 10, columnCount: 5, defaultColumnWidth: 160, defaultRowHeight: 26 }
  const layout = layoutSheetCellRichText(cell, cell.value, { width: 148, height: Infinity, wrap: false, measure })
  assert.equal(calculateSheetAxisAutoFit(sheet, 'row', 0, { measure }), Math.ceil(layout.height + 6))
  assert.equal(calculateSheetAxisAutoFit(sheet, 'column', 0, { measure }), Math.ceil(layout.width + 12))
  assert.equal(layoutSheetCellRichText({ ...cell, rawValue: 123, value: '123' }, '123', { width: 100, height: 100, wrap: false, measure }), undefined)
  assert.equal(layoutSheetCellRichText({ ...cell, formula: '=1' }, cell.value, { width: 100, height: 100, wrap: false, measure }), undefined)
})

test('justification keeps leading whitespace and source offsets for plain and rich text', () => {
  const value = '  ab cd ef gh ij kl'
  for (const rich of [false, true]) {
    const cell = { value, rawValue: value, align: 'justify', wrapStrategy: 'wrap', fontSize: 11,
      ...(rich ? { richText: [{ text: value, font: { italic: true } }] } : {}) }
    const commands = createSheetCellTextCommands({ cell, presentation: { text: value }, rect: { ...rect, width: 112, height: 300 }, clip: rect, zoom: 1, measure })
    assert.equal(commands.filter(command => command.kind === 'text').map(command => command.text).join(''), value)
  }
})

test('rich editor hit geometry respects supplementary characters and the fonts that Canvas actually draws', () => {
  const cell = { value: 'A🙂BC', rawValue: 'A🙂BC', richText: [{ text: 'A', fontSize: 9 }, { text: '🙂BC', fontSize: 18 }] }
  const scene = createSheetRichTextEditorScene(cell, rect, rect, 1, { start: 1, end: 3 }, appearance)
  const line = scene.lines[0]
  assert.deepEqual(line.offsets, [0, 1, 3, 4, 5])
  assert.equal(hitSheetRichTextEditor(scene, line.xAt(1), line.top + 2), 1)
  assert.equal(hitSheetRichTextEditor(scene, line.xAt(3), line.top + 2), 3)
  assert.equal(scene.commands.filter(command => command.kind === 'rect').length, 2)
  assert.equal(sheetRichTextEditorCaret(scene, 5).x, line.xAt(5))
})

test('visual line navigation and scroll share variable-height line coordinates', () => {
  const cell = mixed(), scene = createSheetRichTextEditorScene(cell, rect, rect, 1, { start: 1, end: 1 }, { ...appearance, scrollLeft: 2, scrollTop: 4 })
  const down = moveSheetRichTextEditorCaret(scene, 1, 'ArrowDown')
  assert.equal(down.offset, 4); assert.equal(moveSheetRichTextEditorCaret(scene, 4, 'Home').offset, 3)
  assert.equal(moveSheetRichTextEditorCaret(scene, 4, 'End').offset, 5)
  assert.equal(scene.lines[0].top, -1); assert.ok(scene.contentHeight > 43)
})

test('editor default colors preserve explicit run colors, justification is disabled while editing and IME uses the native host', () => {
  const cell = { value: 'abcd', rawValue: 'abcd', align: 'justify', richText: [{ text: 'ab' }, { text: 'cd', color: '#ff0000' }] }
  const scene = createSheetRichTextEditorScene(cell, rect, rect, 1, { start: 4, end: 4 }, appearance)
  assert.deepEqual(scene.commands.filter(command => command.kind === 'text').map(command => command.color), ['#ddeeff', '#ff0000'])
  const composing = createSheetRichTextEditorScene(cell, rect, rect, 1, { start: 4, end: 4 }, { ...appearance, composing: true })
  assert.equal(composing.commands.filter(command => command.kind === 'text' || command.kind === 'line').length, 0)
  assert.equal(composing.commands[0].fill, '#ffffff')
})

test('long editor scene construction does not shape every growing character prefix eagerly', () => {
  const value = 'a'.repeat(12_000), cell = { value, rawValue: value, richText: [{ text: value, fontSize: 11 }] }
  let measuredCharacters = 0
  const scene = createSheetRichTextEditorScene(cell, rect, rect, 1, { start: 0, end: 0 }, { ...appearance,
    measure: (text, font) => { measuredCharacters += text.length; return measure(text, font) } })
  assert.ok(measuredCharacters < value.length * 8)
  assert.equal(scene.lines[0].offsets.length, value.length + 1)
  assert.ok(hitSheetRichTextEditor(scene, 40, 5) > 0)
})

test('wrapped boundaries keep distinct visual caret rows for one native UTF-16 offset', () => {
  const cell = { value: 'abcd', rawValue: 'abcd', wrapStrategy: 'wrap', richText: [{ text: 'abcd', fontSize: 9 }] }
  const area = { ...rect, width: 24 }, scene = createSheetRichTextEditorScene(cell, area, area, 1, { start: 2, end: 2 }, appearance)
  assert.deepEqual(scene.lines.map(line => [line.start, line.end]), [[0, 2], [2, 4]])
  assert.equal(sheetRichTextEditorCaret(scene, 2, 'upstream').lineIndex, 0)
  assert.equal(sheetRichTextEditorCaret(scene, 2, 'downstream').lineIndex, 1)
  const upstream = createSheetRichTextEditorScene(cell, area, area, 1, { start: 2, end: 2 }, { ...appearance, caretAffinity: 'upstream' })
  const downstream = createSheetRichTextEditorScene(cell, area, area, 1, { start: 2, end: 2 }, { ...appearance, caretAffinity: 'downstream' })
  assert.equal(upstream.commands.filter(command => command.kind === 'line').at(-1).y1, upstream.lines[0].top)
  assert.equal(downstream.commands.filter(command => command.kind === 'line').at(-1).y1, downstream.lines[1].top)
})

test('mouse hits at the end and start of adjacent wrapped rows share offsets while retaining their own affinity', () => {
  const cell = { value: 'abcd', rawValue: 'abcd', wrapStrategy: 'wrap', richText: [{ text: 'abcd', fontSize: 9 }] }
  const area = { ...rect, width: 24 }, scene = createSheetRichTextEditorScene(cell, area, area, 1, { start: 0, end: 0 }, appearance)
  const [first, second] = scene.lines
  const end = hitSheetRichTextEditorPosition(scene, first.xAt(first.end) + 100, first.top + first.height / 2)
  const start = hitSheetRichTextEditorPosition(scene, second.xAt(second.start) - 100, second.top + second.height / 2)
  assert.deepEqual(end, { offset: 2, affinity: 'upstream' }); assert.deepEqual(start, { offset: 2, affinity: 'downstream' })
  assert.equal(sheetRichTextEditorCaret(scene, end.offset, end.affinity).lineIndex, 0)
  assert.equal(sheetRichTextEditorCaret(scene, start.offset, start.affinity).lineIndex, 1)
})

test('repeated End and subsequent Home stay on the selected visual line and vertical motion preserves its target end', () => {
  const cell = { value: 'abcdef', rawValue: 'abcdef', wrapStrategy: 'wrap', richText: [{ text: 'abcdef', fontSize: 9 }] }
  const area = { ...rect, width: 24 }, scene = createSheetRichTextEditorScene(cell, area, area, 1, { start: 0, end: 0 }, appearance)
  const end = moveSheetRichTextEditorCaret(scene, 0, 'End')
  assert.deepEqual(end, { offset: 2, affinity: 'upstream' })
  const focused = { ...scene, caretAffinity: end.affinity }
  assert.deepEqual(moveSheetRichTextEditorCaret(focused, end.offset, 'End'), end)
  assert.deepEqual(moveSheetRichTextEditorCaret(focused, end.offset, 'Home'), { offset: 0, affinity: 'downstream' })
  const down = moveSheetRichTextEditorCaret(focused, end.offset, 'ArrowDown')
  assert.equal(down.offset, 4); assert.equal(down.affinity, 'upstream')
  assert.equal(sheetRichTextEditorCaret({ ...scene, caretAffinity: down.affinity }, down.offset).lineIndex, 1)
  const up = moveSheetRichTextEditorCaret({ ...scene, caretAffinity: down.affinity }, down.offset, 'ArrowUp', down.preferredX)
  assert.equal(up.offset, 2); assert.equal(up.affinity, 'upstream')
})

