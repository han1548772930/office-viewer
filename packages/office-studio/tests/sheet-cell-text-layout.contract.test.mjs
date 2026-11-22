import assert from 'node:assert/strict'
import { test } from 'vitest'
import { layoutSheetCellText, layoutSheetStackedText, rotatedSheetTextSize, sheetCellTextFont, sheetTextGraphemes,
  canSheetCellTextOverflow, isSheetCellCoverable } from '../src/features/sheet-workbench/sheet-cell-text-layout.ts'
import { calculateSheetAxisAutoFit } from '../src/features/sheet-workbench/sheet-auto-fit.ts'
import { createSheetCellTextCommands } from '../src/rendering/spreadsheet/cell-text.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'
import { SpreadsheetGeometry, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../src/rendering/spreadsheet/layout.ts'

const widthOf = text => sheetTextGraphemes(text).length
const measure = (text, font) => widthOf(text) * Number(/([\d.]+)px/u.exec(font)?.[1] ?? 12) / 2
const copy = value => structuredClone(value)
const view = cells => ({ id: 'text-layout', name: 'Text', cells, rowCount: 20, columnCount: 26, defaultColumnWidth: 60, defaultRowHeight: 26 })
const viewport = { width: 360, height: 220, scrollLeft: 0, scrollTop: 0, zoom: 1 }
const textCommands = scene => scene.commands.filter(command => command.kind === 'text' && command.clip?.y >= GRID_HEADER_HEIGHT)
const paint = (cell, overrides = {}) => createSheetCellTextCommands({ cell, presentation: { text: cell.value },
  rect: { x: 10, y: 20, width: 84, height: 100 }, clip: { x: 10, y: 20, width: 84, height: 100 }, zoom: 1, measure, ...overrides })

test('explicit CR/LF/CRLF breaks retain empty paragraphs and original UTF-16 ranges', () => {
  const text = 'a\r\n\r\nb\rc\n'
  const layout = layoutSheetCellText(text, { width: 20, lineHeight: 12, wrap: false, measure: widthOf })
  assert.deepEqual(layout.lines.map(line => [line.text, line.start, line.end, line.paragraphEnd]), [
    ['a', 0, 1, true], ['', 3, 3, true], ['b', 5, 6, true], ['c', 7, 8, true], ['', 9, 9, true],
  ])
  assert.equal(layout.width, 1)
  assert.equal(layout.height, 60)
})

test('word boundaries preserve spaces and long words break without dropping content', () => {
  const text = 'alpha beta gamma'
  const layout = layoutSheetCellText(text, { width: 6, lineHeight: 10, wrap: true, measure: widthOf })
  assert.deepEqual(layout.lines.map(line => line.text), ['alpha ', 'beta ', 'gamma'])
  assert.deepEqual(layout.lines.map(line => line.paragraphEnd), [false, false, true])
  assert.equal(layout.lines.map(line => line.text).join(''), text)
  const longWord = layoutSheetCellText('abcdefghij', { width: 3, lineHeight: 10, wrap: true, measure: widthOf })
  assert.deepEqual(longWord.lines.map(line => line.text), ['abc', 'def', 'ghi', 'j'])
})

test('emoji ZWJ sequences and combining marks stay intact when wrapping a long word', () => {
  const text = 'A👩🏽‍💻e\u0301B'
  const layout = layoutSheetCellText(text, { width: 2, lineHeight: 10, wrap: true, measure: widthOf })
  assert.deepEqual(layout.lines.map(line => line.text), ['A👩🏽‍💻', 'e\u0301B'])
  for (const line of layout.lines) assert.equal(text.slice(line.start, line.end), line.text)
  assert.equal(layout.lines.map(line => line.text).join(''), text)
})

test('CJK punctuation stays with surrounding text and a zero-width column still advances', () => {
  const layout = layoutSheetCellText('甲乙，丙丁。', { width: 3, lineHeight: 10, wrap: true, measure: widthOf })
  assert.deepEqual(layout.lines.map(line => line.text), ['甲乙，', '丙丁。'])
  const narrow = layoutSheetCellText('（甲）🙂', { width: 0, lineHeight: 10, wrap: true, measure: widthOf })
  assert.equal(narrow.lines.map(line => line.text).join(''), '（甲）🙂')
  assert.ok(narrow.lines.every(line => line.end > line.start))
})

test('a long unbroken value fitting a wide merged cell does not measure every growing prefix', () => {
  const text = 'x'.repeat(32_767)
  let calls = 0
  const layout = layoutSheetCellText(text, { width: 40_000, lineHeight: 10, wrap: true,
    measure: value => { calls++; return value.length } })
  assert.equal(layout.lines.length, 1)
  assert.equal(layout.lines[0].text, text)
  assert.ok(calls < 100, 'avoid a measureText call for every prefix of one long value')
})

test('text font units, explicit font overrides and editor zoom use the same point-to-pixel conversion', () => {
  const cell = { value: 'font', fontFamily: 'Test Font', fontSize: 12, bold: false, italic: true }
  const font = sheetCellTextFont(cell, 2, { bold: true, italic: false })
  assert.equal(font.fontSize, 32)
  assert.equal(font.lineHeight, 38.4)
  assert.equal(font.font, '600 32px Test Font')
  assert.equal(sheetCellTextFont({ value: '', fontSize: NaN }).fontSize, 11 * 4 / 3)
})

test('stacked text uses grapheme rows, explicit columns and available-height wrapping', () => {
  const layout = layoutSheetStackedText('A🙂B\n中e\u0301', { height: 20, lineHeight: 10, wrap: true, measure: widthOf })
  assert.deepEqual(layout.columns.map(column => column.glyphs.map(glyph => glyph.text)), [['A', '🙂'], ['B'], ['中', 'e\u0301']])
  assert.equal(layout.width, 3)
  assert.equal(layout.height, 20)
  const empty = layoutSheetStackedText('\n', { height: 20, lineHeight: 10, wrap: true, measure: widthOf })
  assert.equal(empty.columns.length, 2)
  assert.equal(empty.width, 20)
  assert.equal(empty.height, 10)
})

test('overflow follows stored scalar types and excludes clipping, wrapping, fit and stacked text', () => {
  assert.equal(canSheetCellTextOverflow({ value: '1234', rawValue: '1234', dataType: 'text' }), true)
  for (const cell of [
    { value: '1234', rawValue: 1234, dataType: 'number' }, { value: 'TRUE', rawValue: true, dataType: 'boolean' },
    { value: '', formula: '=1', cachedValue: '1', rawValue: null, dataType: 'formula', cachedDataType: 'number' },
    { value: 'text', wrap: true }, { value: 'text', wrapStrategy: 'clip' }, { value: 'text', shrinkToFit: true },
    { value: 'text', rotation: 255 }, { value: 'text', align: 'justify' },
  ]) assert.equal(canSheetCellTextOverflow(cell), false)
  assert.equal(isSheetCellCoverable({ value: '', background: '#ff0000', border: { left: 'thin' } }), true)
  for (const cell of [{ value: '0' }, { value: '', rawValue: 0 }, { value: '', formula: '=IF(1,"","")' },
    { value: '', sparkline: { type: 'line', sourceRange: 'A2:B2' } }]) assert.equal(isSheetCellCoverable(cell), false)
})

test('sparse neighbor queries ignore styled empties and hidden columns while formula cells block overflow', () => {
  const sheet = { ...view({ A1: { value: 'anchor' }, B1: { value: '', background: '#f00' }, D1: { value: 'hidden' },
    G1: { value: '', formula: '=IF(1,"","")' }, XFD1: { value: 'last' } }), rowCount: 1_048_576, columnCount: 16_384, hiddenColumns: [3] }
  const geometry = new SpreadsheetGeometry(sheet)
  assert.equal(geometry.contentColumnAfter(0, 0), 6)
  assert.equal(geometry.contentColumnBefore(0, 16_383), 6)
  assert.deepEqual(geometry.overflowColumnBounds(0, 0), { startColumn: 0, endColumn: 5 })
  assert.deepEqual(geometry.overflowColumnBounds(0, 16_383), { startColumn: 7, endColumn: 16_383 })
  assert.equal(Object.keys(sheet.cells).length, 5)
})

test('merged ranges and frozen columns form overflow barriers even when their cells are empty', () => {
  let geometry = new SpreadsheetGeometry({ ...view({ A1: { value: 'left' }, E1: { value: 'right' } }), merges: ['B1:D3'] })
  assert.deepEqual(geometry.overflowColumnBounds(0, 0), { startColumn: 0, endColumn: 0 })
  assert.deepEqual(geometry.overflowColumnBounds(0, 4), { startColumn: 4, endColumn: 25 })
  geometry = new SpreadsheetGeometry({ ...view({ A1: { value: 'frozen' }, C1: { value: 'body' } }), freeze: { rows: 0, columns: 2 } })
  assert.deepEqual(geometry.overflowColumnBounds(0, 0), { startColumn: 0, endColumn: 1 })
  assert.deepEqual(geometry.overflowColumnBounds(0, 2), { startColumn: 2, endColumn: 25 })
})

test('scene extends left, right and centered text clips only through each available empty lane', () => {
  const geometry = new SpreadsheetGeometry(view({ A1: { value: 'left '.repeat(30) }, C1: { value: 'barrier' },
    A2: { value: 'barrier' }, C2: { value: 'right '.repeat(30), align: 'right' },
    B3: { value: 'center '.repeat(30), align: 'center' }, D3: { value: 'barrier' },
    A4: { value: '123456789123456789', rawValue: 123456789123456789, dataType: 'number' } }))
  const scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'A20', selectedCells: [] })
  const texts = textCommands(scene)
  const left = texts.find(command => command.text.startsWith('left '))
  const right = texts.find(command => command.text.startsWith('right '))
  const center = texts.find(command => command.text.startsWith('center '))
  const number = texts.find(command => command.y > geometry.cellRect(3, 0, viewport).y && command.y < geometry.cellRect(3, 0, viewport).y + 26)
  assert.deepEqual([left.clip.x, left.clip.width], [47, 114])
  assert.deepEqual([right.clip.x, right.clip.width], [107, 114])
  assert.deepEqual([center.clip.x, center.clip.width], [47, 174])
  assert.equal(number.clip.width, 54)
})

test('text foreground is submitted after neighboring backgrounds, borders and crosshair fills', () => {
  const geometry = new SpreadsheetGeometry(view({ A1: { value: 'overflow '.repeat(30), underline: true }, B1: { value: '', background: '#abc123' } }))
  const scene = createSpreadsheetScene(geometry, viewport, { activeCell: 'A1', selectedCells: ['A1'], crosshairHighlight: { enabled: true, color: '#123abc' } })
  const textIndex = scene.commands.findIndex(command => command.kind === 'text' && command.text.startsWith('overflow '))
  assert.ok(textIndex > scene.commands.findIndex(command => command.kind === 'rect' && command.fill === '#abc123'))
  assert.ok(textIndex > scene.commands.findIndex(command => command.kind === 'rect' && command.fill === '#123abc'))
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.clip?.width > 54))
})

test('viewport-independent neighbor anchors preserve text entering from either horizontal side', () => {
  const sheet = view({ A1: { value: 'left anchor '.repeat(50) }, P2: { value: 'right anchor '.repeat(50), align: 'right' },
    A3: { value: '123', rawValue: 123, dataType: 'number' } })
  const scrolled = { ...viewport, width: 224, scrollLeft: 600 }
  const scene = createSpreadsheetScene(new SpreadsheetGeometry(sheet), scrolled, { activeCell: 'K10', selectedCells: [] })
  const texts = textCommands(scene)
  assert.ok(texts.some(command => command.text.startsWith('left anchor ') && command.clip.x === GRID_HEADER_WIDTH))
  assert.ok(texts.some(command => command.text.startsWith('right anchor ') && command.clip.x === GRID_HEADER_WIDTH))
  assert.equal(texts.some(command => command.text === '123'), false)
})

test('offscreen anchors stay in their frozen pane and merged origins never overflow', () => {
  const sheet = { ...view({ A1: { value: 'frozen '.repeat(50) }, C1: { value: 'body '.repeat(50) }, A2: { value: 'merged '.repeat(50) } }),
    freeze: { rows: 0, columns: 1 }, merges: ['A2:B2'] }
  const scrolled = { ...viewport, width: 280, scrollLeft: 600 }
  const texts = textCommands(createSpreadsheetScene(new SpreadsheetGeometry(sheet), scrolled, { activeCell: 'K10', selectedCells: [] }))
  const frozen = texts.filter(command => command.text.startsWith('frozen '))
  assert.equal(frozen.length, 1)
  assert.ok(frozen[0].clip.x + frozen[0].clip.width <= 104)
  assert.ok(texts.find(command => command.text.startsWith('body ')).clip.x >= 104)
  assert.ok(texts.find(command => command.text.startsWith('merged ')).clip.x + texts.find(command => command.text.startsWith('merged ')).clip.width <= 104)
})

test('wrapped and explicitly broken blocks keep line spacing and vertical alignment as one block', () => {
  const cell = { value: 'alpha beta gamma', wrap: true, fontSize: 12, vertical: 'top', underline: true }
  const top = paint(cell).filter(command => command.kind === 'text')
  assert.ok(top.length > 1)
  assert.equal(top.map(command => command.text).join(''), cell.value)
  for (let index = 1; index < top.length; index++) assert.ok(Math.abs(top[index].y - top[index - 1].y - 19.2) < 1e-9)
  const bottom = paint({ ...cell, vertical: 'bottom' }).filter(command => command.kind === 'text')
  const middle = paint({ ...cell, vertical: 'middle' }).filter(command => command.kind === 'text')
  assert.ok(Math.abs(middle[0].y - (top[0].y + bottom[0].y) / 2) < 1e-9)
  const clip = paint({ value: 'one\ntwo', wrapStrategy: 'clip' }).filter(command => command.kind === 'text')
  assert.deepEqual(clip.map(command => command.text), ['one', 'two'])
})

test('shrink-to-fit uniformly scales measured fonts, rotation and decorations without mutating styles', () => {
  const cell = { value: 'Long heading requiring fit', fontSize: 12, shrinkToFit: true, rotation: 45, underline: true }
  const before = copy(cell)
  const commands = paint(cell, { rect: { x: 10, y: 20, width: 64, height: 26 }, clip: { x: 10, y: 20, width: 64, height: 26 } })
  const text = commands.find(command => command.kind === 'text')
  const fittedSize = Number(/([\d.]+)px/u.exec(text.font)?.[1])
  assert.ok(fittedSize > 0 && fittedSize < 16)
  assert.equal(text.maxWidth, undefined)
  const size = rotatedSheetTextSize(measure(cell.value, text.font), fittedSize * 1.2, 45)
  assert.ok(size.width <= 52 + 1e-9 && size.height <= 20 + 1e-9)
  assert.ok(commands.some(command => command.kind === 'line' && command.y1 !== command.y2))
  assert.deepEqual(cell, before)
  const wrapped = paint({ ...cell, wrapStrategy: 'wrap', rotation: 0 }).find(command => command.kind === 'text')
  assert.equal(Number(/([\d.]+)px/u.exec(wrapped.font)?.[1]), 16)
})

test('number-format reservations keep their positions while shrink uses a smaller complete font', () => {
  const commands = createSheetCellTextCommands({ cell: { value: '123456', fontSize: 12, shrinkToFit: true },
    presentation: { text: '123456', runs: [{ kind: 'space', character: ')' }, { kind: 'text', text: '123456' }] },
    rect: { x: 0, y: 0, width: 40, height: 26 }, clip: { x: 0, y: 0, width: 40, height: 26 }, zoom: 1, measure })
  const text = commands.find(command => command.kind === 'text')
  assert.equal(text.maxWidth, undefined)
  assert.ok(Number(/([\d.]+)px/u.exec(text.font)?.[1]) < 16)
  assert.ok(text.x > 6)
  assert.ok(text.x + measure(text.text, text.font) <= 34 + 1e-9)
})

test('auto-fit uses point sizes, actual line breaks and the same wrap layout without changing content', () => {
  const sheet = view({ A1: { value: 'a\nlongline', fontSize: 12 }, B1: { value: 'alpha beta gamma', fontSize: 12, wrapStrategy: 'wrap' } })
  const before = copy(sheet)
  assert.equal(calculateSheetAxisAutoFit(sheet, 'column', 0, { measure }), 76)
  const layout = layoutSheetCellText('alpha beta gamma', { width: 48, lineHeight: 19.2, wrap: true, measure: text => measure(text, '16px Test') })
  assert.equal(calculateSheetAxisAutoFit(sheet, 'row', 0, { measure }), Math.ceil(layout.height + 6))
  assert.deepEqual(sheet, before)
})

test('auto-fit skips hidden rows and merged columns, retains empty widths and respects finite bounds', () => {
  const sheet = { ...view({ A1: { value: 'visible', fontSize: 12 }, A2: { value: 'hidden '.repeat(50), fontSize: 12 },
    B1: { value: 'merged '.repeat(50), fontSize: 12 }, C1: { value: 'x'.repeat(300), fontSize: 12 } }),
    hiddenRows: [1], columnWidths: { 3: 77 }, merges: ['B1:B3'] }
  const geometry = new SpreadsheetGeometry(sheet), options = { measure, mergeAt: (row, column) => geometry.mergeAt(row, column) }
  assert.equal(calculateSheetAxisAutoFit(sheet, 'column', 0, options), 68)
  assert.equal(calculateSheetAxisAutoFit(sheet, 'column', 1, options), 60)
  assert.equal(calculateSheetAxisAutoFit(sheet, 'column', 3, options), 77)
  assert.equal(calculateSheetAxisAutoFit(sheet, 'column', 2, options), 1000)
  assert.throws(() => calculateSheetAxisAutoFit(sheet, 'row', -1, options), RangeError)
  assert.throws(() => calculateSheetAxisAutoFit(sheet, 'column', 26, options), RangeError)
})
