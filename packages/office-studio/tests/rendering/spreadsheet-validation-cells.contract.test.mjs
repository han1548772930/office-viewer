import assert from 'node:assert/strict'
import { test } from 'vitest'
import { SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../../src/rendering/spreadsheet/scene.ts'
import { sheetValidationCellControlRect } from '../../src/rendering/spreadsheet/validation-cell-layout.ts'
import { createSheetValidationTagCommands } from '../../src/rendering/spreadsheet/validation-cell-tags.ts'
import { serializeValidationListItems } from '../../src/features/sheet-validation/validation-engine.ts'

// Unexecuted implementation-phase contracts. These checks do not stand in for browser/backend acceptance.
const area = (row, column, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
function sheet() {
  return { id: 'data', name: 'Data', rowCount: 100, columnCount: 20, defaultRowHeight: 30, defaultColumnWidth: 140,
    cells: { A1: { value: 'Red,"Green, dark",It""s blue', rawValue: 'Red,"Green, dark",It""s blue', dataType: 'text' }, B1: { value: 'other', rawValue: 'other' } },
    dataValidations: [{ id: 'list', type: 'listMultiple', ranges: [area(0, 0)], formula1: '"Red,"Green, dark",It""s blue"', formula2: '#ff0000,#00ff00,#0000ff' }],
    merges: [], showGridlines: true, columnWidths: {}, rowHeights: {}, freeze: { rows: 0, columns: 0 } }
}
const viewport = (extra = {}) => ({ width: 720, height: 360, scrollLeft: 0, scrollTop: 0, zoom: 1, ...extra })
const theme = { fill: '#e7e9ed', text: '#222222' }, measure = (_font, text) => text.length * 7
const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }

test('a CSV multi-value cell paints separate colored labels and retains separators and original model data', () => {
  const source = sheet(), before = JSON.stringify(source)
  const commands = createSheetValidationTagCommands(source.dataValidations[0], source.cells.A1,
    { x: 50, y: 30, width: 400, height: 40 }, { x: 50, y: 30, width: 400, height: 40 }, 1, theme, measure)
  assert.deepEqual(commands.filter(command => command.kind === 'text').map(command => command.text), ['Red', 'Green, dark', 'It"s blue'])
  assert.deepEqual(commands.filter(command => command.kind === 'polygon').map(command => command.fill), ['#ff0000', '#00ff00', '#0000ff'])
  assert.equal(JSON.stringify(source), before)
  for (const command of commands) { assert.ok(command.clip.width > 0 && command.clip.height > 0); assert.ok(command.clip.x >= 50 && command.clip.y >= 30) }
})

test('tag wrapping emits only visible labels even when a cell contains thousands of selections', () => {
  const labels = Array.from({ length: 3000 }, (_, index) => `v${index}`), value = serializeValidationListItems(labels).slice(1, -1)
  const rule = { type: 'listMultiple', renderMode: 'custom', formula1: serializeValidationListItems(labels) }
  const commands = createSheetValidationTagCommands(rule, { value, rawValue: value, vertical: 'top', fontSize: 11 },
    { x: 0, y: 0, width: 220, height: 24 }, { x: 0, y: 0, width: 220, height: 24 }, 1, theme, measure)
  assert.ok(commands.length > 0 && commands.length < 30)
  for (const command of commands) assert.ok(command.clip.height <= 24 && command.clip.width <= 220)
})

test('default tokens follow the supplied theme while authored label colors and font decorations remain observable', () => {
  const rule = { type: 'listMultiple', formula1: '"Red,Blue"', formula2: '#ff0000,' }, cell = { value: 'Red,Blue', rawValue: 'Red,Blue', bold: true, italic: true, underline: true, strike: true }
  const rect = { x: 0, y: 0, width: 280, height: 35 }
  const light = createSheetValidationTagCommands(rule, cell, rect, rect, 1, theme, measure)
  const dark = createSheetValidationTagCommands(rule, cell, rect, rect, 1, { fill: '#333333', text: '#eeeeee' }, measure)
  assert.deepEqual(light.filter(command => command.kind === 'polygon').map(command => command.fill), ['#ff0000', '#e7e9ed'])
  assert.deepEqual(dark.filter(command => command.kind === 'polygon').map(command => command.fill), ['#ff0000', '#333333'])
  assert.equal(dark.filter(command => command.kind === 'line').length, 4)
  for (const command of dark.filter(command => command.kind === 'text')) { assert.equal(command.color, '#eeeeee'); assert.match(command.font, /italic 600/) }
})

test.each([0.5, 1, 2])('zoom %s uses one visible control inside the proper merged/frozen pane', zoom => {
  const source = sheet(); source.freeze = { rows: 1, columns: 1 }; source.merges = ['A1:C3']
  const geometry = new SpreadsheetGeometry(source), view = viewport({ zoom, scrollLeft: 30, scrollTop: 20 }), before = JSON.stringify(source)
  const rect = sheetValidationCellControlRect(geometry, view, 'B2')
  assert.ok(rect && rect.width >= 10 && rect.height >= 10)
  const panes = geometry.panes(view)
  assert.ok(panes.some(pane => rect.x >= pane.clip.x && rect.x + rect.width <= pane.clip.x + pane.clip.width
    && rect.y >= pane.clip.y && rect.y + rect.height <= pane.clip.y + pane.clip.height))
  assert.equal(JSON.stringify(source), before)
})

test('offscreen, hidden and invalid targets have no floating control, while small visible cells remain bounded', () => {
  const source = sheet(), geometry = new SpreadsheetGeometry(source)
  assert.equal(sheetValidationCellControlRect(geometry, viewport(), 'Z999'), undefined)
  assert.equal(sheetValidationCellControlRect(geometry, viewport({ zoom: 0 }), 'A1'), undefined)
  assert.equal(sheetValidationCellControlRect(geometry, viewport({ scrollLeft: 500, scrollTop: 100 }), 'A1'), undefined)
  source.hiddenRows = [0]; source.hiddenColumns = [0]
  assert.equal(sheetValidationCellControlRect(new SpreadsheetGeometry(source), viewport(), 'A1'), undefined)
  const small = sheet(); small.defaultRowHeight = 14; small.defaultColumnWidth = 15
  const visible = sheetValidationCellControlRect(new SpreadsheetGeometry(small), viewport(), 'A1')
  assert.ok(visible && visible.width <= 13 && visible.height <= 12)
})

test('scene integration draws tags instead of raw CSV while keeping selection and fill-handle commands', () => {
  const source = sheet(); source.columnWidths[0] = 420
  const geometry = new SpreadsheetGeometry(source), scene = createSpreadsheetScene(geometry, viewport(), {
    activeCell: 'A1', selectedCells: [], selectedRanges: [range], validationTagTheme: theme,
  })
  assert.ok(scene.commands.some(command => command.kind === 'polygon' && command.fill === '#ff0000'))
  assert.equal(scene.commands.some(command => command.kind === 'text' && command.text === source.cells.A1.value), false)
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#246bfd'))
  assert.ok(scene.commands.some(command => command.kind === 'rect' && command.fill === '#246bfd'))
})

test('view-denied cells disclose neither normal text nor colored tags in the Canvas command stream', () => {
  const source = sheet(), geometry = new SpreadsheetGeometry(source)
  const scene = createSpreadsheetScene(geometry, viewport(), { activeCell: 'A1', selectedCells: [], selectedRanges: [range],
    validationTagTheme: theme, canViewCell: (row, column) => row !== 0 || column > 1 })
  assert.equal(scene.commands.some(command => command.kind === 'text' && ['Red', 'Green, dark', 'It"s blue', 'other', source.cells.A1.value].includes(command.text)), false)
  assert.equal(scene.commands.some(command => command.kind === 'polygon'), false)
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#246bfd'))
})

test('text render mode and plain single-arrow cells retain ordinary text; custom single cells use a token', () => {
  const rect = { x: 0, y: 0, width: 280, height: 35 }, cell = { value: '1', rawValue: '1', dataType: 'text' }
  assert.equal(createSheetValidationTagCommands({ type: 'listMultiple', renderMode: 'text' }, cell, rect, rect, 1, theme, measure), undefined)
  assert.equal(createSheetValidationTagCommands({ type: 'list', renderMode: 'arrow', formula1: '"1,2"' }, cell, rect, rect, 1, theme, measure), undefined)
  const custom = createSheetValidationTagCommands({ type: 'list', renderMode: 'custom', formula1: '"1,2"' }, cell, rect, rect, 1, theme, measure)
  assert.equal(custom.filter(command => command.kind === 'polygon').length, 1); assert.equal(custom.find(command => command.kind === 'text').text, '1')
})

test.each(['arrow', 'text', 'custom'])('referenced single %s cells use the supplied formatted colors without evaluating a source in the renderer', renderMode => {
  const source = sheet()
  source.cells = { A1: { value: 'red', rawValue: 'red', dataType: 'text' } }
  source.dataValidations = [{ id: 'referenced', type: 'list', ranges: [area(0, 0)], formula1: '=$D$1:$D$2', renderMode }]
  const geometry = new SpreadsheetGeometry(source), before = JSON.stringify(source)
  let reads = 0
  const scene = createSpreadsheetScene(geometry, viewport(), { activeCell: 'A1', selectedCells: [], validationTagTheme: theme,
    validationTagOptions: (row, column) => { reads++; assert.deepEqual([row, column], [0, 0]); return [{ label: 'Red', color: '#008800' }] } })
  assert.equal(reads, 1)
  assert.ok(scene.commands.some(command => command.kind === (renderMode === 'custom' ? 'polygon' : 'rect') && command.fill === '#008800'))
  assert.ok(scene.commands.some(command => command.kind === 'text' && command.text === (renderMode === 'custom' ? 'Red' : 'red')))
  assert.equal(JSON.stringify(source), before)
})

test.each([
  ['arrow', true, 22], ['arrow', false, 0], ['text', true, 0],
])('%s with showDropDown=%s clips to its own text rectangle and reserves %s trailing pixels', (renderMode, showDropDown, trailing) => {
  const source = sheet(), value = 'a long value that would normally overflow neighboring empty cells'
  source.defaultColumnWidth = 80; source.cells = { A1: { value, rawValue: value, dataType: 'text', wrapStrategy: 'overflow' } }
  source.dataValidations = [{ id: 'list', type: 'list', ranges: [area(0, 0)], formula1: serializeValidationListItems([value]), renderMode, showDropDown }]
  const geometry = new SpreadsheetGeometry(source), view = viewport(), rect = geometry.cellRect(0, 0, view)
  const scene = createSpreadsheetScene(geometry, view, { activeCell: 'A1', selectedCells: [] })
  const command = scene.commands.find(command => command.kind === 'text' && command.text === value)
  assert.ok(command); assert.ok(command.clip.x >= rect.x); assert.ok(command.clip.width <= rect.width - trailing - 6 + 1e-8)
  assert.ok(command.clip.x + command.clip.width <= rect.x + rect.width - trailing)
  const offscreen = createSpreadsheetScene(geometry, viewport({ scrollLeft: 200 }), { activeCell: 'A1', selectedCells: [] })
  assert.equal(offscreen.commands.some(command => command.kind === 'text' && command.text === value), false)
})

test('a view-denied target does not invoke the formatted candidate callback or disclose its selected background', () => {
  const source = sheet(); source.cells = { A1: { value: 'red', rawValue: 'red' } }
  source.dataValidations = [{ id: 'list', type: 'list', ranges: [area(0, 0)], formula1: '=$D$1', renderMode: 'arrow' }]
  let reads = 0
  const scene = createSpreadsheetScene(new SpreadsheetGeometry(source), viewport(), { activeCell: 'A1', selectedCells: [],
    canViewCell: (row, column) => row !== 0 || column !== 0,
    validationTagOptions: () => { reads++; return [{ label: 'red', color: '#123456' }] } })
  assert.equal(reads, 0)
  assert.equal(scene.commands.some(command => command.fill === '#123456' || command.kind === 'text' && command.text === 'red'), false)
})
