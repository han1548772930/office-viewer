import assert from 'node:assert/strict'
import { test } from 'vitest'
import { SpreadsheetGeometry, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../src/rendering/spreadsheet/layout.ts'
import { sheetRangeDashedOutlineCommands } from '../src/rendering/spreadsheet/selection.ts'
import { sheetClipboardOverlayPosition } from '../src/components/spreadsheet/sheet-clipboard-overlay.ts'

// Geometry drafts only. Canvas2D/WebGPU, themes and real pointer focus remain unverified.
const sheet = extra => ({ id: 'overlay-sheet', name: 'Overlay', rowCount: 100, columnCount: 20,
  defaultRowHeight: 26, defaultColumnWidth: 96, cells: {}, ...extra })
const viewport = extra => ({ width: 300, height: 200, zoom: 1, scrollLeft: 0, scrollTop: 0, ...extra })
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })

test('the paste button follows the actual endpoint and stays in CSS pixels at editor zoom', () => {
  const geometry = new SpreadsheetGeometry(sheet())
  assert.deepEqual(sheetClipboardOverlayPosition(geometry, viewport(), range(1, 1)), { left: 240, top: 82 })
  assert.deepEqual(sheetClipboardOverlayPosition(geometry, viewport({ width: 500, zoom: 2 }), range(0, 0)), { left: 240, top: 82 })
})

test('frozen endpoints stay anchored while scrolling endpoints hidden behind frozen panes disappear', () => {
  const geometry = new SpreadsheetGeometry(sheet({ freeze: { rows: 1, columns: 1 } }))
  const view = viewport({ width: 500, scrollLeft: 700, scrollTop: 700 })
  assert.deepEqual(sheetClipboardOverlayPosition(geometry, view, range(0, 0)), { left: 144, top: 56 })
  assert.equal(sheetClipboardOverlayPosition(geometry, view, range(1, 1)), undefined)
  assert.equal(sheetClipboardOverlayPosition(new SpreadsheetGeometry(sheet()), view, range(0, 0)), undefined)
})

test('partially visible endpoints keep the button inside the viewport and tiny or invalid viewports hide it', () => {
  const geometry = new SpreadsheetGeometry(sheet())
  assert.deepEqual(sheetClipboardOverlayPosition(geometry, viewport({ width: 200, height: 100 }), range(1, 1)), { left: 152, top: 72 })
  for (const view of [viewport({ width: 70 }), viewport({ height: 50 }), viewport({ zoom: 0 }), viewport({ zoom: Infinity }),
    viewport({ width: Infinity }), viewport({ height: NaN }), viewport({ scrollTop: Infinity }), viewport({ scrollLeft: NaN })]) {
    assert.equal(sheetClipboardOverlayPosition(geometry, view, range(1, 1)), undefined)
  }
  assert.equal(sheetClipboardOverlayPosition(geometry, viewport(), range(100, 0)), undefined)
})

test('hidden or zero-size endpoint axes do not leave a clickable paste button over unrelated cells', () => {
  for (const geometry of [new SpreadsheetGeometry(sheet({ hiddenRows: [0] })), new SpreadsheetGeometry(sheet({ hiddenColumns: [0] })),
    new SpreadsheetGeometry(sheet({ rowHeights: { 0: 0 } })), new SpreadsheetGeometry(sheet({ columnWidths: { 0: 0 } }))]) {
    assert.equal(sheetClipboardOverlayPosition(geometry, viewport(), range(0, 0)), undefined)
  }
})

test('the existing copy outline emits clipped line commands for whole-sheet and frozen selections', () => {
  const geometry = new SpreadsheetGeometry(sheet({ rowCount: 1_048_576, columnCount: 16_384, freeze: { rows: 1, columns: 1 } }))
  const view = viewport({ width: 600, height: 400, scrollLeft: 200, scrollTop: 500 })
  const commands = sheetRangeDashedOutlineCommands(geometry, { startRow: 0, endRow: 1_048_575, startColumn: 0, endColumn: 16_383 }, view, '#3064c7')
  assert.ok(commands.length > 0 && commands.length < 500)
  for (const command of commands) {
    assert.equal(command.kind, 'line'); assert.equal(command.color, '#3064c7')
    assert.ok(command.clip.x >= GRID_HEADER_WIDTH && command.clip.y >= GRID_HEADER_HEIGHT)
    assert.ok(command.clip.x + command.clip.width <= view.width && command.clip.y + command.clip.height <= view.height)
  }
})
