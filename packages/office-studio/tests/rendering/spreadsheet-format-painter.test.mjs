import assert from 'node:assert/strict'
import { test } from 'vitest'
import { SpreadsheetGeometry, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../../src/rendering/spreadsheet/scene.ts'
import { sheetRangeDashedOutlineCommands } from '../../src/rendering/spreadsheet/selection.ts'

const source = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 }
const viewport = { width: 400, height: 220, scrollLeft: 0, scrollTop: 0, zoom: 1 }
function geometry(extra = {}) {
  return new SpreadsheetGeometry({ id: 'source-sheet', name: 'Source', cells: {}, rowCount: 100, columnCount: 100,
    defaultColumnWidth: 80, defaultRowHeight: 24, ...extra })
}

test('format-painter source has four dashed edges and uses the requested theme color', () => {
  const lines = sheetRangeDashedOutlineCommands(geometry(), source, viewport, '#8c6ff0')
  assert.ok(lines.length > 4)
  assert.deepEqual([...new Set(lines.filter(line => line.y1 === line.y2).map(line => line.y1))].sort((a, b) => a - b), [26.5, 122.5])
  assert.deepEqual([...new Set(lines.filter(line => line.x1 === line.x2).map(line => line.x1))].sort((a, b) => a - b), [44.5, 364.5])
  assert.ok(lines.every(line => line.color === '#8c6ff0' && Math.hypot(line.x2 - line.x1, line.y2 - line.y1) <= 4))
  const top = lines.filter(line => line.y1 === 26.5 && line.y2 === 26.5)
  assert.ok(top.some((line, index) => index > 0 && line.x1 > top[index - 1].x2))
})

test('scrolling across frozen rows and columns keeps the source perimeter on its owning panes without drawing internal seams', () => {
  const lines = sheetRangeDashedOutlineCommands(geometry({ freeze: { rows: 1, columns: 1 } }), source,
    { ...viewport, scrollLeft: 100, scrollTop: 48 })
  assert.ok(lines.length > 0)
  assert.deepEqual([...new Set(lines.filter(line => line.y1 === line.y2).map(line => line.y1))].sort((a, b) => a - b), [26.5, 74.5])
  assert.deepEqual([...new Set(lines.filter(line => line.x1 === line.x2).map(line => line.x1))].sort((a, b) => a - b), [44.5, 264.5])
  for (const line of lines) {
    assert.ok(line.x1 >= line.clip.x && line.x2 <= line.clip.x + line.clip.width)
    assert.ok(line.y1 >= line.clip.y && line.y2 <= line.clip.y + line.clip.height)
  }
})

test('offscreen and fully hidden source ranges emit no marks', () => {
  assert.deepEqual(sheetRangeDashedOutlineCommands(geometry(), source, { ...viewport, scrollLeft: 700, scrollTop: 700 }), [])
  assert.deepEqual(sheetRangeDashedOutlineCommands(geometry({ hiddenRows: [0, 1, 2, 3] }), source, viewport), [])
})

test('a whole-sheet source emits only a viewport-sized set of dashes at every zoom', () => {
  const huge = geometry({ rowCount: 1_048_576, columnCount: 16_384 })
  const range = { startRow: 0, endRow: 1_048_575, startColumn: 0, endColumn: 16_383 }
  for (const zoom of [0.5, 1, 2]) {
    const lines = sheetRangeDashedOutlineCommands(huge, range, { ...viewport, zoom })
    assert.ok(lines.length > 0 && lines.length < (viewport.width + viewport.height) / 7 + 4)
    assert.ok(lines.every(line => line.x1 >= GRID_HEADER_WIDTH && line.y1 >= GRID_HEADER_HEIGHT))
  }
})

test('source marking suppresses the fill handle while leaving the actual selection outline visible', () => {
  const options = { activeCell: 'A1', selectedCells: [], selectedRanges: [source] }
  const normal = createSpreadsheetScene(geometry(), viewport, options)
  const painting = createSpreadsheetScene(geometry(), viewport, { ...options, suppressFillHandle: true, formatPainterSource: source })
  const handles = scene => scene.commands.filter(command => command.kind === 'rect' && command.width === 5 && command.height === 5 && command.fill === '#246bfd')
  assert.equal(handles(normal).length, 1); assert.equal(handles(painting).length, 0)
  assert.ok(painting.commands.some(command => command.kind === 'line' && command.color === '#246bfd'
    && Math.hypot(command.x2 - command.x1, command.y2 - command.y1) > 4))
})
