import assert from 'node:assert/strict'
import { test } from 'vitest'
import { intersectRect } from '../../src/rendering/geometry.ts'
import { SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../../src/rendering/spreadsheet/scene.ts'
import { sheetCheckboxCellPositions, sheetCheckboxRect } from '../../src/rendering/spreadsheet/validation-checkbox-layout.ts'
import { sheetValidationCellVisibleRect } from '../../src/rendering/spreadsheet/validation-cell-layout.ts'
import { createSheetValidationInvalidMarker } from '../../src/rendering/spreadsheet/validation-cell-markers.ts'

// Shared geometry/Canvas command drafts. Ant focus, browser clipping and both render backends remain unexecuted.
const viewport = (extra = {}) => ({ width: 720, height: 360, scrollLeft: 0, scrollTop: 0, zoom: 1, ...extra })
const selection = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
function sheet() {
  return { id: 'data', name: 'Data', rowCount: 100, columnCount: 20, defaultRowHeight: 30, defaultColumnWidth: 120,
    cells: { A1: { value: '0', rawValue: 0, fontSize: 10 }, B1: { value: '1', rawValue: 1, fontSize: 10 },
      A2: { value: 'invalid checkbox text', rawValue: 'invalid checkbox text', fontSize: 10 }, B2: { value: '1', rawValue: 1, fontSize: 10 } },
    merges: [], columnWidths: {}, rowHeights: {}, freeze: { rows: 0, columns: 0 }, showGridlines: true }
}
const bounds = { x: 50, y: 30, width: 100, height: 40 }
test.each([
  ['left', 'top', 56, 30], ['center', 'middle', 92, 42], ['right', 'bottom', 128, 54],
])('checkbox %s/%s alignment follows the real cell rectangle and fixed font sizing', (align, vertical, x, y) => {
  assert.deepEqual(sheetCheckboxRect({ fontSize: 10, align, vertical }, bounds, 1), { x, y, width: 16, height: 16 })
  const enlarged = sheetCheckboxRect({ fontSize: 10, align, vertical }, { x: 75, y: 45, width: 150, height: 60 }, 1.5)
  assert.deepEqual(enlarged, { x: x * 1.5, y: y * 1.5, width: 24, height: 24 })
})

test('invalid or absent font sizes use the same default without invalid dimensions', () => {
  for (const fontSize of [undefined, 0, -3, NaN, Infinity]) {
    assert.deepEqual(sheetCheckboxRect({ fontSize }, bounds, 1), { x: 92, y: 42, width: 16, height: 16 })
  }
})

test('frozen row/column controls share viewport hit testing through zoom and independent scroll offsets', () => {
  const source = sheet(); source.freeze = { rows: 1, columns: 1 }
  const geometry = new SpreadsheetGeometry(source), view = viewport({ zoom: 1.25, scrollLeft: 40, scrollTop: 10 })
  const positions = sheetCheckboxCellPositions(geometry, view, (row, column) => row < 2 && column < 2)
  assert.deepEqual(new Set(positions.map(position => position.key)), new Set(['A1', 'B1', 'A2', 'B2']))
  for (const position of positions) {
    const center = { x: position.rect.x + position.rect.width / 2, y: position.rect.y + position.rect.height / 2 }
    assert.equal(geometry.hitTest(center.x, center.y, view).key, position.key)
    assert.equal(position.rect.width, 20); assert.ok(position.clips.length >= 1)
    assert.ok(position.clips.some(clip => intersectRect(position.rect, clip).width > 0))
  }
})

test('one merged Ant checkbox combines all visible frozen-pane clips instead of truncating the control to a single pane', () => {
  const source = sheet(); source.merges = ['A1:B2']; source.freeze = { rows: 1, columns: 1 }
  const geometry = new SpreadsheetGeometry(source), calls = [], positions = sheetCheckboxCellPositions(geometry, viewport(), (row, column) => {
    calls.push([row, column]); return row === 0 && column === 0
  })
  assert.equal(positions.length, 1); assert.equal(positions[0].key, 'A1'); assert.equal(positions[0].clips.length, 4)
  assert.equal(calls.filter(([row, column]) => row === 0 && column === 0).length, 1)
  const position = positions[0], visibleArea = position.clips.reduce((sum, clip) => {
    const part = intersectRect(position.rect, clip); return sum + part.width * part.height
  }, 0)
  assert.equal(visibleArea, position.rect.width * position.rect.height)
  assert.deepEqual(position.clip, geometry.cellRect(0, 0, viewport()))
})

test('a scrolled merged checkbox is clipped to the remaining actual cell fragments', () => {
  const source = sheet(); source.merges = ['A1:B2']; source.freeze = { rows: 1, columns: 1 }
  const geometry = new SpreadsheetGeometry(source), view = viewport({ scrollLeft: 240, scrollTop: 120 })
  const positions = sheetCheckboxCellPositions(geometry, view, (row, column) => row === 0 && column === 0)
  assert.equal(positions.length, 1); assert.equal(positions[0].clips.length, 1)
  const part = intersectRect(positions[0].rect, positions[0].clip)
  assert.equal(part.width, 8); assert.equal(part.height, 8)
})

test('hidden axes, absent candidates and invalid zoom produce no unreachable controls', () => {
  const source = sheet(); source.hiddenRows = [1]; source.hiddenColumns = [1]
  const geometry = new SpreadsheetGeometry(source)
  assert.deepEqual(sheetCheckboxCellPositions(geometry, viewport(), (row, column) => row < 2 && column < 2).map(position => position.key), ['A1'])
  assert.deepEqual(sheetCheckboxCellPositions(geometry, viewport(), () => false), [])
  for (const zoom of [0, -1, NaN, Infinity]) assert.deepEqual(sheetCheckboxCellPositions(geometry, viewport({ zoom }), () => true), [])
})

test('notice anchors use a visible merged/frozen fragment and disappear outside the viewport', () => {
  const source = sheet(); source.merges = ['A1:B2']; source.freeze = { rows: 1, columns: 1 }
  const geometry = new SpreadsheetGeometry(source), position = sheetValidationCellVisibleRect(geometry, viewport(), 'B2')
  assert.ok(position.width > 0 && position.height > 0)
  assert.ok(position.x >= geometry.headerWidth && position.y >= geometry.headerHeight)
  assert.equal(sheetValidationCellVisibleRect(geometry, viewport(), 'T100'), undefined)
  assert.equal(sheetValidationCellVisibleRect(geometry, viewport(), 'invalid'), undefined)
})

test('a marker is a six-pixel zoomed corner triangle using the supplied theme color and clip', () => {
  const clip = { x: 50, y: 30, width: 100, height: 40 }
  assert.deepEqual(createSheetValidationInvalidMarker(bounds, clip, 1, '#aabbcc'), [{ kind: 'polygon', fill: '#aabbcc', clip,
    points: [{ x: 144, y: 30 }, { x: 150, y: 30 }, { x: 150, y: 36 }] }])
  const doubled = createSheetValidationInvalidMarker(bounds, clip, 2, '#111111')[0]
  assert.deepEqual(doubled.points, [{ x: 138, y: 30 }, { x: 150, y: 30 }, { x: 150, y: 42 }]); assert.equal(doubled.fill, '#111111')
  assert.deepEqual(createSheetValidationInvalidMarker(bounds, { x: 0, y: 0, width: 20, height: 20 }, 1, '#111111'), [])
  assert.deepEqual(createSheetValidationInvalidMarker(bounds, clip, 0, '#111111'), [])
})

test('the Canvas preserves selection and cell style while valid Ant checkboxes suppress only their scalar text', () => {
  const source = sheet(), before = JSON.stringify(source), geometry = new SpreadsheetGeometry(source)
  const scene = createSpreadsheetScene(geometry, viewport(), { activeCell: 'A1', selectedCells: ['A1'], selectedRanges: [selection],
    checkboxRenderable: (row, column) => row === 0 && column === 0, validationInvalid: (row, column) => row === 1 && column === 0,
    validationInvalidColor: '#bb3344' })
  assert.equal(scene.commands.some(command => command.kind === 'text' && command.text === '0'), false)
  assert.ok(scene.commands.some(command => command.kind === 'text' && command.text === 'invalid checkbox text'))
  assert.ok(scene.commands.some(command => command.kind === 'polygon' && command.fill === '#bb3344'))
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#246bfd'))
  assert.equal(JSON.stringify(source), before)
})

test('View-denied cells expose neither scalar content nor invalid markers in either backend command stream', () => {
  const geometry = new SpreadsheetGeometry(sheet()), scene = createSpreadsheetScene(geometry, viewport(), { activeCell: 'A1', selectedCells: [],
    selectedRanges: [selection], validationInvalid: () => true, validationInvalidColor: '#bb3344', canViewCell: () => false })
  assert.equal(scene.commands.some(command => command.kind === 'polygon' && command.fill === '#bb3344'), false)
  assert.equal(scene.commands.some(command => command.kind === 'text' && command.text === 'invalid checkbox text'), false)
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#246bfd'))
})
