import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetBorderLines, sheetBorderDashPattern } from '../../src/rendering/spreadsheet/borders.ts'
import { createSpreadsheetScene } from '../../src/rendering/spreadsheet/scene.ts'
import { SpreadsheetGeometry } from '../../src/rendering/spreadsheet/layout.ts'

const round = value => Math.round(value * 1e6) / 1e6
const viewport = { width: 420, height: 230, scrollLeft: 0, scrollTop: 0, zoom: 1 }
const pen = { style: 'thin', color: '#123456' }
const geometry = (cells, extra = {}) => new SpreadsheetGeometry({ id: 'merged-borders', name: 'Borders', cells,
  rowCount: 20, columnCount: 8, defaultRowHeight: 24, defaultColumnWidth: 80, ...extra })
const scene = (grid, view = viewport) => createSpreadsheetScene(grid, view, { activeCell: 'A5', selectedCells: [] })
const colored = (result, color = '#123456') => result.commands.filter(command => command.kind === 'line' && command.color === color)

test('dash clipping retains its original phase even when the visible side starts inside a distant cycle', () => {
  const lines = createSheetBorderLines(0, 0, 1000, 0, 'dashed', '#123456', 1, { x: 101, y: -2, width: 18, height: 4 })
  assert.deepEqual(lines.map(line => [round(line.x1), round(line.x2)]), [[102, 106], [108, 112], [114, 118]])
  const reverse = createSheetBorderLines(1000, 0, 0, 0, 'dashed', '#123456', 1, { x: 881, y: -2, width: 18, height: 4 })
  assert.deepEqual(reverse.map(line => [round(line.x1), round(line.x2)]), [[898, 894], [892, 888], [886, 882]])
  assert.deepEqual(sheetBorderDashPattern('slant-dash-dot'), [4, 2, 1, 2])
})

test('dash geometry is bounded by the visible interval for long horizontal and diagonal borders', () => {
  const clip = { x: 10_000_001, y: -2, width: 300, height: 4 }
  const lines = createSheetBorderLines(0, 0, 30_000_000, 0, 'dotted', '#123456', 1, clip)
  assert.ok(lines.length > 90 && lines.length < 110)
  assert.ok(lines.every(line => line.x1 >= clip.x - .5 && line.x2 <= clip.x + clip.width + .5))
  const diagonal = createSheetBorderLines(0, 0, 30_000_000, 30_000_000, 'dash-dot', '#123456', 1,
    { x: 20_000_000, y: 20_000_000, width: 300, height: 300 })
  assert.ok(diagonal.length > 50 && diagonal.length < 120)
  assert.deepEqual(createSheetBorderLines(0, 0, 1000, 0, 'thin', '#123456', 1, { x: 100, y: 5, width: 300, height: 10 }), [])
})

test('double and hair pens use zoomed width and perpendicular offsets on shared backend commands', () => {
  const clip = { x: -10, y: -10, width: 100, height: 100 }
  const double = createSheetBorderLines(0, 0, 40, 40, 'double', '#123456', 2, clip)
  assert.equal(double.length, 2)
  assert.ok(Math.abs(double[0].x1 + double[0].y1) < 1e-12)
  assert.ok(Math.abs(Math.hypot(double[0].x1, double[0].y1) - 1.6) < 1e-12)
  assert.equal(double[0].width, 2)
  assert.equal(createSheetBorderLines(0, 0, 40, 0, 'hair', '#123456', 2, clip)[0].width, 1)
  assert.deepEqual(createSheetBorderLines(0, 0, 40, 0, 'none', '#123456', 1, clip), [])
  assert.deepEqual(createSheetBorderLines(0, 0, 40, 0, 'thin', '#123456', 1, { ...clip, width: 0 }), [])
  assert.deepEqual(createSheetBorderLines(NaN, 0, 40, 0, 'thin', '#123456', 1, clip), [])
})

test('an explicitly cleared physical segment interrupts the imported main-cell merged side', () => {
  const result = scene(geometry({ B1: { value: '', border: { left: pen } }, B2: { value: '', border: { left: null } } }, { merges: ['B1:B3'] }))
  assert.deepEqual(colored(result).map(line => [line.x1, line.y1, line.x2, line.y2]), [
    [124, 26, 124, 50], [124, 74, 124, 98],
  ])
})

test('a merged non-primary edge is painted even when the main cell has no border', () => {
  const result = scene(geometry({ B3: { value: '', border: { bottom: pen } } }, { merges: ['B1:B3'] }))
  assert.deepEqual(colored(result).map(line => [line.x1, line.y1, line.x2, line.y2]), [[124, 98, 204, 98]])
})

test('pane backgrounds and selection fills are flushed before explicit cell borders', () => {
  const result = scene(geometry({ A1: { value: '', border: { right: pen } }, B1: { value: '', background: '#ff0000' } }))
  const edge = result.commands.findIndex(command => command.kind === 'line' && command.color === pen.color)
  const background = result.commands.findIndex(command => command.kind === 'rect' && command.fill === '#ff0000')
  assert.ok(edge > background)
})

test('a million-row merged border produces only current pane segments, including a scrolled main-cell fallback', () => {
  const grid = geometry({ B1: { value: '', border: { left: { ...pen, style: 'dashed' }, diagonalDown: { ...pen, style: 'dotted' } } } },
    { rowCount: 1_048_576, columnCount: 2, merges: ['B1:B1048576'] })
  for (const scrollTop of [0, 20_000_000]) {
    const lines = colored(scene(grid, { ...viewport, scrollTop }))
    assert.ok(lines.length > 10 && lines.length < 250)
    assert.ok(lines.every(line => line.clip.x >= 44 && line.clip.y >= 26))
  }
})
