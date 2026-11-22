import assert from 'node:assert/strict'
import { test } from 'vitest'
import { applySheetBorders } from '../src/features/sheet-workbench/sheet-toolbar-format.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'

test('border options preserve perimeter behavior and apply the selected style and color', () => {
  const cells = { A1: { value: 'keep', border: { bottom: { style: 'dashed', color: '#123456' } } } }
  const ranges = [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }]
  applySheetBorders(cells, ranges, 'outside', { style: 'medium-dash-dot', color: '#cc3300' })
  assert.deepEqual(cells.A1.border, {
    bottom: { style: 'dashed', color: '#123456' },
    top: { style: 'medium-dash-dot', color: '#cc3300' },
    left: { style: 'medium-dash-dot', color: '#cc3300' },
  })
  assert.deepEqual(cells.B2.border, {
    bottom: { style: 'medium-dash-dot', color: '#cc3300' },
    right: { style: 'medium-dash-dot', color: '#cc3300' },
  })
  assert.equal(cells.A1.value, 'keep')
})

test('inside, horizontal, vertical and diagonal border operations use their own positions', () => {
  const range = [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }]
  const cells = {}
  applySheetBorders(cells, range, 'inside')
  assert.deepEqual(cells.A1.border, { right: { style: 'thin', color: '#1f2937' }, bottom: { style: 'thin', color: '#1f2937' } })
  applySheetBorders(cells, range, 'horizontal')
  assert.equal(cells.A1.border.top, undefined)
  assert.ok(cells.A1.border.bottom)
  assert.equal(cells.A1.border.left, undefined)
  applySheetBorders(cells, range, 'vertical')
  assert.equal(cells.A1.border.left, undefined)
  assert.ok(cells.A1.border.right)
  applySheetBorders(cells, [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }], 'diagonal-down', { style: 'double', color: '#008844' })
  assert.deepEqual(cells.A1.border.diagonalDown, { style: 'double', color: '#008844' })
  applySheetBorders(cells, range, 'diagonal-down-triple', { style: 'thin', color: '#ee4111' })
  for (const key of ['topLeftMiddleRight', 'topLeftBottomCenter', 'diagonalDown']) assert.deepEqual(cells.B2.border[key], { style: 'thin', color: '#ee4111' })
  applySheetBorders(cells, range, 'none')
  assert.deepEqual(cells.A1.border, {})
})

test('Canvas scene expands dashed and double borders and renders diagonal borders', () => {
  const geometry = new SpreadsheetGeometry({
    id: 's', name: 'Borders', rowCount: 2, columnCount: 2,
    cells: { A1: { value: '', border: {
      bottom: { style: 'dotted', color: '#123456' },
      diagonalDown: { style: 'double', color: '#cc3300' },
    } } },
  })
  const scene = createSpreadsheetScene(geometry, { width: 500, height: 200, scrollLeft: 0, scrollTop: 0, zoom: 1 }, { activeCell: 'B2', selectedCells: [] })
  const coloredLines = scene.commands.filter(command => command.kind === 'line' && ['#123456', '#cc3300'].includes(command.color))
  assert.ok(coloredLines.filter(command => command.color === '#123456').length > 1)
  assert.equal(coloredLines.filter(command => command.color === '#cc3300').length, 2)
  assert.ok(coloredLines.some(command => command.color === '#cc3300' && command.width === 1))
})
