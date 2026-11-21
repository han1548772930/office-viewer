import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'vitest'
import { crosshairHighlightRanges } from '../src/features/sheet-crosshair/crosshair-highlight.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'

const selected = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }

test('crosshair ranges cover the selected row and column outside the selected cells', () => {
  const ranges = crosshairHighlightRanges([selected], 4, 4)
  assert.equal(ranges.length, 4)
  for (const range of ranges) {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        assert.ok(row === selected.startRow || column === selected.startColumn)
        assert.ok(!(row === selected.startRow && column === selected.startColumn))
      }
    }
  }
})

test('crosshair combines multiple cell ranges and excludes whole-row or whole-column selections', () => {
  const ranges = crosshairHighlightRanges([
    selected,
    { startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 },
    { startRow: 0, endRow: 3, startColumn: 3, endColumn: 3 },
  ], 4, 4)
  for (const range of ranges) {
    assert.ok(!(range.startRow === 0 && range.endRow === 0 && range.startColumn === 0 && range.endColumn === 3))
    assert.ok(!(range.startColumn === 3 && range.endColumn === 3 && range.startRow === 0 && range.endRow === 3))
  }
  assert.deepEqual(crosshairHighlightRanges([{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 }], 4, 4), [])
  assert.deepEqual(crosshairHighlightRanges([{ startRow: 0, endRow: 3, startColumn: 2, endColumn: 2 }], 4, 4), [])
})

test('crosshair is painted as a view overlay while the selection outline remains visible', () => {
  const geometry = new SpreadsheetGeometry({ id: 'crosshair', name: 'Crosshair', cells: {}, rowCount: 4, columnCount: 4 })
  const scene = createSpreadsheetScene(
    geometry,
    { width: 500, height: 300, scrollLeft: 0, scrollTop: 0, zoom: 1 },
    { activeCell: 'B2', selectedCells: [], selectedRanges: [selected], crosshairHighlight: { enabled: true, color: 'rgba(91, 141, 239, 0.18)' } },
  )
  assert.ok(scene.commands.some(command => command.kind === 'rect' && command.fill === 'rgba(91, 141, 239, 0.18)'))
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#246bfd'))
  assert.ok(!scene.commands.some(command => command.kind === 'rect' && command.fill === 'rgba(91, 141, 239, 0.18)' && command.x === geometry.cellRect(1, 1, { width: 500, height: 300, scrollLeft: 0, scrollTop: 0, zoom: 1 }).x))
})

test('crosshair setting and color picker remain operations rather than workbook content mutations', () => {
  const ribbon = readFileSync(new URL('../src/components/spreadsheet/sheet-ribbon.ts', import.meta.url), 'utf8')
  const workbench = readFileSync(new URL('../src/components/ExcelWorkbench.vue', import.meta.url), 'utf8')
  assert.match(ribbon, /sheet\.operation\.toggle-crosshair-highlight/)
  assert.match(ribbon, /sheet\.operation\.set-crosshair-highlight-color/)
  assert.match(workbench, /crosshairEnabled\.value = !crosshairEnabled\.value/)
  assert.match(workbench, /crosshairColor\.value = params/)
  assert.match(workbench, /:crosshair-highlight="\{ enabled: crosshairEnabled, color: crosshairColor \}"/)
})
