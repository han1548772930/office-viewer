import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createSheetRibbon } from '../src/components/spreadsheet/sheet-ribbon.ts'
import { ribbonInvocation } from '../src/platform/ribbon/ribbon-schema.ts'
import { createRibbonMenuModel, ribbonMenuInvocation } from '../src/platform/ribbon/ribbon-command-menu.ts'
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
  const viewport = { width: 500, height: 300, scrollLeft: 0, scrollTop: 0, zoom: 1 }
  const scene = createSpreadsheetScene(
    geometry,
    viewport,
    { activeCell: 'B2', selectedCells: [], selectedRanges: [selected], crosshairHighlight: { enabled: true, color: 'rgba(91, 141, 239, 0.18)' } },
  )
  const highlights = scene.commands.filter(command => command.kind === 'rect' && command.fill === 'rgba(91, 141, 239, 0.18)')
  const rect = geometry.cellRect(1, 1, viewport)
  const center = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }
  assert.ok(highlights.length > 0)
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#246bfd'))
  assert.ok(highlights.every(command => center.x < command.x || center.x >= command.x + command.width
    || center.y < command.y || center.y >= command.y + command.height))
})

test('crosshair Ribbon and color menu dispatch operations and reflect the current view state', () => {
  const color = 'rgba(141, 99, 201, 0.16)'
  const view = createSheetRibbon({
    canUndo: false, canRedo: false, hasSheet: true, canMerge: false, canUnmerge: false,
    canDeleteRow: true, canDeleteColumn: true, chartAvailable: true, dataValidationAvailable: true,
    canSort: true, fontFamily: 'Aptos', fontSize: 11, numberFormat: 'general',
    gridlines: true, frozen: false, zoom: 100, fullscreen: false, crosshairEnabled: true, crosshairColor: color,
  }).find(tab => tab.id === 'view')
  const item = view.groups[0].items.find(item => item.id === 'sheet.operation.toggle-crosshair-highlight')
  assert.equal(item.activated, true)
  assert.deepEqual(ribbonInvocation(item), { commandId: 'sheet.operation.toggle-crosshair-highlight', params: undefined })
  const menu = createRibbonMenuModel(item.children)
  assert.deepEqual(menu.selectedKeys, ['sheet.operation.crosshair-color.purple'])
  assert.deepEqual(ribbonMenuInvocation(menu, 'sheet.operation.crosshair-color.purple'), {
    commandId: 'sheet.operation.set-crosshair-highlight-color', params: color,
  })
})

test('crosshair clips each visible cell once across frozen panes, hidden axes, scrolling and zoom', () => {
  const sheet = { id: 'crosshair-panes', name: 'Crosshair panes', cells: {}, rowCount: 12, columnCount: 12,
    freeze: { rows: 2, columns: 2 }, hiddenRows: [3], hiddenColumns: [3], defaultRowHeight: 26, defaultColumnWidth: 80 }
  const before = structuredClone(sheet)
  const geometry = new SpreadsheetGeometry(sheet)
  const viewport = { width: 580, height: 340, scrollLeft: 135, scrollTop: 52, zoom: 1.25 }
  const selections = [selected, { startRow: 4, endRow: 4, startColumn: 4, endColumn: 4 }]
  const color = 'rgba(141, 99, 201, 0.16)'
  const scene = createSpreadsheetScene(geometry, viewport, {
    activeCell: 'E5', selectedCells: [], selectedRanges: selections, crosshairHighlight: { enabled: true, color },
  })
  const highlights = scene.commands.filter(command => command.kind === 'rect' && command.fill === color)
  const panes = geometry.panes(viewport)
  assert.ok(highlights.length > 0)
  for (const command of highlights) assert.ok(panes.some(pane => JSON.stringify(pane.clip) === JSON.stringify(command.clip)))
  for (const pane of panes) {
    for (const row of pane.rows) {
      for (const column of pane.columns) {
        const rect = geometry.cellRect(row, column, viewport, false)
        const x = rect.x + rect.width / 2, y = rect.y + rect.height / 2
        if (x < pane.clip.x || x >= pane.clip.x + pane.clip.width || y < pane.clip.y || y >= pane.clip.y + pane.clip.height) continue
        const selectedCell = selections.some(range => row === range.startRow && column === range.startColumn)
        const expected = !selectedCell && selections.some(range => row === range.startRow || column === range.startColumn)
        const painted = highlights.filter(command => x >= command.x && x < command.x + command.width
          && y >= command.y && y < command.y + command.height && x >= command.clip.x && x < command.clip.x + command.clip.width
          && y >= command.clip.y && y < command.clip.y + command.clip.height)
        assert.equal(painted.length, expected ? 1 : 0, `cell ${row}:${column}`)
      }
    }
  }
  assert.deepEqual(sheet, before)
  const disabled = createSpreadsheetScene(geometry, viewport, {
    activeCell: 'E5', selectedCells: [], selectedRanges: selections, crosshairHighlight: { enabled: false, color },
  })
  assert.ok(!disabled.commands.some(command => command.kind === 'rect' && command.fill === color))
})
