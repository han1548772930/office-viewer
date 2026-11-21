import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { toExcelWorkbookView, applyExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createSparklineCommands, parseSparklineRange, readSparklineValues } from '../src/features/sparkline/sparkline-render.ts'
import { createSparklineGroupPatches, createSparklineSpec, translateSparklineForMove } from '../src/features/sparkline/sparkline-model.ts'
import { createSpreadsheetScene } from '../src/rendering/spreadsheet/scene.ts'
import { SpreadsheetGeometry } from '../src/rendering/spreadsheet/layout.ts'

const lineSpec = {
  type: 'line', sourceRange: 'B1:E1', color: '#167a55', negativeColor: '#c43b3b',
  emptyCells: 'gap', axis: 'individual',
}

function renderSheet() {
  return {
    id: 'sheet-1', name: 'Sheet1', rowCount: 3, columnCount: 6, showGridlines: true,
    cells: {
      A1: { value: '', sparkline: lineSpec },
      B1: { value: '2', rawValue: 2 }, C1: { value: '', rawValue: null },
      D1: { value: '-2', rawValue: -2 }, E1: { value: '4', rawValue: 4 },
    },
  }
}

test('sparkline range parsing and live source values handle blanks and formula results', () => {
  assert.deepEqual(parseSparklineRange('$B$2:D4'), { startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 })
  assert.equal(parseSparklineRange('A0:B2'), undefined)
  const sheet = renderSheet()
  assert.deepEqual(readSparklineValues(lineSpec, sheet), [2, null, -2, 4])
  sheet.cells.E1.rawValue = 6
  assert.deepEqual(readSparklineValues(lineSpec, sheet), [2, null, -2, 6])
})

test('sparkline drawing supports line gaps, connected blanks, custom range, column colors and win-loss', () => {
  const bounds = { x: 10, y: 20, width: 80, height: 24 }
  const gap = createSparklineCommands(lineSpec, [2, null, -2, 4], bounds)
  assert.equal(gap.filter(command => command.kind === 'line' && command.color === '#167a55').length, 1)
  const connected = createSparklineCommands({ ...lineSpec, emptyCells: 'connect' }, [2, null, -2, 4], bounds)
  assert.equal(connected.filter(command => command.kind === 'line').length, 2)
  const connectedLines = connected.filter(command => command.kind === 'line')
  assert.ok(Math.abs(connectedLines[0].x2 - (bounds.x + 3 + (bounds.width - 6) * 2 / 3)) < 0.001)
  const columns = createSparklineCommands({ ...lineSpec, type: 'column', emptyCells: 'zero' }, [2, -2], bounds)
  assert.deepEqual(columns.filter(command => command.kind === 'rect').map(command => command.fill), ['#167a55', '#c43b3b'])
  const winloss = createSparklineCommands({ ...lineSpec, type: 'winloss' }, [5, -1, 0], bounds)
  assert.equal(winloss.filter(command => command.kind === 'rect').length, 2)
  assert.equal(createSparklineCommands({ ...lineSpec, axis: 'custom', min: 4, max: -4 }, [2], bounds).length, 0)
})

test('sparkline configuration survives workbook view conversion and renders in the Canvas scene', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.cells['0:0'] = {
    id: '0:0', row: 0, col: 0, value: null, formula: null, display: '',
    style: { font: { family: 'Aptos', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#111' }, fill: '#fff', horizontal: 'left', vertical: 'middle', wrap: false, numberFormat: 'General' },
    sparkline: lineSpec,
  }
  const view = toExcelWorkbookView(workbook)
  const restored = applyExcelWorkbookView(workbook, view)
  assert.deepEqual(restored.sheets[0].cells['0:0'].sparkline, lineSpec)

  const renderSheetView = renderSheet()
  const scene = createSpreadsheetScene(new SpreadsheetGeometry({ ...renderSheetView, merges: [], freeze: { rows: 0, columns: 0 } }),
    { width: 500, height: 180, scrollLeft: 0, scrollTop: 0, zoom: 1 }, { activeCell: 'F3', selectedCells: [] })
  assert.ok(scene.commands.some(command => command.kind === 'line' && command.color === '#167a55'))
})

test('sparkline group patches validate target series and move source references', () => {
  const request = { targetRange: 'F1:F2', sourceRange: 'B1:D2', type: 'column', color: '#167a55', negativeColor: '#c43b3b', axis: 'custom', min: -10, max: 10 }
  const patches = createSparklineGroupPatches(request)
  assert.deepEqual(Object.keys(patches), ['F1', 'F2'])
  assert.equal(patches.F1.sourceRange, 'B1:D1')
  assert.equal(patches.F2.sourceRange, 'B2:D2')
  assert.throws(() => createSparklineGroupPatches({ ...request, targetRange: 'F1:G2' }), /单行或单列/)
  const spec = createSparklineSpec({ ...request, targetRange: 'F1' })
  assert.equal(translateSparklineForMove(spec, 1, 2).sourceRange, 'D2:F3')
  assert.throws(() => translateSparklineForMove(spec, -1, 0), /工作表边界/)
})
