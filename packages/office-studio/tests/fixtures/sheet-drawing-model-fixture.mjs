import { createBlankWorkbook, createCell } from '../../src/domain/model.ts'
import { drawingAffineFromFrame } from '../../src/domain/drawing-affine.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../src/rendering/spreadsheet/layout.ts'
import { createChartDefinition } from '../../src/features/charts/chart-engine.ts'
import { normalizeSheetFloatingImage } from '../../src/features/sheet-images/sheet-image-model.ts'
import { createSheetTable } from '../../src/features/sheet-tables/sheet-tables.ts'
import { createSlicer } from '../../src/features/slicer/slicer-engine.ts'

// A native model fixture only. Its HTTP media is never fetched/decoded and is not an accepted Office/image sample.
export const drawingRef = (kind, id) => ({ kind, id })
export const drawingFrame = (left, top, width, height) => drawingAffineFromFrame({ left, top, width, height })
export function sheetDrawingWorkbook({ id = 'drawing-book', slicer = false } = {}) {
  let workbook = createBlankWorkbook()
  workbook.id = id; workbook.calculationMode = 'manual'
  const sheet = workbook.sheets[0]
  sheet.id = 'data'; workbook.activeSheetId = sheet.id; sheet.rowCount = 100; sheet.columnCount = 20
  sheet.defaultRowHeight = 20; sheet.defaultColumnWidth = 100; sheet.defaultRowAutoHeight = false
  for (const [row, values] of [['Name', 'Value'], ['North', 7], ['South', 4]].entries()) {
    values.forEach((value, col) => { sheet.cells[row + ':' + col] = createCell(row, col, value) })
  }
  const resource = { id: 'shared-media', name: 'drawing.png', mimeType: 'image/png', source: 'https://example.invalid/drawing.png',
    width: 20, height: 10, byteLength: 8, metadata: { owner: 'fixture' } }
  sheet.images = [
    normalizeSheetFloatingImage({ id: 'a', resource, anchor: { row: 0, column: 0, offsetX: 10, offsetY: 20 },
      anchorType: '2', absolutePosition: { left: 10, top: 20 }, size: { width: 20, height: 10 },
      keepRatio: true, opacity: 0.8, title: 'Image', altText: 'Keep original', crop: { top: 1, right: 2, bottom: 3, left: 4 }, flipX: true, flipY: false }),
    normalizeSheetFloatingImage({ id: 'd', resource, anchor: { row: 1, column: 0, offsetX: 90, offsetY: 0 },
      anchorType: '2', absolutePosition: { left: 90, top: 20 }, size: { width: 20, height: 10 },
      keepRatio: false, opacity: 1, title: 'Other owner', altText: '' }),
  ]
  sheet.shapes = [{ id: 'b', sheetId: sheet.id, geometry: drawingFrame(30, 20, 20, 10), source: { kind: 'preset', preset: 'rect' },
    style: { fill: '#00aa66', line: '#003322' }, text: { content: 'Keep shape text', style: { fontSize: 12 } }, metadata: { imported: true } }]
  workbook.charts = [createChartDefinition(workbook, sheet.id, { type: 'line',
    source: { sheetId: sheet.id, orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 } },
    position: { x: GRID_HEADER_WIDTH + 50, y: GRID_HEADER_HEIGHT + 40 }, size: { width: 30, height: 20 },
    valueFields: [1], title: 'Keep chart title', options: { titleStyle: { rotate: 30 } },
  }, 'c')]
  if (slicer) {
    workbook = createSheetTable(workbook, { id: 'table-source', sheetId: sheet.id, name: 'DrawingTable', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 } })
    createSlicer(workbook, { id: 's', unitId: workbook.id, hostSheetId: sheet.id, type: 'slicer', tableId: 'table-source',
      sourceFieldId: workbook.tables[0].columns[0].id, layout: { left: GRID_HEADER_WIDTH + 120, top: GRID_HEADER_HEIGHT + 80, width: 80, height: 40 } })
  }
  return workbook
}
