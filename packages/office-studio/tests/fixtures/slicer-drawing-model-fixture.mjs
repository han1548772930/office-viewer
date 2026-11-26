import { createBlankSheet, createCell } from '../../src/domain/model.ts'
import { initializeWorkbookDrawingStates } from '../../src/features/sheet-drawings/sheet-drawing-model-lifecycle.ts'
import { createPivotTable, refreshPivotTable, setPivotField } from '../../src/features/pivot/services/pivot-engine.ts'
import { createSlicer, getSlicerItems, setSlicerItems, setSlicerTimelineRange, updateSlicerPresentation } from '../../src/features/slicer/slicer-engine.ts'
import { slicerWithDrawingGeometry } from '../../src/features/slicer/slicer-drawing.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../src/rendering/spreadsheet/layout.ts'
import { sheetDrawingWorkbook } from './sheet-drawing-model-fixture.mjs'

// Deferred native model fixture: a real Table slicer and a real date-field Pivot timeline.
// HTTP media from the base fixture is not decoded and is not an Office/image acceptance sample.
export const slicerDrawingRefs = [
  { kind: 'image', id: 'a' }, { kind: 'shape', id: 'b' }, { kind: 'chart', id: 'c' },
  { kind: 'slicer', id: 's' }, { kind: 'timeline', id: 't' },
]
export function slicerDrawingWorkbook({ id = 'slicer-drawing-book', affine = true, filtered = false, scrollable = false } = {}) {
  const workbook = sheetDrawingWorkbook({ id, slicer: true }), sheet = workbook.sheets[0]
  const slicer = workbook.slicers.find(control => control.id === 's')
  if (scrollable) {
    const table = workbook.tables.find(table => table.id === slicer.tableId)
    table.range.endRow = 48; sheet.rowCount = 200
    for (let row = 3; row <= table.range.endRow; row++) {
      sheet.cells[`${row}:0`] = createCell(row, 0, `Region ${String(row).padStart(2, '0')}`)
      sheet.cells[`${row}:1`] = createCell(row, 1, row * 10)
    }
  }
  updateSlicerPresentation(slicer, { layout: { left: GRID_HEADER_WIDTH + 150, top: GRID_HEADER_HEIGHT + 60, width: 180, height: 160 } }, workbook)
  const dates = createBlankSheet('Dates'); dates.id = 'dates'; dates.rowCount = 50; dates.columnCount = 10
  dates.defaultRowHeight = 20; dates.defaultColumnWidth = 100; dates.defaultRowAutoHeight = false
  const dateRows = scrollable ? [['Date', 'Amount'], ...Array.from({ length: 36 }, (_, index) =>
    [`${2024 + Math.floor(index / 12)}-${String(index % 12 + 1).padStart(2, '0')}-15`, (index + 1) * 10])]
    : [['Date', 'Amount'], ['2026-01-15', 10], ['2026-02-03', 20], ['2026-03-02', 30]]
  for (const [row, values] of dateRows.entries()) {
    values.forEach((value, col) => { dates.cells[row + ':' + col] = createCell(row, col, value) })
  }
  workbook.sheets.push(dates)
  const pivot = createPivotTable(workbook, { unitId: id, subUnitId: 'dates', sheetName: 'Dates',
    range: { startRow: 0, startColumn: 0, endRow: dateRows.length - 1, endColumn: 1 } },
  { unitId: id, subUnitId: sheet.id, row: scrollable ? 70 : 40, col: 10 }, 'date-pivot')
  setPivotField(workbook, pivot, '0', 'row'); setPivotField(workbook, pivot, '1', 'value')
  if (scrollable) refreshPivotTable(workbook, pivot)
  const timeline = createSlicer(workbook, { id: 't', unitId: id, hostSheetId: sheet.id, type: 'timeline',
    pivotTableId: pivot.id, sourceFieldId: '0', layout: { left: GRID_HEADER_WIDTH + 450, top: GRID_HEADER_HEIGHT + 120, width: 300, height: 130 } })
  if (affine) {
    Object.assign(slicer, slicerWithDrawingGeometry(workbook, slicer,
      { width: 180, height: 160, matrix: { a: 1, b: 0.25, c: 0.3, d: 1, e: 150, f: 60 } }))
    Object.assign(timeline, slicerWithDrawingGeometry(workbook, timeline,
      { width: 300, height: 130, matrix: { a: 1, b: -0.25, c: 0.2, d: 1, e: 450, f: 200 } }))
  }
  if (filtered) {
    setSlicerItems(workbook, slicer, ['North'])
    setSlicerTimelineRange(workbook, timeline, { start: '2026-02-01', end: '2026-02-28', level: 'month' })
  }
  initializeWorkbookDrawingStates(workbook)
  return workbook
}
export function slicerDrawingControls(workbook) {
  return (workbook.slicers ?? []).map(model => ({ model, items: getSlicerItems(workbook, model) }))
}
export function slicerDrawingAuthor(model) {
  const { position, size, layout, drawing, drawingGeometry, revision, ...author } = structuredClone(model)
  return { ...author, size, anchorType: drawing?.anchorType, keepRatio: drawing?.keepRatio, zIndex: drawing?.zIndex }
}
