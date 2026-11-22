import type { SlicerModel, SpreadsheetSheet, WorkbookModel } from '../../domain/model'
import { AxisIndex, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, type SpreadsheetGeometry } from '../../rendering/spreadsheet/layout'
import { getCollapsedOutlineIndexes } from '../sheet-outline/sheet-outline'
import { effectiveSheetRowSizes, sheetAxisSpanAt } from '../../domain/sheet-axis.ts'

export type SlicerDrawing = NonNullable<SlicerModel['drawing']>
export type SlicerAnchorType = SlicerDrawing['anchorType']
export type SlicerBounds = NonNullable<SlicerModel['layout']>
export type SlicerArrange = 'forward' | 'backward' | 'front' | 'back'
interface SlicerAxes { rows: AxisIndex; columns: AxisIndex }

/** Layout coordinates retain the workbook's existing header-inclusive CSS-pixel convention. */
export function captureSlicerDrawing(workbook: WorkbookModel, slicer: SlicerModel, bounds: SlicerBounds, anchorType: SlicerAnchorType): SlicerDrawing {
  const sheet = workbook.sheets.find((candidate) => candidate.id === slicer.hostSheetId)
  if (!sheet) throw new Error(`Slicer host worksheet not found: ${slicer.hostSheetId}`)
  const axes = workbookAxes(workbook, sheet)
  return {
    anchorType,
    from: pointAnchor(axes, bounds.left - GRID_HEADER_WIDTH, bounds.top - GRID_HEADER_HEIGHT),
    to: pointAnchor(axes, bounds.left + bounds.width - GRID_HEADER_WIDTH, bounds.top + bounds.height - GRID_HEADER_HEIGHT),
    keepRatio: slicer.drawing?.keepRatio ?? false,
    zIndex: slicer.drawing?.zIndex ?? nextSlicerZIndex(workbook, slicer.hostSheetId),
  }
}

export function resolveSlicerBounds(model: SlicerModel, geometry: Pick<SpreadsheetGeometry, 'rows' | 'columns'>): SlicerBounds {
  const absolute = model.layout ?? {
    left: GRID_HEADER_WIDTH + geometry.columns.offset(model.position.column),
    top: GRID_HEADER_HEIGHT + geometry.rows.offset(model.position.row),
    width: geometry.columns.offset(model.position.column + model.size.columns) - geometry.columns.offset(model.position.column),
    height: geometry.rows.offset(model.position.row + model.size.rows) - geometry.rows.offset(model.position.row),
  }
  const drawing = model.drawing
  if (!drawing || drawing.anchorType === 'none') return { ...absolute }
  const left = GRID_HEADER_WIDTH + geometry.columns.offset(drawing.from.column) + drawing.from.offsetX
  const top = GRID_HEADER_HEIGHT + geometry.rows.offset(drawing.from.row) + drawing.from.offsetY
  if (drawing.anchorType === 'position') return { ...absolute, left, top }
  return {
    left, top,
    width: Math.max(1, GRID_HEADER_WIDTH + geometry.columns.offset(drawing.to.column) + drawing.to.offsetX - left),
    height: Math.max(1, GRID_HEADER_HEIGHT + geometry.rows.offset(drawing.to.row) + drawing.to.offsetY - top),
  }
}

export function slicerBoundsInWorkbook(workbook: WorkbookModel, model: SlicerModel): SlicerBounds {
  const sheet = workbook.sheets.find((candidate) => candidate.id === model.hostSheetId)
  if (!sheet) throw new Error(`Slicer host worksheet not found: ${model.hostSheetId}`)
  return resolveSlicerBounds(model, workbookAxes(workbook, sheet))
}

export function nextSlicerZIndex(workbook: WorkbookModel, hostSheetId: string): number {
  return Math.max(-1, ...(workbook.slicers ?? []).filter((control) => control.hostSheetId === hostSheetId).map((control, index) => control.drawing?.zIndex ?? index)) + 1
}

export function arrangeSlicer(workbook: WorkbookModel, slicer: SlicerModel, action: SlicerArrange): void {
  if (!['forward', 'backward', 'front', 'back'].includes(action)) throw new TypeError('Invalid slicer arrangement')
  const controls = (workbook.slicers ?? []).filter((control) => control.hostSheetId === slicer.hostSheetId)
    .map((control, index) => ({ control, order: control.drawing?.zIndex ?? index }))
    .sort((first, second) => first.order - second.order).map(({ control }) => control)
  const index = controls.findIndex((control) => control.id === slicer.id)
  const destination = action === 'front' ? controls.length - 1 : action === 'back' ? 0
    : Math.max(0, Math.min(controls.length - 1, index + (action === 'forward' ? 1 : -1)))
  if (index < 0 || index === destination) return
  controls.splice(index, 1)
  controls.splice(destination, 0, slicer)
  controls.forEach((control, order) => {
    if (!control.drawing) control.drawing = captureSlicerDrawing(workbook, control, slicerBoundsInWorkbook(workbook, control), 'position')
    if (control.drawing.zIndex !== order) { control.drawing.zIndex = order; control.revision += 1 }
  })
}

export function validateSlicerDrawing(drawing: SlicerDrawing): void {
  if (!['both', 'position', 'none'].includes(drawing.anchorType)) throw new TypeError('Invalid slicer anchor type')
  for (const point of [drawing.from, drawing.to]) {
    if (!point || !Number.isInteger(point.row) || point.row < 0 || !Number.isInteger(point.column) || point.column < 0
      || !Number.isFinite(point.offsetX) || !Number.isFinite(point.offsetY)) throw new RangeError('Invalid slicer drawing anchor')
  }
  if (drawing.keepRatio !== undefined && typeof drawing.keepRatio !== 'boolean') throw new TypeError('Slicer aspect ratio lock must be a boolean')
  if (drawing.zIndex !== undefined && !Number.isFinite(drawing.zIndex)) throw new RangeError('Slicer drawing order must be finite')
}

function pointAnchor(axes: SlicerAxes, x: number, y: number): SlicerDrawing['from'] {
  const column = axes.columns.locate(Math.max(0, x))
  const row = axes.rows.locate(Math.max(0, y))
  return { row, column, offsetX: x - axes.columns.offset(column), offsetY: y - axes.rows.offset(row) }
}

function workbookAxes(workbook: WorkbookModel, sheet: SpreadsheetSheet): SlicerAxes {
  const counts = Object.values(sheet.cells).reduce((size, cell) => ({ rows: Math.max(size.rows, cell.row + 1), columns: Math.max(size.columns, cell.col + 1) }), { rows: sheet.rowCount ?? 100, columns: sheet.columnCount ?? 26 })
  const rowCount = Math.max(100, counts.rows)
  const columnCount = Math.max(26, counts.columns)
  const outlineRows = sheet.outline ? [...getCollapsedOutlineIndexes({ count: rowCount, groups: sheet.outline.groups.filter((group) => group.axis === 'row') })] : []
  const outlineColumns = sheet.outline ? [...getCollapsedOutlineIndexes({ count: columnCount, groups: sheet.outline.groups.filter((group) => group.axis === 'column') })] : []
  const tableRows = (workbook.tables ?? []).filter((table) => table.sheetId === sheet.id).flatMap((table) => table.autoFilter?.filteredRows ?? [])
  const rows = effectiveSheetRowSizes(sheet)
  const columns = sheet.defaultColumnWidth === undefined && !sheetAxisSpanAt(sheet.columnWidthSpans, 0) ? { 0: 152, ...sheet.columnWidths } : sheet.columnWidths
  return {
    rows: new AxisIndex(rowCount, sheet.defaultRowHeight ?? 26, rows.sizes, [...(sheet.hiddenRows ?? []), ...(sheet.autoFilter?.filteredRows ?? []), ...tableRows, ...outlineRows], rows.spans, sheet.rowHiddenSpans),
    columns: new AxisIndex(columnCount, sheet.defaultColumnWidth ?? 112, columns, [...(sheet.hiddenColumns ?? []), ...outlineColumns], sheet.columnWidthSpans, sheet.columnHiddenSpans),
  }
}
