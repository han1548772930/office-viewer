import type { SlicerModel, WorkbookModel } from '../../domain/model'
import { cloneModel } from '../../domain/serialization.ts'
import { drawingAffineBounds, drawingAffineFromFrame, multiplyDrawingAffine, normalizeDrawingAffineGeometry,
  type DrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, type SpreadsheetGeometry } from '../../rendering/spreadsheet/layout'
import { sheetClipboardObjectAxes } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'

export type SlicerDrawing = NonNullable<SlicerModel['drawing']>
export type SlicerAnchorType = SlicerDrawing['anchorType']
export type SlicerBounds = NonNullable<SlicerModel['layout']>
export type SlicerArrange = 'forward' | 'backward' | 'front' | 'back'
export type SlicerAxes = Pick<SpreadsheetGeometry, 'rows' | 'columns'>

/** Layout coordinates retain the workbook's existing header-inclusive CSS-pixel convention. */
export function captureSlicerDrawing(workbook: WorkbookModel, slicer: SlicerModel, bounds: SlicerBounds, anchorType: SlicerAnchorType): SlicerDrawing {
  const sheet = workbook.sheets.find((candidate) => candidate.id === slicer.hostSheetId)
  if (!sheet) throw new Error(`Slicer host worksheet not found: ${slicer.hostSheetId}`)
  const axes = sheetClipboardObjectAxes(workbook, sheet)
  return {
    anchorType,
    from: pointAnchor(axes, bounds.left - GRID_HEADER_WIDTH, bounds.top - GRID_HEADER_HEIGHT),
    to: pointAnchor(axes, bounds.left + bounds.width - GRID_HEADER_WIDTH, bounds.top + bounds.height - GRID_HEADER_HEIGHT),
    keepRatio: slicer.drawing?.keepRatio ?? false,
    zIndex: slicer.drawing?.zIndex ?? nextSlicerZIndex(workbook, slicer.hostSheetId),
  }
}

/** Cell markers describe the world AABB, independently of the authored affine basis. */
export function resolveSlicerAnchorBounds(model: SlicerModel, geometry: SlicerAxes): SlicerBounds {
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
    width: Math.max(0, GRID_HEADER_WIDTH + geometry.columns.offset(drawing.to.column) + drawing.to.offsetX - left),
    height: Math.max(0, GRID_HEADER_HEIGHT + geometry.rows.offset(drawing.to.row) + drawing.to.offsetY - top),
  }
}

/** A collapsed Both placement is unavailable, but its reversible author geometry is retained. */
export function resolveSlicerBounds(model: SlicerModel, axes: SlicerAxes): SlicerBounds {
  const anchored = resolveSlicerAnchorBounds(model, axes)
  if (!model.drawingGeometry || anchored.width <= 0 || anchored.height <= 0) return anchored
  const bounds = drawingAffineBounds(normalizeDrawingAffineGeometry(model.drawingGeometry))
  return { ...bounds, left: bounds.left + GRID_HEADER_WIDTH, top: bounds.top + GRID_HEADER_HEIGHT }
}

export function slicerDrawingGeometry(model: SlicerModel, axes: SlicerAxes): DrawingAffineGeometry {
  if (model.drawingGeometry) return normalizeDrawingAffineGeometry(model.drawingGeometry)
  const bounds = resolveSlicerAnchorBounds(model, axes)
  return drawingAffineFromFrame({ ...bounds, left: bounds.left - GRID_HEADER_WIDTH, top: bounds.top - GRID_HEADER_HEIGHT })
}

/** Change a world AABB without dropping rotation, reflection, shear or the local control extent. */
export function slicerGeometryWithBounds(geometry: DrawingAffineGeometry, bounds: SlicerBounds): DrawingAffineGeometry {
  const valid = normalizeDrawingAffineGeometry(geometry), before = drawingAffineBounds(valid)
  if (![bounds.left, bounds.top, bounds.width, bounds.height].every(Number.isFinite) || bounds.width <= 0 || bounds.height <= 0) {
    throw new RangeError('切片器位置或尺寸无效。')
  }
  const a = bounds.width / before.width, d = bounds.height / before.height
  return normalizeDrawingAffineGeometry({ width: valid.width, height: valid.height,
    matrix: multiplyDrawingAffine({ a, b: 0, c: 0, d,
      e: bounds.left - GRID_HEADER_WIDTH - a * before.left, f: bounds.top - GRID_HEADER_HEIGHT - d * before.top }, valid.matrix) })
}

/** Stage a native placement; callers own revision/history and publish only after the complete transaction validates. */
export function slicerWithDrawingGeometry(workbook: WorkbookModel, model: SlicerModel, geometry: DrawingAffineGeometry): SlicerModel {
  const valid = normalizeDrawingAffineGeometry(geometry), bounds = drawingAffineBounds(valid), next = cloneModel(model)
  next.drawingGeometry = valid
  next.layout = { ...bounds, left: bounds.left + GRID_HEADER_WIDTH, top: bounds.top + GRID_HEADER_HEIGHT }
  next.drawing = captureSlicerDrawing(workbook, next, next.layout, next.drawing?.anchorType ?? 'position')
  return next
}

export function slicerBoundsInWorkbook(workbook: WorkbookModel, model: SlicerModel): SlicerBounds {
  const sheet = workbook.sheets.find((candidate) => candidate.id === model.hostSheetId)
  if (!sheet) throw new Error(`Slicer host worksheet not found: ${model.hostSheetId}`)
  return resolveSlicerBounds(model, sheetClipboardObjectAxes(workbook, sheet))
}

export function nextSlicerZIndex(workbook: WorkbookModel, hostSheetId: string): number {
  return Math.max(-1, ...(workbook.slicers ?? []).filter((control) => control.hostSheetId === hostSheetId).map((control, index) => control.drawing?.zIndex ?? index)) + 1
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

