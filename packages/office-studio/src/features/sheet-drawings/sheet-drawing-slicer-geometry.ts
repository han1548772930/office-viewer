import type { SlicerModel, WorkbookModel } from '../../domain/model.ts'
import { drawingAffineBounds } from '../../domain/drawing-affine.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH } from '../../rendering/spreadsheet/layout.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { sheetClipboardObjectAxes } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { resolveSlicerAnchorBounds, slicerGeometryWithBounds } from '../slicer/slicer-drawing.ts'
import { sheetDrawingAxisGeometryInput } from './sheet-drawing-axis-geometry.ts'

function placement(model: SlicerModel) {
  return { position: model.position, size: model.size, layout: model.layout, drawing: model.drawing }
}

/** Final projection after structure, filtering and auto-height; explicit new matrices are already authoritative. */
export function reconcileWorkbookSlicerGeometryChanges(before: WorkbookModel, after: WorkbookModel): void {
  if (!after.slicers?.some(model => model.drawingGeometry && model.drawing?.anchorType !== 'none')) return
  const originals = new Map((before.slicers ?? []).map(model => [model.id, model]))
  for (const sheet of after.sheets) {
    const previous = before.sheets.find(candidate => candidate.id === sheet.id)
    if (!previous) continue
    const axesChanged = !equivalentModelData(sheetDrawingAxisGeometryInput(before, previous), sheetDrawingAxisGeometryInput(after, sheet))
    let axes: ReturnType<typeof sheetClipboardObjectAxes> | undefined
    for (const model of after.slicers) {
      const original = originals.get(model.id)
      if (model.hostSheetId !== sheet.id || !model.drawingGeometry || model.drawing?.anchorType === 'none'
        || !original?.drawingGeometry || original.hostSheetId !== sheet.id
        || !equivalentModelData(original.drawingGeometry, model.drawingGeometry)
        || !axesChanged && equivalentModelData(placement(original), placement(model))) continue
      axes ??= sheetClipboardObjectAxes(after, sheet)
      const bounds = resolveSlicerAnchorBounds(model, axes), geometry = model.drawingGeometry
      if (bounds.width <= 0 || bounds.height <= 0) {
        // Collapsed axes do not create a singular matrix or overwrite a positive authored extent.
        const retained = drawingAffineBounds(geometry)
        model.layout = { ...retained, left: retained.left + GRID_HEADER_WIDTH, top: retained.top + GRID_HEADER_HEIGHT }
        continue
      }
      const projected = slicerGeometryWithBounds(geometry, bounds)
      if (equivalentModelData(projected, geometry) && equivalentModelData(model.layout, bounds)) continue
      model.drawingGeometry = projected
      model.layout = { ...bounds }
      model.revision++
    }
  }
}
