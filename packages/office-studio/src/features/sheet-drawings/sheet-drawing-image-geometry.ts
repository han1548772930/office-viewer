import type { WorkbookModel } from '../../domain/model.ts'
import type { SheetFloatingImage } from '../../domain/sheet-images.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { sheetClipboardObjectAxes } from '../sheet-workbench/sheet-clipboard-object-geometry.ts'
import { transformSheetImagesForGeometryChange } from '../sheet-images/sheet-image-structure.ts'
import { sheetDrawingAxisGeometryInput } from './sheet-drawing-axis-geometry.ts'

function placement(image: SheetFloatingImage) {
  return { anchor: image.anchor, endAnchor: image.endAnchor, anchorType: image.anchorType, absolutePosition: image.absolutePosition,
    size: image.size, drawingGeometry: image.drawingGeometry, rotation: image.rotation, crop: image.crop, flipX: image.flipX, flipY: image.flipY }
}

/** Run after filter and automatic row-height computation, before fitting parent groups. */
export function reconcileWorkbookImageGeometryChanges(before: WorkbookModel, after: WorkbookModel): void {
  for (const sheet of after.sheets) {
    const previous = before.sheets.find(candidate => candidate.id === sheet.id)
    if (!previous || !sheet.images?.some(image => image.drawingGeometry && image.anchorType !== '2')
      || equivalentModelData(sheetDrawingAxisGeometryInput(before, previous), sheetDrawingAxisGeometryInput(after, sheet))) continue
    const originals = new Map((previous.images ?? []).map(image => [image.id, image]))
    // Structure and explicit drawing commands already publish their final geometry. Only untouched placements need projection.
    const pending = sheet.images.filter(image => {
      const original = originals.get(image.id)
      return image.drawingGeometry && image.anchorType !== '2' && original && equivalentModelData(placement(original), placement(image))
    })
    if (!pending.length) continue
    const projected = new Map(transformSheetImagesForGeometryChange(pending, {
      before: sheetClipboardObjectAxes(before, previous), after: sheetClipboardObjectAxes(after, sheet),
    }).map(image => [image.id, image]))
    sheet.images = sheet.images.map(image => projected.get(image.id) ?? image)
  }
}
