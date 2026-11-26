import { createBlankWorkbook } from '../../src/domain/model.ts'
import { sheetClipboardObjectAxes } from '../../src/features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { normalizeSheetFloatingImage } from '../../src/features/sheet-images/sheet-image-model.ts'
import { sheetImageResourceFromDataUrl } from '../../src/features/sheet-images/sheet-image-source.ts'
import { sheetImageWithDrawingGeometry, sheetImageWithFrame } from '../../src/features/sheet-images/sheet-image-placement.ts'

export const sheetImagePng = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
export const sheetImageCopy = value => structuredClone(value)
export function sheetImageFixture() {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'sheet-images-book'; sheet.id = 'source'; workbook.activeSheetId = sheet.id
  sheet.rowCount = 80; sheet.columnCount = 40
  const destination = { ...sheetImageCopy(sheet), id: 'destination', name: 'Destination', cells: {}, images: [] }
  workbook.sheets.push(destination)
  return { workbook, sheet, destination, axes: sheetClipboardObjectAxes(workbook, sheet) }
}
export function sheetImageAt(id, axes, frame = { left: 80, top: 60, width: 80, height: 40 }, patch = {}) {
  const { drawingGeometry, ...fields } = patch
  const base = normalizeSheetFloatingImage({ id, resource: sheetImageResourceFromDataUrl(sheetImagePng, 'sample.png', 'media-' + id),
    anchor: { row: 0, column: 0, offsetX: 0, offsetY: 0 }, size: { width: 80, height: 40 },
    endAnchor: patch.anchorType === '1' ? { row: 1, column: 1, offsetX: 0, offsetY: 0 } : undefined,
    absolutePosition: patch.anchorType === '2' ? { left: frame.left, top: frame.top } : undefined,
    title: '<Picture & title>', altText: '"Alternative"', keepRatio: false, opacity: 0.7, ...fields })
  return drawingGeometry ? sheetImageWithDrawingGeometry(base, drawingGeometry, axes) : sheetImageWithFrame(base, frame, axes)
}
export function sheetImageAffine(id, axes, patch = {}) {
  return sheetImageAt(id, axes, undefined, { crop: { left: 8, right: 16, top: 6, bottom: 10 }, flipX: true, flipY: true,
    metadata: { opaqueAuthor: { retained: ['value', 4] } }, ...patch,
    drawingGeometry: patch.drawingGeometry ?? { width: 60, height: 40, matrix: { a: -1, b: 0.2, c: 0.3, d: 1, e: 160, f: 80 } } })
}
export function sheetImageTestAxes(rowSizes = [], columnSizes = [], hiddenRows = []) {
  const axis = (sizes, fallback, hidden = []) => {
    const count = 64, hiddenSet = new Set(hidden), offsets = [0]
    const size = index => hiddenSet.has(index) ? 0 : sizes[index] ?? fallback
    for (let index = 0; index < count; index++) offsets.push(offsets[index] + size(index))
    return { count, size, offset: index => offsets[index], locate: coordinate => {
      let index = 0
      while (index < count - 1 && offsets[index + 1] <= coordinate) index++
      return index
    } }
  }
  return { rows: axis(rowSizes, 20, hiddenRows), columns: axis(columnSizes, 64) }
}
