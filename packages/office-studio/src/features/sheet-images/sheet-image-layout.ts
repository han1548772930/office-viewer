import type { SheetFloatingImage, SheetImageResource } from '../../domain/sheet-images.ts'
import type { CellRange } from '../../domain/model.ts'
import type { DrawCommand, RenderRect } from '../../rendering/contracts.ts'
import { drawingAffineBounds, drawingAffinePoint, invertDrawingAffine, multiplyDrawingAffine, type DrawingAffineMatrix } from '../../domain/drawing-affine.ts'
import { intersectRect } from '../../rendering/geometry.ts'
import { cellAddress, type GridViewport, type SpreadsheetGeometry } from '../../rendering/spreadsheet/layout.ts'
import { patchSheetFloatingImage, sheetImageFootprint, sheetImageFrame, sheetImageLocalFrameDelta, sheetImageWorldFrameDelta,
  sheetImageNormalizedCrop, sheetImageRotatedBounds, sheetImageWithFrame } from './sheet-image-placement.ts'

export interface SheetImagePaneLayout { image: SheetFloatingImage; rect: RenderRect; bounds: RenderRect; clip: RenderRect; frozenX: boolean; frozenY: boolean; matrix?: DrawingAffineMatrix }
export const SHEET_IMAGE_HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
export type SheetImageResizeHandle = typeof SHEET_IMAGE_HANDLES[number]
export type SheetImageInteractionHandle = SheetImageResizeHandle | 'rotate'
export function sheetImageGeometryRange(geometry: SpreadsheetGeometry, image: SheetFloatingImage): CellRange {
  const frame = sheetImageFootprint(image, geometry)
  const locate = (axis: SpreadsheetGeometry['rows'], coordinate: number) => Math.max(0, Math.min(axis.count - 1, axis.locate(Math.max(0, coordinate))))
  const interior = (end: number, start: number) => end > start ? Math.max(start, end - Math.max(1e-7, Math.abs(end) * Number.EPSILON)) : start
  return { startRow: locate(geometry.rows, frame.top), endRow: locate(geometry.rows, interior(frame.top + frame.height, frame.top)),
    startCol: locate(geometry.columns, frame.left), endCol: locate(geometry.columns, interior(frame.left + frame.width, frame.left)) }
}
export function sheetImageHandlePoints(rect: RenderRect, rotation = 0, matrix?: DrawingAffineMatrix) {
  const points = { nw: { x: rect.x, y: rect.y }, n: { x: rect.x + rect.width / 2, y: rect.y }, ne: { x: rect.x + rect.width, y: rect.y },
    e: { x: rect.x + rect.width, y: rect.y + rect.height / 2 }, se: { x: rect.x + rect.width, y: rect.y + rect.height },
    s: { x: rect.x + rect.width / 2, y: rect.y + rect.height }, sw: { x: rect.x, y: rect.y + rect.height }, w: { x: rect.x, y: rect.y + rect.height / 2 } }
  const angle = (rotation % 360) * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle), cx = rect.x + rect.width / 2, cy = rect.y + rect.height / 2
  return Object.fromEntries(Object.entries(points).map(([name, point]) => {
    const dx = point.x - cx, dy = point.y - cy
    const x = cx + dx * cos - dy * sin, y = cy + dx * sin + dy * cos
    return [name, matrix ? drawingAffinePoint(matrix, x, y) : { x, y }]
  })) as Record<SheetImageResizeHandle, { x: number; y: number }>
}
export function sheetImageRotationPoint(rect: RenderRect, rotation = 0, matrix?: DrawingAffineMatrix) {
  if (matrix) {
    const center = drawingAffinePoint(matrix, rect.x + rect.width / 2, rect.y + rect.height / 2),
      top = drawingAffinePoint(matrix, rect.x + rect.width / 2, rect.y), dx = top.x - center.x, dy = top.y - center.y,
      length = Math.hypot(dx, dy)
    return { x: top.x + dx / length * 24, y: top.y + dy / length * 24 }
  }
  const angle = (rotation % 360) * Math.PI / 180, distance = rect.height / 2 + 24
  return { x: rect.x + rect.width / 2 + Math.sin(angle) * distance, y: rect.y + rect.height / 2 - Math.cos(angle) * distance }
}
export function sheetImagePaneLayouts(geometry: SpreadsheetGeometry, viewport: GridViewport, images = geometry.sheet.images ?? []): SheetImagePaneLayout[] {
  return images.flatMap(image => geometry.panes(viewport).flatMap(pane => {
    if (!pane.clip.width || !pane.clip.height) return []
    const frozenX = geometry.frozenColumns > 0 && pane.clip.x === geometry.headerWidth
    const frozenY = geometry.frozenRows > 0 && pane.clip.y === geometry.headerHeight
    const frame = sheetImageFrame(image, geometry)
    if (frame.width <= 0 || frame.height <= 0) return []
    let rect = { x: geometry.headerWidth + frame.left * viewport.zoom - (frozenX ? 0 : viewport.scrollLeft),
      y: geometry.headerHeight + frame.top * viewport.zoom - (frozenY ? 0 : viewport.scrollTop),
      width: frame.width * viewport.zoom, height: frame.height * viewport.zoom }
    const affine = image.drawingGeometry, matrix = affine ? multiplyDrawingAffine({ a: viewport.zoom, b: 0, c: 0, d: viewport.zoom,
      e: geometry.headerWidth - (frozenX ? 0 : viewport.scrollLeft), f: geometry.headerHeight - (frozenY ? 0 : viewport.scrollTop) }, affine.matrix) : undefined
    if (affine) rect = { x: 0, y: 0, width: affine.width, height: affine.height }
    const footprint = matrix ? drawingAffineBounds({ width: rect.width, height: rect.height, matrix })
      : sheetImageRotatedBounds({ left: rect.x, top: rect.y, width: rect.width, height: rect.height }, image.rotation),
      bounds = { x: footprint.left, y: footprint.top, width: footprint.width, height: footprint.height }
    const visible = intersectRect(bounds, pane.clip)
    return visible.width > 0 && visible.height > 0 ? [{ image, rect, bounds, clip: pane.clip, frozenX, frozenY, matrix }] : []
  }))
}
export function hitSheetFloatingImage(layouts: readonly SheetImagePaneLayout[], x: number, y: number, selectedId?: string, allowRotation = true): { layout: SheetImagePaneLayout; handle?: SheetImageInteractionHandle } | undefined {
  const inside = (rect: RenderRect) => x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height
  // Selection controls are drawn above every image and receive the same hit priority.
  for (const layout of [...layouts].reverse()) {
    if (!inside(layout.clip)) continue
    if (layout.image.id === selectedId) {
      const rotation = sheetImageRotationPoint(layout.rect, layout.matrix ? 0 : layout.image.rotation, layout.matrix)
      if (allowRotation && Math.hypot(x - rotation.x, y - rotation.y) <= 7) return { layout, handle: 'rotate' }
      for (const [handle, point] of Object.entries(sheetImageHandlePoints(layout.rect, layout.matrix ? 0 : layout.image.rotation, layout.matrix))) {
        if (Math.abs(x - point.x) <= 5 && Math.abs(y - point.y) <= 5) return { layout, handle: handle as SheetImageResizeHandle }
      }
    }
  }
  for (const layout of [...layouts].reverse()) {
    if (!inside(layout.clip)) continue
    const point = layout.matrix ? drawingAffinePoint(invertDrawingAffine(layout.matrix), x, y) : { x, y },
      angle = (layout.matrix ? 0 : (layout.image.rotation ?? 0) % 360) * Math.PI / 180, cx = layout.rect.x + layout.rect.width / 2, cy = layout.rect.y + layout.rect.height / 2,
      dx = point.x - cx, dy = point.y - cy, localX = dx * Math.cos(angle) + dy * Math.sin(angle), localY = -dx * Math.sin(angle) + dy * Math.cos(angle)
    if (Math.abs(localX) <= layout.rect.width / 2 && Math.abs(localY) <= layout.rect.height / 2) return { layout }
  }
  return undefined
}
export function sheetImageDrawCommands(layouts: readonly SheetImagePaneLayout[], source: (resource: SheetImageResource) => CanvasImageSource | undefined,
  selectedId?: string, color = '#246bfd', handleFill = '#ffffff', options: { crop?: boolean; selectedIds?: readonly string[]; outlineOnly?: boolean } = {}): DrawCommand[] {
  const content: DrawCommand[] = [], controls: DrawCommand[] = []
  for (const layout of layouts) {
    const result = controls, decoded = source(layout.image.resource)
    if (decoded) content.push({ kind: 'image', drawingRef: { kind: 'image', id: layout.image.id }, ...layout.rect, source: decoded, opacity: layout.image.opacity, clip: layout.clip,
      sourceRect: sheetImageNormalizedCrop(layout.image.crop, layout.image.size), rotation: layout.matrix ? undefined : layout.image.rotation,
      matrix: layout.matrix, flipX: layout.image.flipX, flipY: layout.image.flipY })
    if (layout.image.id === selectedId || options.selectedIds?.includes(layout.image.id)) {
      const handles = sheetImageHandlePoints(layout.rect, layout.matrix ? 0 : layout.image.rotation, layout.matrix), corners = [handles.nw, handles.ne, handles.se, handles.sw]
      for (let index = 0; index < corners.length; index++) result.push({ kind: 'line', x1: corners[index]!.x, y1: corners[index]!.y,
        x2: corners[(index + 1) % corners.length]!.x, y2: corners[(index + 1) % corners.length]!.y, color, width: 1, clip: layout.clip })
      if (options.outlineOnly) continue
      const rotation = sheetImageRotationPoint(layout.rect, layout.matrix ? 0 : layout.image.rotation, layout.matrix)
      if (!options.crop) result.push({ kind: 'line', x1: handles.n.x, y1: handles.n.y, x2: rotation.x, y2: rotation.y, color, width: 1, clip: layout.clip })
      for (const point of [...Object.values(handles), ...(options.crop ? [] : [rotation])]) result.push({ kind: 'rect', x: point.x - 3, y: point.y - 3,
        width: 6, height: 6, fill: handleFill, stroke: color, strokeWidth: 1, clip: layout.clip })
    }
  }
  return [...content, ...controls]
}
/** Only viewport media is decoded. Cell images use the same merged-cell owner as text and hit-testing. */
export function visibleSheetImageResources(geometry: SpreadsheetGeometry, viewport: GridViewport, visible: (row: number, column: number) => boolean = () => true, backgroundVisible = true,
  floatingVisible: (image: SheetFloatingImage) => boolean = image => visible(image.anchor.row, image.anchor.column)): SheetImageResource[] {
  const resources = new Map<string, SheetImageResource>()
  const add = (resource: SheetImageResource) => resources.set(`${resource.id}\0${resource.source}`, resource)
  if (backgroundVisible && geometry.sheet.backgroundImage?.resource) add(geometry.sheet.backgroundImage.resource)
  for (const layout of sheetImagePaneLayouts(geometry, viewport)) if (floatingVisible(layout.image)) add(layout.image.resource)
  for (const pane of geometry.panes(viewport)) for (const row of pane.rows) for (const column of pane.columns) {
    const merge = geometry.mergeAt(row, column), cell = geometry.sheet.cells[cellAddress(merge?.startRow ?? row, merge?.startColumn ?? column)]
    if (cell?.image && visible(merge?.startRow ?? row, merge?.startColumn ?? column)) add(cell.image.resource)
  }
  return [...resources.values()]
}
export function fitSheetCellImage(rect: RenderRect, width: number, height: number, padding = 2): RenderRect {
  const availableWidth = Math.max(0, rect.width - padding * 2), availableHeight = Math.max(0, rect.height - padding * 2)
  const scale = Math.min(availableWidth / width, availableHeight / height)
  const resultWidth = width * scale, resultHeight = height * scale
  return { x: rect.x + (rect.width - resultWidth) / 2, y: rect.y + (rect.height - resultHeight) / 2, width: resultWidth, height: resultHeight }
}
/** Drag coordinates are worksheet CSS pixels; resizing retains the opposite edge and optional aspect ratio. */
export function moveSheetImage(geometry: SpreadsheetGeometry, image: SheetFloatingImage, dx: number, dy: number, handle?: SheetImageResizeHandle): SheetFloatingImage {
  const frame = sheetImageFrame(image, geometry), startX = frame.left, startY = frame.top
  let x = startX, y = startY, width = frame.width, height = frame.height
  if (!handle) { x = Math.max(0, x + dx); y = Math.max(0, y + dy) }
  else {
    const local = sheetImageLocalFrameDelta(image, dx, dy), localX = local.x, localY = local.y
    const west = handle.includes('w'), east = handle.includes('e'), north = handle.includes('n'), south = handle.includes('s')
    const oldWidth = width, oldHeight = height
    if (west || east) width = Math.max(1, oldWidth + (west ? -localX : localX))
    if (north || south) height = Math.max(1, oldHeight + (north ? -localY : localY))
    if (image.keepRatio) { if (west || east) height = width * oldHeight / oldWidth; else width = height * oldWidth / oldHeight }
    const centerDx = west ? (oldWidth - width) / 2 : east ? (width - oldWidth) / 2 : 0,
      centerDy = north ? (oldHeight - height) / 2 : south ? (height - oldHeight) / 2 : 0,
      delta = sheetImageWorldFrameDelta(image, centerDx, centerDy)
    x = startX + oldWidth / 2 + delta.x - width / 2
    y = startY + oldHeight / 2 + delta.y - height / 2
  }
  return sheetImageWithFrame(image, { left: x, top: y, width, height }, geometry)
}
export function rotateSheetImage(geometry: SpreadsheetGeometry, image: SheetFloatingImage, delta: number, snap = false): SheetFloatingImage {
  const angle = (image.rotation ?? 0) + delta
  return patchSheetFloatingImage(image, { rotation: snap ? Math.round(angle / 15) * 15 : angle }, geometry)
}
