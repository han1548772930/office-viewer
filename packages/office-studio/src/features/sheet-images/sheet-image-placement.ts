import type { SheetDrawingAnchorType, SheetFloatingImage, SheetImageAnchor, SheetImageBounds, SheetImageCellMarker,
  SheetImagePlacement, SheetImagePlacementInput, SheetImageSourceRect } from '../../domain/sheet-images.ts'
import type { RenderRect } from '../../rendering/contracts.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { drawingAffineBounds, drawingAffineFrame, drawingAffineFromFrame, drawingAffinePoint, normalizeDrawingAffineGeometry,
  invertDrawingAffine, multiplyDrawingAffine, type DrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { normalizeSheetFloatingImage, sheetImageIndex, sheetImageNumber, sheetImageRecord } from './sheet-image-model.ts'

export interface SheetImageAxis { count: number; offset(index: number): number; size(index: number): number; locate(coordinate: number): number }
export interface SheetImageAxes { rows: SheetImageAxis; columns: SheetImageAxis }
export type SheetFloatingImagePatch = Partial<Pick<SheetFloatingImage, 'anchor' | 'endAnchor' | 'absolutePosition' | 'anchorType' | 'size' |
  'opacity' | 'keepRatio' | 'altText' | 'title' | 'resource' | 'crop' | 'rotation' | 'flipX' | 'flipY' | 'drawingGeometry'>> & { placement?: SheetImagePlacementInput }

export function sheetImageAxisSwapped(rotation = 0): boolean {
  const angle = ((rotation % 360) + 360) % 360
  return angle >= 45 && angle < 135 || angle >= 225 && angle < 315
}
function projectOffset(axis: SheetImageAxis, index: number, offset: number): number {
  const extent = axis.size(index)
  return axis.offset(index) + (extent === 0 ? 0 : index === axis.count - 1 ? offset : Math.min(offset, extent))
}
export function sheetImageProjectAnchor(axes: SheetImageAxes, anchor: SheetImageAnchor) {
  return { left: projectOffset(axes.columns, anchor.column, anchor.offsetX), top: projectOffset(axes.rows, anchor.row, anchor.offsetY) }
}
export function sheetImageAnchorAt(axes: SheetImageAxes, left: number, top: number): SheetImageAnchor {
  const column = axes.columns.locate(Math.max(0, left)), row = axes.rows.locate(Math.max(0, top))
  return { row, column, offsetX: left - axes.columns.offset(column), offsetY: top - axes.rows.offset(row) }
}
export const sheetImagePublicMarker = (anchor: SheetImageAnchor): SheetImageCellMarker =>
  ({ row: anchor.row, column: anchor.column, rowOffset: anchor.offsetY, columnOffset: anchor.offsetX })
function internalMarker(value: unknown): SheetImageAnchor {
  const marker = sheetImageRecord(value, '图片锚点')
  return { row: sheetImageIndex(marker.row, 'row'), column: sheetImageIndex(marker.column, 'column'),
    offsetX: sheetImageNumber(marker.columnOffset, '图片列偏移', -Number.MAX_SAFE_INTEGER),
    offsetY: sheetImageNumber(marker.rowOffset, '图片行偏移', -Number.MAX_SAFE_INTEGER) }
}
export function sheetImageMajorBounds(frame: SheetImageBounds, rotation = 0): SheetImageBounds {
  return sheetImageAxisSwapped(rotation) ? { left: frame.left + (frame.width - frame.height) / 2,
    top: frame.top + (frame.height - frame.width) / 2, width: frame.height, height: frame.width } : { ...frame }
}
/** One geometry authority for rendering, permissions, placement, clipboard and structural edits. */
export function sheetImageFrame(image: SheetFloatingImage, axes: SheetImageAxes): SheetImageBounds {
  if (image.drawingGeometry) {
    const geometry = image.drawingGeometry, m = geometry.matrix, center = drawingAffinePoint(m, geometry.width / 2, geometry.height / 2),
      width = Math.hypot(m.a, m.b) * geometry.width, height = Math.hypot(m.c, m.d) * geometry.height
    if (image.anchorType === '1') {
      if (!image.endAnchor) throw new TypeError('随单元格缩放的图片缺少结束锚点。')
      const origin = sheetImageProjectAnchor(axes, image.anchor), end = sheetImageProjectAnchor(axes, image.endAnchor)
      let projectedWidth = Math.max(0, end.left - origin.left), projectedHeight = Math.max(0, end.top - origin.top)
      if (sheetImageAxisSwapped(image.rotation)) [projectedWidth, projectedHeight] = [projectedHeight, projectedWidth]
      if (projectedWidth <= 0 || projectedHeight <= 0) return { left: center.x - width / 2, top: center.y - height / 2,
        width: projectedWidth <= 0 ? 0 : width, height: projectedHeight <= 0 ? 0 : height }
    }
    return { left: center.x - width / 2, top: center.y - height / 2, width, height }
  }
  if (image.anchorType === '2' && image.absolutePosition) return { ...image.absolutePosition, ...image.size }
  const origin = sheetImageProjectAnchor(axes, image.anchor)
  let width = image.size.width, height = image.size.height
  if (image.anchorType === '1') {
    if (!image.endAnchor) throw new TypeError('随单元格缩放的图片缺少结束锚点。')
    const end = sheetImageProjectAnchor(axes, image.endAnchor)
    width = Math.max(0, end.left - origin.left); height = Math.max(0, end.top - origin.top)
    if (sheetImageAxisSwapped(image.rotation)) [width, height] = [height, width]
  }
  if (sheetImageAxisSwapped(image.rotation)) {
    origin.left += (height - width) / 2; origin.top += (width - height) / 2
  }
  return { ...origin, width, height }
}
export function sheetImageRotatedBounds(frame: SheetImageBounds, rotation = 0): SheetImageBounds {
  const angle = (rotation % 360) * Math.PI / 180, width = Math.abs(frame.width * Math.cos(angle)) + Math.abs(frame.height * Math.sin(angle)),
    height = Math.abs(frame.width * Math.sin(angle)) + Math.abs(frame.height * Math.cos(angle))
  return { left: frame.left + (frame.width - width) / 2, top: frame.top + (frame.height - height) / 2, width, height }
}
export function sheetImageFootprint(image: SheetFloatingImage, axes: SheetImageAxes): SheetImageBounds {
  if (image.drawingGeometry) {
    const frame = sheetImageFrame(image, axes)
    return frame.width <= 0 || frame.height <= 0 ? { ...frame, width: 0, height: 0 } : drawingAffineBounds(image.drawingGeometry)
  }
  return sheetImageRotatedBounds(sheetImageFrame(image, axes), image.rotation)
}
export function sheetImageDrawingGeometry(image: SheetFloatingImage, axes: SheetImageAxes): DrawingAffineGeometry {
  return image.drawingGeometry ? normalizeDrawingAffineGeometry(image.drawingGeometry) : drawingAffineFromFrame(sheetImageFrame(image, axes), image.rotation)
}
/** Express a world pointer delta in the image's two displayed frame axes, including shear. */
export function sheetImageLocalFrameDelta(image: SheetFloatingImage, dx: number, dy: number): { x: number; y: number } {
  const geometry = image.drawingGeometry
  if (geometry) {
    const inverse = invertDrawingAffine(geometry.matrix), m = geometry.matrix
    return { x: (inverse.a * dx + inverse.c * dy) * Math.hypot(m.a, m.b),
      y: (inverse.b * dx + inverse.d * dy) * Math.hypot(m.c, m.d) }
  }
  const angle = ((image.rotation ?? 0) % 360) * Math.PI / 180
  return { x: dx * Math.cos(angle) + dy * Math.sin(angle), y: -dx * Math.sin(angle) + dy * Math.cos(angle) }
}
export function sheetImageWorldFrameDelta(image: SheetFloatingImage, dx: number, dy: number): { x: number; y: number } {
  const geometry = image.drawingGeometry
  if (geometry) {
    const m = geometry.matrix, x = dx / Math.hypot(m.a, m.b), y = dy / Math.hypot(m.c, m.d)
    return { x: m.a * x + m.c * y, y: m.b * x + m.d * y }
  }
  const angle = ((image.rotation ?? 0) % 360) * Math.PI / 180
  return { x: dx * Math.cos(angle) - dy * Math.sin(angle), y: dx * Math.sin(angle) + dy * Math.cos(angle) }
}
export function sheetImagePlacement(image: SheetFloatingImage, axes: SheetImageAxes): SheetImagePlacement {
  if (image.anchorType === '2') return { kind: '2', ...sheetImageFrame(image, axes) }
  if (image.anchorType === '1') {
    if (!image.endAnchor) throw new TypeError('图片缺少结束锚点。')
    return { kind: '1', from: sheetImagePublicMarker(image.anchor), to: sheetImagePublicMarker(image.endAnchor) }
  }
  return { kind: '0', from: sheetImagePublicMarker(image.anchor), width: image.size.width, height: image.size.height }
}
function validateFrame(frame: SheetImageBounds): void {
  sheetImageNumber(frame.left, '图片水平位置', -Number.MAX_SAFE_INTEGER)
  sheetImageNumber(frame.top, '图片垂直位置', -Number.MAX_SAFE_INTEGER)
  sheetImageNumber(frame.width, '图片宽度', 1, 100_000); sheetImageNumber(frame.height, '图片高度', 1, 100_000)
}
/** The authored frame changes crop offsets proportionally; changing cells never re-anchors its markers. */
export function sheetImageWithFrame(image: SheetFloatingImage, frame: SheetImageBounds, axes: SheetImageAxes,
  kind: SheetDrawingAnchorType = image.anchorType ?? '0'): SheetFloatingImage {
  validateFrame(frame)
  if (!['0', '1', '2'].includes(kind)) throw new TypeError('图片锚点类型无效。')
  const major = sheetImageMajorBounds(frame, image.rotation), next = cloneModel(image)
  next.anchorType = kind; next.anchor = sheetImageAnchorAt(axes, major.left, major.top); next.size = { width: frame.width, height: frame.height }
  delete next.endAnchor; delete next.absolutePosition
  if (kind === '1') next.endAnchor = sheetImageAnchorAt(axes, major.left + major.width, major.top + major.height)
  if (kind === '2') next.absolutePosition = { left: frame.left, top: frame.top }
  if (image.crop) next.crop = { top: image.crop.top * frame.height / image.size.height, bottom: image.crop.bottom * frame.height / image.size.height,
    left: image.crop.left * frame.width / image.size.width, right: image.crop.right * frame.width / image.size.width }
  if (image.drawingGeometry) {
    const current = sheetImageFrame(image, axes), before = image.drawingGeometry,
      width = before.width * frame.width / current.width, height = before.height * frame.height / current.height,
      matrix = { ...before.matrix }, center = drawingAffinePoint(matrix, width / 2, height / 2)
    matrix.e += frame.left + frame.width / 2 - center.x; matrix.f += frame.top + frame.height / 2 - center.y
    next.drawingGeometry = normalizeDrawingAffineGeometry({ width, height, matrix })
  }
  return normalizeSheetFloatingImage(next)
}
/** Bake compatibility markers from the affine frame while retaining the complete world transform. */
export function sheetImageWithDrawingGeometry(image: SheetFloatingImage, geometry: DrawingAffineGeometry, axes: SheetImageAxes): SheetFloatingImage {
  const valid = normalizeDrawingAffineGeometry(geometry), m = valid.matrix,
    center = drawingAffinePoint(m, valid.width / 2, valid.height / 2), width = Math.hypot(m.a, m.b) * valid.width,
    height = Math.hypot(m.c, m.d) * valid.height, base = cloneModel(image)
  const before = image.drawingGeometry
  if (before && before.width === valid.width && before.height === valid.height
    && (['a', 'b', 'c', 'd'] as const).every(key => before.matrix[key] === m[key])) {
    const frame = drawingAffineFrame(valid), major = sheetImageMajorBounds(frame, image.rotation)
    base.drawingGeometry = valid
    base.anchor = sheetImageAnchorAt(axes, major.left, major.top)
    if (image.anchorType === '1') base.endAnchor = sheetImageAnchorAt(axes, major.left + major.width, major.top + major.height)
    if (image.anchorType === '2') base.absolutePosition = { left: frame.left, top: frame.top }
    // Both retains its original size/crop author basis after row/column dimensions change.
    return normalizeSheetFloatingImage(base)
  }
  delete base.drawingGeometry
  base.rotation = Math.atan2(m.b, m.a) * 180 / Math.PI
  const next = sheetImageWithFrame(base, { left: center.x - width / 2, top: center.y - height / 2, width, height }, axes)
  next.drawingGeometry = valid
  return next
}
export function applySheetImagePlacement(image: SheetFloatingImage, value: SheetImagePlacementInput, axes: SheetImageAxes): SheetFloatingImage {
  const input = sheetImageRecord(value, '图片放置'), kind = input.kind
  if (kind !== '0' && kind !== '1' && kind !== '2') throw new TypeError('图片锚点类型无效。')
  if (input.bounds !== undefined || kind === '2') {
    const frame = sheetImageRecord(input.bounds ?? input, '图片位置')
    return sheetImageWithFrame(image, { left: sheetImageNumber(frame.left, '图片水平位置', -Number.MAX_SAFE_INTEGER),
      top: sheetImageNumber(frame.top, '图片垂直位置', -Number.MAX_SAFE_INTEGER), width: sheetImageNumber(frame.width, '图片宽度', 1, 100_000),
      height: sheetImageNumber(frame.height, '图片高度', 1, 100_000) }, axes, kind)
  }
  const next = cloneModel(image); next.anchorType = kind; next.anchor = internalMarker(input.from)
  delete next.absolutePosition; delete next.endAnchor
  delete next.drawingGeometry
  if (kind === '1') next.endAnchor = internalMarker(input.to)
  else next.size = { width: sheetImageNumber(input.width, '图片宽度', 1, 100_000), height: sheetImageNumber(input.height, '图片高度', 1, 100_000) }
  const frame = sheetImageFrame(next, axes); validateFrame(frame)
  if (image.drawingGeometry) {
    const transformed = sheetImageWithFrame(image, frame, axes, kind)
    transformed.anchor = next.anchor
    if (next.endAnchor) transformed.endAnchor = next.endAnchor
    return transformed
  }
  if (image.crop) next.crop = { top: image.crop.top * frame.height / image.size.height, bottom: image.crop.bottom * frame.height / image.size.height,
    left: image.crop.left * frame.width / image.size.width, right: image.crop.right * frame.width / image.size.width }
  next.size = { width: frame.width, height: frame.height }
  return normalizeSheetFloatingImage(next)
}
export function patchSheetFloatingImage(image: SheetFloatingImage, patch: SheetFloatingImagePatch, axes?: SheetImageAxes): SheetFloatingImage {
  const { placement, ...fields } = patch
  if (!axes) {
    if (placement || fields.drawingGeometry || fields.anchorType !== undefined || fields.endAnchor || fields.absolutePosition || fields.rotation !== undefined && fields.rotation !== image.rotation)
      throw new TypeError('图片位置更新需要目标工作表的尺寸索引。')
    return normalizeSheetFloatingImage({ ...image, ...fields, id: image.id })
  }
  const before = sheetImageFrame(image, axes), changedOrientation = fields.rotation !== undefined && fields.rotation !== image.rotation
  let next: SheetFloatingImage = { ...cloneModel(image), ...cloneModel(fields), id: image.id }
  if (fields.drawingGeometry) {
    // Compare with the old geometry; merging the new matrix first would make every patch look like a translation.
    if (image.drawingGeometry) next.drawingGeometry = cloneModel(image.drawingGeometry)
    else delete next.drawingGeometry
    return sheetImageWithDrawingGeometry(next, fields.drawingGeometry, axes)
  }
  if (image.drawingGeometry && changedOrientation) {
    const angle = ((fields.rotation! - (image.rotation ?? 0)) % 360) * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle),
      cx = before.left + before.width / 2, cy = before.top + before.height / 2,
      delta = { a: cos, b: sin, c: -sin, d: cos, e: cx - cos * cx + sin * cy, f: cy - sin * cx - cos * cy }
    next.drawingGeometry = { ...image.drawingGeometry, matrix: multiplyDrawingAffine(delta, image.drawingGeometry.matrix) }
  }
  if (placement) next = applySheetImagePlacement(next, placement, axes)
  else if (fields.absolutePosition) next = sheetImageWithFrame(next, { ...fields.absolutePosition, ...(fields.size ?? { width: before.width, height: before.height }) }, axes, '2')
  else if (fields.endAnchor) next = applySheetImagePlacement(next, { kind: '1', from: sheetImagePublicMarker(fields.anchor ?? image.anchor), to: sheetImagePublicMarker(fields.endAnchor) }, axes)
  else if (fields.anchor || fields.size || fields.anchorType !== undefined || changedOrientation) {
    let frame = { ...before, ...(fields.size ?? {}) }
    if (fields.anchor) {
      const start = sheetImageProjectAnchor(axes, fields.anchor), major = sheetImageMajorBounds(frame, next.rotation)
      frame = { ...frame, left: start.left + frame.left - major.left, top: start.top + frame.top - major.top }
    }
    // Use the old crop basis even when a size patch has already changed the compatibility extent.
    next.size = { ...image.size }
    next = sheetImageWithFrame(next, frame, axes, fields.anchorType ?? image.anchorType ?? '0')
  }
  if (fields.crop !== undefined) next.crop = cloneModel(fields.crop)
  return normalizeSheetFloatingImage(next)
}
export function sheetImageNormalizedCrop(crop: SheetImageSourceRect | undefined, size: SheetFloatingImage['size']): RenderRect | undefined {
  if (!crop) return undefined
  const width = size.width + crop.left + crop.right, height = size.height + crop.top + crop.bottom
  if (width <= 0 || height <= 0) return undefined
  return { x: crop.left / width, y: crop.top / height, width: size.width / width, height: size.height / height }
}
