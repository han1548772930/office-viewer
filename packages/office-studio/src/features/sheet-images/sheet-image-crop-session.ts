import type { SheetFloatingImage, SheetImageBounds, SheetImageSourceRect } from '../../domain/sheet-images.ts'
import type { SheetImageResizeHandle } from './sheet-image-layout.ts'
import { normalizeSheetFloatingImage, sheetImageNumber } from './sheet-image-model.ts'
import { sheetImageFrame, sheetImageWithFrame, sheetImageLocalFrameDelta, sheetImageWorldFrameDelta, type SheetImageAxes } from './sheet-image-placement.ts'
import { drawingAffinePoint, normalizeDrawingAffineGeometry, type DrawingAffineGeometry } from '../../domain/drawing-affine.ts'

export const SHEET_IMAGE_CROP_OPERATIONS = { open: 'sheet.operation.open-image-crop', close: 'sheet.operation.close-image-crop', auto: 'sheet.operation.Auto-image-crop' } as const
export type SheetImageCropOperation = typeof SHEET_IMAGE_CROP_OPERATIONS[keyof typeof SHEET_IMAGE_CROP_OPERATIONS]
export interface SheetImageCropResult { handled: boolean; image?: SheetFloatingImage; expectedRevision?: number }
const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value))
function cropGeometry(image: SheetFloatingImage, axes: SheetImageAxes): { frame: SheetImageBounds; crop: SheetImageSourceRect; width: number; height: number } {
  const frame = sheetImageFrame(image, axes), authored = image.crop ?? { left: 0, right: 0, top: 0, bottom: 0 },
    crop = { left: authored.left * frame.width / image.size.width, right: authored.right * frame.width / image.size.width,
      top: authored.top * frame.height / image.size.height, bottom: authored.bottom * frame.height / image.size.height }
  if (frame.width <= 0 || frame.height <= 0) throw new RangeError('隐藏行列中的零尺寸图片暂时无法裁剪。')
  return { frame, crop, width: frame.width + crop.left + crop.right, height: frame.height + crop.top + crop.bottom }
}
/** The translucent source uses the same world transform as the visible crop, including signed outcropping. */
export function sheetImageUncroppedFrame(image: SheetFloatingImage, axes: SheetImageAxes): SheetImageBounds {
  const { frame, crop, width, height } = cropGeometry(image, axes),
    delta = sheetImageWorldFrameDelta(image, (crop.right - crop.left) / 2 * (image.flipX ? -1 : 1), (crop.bottom - crop.top) / 2 * (image.flipY ? -1 : 1))
  return { left: frame.left + frame.width / 2 + delta.x - width / 2,
    top: frame.top + frame.height / 2 + delta.y - height / 2, width, height }
}
export function sheetImageUncroppedDrawingGeometry(image: SheetFloatingImage, axes: SheetImageAxes): DrawingAffineGeometry | undefined {
  const geometry = image.drawingGeometry
  if (!geometry) return undefined
  const frame = sheetImageUncroppedFrame(image, axes), visible = sheetImageFrame(image, axes),
    width = geometry.width * frame.width / visible.width, height = geometry.height * frame.height / visible.height,
    matrix = { ...geometry.matrix }, center = drawingAffinePoint(matrix, width / 2, height / 2)
  matrix.e += frame.left + frame.width / 2 - center.x; matrix.f += frame.top + frame.height / 2 - center.y
  return normalizeDrawingAffineGeometry({ width, height, matrix })
}
/** Dragging inside the crop moves its source; handles change the window and retain the opposite world edge. */
export function dragSheetImageCrop(image: SheetFloatingImage, axes: SheetImageAxes, dx: number, dy: number,
  handle?: SheetImageResizeHandle): SheetFloatingImage {
  sheetImageNumber(dx, '图片裁剪水平位移', -Number.MAX_SAFE_INTEGER); sheetImageNumber(dy, '图片裁剪垂直位移', -Number.MAX_SAFE_INTEGER)
  if (dx === 0 && dy === 0) return normalizeSheetFloatingImage(image)
  const { frame, crop, width: sourceWidth, height: sourceHeight } = cropGeometry(image, axes),
    local = sheetImageLocalFrameDelta(image, dx, dy), localX = local.x, localY = local.y
  if (!handle) {
    const left = clamp(crop.left - localX * (image.flipX ? -1 : 1), Math.min(0, sourceWidth - frame.width, crop.left), Math.max(0, sourceWidth - frame.width, crop.left)),
      top = clamp(crop.top - localY * (image.flipY ? -1 : 1), Math.min(0, sourceHeight - frame.height, crop.top), Math.max(0, sourceHeight - frame.height, crop.top)),
      next = sheetImageWithFrame(image, frame, axes)
    next.crop = { left, right: sourceWidth - frame.width - left, top, bottom: sourceHeight - frame.height - top }
    return normalizeSheetFloatingImage(next)
  }
  let left = 0, top = 0, right = frame.width, bottom = frame.height
  if (handle.includes('w')) left = clamp(localX, Math.min(0, -(image.flipX ? crop.right : crop.left)), right - 1)
  if (handle.includes('e')) right = clamp(frame.width + localX, left + 1, Math.max(frame.width, frame.width + (image.flipX ? crop.left : crop.right)))
  if (handle.includes('n')) top = clamp(localY, Math.min(0, -(image.flipY ? crop.bottom : crop.top)), bottom - 1)
  if (handle.includes('s')) bottom = clamp(frame.height + localY, top + 1, Math.max(frame.height, frame.height + (image.flipY ? crop.top : crop.bottom)))
  const width = right - left, height = bottom - top, centerX = (left + right - frame.width) / 2, centerY = (top + bottom - frame.height) / 2,
    delta = sheetImageWorldFrameDelta(image, centerX, centerY),
    next = sheetImageWithFrame(image, { left: frame.left + frame.width / 2 + delta.x - width / 2,
      top: frame.top + frame.height / 2 + delta.y - height / 2, width, height }, axes)
  next.crop = { left: crop.left + (image.flipX ? frame.width - right : left), right: crop.right + (image.flipX ? left : frame.width - right),
    top: crop.top + (image.flipY ? frame.height - bottom : top), bottom: crop.bottom + (image.flipY ? top : frame.height - bottom) }
  return normalizeSheetFloatingImage(next)
}
