import type { SheetFloatingImage } from '../../domain/sheet-images.ts'
import { normalizeSheetFloatingImage, sheetImageNumber } from './sheet-image-model.ts'
import { sheetImageFrame, sheetImageWithFrame, type SheetImageAxes } from './sheet-image-placement.ts'

export const SHEET_IMAGE_CROP_TYPES = { FREE: '0', R1_1: '1', R16_9: '2', R9_16: '3', R5_4: '4', R4_5: '5', R4_3: '6', R3_4: '7', R3_2: '8', R2_3: '9' } as const
export type SheetImageCropType = typeof SHEET_IMAGE_CROP_TYPES[keyof typeof SHEET_IMAGE_CROP_TYPES]
export const SHEET_IMAGE_CROP_PRESETS: ReadonlyArray<{ value: SheetImageCropType; label: string; ratio?: number }> = [
  { value: '0', label: '自由裁剪' }, { value: '1', label: '1:1', ratio: 1 }, { value: '2', label: '16:9', ratio: 16 / 9 },
  { value: '3', label: '9:16', ratio: 9 / 16 }, { value: '4', label: '5:4', ratio: 5 / 4 }, { value: '5', label: '4:5', ratio: 4 / 5 },
  { value: '6', label: '4:3', ratio: 4 / 3 }, { value: '7', label: '3:4', ratio: 3 / 4 }, { value: '8', label: '3:2', ratio: 3 / 2 }, { value: '9', label: '2:3', ratio: 2 / 3 },
]
/** Center a ratio window in the complete authored source frame while retaining its pixel scale and rotation. */
export function cropSheetImageToRatio(image: SheetFloatingImage, axes: SheetImageAxes, ratio: number): SheetFloatingImage {
  sheetImageNumber(ratio, '图片裁剪比例', 0.01, 100)
  const frame = sheetImageFrame(image, axes), crop = image.crop ?? { top: 0, left: 0, bottom: 0, right: 0 },
    left = crop.left * frame.width / image.size.width, right = crop.right * frame.width / image.size.width,
    top = crop.top * frame.height / image.size.height, bottom = crop.bottom * frame.height / image.size.height,
    fullWidth = frame.width + left + right, fullHeight = frame.height + top + bottom,
    width = Math.min(fullWidth, fullHeight * ratio), height = width / ratio,
    dx = (right - left) / 2 * (image.flipX ? -1 : 1), dy = (bottom - top) / 2 * (image.flipY ? -1 : 1), angle = ((image.rotation ?? 0) % 360) * Math.PI / 180
  const next = sheetImageWithFrame(image, { left: frame.left + frame.width / 2 + dx * Math.cos(angle) - dy * Math.sin(angle) - width / 2,
    top: frame.top + frame.height / 2 + dx * Math.sin(angle) + dy * Math.cos(angle) - height / 2, width, height }, axes)
  next.crop = { left: (fullWidth - width) / 2, right: (fullWidth - width) / 2, top: (fullHeight - height) / 2, bottom: (fullHeight - height) / 2 }
  return normalizeSheetFloatingImage(next)
}
