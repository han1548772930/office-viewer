import type { SheetFloatingImage, SheetImageBounds } from '../../domain/sheet-images.ts'
import type { SheetImageResizeHandle } from './sheet-image-layout.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { sheetImageNumber, SHEET_IMAGE_MAX_COUNT } from './sheet-image-model.ts'
import { patchSheetFloatingImage, sheetImageFootprint, sheetImageFrame, sheetImageWithFrame, type SheetImageAxes } from './sheet-image-placement.ts'

export const SHEET_IMAGE_ALIGN_TYPES = { default: '0', left: '1', center: '2', right: '3', top: '4', middle: '5', bottom: '6', horizon: '7', vertical: '8' } as const
export type SheetImageAlignType = typeof SHEET_IMAGE_ALIGN_TYPES[keyof typeof SHEET_IMAGE_ALIGN_TYPES]
export const SHEET_IMAGE_ALIGN_OPERATIONS = { align: 'sheet.operation.set-image-align', left: 'sheet.operation.set-drawing-align-left',
  center: 'sheet.operation.set-drawing-align-center', right: 'sheet.operation.set-drawing-align-right', top: 'sheet.operation.set-drawing-align-top',
  middle: 'sheet.operation.set-drawing-align-middle', bottom: 'sheet.operation.set-drawing-align-bottom',
  horizon: 'sheet.operation.set-drawing-align-horizon', vertical: 'sheet.operation.set-drawing-align-vertical' } as const
export const SHEET_IMAGE_ALIGN_PRESETS = [{ value: '1', label: '左对齐' }, { value: '2', label: '水平居中' }, { value: '3', label: '右对齐' },
  { value: '4', label: '顶部对齐' }, { value: '5', label: '垂直居中' }, { value: '6', label: '底部对齐' },
  { value: '7', label: '水平分布' }, { value: '8', label: '垂直分布' }] as const
export function sheetImageAlignTypeForOperation(id: string, alignType?: SheetImageAlignType): SheetImageAlignType | undefined {
  if (id === SHEET_IMAGE_ALIGN_OPERATIONS.align) return alignType ?? '0'
  const key = (Object.keys(SHEET_IMAGE_ALIGN_OPERATIONS) as Array<keyof typeof SHEET_IMAGE_ALIGN_OPERATIONS>).find(key => SHEET_IMAGE_ALIGN_OPERATIONS[key] === id)
  return key && key !== 'align' ? SHEET_IMAGE_ALIGN_TYPES[key] : undefined
}
function assertSelection(images: readonly SheetFloatingImage[]): void {
  if (!images.length || images.length > SHEET_IMAGE_MAX_COUNT || new Set(images.map(image => image.id)).size !== images.length) throw new RangeError('图片多选对象无效。')
}
export function sheetImageSelectionBounds(images: readonly SheetFloatingImage[], axes: SheetImageAxes, rotated = true): SheetImageBounds {
  assertSelection(images)
  const frames = images.map(image => rotated ? sheetImageFootprint(image, axes) : sheetImageFrame(image, axes)),
    left = Math.min(...frames.map(frame => frame.left)), top = Math.min(...frames.map(frame => frame.top))
  return { left, top, width: Math.max(...frames.map(frame => frame.left + frame.width)) - left,
    height: Math.max(...frames.map(frame => frame.top + frame.height)) - top }
}
/** A rigid selection transform preserves each image's rotation, crop scale, anchor mode and relative position. */
export function transformSheetImageSelection(images: readonly SheetFloatingImage[], axes: SheetImageAxes, dx: number, dy: number,
  handle?: SheetImageResizeHandle, rotationDelta?: number): SheetFloatingImage[] {
  assertSelection(images); sheetImageNumber(dx, '图片水平位移', -Number.MAX_SAFE_INTEGER); sheetImageNumber(dy, '图片垂直位移', -Number.MAX_SAFE_INTEGER)
  const bounds = sheetImageSelectionBounds(images, axes), frames = images.map(image => sheetImageFrame(image, axes))
  if (rotationDelta !== undefined) {
    sheetImageNumber(rotationDelta, '图片旋转位移', -Number.MAX_SAFE_INTEGER)
    const angle = (rotationDelta % 360) * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle), cx = bounds.left + bounds.width / 2, cy = bounds.top + bounds.height / 2
    return images.map((image, index) => {
      const frame = frames[index]!, x = frame.left + frame.width / 2 - cx, y = frame.top + frame.height / 2 - cy,
        oriented = patchSheetFloatingImage(image, { rotation: (image.rotation ?? 0) + rotationDelta }, axes)
      return sheetImageWithFrame(oriented, { ...frame, left: cx + x * cos - y * sin - frame.width / 2,
        top: cy + x * sin + y * cos - frame.height / 2 }, axes)
    })
  }
  if (!handle) {
    const offsetX = Math.max(-bounds.left, dx), offsetY = Math.max(-bounds.top, dy)
    return images.map((image, index) => sheetImageWithFrame(image, { ...frames[index]!, left: frames[index]!.left + offsetX, top: frames[index]!.top + offsetY }, axes))
  }
  if (bounds.width <= 0 || bounds.height <= 0 || frames.some(frame => frame.width <= 0 || frame.height <= 0)) throw new RangeError('零尺寸图片多选暂时无法缩放。')
  const west = handle.includes('w'), east = handle.includes('e'), north = handle.includes('n'), south = handle.includes('s'),
    horizontal = west || east, vertical = north || south,
    factorX = 1 + dx * (west ? -1 : 1) / bounds.width, factorY = 1 + dy * (north ? -1 : 1) / bounds.height,
    requested = horizontal && (!vertical || Math.abs(dx / bounds.width) >= Math.abs(dy / bounds.height)) ? factorX : factorY,
    minimum = Math.max(...frames.flatMap(frame => [1 / frame.width, 1 / frame.height])),
    maximum = Math.min(...frames.flatMap(frame => [100_000 / frame.width, 100_000 / frame.height])), scale = Math.max(minimum, Math.min(maximum, requested)),
    width = bounds.width * scale, height = bounds.height * scale,
    left = west ? bounds.left + bounds.width - width : east ? bounds.left : bounds.left + (bounds.width - width) / 2,
    top = north ? bounds.top + bounds.height - height : south ? bounds.top : bounds.top + (bounds.height - height) / 2
  return images.map((image, index) => {
    const frame = frames[index]!
    return sheetImageWithFrame(image, { left: left + (frame.left - bounds.left) * scale, top: top + (frame.top - bounds.top) * scale,
      width: frame.width * scale, height: frame.height * scale }, axes)
  })
}
/** Mirrors v1.6.0 DrawingUpdateController: transform-frame alignment and its rounded span/count distribution. */
export function alignSheetImageSelection(images: readonly SheetFloatingImage[], axes: SheetImageAxes, alignType: SheetImageAlignType): SheetFloatingImage[] {
  assertSelection(images)
  if (!(Object.values(SHEET_IMAGE_ALIGN_TYPES) as readonly string[]).includes(alignType)) throw new TypeError('图片对齐方式无效。')
  if (alignType === '0' || images.length < 2) return images.map(cloneModel)
  const bounds = sheetImageSelectionBounds(images, axes, false), items = images.map((image, index) => ({ image, frame: sheetImageFrame(image, axes), index }))
  if (alignType === '7') items.sort((a, b) => a.frame.left + a.frame.width / 2 - b.frame.left - b.frame.width / 2 || a.index - b.index)
  if (alignType === '8') items.sort((a, b) => a.frame.top + a.frame.height / 2 - b.frame.top - b.frame.height / 2 || a.index - b.index)
  const horizontal = Math.round(bounds.width / images.length * 10) / 10, vertical = Math.round(bounds.height / images.length * 10) / 10,
    results = new Map(items.map(({ image, frame }, index) => {
      const left = alignType === '1' ? bounds.left : alignType === '2' ? bounds.left + (bounds.width - frame.width) / 2
        : alignType === '3' ? bounds.left + bounds.width - frame.width : alignType === '7' ? bounds.left + horizontal * index : frame.left,
        top = alignType === '4' ? bounds.top : alignType === '5' ? bounds.top + (bounds.height - frame.height) / 2
          : alignType === '6' ? bounds.top + bounds.height - frame.height : alignType === '8' ? bounds.top + vertical * index : frame.top
      return [image.id, left === frame.left && top === frame.top ? cloneModel(image) : sheetImageWithFrame(image, { ...frame, left, top }, axes)] as const
    }))
  return images.map(image => results.get(image.id)!)
}
