import type { SheetCellImage, SheetDrawingAnchorType, SheetFloatingImage, SheetImageSourceRect } from '../../domain/sheet-images.ts'
import type { SheetImageCommandInput } from './sheet-image-commands.ts'
import { sheetImageIndex, sheetImageNumber, sheetImageText } from './sheet-image-model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { patchSheetFloatingImage, sheetImageFrame, type SheetImageAxes } from './sheet-image-placement.ts'
import { cropSheetImageToRatio } from './sheet-image-crop.ts'

export interface SheetImageEditorDraft {
  floating: boolean
  row: number
  column: number
  offsetX: number
  offsetY: number
  width: number
  height: number
  ratio: number
  keepRatio: boolean
  opacity: number
  title: string
  altText: string
  anchorType: SheetDrawingAnchorType
  left: number
  top: number
  rotation: number
  flipX: boolean
  flipY: boolean
  crop: SheetImageSourceRect
  initial?: Pick<SheetImageEditorDraft, 'row' | 'column' | 'offsetX' | 'offsetY' | 'width' | 'height' | 'anchorType' | 'left' | 'top' | 'rotation' | 'flipX' | 'flipY' | 'crop'>
  axes?: SheetImageAxes
  originalImage?: SheetFloatingImage
}
export function createSheetImageEditorDraft(image: SheetCellImage | SheetFloatingImage, row = 0, column = 0, axes?: SheetImageAxes): SheetImageEditorDraft {
  const floating = 'anchor' in image, frame = floating && axes ? sheetImageFrame(image, axes) : undefined,
    width = frame?.width ?? (floating ? image.size.width : image.resource.width), height = frame?.height ?? (floating ? image.size.height : image.resource.height),
    crop = floating && image.crop ? { left: image.crop.left * width / image.size.width, right: image.crop.right * width / image.size.width,
      top: image.crop.top * height / image.size.height, bottom: image.crop.bottom * height / image.size.height } : { top: 0, left: 0, bottom: 0, right: 0 }
  const draft: SheetImageEditorDraft = { floating, row: (floating ? image.anchor.row : row) + 1, column: (floating ? image.anchor.column : column) + 1,
    offsetX: floating ? image.anchor.offsetX : 0, offsetY: floating ? image.anchor.offsetY : 0, width, height, ratio: width / height,
    keepRatio: floating ? image.keepRatio : true, opacity: floating ? image.opacity * 100 : 100, title: image.title, altText: image.altText,
    anchorType: floating ? image.anchorType ?? '0' : '0', left: frame?.left ?? (floating ? image.absolutePosition?.left ?? 0 : 0),
    top: frame?.top ?? (floating ? image.absolutePosition?.top ?? 0 : 0), rotation: floating ? image.rotation ?? 0 : 0,
    flipX: floating ? image.flipX ?? false : false, flipY: floating ? image.flipY ?? false : false, crop, axes, originalImage: floating ? cloneModel(image) : undefined }
  draft.initial = { row: draft.row, column: draft.column, offsetX: draft.offsetX, offsetY: draft.offsetY, width, height,
    anchorType: draft.anchorType, left: draft.left, top: draft.top, rotation: draft.rotation, flipX: draft.flipX, flipY: draft.flipY, crop: { ...crop } }
  return draft
}
export function resizeSheetImageEditorDraft(draft: SheetImageEditorDraft, axis: 'width' | 'height', value: unknown): void {
  const size = sheetImageNumber(value, axis === 'width' ? '图片宽度' : '图片高度', 1, 100_000)
  if (draft.keepRatio) {
    const other = axis === 'width' ? size / draft.ratio : size * draft.ratio
    sheetImageNumber(other, '图片尺寸', 1, 100_000)
    draft[axis === 'width' ? 'height' : 'width'] = other
  }
  draft[axis] = size
}
export function sheetImagePatchFromDraft(draft: SheetImageEditorDraft): NonNullable<SheetImageCommandInput['patch']> {
  const patch: NonNullable<SheetImageCommandInput['patch']> = { title: sheetImageText(draft.title, '图片标题'), altText: sheetImageText(draft.altText, '替代文本') }
  if (draft.floating) {
    const anchor = { row: sheetImageIndex(draft.row - 1, 'row'), column: sheetImageIndex(draft.column - 1, 'column'),
      offsetX: sheetImageNumber(draft.offsetX, '水平偏移', -Number.MAX_SAFE_INTEGER), offsetY: sheetImageNumber(draft.offsetY, '垂直偏移', -Number.MAX_SAFE_INTEGER) },
      size = { width: sheetImageNumber(draft.width, '图片宽度', 1, 100_000), height: sheetImageNumber(draft.height, '图片高度', 1, 100_000) }, initial = draft.initial,
      crop = { top: sheetImageNumber(draft.crop.top, '顶部裁剪', -100_000, 100_000), left: sheetImageNumber(draft.crop.left, '左侧裁剪', -100_000, 100_000),
        bottom: sheetImageNumber(draft.crop.bottom, '底部裁剪', -100_000, 100_000), right: sheetImageNumber(draft.crop.right, '右侧裁剪', -100_000, 100_000) }
    if (size.width + crop.left + crop.right <= 0 || size.height + crop.top + crop.bottom <= 0) throw new RangeError('裁剪后的原图绘制尺寸必须大于零。')
    if (!initial || (['row', 'column', 'offsetX', 'offsetY'] as const).some(key => draft[key] !== initial[key])) patch.anchor = anchor
    const cropChanged = !initial || (['top', 'left', 'bottom', 'right'] as const).some(key => crop[key] !== initial.crop[key])
    if (!initial || size.width !== initial.width || size.height !== initial.height || cropChanged) patch.size = size
    if (cropChanged) patch.crop = crop
    if (!initial || draft.anchorType !== initial.anchorType) patch.anchorType = draft.anchorType
    if (draft.anchorType === '2' && (!initial || initial.anchorType !== '2' || draft.left !== initial.left || draft.top !== initial.top))
      patch.absolutePosition = { left: sheetImageNumber(draft.left, '图片水平位置', -Number.MAX_SAFE_INTEGER), top: sheetImageNumber(draft.top, '图片垂直位置', -Number.MAX_SAFE_INTEGER) }
    if (!initial || draft.rotation !== initial.rotation) patch.rotation = sheetImageNumber(draft.rotation, '图片旋转角度', -Number.MAX_SAFE_INTEGER)
    if (!initial || draft.flipX !== initial.flipX) patch.flipX = draft.flipX
    if (!initial || draft.flipY !== initial.flipY) patch.flipY = draft.flipY
    patch.keepRatio = draft.keepRatio; patch.opacity = sheetImageNumber(draft.opacity, '不透明度', 0, 100) / 100
  }
  return patch
}
export function applySheetImageEditorCropRatio(draft: SheetImageEditorDraft, ratio: number): void {
  if (!draft.floating || !draft.originalImage || !draft.axes) throw new Error('比例裁剪需要目标工作表的尺寸索引。')
  const current = patchSheetFloatingImage(draft.originalImage, sheetImagePatchFromDraft(draft), draft.axes), next = cropSheetImageToRatio(current, draft.axes, ratio), frame = sheetImageFrame(next, draft.axes)
  draft.row = next.anchor.row + 1; draft.column = next.anchor.column + 1; draft.offsetX = next.anchor.offsetX; draft.offsetY = next.anchor.offsetY
  draft.width = frame.width; draft.height = frame.height; draft.ratio = frame.width / frame.height; draft.left = frame.left; draft.top = frame.top; draft.crop = { ...next.crop! }
}
