import type { SheetBackgroundImage, SheetCellImage, SheetFloatingImage, SheetImageResource } from '../../domain/sheet-images.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { normalizeDrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { decodeSheetImageDataUrl, SHEET_IMAGE_MAX_FILE_BYTES, SHEET_IMAGE_MAX_PIXELS } from './sheet-image-bytes.ts'
export { SHEET_IMAGE_MAX_FILE_BYTES, SHEET_IMAGE_MAX_PIXELS } from './sheet-image-bytes.ts'

export const SHEET_IMAGE_MIME_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp'] as const
export const SHEET_IMAGE_ACCEPT = '.png,.jpg,.jpeg,.gif,.bmp,.webp'
export const SHEET_IMAGE_MAX_BATCH_BYTES = 32 * 1024 * 1024
export const SHEET_IMAGE_MAX_COUNT = 1000

export function sheetImageRecord(value: unknown, label = '图片'): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label}参数必须是对象。`)
  return value as Record<string, unknown>
}
export function sheetImageNumber(value: unknown, label: string, min: number, max = Number.MAX_SAFE_INTEGER): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) throw new RangeError(`${label}无效。`)
  return value
}
export function sheetImageIndex(value: unknown, axis: 'row' | 'column'): number {
  const index = sheetImageNumber(value, axis === 'row' ? '行号' : '列号', 0, axis === 'row' ? 1_048_575 : 16_383)
  if (!Number.isInteger(index)) throw new RangeError('图片锚点必须是整数行列。')
  return index
}
export function sheetImageText(value: unknown, label: string, fallback = ''): string {
  if (value === undefined) return fallback
  if (typeof value !== 'string' || value.length > 4096) throw new TypeError(`${label}必须是最多 4096 个字符的文本。`)
  return value
}
export function sheetImageId(value: unknown): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 255) throw new TypeError('图片 ID 无效。')
  return value
}
export function normalizeSheetImageResource(value: unknown): SheetImageResource {
  const input = sheetImageRecord(value, '图片资源'), id = sheetImageId(input.id)
  if (!SHEET_IMAGE_MIME_TYPES.includes(input.mimeType as SheetImageResource['mimeType'])) throw new TypeError('请选择 PNG、JPEG、GIF、BMP 或 WebP 图片。')
  const source = input.source
  if (typeof source !== 'string' || !/^(?:data:image\/(?:png|jpeg|jpg|gif|bmp|webp);base64,|https?:\/\/)/i.test(source)) throw new TypeError('图片来源必须是受支持的 data URL 或 HTTP 图片地址。')
  if (source.length > Math.ceil(SHEET_IMAGE_MAX_FILE_BYTES / 3) * 4 + 128) throw new RangeError('图片超过 5 MiB 限制。')
  const width = sheetImageNumber(input.width, '图片宽度', 1), height = sheetImageNumber(input.height, '图片高度', 1)
  if (width * height > SHEET_IMAGE_MAX_PIXELS) throw new RangeError('图片解码尺寸超过内存预算。')
  const byteLength = sheetImageNumber(input.byteLength, '图片字节数', 0, SHEET_IMAGE_MAX_FILE_BYTES)
  if (!Number.isSafeInteger(byteLength)) throw new RangeError('图片字节数必须是整数。')
  if (/^data:/i.test(source)) {
    const actual = decodeSheetImageDataUrl(source)
    if (actual.bytes.byteLength !== byteLength || actual.header.width !== width || actual.header.height !== height || actual.header.mimeType !== input.mimeType) {
      throw new TypeError('图片尺寸、字节数或 MIME 类型与实际内容不一致。')
    }
  }
  const result: SheetImageResource = { ...cloneModel(input), id, name: sheetImageText(input.name, '文件名', 'image'),
    mimeType: input.mimeType as SheetImageResource['mimeType'], source, width, height, byteLength }
  if (input.scaleX !== undefined) result.scaleX = sheetImageNumber(input.scaleX, '水平分辨率比例', 0.001, 1000)
  if (input.scaleY !== undefined) result.scaleY = sheetImageNumber(input.scaleY, '垂直分辨率比例', 0.001, 1000)
  if (input.metadata !== undefined) result.metadata = cloneModel(sheetImageRecord(input.metadata, '图片扩展'))
  return result
}
export function normalizeSheetCellImage(value: unknown): SheetCellImage {
  const input = sheetImageRecord(value)
  const result: SheetCellImage = { ...cloneModel(input), id: sheetImageId(input.id), resource: normalizeSheetImageResource(input.resource),
    altText: sheetImageText(input.altText, '替代文本'), title: sheetImageText(input.title, '图片标题') }
  if (input.metadata !== undefined) result.metadata = cloneModel(sheetImageRecord(input.metadata, '图片扩展'))
  return result
}
export function normalizeSheetFloatingImage(value: unknown): SheetFloatingImage {
  const input = sheetImageRecord(value), anchor = sheetImageRecord(input.anchor, '图片锚点'), size = sheetImageRecord(input.size, '图片尺寸')
  if (typeof input.keepRatio !== 'boolean') throw new TypeError('锁定纵横比必须是布尔值。')
  const result: SheetFloatingImage = { ...normalizeSheetCellImage(input), anchor: { ...cloneModel(anchor), row: sheetImageIndex(anchor.row, 'row'), column: sheetImageIndex(anchor.column, 'column'),
    offsetX: sheetImageNumber(anchor.offsetX, '水平偏移', -Number.MAX_SAFE_INTEGER), offsetY: sheetImageNumber(anchor.offsetY, '垂直偏移', -Number.MAX_SAFE_INTEGER) },
    size: { ...cloneModel(size), width: sheetImageNumber(size.width, '图片宽度', 1, 100_000), height: sheetImageNumber(size.height, '图片高度', 1, 100_000) },
    opacity: sheetImageNumber(input.opacity, '图片不透明度', 0, 1), keepRatio: input.keepRatio }
  const kind = input.anchorType ?? '0'
  if (kind !== '0' && kind !== '1' && kind !== '2') throw new TypeError('图片锚点类型无效。')
  result.anchorType = kind
  delete result.endAnchor; delete result.absolutePosition
  if (kind === '1') {
    const end = sheetImageRecord(input.endAnchor, '图片结束锚点')
    result.endAnchor = { row: sheetImageIndex(end.row, 'row'), column: sheetImageIndex(end.column, 'column'),
      offsetX: sheetImageNumber(end.offsetX, '图片结束水平偏移', -Number.MAX_SAFE_INTEGER), offsetY: sheetImageNumber(end.offsetY, '图片结束垂直偏移', -Number.MAX_SAFE_INTEGER) }
  }
  if (kind === '2') {
    const position = sheetImageRecord(input.absolutePosition, '图片绝对位置')
    result.absolutePosition = { left: sheetImageNumber(position.left, '图片水平位置', -Number.MAX_SAFE_INTEGER),
      top: sheetImageNumber(position.top, '图片垂直位置', -Number.MAX_SAFE_INTEGER) }
  }
  if (input.rotation !== undefined) result.rotation = sheetImageNumber(input.rotation, '图片旋转角度', -Number.MAX_SAFE_INTEGER)
  if (input.drawingGeometry !== undefined) result.drawingGeometry = normalizeDrawingAffineGeometry(input.drawingGeometry)
  for (const key of ['flipX', 'flipY'] as const) if (input[key] !== undefined) {
    if (typeof input[key] !== 'boolean') throw new TypeError('图片翻转参数必须是布尔值。')
    result[key] = input[key]
  }
  if (input.crop !== undefined) {
    const crop = sheetImageRecord(input.crop, '图片裁剪')
    result.crop = { top: sheetImageNumber(crop.top, '顶部裁剪', -100_000, 100_000), left: sheetImageNumber(crop.left, '左侧裁剪', -100_000, 100_000),
      bottom: sheetImageNumber(crop.bottom, '底部裁剪', -100_000, 100_000), right: sheetImageNumber(crop.right, '右侧裁剪', -100_000, 100_000) }
    if (result.size.width + result.crop.left + result.crop.right <= 0 || result.size.height + result.crop.top + result.crop.bottom <= 0) throw new RangeError('裁剪后的原图绘制尺寸必须大于零。')
  }
  return result
}
export function normalizeSheetBackgroundImage(value: unknown): SheetBackgroundImage {
  const input = sheetImageRecord(value, '背景图')
  if (input.resource === undefined) {
    if (typeof input.source !== 'string' || !input.source || input.imageSourceType !== 'URL' && input.imageSourceType !== 'UUID'
      || input.imageSourceType === 'URL' && !/^https?:\/\//i.test(input.source) || input.source.length > 16_384) throw new TypeError('背景图片来源无效。')
    return { source: input.source, imageSourceType: input.imageSourceType,
      scaleX: sheetImageNumber(input.scaleX ?? 1, '背景图水平比例', 0.001, 1000),
      scaleY: sheetImageNumber(input.scaleY ?? 1, '背景图垂直比例', 0.001, 1000),
      ...(input.metadata === undefined ? {} : { metadata: cloneModel(sheetImageRecord(input.metadata, '背景图扩展')) }) }
  }
  const resource = normalizeSheetImageResource(input.resource)
  return { ...cloneModel(input), resource,
    scaleX: sheetImageNumber(input.scaleX ?? resource.scaleX ?? 1, '背景图水平比例', 0.001, 1000),
    scaleY: sheetImageNumber(input.scaleY ?? resource.scaleY ?? 1, '背景图垂直比例', 0.001, 1000) }
}
export function createSheetFloatingImage(resource: SheetImageResource, row: number, column: number, index = 0): SheetFloatingImage {
  const normalized = normalizeSheetImageResource(resource), scale = Math.min(1, 500 / normalized.width, 500 / normalized.height)
  return normalizeSheetFloatingImage({ id: crypto.randomUUID(), resource: normalized, title: normalized.name, altText: '',
    anchor: { row, column, offsetX: 12 * index, offsetY: 12 * index },
    size: { width: Math.max(1, normalized.width * scale), height: Math.max(1, normalized.height * scale) }, keepRatio: true, opacity: 1 })
}
export function copySheetCellImage(image: SheetCellImage): SheetCellImage { return { ...cloneModel(image), id: crypto.randomUUID() } }
