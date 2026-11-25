import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentBlockChildContainers } from '../model/document-structural-blocks'
import { imageForBlock, textTargets, type DocumentImage, type DocumentRun, type TextTarget } from '../model/document-model'
import { documentImageAtom, setDocumentImageAtom } from './document-image-atom'

export const DOCUMENT_IMAGE_WRAPPING_STYLES = ['inline', 'behindText', 'inFrontOfText', 'wrapSquare', 'wrapTopAndBottom'] as const
export type DocumentImageWrappingStyle = typeof DOCUMENT_IMAGE_WRAPPING_STYLES[number]
export type DocumentImageWrapSide = 'both' | 'left' | 'right' | 'largest'
export interface DocumentImageCrop { top: number; right: number; bottom: number; left: number }
/** Univer offsets are CSS pixels outside the visible frame, including signed outcropping. */
export interface DocumentImageSourceRect { top?: number; right?: number; bottom?: number; left?: number }
export interface DocumentImageDistances { top: number; right: number; bottom: number; left: number }
export interface DocumentImagePosition { relativeFrom: number; posOffset?: number; align?: number; percent?: number }
export type DocumentImagePatch = Partial<Omit<DocumentImage, 'src' | 'resourceId' | 'mimeType' | 'byteLength' | 'intrinsicWidth' | 'intrinsicHeight' | 'originalSource'>>
export type DocumentImageArrange = 'front' | 'forward' | 'back' | 'backward'
export interface DocumentImageEntry {
  imageId: string; block: DocumentBlock; image: DocumentImage; container: DocumentBlock[]
  inline?: { run: DocumentRun; target: TextTarget; startOffset: number; endOffset: number }
}

export function documentImageEntries(model: Pick<DocumentModel, 'blocks'>): DocumentImageEntry[] {
  const result: DocumentImageEntry[] = []
  const targets = new Map(textTargets(model as DocumentModel).map(target => [target.cellBlockId ?? target.blockId, target]))
  const visit = (blocks: DocumentBlock[]) => { for (const block of blocks) {
    const image = block.type === 'image' ? imageForBlock(block) : undefined
    if (image) result.push({ imageId: block.id, block, image, container: blocks })
    let offset = 0
    for (const run of block.runs as DocumentRun[]) {
      const atom = documentImageAtom(run)
      if (atom) result.push({ imageId: atom.imageId, block, image: atom.image, container: blocks,
        inline: { run, target: targets.get(block.id) ?? { blockId: block.id }, startOffset: offset, endOffset: offset + run.text.length } })
      offset += run.text.length
    }
    for (const content of documentBlockChildContainers(block)) visit(content)
  } }
  visit(model.blocks)
  return result
}

export function documentImageEntryMetadata(entry: DocumentImageEntry): Record<string, unknown> {
  return entry.inline ? entry.inline.run.metadata ?? {} : entry.block.metadata
}
export function setDocumentImageEntry(entry: DocumentImageEntry, image: DocumentImage): void {
  if (entry.inline) setDocumentImageAtom(entry.inline.run, entry.imageId, image)
  else entry.block.metadata.image = image
  entry.image = image
}

/** Adjacent moves keep selected objects in their existing stacking order. */
export function arrangeDocumentImageOrder(order: readonly string[], imageIds: readonly string[], arrangement: DocumentImageArrange | number): string[] {
  if (!Array.isArray(imageIds) || !imageIds.length || imageIds.length > 1000
    || imageIds.some(id => typeof id !== 'string' || !id.trim()) || new Set(imageIds).size !== imageIds.length
    || imageIds.some(id => !order.includes(id))) throw new TypeError('图片排列目标无效。')
  const selected = new Set(imageIds), result = [...order]
  if (typeof arrangement === 'number') {
    if (imageIds.length !== 1) throw new RangeError('绝对层级只能指定一张图片。')
    const current = result.indexOf(imageIds[0]!), target = Math.max(0, Math.min(Math.floor(arrangement), result.length - 1)) || 0
    result.splice(target, 0, result.splice(current, 1)[0]!)
    return result
  }
  if (!['front', 'forward', 'back', 'backward'].includes(arrangement)) throw new RangeError('图片层级动作无效。')
  if (arrangement === 'front' || arrangement === 'back') {
    const picked = result.filter(id => selected.has(id)), others = result.filter(id => !selected.has(id))
    return arrangement === 'front' ? [...others, ...picked] : [...picked, ...others]
  }
  const direction = arrangement === 'forward' ? 1 : -1
  const indices = Array.from({ length: Math.max(0, result.length - 1) }, (_, index) => direction === 1 ? result.length - 2 - index : index + 1)
  for (const index of indices) if (selected.has(result[index]!) && !selected.has(result[index + direction]!)) {
    const adjacent = result[index + direction]!
    result[index + direction] = result[index]!; result[index] = adjacent
  }
  return result
}

function finite(value: unknown, label: string, minimum: number, maximum: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < minimum || value > maximum) throw new RangeError(`${label}无效。`)
  return value
}
function box(value: unknown, label: string, maximum: number, minimum = 0): DocumentImageCrop {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label}无效。`)
  const fields = ['top', 'right', 'bottom', 'left'] as const
  if (Object.keys(value).some(key => !fields.some(field => field === key))) throw new TypeError(`${label}字段无效。`)
  return { top: finite(Reflect.get(value, 'top'), label, minimum, maximum), right: finite(Reflect.get(value, 'right'), label, minimum, maximum),
    bottom: finite(Reflect.get(value, 'bottom'), label, minimum, maximum), left: finite(Reflect.get(value, 'left'), label, minimum, maximum) }
}
export function documentImageCropFromSourceRect(value: DocumentImageSourceRect | null, width: number, height: number): DocumentImageCrop {
  finite(width, '图片宽度', 0.1, 100_000); finite(height, '图片高度', 0.1, 100_000)
  if (value === null) return { top: 0, right: 0, bottom: 0, left: 0 }
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.keys(value).some(key => !['top', 'right', 'bottom', 'left'].includes(key))) throw new TypeError('图片原始裁剪框无效。')
  const offsets = { top: finite(value.top ?? 0, '图片裁剪偏移', -100_000, 100_000), right: finite(value.right ?? 0, '图片裁剪偏移', -100_000, 100_000),
    bottom: finite(value.bottom ?? 0, '图片裁剪偏移', -100_000, 100_000), left: finite(value.left ?? 0, '图片裁剪偏移', -100_000, 100_000) }
  const fullWidth = width + offsets.left + offsets.right, fullHeight = height + offsets.top + offsets.bottom
  if (fullWidth <= 0 || fullHeight <= 0) throw new RangeError('图片裁剪前的原始框必须大于零。')
  return normalizeDocumentImageCrop({ top: offsets.top / fullHeight, right: offsets.right / fullWidth, bottom: offsets.bottom / fullHeight, left: offsets.left / fullWidth })
}
export function normalizeDocumentImageCrop(value: DocumentImageCrop): DocumentImageCrop {
  const result = box(value, '图片裁剪比例', 0.999999, -1000)
  if (1 - result.left - result.right < 0.000001 || 1 - result.top - result.bottom < 0.000001) throw new RangeError('裁剪后必须保留图片内容。')
  return result
}
export function documentImageSourceRectFromCrop(value: DocumentImageCrop, width: number, height: number): DocumentImageSourceRect | null {
  const crop = normalizeDocumentImageCrop(value)
  if (!crop.top && !crop.right && !crop.bottom && !crop.left) return null
  const fullWidth = width / (1 - crop.left - crop.right), fullHeight = height / (1 - crop.top - crop.bottom)
  return { top: crop.top * fullHeight, right: crop.right * fullWidth, bottom: crop.bottom * fullHeight, left: crop.left * fullWidth }
}
export function normalizeDocumentImagePosition(value: DocumentImagePosition, vertical = false): DocumentImagePosition {
  if (!value || typeof value !== 'object' || Array.isArray(value)
    || Object.keys(value).some(key => !['relativeFrom', 'posOffset', 'align', 'percent'].includes(key))) throw new TypeError('图片定位参数无效。')
  if (!Number.isInteger(value.relativeFrom) || value.relativeFrom < 0 || value.relativeFrom > 7) throw new RangeError('图片定位基准无效。')
  if (value.align !== undefined && (!Number.isInteger(value.align) || value.align < 0 || value.align > (vertical ? 4 : 6))) throw new RangeError('图片对齐参数无效。')
  if (value.posOffset !== undefined) finite(value.posOffset, '图片偏移', -100_000, 100_000)
  if (value.percent !== undefined) finite(value.percent, '图片定位百分比', -1000, 1000)
  if ([value.align, value.posOffset, value.percent].filter(field => field !== undefined).length > 1) throw new RangeError('图片定位同时指定多个定位值。')
  if (value.percent !== undefined) return { relativeFrom: value.relativeFrom, percent: value.percent }
  return { relativeFrom: value.relativeFrom, ...(value.align === undefined ? { posOffset: value.posOffset ?? 0 } : { align: value.align }) }
}
export function normalizeDocumentImage(image: DocumentImage): DocumentImage {
  if (!image || typeof image !== 'object' || Array.isArray(image) || typeof image.src !== 'string' || !image.src.trim()
    || image.src.length > 7_100_000 || !/^(?:https?:\/\/|blob:|data:image\/(?:png|jpe?g|gif|webp|bmp);base64,)/i.test(image.src.trim())) throw new TypeError('图片来源无效。')
  if (typeof image.alt !== 'string' || image.alt.length > 32_768) throw new RangeError('图片说明超过预算。')
  const width = finite(image.width, '图片宽度', 0.1, 100_000), height = finite(image.height, '图片高度', 0.1, 100_000)
  const result = { ...cloneModel(image), src: image.src.trim(), width, height }
  if (image.angle !== undefined) result.angle = ((finite(image.angle, '图片角度', -100_000, 100_000) % 360) + 360) % 360
  for (const key of ['flipX', 'flipY'] as const) if (image[key] !== undefined && typeof image[key] !== 'boolean') throw new TypeError('图片翻转值必须为布尔值。')
  if (image.keepAspectRatio !== undefined && typeof image.keepAspectRatio !== 'boolean') throw new TypeError('图片比例锁定值无效。')
  if (image.wrappingStyle !== undefined && !DOCUMENT_IMAGE_WRAPPING_STYLES.includes(image.wrappingStyle)) throw new RangeError('图片环绕方式无效。')
  if (image.wrapSide !== undefined && !['both', 'left', 'right', 'largest'].includes(image.wrapSide)) throw new RangeError('图片环绕方向无效。')
  if (image.positionH) result.positionH = normalizeDocumentImagePosition(image.positionH)
  if (image.positionV) result.positionV = normalizeDocumentImagePosition(image.positionV, true)
  if (image.crop) {
    result.crop = normalizeDocumentImageCrop(image.crop)
  }
  if (image.distances) result.distances = box(image.distances, '图片文字间距', 100_000)
  if (image.layer !== undefined && (!Number.isSafeInteger(image.layer) || Math.abs(image.layer) > 1_000_000)) throw new RangeError('图片层级无效。')
  for (const key of ['intrinsicWidth', 'intrinsicHeight', 'byteLength'] as const) if (image[key] !== undefined) finite(image[key], '图片资源尺寸', 1, key === 'byteLength' ? 5 * 1024 * 1024 : 100_000)
  return result
}

export function patchDocumentImage(image: DocumentImage, patch: DocumentImagePatch): DocumentImage {
  if (patch && typeof patch === 'object' && 'src' in patch) throw new TypeError('替换图片来源必须经过异步资源读取。')
  if (!patch || typeof patch !== 'object' || Array.isArray(patch) || Object.keys(patch).some(key => ![
    'alt', 'width', 'height', 'angle', 'flipX', 'flipY', 'keepAspectRatio', 'wrappingStyle', 'wrapSide', 'positionH', 'positionV', 'crop', 'distances', 'layer',
  ].includes(key))) throw new TypeError('图片修改字段无效。')
  const next = { ...image, ...cloneModel(patch) }
  if ((patch.keepAspectRatio ?? image.keepAspectRatio) !== false) {
    if (patch.width !== undefined && patch.height === undefined) next.height = patch.width * image.height / image.width
    else if (patch.height !== undefined && patch.width === undefined) next.width = patch.height * image.width / image.height
  }
  return normalizeDocumentImage(next)
}
