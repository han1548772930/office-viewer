import type { DocumentImage } from '../features/document/model/document-model'
import { bodyNumber, bodyRecord, type UniverBodyRecord } from './document-body-style'
import { cloneModel } from '../domain/serialization'
import { documentImageCropFromSourceRect, documentImageSourceRectFromCrop, normalizeDocumentImageCrop,
  type DocumentImagePosition, type DocumentImageWrappingStyle } from '../features/document/images/document-image-model'

const layouts: Record<DocumentImageWrappingStyle, number> = { inline: 0, behindText: 1, inFrontOfText: 1, wrapSquare: 3, wrapTopAndBottom: 6 }
const wrapSides = ['both', 'left', 'right', 'largest'] as const
function position(raw: unknown): DocumentImagePosition | undefined {
  const data = bodyRecord(raw), relative = bodyNumber(data.relativeFrom)
  if (relative === undefined || !Number.isInteger(relative) || relative < 0 || relative > 7) return undefined
  const align = bodyNumber(data.align), percent = bodyNumber(data.percent), offset = bodyNumber(data.posOffset)
  return { relativeFrom: relative, ...(align === undefined ? percent === undefined ? { posOffset: offset ?? 0 } : { percent } : { align }) }
}
function crop(raw: unknown): DocumentImage['crop'] {
  const value = bodyRecord(raw), top = bodyNumber(value.top), right = bodyNumber(value.right), bottom = bodyNumber(value.bottom), left = bodyNumber(value.left)
  if (top === undefined || right === undefined || bottom === undefined || left === undefined) return undefined
  try { return normalizeDocumentImageCrop({ top, right, bottom, left }) } catch { return undefined }
}
function sourceCrop(raw: unknown, width: number, height: number): DocumentImage['crop'] {
  if (raw === undefined) return undefined
  if (raw === null) return { top: 0, right: 0, bottom: 0, left: 0 }
  const value = bodyRecord(raw)
  if (!Object.keys(value).length) return undefined
  if (!['top', 'right', 'bottom', 'left'].some(key => value[key] !== undefined)) return undefined
  const top = bodyNumber(value.top), right = bodyNumber(value.right), bottom = bodyNumber(value.bottom), left = bodyNumber(value.left)
  if (['top', 'right', 'bottom', 'left'].some(key => value[key] !== undefined && bodyNumber(value[key]) === undefined)) return undefined
  // Unknown/invalid source records remain in univerBodyDrawingSource for later adapters.
  try { return documentImageCropFromSourceRect({ top, right, bottom, left }, width, height) } catch { return undefined }
}
export function documentImageFromDrawing(drawing: UniverBodyRecord, source: string, width: number, height: number, layer: number): DocumentImage {
  const transform = bodyRecord(drawing.docTransform), extension = bodyRecord(drawing.officeStudioImage)
  const layout = bodyNumber(drawing.layoutType), wrappingStyle = layout === 0 || layout === undefined ? 'inline'
    : layout === 6 ? 'wrapTopAndBottom' : layout === 1 ? drawing.behindDoc === 1 ? 'behindText' : 'inFrontOfText' : 'wrapSquare'
  const intrinsicWidth = bodyNumber(extension.intrinsicWidth), intrinsicHeight = bodyNumber(extension.intrinsicHeight)
  const imageCrop = crop(extension.crop) ?? sourceCrop(drawing.srcRect, width, height)
  return { src: source, alt: typeof drawing.description === 'string' ? drawing.description : '', width, height, wrappingStyle,
    angle: bodyNumber(transform.angle) ?? 0, keepAspectRatio: extension.keepAspectRatio !== false, positionH: position(transform.positionH), positionV: position(transform.positionV),
    ...(typeof transform.flipX === 'boolean' ? { flipX: transform.flipX } : {}), ...(typeof transform.flipY === 'boolean' ? { flipY: transform.flipY } : {}),
    wrapSide: wrapSides[bodyNumber(drawing.wrapText) ?? 0] ?? 'both', layer,
    distances: { top: bodyNumber(drawing.distT) ?? 0, right: bodyNumber(drawing.distR) ?? 0, bottom: bodyNumber(drawing.distB) ?? 0, left: bodyNumber(drawing.distL) ?? 0 },
    ...(intrinsicWidth ? { intrinsicWidth } : {}), ...(intrinsicHeight ? { intrinsicHeight } : {}),
    ...(typeof extension.resourceId === 'string' ? { resourceId: extension.resourceId } : {}),
    ...(typeof extension.originalSource === 'string' ? { originalSource: extension.originalSource } : {}),
    ...(typeof extension.mimeType === 'string' ? { mimeType: extension.mimeType } : {}),
    ...(bodyNumber(extension.byteLength) ? { byteLength: bodyNumber(extension.byteLength) } : {}),
    ...(imageCrop ? { crop: imageCrop } : {}) }
}
export function documentImageToDrawing(image: DocumentImage, drawing: UniverBodyRecord): UniverBodyRecord {
  const style = image.wrappingStyle ?? 'inline'
  const sourceSize = bodyRecord(bodyRecord(drawing.docTransform).size), originalCrop = sourceCrop(drawing.srcRect, image.width, image.height)
  const computedRect = image.crop ? documentImageSourceRectFromCrop(image.crop, image.width, image.height) : undefined
  const unchangedCrop = originalCrop && image.crop && JSON.stringify(originalCrop) === JSON.stringify(image.crop)
    && bodyNumber(sourceSize.width) === image.width && bodyNumber(sourceSize.height) === image.height
  const srcRect = unchangedCrop ? cloneModel(drawing.srcRect) : computedRect === null ? null : computedRect ? { ...bodyRecord(drawing.srcRect), ...computedRect } : undefined
  const outputPosition = (value: DocumentImagePosition, raw: unknown) => {
    const retained = { ...bodyRecord(raw) }
    for (const field of ['align', 'percent', 'posOffset']) delete retained[field]
    return { ...retained, ...cloneModel(value) }
  }
  return { ...drawing, source: image.src, imageSourceType: /^data:/i.test(image.src) ? 'BASE64' : drawing.imageSourceType ?? 'URL', description: image.alt,
    ...(image.crop ? { srcRect } : {}),
    // Preserve an imported polygon/tight/through layout until the author changes its wrapping semantics.
    layoutType: style === 'wrapSquare' && [2, 4, 5].includes(bodyNumber(drawing.layoutType) ?? -1) ? drawing.layoutType : layouts[style],
    behindDoc: style === 'behindText' ? 1 : 0, wrapText: wrapSides.indexOf(image.wrapSide ?? 'both'),
    distT: image.distances?.top ?? 0, distR: image.distances?.right ?? 0, distB: image.distances?.bottom ?? 0, distL: image.distances?.left ?? 0,
    docTransform: { ...bodyRecord(drawing.docTransform), size: { ...bodyRecord(bodyRecord(drawing.docTransform).size), width: image.width, height: image.height },
      angle: image.angle ?? 0, ...(image.flipX !== undefined ? { flipX: image.flipX } : {}), ...(image.flipY !== undefined ? { flipY: image.flipY } : {}),
      ...(image.positionH ? { positionH: outputPosition(image.positionH, bodyRecord(drawing.docTransform).positionH) } : {}),
      ...(image.positionV ? { positionV: outputPosition(image.positionV, bodyRecord(drawing.docTransform).positionV) } : {}) },
    officeStudioImage: { ...bodyRecord(drawing.officeStudioImage), resourceId: image.resourceId, intrinsicWidth: image.intrinsicWidth,
      intrinsicHeight: image.intrinsicHeight, originalSource: image.originalSource, keepAspectRatio: image.keepAspectRatio, crop: image.crop,
      byteLength: image.byteLength, mimeType: image.mimeType, layer: image.layer } }
}
