import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import type { DocumentImage } from '../model/document-model'
import type { DocumentLayout, LayoutImage } from '../layout/document-layout-types'
import type { DocumentImageFrame } from '../layout/document-image-layout'
import { documentImageEntries, patchDocumentImage, type DocumentImagePatch } from './document-image-model'

/** The values and ordering match the fixed drawing-ui AlignType public enum. */
export const DOCUMENT_IMAGE_ALIGN_TYPES = { default: '0', left: '1', center: '2', right: '3', top: '4', middle: '5', bottom: '6', horizon: '7', vertical: '8' } as const
export type DocumentImageAlign = typeof DOCUMENT_IMAGE_ALIGN_TYPES[keyof typeof DOCUMENT_IMAGE_ALIGN_TYPES]
export interface DocumentImageBatchMember { imageId: string; image: DocumentImage; rectangle: LayoutImage; pageOrigin: { x: number; y: number } }
export interface DocumentImageBatchTransform { x?: number; y?: number; scale?: number; angle?: number; origin?: { x: number; y: number } }
export interface DocumentImageBatchUpdate { imageId: string; patch: DocumentImagePatch; rectangle: DocumentImageFrame }

function checkedIds(imageIds: readonly string[]): string[] {
  if (!Array.isArray(imageIds) || !imageIds.length || imageIds.length > 1000 || imageIds.some(id => typeof id !== 'string' || !id.trim())
    || new Set(imageIds).size !== imageIds.length) throw new TypeError('图片选择需要不重复的有效对象身份。')
  return [...imageIds]
}
function number(value: unknown, label: string, minimum = -100_000, maximum = 100_000): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < minimum || value > maximum) throw new RangeError(`${label}无效。`)
  return value
}
export function documentImageBatchMembers(model: DocumentModel, layout: DocumentLayout, imageIds: readonly string[], preferredPage?: number): DocumentImageBatchMember[] {
  if (preferredPage !== undefined && (!Number.isSafeInteger(preferredPage) || preferredPage < 0)) throw new RangeError('图片所在页无效。')
  const entries = documentImageEntries(model), ids = checkedIds(imageIds)
  return ids.map(imageId => {
    const entry = entries.find(entry => entry.imageId === imageId)
    const rectangles = layout.images.filter(rectangle => rectangle.objectImage && rectangle.blockId === imageId)
    const rectangle = rectangles.find(rectangle => rectangle.page === preferredPage) ?? rectangles.find(rectangle => !rectangle.repeatedHeader) ?? rectangles[0]
    if (!entry || !rectangle) throw new RangeError('图片尚无当前排版位置，请等待排版或重新选择。')
    const page = layout.pages[rectangle.page]
    if (!page || page.isParityBlank) throw new RangeError('图片不在可编辑的文档页中。')
    return { imageId, image: cloneModel(entry.image), rectangle: cloneModel(rectangle), pageOrigin: { x: page.x, y: page.y } }
  })
}

/** These are derived layout snapshots, never a second persistent image model. */
export function assertDocumentImageBatchMembers(model: DocumentModel, members: readonly DocumentImageBatchMember[]): void {
  if (!Array.isArray(members)) throw new TypeError('图片排版快照无效。')
  checkedIds(members.map(member => member.imageId))
  const entries = documentImageEntries(model)
  const first = entries.find(entry => entry.imageId === members[0]!.imageId)
  for (const member of members) {
    const entry = entries.find(entry => entry.imageId === member.imageId), rectangle = member.rectangle
    if (!entry || JSON.stringify(entry.image) !== JSON.stringify(member.image)) throw new RangeError('图片内容或排版输入已变化，请重新选择。')
    if (entry.container !== first?.container || rectangle.page !== members[0]!.rectangle.page
      || member.pageOrigin.x !== members[0]!.pageOrigin.x || member.pageOrigin.y !== members[0]!.pageOrigin.y) throw new RangeError('图片几何操作需要同一内容容器和同一页中的对象。')
    if (!rectangle.objectImage || rectangle.blockId !== member.imageId) throw new TypeError('图片排版身份无效。')
    if (!Number.isSafeInteger(rectangle.page) || rectangle.page < 0) throw new TypeError('图片排版页无效。')
    number(rectangle.x, '图片横向位置'); number(rectangle.y, '图片纵向位置', -100_000, 100_000_000)
    number(rectangle.width, '图片排版宽度', 0.000001, 200_000); number(rectangle.height, '图片排版高度', 0.000001, 200_000)
    if (rectangle.paintWidth !== undefined) number(rectangle.paintWidth, '图片绘制宽度', 0.000001, 100_000)
    if (rectangle.paintHeight !== undefined) number(rectangle.paintHeight, '图片绘制高度', 0.000001, 100_000)
    number(rectangle.angle ?? 0, '图片排版角度'); number(member.image.width, '图片宽度', 0.1, 100_000); number(member.image.height, '图片高度', 0.1, 100_000)
    number(member.pageOrigin.x, '图片所属页横向位置'); number(member.pageOrigin.y, '图片所属页纵向位置', -100_000, 100_000_000)
  }
}
export function documentImageBatchBounds(members: readonly DocumentImageBatchMember[]): DocumentImageFrame {
  if (!members.length) throw new RangeError('图片选择为空。')
  const x = Math.min(...members.map(member => member.rectangle.x)), y = Math.min(...members.map(member => member.rectangle.y))
  return { x, y, width: Math.max(...members.map(member => member.rectangle.x + member.rectangle.width)) - x,
    height: Math.max(...members.map(member => member.rectangle.y + member.rectangle.height)) - y, angle: 0 }
}
function positioned(member: DocumentImageBatchMember, rectangle: DocumentImageFrame, extra: DocumentImagePatch = {}): DocumentImageBatchUpdate {
  const inline = (member.image.wrappingStyle ?? 'inline') === 'inline'
  const patch: DocumentImagePatch = { ...(inline ? { width: member.rectangle.paintWidth ?? member.image.width, height: member.rectangle.paintHeight ?? member.image.height } : {}), ...extra,
    wrappingStyle: (member.image.wrappingStyle ?? 'inline') === 'inline' ? 'inFrontOfText' : member.image.wrappingStyle,
    positionH: { relativeFrom: 0, posOffset: rectangle.x - member.pageOrigin.x },
    positionV: { relativeFrom: 0, posOffset: rectangle.y - member.pageOrigin.y } }
  patchDocumentImage(member.image, patch)
  return { imageId: member.imageId, rectangle, patch }
}

/** Uniform multi-selection scaling and rotation preserve each image's exact local frame. */
export function transformDocumentImageBatch(members: readonly DocumentImageBatchMember[], transform: DocumentImageBatchTransform): DocumentImageBatchUpdate[] {
  checkedIds(members.map(member => member.imageId))
  if (!transform || typeof transform !== 'object' || Array.isArray(transform)
    || Object.keys(transform).some(key => !['x', 'y', 'scale', 'angle', 'origin'].includes(key))) throw new TypeError('图片批量变换参数无效。')
  const dx = number(transform.x ?? 0, '图片横向位移'), dy = number(transform.y ?? 0, '图片纵向位移')
  const scale = number(transform.scale ?? 1, '图片缩放比例', 0.000001, 1_000_000), angle = number(transform.angle ?? 0, '图片旋转角度') % 360
  const bounds = documentImageBatchBounds(members), origin = transform.origin ?? { x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 }
  if (!origin || typeof origin !== 'object' || Array.isArray(origin) || Object.keys(origin).some(key => key !== 'x' && key !== 'y')) throw new TypeError('图片变换中心无效。')
  // A visible rotated box can extend beyond its valid top-left coordinate.
  number(origin.x, '图片变换中心', -100_000, 300_000); number(origin.y, '图片变换中心', -100_000, 100_200_000)
  if (!dx && !dy && scale === 1 && !angle) return []
  const radians = angle * Math.PI / 180, cosine = Math.cos(radians), sine = Math.sin(radians)
  return members.map(member => {
    const old = member.rectangle, localX = (old.x + old.width / 2 - origin.x) * scale, localY = (old.y + old.height / 2 - origin.y) * scale
    const centerX = origin.x + localX * cosine - localY * sine + dx, centerY = origin.y + localX * sine + localY * cosine + dy
    const nextAngle = ((member.image.angle ?? 0) + angle + 360) % 360, rotation = nextAngle * Math.PI / 180
    const paintWidth = (old.paintWidth ?? old.width) * scale, paintHeight = (old.paintHeight ?? old.height) * scale
    const width = Math.abs(Math.cos(rotation)) * paintWidth + Math.abs(Math.sin(rotation)) * paintHeight
    const height = Math.abs(Math.sin(rotation)) * paintWidth + Math.abs(Math.cos(rotation)) * paintHeight
    return positioned(member, { x: centerX - width / 2, y: centerY - height / 2, width, height, paintWidth, paintHeight, angle: nextAngle },
      { width: paintWidth, height: paintHeight, angle: nextAngle })
  })
}

/** Alignment uses visible rotated bounds; distribution uses their original span and permits overlap. */
export function alignDocumentImageBatch(members: readonly DocumentImageBatchMember[], align: DocumentImageAlign): DocumentImageBatchUpdate[] {
  checkedIds(members.map(member => member.imageId))
  if (!Object.values(DOCUMENT_IMAGE_ALIGN_TYPES).includes(align)) throw new TypeError('图片对齐类型无效。')
  if (align === '0') return []
  if (members.length < 2) throw new RangeError('图片对齐至少需要选择两张图片。')
  const bounds = documentImageBatchBounds(members), positions = new Map<string, { x: number; y: number }>()
  if (align === '7' || align === '8') {
    if (members.length < 3) return []
    const horizontal = align === '7', sorted = members.map((member, index) => ({ member, index }))
      .sort((a, b) => (horizontal ? a.member.rectangle.x - b.member.rectangle.x : a.member.rectangle.y - b.member.rectangle.y) || a.index - b.index)
    const start = horizontal ? bounds.x : bounds.y, end = horizontal ? bounds.x + bounds.width : bounds.y + bounds.height
    const total = sorted.reduce((value, entry) => value + (horizontal ? entry.member.rectangle.width : entry.member.rectangle.height), 0)
    const gap = (end - start - total) / (sorted.length - 1)
    let position = start
    for (const { member } of sorted) {
      positions.set(member.imageId, { x: horizontal ? position : member.rectangle.x, y: horizontal ? member.rectangle.y : position })
      position += (horizontal ? member.rectangle.width : member.rectangle.height) + gap
    }
  } else for (const member of members) {
    const old = member.rectangle
    positions.set(member.imageId, { x: align === '1' ? bounds.x : align === '2' ? bounds.x + (bounds.width - old.width) / 2
      : align === '3' ? bounds.x + bounds.width - old.width : old.x,
      y: align === '4' ? bounds.y : align === '5' ? bounds.y + (bounds.height - old.height) / 2
        : align === '6' ? bounds.y + bounds.height - old.height : old.y })
  }
  return members.flatMap(member => {
    const next = positions.get(member.imageId)!
    return Math.abs(next.x - member.rectangle.x) < 1e-9 && Math.abs(next.y - member.rectangle.y) < 1e-9 ? []
      : [positioned(member, { ...member.rectangle, ...next })]
  })
}
