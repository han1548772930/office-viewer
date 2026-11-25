import type { DocumentImage } from '../model/document-model'
import type { DocumentImagePosition } from '../images/document-image-model'
import type { LayoutImage, LayoutLine, LayoutRect, LayoutPage, LayoutColumn } from './document-layout-types'

export interface DocumentImageBox { width: number; height: number; paintWidth: number; paintHeight: number }
export type DocumentImageFrame = LayoutRect & { paintWidth?: number; paintHeight?: number; angle?: number }
export function documentImageFrameCorners(image: DocumentImageFrame) {
  const centerX = image.x + image.width / 2, centerY = image.y + image.height / 2, angle = (image.angle ?? 0) * Math.PI / 180
  const cosine = Math.cos(angle), sine = Math.sin(angle), width = image.paintWidth ?? image.width, height = image.paintHeight ?? image.height
  return ([['nw', -1, -1], ['ne', 1, -1], ['se', 1, 1], ['sw', -1, 1]] as const).map(([corner, dx, dy]) => ({ corner,
    x: centerX + dx * width * cosine / 2 - dy * height * sine / 2, y: centerY + dx * width * sine / 2 + dy * height * cosine / 2 }))
}
export function documentImageBox(image: DocumentImage, available: number, maximumHeight: number): DocumentImageBox {
  const angle = (image.angle ?? 0) * Math.PI / 180
  // Width/height describe the visible crop frame; fractions select source pixels inside it.
  const paintWidth = Math.max(1, image.width || 320)
  const paintHeight = Math.max(1, image.height || (image.width || 320) * 0.625)
  const width = Math.abs(Math.cos(angle)) * paintWidth + Math.abs(Math.sin(angle)) * paintHeight
  const height = Math.abs(Math.sin(angle)) * paintWidth + Math.abs(Math.cos(angle)) * paintHeight
  const scale = (image.wrappingStyle ?? 'inline') === 'inline'
    ? Math.max(0.000001, Math.min(1, Math.max(1, available) / width, Math.max(1, maximumHeight) / height)) : 1
  return { width: width * scale, height: height * scale, paintWidth: paintWidth * scale, paintHeight: paintHeight * scale }
}

/** Bitmap objects share their source glyph's baseline, host and UTF-16 address. */
export function appendDocumentInlineImages(line: LayoutLine, images: LayoutImage[]): void {
  for (const glyph of line.glyphs) {
    const atom = glyph.imageAtom
    if (!atom || (atom.image.wrappingStyle ?? 'inline') !== 'inline') continue
    const image = atom.image
    images.push({ blockId: atom.imageId, src: image.src, alt: image.alt, x: glyph.x, y: glyph.y, ...atom.box, page: line.page,
      sectionId: line.sectionId, sectionFragment: line.sectionFragment, column: line.column, host: line.host, clip: line.clip,
      tableCellId: line.tableCellId, repeatedHeader: line.repeatedHeader, objectImage: true, angle: image.angle, flipX: image.flipX, flipY: image.flipY,
      crop: image.crop, wrappingStyle: 'inline', layer: image.layer ?? 0,
      anchor: { target: line.target, startOffset: glyph.offset, endOffset: glyph.end } })
  }
}

function aligned(position: DocumentImagePosition | undefined, start: number, span: number, size: number, vertical: boolean, odd: boolean): number {
  if (!position) return start
  if (position.percent !== undefined) return start + span * position.percent / 100
  if (position.align === undefined) return start + (position.posOffset ?? 0)
  const align = position.align
  if (vertical) return start + (align === 0 || align === 3 && odd || align === 2 && !odd ? span - size : align === 1 ? (span - size) / 2 : 0)
  return start + (align === 4 || align === 3 && odd || align === 1 && !odd ? span - size : align === 0 || align === 5 || align === 6 ? (span - size) / 2 : 0)
}
export function positionDocumentImage(image: DocumentImage, page: Pick<LayoutPage, 'x' | 'y' | 'width' | 'height' | 'content' | 'pageNumber'>,
  column: Pick<LayoutColumn, 'x' | 'y' | 'width' | 'height'>, paragraphY: number, box: DocumentImageBox): { x: number; y: number } {
  const odd = page.pageNumber % 2 !== 0, content = page.content, h = image.positionH?.relativeFrom ?? 1, v = image.positionV?.relativeFrom ?? 1
  const leftMargin = { start: page.x, span: Math.max(0, content.x - page.x) }, rightMargin = { start: content.x + content.width, span: Math.max(0, page.x + page.width - content.x - content.width) }
  const horizontal = h === 0 ? { start: page.x, span: page.width } : h === 1 || h === 2 ? { start: column.x, span: h === 2 ? 0 : column.width }
    : h === 3 ? { start: content.x, span: content.width } : h === 6 || h === 4 && odd || h === 5 && !odd ? leftMargin : rightMargin
  const topMargin = { start: page.y, span: Math.max(0, content.y - page.y) }, bottomMargin = { start: content.y + content.height, span: Math.max(0, page.y + page.height - content.y - content.height) }
  const vertical = v === 0 ? { start: page.y, span: page.height } : v === 1 || v === 2 ? { start: paragraphY, span: 0 }
    : v === 3 ? { start: content.y, span: content.height } : v === 4 || v === 6 && odd || v === 7 && !odd ? topMargin : bottomMargin
  return { x: aligned(image.positionH, horizontal.start, horizontal.span, box.width, false, odd),
    y: aligned(image.positionV, vertical.start, vertical.span, box.height, true, odd) }
}

export interface DocumentImageObstacle extends LayoutRect { wrappingStyle: 'wrapSquare' | 'wrapTopAndBottom'; wrapSide: 'both' | 'left' | 'right' | 'largest' }
export function documentImageObstacle(image: LayoutImage, data: DocumentImage): DocumentImageObstacle | undefined {
  if (data.wrappingStyle !== 'wrapSquare' && data.wrappingStyle !== 'wrapTopAndBottom') return undefined
  const distances = data.distances ?? { top: 0, right: 0, bottom: 0, left: 0 }
  return { x: image.x - distances.left, y: image.y - distances.top, width: image.width + distances.left + distances.right,
    height: image.height + distances.top + distances.bottom, wrappingStyle: data.wrappingStyle, wrapSide: data.wrapSide ?? 'both' }
}
/** Available intervals stay in reading order and share one physical line. */
export function documentImageWrapBands(area: LayoutRect, y: number, height: number, obstacles: readonly DocumentImageObstacle[]): LayoutRect[] {
  for (let attempt = 0; attempt <= obstacles.length; attempt++) {
    const hits = obstacles.filter(item => item.y < y + height - 1e-6 && item.y + item.height > y + 1e-6
      && item.x < area.x + area.width && item.x + item.width > area.x)
    if (!hits.length) return [{ x: area.x, y, width: area.width, height }]
    let intervals = [{ start: area.x, end: area.x + area.width }]
    for (const hit of hits) {
      if (hit.wrappingStyle === 'wrapTopAndBottom') { intervals = []; break }
      intervals = intervals.flatMap(interval => {
        const left = { start: interval.start, end: Math.min(interval.end, hit.x) }, right = { start: Math.max(interval.start, hit.x + hit.width), end: interval.end }
        return (hit.wrapSide === 'left' ? [left] : hit.wrapSide === 'right' ? [right] : [left, right]).filter(item => item.end - item.start >= 24)
      })
    }
    if (intervals.length) {
      if (hits.some(hit => hit.wrapSide === 'largest')) intervals = intervals.sort((a, b) => b.end - b.start - (a.end - a.start) || a.start - b.start).slice(0, 1)
      return intervals.sort((a, b) => a.start - b.start).map(interval => ({ x: interval.start, y, width: interval.end - interval.start, height }))
    }
    y = Math.max(y + 1e-6, Math.min(...hits.map(hit => hit.y + hit.height)))
  }
  return [{ x: area.x, y, width: area.width, height }]
}

/** Compatibility for callers that explicitly request a single interval. */
export function documentImageWrapBand(area: LayoutRect, y: number, height: number, obstacles: readonly DocumentImageObstacle[]): LayoutRect {
  return documentImageWrapBands(area, y, height, obstacles).sort((a, b) => b.width - a.width || a.x - b.x)[0]!
}
