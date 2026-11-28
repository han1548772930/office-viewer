import type { DocumentLayout, LayoutRect } from './document-layout'
import type { DocumentViewport } from '../rendering/document-scene'

export interface DocumentViewportLocation {
  reason?: 'scroll' | 'layout' | 'resize' | 'zoom' | 'selection'
  pageIndex: number | undefined
  firstVisibleBlockId?: string
  /** Present when the visible paragraph is nested inside a table cell. */
  firstVisibleParagraphId?: string
  top: number
  bottom: number
  left: number
  right: number
}

export interface DocumentScrollOffset { left: number; top: number }

type ViewportLayout = Pick<DocumentLayout, 'width' | 'height' | 'pages' | 'lines'>

function scale(zoom: number): number {
  return Number.isFinite(zoom) && zoom > 0 ? zoom / 100 : 1
}

function centeredOffset(width: number, contentWidth: number, zoom: number): number {
  return Math.max(0, (width - contentWidth * scale(zoom)) / 2)
}

function clampScroll(layout: ViewportLayout, viewport: DocumentViewport, offset: DocumentScrollOffset): DocumentScrollOffset {
  const ratio = scale(viewport.zoom)
  return {
    left: Math.max(0, Math.min(offset.left, Math.max(0, layout.width * ratio - viewport.width))),
    top: Math.max(0, Math.min(offset.top, Math.max(0, layout.height * ratio - viewport.height))),
  }
}

/** Layout coordinates, shared by the outline and page controls; scroll never changes selection. */
export function documentViewportLocation(layout: ViewportLayout, viewport: DocumentViewport): DocumentViewportLocation {
  const ratio = scale(viewport.zoom)
  const top = viewport.scrollTop / ratio
  const bottom = (viewport.scrollTop + viewport.height) / ratio
  const left = (viewport.scrollLeft - centeredOffset(viewport.width, layout.width, viewport.zoom)) / ratio
  const right = left + viewport.width / ratio
  const center = (top + bottom) / 2
  const page = viewport.height > 0
    ? layout.pages.find((candidate) => center <= candidate.y + candidate.height) ?? layout.pages.at(-1)
    : undefined
  const line = viewport.height > 0
    ? layout.lines.find((candidate) => candidate.y < bottom && candidate.y + candidate.height > top)
    : undefined
  return { pageIndex: page?.index, firstVisibleBlockId: line?.target.blockId,
    ...(line?.target.cellBlockId ? { firstVisibleParagraphId: line.target.cellBlockId } : {}), top, bottom, left, right }
}

/** Keep the reading position under the same CSS-pixel anchor when the editor zoom changes. */
export function zoomedDocumentScrollOffset(layout: ViewportLayout, viewport: DocumentViewport, zoom: number, anchor = { x: viewport.width / 2, y: viewport.height / 2 }): DocumentScrollOffset {
  const previous = scale(viewport.zoom)
  const next = scale(zoom)
  const logicalX = (viewport.scrollLeft + anchor.x - centeredOffset(viewport.width, layout.width, viewport.zoom)) / previous
  const logicalY = (viewport.scrollTop + anchor.y) / previous
  return clampScroll(layout, { ...viewport, zoom }, {
    left: centeredOffset(viewport.width, layout.width, zoom) + logicalX * next - anchor.x,
    top: logicalY * next - anchor.y,
  })
}

/** Minimal scrolling for a caret or object, including centered pages and document edges. */
export function documentRectScrollOffset(layout: ViewportLayout, viewport: DocumentViewport, rect: LayoutRect, padding = 28): DocumentScrollOffset {
  const ratio = scale(viewport.zoom)
  const insetX = Math.min(Math.max(0, padding), Math.max(0, (viewport.width - rect.width * ratio) / 2))
  const insetY = Math.min(Math.max(0, padding), Math.max(0, (viewport.height - rect.height * ratio) / 2))
  const x = centeredOffset(viewport.width, layout.width, viewport.zoom) + rect.x * ratio
  const y = rect.y * ratio
  let left = viewport.scrollLeft
  let top = viewport.scrollTop
  if (x < left + insetX) left = x - insetX
  else if (x + rect.width * ratio > left + viewport.width - insetX) left = x + rect.width * ratio - viewport.width + insetX
  if (y < top + insetY) top = y - insetY
  else if (y + rect.height * ratio > top + viewport.height - insetY) top = y + rect.height * ratio - viewport.height + insetY
  return clampScroll(layout, viewport, { left, top })
}
