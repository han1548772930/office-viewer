import type { DocumentModel } from '../../../domain/model.ts'

export interface DocumentOutlineItem {
  /** Stable paragraph ID, or a view-only ID for the document title. */
  id: string
  blockId?: string
  text: string
  level: number
  depth: number
  isTitle: boolean
  blockIndex: number
}

export interface DocumentOutlineOptions {
  includeDocumentTitle?: boolean
}

export interface DocumentOutlineBlockBound {
  blockId: string
  top: number
  bottom: number
}

export interface DocumentOutlineRect {
  x: number
  y: number
  width: number
  height: number
}

export interface DocumentOutlineTableViewport {
  /** Table content bounds before its own horizontal scroll, in CSS pixels. */
  bounds: DocumentOutlineRect
  scrollLeft?: number
  /** Visible table viewport in the same CSS coordinate system. */
  clip?: DocumentOutlineRect
}

export interface DocumentOutlinePlacement {
  availableWidth: number
  pageWidth: number
  /** Workbench percentage, e.g. 100 rather than 1. */
  zoom: number
  outlineWidth?: number
  canvasPadding?: number
  outlineBounds?: DocumentOutlineRect
  tables?: readonly DocumentOutlineTableViewport[]
}

const DOCUMENT_TITLE_ID = '__document_outline_title__'

function record(value: unknown): Record<string, unknown> | undefined {
  return value != null && typeof value === 'object' && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined
}

/**
 * An explicit workbench setting wins over the document flavor. Univer's
 * DocSideMenu defaults to enabled for modern documents and disabled for
 * traditional documents. A model without a declared flavor is traditional.
 */
export function documentOutlineEnabled(model: DocumentModel, configured?: boolean): boolean {
  if (configured !== undefined) return configured
  const metadata = model.metadata
  const source = record(metadata?.source)
  const flavor = metadata?.documentFlavor
    ?? record(metadata?.documentStyle)?.documentFlavor
    ?? record(source?.documentStyle)?.documentFlavor
  return flavor === 2 || (typeof flavor === 'string' && flavor.toLowerCase() === 'modern')
}

/**
 * The editable model's current block type is authoritative. In particular,
 * preserved parser/source metadata must not turn a paragraph back into a
 * heading after a user changes its style. Title and Subtitle are not separate
 * block types in this model; the document name supplies the view-only title.
 */
export function buildDocumentOutline(model: DocumentModel, options: DocumentOutlineOptions = {}): DocumentOutlineItem[] {
  const items: DocumentOutlineItem[] = []
  const ids = new Set<string>()
  for (const [blockIndex, block] of model.blocks.entries()) {
    if (block.type !== 'heading' || !block.id || ids.has(block.id)) continue
    const text = block.runs.map((run) => run.text).join('').replace(/[\r\n\t]/g, ' ').trim()
    if (!text) continue
    const level = Number.isFinite(block.level) ? Math.min(6, Math.max(1, Math.floor(block.level))) : 1
    items.push({ id: block.id, blockId: block.id, text, level, depth: level - 1, isTitle: false, blockIndex })
    ids.add(block.id)
  }

  if (options.includeDocumentTitle !== false && model.name.trim()) {
    // Avoid colliding with a real paragraph ID without changing any model ID.
    let titleId = DOCUMENT_TITLE_ID
    const blockIds = new Set(model.blocks.map((block) => block.id))
    while (blockIds.has(titleId)) titleId += '_'
    items.unshift({ id: titleId, text: model.name.trim(), level: 1, depth: 0, isTitle: true, blockIndex: -1 })
  }
  return items
}

/** Return the current section when the caret is in a heading or its body. */
export function activeOutlineIdForBlock(items: readonly DocumentOutlineItem[], model: DocumentModel, blockId: string): string | undefined {
  const blockIndex = model.blocks.findIndex((block) => block.id === blockId)
  if (blockIndex < 0) return undefined
  let activeId: string | undefined
  for (const item of items) {
    if (item.blockIndex > blockIndex) break
    activeId = item.id
  }
  return activeId
}

/**
 * Bounds are supplied in document order and document coordinates by layout.
 * Match DocSideMenu: prefer the first heading visible in the viewport, then
 * the nearest heading preceding the first visible paragraph. Touching an edge
 * does not make a paragraph visible, and zero-height/invalid bounds are ignored.
 */
export function activeOutlineIdInViewport(
  items: readonly DocumentOutlineItem[],
  bounds: readonly DocumentOutlineBlockBound[],
  top: number,
  bottom: number,
): string | undefined {
  if (!Number.isFinite(top) || !Number.isFinite(bottom) || bottom <= top) return undefined
  const headings = new Map(items.filter((item) => item.blockId).map((item) => [item.blockId, item.id]))
  const valid = (bound: DocumentOutlineBlockBound): boolean => Number.isFinite(bound.top) && Number.isFinite(bound.bottom) && bound.bottom > bound.top
  const visible = (bound: DocumentOutlineBlockBound): boolean => valid(bound) && bound.top < bottom && bound.bottom > top
  let firstVisible = -1
  let lastVisible = -1
  for (const [index, bound] of bounds.entries()) {
    if (visible(bound)) {
      if (firstVisible < 0) firstVisible = index
      lastVisible = index
    }
  }
  if (firstVisible < 0) return undefined
  for (let index = firstVisible; index <= lastVisible; index++) {
    if (!visible(bounds[index]!)) continue
    const id = headings.get(bounds[index]!.blockId)
    if (id) return id
  }
  for (let index = firstVisible; index >= 0; index--) {
    if (!valid(bounds[index]!)) continue
    const id = headings.get(bounds[index]!.blockId)
    if (id) return id
  }
  return items.find((item) => item.isTitle)?.id
}

/** Focus navigation is separate from activating a heading and moving the caret. */
export function nextOutlineFocusId(items: readonly DocumentOutlineItem[], currentId: string | undefined, key: string): string | undefined {
  if (!items.length || !['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(key)) return undefined
  const index = items.findIndex((item) => item.id === currentId)
  if (key === 'Home') return items[0]!.id
  if (key === 'End') return items.at(-1)!.id
  if (index < 0) return items[0]!.id
  return items[Math.max(0, Math.min(items.length - 1, index + (key === 'ArrowDown' ? 1 : -1)))]!.id
}

function validRect(rect: DocumentOutlineRect): boolean {
  return [rect.x, rect.y, rect.width, rect.height].every(Number.isFinite) && rect.width > 0 && rect.height > 0
}

/**
 * Compare visible content, not a table viewport's empty leading inset. Bounds
 * and clipping are explicit so a horizontal table scroll never gets applied
 * twice, and a clipped wide table does not unnecessarily hide the outline.
 */
export function documentOutlineOverlapsTable(outline: DocumentOutlineRect, table: DocumentOutlineTableViewport): boolean {
  if (!validRect(outline) || !validRect(table.bounds) || (table.clip && !validRect(table.clip))) return false
  const scrollLeft = Number.isFinite(table.scrollLeft) ? table.scrollLeft ?? 0 : 0
  let left = table.bounds.x - scrollLeft
  let right = left + table.bounds.width
  let top = table.bounds.y
  let bottom = top + table.bounds.height
  if (table.clip) {
    left = Math.max(left, table.clip.x)
    right = Math.min(right, table.clip.x + table.clip.width)
    top = Math.max(top, table.clip.y)
    bottom = Math.min(bottom, table.clip.y + table.clip.height)
  }
  return right > left && bottom > top
    && left < outline.x + outline.width && right > outline.x
    && top < outline.y + outline.height && bottom > outline.y
}

/** Choose an overlay when docking would squeeze the page or cover visible table content. */
export function documentOutlinePlacement(input: DocumentOutlinePlacement): 'side-bar' | 'float' {
  if (![input.availableWidth, input.pageWidth, input.zoom].every(Number.isFinite)
    || input.availableWidth <= 0 || input.pageWidth <= 0 || input.zoom <= 0) return 'float'
  const outlineWidth = Math.min(256, Math.max(32, Number.isFinite(input.outlineWidth) ? input.outlineWidth ?? 256 : 256))
  const padding = Math.max(0, Number.isFinite(input.canvasPadding) ? input.canvasPadding ?? 48 : 48)
  if (input.availableWidth < input.pageWidth * input.zoom / 100 + padding + outlineWidth) return 'float'
  const outlineBounds = input.outlineBounds
  if (outlineBounds && input.tables?.some((table) => documentOutlineOverlapsTable(outlineBounds, table))) return 'float'
  return 'side-bar'
}
