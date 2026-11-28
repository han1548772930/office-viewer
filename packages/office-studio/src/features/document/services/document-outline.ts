import type { DocumentModel } from '../../../domain/model.ts'
import { textTargets, type TextTarget } from '../model/document-model'
import { documentBlockChildContainers } from '../model/document-structural-blocks'
import type { LayoutLine } from '../layout/document-layout-types'

export interface DocumentOutlineItem {
  /** Stable paragraph ID, or a view-only ID for the document title. */
  id: string
  blockId?: string
  text: string
  level: number
  depth: number
  isTitle: boolean
  blockIndex: number
  /** Canonical paragraph address, including a table cell's root/row/column. */
  target?: TextTarget
  sourceOrder?: number
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
 * An explicit workbench setting wins over the document flavor. The outline
 * defaults to enabled for modern documents and disabled for
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
 * heading after a user changes its style. An authored Title paragraph takes
 * precedence over the view-only document name; Subtitle remains excluded.
 */
export function buildDocumentOutline(model: DocumentModel, options: DocumentOutlineOptions = {}): DocumentOutlineItem[] {
  const items: DocumentOutlineItem[] = []
  const ids = new Set<string>()
  const blocks = new Map<string, { block: DocumentModel['blocks'][number]; blockIndex: number }>()
  const collect = (values: DocumentModel['blocks'], rootIndex: number, depth = 0) => {
    if (depth > 64) throw new RangeError('文档大纲结构超过 64 层。')
    for (const block of values) {
      if (!blocks.has(block.id)) blocks.set(block.id, { block, blockIndex: rootIndex })
      for (const children of documentBlockChildContainers(block)) collect(children, rootIndex, depth + 1)
    }
  }
  model.blocks.forEach((block, blockIndex) => collect([block], blockIndex))
  const targets = textTargets(model)
  for (const [sourceOrder, target] of targets.entries()) {
    const entry = blocks.get(target.cellBlockId ?? target.blockId)
    if (!entry) continue
    const { block, blockIndex } = entry
    const isTitle = block.type === 'paragraph' && block.metadata.namedStyleType === 'title'
    if ((block.type !== 'heading' && !isTitle) || !block.id || ids.has(block.id)) continue
    const text = block.runs.map((run) => run.text).join('').replace(/[\r\n\t]/g, ' ').trim()
    if (!text) continue
    const level = isTitle ? 1 : Number.isFinite(block.level) ? Math.min(6, Math.max(1, Math.floor(block.level))) : 1
    items.push({ id: block.id, blockId: block.id, text, level, depth: level - 1, isTitle, blockIndex, target: { ...target }, sourceOrder })
    ids.add(block.id)
  }

  if (options.includeDocumentTitle !== false && !items.some(item => item.isTitle) && model.name.trim()) {
    // Avoid colliding with a real paragraph ID without changing any model ID.
    let titleId = DOCUMENT_TITLE_ID
    const blockIds = new Set(blocks.keys())
    while (blockIds.has(titleId)) titleId += '_'
    items.unshift({ id: titleId, text: model.name.trim(), level: 1, depth: 0, isTitle: true, blockIndex: -1 })
  }
  return items
}

/** Return the current section when the caret is in a heading or its body. */
export function activeOutlineIdForBlock(items: readonly DocumentOutlineItem[], model: DocumentModel, blockId: string): string | undefined {
  const targets = textTargets(model), sourceOrder = targets.findIndex(target => (target.cellBlockId ?? target.blockId) === blockId)
  const blockIndex = model.blocks.findIndex((block) => block.id === blockId)
  if (sourceOrder < 0 && blockIndex < 0) return undefined
  let activeId: string | undefined
  for (const item of items) {
    if (sourceOrder >= 0 ? (item.sourceOrder ?? -1) > sourceOrder : item.blockIndex > blockIndex) break
    activeId = item.id
  }
  return activeId
}

/** Preserve paragraph order across table cells, columns, wrapping and page splits. */
export function documentOutlineParagraphBounds(model: DocumentModel, lines: readonly Pick<LayoutLine, 'target' | 'y' | 'height'>[]): DocumentOutlineBlockBound[] {
  const bounds = new Map<string, DocumentOutlineBlockBound>()
  for (const line of lines) {
    if (!Number.isFinite(line.y) || !Number.isFinite(line.height) || line.height <= 0) continue
    const id = line.target.cellBlockId ?? line.target.blockId
    const bound = bounds.get(id)
    if (bound) { bound.top = Math.min(bound.top, line.y); bound.bottom = Math.max(bound.bottom, line.y + line.height) }
    else bounds.set(id, { blockId: id, top: line.y, bottom: line.y + line.height })
  }
  return textTargets(model).flatMap(target => {
    const bound = bounds.get(target.cellBlockId ?? target.blockId)
    return bound ? [bound] : []
  })
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
