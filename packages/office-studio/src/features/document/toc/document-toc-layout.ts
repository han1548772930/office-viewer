import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import type { DocumentLayout, LayoutRect, LayoutRightTabStop, LayoutTabLeader } from '../layout/document-layout-types'
import { runText, type DocumentSelection } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentTextFields } from '../model/document-text-fields'
import { documentTocIdAtBlock, documentTocIdForSelection, documentTocOptionsAtBlock } from './document-toc'

function record(value: unknown): Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
}

function nativeLeader(value: unknown): LayoutTabLeader | undefined {
  return value === undefined || value === 'none' ? 'none' : value === 'dot' ? 'dots' : value === 'hyphen' ? 'dashes'
    : value === 'underscore' ? 'underline' : value === 'middleDot' ? 'middleDots' : undefined
}

export interface DocumentTocPageTab {
  /** UTF-16 position of the actual tab immediately before a PAGEREF result. */
  offset: number
  stops: readonly Omit<LayoutRightTabStop, 'offset'>[]
}

/** Native w:pos is twips; 1 CSS px = 15 twips at 96 dpi. Nothing is written to the model. */
export function documentTocPageTab(block: DocumentBlock): DocumentTocPageTab | undefined {
  const metadata = record(block.metadata.documentToc), options = documentTocOptionsAtBlock(block)
  if (metadata.role !== 'entry' || !options) return undefined
  const text = runText(block.runs)
  const page = documentTextFields(block).find(field => field.field.toUpperCase() === 'PAGEREF' && (!field.endBlockId || field.endBlockId === block.id)
    && field.startOffset > 0 && text[field.startOffset - 1] === '\t')
  if (!page) return undefined
  const source = record(block.metadata.source), details = record(source.details), style = record(details.style)
  const sourceStops = Array.isArray(style.tabStops) ? style.tabStops : []
  const nativeStops = sourceStops.flatMap(value => {
    const stop = record(value), leader = nativeLeader(stop.leader)
    if (stop.alignment !== 'right' || stop.clear === true || leader === undefined || typeof stop.positionTwips !== 'number'
      || !Number.isSafeInteger(stop.positionTwips) || stop.positionTwips < 0) return []
    return [{ right: stop.positionTwips / 15, leader }]
  }).sort((a, b) => a.right - b.right)
  // The retained result and its actual paragraph stops can differ from the TOC
  // instruction. Use a supported source stop even when the instruction says \p.
  if (nativeStops.length) return { offset: page.startOffset - 1, stops: nativeStops }
  // An unsupported explicit stop must not silently become a generated dot leader.
  if (sourceStops.length || !options.rightAlignPageNumbers) return undefined
  return { offset: page.startOffset - 1, stops: [{ leader: options.tabLeader }] }
}

export interface DocumentTocHighlightRect extends LayoutRect { page: number }
export interface DocumentTocHighlightLayout { tocId: string; ranges: DocumentTocHighlightRect[]; outlines: DocumentTocHighlightRect[] }

/** Read-only context fill and four-pixel outlines per physical page/flow fragment. */
export function documentTocHighlightLayout(model: DocumentModel, layout: DocumentLayout, selection: DocumentSelection): DocumentTocHighlightLayout | undefined {
  const tocId = documentTocIdForSelection(model, selection)
  if (!tocId) return undefined
  for (const point of [selection.anchor, selection.focus]) {
    const block = model.blocks.find(block => block.id === point.blockId)
    if (!block || point.cellBlockId || point.row != null || point.col != null || !isDocumentTextBoundary(runText(block.runs), point.offset)) return undefined
  }
  const ids = new Set(model.blocks.filter(block => documentTocIdAtBlock(block) === tocId).map(block => block.id))
  const ranges: DocumentTocHighlightRect[] = [], bounds = new Map<string, DocumentTocHighlightRect>()
  for (const line of layout.lines) {
    const page = layout.pages[line.page]
    if (!ids.has(line.target.blockId) || line.target.cellBlockId || !page || page.isParityBlank) continue
    const range = { page: line.page, x: line.x, y: line.y, width: Math.max(2, line.width), height: line.height }
    ranges.push(range)
    const key = `${line.page}:${line.sectionFragment ?? ''}:${line.column ?? 0}`, previous = bounds.get(key)
    if (!previous) { bounds.set(key, { ...range }); continue }
    const left = Math.min(previous.x, range.x), top = Math.min(previous.y, range.y)
    const right = Math.max(previous.x + previous.width, range.x + range.width), bottom = Math.max(previous.y + previous.height, range.y + range.height)
    bounds.set(key, { page: line.page, x: left, y: top, width: right - left, height: bottom - top })
  }
  if (!ranges.length) return undefined
  return { tocId, ranges, outlines: [...bounds.values()].map(bound => ({ ...bound, x: bound.x - 4, y: bound.y - 4,
    width: bound.width + 8, height: bound.height + 8 })) }
}
