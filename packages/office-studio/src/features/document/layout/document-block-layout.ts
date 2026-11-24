import type { DocumentBlock, DocumentSectionBreakType } from '../../../domain/model'
import { paragraphListMarkers, type ParagraphListMarker } from '../model/document-lists'
import { imageForBlock } from '../model/document-model'
import { balanceDocumentSection } from './document-column-balance'
import { layoutDocumentParagraph, paragraphGeometry } from './document-paragraph-layout'
import { layoutDocumentTable } from './document-table-layout'
import { documentPictureInputAnchors, layoutDocumentPictureInputAnchor } from './document-picture-input-anchor'
import type { DocumentSectionFlow } from './document-section-flow'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutMeasurer } from './document-layout-types'

export interface DocumentBlockLayoutOptions {
  markers?: Map<string, ParagraphListMarker>
  beforeSectionBreak?: (sectionId: string, type: DocumentSectionBreakType) => void
  textDisplay?: DocumentTextLayoutDisplay
}

/** Body, notes and separator drafts share paragraph/table/object pagination. */
export function appendDocumentBlocks(blocks: DocumentBlock[], flow: DocumentSectionFlow, result: DocumentLayout,
  measurer: LayoutMeasurer, options: DocumentBlockLayoutOptions = {}): void {
  const markers = options.markers ?? paragraphListMarkers(blocks)
  const pictureAnchors = documentPictureInputAnchors(blocks), anchorIds = new Set([...pictureAnchors.values()].map(block => block.id))
  for (const block of blocks) {
    if (anchorIds.has(block.id)) continue
    if (block.type === 'pageBreak') { flow.pageBreak(); continue }
    if (block.type === 'columnBreak') { flow.nextColumn(true); continue }
    if (block.type === 'sectionBreak') {
      const type = block.sectionBreak?.type ?? 'nextPage'
      if (type === 'continuous' && flow.columnsChange(block.id)) balanceDocumentSection(result, flow.section)
      options.beforeSectionBreak?.(block.id, type)
      flow.startSection(block.id, type)
      continue
    }
    flow.ensureRoom(1)
    const inputAnchor = pictureAnchors.get(block.id), style = inputAnchor?.style ?? block.style
    flow.addSpace(style.spaceBefore)
    if (block.type === 'table') layoutDocumentTable(block, markers, flow, result, measurer, options.textDisplay)
    else if (block.type === 'image') {
      const image = imageForBlock(block)
      if (!image) continue
      const nativeWidth = Math.max(1, image.width || 320), nativeHeight = Math.max(1, image.height || nativeWidth * 0.625)
      const size = () => {
        const available = inputAnchor ? paragraphGeometry(inputAnchor, markers.get(inputAnchor.id), flow.column.width, measurer).width(0) : flow.column.width
        const scale = Math.min(1, available / nativeWidth, flow.fullPageContentHeight / nativeHeight)
        return { width: nativeWidth * scale, height: nativeHeight * scale }
      }
      let prepared = size()
      while (true) {
        flow.ensureRoom(prepared.height)
        const next = size()
        if (Math.abs(next.width - prepared.width) < 1e-6 && Math.abs(next.height - prepared.height) < 1e-6) break
        prepared = next
      }
      const geometry = inputAnchor && paragraphGeometry(inputAnchor, markers.get(inputAnchor.id), flow.column.width, measurer)
      const available = geometry ? geometry.width(0) : flow.column.width
      const shift = style.align === 'center' ? (available - prepared.width) / 2 : style.align === 'right' ? available - prepared.width : 0
      const positioned = { x: flow.column.x + (geometry?.firstIndent ?? 0) + shift, y: flow.cursor, ...prepared, blockId: block.id,
        src: image.src, alt: image.alt, page: flow.page.index, ...flow.location }
      result.images.push(positioned)
      if (inputAnchor) result.lines.push(layoutDocumentPictureInputAnchor(inputAnchor, positioned, { blockId: inputAnchor.id },
        flow.column.width, flow.column.x, measurer, result.images, markers.get(inputAnchor.id), options.textDisplay))
      flow.consume(prepared.height)
    } else if (block.type === 'divider') {
      flow.ensureRoom(18)
      const fraction = typeof block.metadata.noteSeparatorWidthFraction === 'number' ? Math.max(0, Math.min(1, block.metadata.noteSeparatorWidthFraction)) : 1
      result.dividers.push({ x: flow.column.x, y: flow.cursor + 9, width: flow.column.width * fraction, height: 1, page: flow.page.index, ...flow.location,
        ...(typeof block.metadata.noteSeparatorColor === 'string' ? { color: block.metadata.noteSeparatorColor } : {}) })
      flow.consume(18)
    } else layoutDocumentParagraph(block, markers.get(block.id), flow, result, measurer, options.textDisplay)
    flow.addSpace(style.spaceAfter)
  }
}
