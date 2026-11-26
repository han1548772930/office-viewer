import type { DocumentBlock, DocumentCodeConfig, DocumentSectionBreakType } from '../../../domain/model'
import { paragraphListMarkers, type ParagraphListMarker } from '../model/document-lists'
import { imageForBlock } from '../model/document-model'
import { documentHorizontalLineBorder } from '../model/document-horizontal-line'
import { balanceDocumentSection } from './document-column-balance'
import { layoutDocumentParagraph, paragraphGeometry } from './document-paragraph-layout'
import { layoutDocumentTable } from './document-table-layout'
import { documentPictureInputAnchors, layoutDocumentPictureInputAnchor } from './document-picture-input-anchor'
import { isDocumentStructuredBlock } from '../model/document-structural-blocks'
import { layoutDocumentStructuredBlock } from './document-structural-layout'
import type { DocumentSectionFlow } from './document-section-flow'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutMeasurer } from './document-layout-types'
import { documentParagraphSuccessorHeight, prepareDocumentBlockPagination } from './document-pagination-layout'
import { documentImageBox, documentImageObstacle, positionDocumentImage } from './document-image-layout'

export interface DocumentBlockLayoutOptions {
  markers?: Map<string, ParagraphListMarker>
  beforeSectionBreak?: (sectionId: string, type: DocumentSectionBreakType) => void
  textDisplay?: DocumentTextLayoutDisplay
  code?: DocumentCodeConfig
  /** Bottom padding of a decorated host, consumed after its last child. */
  endInset?: number
}

/** Body, notes and separator drafts share paragraph/table/object pagination. */
export function appendDocumentBlocks(blocks: DocumentBlock[], flow: DocumentSectionFlow, result: DocumentLayout,
  measurer: LayoutMeasurer, options: DocumentBlockLayoutOptions = {}): void {
  const markers = options.markers ?? paragraphListMarkers(blocks)
  const pictureAnchors = documentPictureInputAnchors(blocks), anchorIds = new Set([...pictureAnchors.values()].map(block => block.id))
  const paginationContext = { markers, measurer, display: options.textDisplay, code: options.code }
  for (const [index, block] of blocks.entries()) {
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
    prepareDocumentBlockPagination(blocks, index, flow, paginationContext)
    flow.ensureRoom(1)
    const inputAnchor = pictureAnchors.get(block.id), style = inputAnchor?.style ?? block.style
    flow.addSpace(style.spaceBefore)
    if (isDocumentStructuredBlock(block)) layoutDocumentStructuredBlock(block, flow, result,
      (children, childFlow, endInset = 0) => appendDocumentBlocks(children, childFlow, result, measurer, { ...options, markers, endInset, code: block.code ?? options.code }))
    else if (block.type === 'table') layoutDocumentTable(block, markers, flow, result, measurer, options.textDisplay)
    else if (block.type === 'image') {
      const image = imageForBlock(block)
      if (!image) continue
      const size = () => {
        const available = inputAnchor ? paragraphGeometry(inputAnchor, markers.get(inputAnchor.id), flow.column.width, measurer).width(0) : flow.column.width
        return documentImageBox(image, available, flow.fullPageContentHeight)
      }
      let prepared = size()
      while (true) {
        flow.ensureRoom((image.wrappingStyle ?? 'inline') === 'inline' ? prepared.height : 1)
        const next = size()
        if (Math.abs(next.width - prepared.width) < 1e-6 && Math.abs(next.height - prepared.height) < 1e-6) break
        prepared = next
      }
      const geometry = inputAnchor && paragraphGeometry(inputAnchor, markers.get(inputAnchor.id), flow.column.width, measurer)
      const available = geometry ? geometry.width(0) : flow.column.width
      const shift = style.align === 'center' ? (available - prepared.width) / 2 : style.align === 'right' ? available - prepared.width : 0
      const floated = (image.wrappingStyle ?? 'inline') !== 'inline'
      const position = floated ? positionDocumentImage(image, flow.page, flow.column, flow.cursor, prepared)
        : { x: flow.column.x + (geometry?.firstIndent ?? 0) + shift, y: flow.cursor }
      const positioned = { ...position, ...prepared, blockId: block.id, src: image.src, alt: image.alt, page: flow.page.index, ...flow.location,
        objectImage: true, angle: image.angle, flipX: image.flipX, flipY: image.flipY, crop: image.crop, wrappingStyle: image.wrappingStyle ?? 'inline', layer: image.layer ?? 0 }
      result.images.push(positioned)
      const obstacle = documentImageObstacle(positioned, image)
      if (obstacle) flow.addImageObstacle(obstacle)
      if (inputAnchor) result.lines.push(layoutDocumentPictureInputAnchor(inputAnchor, positioned, { blockId: inputAnchor.id },
        flow.column.width, flow.column.x, measurer, result.images, markers.get(inputAnchor.id), options.textDisplay))
      if (!floated) flow.consume(prepared.height)
    } else if (block.type === 'divider') {
      const border = documentHorizontalLineBorder(block), height = Math.max(18, border.padding * 2 + border.width)
      flow.ensureRoom(height)
      const fraction = typeof block.metadata.noteSeparatorWidthFraction === 'number' ? Math.max(0, Math.min(1, block.metadata.noteSeparatorWidthFraction)) : 1
      const separator = block.metadata.noteSeparatorWidthFraction !== undefined
      result.dividers.push({ x: flow.column.x, y: flow.cursor + (separator ? 9 : border.padding + border.width / 2), width: flow.column.width * fraction,
        height: border.width, page: flow.page.index, ...flow.location, style: border.style,
        color: typeof block.metadata.noteSeparatorColor === 'string' ? block.metadata.noteSeparatorColor : separator ? '#b3bdca' : border.color })
      flow.consume(height)
    } else {
      const endInset = index === blocks.length - 1 && options.endInset ? options.endInset + style.spaceAfter : 0
      const successor = documentParagraphSuccessorHeight(blocks, index, flow, paginationContext)
      layoutDocumentParagraph(block, markers.get(block.id), flow, result, measurer, options.textDisplay, options.code, Math.max(endInset, successor))
    }
    flow.addSpace(style.spaceAfter)
  }
}
