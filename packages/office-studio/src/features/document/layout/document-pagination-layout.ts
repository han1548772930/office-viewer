import type { DocumentBlock, DocumentCodeConfig } from '../../../domain/model'
import { isDocumentStructuredBlock, isDocumentTextBlock } from '../model/document-structural-blocks'
import type { ParagraphListMarker } from '../model/document-lists'
import type { DocumentSectionFlow } from './document-section-flow'
import type { DocumentTextLayoutDisplay, LayoutLine, LayoutMeasurer } from './document-layout-types'
import type { RawLine } from './document-text-layout'
import { measureDocumentParagraphLines } from './document-paragraph-layout'
import { documentDecoratedBlockInsets } from './document-structural-layout'
import { measureDocumentBlockCollections } from './document-block-measurement'
import { measureDocumentTable } from './document-table-measurement'

export const DOCUMENT_KEEP_NEXT_SCAN_LIMIT = 32
const EPSILON = 1e-6
const textParagraph = (block: DocumentBlock | undefined): block is DocumentBlock => !!block && isDocumentTextBlock(block) && !isDocumentStructuredBlock(block)
interface PaginationContext {
  markers: Map<string, ParagraphListMarker>
  measurer: LayoutMeasurer
  display?: DocumentTextLayoutDisplay
  code?: DocumentCodeConfig
}

function linesFloor(lines: readonly (RawLine | LayoutLine)[], flow: DocumentSectionFlow): number {
  return lines.reduce((page, line) => Math.max(page,
    flow.referencePageFloor(line.glyphs.filter(glyph => glyph.note?.role === 'reference').map(glyph => glyph.run.id))), 0)
}

/** The last member only contributes its opening unless it has keepLines.
 * Explicit break tokens, object hosts and pageBreakBefore stop adjacency. */
function measureKeepChain(blocks: readonly DocumentBlock[], index: number, width: number, flow: DocumentSectionFlow, context: PaginationContext) {
  let height = 0, count = 0
  const lines: RawLine[] = []
  for (let at = index; at < blocks.length && count < DOCUMENT_KEEP_NEXT_SCAN_LIMIT; at++) {
    const block = blocks[at]
    if (!textParagraph(block)) break
    const rules = flow.paragraphPagination(block.style)
    if (at > index && rules.pageBreakBefore) break
    const measured = measureDocumentParagraphLines(block, context.markers.get(block.id), width, context.measurer, context.display, context.code, 0, 0, flow.fullPageContentHeight)
    const following = blocks[at + 1], linked = rules.keepNext && textParagraph(following) && !flow.paragraphPagination(following.style).pageBreakBefore
    const whole = at === index || rules.keepLines || linked
    const opening = rules.widowControl ? Math.min(2, measured.lines.length) : 1
    height += block.style.spaceBefore + (whole ? measured.height + block.style.spaceAfter : measured.heights.slice(0, opening).reduce((sum, value) => sum + value, 0))
    lines.push(...(whole ? measured.lines : measured.lines.slice(0, opening)))
    count++
    if (!linked) break
  }
  return { height, count, lines }
}

function reserveTogether(height: number, lines: readonly (RawLine | LayoutLine)[], flow: DocumentSectionFlow): void {
  if (height <= flow.fullPageContentHeight + EPSILON) flow.advanceToPage(linesFloor(lines, flow))
  flow.keepContentTogether(height)
}

/** Run before consuming spacing or capturing decoration starts. Measurements
 * are repeated only if moving to a new frame changes the actual wrap width. */
export function prepareDocumentBlockPagination(blocks: readonly DocumentBlock[], index: number, flow: DocumentSectionFlow, context: PaginationContext): void {
  const block = blocks[index]!
  if (isDocumentStructuredBlock(block) && block.type !== 'columns') {
    const first = block.children?.[0]
    if (!first) return
    if (flow.paginationEnabled && textParagraph(first) && flow.paragraphPagination(first.style).pageBreakBefore) flow.paragraphPageBreakBefore()
    while (true) {
      const available = flow.column.width, inset = documentDecoratedBlockInsets(block), sum = inset.left + inset.right
      const scale = sum > available - 1 ? Math.max(0, available - 1) / Math.max(1, sum) : 1, width = Math.max(1, available - sum * scale)
      if (textParagraph(first)) {
        const innerContext = { ...context, code: block.code ?? context.code }
        const measured = measureDocumentParagraphLines(first, context.markers.get(first.id), width, context.measurer, context.display, innerContext.code, 0, 0, flow.fullPageContentHeight)
        const rules = flow.paragraphPagination(first.style), prefix = block.style.spaceBefore + inset.top
        const opening = measured.heights.slice(0, flow.paginationEnabled && rules.widowControl ? 2 : 1).reduce((height, value) => height + value, 0)
        if (flow.paginationEnabled && (rules.keepLines || rules.keepNext)) {
          const chain = measureKeepChain(block.children!, 0, width, flow, innerContext)
          const tail = chain.count === block.children!.length ? inset.bottom + block.style.spaceAfter : 0
          const requested = prefix + chain.height + tail
          if (requested <= flow.fullPageContentHeight + EPSILON) reserveTogether(requested, chain.lines, flow)
          else reserveTogether(prefix + first.style.spaceBefore + measured.height + first.style.spaceAfter, measured.lines, flow)
        } else flow.ensureRoom(prefix + first.style.spaceBefore + opening)
      } else {
        // Object-led containers share the existing table/image measurement;
        // this preflight only ensures that a first visible item can begin.
        const measured = measureDocumentBlockCollections([[block]], { width: available, markers: context.markers, measurer: context.measurer,
          measureTable: measureDocumentTable, maximumObjectHeight: flow.fullPageContentHeight, display: context.display, code: context.code })
        const firstHeight = measured.lines[0]?.height ?? measured.images[0]?.height ?? 1
        flow.ensureRoom(block.style.spaceBefore + inset.top + first.style.spaceBefore + firstHeight)
      }
      if (Math.abs(available - flow.column.width) <= EPSILON) return
    }
  }
  if (!textParagraph(block)) return
  const rules = flow.paragraphPagination(block.style)
  if (flow.paginationEnabled && rules.pageBreakBefore) flow.paragraphPageBreakBefore()
  if (!block.style.spaceBefore && (!flow.paginationEnabled || !rules.keepLines && !rules.keepNext && !rules.widowControl)) return
  while (true) {
    const width = flow.column.width, measured = measureDocumentParagraphLines(block, context.markers.get(block.id), width, context.measurer, context.display, context.code, 0, 0, flow.fullPageContentHeight)
    if (flow.paginationEnabled && (rules.keepLines || rules.keepNext)) {
      const chain = measureKeepChain(blocks, index, width, flow, context)
      if (chain.height <= flow.fullPageContentHeight + EPSILON) reserveTogether(chain.height, chain.lines, flow)
      else reserveTogether(block.style.spaceBefore + measured.height + block.style.spaceAfter, measured.lines, flow)
    } else {
      const opening = measured.heights.slice(0, flow.paginationEnabled && rules.widowControl ? 2 : 1).reduce((height, value) => height + value, 0)
      flow.ensureRoom(block.style.spaceBefore + opening)
    }
    if (Math.abs(width - flow.column.width) <= EPSILON) return
  }
}

/** Reserve the successor opening on the current paragraph's last fragment,
 * so even an oversized predecessor can retain a feasible keepNext boundary. */
export function documentParagraphSuccessorHeight(blocks: readonly DocumentBlock[], index: number, flow: DocumentSectionFlow, context: PaginationContext): number {
  const block = blocks[index]!, next = blocks[index + 1]
  if (!flow.paginationEnabled || !flow.paragraphPagination(block.style).keepNext || !textParagraph(next)) return 0
  const rules = flow.paragraphPagination(next.style)
  if (rules.pageBreakBefore) return 0
  const measured = measureDocumentParagraphLines(next, context.markers.get(next.id), flow.column.width, context.measurer, context.display, context.code)
  const height = rules.keepLines && measured.height + next.style.spaceBefore + next.style.spaceAfter <= flow.fullPageContentHeight + EPSILON
    ? measured.height + next.style.spaceAfter : measured.heights.slice(0, rules.widowControl ? 2 : 1).reduce((sum, value) => sum + value, 0)
  return block.style.spaceAfter + next.style.spaceBefore + height
}
