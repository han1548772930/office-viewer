import type { DocumentBlock, DocumentModel } from '../domain/model.ts'
import { flattenDocumentBody, type DocumentBodySegment } from '../features/document/services/document-facade.ts'
import { isDocumentTextBoundary } from '../features/document/model/document-story.ts'
import { findDocumentBlock, findDocumentBlockContainer, sameTarget, type DocumentPoint, type DocumentSelection } from '../features/document/model/document-model.ts'

/** Facade offsets keep the existing synthetic-newline and independent-cell contract. */
export function documentSelectionAtBodyRange(model: DocumentModel, startOffset: number, endOffset: number): DocumentSelection | undefined {
  const body = flattenDocumentBody(model)
  if (!Number.isSafeInteger(startOffset) || !Number.isSafeInteger(endOffset) || startOffset < 0 || endOffset < startOffset
    || endOffset > body.text.length || !isDocumentTextBoundary(body.text, startOffset) || !isDocumentTextBoundary(body.text, endOffset)) {
    throw new RangeError(`Invalid document text range [${startOffset}, ${endOffset}) for body length ${body.text.length}.`)
  }
  const start = pointAtOffset(body.segments, startOffset), end = pointAtOffset(body.segments, endOffset)
  if (!start || !end) return undefined
  const first = targetContainer(model, start), last = targetContainer(model, end)
  if (!first || !last) throw new RangeError('文档正文范围的内容已不存在。')
  if (first.blocks !== last.blocks) throw new RangeError('A document text range cannot cross independent table cells.')
  if (!sameTarget(start, end) && first.blocks.slice(first.index + 1, last.index)
    .some(block => !['paragraph', 'heading', 'bullet', 'numbered', 'checklist', 'quote'].includes(block.type))) {
    throw new RangeError('A document text range cannot cross non-text blocks.')
  }
  return { anchor: start, focus: end }
}

function pointAtOffset(segments: readonly DocumentBodySegment[], offset: number): DocumentPoint | undefined {
  const segment = segments.find(candidate => offset <= candidate.endOffset) ?? segments.at(-1)
  if (!segment) return undefined
  const { blockId, row, col, cellBlockId } = segment
  return { blockId, ...(row === undefined ? {} : { row }), ...(col === undefined ? {} : { col }),
    ...(cellBlockId === undefined ? {} : { cellBlockId }), offset: Math.max(0, Math.min(segment.text.length, offset - segment.startOffset)) }
}

function targetContainer(model: DocumentModel, point: DocumentPoint): { blocks: DocumentBlock[]; index: number } | undefined {
  const block = findDocumentBlock(model, point.blockId)
  if (!block) return undefined
  const blocks = findDocumentBlockContainer(model, point.cellBlockId ?? point.blockId)
  if (!blocks) return undefined
  const index = blocks.findIndex(candidate => candidate.id === (point.cellBlockId ?? point.blockId))
  return index < 0 ? undefined : { blocks, index }
}
