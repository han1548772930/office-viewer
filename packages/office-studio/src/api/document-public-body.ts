import type { DocumentModel, DocumentStoryContent } from '../domain/model'
import { documentDefaultStyle } from '../features/document/model/document-default-style'
import type { DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { isDocumentTextBoundary } from '../features/document/model/document-story'
import { sameTarget, textTargets, type DocumentPoint, type DocumentSelection } from '../features/document/model/document-model'
import { documentModelForBlockHost, resolveDocumentBlockHost } from '../features/document/services/document-block-host'
import { findDocumentParagraphs, flattenDocumentBody, type DocumentParagraphInfo } from '../features/document/services/document-facade'
import { encodeDocumentBodyWithPositions, documentBodyContentControls } from './document-body-codec'
import { documentBodyLocalOffset, type DocumentBodyEncodeResult } from './document-body-types'
import { documentSelectionAtBodyRange } from './document-indexed-input'

export interface DocumentPublicBodyIndex extends DocumentBodyEncodeResult { model: DocumentModel }

/** Public stream coordinates include tables, fields, ranges and one-token note references. */
export function documentPublicBodyIndex(model: DocumentModel, host?: DocumentAuxiliaryBlockHost): DocumentPublicBodyIndex {
  const scoped = host ? documentModelForBlockHost(model, host) : model
  const content: DocumentStoryContent = host ? resolveDocumentBlockHost(model, host).content
    : { blocks: model.blocks, runs: [], paragraphs: [], metadata: model.metadata }
  return { ...encodeDocumentBodyWithPositions(content, documentDefaultStyle(scoped), documentBodyContentControls(content, scoped.contentControls)), model: scoped }
}

export function assertDocumentPublicRange(index: DocumentPublicBodyIndex, start: number, end: number): void {
  const stream = index.data.body.dataStream
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > stream.length
    || !isDocumentTextBoundary(stream, start) || !isDocumentTextBoundary(stream, end)) throw new RangeError('文档公开范围的 UTF-16 边界无效。')
}

function pointAt(index: DocumentPublicBodyIndex, offset: number, affinity: 'start' | 'end', allowStructure: boolean): DocumentPoint | undefined {
  let position = index.positions.find(position => offset >= position.start && offset <= position.end)
  if (!position && allowStructure) position = affinity === 'start' ? index.positions.find(position => position.start > offset) ?? index.positions.at(-1)
    : [...index.positions].reverse().find(position => position.end < offset) ?? index.positions[0]
  if (!position) return undefined
  const target = textTargets(index.model).find(target => (target.cellBlockId ?? target.blockId) === position.block.id)
  return target ? { ...target, offset: documentBodyLocalOffset(position, offset, affinity) } : undefined
}

/** Read/style adapters may cover structural gaps; text editing still validates actual block containers. */
export function documentPublicRangeToFlat(index: DocumentPublicBodyIndex, start: number, end: number): { startOffset: number; endOffset: number } {
  assertDocumentPublicRange(index, start, end)
  const anchor = pointAt(index, start, 'start', true), focus = pointAt(index, end, 'end', true), flat = flattenDocumentBody(index.model)
  if (!anchor || !focus) return { startOffset: 0, endOffset: 0 }
  const first = flat.segments.find(segment => sameTarget(segment, anchor)), last = flat.segments.find(segment => sameTarget(segment, focus))
  if (!first || !last) throw new RangeError('文档公开范围的目标段落已不存在。')
  const a = first.startOffset + anchor.offset, b = last.startOffset + focus.offset
  const intersectsText = start < end && index.positions.some(position => position.spans.some(span =>
    span.targetEnd > span.targetStart && span.sourceStart < end && span.sourceEnd > start))
  // A range containing only delimiters must not expand into adjacent text.
  return { startOffset: a, endOffset: intersectsText && b >= a ? b : a }
}

export function documentSelectionAtPublicBodyRange(model: DocumentModel, start: number, end: number,
  host?: DocumentAuxiliaryBlockHost): DocumentSelection | undefined {
  const index = documentPublicBodyIndex(model, host)
  assertDocumentPublicRange(index, start, end)
  const anchor = pointAt(index, start, 'start', false), focus = pointAt(index, end, 'end', false)
  if (!anchor || !focus) return undefined
  const flat = flattenDocumentBody(index.model)
  const first = flat.segments.find(segment => sameTarget(segment, anchor)), last = flat.segments.find(segment => sameTarget(segment, focus))
  if (!first || !last) return undefined
  return documentSelectionAtBodyRange(index.model, first.startOffset + anchor.offset, last.startOffset + focus.offset)
}

export function documentPublicParagraphs(index: DocumentPublicBodyIndex): DocumentParagraphInfo[] {
  const positions = new Map(index.positions.map(position => [position.block.id, position]))
  return findDocumentParagraphs(index.model, {}).flatMap(info => {
    const position = positions.get(info.paragraphId)
    return position ? [{ ...info, startOffset: position.start, endOffset: position.end }] : []
  })
}

export function documentPublicOffsetAtFlat(index: DocumentPublicBodyIndex, offset: number, affinity: 'start' | 'end' = 'start'): number {
  const flat = flattenDocumentBody(index.model), segment = flat.segments.find(segment => offset <= segment.endOffset) ?? flat.segments.at(-1)
  const position = segment && index.positions.find(position => position.block.id === (segment.cellBlockId ?? segment.blockId))
  if (!segment || !position) return 0
  const local = Math.max(0, Math.min(segment.text.length, offset - segment.startOffset))
  const spans = position.spans.filter(span => span.targetEnd > span.targetStart)
  for (const span of spans) {
    if (local > span.targetStart && local < span.targetEnd) {
      if (span.targetEnd - span.targetStart === span.sourceEnd - span.sourceStart) return span.sourceStart + local - span.targetStart
      return affinity === 'start' ? span.sourceStart : span.sourceEnd
    }
  }
  const next = spans.find(span => span.targetStart === local), previous = [...spans].reverse().find(span => span.targetEnd === local)
  if (affinity === 'start') return next?.sourceStart ?? previous?.sourceEnd ?? position.end
  return previous?.sourceEnd ?? next?.sourceStart ?? position.start
}
