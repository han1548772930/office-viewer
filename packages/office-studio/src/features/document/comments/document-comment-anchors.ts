import type { DocumentCommentAnchorSegment, DocumentModel, DocumentThreadComment } from '../../../domain/model'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { isDocumentTextBoundary } from '../model/document-story'
import { assertDocumentNoteReferenceBoundaryRank, remapDocumentNoteReferenceBoundary, type DocumentNoteReferenceOffset } from '../model/document-note-boundary'

interface SourceParagraph {
  block: { id: string }
  absoluteStart: number
  absoluteEnd: number
  noteReferences?: readonly DocumentNoteReferenceOffset[]
}

interface DestinationParagraph extends SourceParagraph { text: string }

export interface DocumentCommentParagraphEdit {
  blockId: string
  startOffset: number
  endOffset: number
  insertedText: string
  /** A removed paragraph can collapse to a boundary in a surviving paragraph. */
  targetOffset?: number
}

function mappedAnchorRange(from: number, to: number, start: number, end: number, length: number): { start: number; end: number } {
  const delta = length - (end - start)
  if (from === to) {
    const point = from < start ? from : from >= end ? from + delta : start + Math.min(from - start, length)
    return { start: point, end: point }
  }
  if (start === end) {
    if (start <= from) return { start: from + length, end: to + length }
    return { start: from, end: start < to ? to + length : to }
  }
  if (end <= from) return { start: from + delta, end: to + delta }
  if (start >= to) return { start: from, end: to }
  return { start: from >= start ? start : from, end: to >= end ? to + delta : start + length }
}

function unicodeBoundary(text: string, offset: number, direction: -1 | 1): number {
  return isDocumentTextBoundary(text, offset) ? offset : offset + direction
}

function anchorsForMappedRange(
  anchor: DocumentCommentAnchorSegment,
  start: number,
  end: number,
  paragraphs: readonly DestinationParagraph[],
): DocumentCommentAnchorSegment[] {
  const first = paragraphs.find(paragraph => start >= paragraph.absoluteStart && start <= paragraph.absoluteEnd)
  const last = paragraphs.find(paragraph => end >= paragraph.absoluteStart && end <= paragraph.absoluteEnd)
  if (!first || !last) throw new RangeError('正文编辑后的批注锚点无法映射到文本段落。')
  if (start === end) {
    const offset = unicodeBoundary(first.text, start - first.absoluteStart, -1)
    return [{ ...anchor, blockId: first.block.id, startOffset: offset, endOffset: offset, text: '' }]
  }
  const result = paragraphs.flatMap(paragraph => {
    const from = Math.max(start, paragraph.absoluteStart), to = Math.min(end, paragraph.absoluteEnd)
    if (from >= to) return []
    const startOffset = unicodeBoundary(paragraph.text, from - paragraph.absoluteStart, -1)
    const endOffset = unicodeBoundary(paragraph.text, to - paragraph.absoluteStart, 1)
    return [{ ...anchor, blockId: paragraph.block.id, startOffset, endOffset, text: paragraph.text.slice(startOffset, endOffset) }]
  })
  // A comment whose complete text was removed remains a resolvable empty anchor.
  return result.length ? result : [{ ...anchor, blockId: first.block.id,
    startOffset: start - first.absoluteStart, endOffset: start - first.absoluteStart, text: '' }]
}

function mergeAnchors(
  anchors: readonly DocumentCommentAnchorSegment[],
  paragraphs: ReadonlyMap<string, DestinationParagraph>,
  preserved: ReadonlySet<DocumentCommentAnchorSegment>,
): DocumentCommentAnchorSegment[] {
  const result: DocumentCommentAnchorSegment[] = []
  let pending: DocumentCommentAnchorSegment[] = []
  const flush = () => {
    const groupStart = result.length
    const rank = (anchor: DocumentCommentAnchorSegment, endpoint: 'start' | 'end') => assertDocumentNoteReferenceBoundaryRank(
      paragraphs.get(anchor.blockId)!.noteReferences ?? [], anchor[`${endpoint}Offset`], anchor[`${endpoint}NoteReferenceCount`])
    pending.sort((left, right) => paragraphs.get(left.blockId)!.absoluteStart - paragraphs.get(right.blockId)!.absoluteStart
      || left.startOffset - right.startOffset || rank(left, 'start') - rank(right, 'start')
      || left.endOffset - right.endOffset || rank(left, 'end') - rank(right, 'end'))
    for (const anchor of pending) {
      const previous = result.length > groupStart ? result.at(-1) : undefined, paragraph = paragraphs.get(anchor.blockId)!
      if (previous && previous.blockId === anchor.blockId && (anchor.startOffset < previous.endOffset
        || anchor.startOffset === previous.endOffset && rank(anchor, 'start') <= rank(previous, 'end'))
        && previous.startOffset <= anchor.endOffset) {
        if (anchor.startOffset < previous.startOffset || anchor.startOffset === previous.startOffset && rank(anchor, 'start') < rank(previous, 'start')) {
          previous.startOffset = anchor.startOffset
          if (anchor.startNoteReferenceCount === undefined) delete previous.startNoteReferenceCount
          else previous.startNoteReferenceCount = anchor.startNoteReferenceCount
        }
        if (anchor.endOffset > previous.endOffset || anchor.endOffset === previous.endOffset && rank(anchor, 'end') > rank(previous, 'end')) {
          previous.endOffset = anchor.endOffset
          if (anchor.endNoteReferenceCount === undefined) delete previous.endNoteReferenceCount
          else previous.endNoteReferenceCount = anchor.endNoteReferenceCount
        }
        previous.text = paragraph.text.slice(previous.startOffset, previous.endOffset)
      } else result.push({ ...anchor })
    }
    pending = []
  }
  for (const anchor of anchors) {
    if (paragraphs.has(anchor.blockId) && !preserved.has(anchor)) pending.push(anchor)
    else { flush(); result.push({ ...anchor }) }
  }
  flush()
  return result
}

function transformKnownCommentAnchors(
  model: DocumentModel,
  sourceParagraphs: readonly SourceParagraph[],
  destinationParagraphs: readonly DestinationParagraph[],
  map: (anchor: DocumentCommentAnchorSegment, paragraph: SourceParagraph) => DocumentCommentAnchorSegment[],
  timestamp: number,
): void {
  if (!Object.keys(model.comments ?? {}).length) return
  const updatedAt = new Date(timestamp).toISOString()
  const source = new Map(sourceParagraphs.map(paragraph => [paragraph.block.id, paragraph]))
  const destination = new Map(destinationParagraphs.map(paragraph => [paragraph.block.id, paragraph]))
  const plans: Array<{ comment: DocumentThreadComment; anchor: DocumentCommentAnchorSegment[] }> = []
  for (const comment of Object.values(model.comments ?? {})) {
    const preserved = new Set<DocumentCommentAnchorSegment>()
    const mapped = comment.anchor.flatMap(anchor => {
      const paragraph = source.get(anchor.blockId)
      if (!paragraph) {
        const unchanged = { ...anchor }
        preserved.add(unchanged)
        return [unchanged]
      }
      if (!Number.isSafeInteger(anchor.startOffset) || !Number.isSafeInteger(anchor.endOffset) || anchor.startOffset < 0
        || anchor.endOffset < anchor.startOffset || anchor.endOffset > paragraph.absoluteEnd - paragraph.absoluteStart) {
        throw new RangeError(`批注 ${comment.id} 的源 UTF-16 锚点无效。`)
      }
      const mapped = map(anchor, paragraph)
      return mapped.map((next, index) => {
        const { startNoteReferenceCount: _startRank, endNoteReferenceCount: _endRank, ...result } = next
        const target = destination.get(result.blockId)
        if (!target) throw new RangeError('批注编辑后的注释引用宿主不存在。')
        const references = target.noteReferences ?? []
        // A split keeps source bookmarks only at its outside endpoints.
        const first = index === 0 && paragraph.noteReferences
          ? remapDocumentNoteReferenceBoundary(paragraph.noteReferences, references, anchor.startOffset, result.startOffset, anchor.startNoteReferenceCount)
          : { offset: result.startOffset, rank: undefined }
        const last = index === mapped.length - 1 && paragraph.noteReferences
          ? remapDocumentNoteReferenceBoundary(paragraph.noteReferences, references, anchor.endOffset, result.endOffset, anchor.endNoteReferenceCount)
          : { offset: result.endOffset, rank: undefined }
        const start = assertDocumentNoteReferenceBoundaryRank(references, first.offset, first.rank)
        const end = assertDocumentNoteReferenceBoundaryRank(references, last.offset, last.rank)
        result.startOffset = first.offset
        result.endOffset = Math.max(first.offset, last.offset)
        result.text = target.text.slice(result.startOffset, result.endOffset)
        if (first.offset > last.offset || first.offset === last.offset && end < start) {
          // Deleting a range collapses both bookmarks to the same current position.
          return { ...result, ...(first.rank === undefined ? {} : { startNoteReferenceCount: first.rank, endNoteReferenceCount: first.rank }) }
        }
        return { ...result, ...(first.rank === undefined ? {} : { startNoteReferenceCount: first.rank }),
          ...(last.rank === undefined ? {} : { endNoteReferenceCount: last.rank }) }
      })
    })
    const anchor = mergeAnchors(mapped, destination, preserved)
    if (JSON.stringify(anchor) !== JSON.stringify(comment.anchor)) plans.push({ comment, anchor })
  }
  for (const { comment, anchor } of plans) {
    comment.anchor = anchor
    comment.referenceText = anchor.map(segment => segment.text).join('\n')
    comment.updatedAt = updatedAt
  }
}

/** Split/merge logical anchors after an edit, preparing every thread before publication. */
export function transformDocumentCommentAnchorsForBodyRangeEdit(
  model: DocumentModel,
  sourceParagraphs: readonly SourceParagraph[],
  destinationParagraphs: readonly DestinationParagraph[],
  startOffset: number,
  endOffset: number,
  insertedText: string,
  timestamp = Date.now(),
): void {
  transformKnownCommentAnchors(model, sourceParagraphs, destinationParagraphs, (anchor, paragraph) => {
    const range = mappedAnchorRange(paragraph.absoluteStart + anchor.startOffset, paragraph.absoluteStart + anchor.endOffset,
      startOffset, endOffset, insertedText.length)
    return anchorsForMappedRange(anchor, range.start, range.end, destinationParagraphs)
  }, timestamp)
}

/** A structural rebuild keeps anchors on surviving paragraph IDs and maps only changed row text. */
export function transformDocumentCommentAnchorsForParagraphEdits(
  model: DocumentModel,
  sourceParagraphs: readonly SourceParagraph[],
  destinationParagraphs: readonly DestinationParagraph[],
  edits: ReadonlyMap<string, DocumentCommentParagraphEdit | readonly DocumentCommentParagraphEdit[]>,
  timestamp = Date.now(),
): void {
  const destination = new Map(destinationParagraphs.map(paragraph => [paragraph.block.id, paragraph]))
  transformKnownCommentAnchors(model, sourceParagraphs, destinationParagraphs, anchor => {
    const value = edits.get(anchor.blockId)
    if (!value) return [{ ...anchor }]
    const changes = 'blockId' in value ? [value] : value
    let range = { start: anchor.startOffset, end: anchor.endOffset }, blockId = anchor.blockId
    for (const edit of changes) {
      const mapped = mappedAnchorRange(range.start, range.end, edit.startOffset, edit.endOffset, edit.insertedText.length)
      range = { start: mapped.start + (edit.targetOffset ?? 0), end: mapped.end + (edit.targetOffset ?? 0) }
      blockId = edit.blockId
    }
    const paragraph = destination.get(blockId)
    if (!paragraph) throw new RangeError('正文重建后的批注目标段落不存在。')
    const offset = paragraph.absoluteStart
    return anchorsForMappedRange(anchor, offset + range.start, offset + range.end, destinationParagraphs)
  }, timestamp)
}

/** The paragraph's runs already contain the replacement; recover source geometry from its delta. */
export function transformDocumentCommentAnchorsForTextEdit(
  model: DocumentModel,
  blockId: string,
  startOffset: number,
  endOffset: number,
  insertedText: string,
  timestamp = Date.now(),
  sourceNoteReferences?: readonly DocumentNoteReferenceOffset[],
): void {
  if (!Object.keys(model.comments ?? {}).length) return
  const destination = documentTextFieldParagraphs(model)
  const paragraph = destination.find(paragraph => paragraph.block.id === blockId)
  if (!paragraph) return
  const delta = insertedText.length - (endOffset - startOffset)
  const source = destination.map(current => current.block.id === blockId
    ? { ...current, absoluteEnd: current.absoluteEnd - delta, noteReferences: sourceNoteReferences }
    : current.absoluteStart > paragraph.absoluteStart ? { ...current, absoluteStart: current.absoluteStart - delta, absoluteEnd: current.absoluteEnd - delta } : current)
  transformDocumentCommentAnchorsForBodyRangeEdit(model, source, destination,
    paragraph.absoluteStart + startOffset, paragraph.absoluteStart + endOffset, insertedText, timestamp)
}
