import type { DocumentNoteReferenceOffset } from './document-note-boundary'
import { assertDocumentNoteReferenceBoundaryRank, remapDocumentNoteReferenceBoundary } from './document-note-boundary'
import { isDocumentTextBoundary } from './document-story'

export function compareDocumentNoteBoundaries(leftOffset: number, leftRank: number, rightOffset: number, rightRank: number): number {
  return leftOffset - rightOffset || leftRank - rightRank
}

/** Source geometry has no text buffer here, but identities, ordering and safe
 * integer coordinates must still be valid before cumulative ranks are used. */
export function assertDocumentNoteReferenceSourceOffsets(references: readonly DocumentNoteReferenceOffset[]): void {
  const ids = new Set<string>()
  let end = 0
  for (const reference of references) {
    if (!reference || typeof reference.runId !== 'string' || !reference.runId || ids.has(reference.runId)
      || !Number.isSafeInteger(reference.startOffset) || !Number.isSafeInteger(reference.endOffset)
      || reference.startOffset < end || reference.endOffset < reference.startOffset) {
      throw new RangeError('注释引用的源几何或身份无效。')
    }
    ids.add(reference.runId); end = reference.endOffset
  }
}

/** Temporary geometry is checked against the intended text before a mutation publishes it. */
export function assertDocumentNoteReferenceOffsets(references: readonly DocumentNoteReferenceOffset[], text: string): void {
  const ids = new Set<string>()
  let previousEnd = 0
  for (const reference of references) {
    if (!reference || typeof reference.runId !== 'string' || !reference.runId || ids.has(reference.runId)
      || !Number.isSafeInteger(reference.startOffset) || !Number.isSafeInteger(reference.endOffset)
      || reference.startOffset < previousEnd || reference.endOffset < reference.startOffset
      || !isDocumentTextBoundary(text, reference.startOffset) || !isDocumentTextBoundary(text, reference.endOffset)) {
      throw new RangeError('编辑后的注释引用几何与真实段落不一致。')
    }
    ids.add(reference.runId); previousEnd = reference.endOffset
  }
}

/** Plain text replacement removes selected note atoms; following empty atoms retain their IDs. */
export function documentNoteReferenceOffsetsAfterTextEdit(
  references: readonly DocumentNoteReferenceOffset[], start: number, end: number, insertedLength: number,
  sourceNoteRanks?: DocumentNoteEditRanks,
): DocumentNoteReferenceOffset[] {
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start
    || !Number.isSafeInteger(insertedLength) || insertedLength < 0 || !Number.isSafeInteger(start + insertedLength)
    || sourceNoteRanks !== undefined && (!sourceNoteRanks || typeof sourceNoteRanks !== 'object' || Array.isArray(sourceNoteRanks))) {
    throw new RangeError('注释引用文字编辑参数无效。')
  }
  assertDocumentNoteReferenceSourceOffsets(references)
  const startRank = assertDocumentNoteReferenceBoundaryRank(references, start, sourceNoteRanks?.start)
  const endRank = assertDocumentNoteReferenceBoundaryRank(references, end, sourceNoteRanks?.end)
  if (compareDocumentNoteBoundaries(start, startRank, end, endRank) > 0) throw new RangeError('注释引用文字编辑参数无效。')
  const delta = insertedLength - (end - start)
  return references.flatMap((reference, index) => {
    const before = compareDocumentNoteBoundaries(reference.endOffset, index + 1, start, startRank) <= 0
    if (before) return [{ ...reference }]
    if (compareDocumentNoteBoundaries(reference.startOffset, index, end, endRank) >= 0) {
      const startOffset = reference.startOffset + delta, endOffset = reference.endOffset + delta
      if (!Number.isSafeInteger(startOffset) || !Number.isSafeInteger(endOffset) || startOffset < 0) throw new RangeError('编辑后的注释引用位置超出安全整数范围。')
      return [{ ...reference, startOffset, endOffset }]
    }
    return []
  })
}

export interface DocumentNoteBoundaryParagraph {
  id: string
  text: string
  noteReferences: readonly DocumentNoteReferenceOffset[]
}

/** Source ranks are paragraph-local, including all visible and empty notes. */
export interface DocumentNoteEditRanks {
  start: number
  end: number
}

/** A split/merge can move a tied empty run to another paragraph. Resolve its
 * surviving ID across the complete destination before recomputing the local rank. */
export function remapDocumentNoteBoundaryParagraph<T extends DocumentNoteBoundaryParagraph>(
  source: readonly DocumentNoteReferenceOffset[], sourceOffset: number, rank: number | undefined,
  paragraphs: readonly T[], fallback: { paragraph: T; offset: number },
): { paragraph: T; offset: number; rank?: number } {
  const resolved = assertDocumentNoteReferenceBoundaryRank(source, sourceOffset, rank)
  const next = source[resolved], previous = source[resolved - 1]
  const tied = (reference: DocumentNoteReferenceOffset | undefined) => reference?.startOffset === sourceOffset && reference.endOffset === sourceOffset
  let destination = fallback
  for (const [reference, after] of [[next, false], [previous, true]] as const) {
    if (!tied(reference)) continue
    const matches = paragraphs.flatMap(paragraph => paragraph.noteReferences.filter(candidate => candidate.runId === reference!.runId)
      .map(candidate => ({ paragraph, offset: after ? candidate.endOffset : candidate.startOffset })))
    if (matches.length > 1) throw new RangeError('注释引用 run 在编辑后的多个段落重复。')
    if (matches.length) { destination = matches[0]!; break }
  }
  const mapped = remapDocumentNoteReferenceBoundary(source, destination.paragraph.noteReferences, sourceOffset, destination.offset, rank)
  return { paragraph: destination.paragraph, offset: mapped.offset, ...(mapped.rank === undefined ? {} : { rank: mapped.rank }) }
}
