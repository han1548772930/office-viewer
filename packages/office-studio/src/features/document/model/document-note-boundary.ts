import type { InlineRun } from '../../../domain/model'
import { documentRunHasNoteReference } from './document-note-reference'

/** Temporary paragraph geometry. The run remains the authoritative note entity. */
export interface DocumentNoteReferenceOffset {
  runId: string
  startOffset: number
  endOffset: number
}

export function documentNoteReferenceOffsets(runs: readonly Pick<InlineRun, 'id' | 'text' | 'metadata'>[]): DocumentNoteReferenceOffset[] {
  const references: DocumentNoteReferenceOffset[] = []
  let offset = 0
  for (const run of runs) {
    const endOffset = offset + run.text.length
    if (documentRunHasNoteReference(run)) references.push({ runId: run.id, startOffset: offset, endOffset })
    offset = endOffset
  }
  return references
}

/** Rank counts every reference since paragraph start, including visible labels.
 * Only empty labels allow multiple native positions at one display offset. */
export function documentNoteReferenceBoundaryRanks(references: readonly DocumentNoteReferenceOffset[], offset: number): { first: number; last: number } {
  let first = 0, tied = 0
  for (const reference of references) {
    if (reference.startOffset < offset && offset < reference.endOffset) throw new RangeError('文字范围端点不能位于脚注或尾注引用字符内部。')
    if (reference.endOffset < offset || reference.endOffset === offset && reference.startOffset < offset) first++
    else if (reference.startOffset === offset && reference.endOffset === offset) tied++
  }
  return { first, last: first + tied }
}

export function assertDocumentNoteReferenceBoundaryRank(references: readonly DocumentNoteReferenceOffset[], offset: number, rank?: number): number {
  const bounds = documentNoteReferenceBoundaryRanks(references, offset)
  if (rank === undefined) return bounds.first
  if (!Number.isSafeInteger(rank) || rank < bounds.first || rank > bounds.last) throw new RangeError('文字范围的累计注释引用序号与段落位置不匹配。')
  return rank
}

/** Surviving neighboring reference IDs preserve a boundary between empty
 * markers while edits change their cumulative paragraph rank. */
export function remapDocumentNoteReferenceBoundaryRank(
  source: readonly DocumentNoteReferenceOffset[], destination: readonly DocumentNoteReferenceOffset[],
  sourceOffset: number, destinationOffset: number, rank: number | undefined,
): number | undefined {
  const resolved = assertDocumentNoteReferenceBoundaryRank(source, sourceOffset, rank)
  const bounds = documentNoteReferenceBoundaryRanks(destination, destinationOffset)
  const next = source[resolved], previous = source[resolved - 1]
  const output = (value: number): number | undefined => rank === undefined && value === bounds.first ? undefined : value
  if (next) {
    const index = destination.findIndex(reference => reference.runId === next.runId)
    if (index >= bounds.first && index <= bounds.last) return output(index)
  }
  if (previous) {
    const index = destination.findIndex(reference => reference.runId === previous.runId)
    if (index >= 0 && index + 1 >= bounds.first && index + 1 <= bounds.last) return output(index + 1)
  }
  const original = documentNoteReferenceBoundaryRanks(source, sourceOffset)
  return output(bounds.first + Math.min(resolved - original.first, bounds.last - bounds.first))
}

/** An empty marker gaining a label also moves native positions that follow it.
 * UTF-16 delta alone cannot distinguish the two sides of that old empty run. */
export function remapDocumentNoteReferenceBoundary(
  source: readonly DocumentNoteReferenceOffset[], destination: readonly DocumentNoteReferenceOffset[],
  sourceOffset: number, mappedOffset: number, rank: number | undefined,
): { offset: number; rank?: number } {
  const resolved = assertDocumentNoteReferenceBoundaryRank(source, sourceOffset, rank)
  const next = source[resolved], previous = source[resolved - 1]
  const tied = (reference: DocumentNoteReferenceOffset | undefined) => reference?.startOffset === sourceOffset && reference.endOffset === sourceOffset
  const after = tied(next) ? destination.find(reference => reference.runId === next!.runId) : undefined
  const before = tied(previous) ? destination.find(reference => reference.runId === previous!.runId) : undefined
  const offset = after?.startOffset ?? before?.endOffset ?? mappedOffset
  const count = remapDocumentNoteReferenceBoundaryRank(source, destination, sourceOffset, offset, rank)
  return { offset, ...(count === undefined ? {} : { rank: count }) }
}
