import type { DocumentNoteReferenceOffset } from './document-note-boundary'
import { assertDocumentNoteReferenceBoundaryRank } from './document-note-boundary'
import { assertDocumentNoteReferenceOffsets, assertDocumentNoteReferenceSourceOffsets } from './document-note-boundary-edit'
import type { DocumentNoteBoundaryParagraph } from './document-note-boundary-edit'

export interface DocumentNoteInsertionAffinity<T> {
  sourceOffset: number
  sourceRank: number
  before: T
  after: T
}

export interface DocumentNoteEmptyReplacementAffinity<T> {
  sourceOffset: number
  startRank: number
  endRank: number
  before: T
  after: T
}

/** The complete destination owns each note run once, including runs outside
 * the endpoints whose neighboring IDs happen to be used during reanchoring. */
export function assertDocumentNoteReferenceParagraphIds(paragraphs: readonly Pick<DocumentNoteBoundaryParagraph, 'noteReferences'>[]): void {
  const ids = new Set<string>()
  for (const paragraph of paragraphs) for (const reference of paragraph.noteReferences) {
    if (ids.has(reference.runId)) throw new RangeError('注释引用 run 在多个目的段落重复。')
    ids.add(reference.runId)
  }
}

/** A zero-display deletion still has a nonempty native range. It cannot skip a
 * surviving empty atom and claim that both removed atoms were one selection. */
export function documentNoteReferenceEmptyRemovalRange(source: readonly DocumentNoteReferenceOffset[],
  survivingIds: ReadonlySet<string>, offset: number): { startRank: number; endRank: number } | undefined {
  assertDocumentNoteReferenceSourceOffsets(source)
  if (!Number.isSafeInteger(offset) || offset < 0) throw new RangeError('空注释引用删除位置无效。')
  assertDocumentNoteReferenceBoundaryRank(source, offset)
  const removed: number[] = []
  for (const [index, reference] of source.entries()) if (!survivingIds.has(reference.runId)) {
    if (reference.startOffset !== offset || reference.endOffset !== offset) throw new RangeError('零文字长度编辑不能删除其他位置的注释引用。')
    removed.push(index)
  }
  if (!removed.length) return undefined
  if (removed.some((index, position) => index !== removed[0]! + position)) throw new RangeError('注释引用删除范围必须连续。')
  return { startRank: removed[0]!, endRank: removed.at(-1)! + 1 }
}

/** Resolve ordinary insertion from real future atoms. A surviving reference
 * whose label changes length is derived renumbering and keeps ID affinity. */
export function documentNoteReferenceInsertionAffinity<T extends DocumentNoteBoundaryParagraph>(
  source: readonly DocumentNoteReferenceOffset[], paragraphs: readonly T[], offset: number, insertedLength: number,
): DocumentNoteInsertionAffinity<{ paragraph: T; offset: number; rank: number }> | undefined {
  if (!paragraphs.length || !Number.isSafeInteger(offset) || offset < 0
    || !Number.isSafeInteger(insertedLength) || insertedLength < 0) throw new RangeError('注释引用插入位置或长度无效。')
  assertDocumentNoteReferenceSourceOffsets(source)
  for (const paragraph of paragraphs) assertDocumentNoteReferenceOffsets(paragraph.noteReferences, paragraph.text)
  assertDocumentNoteReferenceParagraphIds(paragraphs)
  let base = 0, count = 0
  const outputs = paragraphs.map(paragraph => {
    const value = { paragraph, base, count }
    base += paragraph.text.length + 1; count += paragraph.noteReferences.length
    return value
  })
  const destination = outputs.flatMap(({ paragraph, base }) => paragraph.noteReferences.map(reference => ({ ...reference,
    startOffset: base + reference.startOffset, endOffset: base + reference.endOffset })))
  const sourceIds = new Set(source.map(reference => reference.runId)), destinationIds = new Set(destination.map(reference => reference.runId))
  if (source.some(reference => !destinationIds.has(reference.runId))) return undefined
  const surviving = destination.filter(reference => sourceIds.has(reference.runId))
  if (surviving.some((reference, index) => reference.runId !== source[index]?.runId)) throw new RangeError('注释引用插入不能重排已有引用。')
  if (surviving.some((reference, index) => reference.endOffset - reference.startOffset !== source[index]!.endOffset - source[index]!.startOffset)) return undefined
  const inserted = destination.flatMap((reference, index) => !sourceIds.has(reference.runId) ? [index] : [])
  if (!insertedLength && !inserted.length) return undefined
  if (inserted.some((index, position) => index !== inserted[0]! + position)) throw new RangeError('新注释引用必须属于同一连续插入范围。')
  // Empty references tied at the insertion offset can be on either side of
  // the input. Their actual displacement identifies the source rank.
  const firstMoved = surviving.findIndex((reference, index) => reference.startOffset !== source[index]!.startOffset
    || reference.endOffset !== source[index]!.endOffset)
  const sourceRank = inserted.length ? inserted[0]! : firstMoved < 0 ? source.length : firstMoved
  assertDocumentNoteReferenceBoundaryRank(source, offset, sourceRank)
  if (surviving.some((reference, index) => reference.startOffset !== source[index]!.startOffset + (index < sourceRank ? 0 : insertedLength)
    || reference.endOffset !== source[index]!.endOffset + (index < sourceRank ? 0 : insertedLength))
    || inserted.some(index => destination[index]!.startOffset < offset || destination[index]!.endOffset > offset + insertedLength)) {
    throw new RangeError('注释引用几何不是所声明的连续插入。')
  }
  const boundary = (absoluteOffset: number, rank: number) => {
    const output = outputs.find(value => absoluteOffset >= value.base && absoluteOffset <= value.base + value.paragraph.text.length)
    if (!output) throw new RangeError('注释引用插入端点无法映射到真实目的段落。')
    const localOffset = absoluteOffset - output.base, localRank = rank - output.count
    assertDocumentNoteReferenceBoundaryRank(output.paragraph.noteReferences, localOffset, localRank)
    return { paragraph: output.paragraph, offset: localOffset, rank: localRank }
  }
  return { sourceOffset: offset, sourceRank, before: boundary(offset, sourceRank),
    after: boundary(offset + insertedLength, sourceRank + inserted.length) }
}

/** Positions strictly inside removed empty atoms collapse onto the actual
 * replacement boundary, even when neither neighboring source ID survives. */
export function documentNoteReferenceEmptyReplacementAffinity<T extends DocumentNoteBoundaryParagraph>(
  source: readonly DocumentNoteReferenceOffset[], paragraphs: readonly T[], offset: number, insertedLength: number,
  range: { startRank: number; endRank: number },
): DocumentNoteEmptyReplacementAffinity<{ paragraph: T; offset: number; rank: number }> {
  if (!paragraphs.length || !Number.isSafeInteger(offset) || offset < 0
    || !Number.isSafeInteger(insertedLength) || insertedLength < 0
    || !range || !Number.isSafeInteger(range.startRank) || !Number.isSafeInteger(range.endRank)
    || range.startRank < 0 || range.endRank <= range.startRank || range.endRank > source.length) {
    throw new RangeError('空注释引用替换参数无效。')
  }
  assertDocumentNoteReferenceSourceOffsets(source)
  assertDocumentNoteReferenceBoundaryRank(source, offset, range.startRank)
  assertDocumentNoteReferenceBoundaryRank(source, offset, range.endRank)
  if (source.slice(range.startRank, range.endRank).some(reference => reference.startOffset !== offset || reference.endOffset !== offset)) {
    throw new RangeError('空注释引用替换只能移除指定位置的空引用。')
  }
  for (const paragraph of paragraphs) assertDocumentNoteReferenceOffsets(paragraph.noteReferences, paragraph.text)
  assertDocumentNoteReferenceParagraphIds(paragraphs)
  let base = 0, count = 0
  const outputs = paragraphs.map(paragraph => {
    const output = { paragraph, base, count }
    base += paragraph.text.length + 1; count += paragraph.noteReferences.length
    return output
  })
  const destination = outputs.flatMap(({ paragraph, base }) => paragraph.noteReferences.map(reference => ({ ...reference,
    startOffset: base + reference.startOffset, endOffset: base + reference.endOffset })))
  const sourceIds = new Set(source.map(reference => reference.runId))
  const kept = source.filter((_, index) => index < range.startRank || index >= range.endRank)
  const surviving = destination.filter(reference => sourceIds.has(reference.runId))
  const inserted = destination.flatMap((reference, index) => !sourceIds.has(reference.runId) ? [index] : [])
  if (surviving.length !== kept.length || surviving.some((reference, index) => reference.runId !== kept[index]!.runId
    || reference.startOffset !== kept[index]!.startOffset + (index < range.startRank ? 0 : insertedLength)
    || reference.endOffset !== kept[index]!.endOffset + (index < range.startRank ? 0 : insertedLength))
    || inserted.some((index, position) => index !== range.startRank + position
      || destination[index]!.startOffset < offset || destination[index]!.endOffset > offset + insertedLength)) {
    throw new RangeError('空注释引用替换与所声明的连续范围不一致。')
  }
  const boundary = (absoluteOffset: number, rank: number) => {
    const output = outputs.find(value => absoluteOffset >= value.base && absoluteOffset <= value.base + value.paragraph.text.length)
    if (!output) throw new RangeError('空注释引用替换端点无法映射到目的段落。')
    const localOffset = absoluteOffset - output.base, localRank = rank - output.count
    assertDocumentNoteReferenceBoundaryRank(output.paragraph.noteReferences, localOffset, localRank)
    return { paragraph: output.paragraph, offset: localOffset, rank: localRank }
  }
  return { sourceOffset: offset, ...range, before: boundary(offset, range.startRank),
    after: boundary(offset + insertedLength, range.startRank + inserted.length) }
}

export function documentNoteEmptyReplacementBoundary<T>(affinity: DocumentNoteEmptyReplacementAffinity<T> | undefined,
  offset: number, rank: number, edge: 'start' | 'end'): T | undefined {
  if (!affinity || offset !== affinity.sourceOffset || rank < affinity.startRank || rank > affinity.endRank) return undefined
  return edge === 'start' ? rank === affinity.endRank ? affinity.after : affinity.before
    : rank === affinity.startRank ? affinity.before : affinity.after
}

/** Half-open ranges move their start after new content and keep their end
 * before it. A point moves as one entity; an explicit inside edit encloses it. */
export function documentNoteInsertionBoundary<T>(affinity: DocumentNoteInsertionAffinity<T> | undefined,
  offset: number, rank: number, edge: 'start' | 'end', point: boolean, inside = false): T | undefined {
  if (!affinity || offset !== affinity.sourceOffset || rank !== affinity.sourceRank) return undefined
  return inside ? edge === 'start' ? affinity.before : affinity.after
    : point || edge === 'start' ? affinity.after : affinity.before
}
