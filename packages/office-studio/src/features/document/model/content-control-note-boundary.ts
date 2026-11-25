import type { DocumentContentControl, InlineRun } from '../../../domain/model'
import { documentId, runText } from './document-model'
import { assertDocumentNoteReferenceBoundaryRank, documentNoteReferenceOffsets, type DocumentNoteReferenceOffset } from './document-note-boundary'
import { compareDocumentNoteBoundaries } from './document-note-boundary-edit'
import { documentRunHasNoteReference } from './document-note-reference'
import { isDocumentTextBoundary } from './document-story'

export function documentContentControlNoteBoundaryRanks(control: Pick<DocumentContentControl,
  'startOffset' | 'endOffset' | 'startNoteReferenceCount' | 'endNoteReferenceCount'>,
references: readonly DocumentNoteReferenceOffset[]): { start: number; end: number } {
  const start = assertDocumentNoteReferenceBoundaryRank(references, control.startOffset, control.startNoteReferenceCount)
  const end = assertDocumentNoteReferenceBoundaryRank(references, control.endOffset, control.endNoteReferenceCount)
  if (compareDocumentNoteBoundaries(control.startOffset, start, control.endOffset, end) > 0) throw new RangeError('内容控件的累计注释引用端点顺序相反。')
  return { start, end }
}

/** Replacing the input host discards the old host's cumulative reference rank. */
export function reanchorDocumentContentControlNoteParagraph(control: DocumentContentControl, blockId: string, runs: readonly InlineRun[]): void {
  control.blockId = blockId
  control.startOffset = 0; control.endOffset = runText(runs).length
  control.startNoteReferenceCount = 0; control.endNoteReferenceCount = documentNoteReferenceOffsets(runs).length
}

export function documentContentControlNoteRangeContains(parent: DocumentContentControl, child: DocumentContentControl,
  references: readonly DocumentNoteReferenceOffset[]): boolean {
  if (parent.blockId !== child.blockId) return false
  const outer = documentContentControlNoteBoundaryRanks(parent, references), inner = documentContentControlNoteBoundaryRanks(child, references)
  return compareDocumentNoteBoundaries(parent.startOffset, outer.start, child.startOffset, inner.start) <= 0
    && compareDocumentNoteBoundaries(child.endOffset, inner.end, parent.endOffset, outer.end) <= 0
}

export function documentContentControlNoteRangesOverlap(left: DocumentContentControl, right: DocumentContentControl,
  references: readonly DocumentNoteReferenceOffset[]): boolean {
  if (left.blockId !== right.blockId) return false
  const first = documentContentControlNoteBoundaryRanks(left, references), second = documentContentControlNoteBoundaryRanks(right, references)
  return compareDocumentNoteBoundaries(left.startOffset, first.start, right.endOffset, second.end) < 0
    && compareDocumentNoteBoundaries(right.startOffset, second.start, left.endOffset, first.end) < 0
}

export function assignDocumentContentControlNoteBoundaries(control: DocumentContentControl,
  start: { offset: number; rank?: number }, end: { offset: number; rank?: number }): void {
  control.startOffset = start.offset; control.endOffset = end.offset
  delete control.startNoteReferenceCount; delete control.endNoteReferenceCount
  if (start.rank !== undefined) control.startNoteReferenceCount = start.rank
  if (end.rank !== undefined) control.endNoteReferenceCount = end.rank
}

/** A point child at its enclosing wrapper's end follows that end when the
 * insertion is outside the wrapper. Independent points keep their own affinity. */
export function documentContentControlPointBeforeInsertion(control: DocumentContentControl,
  controls: ReadonlyMap<string, DocumentContentControl>, references: readonly DocumentNoteReferenceOffset[],
  offset: number, rank: number, inside: ReadonlySet<string>): boolean {
  const bounds = documentContentControlNoteBoundaryRanks(control, references)
  if (inside.has(control.id) || control.startOffset !== offset || bounds.start !== rank
    || compareDocumentNoteBoundaries(control.startOffset, bounds.start, control.endOffset, bounds.end) !== 0) return false
  const seen = new Set<string>([control.id])
  let parentId = control.parentId
  while (parentId) {
    if (seen.has(parentId)) throw new RangeError('内容控件插入边界的父级关系循环。')
    seen.add(parentId)
    const parent = controls.get(parentId)
    if (!parent) throw new RangeError('内容控件插入边界的父级不存在。')
    if (inside.has(parent.id) || parent.blockRange || parent.tableRowRange || parent.blockId !== control.blockId) return false
    const enclosing = documentContentControlNoteBoundaryRanks(parent, references)
    if (parent.endOffset !== offset || enclosing.end !== rank) return false
    if (compareDocumentNoteBoundaries(parent.startOffset, enclosing.start, parent.endOffset, enclosing.end) < 0) return true
    parentId = parent.parentId
  }
  return false
}

/** A SDT may select only one of several empty references at the same text
 * offset. Numeric slicing cannot express its native start/end ranks. */
export function partitionDocumentContentControlNoteRuns<T extends InlineRun>(runs: readonly T[], control: Pick<DocumentContentControl,
  'startOffset' | 'endOffset' | 'startNoteReferenceCount' | 'endNoteReferenceCount'>): { before: T[]; selected: T[]; after: T[] } {
  const text = runText(runs), references = documentNoteReferenceOffsets(runs), range = documentContentControlNoteBoundaryRanks(control, references)
  if (!isDocumentTextBoundary(text, control.startOffset) || !isDocumentTextBoundary(text, control.endOffset)) throw new RangeError('内容控件字符范围不是合法的 UTF-16 边界。')
  const before: T[] = [], selected: T[] = [], after: T[] = []
  let offset = 0, rank = 0
  for (const run of runs) {
    const note = documentRunHasNoteReference(run), from = offset, fromRank = rank, to = from + run.text.length
    offset = to; if (note) rank++
    if (!run.text.length && !note) {
      if (compareDocumentNoteBoundaries(from, fromRank, control.startOffset, range.start) < 0) before.push(run)
      else if (compareDocumentNoteBoundaries(from, fromRank, control.endOffset, range.end) >= 0) after.push(run)
      else selected.push(run)
      continue
    }
    if (compareDocumentNoteBoundaries(to, rank, control.startOffset, range.start) <= 0) { before.push(run); continue }
    if (compareDocumentNoteBoundaries(from, fromRank, control.endOffset, range.end) >= 0) { after.push(run); continue }
    if (note || run.metadata?.latexFormula !== undefined) {
      if (compareDocumentNoteBoundaries(from, fromRank, control.startOffset, range.start) < 0
        || compareDocumentNoteBoundaries(to, rank, control.endOffset, range.end) > 0) throw new RangeError('内容控件值替换或复制不能切开注释引用或公式实体。')
      selected.push(run); continue
    }
    const a = Math.max(from, control.startOffset), b = Math.min(to, control.endOffset)
    const prefix = a > from, content = b > a
    if (prefix) before.push({ ...run, text: run.text.slice(0, a - from) })
    if (content) selected.push({ ...run, id: prefix ? documentId('run') : run.id, text: run.text.slice(a - from, b - from) })
    if (b < to) after.push({ ...run, id: prefix || content ? documentId('run') : run.id, text: run.text.slice(b - from) })
  }
  return { before, selected, after }
}
