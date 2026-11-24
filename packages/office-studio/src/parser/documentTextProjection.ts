import type { DocumentBlock, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { runText, tableForBlock } from '../features/document/model/document-model'
import { documentTableCellContentContainers } from '../features/document/model/document-table-continuations'
import { isDocumentTextBoundary } from '../features/document/model/document-story'
import { array, record, text } from './validation'

interface AtomProjection { sourceStart: number; sourceEnd: number; targetStart: number; targetEnd: number; noteOrdinal?: number }
export interface ParsedDocumentTextProjection {
  changed: boolean
  sourceText: string
  targetText: string
  /** An interior position in an atomic formula has no editable text boundary. */
  mapBoundary(offset: number, noteReferenceCount?: number): number | undefined
}

/** Source XML retains literal UTF-16 coordinates. Native boundary ranks distinguish
 * ranges immediately before and after a reference whose source text is empty. */
export function parsedDocumentTextProjection(block: DocumentBlock): ParsedDocumentTextProjection {
  const targetText = runText(block.runs)
  const formulaRuns = block.runs.filter(run => run.metadata?.latexFormula !== undefined)
  const noteRuns = block.runs.filter(run => record(run.metadata?.noteReference).kind !== undefined || record(run.metadata?.note).type !== undefined)
  const bodyMarkers = block.runs.filter(run => run.metadata?.officeNoteBodyMarkerSource !== undefined)
  if (!formulaRuns.length && !noteRuns.length && !bodyMarkers.length) return { changed: false, sourceText: targetText, targetText,
    mapBoundary: offset => isDocumentTextBoundary(targetText, offset) ? offset : undefined }
  let valid = true
  const inlines = array(record(record(block.metadata.source).details).inlines)
    .map(record).filter(inline => inline.type === 'text' || inline.type === 'form_field')
  const groups: InlineRun[][] = inlines.map(() => [])
  const indexed = block.runs.some(run => run.metadata?.officeParsedRunSource !== undefined)
  let previousIndex = -1
  for (const [runIndex, run] of block.runs.entries()) {
    const source = record(run.metadata?.officeParsedRunSource)
    const index = indexed ? source.version === 1 && typeof source.index === 'number' ? source.index : -1
      : block.runs.length === inlines.length ? runIndex : -1
    if (!Number.isSafeInteger(index) || index < previousIndex || index < 0 || index >= groups.length) { valid = false; continue }
    groups[index]!.push(run); previousIndex = index
  }
  const atoms: AtomProjection[] = [], parts: string[] = []
  let sourceText = '', targetOffset = 0, noteOrdinal = 0, consumedPrefix = 0
  for (const [index, inline] of inlines.entries()) {
    const literal = text(inline.text), sourceStart = sourceText.length
    sourceText += literal
    const runs = groups[index]!, formula = runs.filter(run => run.metadata?.latexFormula !== undefined)
    const atom = formula.length ? runText(formula) : undefined
    if (atom !== undefined && (inline.type !== 'form_field' || !text(inline.blockId) || formula.length !== 1 || runs.length !== 1 || atom !== '\uFFFC')) valid = false
    const note = record(inline.noteReference), noteRun = note.kind !== undefined ? runs.find(run => record(run.metadata?.noteReference).kind !== undefined || record(run.metadata?.note).type !== undefined) : undefined
    const isNote = !!noteRun && (record(noteRun.metadata?.noteReference).kind === note.kind
      || record(noteRun.metadata?.note).type === note.kind)
    const bodyMarker = record(runs.find(run => run.metadata?.officeNoteBodyMarkerSource !== undefined)?.metadata?.officeNoteBodyMarkerSource)
    const bodyPrefix = typeof bodyMarker.length === 'number' && bodyMarker.version === 1 ? bodyMarker.length : 0
    if (bodyPrefix && (!Number.isSafeInteger(bodyPrefix) || bodyPrefix < 0 || typeof bodyMarker.text !== 'string'
      || bodyMarker.text.length !== bodyPrefix || !literal.slice(consumedPrefix).startsWith(bodyMarker.text))) valid = false
    let projected = atom ?? literal.slice(consumedPrefix + bodyPrefix)
    const literalPrefix = consumedPrefix
    consumedPrefix = 0
    if (bodyPrefix) atoms.push({ sourceStart: sourceStart + literalPrefix, sourceEnd: sourceStart + literalPrefix + bodyPrefix,
      targetStart: targetOffset, targetEnd: targetOffset })
    if (note.kind !== undefined) {
      noteOrdinal++
      if (!isNote || runs.length !== 1) valid = false
      else {
        projected = noteRun.text
        const merged = record(noteRun.metadata?.officeNoteReferenceMergedLiteral), next = inlines[index + 1]
        const length = typeof merged.endOffset === 'number' && merged.startOffset === 0 ? merged.endOffset : 0
        if (length) {
          if (!Number.isSafeInteger(length) || length < 0 || !next || typeof merged.text !== 'string'
            || !text(next.text).startsWith(merged.text) || length !== merged.text.length || runText(groups[index + 1] ?? []) !== text(next.text).slice(length)) valid = false
          else consumedPrefix = length
        }
        atoms.push({ sourceStart, sourceEnd: sourceText.length + consumedPrefix, targetStart: targetOffset,
          targetEnd: targetOffset + projected.length, noteOrdinal })
      }
    }
    if (atom !== undefined) {
      atoms.push({ sourceStart: sourceStart + literalPrefix, sourceEnd: sourceText.length, targetStart: targetOffset, targetEnd: targetOffset + projected.length })
    }
    if (!runs.length || runText(runs) !== projected) valid = false
    parts.push(projected); targetOffset += projected.length
  }
  valid &&= !consumedPrefix && parts.join('') === targetText
  return { changed: sourceText !== targetText, sourceText, targetText, mapBoundary(offset, noteReferenceCount) {
    if (!valid || !isDocumentTextBoundary(sourceText, offset) || noteReferenceCount !== undefined && (!Number.isSafeInteger(noteReferenceCount) || noteReferenceCount < 0 || noteReferenceCount > noteOrdinal)) return undefined
    let delta = 0
    for (const atom of atoms) {
      if (offset < atom.sourceStart) break
      if (atom.sourceStart === atom.sourceEnd && offset === atom.sourceStart) {
        if (atom.targetStart === atom.targetEnd) { delta = atom.targetEnd - atom.sourceEnd; continue }
        if (atom.noteOrdinal === undefined || noteReferenceCount === undefined) return undefined
        if (noteReferenceCount < atom.noteOrdinal) return atom.targetStart
        delta = atom.targetEnd - atom.sourceEnd
        continue
      }
      if (offset === atom.sourceStart) return atom.noteOrdinal !== undefined && noteReferenceCount !== undefined && noteReferenceCount >= atom.noteOrdinal ? atom.targetEnd : atom.targetStart
      // Continue through adjacent atoms at the same source boundary. Multiple
      // zero-length references can share an offset and still have distinct ranks.
      if (offset === atom.sourceEnd) { delta = atom.targetEnd - atom.sourceEnd; continue }
      if (offset < atom.sourceEnd) return undefined
      delta = atom.targetEnd - atom.sourceEnd
    }
    const mapped = offset + delta
    return isDocumentTextBoundary(targetText, mapped) ? mapped : undefined
  } }
}

/** Map a complete host before its field validator or story projection runs,
 * including fields whose result ends in another paragraph or a table cell. */
export function projectParsedDocumentFieldRanges(blocks: readonly DocumentBlock[]): void {
  const paragraphs = new Map<string, DocumentBlock>()
  const visit = (items: readonly DocumentBlock[]) => {
    for (const block of items) {
      if (paragraphs.has(block.id)) throw new TypeError('原生字段文字投影的块身份重复。')
      paragraphs.set(block.id, block)
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) visit(content)
    }
  }
  visit(blocks)
  const projections = new Map(Array.from(paragraphs, ([id, block]) => [id, parsedDocumentTextProjection(block)]))
  const sourceIds = new Map<string, string | undefined>()
  for (const block of paragraphs.values()) {
    const id = text(record(block.metadata.source).id)
    if (id) sourceIds.set(id, sourceIds.has(id) ? undefined : block.id)
  }
  for (const block of paragraphs.values()) {
    const sourceCollection = record(record(record(block.metadata.source).details).documentTextFields)
    if (!Object.keys(sourceCollection).length) continue
    if (sourceCollection.version !== 1 || !Array.isArray(sourceCollection.fields)) throw new TypeError('原生字段文字投影的集合无效。')
    const collection = cloneModel(sourceCollection)
    const currentFields = array(record(block.metadata.documentTextFields).fields).map(record)
    collection.fields = sourceCollection.fields.map(value => {
      const field = record(cloneModel(value)), sourceId = text(field.id)
      const current = currentFields.find(candidate => record(record(candidate.metadata).officeParsedFieldSource).id === sourceId || candidate.id === sourceId)
      const sourceEndId = typeof field.endBlockId === 'string' ? field.endBlockId : block.id
      const endId = paragraphs.has(sourceEndId) ? sourceEndId : sourceIds.get(sourceEndId)
      const first = projections.get(block.id)!, last = endId === undefined ? undefined : projections.get(endId)
      if (!last || typeof field.startOffset !== 'number' || typeof field.endOffset !== 'number') throw new RangeError('原生字段文字投影的宿主或偏移无效。')
      const count = (value: unknown): number | undefined => typeof value === 'number' ? value : undefined
      const startOffset = first.mapBoundary(field.startOffset, count(field.startNoteReferenceCount)), endOffset = last.mapBoundary(field.endOffset, count(field.endNoteReferenceCount))
      if (startOffset === undefined || endOffset === undefined) throw new RangeError('原生字段结果边界缺少有效的公式或注释引用来源映射。')
      return { ...field, ...(typeof current?.id === 'string' ? { id: current.id } : {}), startOffset, endOffset,
        ...(field.endBlockId !== undefined ? { endBlockId: endId } : {}),
        metadata: { ...record(field.metadata), ...record(current?.metadata), officeParsedFieldSource: { version: 1, id: sourceId } } }
    })
    block.metadata.documentTextFields = collection
  }
}
