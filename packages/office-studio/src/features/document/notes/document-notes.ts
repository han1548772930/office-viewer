import type { DocumentBlock, DocumentModel, DocumentNote, DocumentStoryContent } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { tableForBlock, type DocumentPoint } from '../model/document-model'
import { documentNoteReferenceRunMatches } from '../model/document-note-reference'
import { projectDocumentStoryBlocks } from '../model/document-story-blocks'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { documentNoteReferenceLocation, documentNoteReferenceRecords } from './document-note-references'
import { documentNoteContentText } from './document-note-content'
import { formatDocumentNoteNumber } from './document-note-numbering'
import { renumberDocumentNotes } from './document-note-reference-sync'

export type { DocumentNoteKind, DocumentNoteInsertOptions, DocumentNotePatch, DocumentNoteMutationRequest, DocumentNoteDecorationControls } from './document-note-types'
export { documentNoteReferenceLocation, getDocumentNotesInReferenceOrder } from './document-note-references'
export { renumberDocumentNotes } from './document-note-reference-sync'
export { insertDocumentNote, updateDocumentNote, patchDocumentNote, deleteDocumentNote, documentSupportsNotes } from './document-note-editing'
export { prepareDocumentNoteMutation, type DocumentNoteMutationPlan } from './document-note-mutations'

function rawRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
}

function rawNoteId(value: unknown, fallback: string): string {
  const raw = rawRecord(value), id = raw.id ?? raw.noteId
  if (typeof id === 'number' && !Number.isSafeInteger(id)) throw new TypeError('导入的注释编号不是安全整数。')
  return typeof id === 'string' || typeof id === 'number' ? String(id) : fallback
}

function rawNoteText(value: unknown): string {
  const raw = rawRecord(value)
  for (const candidate of [raw.text, raw.content, raw.body]) {
    if (typeof candidate === 'string') return candidate
    const nested = rawRecord(candidate)
    if (typeof nested.text === 'string') return nested.text
    if (typeof nested.dataStream === 'string') return nested.dataStream.replace(/[\r\n\u0000]+$/, '').replace(/\r/g, '\n')
  }
  return ''
}

function noteEntries(value: unknown): Array<[string, unknown]> {
  if (Array.isArray(value)) return value.map((entry, index) => [rawNoteId(entry, String(index + 1)), entry])
  return Object.entries(rawRecord(value))
}

function allBlocks(blocks: readonly DocumentBlock[]): DocumentBlock[] {
  const result: DocumentBlock[] = [], ids = new Set<string>()
  const visit = (content: readonly DocumentBlock[], depth: number) => {
    if (depth > 64) throw new RangeError('导入注释超过内容区嵌套预算。')
    for (const block of content) {
      if (ids.has(block.id)) throw new TypeError('导入注释的引用块身份重复：' + block.id)
      ids.add(block.id); result.push(block)
      if (result.length > 250_000) throw new RangeError('导入注释超过块预算。')
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const container of documentTableCellContentContainers(cell)) visit(container, depth + 1)
    }
  }
  visit(blocks, 0)
  return result
}

function decorationNote(value: unknown): boolean {
  const raw = rawRecord(value), metadata = rawRecord(raw.metadata)
  return ['separator', 'continuationSeparator', 'continuationNotice'].includes(String(raw.type ?? raw.noteType ?? ''))
    || metadata.decorationSource !== undefined || raw.decorationSource !== undefined
}

/** Normalize source identity once; importer assembles real note/decorations stories separately. */
export function normalizeDocumentNotes(data: Record<string, unknown>, blocks: DocumentBlock[]): Pick<DocumentModel, 'footnotes' | 'endnotes'> {
  const result: Pick<DocumentModel, 'footnotes' | 'endnotes'> = {}, blockIndex = allBlocks(blocks)
  const sourceRuns = blockIndex.flatMap((block, blockOrder) => block.runs.map((run, runIndex) => ({ block, blockOrder, run, runIndex,
    sourceMetadata: run.metadata && cloneModel(run.metadata) })))
  const claimed = new Map<string, string>()
  for (const kind of ['footnote', 'endnote'] as const) {
    const property = kind === 'footnote' ? 'footnotes' : 'endnotes', notes: Record<string, DocumentNote> = {}
    const entries = new Map<string, unknown>()
    for (const [sourceId, value] of [...noteEntries(data[property]), ...noteEntries(data.notes).filter(([, value]) => rawRecord(value).type === kind)]) {
      if (decorationNote(value)) continue
      const id = rawNoteId(value, sourceId)
      if (entries.has(id) && JSON.stringify(entries.get(id)) !== JSON.stringify(value)) throw new TypeError('导入的注释集合身份重复：' + id)
      entries.set(id, value)
    }
    const references: Array<{ note: DocumentNote; source: typeof sourceRuns[number] }> = []
    for (const [id, value] of entries) {
      const raw = rawRecord(value)
      const note: DocumentNote = { id, text: rawNoteText(value),
        referenceBlockId: typeof raw.referenceBlockId === 'string' ? raw.referenceBlockId : '',
        referenceRunId: typeof raw.referenceRunId === 'string' ? raw.referenceRunId : '',
        ...(raw.metadata && typeof raw.metadata === 'object' && !Array.isArray(raw.metadata) ? { metadata: cloneModel(raw.metadata as Record<string, unknown>) } : {}) }
      const matches = sourceRuns.filter(({ run, sourceMetadata }) => documentNoteReferenceRunMatches({ ...run, metadata: sourceMetadata }, note, kind))
      if (matches.length > 1) throw new TypeError('导入的注释引用字符不是唯一身份：' + id)
      const reference = matches[0]
      // Special source note IDs without a body reference remain importer decoration inputs.
      if (!reference && (id === '-1' || id === '0') && raw.type !== kind && raw.noteType !== kind) continue
      const marker = rawRecord(reference?.sourceMetadata?.noteReference ?? reference?.sourceMetadata?.note)
      const customMark = raw.customMark ?? marker.customMark
      if (typeof customMark === 'string') note.customMark = customMark
      const style = raw.referenceTextStyle ?? marker.referenceTextStyle ?? marker.textStyle
      if (style && typeof style === 'object' && !Array.isArray(style)) note.referenceTextStyle = cloneModel(style as NonNullable<DocumentNote['referenceTextStyle']>)
      if (typeof raw.showBodyReference === 'boolean') note.showBodyReference = raw.showBodyReference
      else if (typeof marker.showBodyReference === 'boolean') note.showBodyReference = marker.showBodyReference
      if (reference) {
        const owner = claimed.get(reference.run.id)
        if (owner && owner !== kind + ':' + id) throw new TypeError('导入的脚注和尾注不能共享一个未定类型的引用字符。')
        claimed.set(reference.run.id, kind + ':' + id)
        note.referenceBlockId = reference.block.id; note.referenceRunId = reference.run.id
        reference.run.metadata = { ...reference.run.metadata, note: { type: kind, noteId: id } }
        // Preserve a reviewable source trace when OOXML customMarkFollows emitted a separate literal.
        if (note.customMark !== undefined && marker.customMarkFollows === true && reference.run.text !== note.customMark) {
          const next = reference.block.runs[reference.runIndex + 1]
          if (next?.text.startsWith(note.customMark) && note.customMark.length) {
            const merged = { runId: next.id, startOffset: 0, endOffset: note.customMark.length, text: note.customMark }
            reference.run.metadata.officeNoteReferenceMergedLiteral = merged
            next.text = next.text.slice(note.customMark.length)
            // Keep the empty source run for parser source-coordinate projection; editor compaction removes it later.
          }
        }
        references.push({ note, source: reference })
      }
      const candidate = raw.content
      if (candidate && typeof candidate === 'object' && !Array.isArray(candidate) && Array.isArray(Reflect.get(candidate, 'blocks'))) {
        note.content = projectDocumentStoryBlocks(cloneModel(candidate as DocumentStoryContent), 'display').content
        note.text = documentNoteContentText(note.content)
      } else if (candidate && typeof candidate === 'object' && !Array.isArray(candidate) && Array.isArray(Reflect.get(candidate, 'runs')) && Array.isArray(Reflect.get(candidate, 'paragraphs'))) {
        note.content = cloneModel(candidate as DocumentStoryContent); note.text = documentNoteContentText(note.content)
      }
      notes[id] = note
    }
    let ordinal = 1
    references.sort((left, right) => left.source.blockOrder - right.source.blockOrder || left.source.runIndex - right.source.runIndex)
    for (const { note, source } of references) source.run.text = note.customMark ?? formatDocumentNoteNumber(ordinal++, kind === 'footnote' ? 'decimal' : 'lowerRoman').text
    if (Object.keys(notes).length) result[property] = notes
  }
  return result
}

/** A structural deletion removes orphaned note identities, then recomputes shared reference labels. */
export function reconcileDocumentNotesAfterStructure(model: DocumentModel, removedBlockIds: ReadonlySet<string>): void {
  let removed = true
  while (removed) {
    removed = false
    const references = documentNoteReferenceRecords(model)
    for (const kind of ['footnote', 'endnote'] as const) {
      const notes = kind === 'footnote' ? model.footnotes : model.endnotes
      if (!notes) continue
      for (const [id, note] of Object.entries(notes)) {
        const reference = references.find(reference => reference.kind === kind && reference.note.id === id)
        if (!reference && (note.referenceBlockId || removedBlockIds.has(note.referenceBlockId))) { delete notes[id]; removed = true }
        else if (reference) { note.referenceBlockId = reference.block.id; note.referenceRunId = reference.run.id }
      }
    }
  }
  renumberDocumentNotes(model)
}

/** Compatibility body query; richer navigation carries the actual story host separately. */
export function noteReferencePoint(model: DocumentModel, note: DocumentNote): DocumentPoint | undefined {
  for (const kind of ['footnote', 'endnote'] as const) {
    if ((kind === 'footnote' ? model.footnotes : model.endnotes)?.[note.id] !== note) continue
    const location = documentNoteReferenceLocation(model, kind, note.id)
    return location?.host ? undefined : location?.point
  }
  const record = documentNoteReferenceRecords(model).find(reference => reference.note.id === note.id && reference.run.id === note.referenceRunId)
  return record?.host || record?.linked ? undefined : record?.point
}
