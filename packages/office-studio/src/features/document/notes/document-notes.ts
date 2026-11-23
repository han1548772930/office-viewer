import type { DocumentBlock, DocumentModel, DocumentNote, DocumentStoryContent } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { projectDocumentStoryBlocks } from '../model/document-story-blocks'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentStoryParagraphs } from '../model/document-story'
import { documentModelForBlockHost, applyDocumentBlockHostModel } from '../services/document-block-host'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { replaceDocumentRange } from '../services/document-facade'
import { blockAtPoint, collapsedSelection, createRun, documentId, marksAtPoint, orderedSelection, runText, runsForTarget, setTargetRuns, sliceRuns, tableForBlock, textTargets, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import { transformDocumentTextFieldsForBlockEdit } from '../model/document-text-fields'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments'
import { getDocumentContentControls, transformContentControlsForBlockEdit } from '../services/content-controls'
import { documentTableCellBlocks } from '../model/document-table-continuations'

export type DocumentNoteKind = 'footnote' | 'endnote'

function rawRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
}

function rawNoteId(value: unknown, fallback: string): string {
  const raw = rawRecord(value)
  const id = raw.id ?? raw.noteId
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

function allBlocks(blocks: DocumentBlock[]): DocumentBlock[] {
  return blocks.flatMap((block) => [block, ...((tableForBlock(block)?.rows.flatMap((row) => row.cells.flatMap((cell) => allBlocks(documentTableCellBlocks(cell)))) ?? []))])
}

export function normalizeDocumentNotes(data: Record<string, unknown>, blocks: DocumentBlock[]): Pick<DocumentModel, 'footnotes' | 'endnotes'> {
  const result: Pick<DocumentModel, 'footnotes' | 'endnotes'> = {}
  const blockIndex = allBlocks(blocks)
  for (const kind of ['footnote', 'endnote'] as const) {
    const property = kind === 'footnote' ? 'footnotes' : 'endnotes'
    const notes: Record<string, DocumentNote> = {}
    const entries = new Map(noteEntries(data[property]).map(([sourceId, value]) => [rawNoteId(value, sourceId), value]))
    // Imported snapshots may store both note kinds in a shared map.
    for (const [sourceId, value] of noteEntries(data.notes)) {
      const id = rawNoteId(value, sourceId)
      if (rawRecord(value).type === kind && !entries.has(id)) entries.set(id, value)
    }
    for (const [sourceId, value] of entries) {
      const id = rawNoteId(value, sourceId)
      const reference = blockIndex.flatMap((block) => block.runs.map((run) => ({ block, run }))).find(({ run }) => {
        const sourceReference = run.metadata?.noteReference ?? run.metadata?.note
        const raw = rawRecord(sourceReference)
        const referenceId = raw.noteId ?? raw.id ?? raw.footnoteId ?? raw.endnoteId ?? sourceReference
        const referenceKind = raw.type ?? raw.kind ?? raw.noteType
        return String(referenceId ?? '') === id && (referenceKind === undefined || referenceKind === kind)
      })
      if (reference) reference.run.metadata = { ...reference.run.metadata, note: { type: kind, noteId: id } }
      notes[id] = {
        id,
        text: rawNoteText(value),
        referenceBlockId: reference?.block.id ?? '',
        referenceRunId: reference?.run.id ?? '',
      }
      const candidate = rawRecord(value).content
      if (candidate && typeof candidate === 'object' && !Array.isArray(candidate) && Array.isArray(Reflect.get(candidate, 'blocks'))) {
        const content = projectDocumentStoryBlocks(cloneModel(candidate as DocumentStoryContent), 'display').content
        notes[id]!.content = content
        notes[id]!.text = runText(content.runs)
      } else if (candidate && typeof candidate === 'object' && !Array.isArray(candidate) && Array.isArray(Reflect.get(candidate, 'runs')) && Array.isArray(Reflect.get(candidate, 'paragraphs'))) {
        notes[id]!.content = cloneModel(candidate as DocumentStoryContent)
        notes[id]!.text = runText(notes[id]!.content!.runs)
      }
    }
    if (Object.keys(notes).length) result[property] = notes
  }
  return result
}

function isReferenceRun(run: { id: string; metadata?: Record<string, unknown> }, note: DocumentNote, kind?: DocumentNoteKind): boolean {
  if (note.referenceRunId && run.id === note.referenceRunId) return true
  const reference = run.metadata?.note
  return !!reference && typeof reference === 'object' && 'noteId' in reference && reference.noteId === note.id
    && (kind === undefined || !('type' in reference) || reference.type === kind)
}

function noteCollection(model: DocumentModel, kind: DocumentNoteKind): Record<string, DocumentNote> {
  if (kind === 'footnote') return model.footnotes ??= {}
  return model.endnotes ??= {}
}

function renumberNotes(model: DocumentModel, kind: DocumentNoteKind): void {
  const notes = noteCollection(model, kind)
  const notesByRunId = new Map(Object.values(notes).filter(note => note.referenceRunId).map(note => [note.referenceRunId, note]))
  let index = 0
  for (const block of allBlocks(model.blocks)) {
    let offset = 0
    for (const run of block.runs) {
      const reference = rawRecord(run.metadata?.note)
      const note = notesByRunId.get(run.id) ?? notes[String(reference.noteId)]
      if (note?.referenceBlockId === block.id && isReferenceRun(run, note, kind)) {
        const result = String(++index)
        if (result !== run.text) {
          const end = offset + run.text.length
          transformDocumentTextFieldsForBlockEdit(block, offset, end, result, undefined, model)
          remapNoteContentControlOffsets(model, block.id, offset, end, result.length)
          run.text = result
          transformDocumentCommentAnchorsForTextEdit(model, block.id, offset, end, result)
        }
      }
      offset += run.text.length
    }
  }
}

/** Table/paragraph structure retains note IDs and removes notes whose references were deleted. */
export function reconcileDocumentNotesAfterStructure(model: DocumentModel, removedBlockIds: ReadonlySet<string>): void {
  const blocks = new Map(allBlocks(model.blocks).map(block => [block.id, block]))
  for (const property of ['footnotes', 'endnotes'] as const) {
    const notes = model[property]
    if (!notes) continue
    for (const [id, note] of Object.entries(notes)) {
      const paragraph = blocks.get(note.referenceBlockId)
      if (removedBlockIds.has(note.referenceBlockId) || note.referenceBlockId && (!paragraph || !paragraph.runs.some(run => isReferenceRun(run, note)))) delete notes[id]
    }
  }
  if (model.footnotes) renumberNotes(model, 'footnote')
  if (model.endnotes) renumberNotes(model, 'endnote')
}

export function insertDocumentNote(model: DocumentModel, selection: DocumentSelection, kind: DocumentNoteKind): { selection: DocumentSelection; note: DocumentNote } {
  if (selection.tableSelection) throw new RangeError('脚注或尾注需要单元格内的文字选区或插入位置')
  const { start, end } = orderedSelection(model, selection)
  if (start.blockId !== end.blockId || start.row !== end.row || start.col !== end.col || start.cellBlockId !== end.cellBlockId) {
    throw new RangeError('A note reference must be inserted within one paragraph or table cell.')
  }
  const runs = runsForTarget(model, start)
  const text = runText(runs)
  const noteId = documentId(kind)
  const referenceRunId = documentId('note-reference')
  const referenceBlockId = start.cellBlockId ?? start.blockId
  const note: DocumentNote = { id: noteId, text: '', referenceBlockId, referenceRunId }
  const before = text.slice(0, start.offset)
  const marks = marksAtPoint(model, start)
  const reference = createRun('1', { ...marks, verticalAlign: 'superscript', metadata: { ...marks.metadata, note: { type: kind, noteId } } })
  const prefix = sliceRuns(runs, 0, start.offset)
  const suffix = sliceRuns(runs, end.offset)
  const controlPlan = { ...model, contentControls: model.contentControls && getDocumentContentControls(model) }
  transformContentControlsForBlockEdit(controlPlan, referenceBlockId, start.offset, end.offset, reference.text.length)
  transformDocumentTextFieldsForBlockEdit(blockAtPoint(model, start)!, start.offset, end.offset, reference.text, undefined, model)
  model.contentControls = controlPlan.contentControls
  setTargetRuns(model, start, [...prefix, reference, ...suffix])
  transformDocumentCommentAnchorsForTextEdit(model, referenceBlockId, start.offset, end.offset, reference.text)
  const persistedReference = runsForTarget(model, start).find((run) => run.metadata?.note && typeof run.metadata.note === 'object' && 'noteId' in run.metadata.note && run.metadata.note.noteId === noteId)
  if (persistedReference) note.referenceRunId = persistedReference.id
  noteCollection(model, kind)[noteId] = note
  renumberNotes(model, kind)
  const point: DocumentPoint = { ...start, offset: before.length + (persistedReference?.text.length ?? reference.text.length) }
  return { selection: collapsedSelection(point), note }
}

export function updateDocumentNote(model: DocumentModel, kind: DocumentNoteKind, id: string, text: string): void {
  const note = noteCollection(model, kind)[id]
  if (!note) throw new RangeError(`Unknown ${kind} id: ${id}`)
  if (note.content?.blocks !== undefined) {
    if (text === note.text) return
    const host = { kind: 'note' as const, noteKind: kind, noteId: id }, scoped = documentModelForBlockHost(cloneModel(model), host, true)
    const paragraphs = documentTextFieldParagraphs(scoped)
    // The legacy text setter remains useful for ordinary notes; real tables require block-target editing.
    if (scoped.blocks.some(block => block.type !== 'paragraph') || paragraphs.some(paragraph => paragraph.container !== scoped.blocks)) throw new RangeError('包含表格或对象的注释需要在画布中编辑对应内容。')
    const updated = replaceDocumentRange(scoped, 0, paragraphs.at(-1)?.absoluteEnd ?? 0, text)
    applyDocumentBlockHostModel(model, host, updated)
    return
  }
  note.text = text
  if (note.content) {
    const defaults = documentDefaultStyle(model)
    note.content = { ...note.content, runs: [createRun(text, defaults.text)], paragraphs: documentStoryParagraphs(text, defaults.paragraph) }
  }
}

export function deleteDocumentNote(model: DocumentModel, kind: DocumentNoteKind, id: string): boolean {
  const notes = noteCollection(model, kind)
  const note = notes[id]
  if (!note) return false
  const paragraph = allBlocks(model.blocks).find((block) => block.id === note.referenceBlockId)
  const runIndex = paragraph?.runs.findIndex((run) => isReferenceRun(run, note, kind)) ?? -1
  if (paragraph && runIndex >= 0) {
    const offset = paragraph.runs.slice(0, runIndex).reduce((total, run) => total + run.text.length, 0), length = paragraph.runs[runIndex]!.text.length
    const controlPlan = { ...model, contentControls: model.contentControls && getDocumentContentControls(model) }
    transformContentControlsForBlockEdit(controlPlan, paragraph.id, offset, offset + length, 0)
    transformDocumentTextFieldsForBlockEdit(paragraph, offset, offset + length, '', undefined, model)
    model.contentControls = controlPlan.contentControls
    paragraph.runs.splice(runIndex, 1)
    if (!paragraph.runs.length) paragraph.runs.push({ id: documentId('run'), text: '', bold: false, italic: false, underline: false, strike: false, color: '#1f2937', fontFamily: 'Arial', fontSize: 11 })
    transformDocumentCommentAnchorsForTextEdit(model, paragraph.id, offset, offset + length, '')
  }
  delete notes[id]
  renumberNotes(model, kind)
  return true
}

/** Note labels are derived numbering; resizing them retains enclosing SDT wrappers. */
function remapNoteContentControlOffsets(model: DocumentModel, blockId: string, start: number, end: number, length: number): void {
  const delta = length - (end - start)
  for (const control of model.contentControls ?? []) {
    if (control.blockId !== blockId) continue
    if (end <= control.startOffset) { control.startOffset += delta; control.endOffset += delta }
    else if (start < control.endOffset && end > control.startOffset) {
      control.startOffset = Math.min(control.startOffset, start)
      control.endOffset = Math.max(control.startOffset, control.endOffset >= end ? control.endOffset + delta : start + length)
    }
  }
}

export function noteReferencePoint(model: DocumentModel, note: DocumentNote): DocumentPoint | undefined {
  const paragraph = allBlocks(model.blocks).find(block => block.id === note.referenceBlockId)
  const runIndex = paragraph?.runs.findIndex(run => isReferenceRun(run, note)) ?? -1
  if (!paragraph || runIndex < 0) return undefined

  // TextTarget resolves a body paragraph or a direct table-cell paragraph. A
  // deeper table path cannot be represented by its single row/col pair.
  const target = textTargets(model).find(candidate => (candidate.cellBlockId ?? candidate.blockId) === paragraph.id)
  if (!target) return undefined
  const offset = paragraph.runs.slice(0, runIndex).reduce((total, run) => total + run.text.length, 0)
  return { ...target, offset }
}
