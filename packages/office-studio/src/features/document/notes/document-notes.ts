import type { DocumentBlock, DocumentModel, DocumentNote } from '../../../domain/model'
import { collapsedSelection, createRun, documentId, marksAtPoint, orderedSelection, runText, runsForTarget, setTargetRuns, sliceRuns, tableForBlock, type DocumentPoint, type DocumentSelection } from '../model/document-model'

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
    if (typeof nested.dataStream === 'string') return nested.dataStream.replace(/[\r\n]+$/, '')
  }
  return ''
}

function noteEntries(value: unknown): Array<[string, unknown]> {
  if (Array.isArray(value)) return value.map((entry, index) => [rawNoteId(entry, String(index + 1)), entry])
  return Object.entries(rawRecord(value))
}

function allBlocks(blocks: DocumentBlock[]): DocumentBlock[] {
  return blocks.flatMap((block) => [block, ...((tableForBlock(block)?.rows.flatMap((row) => row.cells.flatMap((cell) => allBlocks(cell.blocks))) ?? []))])
}

export function normalizeDocumentNotes(data: Record<string, unknown>, blocks: DocumentBlock[]): Pick<DocumentModel, 'footnotes' | 'endnotes'> {
  const result: Pick<DocumentModel, 'footnotes' | 'endnotes'> = {}
  const blockIndex = allBlocks(blocks)
  for (const kind of ['footnote', 'endnote'] as const) {
    const property = kind === 'footnote' ? 'footnotes' : 'endnotes'
    const notes: Record<string, DocumentNote> = {}
    for (const [sourceId, value] of noteEntries(data[property])) {
      const id = rawNoteId(value, sourceId)
      const reference = blockIndex.flatMap((block) => block.runs.map((run) => ({ block, run }))).find(({ run }) => {
        const raw = rawRecord(run.metadata?.noteReference)
        const referenceId = raw.noteId ?? raw.id ?? raw.footnoteId ?? raw.endnoteId ?? run.metadata?.noteReference
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
  const notes = Object.values(noteCollection(model, kind))
    .map((note) => ({ note, point: noteReferencePoint(model, note) }))
    .filter((entry): entry is { note: DocumentNote; point: DocumentPoint } => entry.point !== undefined)
    .sort((a, b) => model.blocks.findIndex((block) => block.id === a.point.blockId) - model.blocks.findIndex((block) => block.id === b.point.blockId) || a.point.offset - b.point.offset)
  notes.forEach(({ note, point }, index) => {
    const runs = runsForTarget(model, point)
    const run = runs.find((candidate) => isReferenceRun(candidate, note, kind))
    if (run) run.text = String(index + 1)
  })
}

export function insertDocumentNote(model: DocumentModel, selection: DocumentSelection, kind: DocumentNoteKind): { selection: DocumentSelection; note: DocumentNote } {
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
  setTargetRuns(model, start, [...prefix, reference, ...suffix])
  const persistedReference = runsForTarget(model, start).find((run) => run.metadata?.note && typeof run.metadata.note === 'object' && 'noteId' in run.metadata.note && run.metadata.note.noteId === noteId)
  if (persistedReference) note.referenceRunId = persistedReference.id
  noteCollection(model, kind)[noteId] = note
  renumberNotes(model, kind)
  const point: DocumentPoint = { ...start, offset: before.length + 1 }
  return { selection: collapsedSelection(point), note }
}

export function updateDocumentNote(model: DocumentModel, kind: DocumentNoteKind, id: string, text: string): void {
  const note = noteCollection(model, kind)[id]
  if (!note) throw new RangeError(`Unknown ${kind} id: ${id}`)
  note.text = text
}

export function deleteDocumentNote(model: DocumentModel, kind: DocumentNoteKind, id: string): boolean {
  const notes = noteCollection(model, kind)
  const note = notes[id]
  if (!note) return false
  for (const block of model.blocks) {
    const table = tableForBlock(block)
    const paragraphs = block.id === note.referenceBlockId ? [block] : table
      ? table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks)).filter((candidate) => candidate.id === note.referenceBlockId)
      : []
    for (const paragraph of paragraphs) {
      const runIndex = paragraph.runs.findIndex((run) => isReferenceRun(run, note, kind))
      if (runIndex < 0) continue
      paragraph.runs.splice(runIndex, 1)
      if (!paragraph.runs.length) paragraph.runs.push({ id: documentId('run'), text: '', bold: false, italic: false, underline: false, strike: false, color: '#1f2937', fontFamily: 'Arial', fontSize: 11 })
    }
  }
  delete notes[id]
  renumberNotes(model, kind)
  return true
}

export function noteReferencePoint(model: DocumentModel, note: DocumentNote): DocumentPoint | undefined {
  for (const block of model.blocks) {
    if (block.id === note.referenceBlockId) {
      let offset = 0
      for (const run of block.runs) {
        if (isReferenceRun(run, note)) return { blockId: block.id, offset }
        offset += run.text.length
      }
    }
    const table = tableForBlock(block)
    if (!table) continue
    for (const [row, rowData] of table.rows.entries()) for (const [col, cell] of rowData.cells.entries()) {
      const paragraph = cell.blocks.find((candidate) => candidate.id === note.referenceBlockId)
      if (!paragraph) continue
      let offset = 0
      for (const run of paragraph.runs) {
        if (isReferenceRun(run, note)) return { blockId: block.id, row, col, cellBlockId: paragraph.id, offset }
        offset += run.text.length
      }
    }
  }
  return undefined
}
