import type { DocumentModel } from '../../../domain/model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { cloneDocumentSelection } from '../model/document-table-selection'
import { collapsedSelection, orderedSelection, runText, runsForTarget, sameTarget, textTargets, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentModelForBlockHost } from '../services/document-block-host'
import { documentSupportsNotes } from '../notes/document-note-editing'
import { documentNoteReferenceLocation, documentNoteReferenceRecords } from '../notes/document-note-references'
import type { DocumentNoteKind } from '../notes/document-note-types'

export interface DocumentNoteNavigationPosition { selection: DocumentSelection; host?: DocumentAuxiliaryBlockHost }
interface DocumentNoteNavigationState { body?: DocumentSelection; footnote?: string; endnote?: string }

/** View memory is isolated by document identity and is never part of content history. */
export class DocumentNoteNavigationMemory {
  private readonly positions = new Map<string, DocumentNoteNavigationState>()
  observe(model: DocumentModel, position: DocumentNoteNavigationPosition): void {
    const key = model.id ?? '', previous = this.positions.get(key) ?? {}
    if (!position.host) previous.body = cloneDocumentSelection(position.selection)
    else if (position.host.kind === 'note') previous[position.host.noteKind] = position.host.noteId
    this.positions.set(key, previous)
  }
  get(model: DocumentModel): Readonly<DocumentNoteNavigationState> { return this.positions.get(model.id ?? '') ?? {} }
  forget(id: string): void { this.positions.delete(id) }
  clear(): void { this.positions.clear() }
}

function comparePoint(model: DocumentModel, left: DocumentPoint, right: DocumentPoint): number {
  if (sameTarget(left, right)) return left.offset - right.offset
  const targets = textTargets(model)
  const a = targets.findIndex(target => sameTarget(target, left)), b = targets.findIndex(target => sameTarget(target, right))
  if (a < 0 || b < 0) throw new RangeError('注释导航的正文位置已失效。')
  return a - b
}

function bodySelectionExists(model: DocumentModel, selection: DocumentSelection): boolean {
  const targets = textTargets(model)
  return [selection.anchor, selection.focus].every(point => {
    if (!targets.some(target => sameTarget(target, point))) return false
    const text = runText(runsForTarget(model, point))
    return Number.isSafeInteger(point.offset) && point.offset >= 0 && point.offset <= text.length && isDocumentTextBoundary(text, point.offset)
  })
}

/** Strict previous/next distance, no wrapping; a cross-kind note reuses that kind's last location. */
export function adjacentDocumentNote(model: DocumentModel, kind: DocumentNoteKind, direction: 'previous' | 'next',
  position: DocumentNoteNavigationPosition, memory: DocumentNoteNavigationMemory): string | undefined {
  if (!documentSupportsNotes(model) || position.selection.tableSelection || position.host && position.host.kind !== 'note') return undefined
  const references = documentNoteReferenceRecords(model).filter(reference => reference.kind === kind && !reference.host && reference.point)
  const remembered = memory.get(model)
  let selection = position.selection
  let noteId = position.host?.kind === 'note' ? position.host.noteId : undefined
  if (position.host?.kind === 'note' && position.host.noteKind !== kind) {
    const last = remembered[kind]
    noteId = last && references.some(reference => reference.note.id === last) ? last : undefined
    const first = textTargets(model)[0]
    if (!first) return undefined
    selection = remembered.body && bodySelectionExists(model, remembered.body) ? remembered.body : collapsedSelection({ ...first, offset: 0 })
  }
  if (!noteId && !bodySelectionExists(model, selection)) return undefined
  const origin = noteId ? references.find(reference => reference.note.id === noteId)?.point
    : direction === 'next' ? orderedSelection(model, selection).end : orderedSelection(model, selection).start
  if (!origin) return undefined
  const ordered = references.sort((a, b) => comparePoint(model, a.point!, b.point!))
  return (direction === 'next' ? ordered : [...ordered].reverse())
    .find(reference => (direction === 'next' ? 1 : -1) * comparePoint(model, reference.point!, origin) > 0)?.note.id
}

/** The upstream navigate operation defaults to the note body when an ID is supplied. */
export function documentNoteNavigationPosition(model: DocumentModel, kind: DocumentNoteKind, noteId: string,
  toReference = false): DocumentNoteNavigationPosition | undefined {
  if (!documentSupportsNotes(model)) return undefined
  if (toReference) {
    const location = documentNoteReferenceLocation(model, kind, noteId)
    return location ? { selection: collapsedSelection(location.point), ...(location.host ? { host: location.host } : {}) } : undefined
  }
  const note = (kind === 'footnote' ? model.footnotes : model.endnotes)?.[noteId]
  if (!note) return undefined
  const host: DocumentAuxiliaryBlockHost = { kind: 'note', noteKind: kind, noteId }
  const target = textTargets(documentModelForBlockHost(model, host))[0]
  return target ? { host, selection: collapsedSelection({ ...target, offset: 0 }) } : undefined
}
