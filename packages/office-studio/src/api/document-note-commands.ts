import type { DocumentModel } from '../domain/model'
import type { OfficeStudio } from './office-studio'
import type { DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { documentStoredBlockHosts, sameDocumentBlockHost } from '../features/document/model/document-block-host'
import { documentNoteDecorationContent } from '../features/document/model/document-note-decoration-host'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection'
import { runsForTarget, type DocumentPoint, type DocumentSelection } from '../features/document/model/document-model'
import { documentModelForBlockHost } from '../features/document/services/document-block-host'
import { sanitizeDocumentInputSelection } from '../features/document/model/document-input-session'
import type { DocumentNoteDecorationControls, DocumentNoteInsertOptions, DocumentNoteKind, DocumentNoteMutationRequest, DocumentNotePatch } from '../features/document/notes/document-note-types'
import type { DocumentNoteSettingsValue } from '../features/document/notes/document-note-settings'

export const DOCUMENT_NOTE_COMMANDS = {
  insert: { footnote: 'doc.command.insert-footnote', endnote: 'doc.command.insert-endnote' },
  delete: { footnote: 'doc.command.delete-footnote', endnote: 'doc.command.delete-endnote' },
  update: { footnote: 'doc.command.update-footnote', endnote: 'doc.command.update-endnote' },
  'update-text': { footnote: 'doc.command.update-footnotes', endnote: 'doc.command.update-endnotes' },
  settings: { footnote: 'doc.command.set-footnote-settings', endnote: 'doc.command.set-endnote-settings' },
} as const

export interface DocumentNoteCommandContext {
  unitId: string
  expectedRevision?: number
  expectedHistoryGeneration?: number
  inputGeneration?: number
  /** Source view for stale-input checks; insertion offsets always address the main body. */
  host?: DocumentAuxiliaryBlockHost
  feedback?: boolean
}

export type DocumentNoteCommandParams = DocumentNoteCommandContext & (
  | { selection: DocumentSelection; options?: DocumentNoteInsertOptions }
  | { noteId: string; patch?: DocumentNotePatch }
  | { values: Record<string, string> }
  | { settings: DocumentNoteSettingsValue | null; sectionId?: string; suppressEndnotes?: boolean; decorationContentControls?: DocumentNoteDecorationControls }
)

export interface DocumentNoteCommandResult {
  changed: boolean
  selectionCurrent: boolean
  selection?: DocumentSelection
  host?: DocumentAuxiliaryBlockHost
  noteId?: string
  modelRevision?: number
  inputGeneration?: number
}

/** Expected request/domain rejection is distinct from unexpected observer or platform failures. */
export class DocumentNoteCommandRejectedError extends Error {
  readonly code = 'DOCUMENT_NOTE_COMMAND_REJECTED'
  constructor(message: string, cause?: unknown) { super(message, { cause }); this.name = 'DocumentNoteCommandRejectedError' }
}

export function documentNoteCommandId(request: DocumentNoteMutationRequest): string {
  if (!Object.hasOwn(DOCUMENT_NOTE_COMMANDS, request.type) || !['footnote', 'endnote'].includes(request.kind)) {
    throw new DocumentNoteCommandRejectedError('注释命令参数无效。')
  }
  return DOCUMENT_NOTE_COMMANDS[request.type][request.kind]
}

function requestForCommand(type: DocumentNoteMutationRequest['type'], kind: DocumentNoteKind, input: DocumentNoteCommandParams): DocumentNoteMutationRequest {
  if (type === 'insert' && 'selection' in input) return { type, kind, selection: input.selection, ...(input.options === undefined ? {} : { options: input.options }) }
  if (type === 'delete' && 'noteId' in input) return { type, kind, noteId: input.noteId }
  if (type === 'update' && 'noteId' in input && input.patch !== undefined) return { type, kind, noteId: input.noteId, patch: input.patch }
  if (type === 'update-text' && 'values' in input) return { type, kind, values: input.values }
  if (type === 'settings' && 'settings' in input) return { type, kind, settings: input.settings,
    ...(input.sectionId === undefined ? {} : { sectionId: input.sectionId }), ...(input.suppressEndnotes === undefined ? {} : { suppressEndnotes: input.suppressEndnotes }),
    ...(input.decorationContentControls === undefined ? {} : { decorationContentControls: input.decorationContentControls }) }
  throw new DocumentNoteCommandRejectedError('注释命令缺少对应参数。')
}

/** Public commands return booleans; editor controllers request guarded selection feedback. */
export function registerDocumentNoteCommands(api: OfficeStudio): void {
  for (const type of Object.keys(DOCUMENT_NOTE_COMMANDS) as DocumentNoteMutationRequest['type'][]) {
    for (const kind of ['footnote', 'endnote'] as const) {
      api.registerCommand<DocumentNoteCommandParams, boolean | DocumentNoteCommandResult>(DOCUMENT_NOTE_COMMANDS[type][kind], input => {
        if (!input || typeof input !== 'object' || Array.isArray(input) || typeof input.unitId !== 'string' || !input.unitId.trim()) {
          throw new DocumentNoteCommandRejectedError('注释命令所属文档无效。')
        }
        const result = api.commitDocumentNoteMutation(input, requestForCommand(type, kind, input))
        return input.feedback ? result : result.changed
      })
    }
  }
}

/** Numbering changes preserve a caret relative to surviving runs, including empty reference entities. */
export function mapDocumentNoteSelection(before: DocumentModel, after: DocumentModel, selection: DocumentSelection,
  host?: DocumentAuxiliaryBlockHost): DocumentSelection {
  return mapDocumentNoteLocation(before, after, selection, host).selection
}

/** Removing a special story also removes its editing address; its caret returns to the body. */
export function mapDocumentNoteLocation(before: DocumentModel, after: DocumentModel, selection: DocumentSelection,
  host?: DocumentAuxiliaryBlockHost): { selection: DocumentSelection; host?: DocumentAuxiliaryBlockHost } {
  if (host) {
    const oldStored = documentStoredBlockHosts(before).some(entry => sameDocumentBlockHost(entry.host, host))
    const nextStored = documentStoredBlockHosts(after).some(entry => sameDocumentBlockHost(entry.host, host))
    let survives = !oldStored || nextStored
    if (host.kind === 'note') survives = !!(host.noteKind === 'footnote' ? after.footnotes : after.endnotes)?.[host.noteId]
    if (host.kind === 'noteDecoration' && host.dormantIndex !== undefined && survives) {
      const source = documentNoteDecorationContent(before, host)?.metadata?.decorationSource
      const nextSource = documentNoteDecorationContent(after, host)?.metadata?.decorationSource
      // A numeric array index cannot identify a different native story after replacement.
      survives = !!source && !!nextSource && ['version', 'kind', 'type', 'noteId', 'partName'].every(key =>
        typeof source === 'object' && typeof nextSource === 'object' && Reflect.get(source, key) === Reflect.get(nextSource, key))
    }
    if (!survives) return { selection: sanitizeDocumentInputSelection(after, { anchor: { ...selection.anchor }, focus: { ...selection.focus } }) }
  }
  const previous = host ? documentModelForBlockHost(before, host) : before
  const next = host ? documentModelForBlockHost(after, host) : after
  const point = (source: DocumentPoint): DocumentPoint => {
    const oldRuns = runsForTarget(previous, source), newRuns = runsForTarget(next, source)
    let oldOffset = 0
    for (let index = 0; index < oldRuns.length; index++) {
      const run = oldRuns[index]!, end = oldOffset + run.text.length
      if (source.offset < end || source.offset === end && index === oldRuns.length - 1) {
        const newIndex = newRuns.findIndex(candidate => candidate.id === run.id)
        if (newIndex >= 0) {
          const prefix = newRuns.slice(0, newIndex).reduce((total, candidate) => total + candidate.text.length, 0)
          return { ...source, offset: prefix + Math.min(source.offset - oldOffset, newRuns[newIndex]!.text.length) }
        }
        break
      }
      oldOffset = end
    }
    return { ...source }
  }
  const mapped = cloneDocumentSelection(selection)
  mapped.anchor = point(mapped.anchor); mapped.focus = point(mapped.focus)
  return { selection: sanitizeDocumentInputSelection(next, mapped), ...(host ? { host: { ...host } } : {}) }
}
