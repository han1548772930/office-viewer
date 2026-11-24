import type { DocumentModel } from '../../../domain/model'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { sameDocumentInputLocation, type DocumentInputSession } from '../model/document-input-session'
import { collapsedSelection, runText, runsForTarget, sameTarget, tableForBlock, textTargets, type DocumentPoint } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { documentModelForBlockHost } from '../services/document-block-host'
import type { DocumentNoteLabel } from '../notes/document-note-numbering'
import type { DocumentNoteKind } from '../notes/document-note-types'
import { documentNoteReferenceLocation } from '../notes/document-note-references'
import { documentNotePermissionTargets } from '../permissions/document-note-permissions'
import { permissionTargetsForSelection, resolveDocumentPermission, type DocumentPermissionContext } from '../permissions/document-permissions'

export type DocumentNoteCanvasNavigationRequest = {
  source: DocumentModel; originHost?: DocumentAuxiliaryBlockHost; inputSession?: DocumentInputSession
} & ({ type: 'note'; kind: DocumentNoteKind; noteId: string; toReference: boolean; point?: DocumentPoint }
  | { type: 'body'; point: DocumentPoint })

export interface DocumentNoteCanvasContext {
  source: DocumentModel; host?: DocumentAuxiliaryBlockHost; session?: DocumentInputSession; requireSession?: boolean; disabled?: boolean;
  permission?: DocumentPermissionContext
}

function validPoint(model: DocumentModel, point: DocumentPoint): boolean {
  return textTargets(model).some(target => sameTarget(target, point)) && Number.isSafeInteger(point.offset)
    && point.offset >= 0 && point.offset <= runText(runsForTarget(model, point)).length
    && isDocumentTextBoundary(runText(runsForTarget(model, point)), point.offset)
}

/** Source, input revision and generation guard both glyph and preview actions. */
export function documentNoteCanvasRequestCurrent(request: DocumentNoteCanvasNavigationRequest, context: DocumentNoteCanvasContext): boolean {
  if (context.disabled || request.source !== context.source || !sameDocumentBlockHost(request.originHost, context.host)) return false
  const captured = request.inputSession, current = context.session
  if (context.requireSession && (!captured?.valid || !current?.valid)) return false
  if (captured && (!captured.valid || !current?.valid || captured.generation !== current.generation
    || captured.modelRevision !== current.modelRevision || !sameDocumentInputLocation(captured.location, current.location))) return false
  const permission = context.permission ?? { documentId: context.source.id ?? '', userId: 'local-user' }
  if (request.type === 'body') return validPoint(context.source, request.point)
    && resolveDocumentPermission(context.source.permissionRules ?? [], permission, 'view',
      permissionTargetsForSelection(context.source, collapsedSelection(request.point))).allowed
  const note = (request.kind === 'footnote' ? context.source.footnotes : context.source.endnotes)?.[request.noteId]
  if (!note) return false
  if (request.toReference) {
    const location = documentNoteReferenceLocation(context.source, request.kind, request.noteId)
    if (!location) return false
    const scoped = location.host ? documentModelForBlockHost(context.source, location.host) : context.source
    return resolveDocumentPermission(context.source.permissionRules ?? [], permission, 'view',
      permissionTargetsForSelection(scoped, collapsedSelection(location.point))).allowed
  }
  if (!resolveDocumentPermission(context.source.permissionRules ?? [], permission, 'view', documentNotePermissionTargets(context.source, request.kind, request.noteId)).allowed) return false
  return !request.point || validPoint(documentModelForBlockHost(context.source,
    { kind: 'note', noteKind: request.kind, noteId: request.noteId }), request.point)
}

export interface DocumentNotePreview { kind: DocumentNoteKind; noteId: string; label: string; text: string; truncated: boolean }

/** Collect at most 500 source characters and six lines without flattening a large note tree. */
export function documentNotePreview(model: DocumentModel, kind: DocumentNoteKind, noteId: string,
  labels?: ReadonlyMap<string, DocumentNoteLabel>, permission?: DocumentPermissionContext): DocumentNotePreview | undefined {
  const note = (kind === 'footnote' ? model.footnotes : model.endnotes)?.[noteId]
  if (!note) return undefined
  if (!resolveDocumentPermission(model.permissionRules ?? [], permission ?? { documentId: model.id ?? '', userId: 'local-user' }, 'view',
    documentNotePermissionTargets(model, kind, noteId)).allowed) return undefined
  let text = '', lines = 1, truncated = false, stopped = false
  const append = (value: string) => {
    if (stopped) { if (value) truncated = true; return }
    for (const character of value) {
      const normalized = character === '\r' ? '' : character
      if (!normalized) continue
      if (text.length + normalized.length > 500 || normalized === '\n' && lines >= 6) { truncated = true; stopped = true; return }
      text += normalized
      if (normalized === '\n') lines++
    }
  }
  let paragraphs = 0
  const visit = (blocks: DocumentModel['blocks'], depth: number) => {
    if (depth > 64) { truncated = true; stopped = true; return }
    for (const block of blocks) {
      if (stopped) { truncated = true; return }
      const table = tableForBlock(block)
      if (table) {
        for (const row of table.rows) for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) {
          visit(content, depth + 1)
          if (stopped) return
        }
      } else {
        if (paragraphs++) append('\n')
        if (block.type === 'image') append('[图片]')
        else if (block.type === 'divider') append('—')
        else for (const run of block.runs) { append(run.text); if (stopped) break }
      }
    }
  }
  if (note.content?.blocks) visit(note.content.blocks, 0)
  else if (note.content) for (const run of note.content.runs) { append(run.text); if (stopped) break }
  else append(note.text)
  return { kind, noteId, label: labels?.get(`${kind}:${noteId}`)?.label ?? note.customMark ?? '', text, truncated }
}
