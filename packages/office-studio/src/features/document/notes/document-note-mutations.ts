import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { collapsedSelection, runText, runsForTarget, type DocumentSelection } from '../model/document-model'
import { documentModelForBlockHost } from '../services/document-block-host'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentNoteMutationPermissionTargets, documentNotePermissionTargets } from '../permissions/document-note-permissions'
import { assertDocumentNoteSerializable } from './document-note-content'
import { deleteDocumentNote, insertDocumentNote, patchDocumentNote } from './document-note-editing'
import { documentNoteReferenceRecords } from './document-note-references'
import { renumberDocumentNotes } from './document-note-reference-sync'
import { setDocumentNoteSettings } from './document-note-settings'
import type { DocumentNoteMutationRequest } from './document-note-types'

export type { DocumentNoteMutationRequest, DocumentNoteInsertOptions, DocumentNoteKind, DocumentNotePatch, DocumentNoteDecorationControls } from './document-note-types'

export interface DocumentNoteMutationPlan {
  model: DocumentModel
  changed: boolean
  selection?: DocumentSelection
  host?: DocumentAuxiliaryBlockHost
  noteId?: string
}

/** UI, Facade and command entry points prepare one detached, fully checked candidate. */
export function prepareDocumentNoteMutation(before: DocumentModel, request: DocumentNoteMutationRequest,
  permission?: DocumentPermissionContext): DocumentNoteMutationPlan {
  assertDocumentNoteSerializable(request)
  if (!request || typeof request !== 'object' || Array.isArray(request) || !['footnote', 'endnote'].includes(request.kind)) throw new TypeError('注释命令参数无效。')
  if (!['insert', 'delete', 'update', 'update-text', 'settings'].includes(request.type)) throw new TypeError('未知的注释命令。')
  if (permission) {
    assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', request.type === 'insert'
      ? permissionTargetsForSelection(before, request.selection) : request.type === 'delete' || request.type === 'update'
        ? documentNotePermissionTargets(before, request.kind, request.noteId) : [])
  }
  const candidate = cloneModel(before), plan: DocumentNoteMutationPlan = { model: candidate, changed: false }
  if (request.type === 'insert') {
    const result = insertDocumentNote(candidate, request.selection, request.kind, request.options)
    Object.assign(plan, { selection: result.selection, host: result.host, noteId: result.note.id })
  } else if (request.type === 'delete') {
    const reference = documentNoteReferenceRecords(before).find(reference => reference.kind === request.kind && reference.note.id === request.noteId)
    const nextRunId = reference?.block.runs[reference.runIndex + 1]?.id
    deleteDocumentNote(candidate, request.kind, request.noteId)
    if (reference?.point && !reference.linked) {
      const scoped = reference.host ? documentModelForBlockHost(candidate, reference.host) : candidate
      const runs = runsForTarget(scoped, reference.point), index = nextRunId ? runs.findIndex(run => run.id === nextRunId) : -1
      const offset = index >= 0 ? runText(runs.slice(0, index)).length : runText(runs).length
      plan.selection = collapsedSelection({ ...reference.point, offset })
      if (reference.host) plan.host = { ...reference.host }
    }
    plan.noteId = request.noteId
  } else if (request.type === 'update') {
    patchDocumentNote(candidate, request.kind, request.noteId, request.patch)
    plan.noteId = request.noteId
  } else if (request.type === 'update-text') {
    if (!request.values || typeof request.values !== 'object' || Array.isArray(request.values)) throw new TypeError('注释文字修改需要按标识提供。')
    for (const [id, text] of Object.entries(request.values)) {
      if (permission) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', documentNotePermissionTargets(before, request.kind, id))
      patchDocumentNote(candidate, request.kind, id, { text })
    }
  } else if (setDocumentNoteSettings(candidate, request.kind, request.settings, request.sectionId, request.suppressEndnotes, request.decorationContentControls)) {
    renumberDocumentNotes(candidate)
  }
  plan.changed = JSON.stringify(before) !== JSON.stringify(candidate)
  if (permission && plan.changed) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', documentNoteMutationPermissionTargets(before, candidate))
  return plan
}
