import type { DocumentModel, DocumentNote } from '../../../domain/model'
import { documentModelForBlockHost } from '../services/document-block-host'
import { documentNoteContentText } from '../notes/document-note-content'
import { resolveDocumentNoteLabels } from '../notes/document-note-numbering'
import type { DocumentNoteKind } from '../notes/document-note-types'
import { assertDocumentRepeatingNoteContentRemovable } from '../services/repeating-content-control-copy'
import { documentNoteEditState } from '../commands/document-note-command-state'
import type { DocumentPermissionContext } from '../permissions/document-permissions'
import { resolveDocumentPermission } from '../permissions/document-permissions'
import { documentNotePermissionTargets } from '../permissions/document-note-permissions'

export interface DocumentNotePanelEntry {
  key: string; kind: DocumentNoteKind; note: DocumentNote; label: string; text: string
  referenceAvailable: boolean; viewable: boolean; textEditable: boolean; editable: boolean; removable: boolean; reason?: string
}

/** Display order and labels use the same reference inventory as layout, never collection indices. */
export function documentNotePanelEntries(model: DocumentModel, labels: Readonly<Record<string, string>> = {},
  permission?: DocumentPermissionContext, disabled = false): DocumentNotePanelEntry[] {
  const numbered = resolveDocumentNoteLabels(model), result: DocumentNotePanelEntry[] = [], seen = new Set<string>()
  const add = (kind: DocumentNoteKind, note: DocumentNote) => {
    const key = `${kind}:${note.id}`, reference = numbered.get(key), scoped = documentModelForBlockHost(model, { kind: 'note', noteKind: kind, noteId: note.id })
    const textEditable = scoped.blocks.every(block => block.type === 'paragraph')
    const context = permission ?? { documentId: model.id ?? '', userId: 'local-user' }
    const viewable = resolveDocumentPermission(model.permissionRules ?? [], context, 'view', documentNotePermissionTargets(model, kind, note.id)).allowed
    const state = viewable ? documentNoteEditState(model, kind, note.id, { permission: context, disabled })
      : { available: false, reason: '此注释没有查看权限。' }
    let removable = state.available, reason = state.reason
    if (removable) {
      try { assertDocumentRepeatingNoteContentRemovable(model, note) }
      catch (error) { if (!(error instanceof Error)) throw error; removable = false; reason = error.message }
    }
    result.push({ key, kind, note, text: viewable ? note.content ? documentNoteContentText(note.content) : note.text : '',
      label: viewable ? labels[key] ?? reference?.label ?? '' : '', referenceAvailable: viewable && !!reference && !reference.reference.linked, viewable,
      textEditable, editable: state.available, removable, ...(reason ? { reason } : {}) })
    seen.add(key)
  }
  for (const [key, entry] of numbered) if (!seen.has(key)) add(entry.reference.kind, entry.reference.note)
  for (const kind of ['footnote', 'endnote'] as const) for (const note of Object.values(kind === 'footnote' ? model.footnotes ?? {} : model.endnotes ?? {})) {
    if (!seen.has(`${kind}:${note.id}`)) add(kind, note)
  }
  return result
}
