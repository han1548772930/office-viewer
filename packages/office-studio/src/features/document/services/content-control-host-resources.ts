import type { DocumentModel } from '../../../domain/model'
import { reconcileDocumentNotesAfterStructure } from '../notes/document-notes'
import { copyDocumentRepeatingContentControlComments, copyDocumentRepeatingContentControlNoteControls,
  copyDocumentRepeatingContentControlNotes, removeDeletedDocumentRepeatingNoteScopes,
  type DocumentRepeatingContentControlCopyIds, type DocumentRepeatingContentControlInlineCopyRange } from './repeating-content-control-copy'

/** Publish a scoped block edit into a detached full model before resolving its cross-story resources. */
export function finalizeDocumentContentControlHostResources(before: DocumentModel, after: DocumentModel,
  ids?: DocumentRepeatingContentControlCopyIds, inline?: DocumentRepeatingContentControlInlineCopyRange): void {
  if (ids) {
    copyDocumentRepeatingContentControlNotes(before, after, ids)
    copyDocumentRepeatingContentControlNoteControls(before, after, ids)
    // The local planner saw only its host's comment segments. Rebuild each copied thread from
    // the full source once, including anchors in its independently copied note scopes.
    for (const id of ids.comment.values()) if (after.comments) delete after.comments[id]
    ids.comment.clear()
    copyDocumentRepeatingContentControlComments(before, after, ids, inline)
  }
  removeDeletedDocumentRepeatingNoteScopes(before, after)
  const keys = (model: DocumentModel) => (['footnotes', 'endnotes'] as const).flatMap(kind => Object.keys(model[kind] ?? {}).map(id => `${kind}:${id}`))
  if (JSON.stringify(keys(before)) !== JSON.stringify(keys(after))) reconcileDocumentNotesAfterStructure(after, new Set())
}
