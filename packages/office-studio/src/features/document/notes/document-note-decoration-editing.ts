import type { DocumentModel, DocumentStoryContent } from '../../../domain/model'
import { documentBlockHostId, documentStoredBlockHosts, type DocumentNoteDecorationHost } from '../model/document-block-host'
import { documentNoteDecorationContent } from '../model/document-note-decoration-host'
import { documentTableDescendantIds } from '../model/document-table-selection'
import { documentFieldLocked, documentTextFieldRanges } from '../model/document-text-fields'
import { documentModelForBlockHost } from '../services/document-block-host'

function decorationBlockIds(model: DocumentModel, host: DocumentNoteDecorationHost): Set<string> {
  return documentNoteDecorationContent(model, host) ? documentTableDescendantIds(documentModelForBlockHost(model, host).blocks) : new Set()
}

/** Anchored resources must use the live host, where commands also publish controls, comments and bindings. */
export function documentNoteDecorationRequiresLiveEditing(model: DocumentModel, host: DocumentNoteDecorationHost): boolean {
  const content = documentNoteDecorationContent(model, host)
  if (!content) return false
  const scoped = documentModelForBlockHost(model, host)
  return !!scoped.contentControls?.length || !!Object.keys(scoped.comments ?? {}).length || !!content.locks?.length
    || documentTextFieldRanges(scoped).length > 0
}

/** Whole-story replacement is a real deletion of its old anchors and obeys every content/field/control lock. */
export function assertDocumentNoteDecorationRemovable(model: DocumentModel, host: DocumentNoteDecorationHost): Set<string> {
  const content = documentNoteDecorationContent(model, host), ids = decorationBlockIds(model, host)
  if (!content) return ids
  const scoped = documentModelForBlockHost(model, host)
  if (content.locks?.some(lock => lock.lock !== 'unlocked') || scoped.contentControls?.some(control =>
    ['controlLocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'].includes(control.lock ?? ''))) {
    throw new RangeError('替换注释分隔内容不能移除锁定的内容或控件。')
  }
  if (documentTextFieldRanges(scoped).some(range => documentFieldLocked(range.field))) {
    throw new RangeError('替换注释分隔内容不能移除锁定字段。')
  }
  const foreign = new Set(documentTableDescendantIds(model.blocks)), own = documentBlockHostId(host)
  for (const entry of documentStoredBlockHosts(model)) if (documentBlockHostId(entry.host) !== own) {
    for (const id of documentTableDescendantIds(entry.blocks)) foreign.add(id)
  }
  if ([...ids].some(id => foreign.has(id))) throw new RangeError('注释分隔内容与其他内容区共享了块身份，不能整体替换。')
  return ids
}

/** Source addresses survive replacement; source XML remains provenance, never a second live story. */
export function preserveDocumentNoteDecorationSource(previous: DocumentStoryContent | undefined, next: DocumentStoryContent): DocumentStoryContent {
  if (!previous?.metadata) return next
  return { ...next, metadata: { ...previous.metadata, ...next.metadata,
    ...(previous.metadata.decorationSource === undefined ? {} : { decorationSource: previous.metadata.decorationSource }),
    ...(previous.metadata.univerNoteDecorationSource === undefined ? {} : { univerNoteDecorationSource: previous.metadata.univerNoteDecorationSource }) } }
}

/** Called only after all replacement plans passed validation, on a detached document candidate. */
export function removeDocumentNoteDecorationAnchors(model: DocumentModel, removed: ReadonlySet<string>): void {
  if (!removed.size) return
  if (model.contentControls) model.contentControls = model.contentControls.filter(control => !removed.has(control.blockId)
    && !control.blockRange?.blockIds.some(id => removed.has(id)) && !removed.has(control.tableRowRange?.tableId ?? ''))
  for (const [id, comment] of Object.entries(model.comments ?? {})) {
    const anchor = comment.anchor.filter(segment => !removed.has(segment.blockId))
    if (!anchor.length) delete model.comments![id]
    else if (anchor.length !== comment.anchor.length) model.comments![id] = { ...comment, anchor, referenceText: anchor.map(segment => segment.text).join('\n') }
  }
}
