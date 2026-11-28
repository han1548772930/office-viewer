import type { OfficeStudio } from './office-studio'
import { captureDocumentTocLayoutOwner, documentTocLayoutSnapshotCurrent } from './document-toc-layout'
import { documentTocIdForSelection, documentTocPageNumbers } from '../features/document/toc/document-toc'
import { documentTocActionState } from '../features/document/toc/document-toc-command'

export const DOCUMENT_TOC_DIALOG_OPERATION = 'doc.operation.open-table-of-contents-dialog'
export interface DocumentTocDialogEditor { openUpdate(): boolean }
export interface DocumentTocDialogEditorHandle { dispose(): void; isCurrent(): boolean }
interface Owner { editor: DocumentTocDialogEditor; generation: number }
const owners = new WeakMap<OfficeStudio, Map<string, Owner>>()

/** UI capability only: no document selection, modal or callback is persisted. */
export function bindDocumentTocDialogEditor(api: OfficeStudio, unitId: string, editor: DocumentTocDialogEditor): DocumentTocDialogEditorHandle {
  if (!api.getDocumentModel(unitId) || !editor || typeof editor.openUpdate !== 'function') throw new TypeError('目录对话框宿主无效。')
  let units = owners.get(api)
  if (!units) { units = new Map(); owners.set(api, units) }
  const owner: Owner = { editor, generation: api.editor.history.captureSavepoint(unitId).generation }
  units.set(unitId, owner)
  const isCurrent = () => !api.isDisposed && units?.get(unitId) === owner && !!api.getDocumentModel(unitId)
    && api.editor.history.captureSavepoint(unitId).generation === owner.generation
  return { isCurrent, dispose: () => { if (units?.get(unitId) === owner) { units.delete(unitId); if (!units.size) owners.delete(api) } } }
}

export function disposeDocumentTocDialogEditors(api: OfficeStudio, unitId?: string): void {
  const units = owners.get(api)
  if (!units) return
  if (unitId === undefined) { owners.delete(api); units.clear() }
  else { units.delete(unitId); if (!units.size) owners.delete(api) }
}

export function registerDocumentTocDialogOperation(api: OfficeStudio): void {
  api.registerCommand<unknown, boolean>(DOCUMENT_TOC_DIALOG_OPERATION, raw => {
    if (raw !== undefined && (!raw || typeof raw !== 'object' || Array.isArray(raw))) throw new TypeError('目录对话框参数无效。')
    const input = (raw ?? {}) as { unitId?: string }
    const unitId = input.unitId ?? api.activeDocumentId
    if (typeof unitId !== 'string' || !unitId.trim()) throw new TypeError('目录对话框的文档 ID 无效。')
    const owner = owners.get(api)?.get(unitId)
    const model = api.getDocumentModel(unitId), session = api.getDocumentInputState(unitId)
    if (!owner || !model || api.editor.history.captureSavepoint(unitId).generation !== owner.generation
      || !session?.valid || session.modelRevision !== api.getDocumentRevision(unitId)
      || session.location.kind !== 'body' || session.location.host) return false
    const tocId = documentTocIdForSelection(model, session.location.selection)
    const layoutOwner = captureDocumentTocLayoutOwner(api, unitId), snapshot = layoutOwner?.provider.read()
    if (!tocId || !layoutOwner || !snapshot || snapshot.source !== model || snapshot.modelRevision !== session.modelRevision
      || !documentTocLayoutSnapshotCurrent(api, unitId, layoutOwner, snapshot)) return false
    const pages = documentTocPageNumbers(snapshot.layout!.headings, snapshot.layout!.pages)
    const context = { documentId: unitId, userId: api.currentUserId, permissionAdminUserId: api.permissionAdminUserId }
    if (!documentTocActionState(model, session.location.selection, tocId, 'pageNumbersOnly', pages, context).available
      && !documentTocActionState(model, session.location.selection, tocId, 'all', pages, context).available) return false
    return owner.editor.openUpdate()
  })
}
