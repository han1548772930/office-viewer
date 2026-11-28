import type { OfficeStudio } from './office-studio'
import { assertDocumentEditorCommandRequest, type DocumentEditorCommandRequest, type DocumentEditorCommandResult } from './document-editor-commands'
import { collapsedSelection } from '../features/document/model/document-model'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection'
import { documentHeadingCollapseState, documentHeadingVisibility, SET_DOCUMENT_HEADING_COLLAPSED_COMMAND } from '../features/document/headings/heading-collapse'

export { SET_DOCUMENT_HEADING_COLLAPSED_COMMAND } from '../features/document/headings/heading-collapse'
export interface DocumentHeadingCollapsedCommandParams { unitId: string; paragraphId: string; collapsed: boolean }

export function registerDocumentHeadingCollapseCommand(api: OfficeStudio,
  commit: (request: DocumentEditorCommandRequest) => DocumentEditorCommandResult): void {
  api.registerCommand<unknown, DocumentEditorCommandResult | boolean>(SET_DOCUMENT_HEADING_COLLAPSED_COMMAND, raw => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return false
    if ('command' in raw) {
      const request = raw as DocumentEditorCommandRequest
      assertDocumentEditorCommandRequest(request, SET_DOCUMENT_HEADING_COLLAPSED_COMMAND)
      return commit(request)
    }
    const input = raw as DocumentHeadingCollapsedCommandParams
    if (typeof input.unitId !== 'string' || !input.unitId.trim() || typeof input.paragraphId !== 'string'
      || !input.paragraphId.trim() || typeof input.collapsed !== 'boolean') return false
    const model = api.getDocumentModel(input.unitId)
    if (!model) return false
    const visibility = documentHeadingVisibility(model), section = visibility.sectionById.get(input.paragraphId)
    if (!section || section.collapsed === input.collapsed || !documentHeadingCollapseState(model, input.paragraphId, input.collapsed,
      { documentId: input.unitId, userId: api.currentUserId, permissionAdminUserId: api.permissionAdminUserId }, visibility).available) return false
    const revision = api.getDocumentRevision(input.unitId), state = api.getDocumentInputState(input.unitId)
    const selection = state?.valid && state.modelRevision === revision && state.location.kind === 'body' && !state.location.host
      ? cloneDocumentSelection(state.location.selection) : collapsedSelection({ ...section.target, offset: 0 })
    return commit({ unitId: input.unitId, selection, selectionHost: null,
      command: { id: SET_DOCUMENT_HEADING_COLLAPSED_COMMAND, paragraphId: input.paragraphId, collapsed: input.collapsed },
      expectedRevision: revision, expectedHistoryGeneration: api.editor.history.captureSavepoint(input.unitId).generation,
      inputGeneration: state?.generation }).changed
  })
}
