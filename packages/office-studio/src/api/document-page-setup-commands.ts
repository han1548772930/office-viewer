import type { OfficeStudio } from './office-studio'
import type { DocumentEditorCommandRequest, DocumentEditorCommandResult } from './document-editor-commands'
import type { DocumentPageSetupRequest } from '../features/document/services/document-settings'
import { collapsedSelection, textTargets } from '../features/document/model/document-model'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection'

export const DOCUMENT_PAGE_SETUP_COMMAND = 'docs.command.page-setup'

export type DocumentPageSetupCommandParams = DocumentPageSetupRequest & Pick<Partial<DocumentEditorCommandRequest>,
  'unitId' | 'selection' | 'expectedRevision' | 'expectedHistoryGeneration' | 'inputGeneration'>

export function registerDocumentPageSetupCommand(api: OfficeStudio): void {
  api.registerCommand<unknown, DocumentEditorCommandResult | boolean>(DOCUMENT_PAGE_SETUP_COMMAND, raw => {
    if (raw === undefined || raw === null) return false
    if (typeof raw !== 'object' || Array.isArray(raw)) throw new TypeError('页面设置参数必须是对象。')
    if ('command' in raw) return api.commitDocumentEditorCommand(DOCUMENT_PAGE_SETUP_COMMAND, raw as DocumentEditorCommandRequest)
    const params = raw as DocumentPageSetupCommandParams
    const unitId = params.unitId ?? api.activeDocumentId
    if (typeof unitId !== 'string' || !unitId.trim()) throw new TypeError('页面设置的文档 ID 无效。')
    const model = api.getDocumentModel(unitId)
    if (!model) return false
    const state = api.getDocumentInputState(unitId)
    const selection = params.selection ?? (state?.valid && state.location.kind === 'body' ? state.location.selection
      : collapsedSelection({ ...textTargets(model)[0]!, offset: 0 }))
    api.commitDocumentEditorCommand(DOCUMENT_PAGE_SETUP_COMMAND, {
      unitId, selection: cloneDocumentSelection(selection),
      command: { id: DOCUMENT_PAGE_SETUP_COMMAND, pageSize: params.pageSize, pageOrient: params.pageOrient, documentFlavor: params.documentFlavor,
        marginTop: params.marginTop, marginBottom: params.marginBottom, marginLeft: params.marginLeft, marginRight: params.marginRight },
      expectedRevision: params.expectedRevision ?? api.getDocumentRevision(unitId),
      expectedHistoryGeneration: params.expectedHistoryGeneration ?? api.editor.history.captureSavepoint(unitId).generation,
      inputGeneration: params.inputGeneration ?? state?.generation,
    })
    return true
  })
}
