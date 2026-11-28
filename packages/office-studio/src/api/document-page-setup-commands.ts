import type { OfficeStudio } from './office-studio'
import { assertDocumentEditorCommandRequest, type DocumentEditorCommandRequest, type DocumentEditorCommandResult } from './document-editor-commands'
import { assertDocumentPageSetupRequest, type DocumentPageSetupRequest } from '../features/document/services/document-settings'
import { collapsedSelection, textTargets } from '../features/document/model/document-model'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { documentModelForBlockHost, validateDocumentBlockHost } from '../features/document/services/document-block-host'

export const DOCUMENT_PAGE_SETUP_COMMAND = 'docs.command.page-setup'

export type DocumentPageSetupCommandParams = DocumentPageSetupRequest & Pick<Partial<DocumentEditorCommandRequest>,
  'unitId' | 'selection' | 'expectedRevision' | 'expectedHistoryGeneration' | 'inputGeneration'> & {
    /** Coordinates for selection; a supplied selection without host belongs to the body. */
    host?: DocumentAuxiliaryBlockHost
  }

/** The application injects its existing atomic editor transaction, retaining its private implementation. */
export function registerDocumentPageSetupCommand(api: OfficeStudio, commit: (request: DocumentEditorCommandRequest) => DocumentEditorCommandResult): void {
  api.registerCommand<unknown, DocumentEditorCommandResult | boolean>(DOCUMENT_PAGE_SETUP_COMMAND, raw => {
    if (raw === undefined || raw === null) return false
    if (typeof raw !== 'object' || Array.isArray(raw)) throw new TypeError('页面设置参数必须是对象。')
    if ('command' in raw) {
      assertDocumentPageSetupRequest(raw.command)
      const request = raw as DocumentEditorCommandRequest
      assertDocumentEditorCommandRequest(request, DOCUMENT_PAGE_SETUP_COMMAND)
      return commit(request)
    }
    assertDocumentPageSetupRequest(raw)
    const params = raw as DocumentPageSetupCommandParams
    const unitId = params.unitId ?? api.activeDocumentId
    if (typeof unitId !== 'string' || !unitId.trim()) throw new TypeError('页面设置的文档 ID 无效。')
    const model = api.getDocumentModel(unitId)
    if (!model) return false
    const state = api.getDocumentInputState(unitId)
    if (params.host !== undefined) validateDocumentBlockHost(params.host)
    const currentInput = state?.valid && state.modelRevision === api.getDocumentRevision(unitId) && state.location.kind === 'body' ? state.location : undefined
    const host = params.host ?? (!params.selection ? currentInput?.host : undefined)
    const scoped = host ? documentModelForBlockHost(model, host) : model
    const first = textTargets(scoped)[0] ?? (scoped.blocks[0] ? { blockId: scoped.blocks[0].id } : undefined)
    const currentSelection = currentInput && sameDocumentBlockHost(currentInput.host, host) ? currentInput.selection : undefined
    if (!params.selection && !currentSelection && !first) return false
    const selection = params.selection ?? currentSelection ?? collapsedSelection({ ...first!, offset: 0 })
    const retainStoryInput = !params.selection && params.host === undefined && state?.valid
      && state.modelRevision === api.getDocumentRevision(unitId) && state.location.kind === 'story'
    const request: DocumentEditorCommandRequest = {
      unitId, selection: cloneDocumentSelection(selection),
      // A flat explicit body selection must not inherit a current note/story
      // host; page settings always mutate the full document independently.
      ...(retainStoryInput ? {} : { selectionHost: host ? { ...host } : null }),
      command: { id: DOCUMENT_PAGE_SETUP_COMMAND, pageSize: params.pageSize, pageOrient: params.pageOrient, documentFlavor: params.documentFlavor,
        marginTop: params.marginTop, marginBottom: params.marginBottom, marginLeft: params.marginLeft, marginRight: params.marginRight },
      expectedRevision: params.expectedRevision ?? api.getDocumentRevision(unitId),
      expectedHistoryGeneration: params.expectedHistoryGeneration ?? api.editor.history.captureSavepoint(unitId).generation,
      inputGeneration: params.inputGeneration ?? state?.generation,
    }
    assertDocumentEditorCommandRequest(request, DOCUMENT_PAGE_SETUP_COMMAND)
    commit(request)
    return true
  })
}
