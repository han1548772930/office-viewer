import type { OfficeStudio } from './office-studio'
import { assertDocumentEditorCommandRequest, type DocumentEditorCommandRequest, type DocumentEditorCommandResult } from './document-editor-commands'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection'
import type { DocumentTocInsertionRequest } from '../features/document/toc/document-toc-command'
import { documentTocPageNumbers } from '../features/document/toc/document-toc'
import { captureDocumentTocLayoutOwner, documentTocLayoutOwnerCurrent, documentTocLayoutSnapshotCurrent, waitForDocumentTocLayout } from './document-toc-layout'
import { scheduleDocumentTocPagination } from './document-toc-pagination'

export const DOCUMENT_TOC_INSERT_COMMAND = 'doc.command.insert-table-of-contents'

/** Fixed upstream flat options plus local layout labels and optimistic guards. */
export type DocumentTocInsertCommandParams = DocumentTocInsertionRequest & Pick<Partial<DocumentEditorCommandRequest>,
  'unitId' | 'expectedRevision' | 'expectedHistoryGeneration' | 'inputGeneration'>

/** Resolve only an existing body input session; never guess a caret or redirect a
 * note/header/footer selection. The Studio captures its owner before hooks run. */
export function registerDocumentTocInsertCommand(api: OfficeStudio, commit: (request: DocumentEditorCommandRequest) => DocumentEditorCommandResult): void {
  const insert = (request: DocumentEditorCommandRequest): DocumentEditorCommandResult => {
    const previous = api.getDocumentModel(request.unitId)
    const layoutGeneration = captureDocumentTocLayoutOwner(api, request.unitId)?.provider.read()?.layoutGeneration
    const result = commit(request)
    if (previous && result.changed) void scheduleDocumentTocPagination(api, request, previous, result, layoutGeneration)
    return result
  }
  api.registerCommand<unknown, DocumentEditorCommandResult | boolean | Promise<boolean>>(DOCUMENT_TOC_INSERT_COMMAND, raw => {
    if (raw !== undefined && (!raw || typeof raw !== 'object' || Array.isArray(raw))) throw new TypeError('目录插入参数必须是对象。')
    const input = (raw ?? {}) as DocumentTocInsertCommandParams
    if ('command' in input) {
      const request = input as DocumentEditorCommandRequest
      assertDocumentEditorCommandRequest(request, DOCUMENT_TOC_INSERT_COMMAND)
      return insert(request)
    }
    const unitId = input.unitId ?? api.activeDocumentId
    if (typeof unitId !== 'string' || !unitId.trim()) throw new TypeError('目录插入的文档 ID 无效。')
    if (!api.getDocumentModel(unitId)) return false
    const revision = api.getDocumentRevision(unitId), session = api.getDocumentInputState(unitId)
    if (!session?.valid || session.modelRevision !== revision || session.location.kind !== 'body' || session.location.host) return false
    const request: DocumentEditorCommandRequest = {
      unitId, selection: cloneDocumentSelection(session.location.selection), selectionHost: null,
      command: { id: DOCUMENT_TOC_INSERT_COMMAND, levels: input.levels, minimumLevel: input.minimumLevel,
        showPageNumbers: input.showPageNumbers, rightAlignPageNumbers: input.rightAlignPageNumbers,
        tabLeader: input.tabLeader, format: input.format, title: input.title, pageNumbers: input.pageNumbers },
      expectedRevision: input.expectedRevision === undefined ? revision : input.expectedRevision,
      expectedHistoryGeneration: input.expectedHistoryGeneration === undefined ? api.editor.history.captureSavepoint(unitId).generation : input.expectedHistoryGeneration,
      inputGeneration: input.inputGeneration === undefined ? session.generation : input.inputGeneration,
    }
    assertDocumentEditorCommandRequest(request, DOCUMENT_TOC_INSERT_COMMAND)
    const owner = captureDocumentTocLayoutOwner(api, unitId), source = api.getDocumentModel(unitId)
    if (!owner || !source) return false
    const savepoint = api.editor.history.captureSavepoint(unitId), userId = api.currentUserId
    const activeId = api.activeDocumentId, activeKind = api.editor.state.activeKind
    const current = () => documentTocLayoutOwnerCurrent(api, unitId, owner)
      && api.getDocumentModel(unitId) === source && api.getDocumentRevision(unitId) === revision
      && api.editor.history.captureSavepoint(unitId).revision === savepoint.revision
      && api.currentUserId === userId && api.activeDocumentId === activeId && api.editor.state.activeKind === activeKind
      && (() => { const latest = api.getDocumentInputState(unitId)
        return latest?.valid === true && latest.modelRevision === revision && latest.generation === session.generation
          && latest.location.kind === 'body' && !latest.location.host
          && JSON.stringify(latest.location.selection) === JSON.stringify(request.selection) })()
    return waitForDocumentTocLayout(api, owner, { unitId, source, modelRevision: revision, historyGeneration: savepoint.generation }, current)
      .then(async snapshot => {
        if (!snapshot || !current() || !documentTocLayoutSnapshotCurrent(api, unitId, owner, snapshot)) return false
        if (request.command.id !== DOCUMENT_TOC_INSERT_COMMAND) return false
        request.command = { ...request.command, pageNumbers: documentTocPageNumbers(snapshot.layout!.headings, snapshot.layout!.pages) }
        const result = commit(request)
        if (!result.changed) return false
        // Match the fixed source: insertion succeeds once committed; a later
        // edit can cancel its page refresh while the public call still waits.
        await scheduleDocumentTocPagination(api, request, source, result, snapshot.layoutGeneration)
        return true
      })
  })
}
