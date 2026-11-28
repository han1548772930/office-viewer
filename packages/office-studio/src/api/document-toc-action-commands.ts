import type { OfficeStudio } from './office-studio'
import { assertDocumentEditorCommandRequest, type DocumentEditorCommandRequest, type DocumentEditorCommandResult } from './document-editor-commands'
import type { HistorySavepoint } from '../platform/history/history-types'
import type { DocumentModel } from '../domain/model'
import { blockAtPoint, collapsedSelection, orderedSelection, type DocumentSelection } from '../features/document/model/document-model'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection'
import { documentTocIdAtBlock, documentTocPageNumbers } from '../features/document/toc/document-toc'
import { documentTextFieldParagraphs, documentTextFieldRanges } from '../features/document/model/document-text-fields'
import { captureDocumentTocLayoutOwner, documentTocLayoutOwnerCurrent, documentTocLayoutSnapshotCurrent, waitForDocumentTocLayout } from './document-toc-layout'
import { cancelDocumentTocPagination, documentTocCommandAmendment, scheduleDocumentTocPagination } from './document-toc-pagination'

export const DOCUMENT_TOC_UPDATE_COMMAND = 'doc.command.update-table-of-contents'
export const DOCUMENT_TOC_DELETE_COMMAND = 'doc.command.delete-table-of-contents'
type Guards = Pick<Partial<DocumentEditorCommandRequest>, 'unitId' | 'expectedRevision' | 'expectedHistoryGeneration' | 'inputGeneration'>
export interface DocumentTocDeleteCommandParams extends Guards { rangeId?: string }
export interface DocumentTocUpdateCommandParams extends DocumentTocDeleteCommandParams { mode?: 'pageNumbersOnly' | 'entireTable' }

/** Public custom-range IDs and the internal stable TOC ID identify the same
 * field. A retained native outer owner can sit outside the generated rows. */
function resolveTocId(model: DocumentModel, rangeId: string | undefined, selection: DocumentSelection | undefined): string | undefined {
  if (rangeId === undefined) return selection && documentTocIdAtBlock(blockAtPoint(model, orderedSelection(model, selection).start))
  const direct = model.blocks.find(block => documentTocIdAtBlock(block) === rangeId)
  if (direct) return documentTocIdAtBlock(direct)
  const range = documentTextFieldRanges(model).find(candidate => candidate.field.id === rangeId && candidate.field.field.toUpperCase() === 'TOC')
  if (!range) return undefined
  const ids = new Set(documentTextFieldParagraphs(model).filter(paragraph => paragraph.absoluteStart >= range.absoluteStart
    && paragraph.absoluteStart <= range.absoluteEnd).map(paragraph => documentTocIdAtBlock(paragraph.block)).filter(Boolean))
  return ids.size === 1 ? [...ids][0] : undefined
}

export function registerDocumentTocActionCommands(api: OfficeStudio,
  commit: (request: DocumentEditorCommandRequest, amendment?: HistorySavepoint) => DocumentEditorCommandResult): void {
  for (const commandId of [DOCUMENT_TOC_UPDATE_COMMAND, DOCUMENT_TOC_DELETE_COMMAND] as const) {
    api.registerCommand<unknown, DocumentEditorCommandResult | boolean | Promise<boolean>>(commandId, raw => {
      if (raw !== undefined && (!raw || typeof raw !== 'object' || Array.isArray(raw))) throw new TypeError('目录操作参数必须为对象。')
      const input = (raw ?? {}) as DocumentTocUpdateCommandParams
      if ('command' in input) {
        const request = input as DocumentEditorCommandRequest
        assertDocumentEditorCommandRequest(request, commandId)
        const amendment = documentTocCommandAmendment(request)
        const previous = api.getDocumentModel(request.unitId)
        const generation = captureDocumentTocLayoutOwner(api, request.unitId)?.provider.read()?.layoutGeneration
        if (!amendment) cancelDocumentTocPagination(api, request.unitId)
        const result = commit(request, amendment)
        if (!amendment && commandId === DOCUMENT_TOC_UPDATE_COMMAND && previous && result.changed) {
          void scheduleDocumentTocPagination(api, request, previous, result, generation)
        }
        return result
      }
      const unitId = input.unitId ?? api.activeDocumentId
      if (typeof unitId !== 'string' || !unitId.trim()
        || input.rangeId !== undefined && (typeof input.rangeId !== 'string' || !input.rangeId.trim())
        || commandId === DOCUMENT_TOC_UPDATE_COMMAND && input.mode !== undefined && !['pageNumbersOnly', 'entireTable'].includes(input.mode)) {
        throw new TypeError('目录目标或更新模式无效。')
      }
      const source = api.getDocumentModel(unitId)
      if (!source) return false
      const revision = api.getDocumentRevision(unitId), savepoint = api.editor.history.captureSavepoint(unitId)
      const session = api.getDocumentInputState(unitId)
      const bodySelection = session?.valid && session.modelRevision === revision && session.location.kind === 'body' && !session.location.host
        ? session.location.selection : undefined
      const tocId = resolveTocId(source, input.rangeId, bodySelection)
      const block = tocId && source.blocks.find(candidate => documentTocIdAtBlock(candidate) === tocId)
      if (!tocId || !block) return false
      const selection = bodySelection ? cloneDocumentSelection(bodySelection) : collapsedSelection({ blockId: block.id, offset: 0 })
      const request: DocumentEditorCommandRequest = {
        unitId, selection, selectionHost: null,
        command: commandId === DOCUMENT_TOC_DELETE_COMMAND ? { id: commandId, tocId, preserveInputSelection: true }
          : { id: commandId, tocId, mode: input.mode === 'pageNumbersOnly' ? 'pageNumbersOnly' : 'all', preserveInputSelection: true },
        expectedRevision: input.expectedRevision === undefined ? revision : input.expectedRevision,
        expectedHistoryGeneration: input.expectedHistoryGeneration === undefined ? savepoint.generation : input.expectedHistoryGeneration,
        inputGeneration: input.inputGeneration === undefined ? session?.generation : input.inputGeneration,
      }
      assertDocumentEditorCommandRequest(request, commandId)
      const userId = api.currentUserId, activeId = api.activeDocumentId, activeKind = api.editor.state.activeKind
      const history = api.editor.history.units[unitId], sessionToken = JSON.stringify(session), savedRevision = history?.savedRevision
      const current = () => !api.isDisposed && api.getDocumentModel(unitId) === source && api.getDocumentRevision(unitId) === revision
        && api.editor.history.units[unitId] === history && api.editor.history.captureSavepoint(unitId).generation === savepoint.generation
        && api.editor.history.captureSavepoint(unitId).revision === savepoint.revision && history?.savedRevision === savedRevision
        && api.currentUserId === userId && api.activeDocumentId === activeId && api.editor.state.activeKind === activeKind
        && JSON.stringify(api.getDocumentInputState(unitId)) === sessionToken
      if (commandId === DOCUMENT_TOC_DELETE_COMMAND) return Promise.resolve().then(() => {
        if (!current()) return false
        cancelDocumentTocPagination(api, unitId)
        return commit(request).changed
      })
      const owner = captureDocumentTocLayoutOwner(api, unitId)
      if (!owner) return false
      return waitForDocumentTocLayout(api, owner, { unitId, source, modelRevision: revision, historyGeneration: savepoint.generation },
        () => current() && documentTocLayoutOwnerCurrent(api, unitId, owner)).then(async snapshot => {
        if (!snapshot || !current() || !documentTocLayoutSnapshotCurrent(api, unitId, owner, snapshot)
          || request.command.id !== DOCUMENT_TOC_UPDATE_COMMAND) return false
        request.command = { ...request.command, pageNumbers: documentTocPageNumbers(snapshot.layout!.headings, snapshot.layout!.pages) }
        cancelDocumentTocPagination(api, unitId)
        const result = commit(request)
        if (!result.changed) return true
        const settled = await scheduleDocumentTocPagination(api, request, source, result, snapshot.layoutGeneration)
        return settled.status === 'settled' || settled.status === 'limit'
      })
    })
  }
}
