import type { OfficeStudio } from './office-studio'
import type { DocumentEditorCommandRequest, DocumentEditorCommandResult } from './document-editor-commands'
import type { HistorySavepoint } from '../platform/history/history-types'
import type { DocumentModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { applyDocumentEditorCommand } from './document-editor-commands'
import { cloneDocumentSelection } from '../features/document/model/document-table-selection'
import { documentTocIdAtBlock, documentTocPageNumbers } from '../features/document/toc/document-toc'
import type { DocumentTocPaginationResult } from '../features/document/toc/document-toc-pagination'
import { captureDocumentTocLayoutOwner, documentTocLayoutOwnerCurrent, waitForDocumentTocLayout } from './document-toc-layout'

export type { DocumentTocPaginationResult } from '../features/document/toc/document-toc-pagination'
interface Job { controller: AbortController; promise: Promise<DocumentTocPaginationResult> }
const jobs = new WeakMap<OfficeStudio, Map<string, Job>>()
const amendments = new WeakMap<DocumentEditorCommandRequest, { savepoint: HistorySavepoint; current(): boolean; payload: string }>()

/** Private capability: JSON DTOs cannot nominate an arbitrary history entry.
 * Resolve after BeforeCommandExecute so hooks cannot change the target/payload. */
export function documentTocCommandAmendment(request: DocumentEditorCommandRequest): HistorySavepoint | undefined {
  const token = amendments.get(request)
  if (!token) return undefined
  if (!token.current() || token.payload !== JSON.stringify(request)) throw new Error('目录重排的文档、输入或历史已变化。')
  return token.savepoint
}

export function waitForDocumentTocPagination(api: OfficeStudio, unitId: string): Promise<DocumentTocPaginationResult | undefined> {
  return jobs.get(api)?.get(unitId)?.promise ?? Promise.resolve(undefined)
}

export function cancelDocumentTocPagination(api: OfficeStudio, unitId?: string): void {
  const units = jobs.get(api)
  if (!units) return
  if (unitId === undefined) { jobs.delete(api); for (const job of units.values()) job.controller.abort(); units.clear() }
  else { units.get(unitId)?.controller.abort(); units.delete(unitId); if (!units.size) jobs.delete(api) }
}

/** The first edit is authoritative immediately. Re-pagination only amends
 * its exact top transaction, never a later user edit, undo/redo, save or reopen. */
export function scheduleDocumentTocPagination(api: OfficeStudio, request: DocumentEditorCommandRequest,
  previous: DocumentModel, insertion: DocumentEditorCommandResult, afterGeneration?: number): Promise<DocumentTocPaginationResult> {
  const unitId = request.unitId, oldIds = new Set(previous.blocks.map(documentTocIdAtBlock).filter(Boolean))
  const source = api.getDocumentModel(unitId), insertedIds = [...new Set(source?.blocks.map(documentTocIdAtBlock).filter((id): id is string => !!id && !oldIds.has(id)) ?? [])]
  const updating = request.command.id === 'doc.command.update-table-of-contents' ? request.command : undefined
  const tocId = updating?.tocId ?? (insertedIds.length === 1 ? insertedIds[0]! : '')
  const mode = updating?.mode ?? 'all'
  // A committed observer may already have closed the unit. Do not recreate
  // its content history merely to record a canceled background refresh.
  if (api.isDisposed || !source) return Promise.resolve<DocumentTocPaginationResult>({ unitId, tocId, status: 'canceled',
    iterations: updating ? 1 : 0, modelRevision: insertion.modelRevision ?? 0 })
  const controller = new AbortController(), owner = captureDocumentTocLayoutOwner(api, unitId)
  let units = jobs.get(api)
  if (!units) { units = new Map(); jobs.set(api, units) }
  units.get(unitId)?.controller.abort()
  let expectedSource: DocumentModel | undefined = source, expectedRevision = insertion.modelRevision ?? api.getDocumentRevision(unitId)
  let savepoint = api.editor.history.captureSavepoint(unitId)
  let input = api.getDocumentInputState(unitId), selection = cloneDocumentSelection(insertion.selection)
  const userId = api.currentUserId, activeId = api.activeDocumentId, activeKind = api.editor.state.activeKind
  const history = api.editor.history.units[unitId]
  const savedRevision = history?.savedRevision
  // Manual updating has already applied its first pass; insertion gets three
  // subsequent refresh passes, matching the two fixed upstream entry points.
  let iterations = updating ? 1 : 0
  const result = (status: DocumentTocPaginationResult['status'], error?: string): DocumentTocPaginationResult =>
    ({ unitId, tocId, status, iterations, modelRevision: expectedRevision, ...(error ? { error } : {}) })
  const current = () => !controller.signal.aborted && !api.isDisposed && !!owner && documentTocLayoutOwnerCurrent(api, unitId, owner)
    && !!expectedSource && api.getDocumentModel(unitId) === expectedSource && api.getDocumentRevision(unitId) === expectedRevision
    && api.editor.history.units[unitId] === history && api.editor.history.canAmend(unitId, savepoint)
    && history?.savedRevision === savedRevision
    && api.editor.history.captureSavepoint(unitId).revision === savepoint.revision
    && api.currentUserId === userId && api.activeDocumentId === activeId && api.editor.state.activeKind === activeKind
    && (() => { const latest = api.getDocumentInputState(unitId)
      if (!input) return !!updating && latest === undefined
      if (!input.valid) return !!updating && JSON.stringify(latest) === JSON.stringify(input)
      if (!updating && (input.location.kind !== 'body' || input.location.host)) return false
      return latest?.valid === true && latest.modelRevision === expectedRevision && latest.generation === input.generation
        && JSON.stringify(latest.location) === JSON.stringify(input.location) })()
  const buildRequest = (pageNumbers: ReturnType<typeof documentTocPageNumbers>): DocumentEditorCommandRequest => ({
    unitId, selection: cloneDocumentSelection(selection), selectionHost: null,
    command: { id: 'doc.command.update-table-of-contents', tocId, mode, pageNumbers, preserveInputSelection: true },
    expectedRevision, expectedHistoryGeneration: savepoint.generation, inputGeneration: input?.generation,
  })
  const promise = Promise.resolve().then(async () => {
    if (!owner) return result('unavailable')
    if (!tocId || !insertion.changed || insertion.selectionCurrent === false || !current()) return result('canceled')
    try {
      let generation = afterGeneration
      for (; iterations < 3;) {
        const snapshot = await waitForDocumentTocLayout(api, owner, { unitId, source: expectedSource!, modelRevision: expectedRevision,
          historyGeneration: savepoint.generation, afterGeneration: generation }, current, controller.signal)
        if (!snapshot || !current()) return result('canceled')
        generation = snapshot.layoutGeneration
        const update = buildRequest(documentTocPageNumbers(snapshot.layout!.headings, snapshot.layout!.pages))
        const layoutCurrent = () => {
          const latest = owner.provider.read()
          return current() && latest?.complete === true && latest.source === snapshot.source
            && latest.modelRevision === snapshot.modelRevision && latest.historyGeneration === snapshot.historyGeneration
            && latest.layoutGeneration === snapshot.layoutGeneration && latest.fontVersion === snapshot.fontVersion
        }
        amendments.set(update, { savepoint, current: layoutCurrent, payload: JSON.stringify(update) })
        let committed: DocumentEditorCommandResult
        try { committed = api.executeCommandSync(update.command.id, update) } finally { amendments.delete(update) }
        if (!committed.changed) return current() ? result('settled') : result('canceled')
        iterations++
        // Command observers run before the call returns; never adopt their
        // newer content/session as this task's own continuation.
        if (committed.selectionCurrent === false || committed.modelRevision !== expectedRevision + 1
          || api.getDocumentRevision(unitId) !== committed.modelRevision) return result('canceled')
        expectedSource = api.getDocumentModel(unitId); expectedRevision = committed.modelRevision
        savepoint = api.editor.history.captureSavepoint(unitId); input = api.getDocumentInputState(unitId)
        selection = cloneDocumentSelection(committed.selection)
        if (!current()) return result('canceled')
      }
      // The fixed source caps updates at three. Read the final real layout to
      // distinguish convergence from the cap without publishing a fourth edit.
      const snapshot = await waitForDocumentTocLayout(api, owner, { unitId, source: expectedSource!, modelRevision: expectedRevision,
        historyGeneration: savepoint.generation, afterGeneration: generation }, current, controller.signal)
      if (!snapshot || !current()) return result('canceled')
      const candidate = cloneModel(expectedSource!)
      applyDocumentEditorCommand(candidate, buildRequest(documentTocPageNumbers(snapshot.layout!.headings, snapshot.layout!.pages)),
        { documentId: unitId, userId, permissionAdminUserId: api.permissionAdminUserId })
      return result(JSON.stringify(candidate) === JSON.stringify(expectedSource) ? 'settled' : 'limit')
    } catch (error) { return current() ? result('rejected', error instanceof Error ? error.message : String(error)) : result('canceled') }
  })
  units.set(unitId, { controller, promise })
  return promise
}
