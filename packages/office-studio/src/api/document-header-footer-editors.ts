import type { OfficeStudio } from './office-studio'
import { DOCUMENT_HEADER_FOOTER_COMMANDS, DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS, resolveDocumentHeaderFooterRequest,
  type DocumentHeaderFooterEditorTarget, type DocumentHeaderFooterRequest } from '../features/document/controllers/document-header-footer-editor'

export type { DocumentHeaderFooterEditorTarget } from '../features/document/controllers/document-header-footer-editor'
export interface DocumentHeaderFooterEditor { execute(target: DocumentHeaderFooterEditorTarget): boolean; release(): void }
export interface DocumentHeaderFooterEditorHandle { execute(commandId: string, params?: unknown): boolean; dispose(): void; isCurrent(): boolean }
interface Owner { editor: DocumentHeaderFooterEditor; historyGeneration: number; released: boolean }
const editors = new WeakMap<OfficeStudio, Map<string, Owner>>()
const targetOwners = new WeakMap<DocumentHeaderFooterEditorTarget, Owner | undefined>()

function currentOwner(api: OfficeStudio, unitId: string, owner: Owner): boolean {
  return !owner.released && editors.get(api)?.get(unitId) === owner && !!api.getDocumentModel(unitId)
    && api.editor.history.captureSavepoint(unitId).generation === owner.historyGeneration
}
function releaseOwner(owner: Owner) {
  if (owner.released) return
  owner.released = true
  owner.editor.release()
}

export function captureDocumentHeaderFooterEditorTarget(api: OfficeStudio, request: DocumentHeaderFooterRequest): DocumentHeaderFooterEditorTarget {
  const unitId = request.unitId ?? api.activeDocumentId
  const source = typeof unitId === 'string' && unitId.trim() ? api.getDocumentModel(unitId) : undefined
  if (!source) throw new RangeError('页眉页脚的目标文档不存在。')
  const revision = api.getDocumentRevision(unitId), historyGeneration = api.editor.history.captureSavepoint(unitId).generation
  const input = api.getDocumentInputState(unitId), inputGeneration = input?.valid ? input.generation : undefined
  if (request.expectedRevision !== undefined && request.expectedRevision !== revision
    || request.expectedHistoryGeneration !== undefined && request.expectedHistoryGeneration !== historyGeneration
    || request.expectedInputGeneration !== undefined && request.expectedInputGeneration !== inputGeneration) throw new RangeError('页眉页脚的内容、历史或输入会话已变化。')
  const target: DocumentHeaderFooterEditorTarget = { ...request, unitId, source, revision, historyGeneration,
    ...(inputGeneration === undefined ? {} : { inputGeneration }) }
  targetOwners.set(target, editors.get(api)?.get(unitId))
  return target
}

/** Extension hooks cannot replace the editor owner between capture and dispatch. */
export function documentHeaderFooterEditorTargetCurrent(api: OfficeStudio, target: DocumentHeaderFooterEditorTarget): boolean {
  const owner = editors.get(api)?.get(target.unitId), input = api.getDocumentInputState(target.unitId)
  return targetOwners.has(target) && targetOwners.get(target) === owner
    && (!owner || currentOwner(api, target.unitId, owner)) && api.getDocumentModel(target.unitId) === target.source
    && api.getDocumentRevision(target.unitId) === target.revision
    && api.editor.history.captureSavepoint(target.unitId).generation === target.historyGeneration
    && (input?.valid ? input.generation : undefined) === target.inputGeneration
    && (!input?.valid || input.modelRevision === target.revision)
}

/** Real workbench owners receive a captured target; an absent owner never reports a successful operation. */
export function requestDocumentHeaderFooterEditor(api: OfficeStudio, target: DocumentHeaderFooterEditorTarget): boolean {
  const owner = editors.get(api)?.get(target.unitId)
  if (!owner || !documentHeaderFooterEditorTargetCurrent(api, target)) return false
  const accepted = owner.editor.execute(target)
  return accepted && currentOwner(api, target.unitId, owner) && api.getDocumentModel(target.unitId) === target.source
    && api.getDocumentRevision(target.unitId) === target.revision
}

export function registerDocumentHeaderFooterEditor(api: OfficeStudio, unitId: string, editor: DocumentHeaderFooterEditor): DocumentHeaderFooterEditorHandle {
  if (!api.getDocumentModel(unitId) || !editor || typeof editor.execute !== 'function' || typeof editor.release !== 'function') throw new TypeError('页眉页脚编辑器目标无效。')
  let units = editors.get(api)
  if (!units) { units = new Map(); editors.set(api, units) }
  const previous = units.get(unitId), owner: Owner = { editor: { execute: target => editor.execute(target), release: () => editor.release() },
    historyGeneration: api.editor.history.captureSavepoint(unitId).generation, released: false }
  units.set(unitId, owner)
  if (previous) releaseOwner(previous)
  return {
    isCurrent: () => currentOwner(api, unitId, owner),
    execute: (commandId, params) => {
      if (!currentOwner(api, unitId, owner)) return false
      const request = resolveDocumentHeaderFooterRequest(commandId, params)
      if (request.unitId !== undefined && request.unitId !== unitId) throw new RangeError('页眉页脚句柄的目标文档不一致。')
      const target = captureDocumentHeaderFooterEditorTarget(api, { ...request, unitId })
      const canonical = target.action === 'open' ? DOCUMENT_HEADER_FOOTER_COMMANDS.open
        : target.action === 'close' ? DOCUMENT_HEADER_FOOTER_COMMANDS.close : DOCUMENT_HEADER_FOOTER_COMMANDS.ribbon
      const guarded = { unitId, expectedRevision: target.revision, expectedHistoryGeneration: target.historyGeneration,
        ...(target.inputGeneration === undefined ? {} : { expectedInputGeneration: target.inputGeneration }),
        ...(target.action === 'open' ? { editArea: target.editArea } : target.action === 'ribbon' ? { value: target.visible ? 'open' : 'close' } : {}) }
      return api.executeCommandSync<typeof guarded, boolean>(canonical, guarded)
    },
    dispose: () => {
      const current = editors.get(api)
      if (current?.get(unitId) === owner) {
        current.delete(unitId)
        if (!current.size) editors.delete(api)
      }
      releaseOwner(owner)
    },
  }
}

export function disposeDocumentHeaderFooterEditors(api: OfficeStudio, unitId?: string): void {
  const current = editors.get(api)
  if (!current) return
  const owners = unitId === undefined ? [...current.values()] : [current.get(unitId)].filter((owner): owner is Owner => !!owner)
  if (unitId === undefined) editors.delete(api)
  else { current.delete(unitId); if (!current.size) editors.delete(api) }
  for (const owner of owners) releaseOwner(owner)
}

export { DOCUMENT_HEADER_FOOTER_COMMANDS, DOCUMENT_HEADER_FOOTER_LOCAL_COMMANDS }
