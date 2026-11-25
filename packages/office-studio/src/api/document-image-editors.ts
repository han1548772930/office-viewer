import type { OfficeStudio } from './office-studio'
import type { DocumentModel } from '../domain/model'
import type { DocumentSelection } from '../features/document/model/document-model'
import type { DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'

export type DocumentImageEditorMode = 'insert' | 'edit' | 'crop' | 'replace'
export interface DocumentImageEditorTarget {
  unitId: string; source: DocumentModel; revision: number; historyGeneration: number; inputGeneration?: number
  segmentId: string; selection: DocumentSelection; host?: DocumentAuxiliaryBlockHost; imageId?: string
  mode: DocumentImageEditorMode; paragraphMenuPlacement?: 'below'
}
export interface DocumentImageEditor { open(target: DocumentImageEditorTarget): boolean; close(force?: boolean): boolean }
export interface DocumentImageEditorHandle { dispose(): void; isCurrent(): boolean }
const editors = new WeakMap<OfficeStudio, Map<string, DocumentImageEditor>>()
const tasks = new Map<DocumentModel, Set<AbortController>>()

/** Only transient IO owners are retained; completed and cancelled tasks release their source. */
export function trackDocumentImageTask(source: DocumentModel, controller: AbortController): () => void {
  let active = tasks.get(source)
  if (!active) { active = new Set(); tasks.set(source, active) }
  active.add(controller)
  return () => { active!.delete(controller); if (!active!.size) tasks.delete(source) }
}
export function cancelDocumentImageTasksForSource(source: DocumentModel): void {
  const active = tasks.get(source)
  tasks.delete(source)
  for (const controller of active ?? []) controller.abort()
}
/** Call before removing the document model from its studio. */
export function disposeDocumentImageTasks(api: OfficeStudio, unitId?: string): void {
  for (const source of [...tasks.keys()]) {
    const id = source.id
    if (typeof id === 'string' && (unitId === undefined || id === unitId) && api.getDocumentModel(id) === source) cancelDocumentImageTasksForSource(source)
  }
}
export function registerDocumentImageEditor(api: OfficeStudio, unitId: string, editor: DocumentImageEditor): DocumentImageEditorHandle {
  if (!api.getDocumentModel(unitId) || !editor || typeof editor.open !== 'function' || typeof editor.close !== 'function') throw new TypeError('图片编辑器目标无效。')
  let units = editors.get(api)
  if (!units) { units = new Map(); editors.set(api, units) }
  const previous = units.get(unitId), owner: DocumentImageEditor = { open: target => editor.open(target), close: force => editor.close(force) }
  units.set(unitId, owner); previous?.close(true)
  return { isCurrent: () => editors.get(api)?.get(unitId) === owner, dispose: () => {
    const current = editors.get(api)
    if (current?.get(unitId) !== owner) return
    current.delete(unitId); if (!current.size) editors.delete(api)
    owner.close(true)
  } }
}
export function requestDocumentImageEditorOpen(api: OfficeStudio, target: DocumentImageEditorTarget): boolean {
  return api.getDocumentModel(target.unitId) === target.source && api.getDocumentRevision(target.unitId) === target.revision
    && api.editor.history.captureSavepoint(target.unitId).generation === target.historyGeneration
    ? editors.get(api)?.get(target.unitId)?.open(target) ?? false : false
}
export function disposeDocumentImageEditors(api: OfficeStudio, unitId?: string): void {
  disposeDocumentImageTasks(api, unitId)
  const current = editors.get(api)
  if (!current) return
  const owners = unitId === undefined ? [...current.values()] : [current.get(unitId)].filter((owner): owner is DocumentImageEditor => !!owner)
  if (unitId === undefined) editors.delete(api)
  else { current.delete(unitId); if (!current.size) editors.delete(api) }
  for (const owner of owners) owner.close(true)
}
