import type { DocumentModel } from '../../../domain/model'
import type { DocumentSelection } from '../model/document-model'

export interface DocumentFieldCacheState {
  unitId: string
  model: DocumentModel
  modelRevision: number
  historyGeneration: number
  inputGeneration?: number
}

export interface DocumentFieldCacheRefreshRequest {
  unitId: string
  expectedRevision: number
  expectedHistoryGeneration: number
  inputGeneration?: number
  locale: string
  timestamp: number
  selection: DocumentSelection
}

export type DocumentFieldCacheRefreshResult =
  | { accepted: false; changed: false }
  | { accepted: true; changed: boolean; model: DocumentModel; selection: DocumentSelection;
    selectionCurrent: boolean; modelRevision: number; historyGeneration: number }

/** A standalone workbench shares its host's authoritative model and history.
 * refresh prepares and publishes derived caches, including retained undo/redo
 * snapshots, without changing dirty/savepoint or typing marks. A stale unit,
 * revision, history or input generation is rejected before publication.
 */
export interface DocumentFieldCacheAdapter {
  /** Stable across workbench remounts and wrappers for the same host. */
  readonly identity: object
  /** Read reactive host state so same-ID reopen generations remain observable. */
  getState(): DocumentFieldCacheState | undefined
  refresh(request: DocumentFieldCacheRefreshRequest): DocumentFieldCacheRefreshResult
}

/** A host's regional settings or adapter identity survive Facade wrappers. */
const refreshed = new WeakMap<object, Map<string, number>>()

/** Reserve before publishing so synchronous observers cannot refresh the same open twice. */
export function reserveDocumentFieldOpenRefresh(host: object, unitId: string, historyGeneration: number): boolean {
  if (!host || typeof host !== 'object' || !unitId.trim() || !Number.isSafeInteger(historyGeneration) || historyGeneration < 1) {
    throw new TypeError('文档打开字段更新的宿主、标识或历史代数无效。')
  }
  let units = refreshed.get(host)
  if (!units) { units = new Map(); refreshed.set(host, units) }
  if (units.get(unitId) === historyGeneration) return false
  units.set(unitId, historyGeneration)
  return true
}

/** A failed attempt can retry; an older attempt cannot release a newer open. */
export function releaseDocumentFieldOpenRefresh(host: object, unitId: string, historyGeneration: number): boolean {
  const units = refreshed.get(host)
  if (units?.get(unitId) !== historyGeneration) return false
  units.delete(unitId)
  if (!units.size) refreshed.delete(host)
  return true
}
