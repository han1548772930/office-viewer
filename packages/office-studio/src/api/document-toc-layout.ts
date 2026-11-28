import type { OfficeStudio } from './office-studio'
import type { DocumentModel } from '../domain/model'
import type { DocumentTocLayoutProvider, DocumentTocLayoutSnapshot } from '../features/document/toc/document-toc-pagination'

export type { DocumentTocLayoutProvider, DocumentTocLayoutSnapshot } from '../features/document/toc/document-toc-pagination'
export interface DocumentTocLayoutHandle { dispose(): void; isCurrent(): boolean }
export interface DocumentTocLayoutOwner {
  readonly provider: DocumentTocLayoutProvider
  readonly historyGeneration: number
  readonly signal: AbortSignal
}
interface Owner extends DocumentTocLayoutOwner { controller: AbortController }
const owners = new WeakMap<OfficeStudio, Map<string, Owner>>()

export function captureDocumentTocLayoutOwner(api: OfficeStudio, unitId: string): DocumentTocLayoutOwner | undefined {
  const owner = owners.get(api)?.get(unitId)
  return owner && documentTocLayoutOwnerCurrent(api, unitId, owner) ? owner : undefined
}

export function documentTocLayoutOwnerCurrent(api: OfficeStudio, unitId: string, owner: DocumentTocLayoutOwner): boolean {
  return !api.isDisposed && !owner.signal.aborted && owners.get(api)?.get(unitId) === owner && !!api.getDocumentModel(unitId)
    && api.editor.history.captureSavepoint(unitId).generation === owner.historyGeneration
}

/** A font/host publication can change between resolving a wait and its next
 * microtask; only the exact complete publication authorizes a page mutation. */
export function documentTocLayoutSnapshotCurrent(api: OfficeStudio, unitId: string, owner: DocumentTocLayoutOwner,
  snapshot: DocumentTocLayoutSnapshot): boolean {
  if (!documentTocLayoutOwnerCurrent(api, unitId, owner) || snapshot.unitId !== unitId
    || snapshot.source !== api.getDocumentModel(unitId) || snapshot.modelRevision !== api.getDocumentRevision(unitId)
    || snapshot.historyGeneration !== owner.historyGeneration || !snapshot.layout || snapshot.layout.inspectionHost
    || !Number.isSafeInteger(snapshot.layoutGeneration) || snapshot.layoutGeneration < 0
    || !Number.isSafeInteger(snapshot.fontVersion) || snapshot.fontVersion < 0) return false
  const latest = owner.provider.read()
  return latest?.complete === true && latest.source === snapshot.source && latest.unitId === snapshot.unitId
    && latest.modelRevision === snapshot.modelRevision && latest.historyGeneration === snapshot.historyGeneration
    && latest.layoutGeneration === snapshot.layoutGeneration && latest.fontVersion === snapshot.fontVersion
}

/** Replacement and same-ID reopen cannot inherit another view's layout task. */
export function bindDocumentTocLayout(api: OfficeStudio, unitId: string, provider: DocumentTocLayoutProvider): DocumentTocLayoutHandle {
  if (!api.getDocumentModel(unitId) || !provider || typeof provider.read !== 'function' || typeof provider.subscribe !== 'function') throw new TypeError('目录排版宿主无效。')
  let units = owners.get(api)
  if (!units) { units = new Map(); owners.set(api, units) }
  const previous = units.get(unitId), controller = new AbortController()
  const owner: Owner = { provider, controller, signal: controller.signal,
    historyGeneration: api.editor.history.captureSavepoint(unitId).generation }
  units.set(unitId, owner)
  previous?.controller.abort()
  return { isCurrent: () => documentTocLayoutOwnerCurrent(api, unitId, owner), dispose: () => {
    if (units?.get(unitId) === owner) { units.delete(unitId); if (!units.size) owners.delete(api) }
    controller.abort()
  } }
}

export function disposeDocumentTocLayouts(api: OfficeStudio, unitId?: string): void {
  const units = owners.get(api)
  if (!units) return
  if (unitId !== undefined) {
    units.get(unitId)?.controller.abort(); units.delete(unitId)
    if (!units.size) owners.delete(api)
  } else { owners.delete(api); for (const owner of units.values()) owner.controller.abort(); units.clear() }
}

export interface DocumentTocLayoutWaitTarget {
  unitId: string
  source: DocumentModel
  modelRevision: number
  historyGeneration: number
  /** A mutation must wait for a computation after its previous layout. */
  afterGeneration?: number
}

/** Reuse the live layout, including synchronous completion before subscribing.
 * A bounded wait mirrors the fixed source's 30 second layout timeout. It never
 * computes guessed headless pages or accepts a previous model/font result. */
export function waitForDocumentTocLayout(api: OfficeStudio, owner: DocumentTocLayoutOwner,
  target: DocumentTocLayoutWaitTarget, current: () => boolean, signal?: AbortSignal, timeoutMs = 30_000): Promise<DocumentTocLayoutSnapshot | undefined> {
  return new Promise(resolve => {
    let done = false, unsubscribe: (() => void) | undefined, timer: ReturnType<typeof setTimeout> | undefined
    const events: Array<{ dispose(): void }> = []
    const finish = (snapshot?: DocumentTocLayoutSnapshot) => {
      if (done) return
      done = true
      if (timer !== undefined) clearTimeout(timer)
      unsubscribe?.(); for (const event of events) event.dispose()
      signal?.removeEventListener('abort', cancel); owner.signal.removeEventListener('abort', cancel)
      resolve(snapshot)
    }
    const cancel = () => finish()
    const check = () => {
      if (done) return
      if (signal?.aborted || !documentTocLayoutOwnerCurrent(api, target.unitId, owner) || !current()) { finish(); return }
      const snapshot = owner.provider.read()
      if (snapshot?.complete && snapshot.source === target.source && snapshot.unitId === target.unitId
        && snapshot.modelRevision === target.modelRevision && snapshot.historyGeneration === target.historyGeneration
        && Number.isSafeInteger(snapshot.layoutGeneration) && snapshot.layoutGeneration > (target.afterGeneration ?? -1)
        && Number.isSafeInteger(snapshot.fontVersion) && snapshot.fontVersion >= 0 && snapshot.layout && !snapshot.layout.inspectionHost) finish(snapshot)
    }
    signal?.addEventListener('abort', cancel, { once: true }); owner.signal.addEventListener('abort', cancel, { once: true })
    for (const name of ['DocumentInputStateChanged', 'CommandExecuted', 'Undo', 'Redo', 'DocDisposed', 'LifeCycleChanged'] as const) events.push(api.events.addEvent(name, check))
    unsubscribe = owner.provider.subscribe(check)
    // A provider may synchronously publish while installing its subscription.
    if (done) { unsubscribe(); return }
    timer = setTimeout(cancel, timeoutMs)
    check()
  })
}
