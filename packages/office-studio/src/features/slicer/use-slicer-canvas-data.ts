import { computed, onScopeDispose, ref, watch } from 'vue'
import { FOffice } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { SlicerModel } from '../../domain/model.ts'
import { SpreadsheetGeometry, type SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout.ts'
import { slicerDrawingGeometry } from './slicer-drawing.ts'
import { SlicerCanvasDataCache, sameSlicerCanvasDataOwner, type SlicerCanvasDataLoader,
  type SlicerCanvasDataOptions, type SlicerCanvasDataOwner } from './slicer-canvas-data.ts'
import { assertSlicerCanvasWindow } from './slicer-canvas-data-validation.ts'
import { slicerMemberViewport } from './slicer-member-viewport.ts'
import type { SlicerCanvasControl } from './slicer-renderer.ts'
import { SlicerConnectionError } from './slicer-connections.ts'
import { assertSlicerSourceReadPermission } from './slicer-permissions.ts'

export interface SlicerCanvasDataContext {
  studio(): OfficeStudio | undefined
  unitId(): string | undefined
  sheetId(): string
  sheet(): SpreadsheetRenderSheet | undefined
  view(): unknown
  active(): boolean
}
interface RequestedWindow { start: number; count?: number }
interface Binding { owner: SlicerCanvasDataOwner; identity: string; loader: SlicerCanvasDataLoader }

function sourceIdentity(studio: OfficeStudio, unitId: string, model: SlicerModel): string {
  const cache = studio.getWorkbookModel(unitId)?.slicerCaches?.find(cache => cache.id === model.cacheId)
  return JSON.stringify([model.id, model.type, model.hostSheetId, model.cacheId, model.sourceFieldId,
    model.tableId, model.pivotTableId, cache?.source, cache?.options])
}

/** Vue observes ownership/windows. Source identities, filtering and member computation stay in the shared service/Worker. */
export function useSlicerCanvasData(context: SlicerCanvasDataContext, options: Omit<SlicerCanvasDataOptions, 'onChange'> = {}) {
  const changed = ref(0), sourceChanged = ref(0), requested = ref(new Map<string, RequestedWindow>()),
    cache = new SlicerCanvasDataCache({ ...options, onChange: () => { changed.value += 1 } }),
    bindings = new Map<string, Binding>()
  let previousOwner: SlicerCanvasDataOwner | undefined
  const candidates = computed<SlicerModel[]>(() => {
    void sourceChanged.value; void context.view()
    const studio = context.studio(), unitId = context.unitId(), sheetId = context.sheetId()
    if (!studio || !unitId || studio.isDisposed || !context.active()) return []
    const sheet = FOffice.newAPI(studio).getWorkbook(unitId)?.getSheetBySheetId(sheetId)
    return (studio.getWorkbookModel(unitId)?.slicers ?? []).filter(model => model.hostSheetId === sheetId).flatMap(model => {
      try { const hosted = sheet?.getSlicerControl(model.id)?.getModel(); return hosted ? [hosted] : [] }
      catch (error) { if (error instanceof SlicerConnectionError) return []; throw error }
    })
  })
  watch(() => context.studio(), (studio, _old, onCleanup) => {
    if (!studio) return
    const update = (unitId?: string) => { if (!unitId || unitId === context.unitId()) sourceChanged.value += 1 }
    const listeners = [studio.events.addEvent('CommandExecuted', event => update(event.unitId)),
      studio.events.addEvent('Undo', event => update(event.unitId)), studio.events.addEvent('Redo', event => update(event.unitId)),
      studio.events.addEvent('SlicerAdded', event => update(event.unitId)), studio.events.addEvent('SlicerChanged', event => update(event.unitId)),
      studio.events.addEvent('SlicerRemoved', event => update(event.unitId)), studio.events.addEvent('PivotTableChanged', event => update(event.unitId)),
      studio.events.addEvent('PivotChartChanged', event => update(event.unitId)),
      studio.events.addEvent('WorkbookDisposed', event => update(event.workbookId)), studio.events.addEvent('LifeCycleChanged', () => update())]
    onCleanup(() => listeners.forEach(listener => listener.dispose()))
  }, { immediate: true, flush: 'sync' })

  function owner(): SlicerCanvasDataOwner | undefined {
    const studio = context.studio(), unitId = context.unitId(), sheetId = context.sheetId()
    if (!studio || !unitId || studio.isDisposed || !context.active()) return undefined
    const workbook = studio.getWorkbookModel(unitId)
    if (!workbook || !workbook.sheets.some(sheet => sheet.id === sheetId)) return undefined
    const revision = studio.getWorkbookRevision(unitId), userId = studio.currentUserId, locale = studio.regionalSettings.locale.value,
      historyGeneration = studio.editor.history.captureSavepoint(unitId).generation
    return { instance: studio, model: workbook, unitId, sheetId, revision, historyGeneration, userId, locale, assertCurrent: () => {
      if (context.studio() !== studio || context.unitId() !== unitId || context.sheetId() !== sheetId || !context.active() || studio.isDisposed
        || studio.currentUserId !== userId || studio.getWorkbookModel(unitId) !== workbook || studio.getWorkbookRevision(unitId) !== revision
        || studio.editor.history.captureSavepoint(unitId).generation !== historyGeneration
        || studio.regionalSettings.locale.value !== locale || !workbook.sheets.some(sheet => sheet.id === sheetId)) {
        throw new Error('切片器画布来源、权限用户或历史代次已改变。')
      }
    } }
  }
  function targetWindow(model: SlicerModel, sheet: SpreadsheetRenderSheet, value?: RequestedWindow) {
    const geometry = slicerDrawingGeometry(model, new SpreadsheetGeometry(sheet)), state = cache.state(model.id),
      viewport = slicerMemberViewport(model, geometry.width, geometry.height, state.total, value?.start ?? model.startItem ?? 0)
    return { start: viewport.start, count: value?.count ?? viewport.count }
  }
  function request(id: string, start?: number, count?: number): void {
    const model = candidates.value.find(model => model.id === id), sheet = context.sheet()
    if (!model || !sheet) return
    if (model.type === 'timeline') { cache.request(id); return }
    const before = requested.value.get(id), value: RequestedWindow = {
      start: start ?? before?.start ?? model.startItem ?? 0, count: count ?? before?.count,
    }
    try { assertSlicerCanvasWindow(value.start, value.count ?? 1, Number.MAX_SAFE_INTEGER) }
    catch { cache.request(id, value.start, value.count ?? 1); return }
    if (start !== undefined || count !== undefined) {
      if (!before || before.start !== value.start || before.count !== value.count) requested.value = new Map(requested.value).set(id, value)
    }
    const target = targetWindow(model, sheet, value)
    cache.request(id, target.start, target.count)
  }
  watch(() => {
    void sourceChanged.value; void context.view()
    const current = owner()
    return { current, models: candidates.value, sheet: context.sheet() }
  }, ({ current, models }) => {
    if (!sameSlicerCanvasDataOwner(current, previousOwner)) { requested.value = new Map(); bindings.clear() }
    previousOwner = current
    cache.reset(current)
    const ids = new Set(models.map(model => model.id))
    cache.retain(ids)
    for (const id of [...bindings.keys()]) if (!ids.has(id)) bindings.delete(id)
    if ([...requested.value.keys()].some(id => !ids.has(id))) requested.value = new Map([...requested.value].filter(([id]) => ids.has(id)))
    if (!current) return
    const studio = context.studio()!, sheet = FOffice.newAPI(studio).getWorkbook(current.unitId)?.getSheetBySheetId(current.sheetId)
    for (const model of models) {
      const identity = sourceIdentity(studio, current.unitId, model), before = bindings.get(model.id)
      if (!before || before.identity !== identity || !sameSlicerCanvasDataOwner(before.owner, current)) {
        const control = sheet?.getSlicerControl(model.id)
        if (!control) continue
        const loader: SlicerCanvasDataLoader = { members: (start, count, signal) => control.getDisplayMemberWindow({ start, count, signal }),
          metadata: signal => control.getCanvasSourceMetadata(signal), assertCurrent: () => {
            current.assertCurrent()
            const workbook = studio.getWorkbookModel(current.unitId)!, hosted = workbook.slicers?.find(candidate => candidate.id === model.id)
            if (!hosted || hosted.hostSheetId !== current.sheetId || sourceIdentity(studio, current.unitId, hosted) !== identity) {
              throw new Error('切片器绘图宿主或绑定来源已改变。')
            }
            assertSlicerSourceReadPermission(workbook, hosted, current.userId)
          } }
        bindings.set(model.id, { owner: current, identity, loader })
        cache.bind(model.id, model.type, loader)
      }
      request(model.id)
    }
  }, { immediate: true, flush: 'sync' })

  // A first result can clamp a persisted/preview start against the real total.
  // Stable bindings and idempotent request avoid replacing the loader or causing an update loop.
  watch(changed, () => { for (const model of candidates.value) request(model.id) }, { flush: 'post' })
  const controls = computed<SlicerCanvasControl[]>(() => {
    void changed.value
    const sheet = context.sheet()
    if (!sheet) return []
    return candidates.value.map(model => {
      const state = cache.state(model.id)
      if (model.type === 'timeline') return { model, items: [], timeline: state.metadata?.timeline ?? { interactive: false }, dataStatus: state }
      const target = targetWindow(model, sheet, requested.value.get(model.id)), window = state.error ? undefined : cache.read(model.id, target.start, target.count)
      return { model, items: [], memberWindow: window ?? { start: target.start, total: state.total ?? 0, items: [] }, dataStatus: state }
    })
  })
  function retry(id: string): void { cache.retry(id) }
  onScopeDispose(() => { cache.dispose(); bindings.clear(); requested.value = new Map() })
  return { controls, request, retry }
}
