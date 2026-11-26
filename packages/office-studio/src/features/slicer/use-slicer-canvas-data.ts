import { computed, onScopeDispose, ref, watch } from 'vue'
import { FOffice } from '../../api/facade.ts'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { SlicerModel } from '../../domain/model.ts'
import { SpreadsheetGeometry, type SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout.ts'
import { slicerDrawingGeometry } from './slicer-drawing.ts'
import { SlicerCanvasDataCache, type SlicerCanvasDataOwner } from './slicer-canvas-data.ts'
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

/** Vue only schedules and observes windows. Source identities, filtering and member computation stay in the shared service/Worker. */
export function useSlicerCanvasData(context: SlicerCanvasDataContext) {
  const changed = ref(0), sourceChanged = ref(0), requested = ref(new Map<string, { start: number; count: number }>()),
    cache = new SlicerCanvasDataCache({ onChange: () => { changed.value += 1 } })
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
      studio.events.addEvent('WorkbookDisposed', event => update(event.workbookId)), studio.events.addEvent('LifeCycleChanged', () => update())]
    onCleanup(() => listeners.forEach(listener => listener.dispose()))
  }, { immediate: true, flush: 'sync' })

  function owner(): SlicerCanvasDataOwner | undefined {
    const studio = context.studio(), unitId = context.unitId(), sheetId = context.sheetId()
    if (!studio || !unitId || studio.isDisposed || !context.active() || !studio.getWorkbookModel(unitId)) return undefined
    const workbook = studio.getWorkbookModel(unitId)!
    const revision = studio.getWorkbookRevision(unitId), userId = studio.currentUserId, locale = studio.regionalSettings.locale.value
    return { instance: studio, model: workbook, unitId, sheetId, revision, userId, locale, assertCurrent: () => {
      if (context.studio() !== studio || context.unitId() !== unitId || context.sheetId() !== sheetId || !context.active() || studio.isDisposed
        || studio.currentUserId !== userId || studio.getWorkbookModel(unitId) !== workbook || studio.getWorkbookRevision(unitId) !== revision
        || studio.regionalSettings.locale.value !== locale) throw new Error('切片器画布来源已改变。')
    } }
  }
  function request(id: string, start?: number, count?: number): void {
    const model = candidates.value.find(model => model.id === id), sheet = context.sheet()
    if (!model || !sheet) return
    if (model.type === 'timeline') { cache.request(id); return }
    const geometry = slicerDrawingGeometry(model, new SpreadsheetGeometry(sheet)), state = cache.state(id),
      viewport = slicerMemberViewport(model, geometry.width, geometry.height, state.total)
    const target = { start: start ?? viewport.start, count: count ?? viewport.count }, before = requested.value.get(id)
    if (!before || before.start !== target.start || before.count !== target.count) requested.value = new Map(requested.value).set(id, target)
    cache.request(id, target.start, target.count)
  }
  watch(() => {
    void sourceChanged.value; void context.view()
    const current = owner()
    return { current, models: candidates.value, sheet: context.sheet() }
  }, ({ current, models }) => {
    if (!current || !previousOwner || current.instance !== previousOwner.instance || current.model !== previousOwner.model || current.unitId !== previousOwner.unitId || current.sheetId !== previousOwner.sheetId
      || current.revision !== previousOwner.revision || current.userId !== previousOwner.userId || current.locale !== previousOwner.locale) requested.value = new Map()
    previousOwner = current
    cache.reset(current)
    cache.retain(new Set(models.map(model => model.id)))
    if (!current) return
    const studio = context.studio()!, sheet = FOffice.newAPI(studio).getWorkbook(current.unitId)?.getSheetBySheetId(current.sheetId)
    for (const model of models) {
      const control = sheet?.getSlicerControl(model.id)
      if (!control) continue
      cache.bind(model.id, model.type, { members: (start, count, signal) => control.getDisplayMemberWindow({ start, count, signal }),
        metadata: signal => control.getCanvasSourceMetadata(signal), assertCurrent: () => {
          current.assertCurrent()
          const workbook = studio.getWorkbookModel(current.unitId)!, hosted = workbook.slicers?.find(candidate => candidate.id === model.id)
          if (!hosted || hosted.hostSheetId !== current.sheetId) throw new Error('切片器绘图宿主已改变。')
          assertSlicerSourceReadPermission(workbook, hosted, current.userId)
        } })
      request(model.id)
    }
  }, { immediate: true, flush: 'post' })

  const controls = computed<SlicerCanvasControl[]>(() => {
    void changed.value
    const sheet = context.sheet()
    if (!sheet) return []
    const grid = new SpreadsheetGeometry(sheet)
    return candidates.value.map(model => {
      const state = cache.state(model.id)
      if (model.type === 'timeline') return { model, items: [], timeline: state.metadata?.timeline ?? { interactive: false }, dataStatus: state }
      const geometry = slicerDrawingGeometry(model, grid), viewport = slicerMemberViewport(model, geometry.width, geometry.height, state.total),
        target = requested.value.get(model.id) ?? viewport, window = cache.read(model.id, target.start, target.count)
      return { model, items: [], memberWindow: window ?? { start: target.start, total: state.total ?? 0, items: [] }, dataStatus: state }
    })
  })
  function retry(id: string): void { cache.retry(id); request(id) }
  onScopeDispose(() => cache.dispose())
  return { controls, request, retry }
}
