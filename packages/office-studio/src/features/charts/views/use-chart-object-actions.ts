import { computed, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { Disposable } from '../../../api/events'
import type { OfficeStudio } from '../../../api/office-studio'
import { randomId } from '../../../api/model-utils'
import { sheetSystemClipboard, sheetSystemClipboardAvailable, type SheetSystemClipboard, type SheetSystemClipboardContent } from '../../../components/spreadsheet/sheet-system-clipboard'
import { CHART_OBJECT_COMMANDS, PASTE_COPIED_CHART_COMMAND, assertChartObjectAction, chartObjectCapabilities, isPrimaryChartObject, registerChartObjectUIHost,
  type ChartObjectAction, type ChartObjectCapabilities, type ChartObjectTarget } from '../chart-object-commands'
import { chartObjectClipboardIdFromContent, chartObjectClipboardIdFromHtml, createChartObjectId } from '../chart-object-clipboard'
import { chartObjectClipboardStore } from '../chart-object-clipboard-store'
import { exportChartCanvasImage } from '../chart-image-export'
import type { ChartEditPanelRoute } from '../chart-edit-panel-route'

export interface ChartObjectChange {
  unitId: string
  sheetId: string
  chartId: string
  action: 'paste' | 'delete'
}
interface ChartObjectActionsHost {
  studio(): OfficeStudio | undefined
  unitId(): string | undefined
  sheetId(): string
  selectedId(): string | null
  active(): boolean
  panelAvailable(): boolean
  authorBusy?(): boolean
  canExport(chartId: string): boolean
  prepareExport(target: ChartObjectTarget): HTMLCanvasElement | undefined
  openPanel(target: ChartObjectTarget, route?: ChartEditPanelRoute): boolean
  changed(change: ChartObjectChange): void
  focus(chartId?: string): void
  notify(message: string): void
  clipboard?: SheetSystemClipboard
}

/** The command owns content; this adapter owns browser I/O and the originating UI session. */
export function useChartObjectActions(host: ChartObjectActionsHost) {
  const busy = ref<ChartObjectAction | 'paste'>(), busyChartId = ref<string>()
  const sessionId = ref(''), error = ref('')
  const clipboardIO = host.clipboard ?? sheetSystemClipboard
  let registration: Disposable | undefined, boundStudio: OfficeStudio | undefined, disposed = false, generation = 0
  let lifecycle: Disposable[] = []
  let exportAbort: AbortController | undefined
  const operation = shallowRef<object>()

  function currentSheet(studio: OfficeStudio | undefined, unitId: string, sheetId: string): boolean {
    return !disposed && host.active() && !!studio && studio === boundStudio && host.studio() === studio
      && host.unitId() === unitId && host.sheetId() === sheetId && studio.activeWorkbookId === unitId
      && studio.editor.state.activeKind === 'workbook' && studio.getWorkbookModel(unitId)?.activeSheetId === sheetId
  }
  function capture(chartId = host.selectedId() ?? ''): ChartObjectTarget | undefined {
    const studio = host.studio(), unitId = host.unitId(), sheetId = host.sheetId()
    if (!studio || !unitId || !sessionId.value || !currentSheet(studio, unitId, sheetId)) return undefined
    const workbook = studio.getWorkbookModel(unitId)
    const chart = workbook?.charts?.find(chart => chart.sheetId === sheetId
      && (chartId ? chart.id === chartId : isPrimaryChartObject(workbook, chart)))
    if (!chart) return undefined
    return { unitId, sheetId, chartId: chart.id, expectedRevision: studio.getWorkbookRevision(unitId), snapshot: JSON.stringify(chart), hostSessionId: sessionId.value }
  }
  function isCurrent(target: ChartObjectTarget): boolean {
    const studio = host.studio()
    const chart = studio?.getWorkbookModel(target.unitId)?.charts?.find(chart => chart.id === target.chartId && chart.sheetId === target.sheetId)
    return currentSheet(studio, target.unitId, target.sheetId) && (!target.hostSessionId || target.hostSessionId === sessionId.value)
      && studio!.getWorkbookRevision(target.unitId) === target.expectedRevision && !!chart && JSON.stringify(chart) === target.snapshot
  }
  function capabilities(target?: ChartObjectTarget): ChartObjectCapabilities {
    const studio = host.studio(), workbook = target && studio?.getWorkbookModel(target.unitId)
    const chart = workbook?.charts?.find(chart => chart.id === target?.chartId && chart.sheetId === target.sheetId)
    if (!target || !studio || !workbook || !chart || !isCurrent(target)) {
      const disabled = { disabled: true, reason: '图表上下文已更新，请重新选择。' }
      return { edit: disabled, copy: disabled, exportImage: disabled, delete: disabled }
    }
    return chartObjectCapabilities(workbook, chart, studio.currentUserId,
      { panel: host.panelAvailable(), image: host.canExport(chart.id), busy: !!busy.value || !!host.authorBusy?.() })
  }
  const currentCapabilities = computed(() => capabilities(capture()))
  function cancel(): void {
    ++generation; exportAbort?.abort(); exportAbort = undefined; operation.value = undefined
    busy.value = undefined; busyChartId.value = undefined; error.value = ''
  }
  function report(reason: unknown, ticket: number): void {
    if (ticket !== generation || disposed || !host.active()) return
    if (reason instanceof Error && reason.name === 'AbortError') return
    error.value = reason instanceof Error ? reason.message : String(reason)
    host.notify(error.value)
  }
  async function execute(action: ChartObjectAction, target = capture(), route?: ChartEditPanelRoute): Promise<boolean> {
    if (!target || busy.value || host.authorBusy?.() || !isCurrent(target)) return false
    const studio = host.studio()!, token = {}, ticket = generation
    operation.value = token; busy.value = action; busyChartId.value = target.chartId; error.value = ''
    try {
      const input = action === 'edit' && route ? { ...target, route } : target
      const accepted = await studio.executeCommand<typeof input, boolean>(CHART_OBJECT_COMMANDS[action], input)
      if (!accepted) return false
      if (ticket === generation && currentSheet(studio, target.unitId, target.sheetId) && action === 'delete') {
        host.changed({ ...target, action })
      }
      return true
    } catch (reason) { report(reason, ticket); return false }
    finally { if (operation.value === token) { operation.value = undefined; busy.value = undefined; busyChartId.value = undefined } }
  }
  async function paste(content?: SheetSystemClipboardContent): Promise<boolean> {
    const studio = host.studio(), unitId = host.unitId(), sheetId = host.sheetId()
    if (!studio || !unitId || busy.value || host.authorBusy?.() || !currentSheet(studio, unitId, sheetId)) return false
    const store = chartObjectClipboardStore(studio.pinia), clipboard = store.clipboard
    if (!clipboard) return false
    const target = { unitId, sheetId, expectedRevision: studio.getWorkbookRevision(unitId), clipboardId: clipboard.id,
      newChartId: createChartObjectId(), hostSessionId: sessionId.value }
    const ticket = generation, token = {}
    operation.value = token; busy.value = 'paste'; busyChartId.value = host.selectedId() ?? undefined; error.value = ''
    try {
      let payload = content
      if (!payload && sheetSystemClipboardAvailable(clipboardIO)) {
        try { payload = await clipboardIO.read() }
        catch (reason) { if (store.systemWritten) throw reason }
      }
      if (payload && chartObjectClipboardIdFromContent(payload) !== clipboard.id) throw new Error('系统剪贴板已更改，请重新复制图表。')
      if (!payload && store.systemWritten) throw new Error('无法确认系统剪贴板中的图表，请重新复制后粘贴。')
      if (ticket !== generation || target.hostSessionId !== sessionId.value || !currentSheet(studio, unitId, sheetId)) return false
      const changed = await studio.executeCommand<typeof target, boolean>(PASTE_COPIED_CHART_COMMAND, target)
      if (changed && ticket === generation && currentSheet(studio, unitId, sheetId)) host.changed({ unitId, sheetId, chartId: target.newChartId, action: 'paste' })
      return changed
    } catch (reason) { report(reason, ticket); return false }
    finally { if (operation.value === token) { operation.value = undefined; busy.value = undefined; busyChartId.value = undefined } }
  }
  function handlePasteEvent(event: ClipboardEvent): boolean {
    const studio = host.studio(), unitId = host.unitId()
    if (!studio || !unitId || !event.clipboardData || !currentSheet(studio, unitId, host.sheetId())) return false
    const content = { text: event.clipboardData.getData('text/plain'), html: event.clipboardData.getData('text/html') }
    const id = chartObjectClipboardIdFromContent(content)
    if (!id && !chartObjectClipboardIdFromHtml(content.html) && !content.text.includes('[office-chart:')) return false
    event.preventDefault(); event.stopImmediatePropagation()
    const clipboard = chartObjectClipboardStore(studio.pinia).clipboard
    if (!clipboard || id !== clipboard.id) {
      error.value = '图表剪贴板来源已失效，请重新复制图表。'; host.notify(error.value); return true
    }
    void paste(content); return true
  }
  function restoreFocus(target?: ChartObjectTarget): void {
    if (disposed || !host.active()) return
    if (target && (!currentSheet(host.studio(), target.unitId, target.sheetId)
      || target.hostSessionId && target.hostSessionId !== sessionId.value)) return
    host.focus(target?.chartId)
  }

  watch([host.studio, host.unitId, host.sheetId, host.active], () => {
    for (const subscription of lifecycle) subscription.dispose()
    lifecycle = []
    cancel(); registration?.dispose(); registration = undefined; boundStudio = undefined; sessionId.value = ''
    const studio = host.studio(), unitId = host.unitId()
    if (!studio || !unitId || !host.active() || disposed) return
    boundStudio = studio; sessionId.value = randomId('chart-ui')
    const ownedSession = sessionId.value
    registration = registerChartObjectUIHost(studio, {
      sessionId: ownedSession,
      target: () => capture(), current: isCurrent,
      currentSheet: (id, sheet) => ownedSession === sessionId.value && currentSheet(studio, id, sheet),
      openPanel: (target, route) => isCurrent(target) && host.openPanel(target, route),
      async writeClipboard(content) {
        try { return await clipboardIO.write(content) }
        catch { return false }
      },
      async exportImage(target) {
        exportAbort?.abort()
        const abort = new AbortController(); exportAbort = abort
        const assertCurrent = () => {
          if (abort.signal.aborted || ownedSession !== sessionId.value || !isCurrent(target)) throw new DOMException('图表导出上下文已更新。', 'AbortError')
          assertChartObjectAction(studio, target, 'exportImage')
        }
        try {
          assertCurrent()
          const surface = host.prepareExport(target)
          if (!surface) throw new Error('图表尚未完成绘制，无法导出图片。')
          await exportChartCanvasImage(surface, target.chartId, { signal: abort.signal, assertCurrent })
          return true
        } finally { if (exportAbort === abort) exportAbort = undefined }
      },
    })
    lifecycle = [studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === unitId) cancel() }),
      studio.events.addEvent('LifeCycleChanged', event => { if (event.stage === 'disposed') { cancel(); registration?.dispose(); boundStudio = undefined; sessionId.value = '' } })]
  }, { immediate: true, flush: 'sync' })
  onScopeDispose(() => { disposed = true; cancel(); registration?.dispose(); registration = undefined; boundStudio = undefined; for (const subscription of lifecycle) subscription.dispose(); lifecycle = [] })
  return { busy, busyChartId, error, sessionId, currentCapabilities, capture, isCurrent, capabilities, execute, paste, handlePasteEvent, restoreFocus, cancel }
}
