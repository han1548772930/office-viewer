import { computed, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { FOffice, type FWorksheet } from '../../../api/facade'
import type { OfficeStudio } from '../../../api/office-studio'
import type { ChartModel } from '../../../domain/model'
import type { ChartConfig, ChartData } from '../chart-engine'
import type { ChartComputedStatistics } from '../chart-derived'
import type { ChartSettingsSourceSheet } from '../chart-settings-draft'
import type { ChartSettingsSourceLoader, ChartSettingsStatisticsLoader } from '../chart-settings-data'
import { ComputeError } from '../../../compute/errors.ts'

type ChartPatch = Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>

interface SheetChartPanelHost {
  studio: () => OfficeStudio | undefined
  unitId: () => string | undefined
  sheetId: () => string | undefined
  sheets: () => ChartSettingsSourceSheet[]
  active: () => boolean
  canRestoreFocus: () => boolean
  refresh: (studio: OfficeStudio, unitId: string) => void
  focus: () => void
  notify: (message: string) => void
}

interface ChartPanelTarget {
  key: number
  studio: OfficeStudio
  unitId: string
  sheetId: string
  sheets: ChartSettingsSourceSheet[]
  setOpen: (open: boolean) => void
  closed: () => void
}
interface ChartPanelReaders { loadSource: ChartSettingsSourceLoader; loadStatistics: ChartSettingsStatisticsLoader }
interface ChartInsertTarget extends ChartPanelTarget, ChartPanelReaders { config: ChartConfig; apply: (config: ChartConfig) => Promise<void> }
interface ChartEditorTarget extends ChartPanelTarget, ChartPanelReaders { chartId: string; chart: ChartModel; apply: (patch: ChartPatch) => Promise<void> }
interface PivotChartTarget extends ChartPanelTarget { applied: () => void; setBusy: (busy: boolean) => void }

/** Panel callbacks retain their originating unit, even after the workbench switches. */
export function useSheetChartPanels(host: SheetChartPanelHost) {
  const chartModalOpen = ref(false), chartEditorOpen = ref(false), pivotChartPanelOpen = ref(false)
  const chartInsertBusy = ref(false), chartEditorBusy = ref(false), pivotChartBusy = ref(false)
  const chartError = ref(''), editorError = ref('')
  const chartInsertContext = shallowRef<ChartInsertTarget>()
  const chartEditorContext = shallowRef<ChartEditorTarget>()
  const pivotChartContext = shallowRef<PivotChartTarget>()
  const sourceRevision = ref(0)
  let generation = 0, disposed = false

  const current = (target: ChartPanelTarget): boolean => !disposed && host.active()
    && host.studio() === target.studio && host.unitId() === target.unitId && host.sheetId() === target.sheetId
  const anyOpen = (): boolean => chartModalOpen.value || chartEditorOpen.value || pivotChartPanelOpen.value
  function restoreFocus(target: ChartPanelTarget): void {
    if (current(target) && !anyOpen() && host.canRestoreFocus()) host.focus()
  }
  function target(): Omit<ChartPanelTarget, 'setOpen' | 'closed'> | undefined {
    const studio = host.studio(), unitId = host.unitId(), sheetId = host.sheetId()
    if (!studio || !unitId || !sheetId || !host.active()) {
      host.notify('图表命令上下文尚未绑定到活动工作簿。')
      return undefined
    }
    return { key: ++generation, studio, unitId, sheetId, sheets: [...host.sheets()] }
  }
  function closeChartPanels(): void {
    chartModalOpen.value = false
    chartEditorOpen.value = false
    pivotChartPanelOpen.value = false
  }
  function releaseChartPanelSessions(): void {
    closeChartPanels()
    chartInsertContext.value = undefined
    chartEditorContext.value = undefined
    pivotChartContext.value = undefined
    chartInsertBusy.value = false
    chartEditorBusy.value = false
    pivotChartBusy.value = false
    chartError.value = ''
    editorError.value = ''
  }
  function refresh(target: ChartPanelTarget): void {
    if (!current(target)) return
    sourceRevision.value += 1
    host.refresh(target.studio, target.unitId)
  }
  function canOpen(): boolean { return !chartInsertBusy.value && !chartEditorBusy.value && !pivotChartBusy.value && !disposed }

  async function readSettings<T>(target: ChartPanelTarget, options: { signal: AbortSignal }, active: () => boolean, read: (sheet: FWorksheet) => Promise<T>): Promise<T> {
    const assertCurrent = (): void => {
      if (options.signal.aborted) throw new ComputeError('COMPUTE_CANCELLED', '图表设置数据读取已取消')
      if (!current(target) || !active()) throw new ComputeError('COMPUTE_STALE_RESULT', '图表设置面板上下文已更新')
    }
    assertCurrent()
    const sheet = FOffice.newAPI(target.studio).getWorkbook(target.unitId)?.getSheetBySheetId(target.sheetId)
    if (!sheet) throw new Error('图表所属的工作簿或工作表已经关闭。')
    const data = await read(sheet)
    assertCurrent()
    return data
  }

  function readSource(target: ChartPanelTarget, config: ChartConfig, options: { signal: AbortSignal }, active: () => boolean): Promise<ChartData> {
    return readSettings(target, options, active, sheet => sheet.getChartSettingsData(config, options))
  }
  function readStatistics(target: ChartPanelTarget, config: ChartConfig, options: { signal: AbortSignal }, active: () => boolean): Promise<ChartComputedStatistics> {
    return readSettings(target, options, active, sheet => sheet.getChartSettingsStatistics(config, options))
  }

  function openInsertChart(config: ChartConfig): void {
    if (!canOpen()) return
    const captured = target()
    if (!captured) return
    closeChartPanels()
    const session: ChartInsertTarget = { ...captured, config,
      loadSource: (config, options) => readSource(session, config, options, () => chartInsertContext.value === session && chartModalOpen.value),
      loadStatistics: (config, options) => readStatistics(session, config, options, () => chartInsertContext.value === session && chartModalOpen.value),
      setOpen(open) {
        if (chartInsertContext.value === session && (open ? current(session) : !chartInsertBusy.value)) chartModalOpen.value = open
      },
      closed() {
        if (chartInsertContext.value !== session || chartModalOpen.value) return
        chartInsertContext.value = undefined
        chartInsertBusy.value = false
        chartError.value = ''
        restoreFocus(session)
      },
      async apply(nextConfig) {
        if (chartInsertContext.value !== session || !chartModalOpen.value || chartInsertBusy.value || !current(session)) return
        chartInsertBusy.value = true
        chartError.value = ''
        try {
          const sheet = FOffice.newAPI(session.studio).getWorkbook(session.unitId)?.getSheetBySheetId(session.sheetId)
          if (!sheet) throw new Error('图表所属的工作簿或工作表已经关闭。')
          await sheet.insertChart(nextConfig)
          if (chartInsertContext.value !== session || !current(session)) return
          refresh(session)
          if (chartInsertContext.value === session) chartModalOpen.value = false
        } catch (reason) {
          if (chartInsertContext.value === session && current(session)) chartError.value = reason instanceof Error ? reason.message : '图表数据范围不可用。'
        } finally {
          if (chartInsertContext.value === session) chartInsertBusy.value = false
        }
      },
    }
    chartInsertContext.value = session
    chartError.value = ''
    chartModalOpen.value = true
  }

  function openChartEditor(chart: ChartModel): void {
    if (!canOpen()) return
    const captured = target()
    if (!captured || chart.sheetId !== captured.sheetId) return
    closeChartPanels()
    const session: ChartEditorTarget = { ...captured, chartId: chart.id, chart,
      loadSource: (config, options) => readSource(session, config, options, () => chartEditorContext.value === session && chartEditorOpen.value),
      loadStatistics: (config, options) => readStatistics(session, config, options, () => chartEditorContext.value === session && chartEditorOpen.value),
      setOpen(open) {
        if (chartEditorContext.value === session && (open ? current(session) : !chartEditorBusy.value)) chartEditorOpen.value = open
      },
      closed() {
        if (chartEditorContext.value !== session || chartEditorOpen.value) return
        chartEditorContext.value = undefined
        chartEditorBusy.value = false
        editorError.value = ''
        restoreFocus(session)
      },
      async apply(patch) {
        if (chartEditorContext.value !== session || !chartEditorOpen.value || chartEditorBusy.value || !current(session)) return
        chartEditorBusy.value = true
        editorError.value = ''
        try {
          const chart = FOffice.newAPI(session.studio).getWorkbook(session.unitId)?.getSheetBySheetId(session.sheetId)?.getChartById(session.chartId)
          if (!chart) throw new Error('图表所属的工作簿或图表已经关闭。')
          await chart.update(patch)
          if (chartEditorContext.value === session && current(session)) refresh(session)
        } catch (reason) {
          if (chartEditorContext.value === session && current(session)) editorError.value = reason instanceof Error ? reason.message : '图表设置未能应用。'
        } finally {
          if (chartEditorContext.value === session) chartEditorBusy.value = false
        }
      },
    }
    chartEditorContext.value = session
    editorError.value = ''
    chartEditorOpen.value = true
  }

  function openPivotChart(): void {
    if (!canOpen()) return
    const captured = target()
    if (!captured) return
    closeChartPanels()
    const session: PivotChartTarget = { ...captured,
      setOpen(open) { if (pivotChartContext.value === session && (!open || current(session))) pivotChartPanelOpen.value = open },
      closed() {
        if (pivotChartContext.value !== session || pivotChartPanelOpen.value) return
        pivotChartContext.value = undefined
        pivotChartBusy.value = false
        restoreFocus(session)
      },
      setBusy(busy) { if (pivotChartContext.value === session) pivotChartBusy.value = busy },
      applied() { if (pivotChartContext.value === session && current(session)) refresh(session) },
    }
    pivotChartContext.value = session
    pivotChartPanelOpen.value = true
  }

  const chartEditorPreview = computed<{ chart: ChartModel | null; error?: string }>(() => {
    sourceRevision.value
    const session = chartEditorContext.value
    if (!session) return { chart: null }
    if (!chartEditorOpen.value || !current(session)) return { chart: session.chart }
    try {
      const chart = FOffice.newAPI(session.studio).getWorkbook(session.unitId)?.getSheetBySheetId(session.sheetId)?.getChartById(session.chartId)
      if (!chart) throw new Error('图表所属的工作簿或图表已经关闭。')
      return { chart: chart.getInfo() }
    } catch (reason) { return { chart: null, error: reason instanceof Error ? reason.message : String(reason) } }
  })
  const chartEditorError = computed(() => editorError.value || chartEditorPreview.value.error || '')
  const chartEditorChart = computed(() => chartEditorPreview.value.chart)
  const chartSourceRevision = computed(() => sourceRevision.value)
  const chartEditorSheets = computed(() => {
    const session = chartEditorContext.value
    return session ? current(session) ? host.sheets() : session.sheets : []
  })
  const chartEditorSheet = computed(() => chartEditorSheets.value.find(sheet => sheet.id === chartEditorContext.value?.sheetId))
  const chartInsertSheets = computed(() => {
    const session = chartInsertContext.value
    return session ? current(session) ? host.sheets() : session.sheets : []
  })

  watch([host.studio, host.unitId, host.sheetId, host.active], releaseChartPanelSessions, { flush: 'sync' })
  watch(() => host.studio(), (studio, _old, onCleanup) => {
    if (!studio) return
    const changed = (unitId: string): void => { if (unitId === host.unitId()) sourceRevision.value += 1 }
    const listeners = [
      studio.events.addEvent('CommandExecuted', event => { if (event.unitId) changed(event.unitId) }),
      studio.events.addEvent('Undo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
      studio.events.addEvent('Redo', event => { if (event.kind === 'workbook') changed(event.unitId) }),
      studio.events.addEvent('SheetValueChanged', event => changed(event.workbookId)),
      studio.events.addEvent('SheetStructureChanged', event => changed(event.workbookId)),
      studio.events.addEvent('PivotTableChanged', event => changed(event.unitId)),
      studio.events.addEvent('PivotTableRendered', event => changed(event.unitId)),
      studio.events.addEvent('WorkbookDisposed', event => { if (event.workbookId === host.unitId()) releaseChartPanelSessions() }),
    ]
    onCleanup(() => listeners.forEach(listener => listener.dispose()))
  }, { immediate: true })
  onScopeDispose(() => { disposed = true; releaseChartPanelSessions() })

  return { chartModalOpen, chartEditorOpen, pivotChartPanelOpen, chartInsertBusy, chartEditorBusy, chartError, chartEditorError,
    chartInsertContext, chartEditorContext, pivotChartContext, chartEditorChart, chartSourceRevision, chartEditorSheet, chartEditorSheets, chartInsertSheets,
    openInsertChart, openChartEditor, openPivotChart }
}
