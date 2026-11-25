import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../../api/office-studio'
import { FOffice } from '../../../api/facade'
import type { ChartModel } from '../../../domain/model'
import type { ChartConfig } from '../chart-engine'
import type { ChartEditPanelRoute } from '../chart-edit-panel-route'
import { chartCanvasElementCompatible, chartCanvasElementDeleteDraft, chartCanvasElementHidden, chartCanvasElementPatch } from '../chart-element-edit'
import { chartCanvasElementAsSeries, chartCanvasElementKey, chartCanvasElementRoute, chartCanvasSeriesMark, supportsChartSeriesStyleMark, type ChartCanvasElement,
  type ChartCanvasElementTarget } from '../chart-element-interaction'
import { chartTitleAuthorContextMatches, type ChartTitleAuthorContext } from '../chart-title-interaction'

export interface ChartCanvasElementSelection {
  chartId: string
  element: ChartCanvasElement
  key: string
  workbookId?: string
  sheetId: string
  studio?: OfficeStudio
  snapshot: string
  context?: ChartTitleAuthorContext
  editable: boolean
  generation: number
}
interface ChartElementEditingHost {
  studio(): OfficeStudio | undefined
  workbookId(): string | undefined
  sheetId(): string
  active(): boolean
  chart(chartId: string): ChartModel | undefined
  target(chartId: string, element: ChartCanvasElement): ChartCanvasElementTarget | undefined
  canEdit(chart: ChartModel): boolean
  busy(chartId: string): boolean
  changed(): void
  focus(chartId: string): void
  openPanel(chartId: string, route: ChartEditPanelRoute): Promise<boolean>
}

/** UI lifetime is separate from the canonical author command. A retired menu
 * cannot submit, rebind to another series, or restore focus to another sheet. */
export function useChartElementEditing(host: ChartElementEditingHost) {
  const selected = shallowRef<ChartCanvasElementSelection | null>(null)
  const commit = shallowRef<ChartCanvasElementSelection | null>(null)
  const error = ref('')
  let generation = 0, disposed = false
  const busy = computed(() => !!commit.value)

  function context(chartId: string): ChartTitleAuthorContext | undefined {
    const studio = host.studio(), workbookId = host.workbookId()
    const chart = studio && workbookId ? studio.getWorkbookModel(workbookId)?.charts?.find(chart => chart.id === chartId && chart.sheetId === host.sheetId()) : undefined
    return studio && workbookId && chart ? { owner: studio, workbookId, sheetId: host.sheetId(), chartId,
      revision: studio.getWorkbookRevision(workbookId), snapshot: JSON.stringify(chart) } : undefined
  }
  function sameLifetime(selection: ChartCanvasElementSelection): boolean {
    const studio = host.studio(), workbookId = host.workbookId()
    return !disposed && host.active() && selection.studio === studio && selection.workbookId === workbookId && selection.sheetId === host.sheetId()
      && (!studio || !!workbookId && studio.activeWorkbookId === workbookId && studio.editor.state.activeKind === 'workbook'
        && studio.getWorkbookModel(workbookId)?.activeSheetId === selection.sheetId)
  }
  function isCurrent(selection: ChartCanvasElementSelection): boolean {
    const chart = host.chart(selection.chartId)
    return sameLifetime(selection) && selection.generation === generation && !!chart
      && (selection.context ? chartTitleAuthorContextMatches(selection.context, context(selection.chartId)) : JSON.stringify(chart) === selection.snapshot)
  }
  function select(chartId: string, target: ChartCanvasElementTarget): ChartCanvasElementSelection | undefined {
    const chart = host.chart(chartId)
    if (!chart || disposed || !host.active() || commit.value) return undefined
    const previous = selected.value
    if (previous?.chartId === chartId && previous.key === target.key && isCurrent(previous)) return previous
    const selection: ChartCanvasElementSelection = { chartId, element: target.element, key: target.key, studio: host.studio(),
      workbookId: host.workbookId(), sheetId: host.sheetId(), snapshot: JSON.stringify(chart), context: context(chartId), editable: host.canEdit(chart), generation: ++generation }
    selected.value = selection; error.value = ''; host.changed()
    return selection
  }
  function close(restoreFocus = false): void {
    const previous = selected.value, ticket = ++generation
    selected.value = null; error.value = ''; host.changed()
    if (restoreFocus && previous) void nextTick(() => {
      if (ticket === generation && !selected.value && sameLifetime(previous) && host.chart(previous.chartId)) host.focus(previous.chartId)
    })
  }
  function targetMatches(selection: ChartCanvasElementSelection): boolean {
    const target = host.target(selection.chartId, selection.element)
    if (!target || target.key !== selection.key) return false
    // Legend controls address the whole legend, including its current page.
    return selection.element.kind === 'legend' || JSON.stringify(target.element) === JSON.stringify(selection.element)
  }
  async function apply(selection: ChartCanvasElementSelection | null, draft: ChartConfig): Promise<void> {
    const chart = selection && host.chart(selection.chartId), studio = host.studio()
    if (!selection || selected.value !== selection || !selection.context || !studio || !chart || !isCurrent(selection)
      || !host.canEdit(chart) || host.busy(selection.chartId) || commit.value || !targetMatches(selection))
      throw new Error('图表元素或权限已变更，请重新选择后编辑。')
    const patch = chartCanvasElementPatch(chart, draft, selection.element)
    if (!Object.keys(patch).length) return
    const live = FOffice.newAPI(studio).getWorkbook(selection.context.workbookId)?.getSheetBySheetId(selection.sheetId)?.getChartById(selection.chartId)
    if (!live) throw new Error('图表元素的所属工作表已关闭。')
    commit.value = selection; error.value = ''
    try {
      // FChart.update -> sheet.command.update-chart owns permission,
      // preparation, revision checking, one history entry and persistence.
      await live.update(patch)
      await nextTick()
      if (selected.value === selection && selection.generation === generation && sameLifetime(selection)) {
        const current = context(selection.chartId), chart = host.chart(selection.chartId)
        if (!current || !chart || chartCanvasElementHidden(chart, selection.element)) close(true)
        else selected.value = { ...selection, context: current, snapshot: current.snapshot }
        host.changed()
      }
    } catch (cause) {
      if (selected.value === selection && sameLifetime(selection)) error.value = cause instanceof Error ? cause.message : '图表元素修改失败。'
      throw cause
    } finally {
      if (commit.value === selection) commit.value = null
      reconcile()
    }
  }
  async function remove(): Promise<boolean> {
    const selection = selected.value, chart = selection && host.chart(selection.chartId)
    if (!selection || !chart || !isCurrent(selection) || commit.value || host.busy(selection.chartId) || !host.canEdit(chart)) return false
    const draft = chartCanvasElementDeleteDraft(chart, selection.element)
    // A point or series selection consumes Delete without deleting source data.
    if (!draft) return false
    try { await apply(selection, draft); return true }
    catch (cause) {
      if (sameLifetime(selection) && selected.value === selection) error.value = cause instanceof Error ? cause.message : '删除图表元素失败。'
      return false
    }
  }
  async function more(selection = selected.value): Promise<boolean> {
    const chart = selection && host.chart(selection.chartId)
    if (!selection || selected.value !== selection || !chart || !isCurrent(selection) || !host.canEdit(chart)
      || commit.value || host.busy(selection.chartId)) return false
    const opened = await host.openPanel(selection.chartId, chartCanvasElementRoute(selection.element, chart.type))
    if (opened && selected.value === selection && sameLifetime(selection)) close()
    return opened
  }
  function promote(selection = selected.value): ChartCanvasElementSelection | undefined {
    if (!selection || selected.value !== selection || !isCurrent(selection) || commit.value || host.busy(selection.chartId)) return undefined
    const series = chartCanvasElementAsSeries(selection.element), chart = host.chart(selection.chartId)
    const element = chart && series.kind === 'series' && supportsChartSeriesStyleMark(chart.type) ? chartCanvasSeriesMark(series.binding) : series
    const target = host.target(selection.chartId, element)
    return target && chartCanvasElementKey(element) !== selection.key ? select(selection.chartId, target) : undefined
  }
  function reconcile(): void {
    const selection = selected.value, chart = selection && host.chart(selection.chartId)
    if (selection && (!sameLifetime(selection) || !chart || !chartCanvasElementCompatible(chart, selection.element)
      || commit.value !== selection && (!isCurrent(selection)
      || !chart || host.canEdit(chart) !== selection.editable))) close()
  }
  const disabled = computed(() => {
    const selection = selected.value, chart = selection && host.chart(selection.chartId)
    return !selection || !chart || !selection.context || !isCurrent(selection) || !host.canEdit(chart) || !!commit.value || host.busy(selection.chartId)
  })
  watch([host.studio, host.workbookId, host.sheetId, host.active, () => {
    const selection = selected.value, studio = host.studio(), workbookId = host.workbookId(), chart = selection && host.chart(selection.chartId)
    return chart ? [chart, host.canEdit(chart), studio && workbookId ? studio.getWorkbookRevision(workbookId) : undefined,
      studio?.activeWorkbookId, studio?.editor.state.activeKind, studio && workbookId ? studio.getWorkbookModel(workbookId)?.activeSheetId : undefined] : undefined
  }], reconcile, { deep: true, flush: 'sync' })
  onScopeDispose(() => { disposed = true; close() })
  return { selected, commit, busy, error, disabled, select, close, isCurrent, apply, remove, more, promote, reconcile }
}
