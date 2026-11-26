import { computed, nextTick, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../../api/office-studio.ts'
import { clearPivotChartAxisSelection, getPivotChartAxisSelection, setPivotChartAxisSelection, type PivotChartAxisSelection } from '../../../api/pivot-chart-axis-selection.ts'
import { canEditPivotChart } from '../../../api/pivot-chart-commands.ts'
import type { ChartModel } from '../../../domain/model.ts'
import type { PivotChartCanvasAxisTarget } from '../../charts/echarts-pivot-axis-interaction.ts'
import { assertChartReadPermission } from '../../charts/services/chart-statistics-service.ts'
import { getPivotChartAxisState, getPivotChartReport } from '../services/pivot-chart-engine.ts'
import { getPivotCollapse } from '../services/pivot-axis-state.ts'
import { pivotChartControls } from '../services/pivot-chart-controls.ts'
import { pivotChartAxisMenu } from '../services/pivot-chart-axis-menu.ts'
import { capturePivotChartPanelOwner, pivotChartPanelOwnerContext, pivotChartPanelOwnerCurrent, type PivotChartPanelOwner } from '../services/pivot-chart-panel-owner.ts'
import { PIVOT_HIERARCHY_COMMANDS } from '../services/pivot-hierarchy-api-types.ts'

export interface PivotChartAxisViewSelection extends PivotChartAxisSelection {
  key: string
  target: PivotChartCanvasAxisTarget
  owner: PivotChartPanelOwner
  generation: number
}
interface PivotChartAxisHost {
  studio(): OfficeStudio | undefined
  unitId(): string | undefined
  sheetId(): string
  active(): boolean
  chart(chartId: string): ChartModel | undefined
  target(chartId: string, key: string): PivotChartCanvasAxisTarget | undefined
  frameRevision(chartId: string): number | undefined
  busy(chartId: string): boolean
  changed(): void
  focus(chartId: string): void
}

export function usePivotChartAxis(host: PivotChartAxisHost) {
  const selected = shallowRef<PivotChartAxisViewSelection>(), menuOpen = ref(false), error = ref('')
  const commit = shallowRef<{ chartId: string; owner: PivotChartPanelOwner; generation: number }>()
  const serviceOwner = {}, busy = computed(() => !!commit.value)
  let generation = 0, disposed = false

  function sameHost(owner: PivotChartPanelOwner): boolean {
    return !disposed && host.active() && owner.studio === host.studio() && owner.unitId === host.unitId() && owner.sheetId === host.sheetId()
  }
  function state(chartId: string) {
    const studio = host.studio(), unitId = host.unitId(), chart = host.chart(chartId), workbook = studio && unitId ? studio.getWorkbookModel(unitId) : undefined
    if (!studio || !workbook || !chart || !chart.ownedPivot && !chart.pivotTableId) return undefined
    try { assertChartReadPermission(workbook, chart, studio.currentUserId); return getPivotChartAxisState(workbook, chart) }
    catch { return undefined }
  }
  function editable(chartId: string): boolean {
    const studio = host.studio(), unitId = host.unitId(), chart = host.chart(chartId), workbook = studio && unitId ? studio.getWorkbookModel(unitId) : undefined
    return !!studio && !!workbook && !!chart && host.active() && !host.busy(chartId) && !commit.value
      && canEditPivotChart(workbook, chart, studio.currentUserId)
  }
  function isCurrent(selection: PivotChartAxisViewSelection): boolean {
    const live = getPivotChartAxisSelection(selection.owner.studio), target = host.target(selection.pivotChartId, selection.key)
    return sameHost(selection.owner) && pivotChartPanelOwnerCurrent(selection.owner) && selection.generation === generation
      && !!live && live.pivotChartId === selection.pivotChartId && live.fieldId === selection.fieldId && live.item === selection.item
      && !!target && JSON.stringify(target.label) === JSON.stringify(selection.target.label)
  }
  function select(chartId: string, target: PivotChartCanvasAxisTarget, openMenu = false): PivotChartAxisViewSelection | undefined {
    const studio = host.studio(), unitId = host.unitId(), chart = host.chart(chartId), workbook = studio && unitId ? studio.getWorkbookModel(unitId) : undefined
    if (disposed || !host.active() || !studio || !unitId || !workbook || !chart || commit.value || host.busy(chartId)
      || target.label.level >= target.label.depth - 1) return undefined
    try {
      assertChartReadPermission(workbook, chart, studio.currentUserId)
      const owner = capturePivotChartPanelOwner(studio, unitId, host.sheetId()), report = getPivotChartReport(workbook, chart)
      if (host.frameRevision(chartId) !== undefined && host.frameRevision(chartId) !== owner.revision) return undefined
      const axis = getPivotChartAxisState(workbook, chart)
      if (axis.fields[target.label.level]?.id !== target.label.fieldId || target.label.level >= axis.fields.length - 1) return undefined
      const selection: PivotChartAxisViewSelection = { ...pivotChartPanelOwnerContext(owner), unitId, subUnitId: host.sheetId(), pivotChartId: chartId,
        fieldId: target.label.fieldId, item: target.label.item, collapsed: getPivotCollapse(report, target.label.fieldId, target.label.item),
        fields: axis.fields, expectedRevision: owner.revision, expectedHistoryGeneration: owner.historyGeneration,
        expectedUserId: owner.userId, host: 'workbench', chartSnapshot: JSON.stringify(chart), reportSnapshot: JSON.stringify(report),
        key: target.key, target, owner, generation: ++generation }
      setPivotChartAxisSelection(studio, serviceOwner, selection)
      selected.value = selection; menuOpen.value = openMenu; error.value = ''; host.changed()
      return selection
    } catch (cause) {
      error.value = cause instanceof Error ? cause.message : '透视图轴目标当前不可用。'
      return undefined
    }
  }
  function close(restoreFocus = false): void {
    const previous = selected.value, ticket = ++generation
    if (previous) clearPivotChartAxisSelection(previous.owner.studio, serviceOwner)
    selected.value = undefined; menuOpen.value = false; error.value = ''; host.changed()
    if (restoreFocus && previous) void nextTick(() => {
      if (generation === ticket && !selected.value && sameHost(previous.owner) && pivotChartPanelOwnerCurrent(previous.owner)) host.focus(previous.pivotChartId)
    })
  }
  function closeMenu(restoreFocus = false): void {
    const previous = selected.value
    menuOpen.value = false
    if (restoreFocus && previous) void nextTick(() => { if (selected.value === previous && isCurrent(previous)) host.focus(previous.pivotChartId) })
  }
  function openSelectedMenu(): boolean {
    const selection = selected.value
    if (!selection || !isCurrent(selection)) return false
    menuOpen.value = true; return true
  }
  const entries = computed(() => selected.value ? pivotChartAxisMenu(selected.value, !isCurrent(selected.value) || !editable(selected.value.pivotChartId)) : [])

  async function execute(key: string): Promise<boolean> {
    const selection = selected.value, entry = entries.value.find(entry => entry.key === key)
    if (!selection || !entry || entry.disabled || !isCurrent(selection) || !editable(selection.pivotChartId)) return false
    const transaction = { chartId: selection.pivotChartId, owner: selection.owner, generation }
    commit.value = transaction; menuOpen.value = false; error.value = ''
    try {
      const accepted = await selection.owner.studio.executeCommand<typeof entry.params, boolean>(entry.commandId, entry.params)
      if (selected.value === selection && sameHost(selection.owner)) close()
      if (accepted && sameHost(selection.owner) && pivotChartPanelOwnerCurrent(selection.owner, true)) {
        const focusGeneration = generation
        await nextTick(); if (generation === focusGeneration && sameHost(selection.owner) && !selected.value) host.focus(selection.pivotChartId)
      }
      return accepted
    } catch (cause) {
      if (generation === transaction.generation && sameHost(selection.owner)) error.value = cause instanceof Error ? cause.message : '透视图展开折叠未能提交。'
      return false
    } finally { if (commit.value === transaction) commit.value = undefined }
  }
  async function step(chartId: string, direction: 'expand' | 'collapse'): Promise<boolean> {
    const axis = state(chartId), chart = host.chart(chartId), studio = host.studio(), unitId = host.unitId()
    if (!axis || !chart || !studio || !unitId || !pivotChartControls(chart).showExpandCollapseButtons || !editable(chartId)
      || !(direction === 'expand' ? axis.canExpand : axis.canCollapse)) return false
    let owner: PivotChartPanelOwner
    try { owner = capturePivotChartPanelOwner(studio, unitId, host.sheetId()) }
    catch (cause) { error.value = cause instanceof Error ? cause.message : '透视图工作台宿主已变更。'; return false }
    if (host.frameRevision(chartId) !== undefined && host.frameRevision(chartId) !== owner.revision) return false
    close()
    const transaction = { chartId, owner, generation }
    commit.value = transaction; error.value = ''
    try {
      const input = { ...pivotChartPanelOwnerContext(owner), unitId, subUnitId: host.sheetId(), pivotChartId: chartId, direction }
      const accepted = await studio.executeCommand<typeof input, boolean>(PIVOT_HIERARCHY_COMMANDS.chartStepCategoryLevel, input)
      if (accepted && sameHost(owner) && pivotChartPanelOwnerCurrent(owner, true)) {
        await nextTick(); if (generation === transaction.generation && sameHost(owner) && !selected.value) host.focus(chartId)
      }
      return accepted
    } catch (cause) {
      if (generation === transaction.generation && sameHost(owner)) error.value = cause instanceof Error ? cause.message : '透视图类别层级未能提交。'
      return false
    } finally { if (commit.value === transaction) commit.value = undefined }
  }
  function reconcile(): void {
    const selection = selected.value
    if (selection && (!sameHost(selection.owner) || commit.value?.chartId !== selection.pivotChartId && !isCurrent(selection))) close()
  }
  watch(() => [host.studio(), host.unitId(), host.sheetId(), host.active()], reconcile, { flush: 'sync' })
  onScopeDispose(() => { close(); disposed = true })
  return { selected, menuOpen, entries, commit, busy, error, state, editable, isCurrent, select, close, closeMenu, openSelectedMenu, execute, step, reconcile }
}
