<script setup lang="ts">
import { computed, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue'
import { Alert, Button, Spin, Tooltip, theme } from 'antdv-next'
import type { OfficeStudio } from '../api/office-studio'
import type { SheetDrawingCommandInput } from '../api/sheet-drawing-commands.ts'
import { FOffice } from '../api/facade'
import type { ChartModel, ChartPlotAreaOptions } from '../domain/model'
import { chartEChartsOption } from '../features/charts/echarts-option'
import { EChartsCanvasRenderer, type EChartsCanvasFrame } from '../features/charts/echarts-renderer'
import type { ChartVector } from '../features/charts/chart-geometry'
import type { ChartCanvasLegendHit } from '../features/charts/echarts-legend'
import type { ChartComputedStatistics } from '../features/charts/chart-derived'
import { chartContentTooltipText, type ChartContentTooltip } from '../features/charts/chart-content-tooltip'
import { useChartStatistics } from '../features/charts/views/use-chart-statistics'
import { canEditWorksheetProtection } from '../features/sheet-protection/range-protection'
import type { DrawCommand } from '../rendering/contracts'
import type { ExcelWorkbookChartView } from './chart-overlay-model'
import { cloneChart, toChartWorkbook } from './chart-overlay-model'
import { mergeChartOptions } from '../features/charts/chart-options'
import { chartTitlePatch, type ChartTitleKind, type ChartTitleObjectSpec } from '../features/charts/chart-titles'
import { chartTitleAuthorContextMatches, chartTitleDragLayout, chartTitleOutlineCommands, type ChartCanvasTitleTarget, type ChartTitleAuthorContext } from '../features/charts/chart-title-interaction'
import ChartTitleFloatMenu from '../features/charts/views/ChartTitleFloatMenu.vue'
import { supportsChartPlotArea, validateChartPlotArea } from '../features/charts/chart-plot-area'
import { CHART_PLOT_AREA_HANDLES, chartPlotAreaDragLayout, chartPlotAreaHandlePoints, chartPlotAreaOutlineCommands, chartPlotAreasEqual,
  type ChartCanvasPlotAreaTarget, type ChartPlotAreaResizeHandle } from '../features/charts/chart-plot-area-interaction'
import ChartPlotAreaFloatMenu from '../features/charts/views/ChartPlotAreaFloatMenu.vue'
import ChartObjectContextMenu from '../features/charts/views/ChartObjectContextMenu.vue'
import { useChartObjectActions, type ChartObjectChange } from '../features/charts/views/use-chart-object-actions'
import { CHART_OBJECT_COMMANDS, PASTE_COPIED_CHART_COMMAND, chartObjectShortcut, type ChartObjectTarget } from '../features/charts/chart-object-commands'
import type { ChartEditPanelRoute } from '../features/charts/chart-edit-panel-route'
import { chartCanvasElementKey, chartCanvasElementOutline, type ChartCanvasElement, type ChartCanvasElementTarget } from '../features/charts/chart-element-interaction'
import type { ChartConfig } from '../features/charts/chart-engine'
import ChartElementFloatMenu from '../features/charts/views/ChartElementFloatMenu.vue'
import { useChartElementEditing, type ChartCanvasElementSelection } from '../features/charts/views/use-chart-element-editing'
import { drawingAffineFrame, drawingAffineWithFrame, type DrawingAffineGeometry } from '../domain/drawing-affine.ts'
import type { SheetDrawingLeafRef, SheetDrawingRef } from '../domain/sheet-drawings.ts'
import { chartDrawingLocalDelta, chartDrawingLocalPoint, chartDrawingOutlineCommands, chartDrawingViewportBounds, chartDrawingViewportPoint, chartDrawingWorldPoint, type ChartDrawingViewport } from '../features/charts/chart-drawing-affine.ts'
import { GRID_HEADER_WIDTH, GRID_HEADER_HEIGHT } from '../rendering/spreadsheet/layout.ts'
import { sheetDrawingMoveDirectionForKey } from '../features/drawing-groups/drawing-group-move.ts'
import { SHEET_DRAWING_GROUP_COMMAND_IDS } from '../features/drawing-groups/index.ts'
import { usePivotChartAxis } from '../features/pivot/views/use-pivot-chart-axis.ts'
import PivotChartAxisMenu from '../features/pivot/views/PivotChartAxisMenu.vue'
import PivotChartHierarchyControls from '../features/pivot/views/PivotChartHierarchyControls.vue'
import PivotChartFieldButtons from '../features/pivot/views/PivotChartFieldButtons.vue'
import { pivotChartControls } from '../features/pivot/services/pivot-chart-controls.ts'
import type { PivotChartCanvasAxisTarget } from '../features/charts/echarts-pivot-axis-interaction.ts'

const props = withDefaults(defineProps<{
  workbook: ExcelWorkbookChartView
  activeSheetId: string
  zoom: number
  headerWidth?: number
  headerHeight?: number
  computedStatistics?: Readonly<Record<string, ChartComputedStatistics>>
  studio?: OfficeStudio
  active?: boolean
  restoreFocusAllowed?: boolean
  openChartEditor?: (target: ChartObjectTarget, route?: ChartEditPanelRoute) => boolean
  openPivotChartSettings?: (chartId: string, fieldId: string) => boolean
  selectedDrawingRefs?: readonly SheetDrawingRef[]
  drawingOrder?: readonly SheetDrawingLeafRef[]
  drawingVisible?: (ref: SheetDrawingLeafRef) => boolean
  routeDrawingEvent?: (event: MouseEvent, chartId: string) => boolean
}>(), { zoom: 100, active: true, restoreFocusAllowed: true })

const emit = defineEmits<{
  (event: 'update-chart', payload: { chartId: string; patch: Partial<Pick<ChartModel, 'position' | 'size'>> }): void
  (event: 'remove-chart', chartId: string): void
  (event: 'overlay-commands', commands: DrawCommand[]): void
  (event: 'select-chart', chartId: string | null): void
  (event: 'object-change', change: ChartObjectChange): void
  (event: 'drawing-command', id: string, input: SheetDrawingCommandInput): void
}>()

interface ChartDisplay {
  chart: ChartModel
  position: { x: number; y: number }
  size: { width: number; height: number }
  drawingGeometry?: DrawingAffineGeometry
}
interface TitleSelection {
  chartId: string
  kind: ChartTitleKind
  workbookId?: string
  sheetId: string
  studio?: OfficeStudio
  snapshot: string
  context?: ChartTitleAuthorContext
  generation: number
}
type PlotAreaSelection = Omit<TitleSelection, 'kind'>
interface DragState {
  chartId: string
  mode: 'move' | 'resize' | 'sankey-node' | 'legend' | 'title' | 'plot-area' | 'element'
  legendHit?: ChartCanvasLegendHit
  titleHit?: ChartCanvasTitleTarget
  titleSelection?: TitleSelection
  plotAreaHit?: ChartCanvasPlotAreaTarget
  plotAreaSelection?: PlotAreaSelection
  plotAreaHandle?: ChartPlotAreaResizeHandle
  elementSelection?: ChartCanvasElementSelection
  pointerId: number
  target: HTMLElement | null
  workbookId?: string
  studio?: OfficeStudio
  sheetId: string
  modelRevision?: number
  chartSnapshot: string
  originX: number
  originY: number
  position: { x: number; y: number }
  size: { width: number; height: number }
  contentSize: { width: number; height: number }
  pane: ChartDrawingViewport
}

const root = ref<HTMLElement | null>(null)
const objectHostActive = ref(false)
const { token } = theme.useToken()
const scrollLeft = ref(0)
const scrollTop = ref(0)
const selectedId = ref<string | null>(null)
const renderError = ref('')
const pointerError = ref('')
const renderedFrames = ref<Record<string, { version: number; width: number; height: number }>>({})
const drafts = ref<Record<string, Pick<ChartDisplay, 'position' | 'size' | 'drawingGeometry'>>>({})
const drag = shallowRef<DragState | null>(null)
const contentTooltip = shallowRef<{ chartId: string; point: ChartVector; content: ChartContentTooltip }>()
const hoveredTitle = shallowRef<{ chartId: string; kind: ChartTitleKind }>()
const selectedTitle = shallowRef<TitleSelection | null>(null)
const titleCommit = shallowRef<TitleSelection | null>(null)
const titleDrafts = ref<Record<string, { kind: ChartTitleKind; value: ChartTitleObjectSpec }>>({})
const titleMenu = ref<{ focusText: (selectAll?: boolean) => Promise<void> } | null>(null)
const hoveredPlotArea = ref<string>()
const selectedPlotArea = shallowRef<PlotAreaSelection | null>(null)
const plotAreaCommit = shallowRef<PlotAreaSelection | null>(null)
const plotAreaDrafts = ref<Record<string, ChartPlotAreaOptions>>({})
const plotAreaMenu = ref<{ focusCoordinate: () => Promise<void> } | null>(null)
const hoveredElement = shallowRef<{ chartId: string; element: ChartCanvasElement }>()
const hoveredPivotAxis = shallowRef<{ chartId: string; target: PivotChartCanvasAxisTarget }>()
const pivotFieldBusy = ref<Record<string, boolean>>({})
const anyPivotFieldBusy = computed(() => Object.values(pivotFieldBusy.value).some(Boolean))
const pivotAxisMenuPoint = ref<ChartVector>({ x: 0, y: 0 })
const elementMenu = ref<{ focusFirst: () => Promise<void> } | null>(null)
let titleGeneration = 0
let plotAreaGeneration = 0
let scroller: HTMLElement | null = null
let resizeObserver: ResizeObserver | undefined
let frameRequest: number | undefined
let pixelRatioQuery: MediaQueryList | undefined
let mounted = false
let contentDirty = true
let pointerDirty = false
let desiredPointer: { chartId: string; point: ChartVector } | undefined
let desiredContentDrag: { chartId: string; pointerId: number; point: ChartVector } | undefined
let activePointerId: string | undefined
let routedDispatch = false
let routedClickChartId: string | undefined
let routedPane: ChartDrawingViewport | undefined
const authorPanes = new Map<string, ChartDrawingViewport>()
let contentCommands: DrawCommand[] = []
const commandIndexes = new Map<string, number>()
const chartRenderer = new EChartsCanvasRenderer()
const objectActions = useChartObjectActions({
  studio: () => props.studio, unitId: () => props.workbook.id, sheetId: () => props.activeSheetId,
  selectedId: () => selectedId.value, active: () => objectHostActive.value && props.active,
  panelAvailable: () => !!props.openChartEditor,
  authorBusy: () => !!titleCommit.value || !!plotAreaCommit.value || elementEditing.busy.value || pivotAxis.busy.value || anyPivotFieldBusy.value,
  canExport: chartId => !!renderedFrames.value[chartId] && !pending.value[chartId] && !titleCommit.value && !plotAreaCommit.value && !elementEditing.busy.value && !pivotAxis.busy.value && !anyPivotFieldBusy.value,
  prepareExport(target) {
    if (titleCommit.value || plotAreaCommit.value || elementEditing.busy.value || pivotAxis.busy.value || anyPivotFieldBusy.value) throw new Error('图表作者修改正在提交，请完成后导出图片。')
    onPointerCancel(); closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); pivotAxis.close(); clearChartPointer()
    const chart = sheetCharts.value.find(chart => chart.id === target.chartId)
    if (!chart || JSON.stringify(chart) !== target.snapshot || pending.value[target.chartId]) throw new Error('图表画面尚未同步到当前作者版本。')
    const snapshot = snapshots.value[target.chartId]
    if (snapshot && snapshot.modelRevision !== target.expectedRevision) throw new Error('图表计算结果尚未同步，请稍后导出。')
    contentCommands = renderCommands(); contentDirty = false
    updatePointerFrames(); pointerDirty = false; queueOverlayFrame()
    return chartRenderer.getImageFrame(target.chartId)?.source
  },
  openPanel: (target, route) => !!props.openChartEditor?.(target, route),
  changed(change) {
    closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); pivotAxis.close()
    if (change.action === 'delete' && selectedId.value === change.chartId) selectedId.value = null
    if (change.action === 'paste') selectedId.value = change.chartId
    emit('select-chart', selectedId.value); emit('object-change', change)
  },
  focus(chartId) {
    if (!props.restoreFocusAllowed) return
    const button = Array.from(root.value?.querySelectorAll<HTMLElement>('.chart-hitbox') ?? []).find(node => node.dataset.chartId === chartId)
    if (button) button.focus({ preventScroll: true })
    else root.value?.parentElement?.querySelector<HTMLElement>('.canvas-scroll canvas[tabindex], .canvas-scroll [tabindex="0"]')?.focus({ preventScroll: true })
  },
  notify: message => { pointerError.value = message },
})

const sheetCharts = computed(() => (props.workbook.charts ?? []).filter(chart => chart.sheetId === props.activeSheetId
  && (props.drawingVisible?.({ kind: 'chart', id: chart.id }) ?? true)))
const domainWorkbook = computed(() => props.studio ? props.studio.getWorkbookModel(props.workbook.id ?? '') : toChartWorkbook(props.workbook))
const { statistics, snapshots, pending, errors: computeErrors, diagnostics } = useChartStatistics({
  studio: () => props.studio, workbook: () => props.studio ? props.studio.getWorkbookModel(props.workbook.id ?? '') : domainWorkbook.value, activeSheetId: () => props.activeSheetId,
  externalStatistics: () => props.computedStatistics,
})
const elementEditing = useChartElementEditing({
  studio: () => props.studio, workbookId: () => props.workbook.id, sheetId: () => props.activeSheetId,
  active: () => objectHostActive.value && props.active,
  chart: chartId => sheetCharts.value.find(chart => chart.id === chartId),
  target: (chartId, element) => chartRenderer.getElementTarget(chartId, element),
  canEdit: canEditChart,
  busy: chartId => !!titleCommit.value || !!plotAreaCommit.value || !!drag.value || !!objectActions.busy.value || !!pending.value[chartId] || pivotAxis.busy.value || anyPivotFieldBusy.value,
  changed: queueOverlayFrame,
  focus: chartId => { if (props.restoreFocusAllowed) objectActions.restoreFocus(objectActions.capture(chartId)) },
  openPanel: (chartId, route) => objectActions.execute('edit', objectActions.capture(chartId), route),
})
const selectedElement = elementEditing.selected
const elementCommit = elementEditing.commit
const elementDisabled = elementEditing.disabled
const pivotAxis = usePivotChartAxis({
  studio: () => props.studio, unitId: () => props.workbook.id, sheetId: () => props.activeSheetId,
  active: () => mounted && objectHostActive.value && props.active,
  chart: chartId => sheetCharts.value.find(chart => chart.id === chartId),
  target: (chartId, key) => chartRenderer.getPivotAxisTarget(chartId, key),
  frameRevision: chartId => snapshots.value[chartId]?.modelRevision,
  busy: chartId => !!drag.value || !!titleCommit.value || !!plotAreaCommit.value || elementEditing.busy.value || objectActions.busy.value || !!pending.value[chartId] || anyPivotFieldBusy.value,
  changed: queueOverlayFrame,
  focus: chartId => { if (props.restoreFocusAllowed) objectActions.restoreFocus(objectActions.capture(chartId)) },
})
const selectedPivotAxis = pivotAxis.selected
const pivotAxisState = computed(() => Object.fromEntries(sheetCharts.value.map(chart => [chart.id, pivotAxis.state(chart.id)])))
const pivotAxisMenuStyle = computed(() => ({ left: `${pivotAxisMenuPoint.value.x}px`, top: `${pivotAxisMenuPoint.value.y}px` }))
const pivotAxisTooltipStyle = computed(() => {
  const hover = hoveredPivotAxis.value, item = displayCharts.value.find(item => item.chart.id === hover?.chartId)
  const bounds = hover && item ? chartDrawingViewportBounds(item, hover.target.bounds, chartView(item.chart.id)) : undefined
  return bounds ? { left: `${bounds.x + bounds.width / 2}px`, top: `${bounds.y}px` } : {}
})
const errorMessage = computed(() => renderError.value || pointerError.value || elementEditing.error.value || pivotAxis.error.value || Object.values(computeErrors.value).join('；'))
const displayCharts = computed<ChartDisplay[]>(() => sheetCharts.value.map((chart) => {
  const draft = drafts.value[chart.id]
  const position = draft?.position ?? chart.position
  const size = draft?.size ?? chart.size
  return { chart, position, size, drawingGeometry: draft?.drawingGeometry ?? chart.drawingGeometry }
}))

const chartTheme = computed(() => ({
  backgroundColor: token.value.colorBgContainer,
  textColor: token.value.colorText,
  axisColor: token.value.colorTextSecondary,
  splitLineColor: token.value.colorBorderSecondary,
  axisPointerLabelBackground: token.value.colorBgSpotlight,
  axisPointerLabelColor: token.value.colorTextLightSolid,
}))
const overlayStyle = computed(() => ({
  '--chart-primary': token.value.colorPrimary,
  '--chart-focus-ring': token.value.colorPrimaryBgHover,
  '--chart-surface': token.value.colorBgContainer,
  '--chart-author-layer': String((props.drawingOrder?.length ?? sheetCharts.value.length) + 10),
}))
const contentTooltipText = computed(() => chartContentTooltipText(contentTooltip.value?.content))
const contentTooltipStyle = computed(() => {
  const tooltip = contentTooltip.value, item = displayCharts.value.find(item => item.chart.id === tooltip?.chartId)
  const point = tooltip && item ? chartDrawingViewportPoint(item, tooltip.point, chartView(item.chart.id)) : undefined
  return point ? { left: `${point.x}px`, top: `${point.y}px` } : {}
})
const selectedTitleChart = computed(() => sheetCharts.value.find(chart => chart.id === selectedTitle.value?.chartId))
const titleMenuApply = computed(() => {
  const selection = selectedTitle.value
  // A removed input's delayed composition/blur event keeps its old token.
  // It cannot apply a draft to another title or a reused chart ID.
  return (value: ChartTitleObjectSpec) => applyTitleSpec(selection, value)
})
const titleMenuStyle = computed(() => {
  const selection = selectedTitle.value, item = displayCharts.value.find(item => item.chart.id === selection?.chartId)
  // Reading the frame version refreshes the measured anchor after each paint.
  const frame = selection ? renderedFrames.value[selection.chartId] : undefined
  const target = selection && frame ? chartRenderer.getTitleTarget(selection.chartId, selection.kind) : undefined
  if (!item || !frame || !root.value || !selection) return { display: 'none' }
  const manual = item.chart.options?.[selection.kind === 'title' ? 'titleStyle' : 'subtitleStyle']?.manualLayout
  // An already selected manual anchor can be clipped out of the frame. Keep
  // its automatic-restore control reachable; this does not create a hit box.
  const bounds = target?.bounds ?? (manual ? { x: manual.left * item.size.width, y: manual.top * item.size.height, width: 0, height: 0 } : undefined)
  if (!bounds) return { display: 'none' }
  const projected = chartDrawingViewportBounds(item, bounds, chartView(item.chart.id))
  const width = Math.min(440, Math.max(0, root.value.clientWidth - 52))
  const left = Math.max(48, Math.min(projected.x, root.value.clientWidth - width - 4))
  const above = projected.y - 190
  const top = Math.max(30, Math.min(above >= 30 ? above : projected.y + projected.height + 8,
    Math.max(30, root.value.clientHeight - 194)))
  return { left: `${left}px`, top: `${top}px`, width: `${width}px` }
})
const selectedPlotAreaChart = computed(() => sheetCharts.value.find(chart => chart.id === selectedPlotArea.value?.chartId))
const selectedPlotAreaTarget = computed(() => selectedPlotArea.value && renderedFrames.value[selectedPlotArea.value.chartId]
  ? chartRenderer.getPlotAreaTarget(selectedPlotArea.value.chartId) : undefined)
const plotAreaMenuApply = computed(() => {
  const selection = selectedPlotArea.value
  return (area: ChartPlotAreaOptions | undefined) => applyPlotArea(selection, area)
})
const plotAreaMenuStyle = computed(() => {
  const selection = selectedPlotArea.value, target = selectedPlotAreaTarget.value
  const item = displayCharts.value.find(item => item.chart.id === selection?.chartId)
  if (!target || !item || !root.value) return { display: 'none' }
  const bounds = target.bounds ?? target.plot
  const projected = chartDrawingViewportBounds(item, bounds, chartView(item.chart.id))
  const width = Math.min(410, Math.max(0, root.value.clientWidth - 52))
  const left = Math.max(48, Math.min(projected.x, root.value.clientWidth - width - 4))
  const below = projected.y + projected.height + 10
  const top = Math.max(30, Math.min(below, Math.max(30, root.value.clientHeight - 252)))
  return { left: `${left}px`, top: `${top}px`, width: `${width}px` }
})
const selectedElementChart = computed(() => sheetCharts.value.find(chart => chart.id === selectedElement.value?.chartId))
const selectedElementTarget = computed(() => selectedElement.value && renderedFrames.value[selectedElement.value.chartId]
  ? chartRenderer.getElementTarget(selectedElement.value.chartId, selectedElement.value.element) : undefined)
const elementMenuApply = computed(() => {
  const selection = selectedElement.value
  return (draft: ChartConfig) => elementEditing.apply(selection, draft)
})
const elementMenuMore = computed(() => {
  const selection = selectedElement.value
  return () => elementEditing.more(selection)
})
const elementMenuPromote = computed(() => {
  const selection = selectedElement.value
  return async () => {
    const next = elementEditing.promote(selection)
    await nextTick()
    if (next && mounted && selectedElement.value === next) await elementMenu.value?.focusFirst()
  }
})
const elementMenuStyle = computed(() => {
  const selection = selectedElement.value, target = selectedElementTarget.value
  const item = displayCharts.value.find(item => item.chart.id === selection?.chartId)
  if (!target || !item || !root.value) return { display: 'none' }
  const projected = chartDrawingViewportBounds(item, target.bounds, chartView(item.chart.id))
  const width = Math.min(420, Math.max(0, root.value.clientWidth - 52)), height = Math.min(420, Math.max(80, root.value.clientHeight - 38))
  const left = Math.max(48, Math.min(projected.x, root.value.clientWidth - width - 4))
  const below = projected.y + projected.height + 8
  const top = Math.max(30, Math.min(below, Math.max(30, root.value.clientHeight - height - 4)))
  return { left: `${left}px`, top: `${top}px`, width: `${width}px`, '--chart-element-max-height': `${height}px` }
})
function plotAreaHandles(item: ChartDisplay): { handle: ChartPlotAreaResizeHandle; label: string; style: Record<string, string> }[] {
  if (selectedPlotArea.value?.chartId !== item.chart.id || !renderedFrames.value[item.chart.id]) return []
  const target = chartRenderer.getPlotAreaTarget(item.chart.id)
  return target ? chartPlotAreaHandlePoints(target).map(point => {
    const definition = CHART_PLOT_AREA_HANDLES.find(entry => entry.handle === point.handle)!
    return { handle: point.handle, label: definition.label,
      style: { left: `${point.x * zoomScale()}px`, top: `${point.y * zoomScale()}px`, cursor: definition.cursor } }
  }) : []
}

function renderCommands(): DrawCommand[] {
  const host = root.value
  commandIndexes.clear()
  if (!host || host.clientWidth <= 0 || host.clientHeight <= 0) {
    chartRenderer.retain(new Set())
    renderedFrames.value = {}
    activePointerId = undefined
    desiredPointer = undefined
    contentTooltip.value = undefined
    return []
  }
  const workbook = props.studio ? props.studio.getWorkbookModel(props.workbook.id ?? '') : domainWorkbook.value
  const dpr = Math.max(0.1, window.devicePixelRatio || 1) * zoomScale()
  const commands: DrawCommand[] = []
  const frames: Record<string, { version: number; width: number; height: number }> = {}
  let errorMessage = ''
  for (const item of displayCharts.value) {
    const computed = statistics.value[item.chart.id]
    if (!computed || !workbook) continue
    try {
      const draft = drafts.value[item.chart.id], titleDraft = titleDrafts.value[item.chart.id], plotAreaDraft = plotAreaDrafts.value[item.chart.id]
      let options = item.chart.options
      if (titleDraft) options = mergeChartOptions(options, chartTitlePatch(titleDraft.value, titleDraft.kind).options)
      if (plotAreaDraft) options = mergeChartOptions(options, { plotArea: plotAreaDraft })
      const model = draft || titleDraft || plotAreaDraft ? { ...cloneChart(item.chart), position: { ...item.position }, size: { ...item.size }, options } : item.chart
      const option = chartEChartsOption(workbook, model, chartTheme.value, computed)
      const frame = chartRenderer.render(item.chart.id, option, {
        width: Math.max(1, item.size.width),
        height: Math.max(1, item.size.height),
        devicePixelRatio: dpr,
      }, chartTheme.value)
      frames[item.chart.id] = { version: frame.sourceVersion, width: frame.source.width, height: frame.source.height }
      commandIndexes.set(item.chart.id, commands.length)
      const affine = item.drawingGeometry
      commands.push({
        kind: 'image',
        drawingRef: { kind: 'chart', id: item.chart.id },
        source: frame.source,
        sourceVersion: frame.sourceVersion,
        x: affine ? 0 : item.position.x,
        y: affine ? 0 : item.position.y,
        width: affine?.width ?? item.size.width,
        height: affine?.height ?? item.size.height,
        matrix: affine ? { ...affine.matrix, e: affine.matrix.e + GRID_HEADER_WIDTH, f: affine.matrix.f + GRID_HEADER_HEIGHT } : undefined,
      })
    } catch (error) {
      errorMessage = error instanceof Error ? error.message : '图表绘制失败'
    }
  }
  renderedFrames.value = frames
  renderError.value = errorMessage
  const validFrames = new Set(Object.keys(frames))
  chartRenderer.retain(validFrames)
  if (activePointerId && !validFrames.has(activePointerId)) activePointerId = undefined
  if (desiredPointer && !validFrames.has(desiredPointer.chartId)) desiredPointer = undefined
  if (contentTooltip.value && !validFrames.has(contentTooltip.value.chartId)) contentTooltip.value = undefined
  return commands
}

function scaleRelationContent(event: WheelEvent, item: ChartDisplay): void {
  if (item.chart.type !== 'relation' || drag.value || !commandIndexes.has(item.chart.id)) return
  event.preventDefault(); event.stopPropagation()
  const point = chartAxisPointerLocalPoint(pointerPoint(event), item.position,
    defaultChartView())
  const delta = event.deltaY * (event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? item.size.height : 1)
  try {
    updatePointerFrame(item.chart.id, chartRenderer.scaleContent(item.chart.id, point, delta))
    desiredPointer = { chartId: item.chart.id, point }; pointerDirty = true
    queueOverlayFrame()
  } catch (error) { pointerError.value = error instanceof Error ? error.message : '关系图视图缩放失败' }
}

function zoomScale(): number { return Math.max(0.1, props.zoom / 100) }
function defaultChartView(): ChartDrawingViewport {
  return { zoom: props.zoom, scrollLeft: scrollLeft.value, scrollTop: scrollTop.value,
    headerWidth: props.headerWidth ?? GRID_HEADER_WIDTH, headerHeight: props.headerHeight ?? GRID_HEADER_HEIGHT }
}
function chartView(chartId?: string): ChartDrawingViewport {
  return routedPane ?? (chartId && drag.value?.chartId === chartId ? drag.value.pane : undefined)
    ?? (chartId ? authorPanes.get(chartId) : undefined) ?? defaultChartView()
}
function viewportX(value: number): number { return (props.headerWidth ?? GRID_HEADER_WIDTH) + (value - GRID_HEADER_WIDTH) * zoomScale() - scrollLeft.value }
function viewportY(value: number): number { return (props.headerHeight ?? GRID_HEADER_HEIGHT) + (value - GRID_HEADER_HEIGHT) * zoomScale() - scrollTop.value }
function pointerPoint(event: Pick<MouseEvent, 'clientX' | 'clientY'>): { x: number; y: number } {
  const rect = root.value?.getBoundingClientRect()
  return rect ? { x: event.clientX - rect.left, y: event.clientY - rect.top } : { x: 0, y: 0 }
}

function chartStyle(item: ChartDisplay): Record<string, string> {
  const rank = props.drawingOrder?.findIndex(ref => ref.kind === 'chart' && ref.id === item.chart.id),
    zIndex = String(rank !== undefined && rank >= 0 ? rank + 1 : 1)
  if (item.drawingGeometry) {
    const geometry = item.drawingGeometry, m = geometry.matrix, scale = zoomScale(),
      sx = geometry.width / item.size.width, sy = geometry.height / item.size.height
    return { zIndex, left: '0px', top: '0px', width: `${item.size.width * scale}px`, height: `${item.size.height * scale}px`,
      transformOrigin: '0 0', transform: `matrix(${m.a * sx},${m.b * sx},${m.c * sy},${m.d * sy},${(props.headerWidth ?? GRID_HEADER_WIDTH) + m.e * scale - scrollLeft.value},${(props.headerHeight ?? GRID_HEADER_HEIGHT) + m.f * scale - scrollTop.value})` }
  }
  return {
    zIndex,
    left: `${viewportX(item.position.x)}px`,
    top: `${viewportY(item.position.y)}px`,
    width: `${item.size.width * zoomScale()}px`,
    height: `${item.size.height * zoomScale()}px`,
  }
}
function chartAxisPointerLocalPoint(point: ChartVector, position: ChartVector, view: ChartDrawingViewport): ChartVector {
  const active = drag.value, item = displayCharts.value.find(item => item.position === position || item.chart.position === position
    || active?.position === position && active.chartId === item.chart.id)
  return chartDrawingLocalPoint(item ?? { position, size: { width: 1, height: 1 } }, point, item ? chartView(item.chart.id) : view)
}

function selectChart(chartId: string): void {
  if (selectedTitle.value && selectedTitle.value.chartId !== chartId) closeTitleSelection()
  if (selectedPlotArea.value && selectedPlotArea.value.chartId !== chartId) closePlotAreaSelection()
  if (selectedElement.value && selectedElement.value.chartId !== chartId) closeElementSelection()
  if (selectedPivotAxis.value && selectedPivotAxis.value.pivotChartId !== chartId) pivotAxis.close()
  selectedId.value = chartId
  emit('select-chart', chartId)
}
function selectWholeChart(chartId: string): void { closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); pivotAxis.close(); selectChart(chartId) }
function activatePivotFieldButtons(chartId: string): void {
  if (!props.active || !mounted || anyPivotFieldBusy.value) return
  clearChartPointer(); selectWholeChart(chartId)
}
function openPivotFieldSettings(chartId: string, fieldId: string): boolean {
  return Boolean(props.openPivotChartSettings?.(chartId, fieldId))
}
function restorePivotFieldFocus(chartId: string): void {
  if (props.restoreFocusAllowed) objectActions.restoreFocus(objectActions.capture(chartId))
}
function setPivotFieldBusy(chartId: string, busy: boolean): void {
  pivotFieldBusy.value = { ...pivotFieldBusy.value, [chartId]: busy }
}
function canEditChart(chart: ChartModel): boolean {
  const workbook = props.studio ? props.studio.getWorkbookModel(props.workbook.id ?? '') : domainWorkbook.value
  const sheet = workbook?.sheets.find(sheet => sheet.id === chart.sheetId)
  return !!sheet && canEditWorksheetProtection(sheet, props.studio?.currentUserId ?? 'local-user')
}
function currentTitleContext(chartId: string): ChartTitleAuthorContext | undefined {
  const studio = props.studio, workbookId = props.workbook.id
  const workbook = studio && workbookId ? studio.getWorkbookModel(workbookId) : undefined
  const chart = workbook?.charts?.find(chart => chart.id === chartId && chart.sheetId === props.activeSheetId)
  return studio && workbookId && chart ? { owner: studio, workbookId, sheetId: props.activeSheetId, chartId,
    revision: studio.getWorkbookRevision(workbookId), snapshot: JSON.stringify(chart) } : undefined
}
function chartElementSelectionIsCurrent(selection: PlotAreaSelection): boolean {
  const chart = sheetCharts.value.find(chart => chart.id === selection.chartId)
  return mounted && selection.studio === props.studio && selection.workbookId === props.workbook.id && selection.sheetId === props.activeSheetId
    && !!chart && (selection.context ? chartTitleAuthorContextMatches(selection.context, currentTitleContext(selection.chartId)) : JSON.stringify(chart) === selection.snapshot)
}
function titleSelectionIsCurrent(selection: TitleSelection): boolean { return chartElementSelectionIsCurrent(selection) }
function selectTitle(chart: ChartModel, kind: ChartTitleKind): TitleSelection {
  pivotAxis.close()
  closePlotAreaSelection()
  closeElementSelection()
  const existing = selectedTitle.value
  if (existing?.chartId === chart.id && existing.kind === kind && titleSelectionIsCurrent(existing)) return existing
  const selection: TitleSelection = { chartId: chart.id, kind, workbookId: props.workbook.id, sheetId: props.activeSheetId, studio: props.studio,
    snapshot: JSON.stringify(chart), context: currentTitleContext(chart.id), generation: ++titleGeneration }
  selectedTitle.value = selection
  selectedId.value = chart.id
  queueOverlayFrame()
  return selection
}
function closeTitleSelection(restoreFocus = false): void {
  const previous = selectedTitle.value
  selectedTitle.value = null
  hoveredTitle.value = undefined
  if (drag.value?.mode === 'title') onPointerCancel()
  titleDrafts.value = {}
  queueOverlayFrame()
  if (restoreFocus && previous) void nextTick(() => {
    if (!mounted || selectedTitle.value || selectedElement.value || previous.studio !== props.studio || previous.workbookId !== props.workbook.id || previous.sheetId !== props.activeSheetId) return
    const button = Array.from(root.value?.querySelectorAll<HTMLElement>('.chart-hitbox') ?? []).find(node => node.dataset.chartId === previous.chartId)
    button?.focus({ preventScroll: true })
  })
}
async function applyTitleSpec(selection: TitleSelection | null, value: ChartTitleObjectSpec | undefined): Promise<void> {
  const chart = selection && sheetCharts.value.find(chart => chart.id === selection.chartId)
  if (!selection || selection !== selectedTitle.value || !selection.context || !props.studio || !chart || !canEditChart(chart)
    || !titleSelectionIsCurrent(selection) || titleCommit.value || plotAreaCommit.value || elementEditing.busy.value || objectActions.busy.value || anyPivotFieldBusy.value || drag.value) throw new Error('图表或权限已变更，请重新选择标题后编辑。')
  const patch = chartTitlePatch(value, selection.kind)
  if (!Object.keys(patch).length) return
  const live = FOffice.newAPI(props.studio).getWorkbook(selection.context.workbookId)?.getSheetBySheetId(selection.sheetId)?.getChartById(selection.chartId)
  if (!live) throw new Error('图表标题的所属工作表已关闭。')
  titleCommit.value = selection
  pointerError.value = ''
  try {
    // The public author command performs preparation, permission checking and
    // revision validation, then adds one document-scoped history entry.
    await live.update(patch)
    await nextTick()
    if (selectedTitle.value === selection && mounted && selection.studio === props.studio && selection.workbookId === props.workbook.id && selection.sheetId === props.activeSheetId) {
      const context = currentTitleContext(selection.chartId)
      if (context) selectedTitle.value = { ...selection, context, snapshot: context.snapshot }
      else closeTitleSelection()
      const current = sheetCharts.value.find(chart => chart.id === selection.chartId)
      if (current && (!current[selection.kind] || current.options?.titleStyle?.visible === false || current.options?.titleStyle?.position === 'hide')) closeTitleSelection(true)
      publishOverlayCommands()
    }
  } catch (error) {
    if (mounted && selection.studio === props.studio && selection.workbookId === props.workbook.id && selection.sheetId === props.activeSheetId
      && (!titleSelectionIsCurrent(selection) || selectedTitle.value !== selection || !titleMenu.value)) pointerError.value = error instanceof Error ? error.message : '标题修改失败。'
    throw error
  } finally {
    if (titleCommit.value === selection) titleCommit.value = null
    if (selectedTitle.value && !titleSelectionIsCurrent(selectedTitle.value)) closeTitleSelection()
  }
}
async function editTitle(event: MouseEvent, item: ChartDisplay): Promise<void> {
  const local = chartAxisPointerLocalPoint(pointerPoint(event), item.position, defaultChartView())
  const hit = chartRenderer.getTitleHit(item.chart.id, local)
  if (!hit) return
  event.preventDefault(); event.stopPropagation()
  if (drag.value) onPointerCancel()
  const selection = selectTitle(item.chart, hit.kind)
  await nextTick()
  if (mounted && selectedTitle.value === selection && canEditChart(item.chart) && selection.context) await titleMenu.value?.focusText(true)
}
async function editChartObject(event: MouseEvent, item: ChartDisplay): Promise<void> {
  const point = chartAxisPointerLocalPoint(pointerPoint(event), item.position, defaultChartView())
  if (chartRenderer.getTitleHit(item.chart.id, point)) { await editTitle(event, item); return }
  const legend = chartRenderer.getLegendHit(item.chart.id, point)
  const element = chartRenderer.getElementHit(item.chart.id, point, event.shiftKey)
  if (element) { event.preventDefault(); event.stopPropagation(); await openCanvasElement(item, element); return }
  if (legend) return
  if (chartRenderer.getPlotAreaHit(item.chart.id, point, 5 / zoomScale(), selectedPlotArea.value?.chartId === item.chart.id)) {
    event.preventDefault(); event.stopPropagation(); await openPlotArea(item); return
  }
  event.preventDefault(); event.stopPropagation(); selectWholeChart(item.chart.id)
  await objectActions.execute('edit', objectActions.capture(item.chart.id))
}
function enterChart(chartId: string): void {
  if (selectedPivotAxis.value?.pivotChartId === chartId && pivotAxis.openSelectedMenu()) return
  selectChart(chartId)
  if (selectedTitle.value?.chartId === chartId) void titleMenu.value?.focusText()
  if (selectedPlotArea.value?.chartId === chartId) void plotAreaMenu.value?.focusCoordinate()
  if (selectedElement.value?.chartId === chartId) void elementMenu.value?.focusFirst()
}
function selectPlotArea(chart: ChartModel): PlotAreaSelection | undefined {
  pivotAxis.close()
  if (!supportsChartPlotArea(chart.type) || !chartRenderer.getPlotAreaTarget(chart.id)) return undefined
  closeTitleSelection()
  closeElementSelection()
  const existing = selectedPlotArea.value
  if (existing?.chartId === chart.id && chartElementSelectionIsCurrent(existing)) return existing
  const selection: PlotAreaSelection = { chartId: chart.id, workbookId: props.workbook.id, sheetId: props.activeSheetId, studio: props.studio,
    snapshot: JSON.stringify(chart), context: currentTitleContext(chart.id), generation: ++plotAreaGeneration }
  selectedPlotArea.value = selection; selectedId.value = chart.id
  queueOverlayFrame()
  return selection
}
async function openPlotArea(item: ChartDisplay): Promise<void> {
  if (drag.value) onPointerCancel()
  const selection = selectPlotArea(item.chart)
  await nextTick()
  if (mounted && selectedPlotArea.value === selection) await plotAreaMenu.value?.focusCoordinate()
}
function closePlotAreaSelection(restoreFocus = false): void {
  const previous = selectedPlotArea.value
  selectedPlotArea.value = null; hoveredPlotArea.value = undefined
  if (drag.value?.mode === 'plot-area') onPointerCancel()
  plotAreaDrafts.value = {}
  queueOverlayFrame()
  if (restoreFocus && previous) void nextTick(() => {
    if (!mounted || selectedPlotArea.value || selectedTitle.value || selectedElement.value || previous.studio !== props.studio || previous.workbookId !== props.workbook.id || previous.sheetId !== props.activeSheetId) return
    const button = Array.from(root.value?.querySelectorAll<HTMLElement>('.chart-hitbox') ?? []).find(node => node.dataset.chartId === previous.chartId)
    button?.focus({ preventScroll: true })
  })
}
async function applyPlotArea(selection: PlotAreaSelection | null, area: ChartPlotAreaOptions | undefined): Promise<void> {
  const chart = selection && sheetCharts.value.find(chart => chart.id === selection.chartId)
  if (!selection || selection !== selectedPlotArea.value || !selection.context || !props.studio || !chart || !canEditChart(chart)
    || !supportsChartPlotArea(chart.type) || !chartElementSelectionIsCurrent(selection) || titleCommit.value || plotAreaCommit.value || elementEditing.busy.value || objectActions.busy.value || anyPivotFieldBusy.value || drag.value)
    throw new Error('图表或权限已变更，请重新选择绘图区后编辑。')
  validateChartPlotArea(area)
  if (chartPlotAreasEqual(chart.options?.plotArea, area)) return
  const live = FOffice.newAPI(props.studio).getWorkbook(selection.context.workbookId)?.getSheetBySheetId(selection.sheetId)?.getChartById(selection.chartId)
  if (!live) throw new Error('绘图区的所属工作表已关闭。')
  plotAreaCommit.value = selection; pointerError.value = ''
  try {
    await live.setOptions({ plotArea: area ? { ...area } : undefined })
    await nextTick()
    if (selectedPlotArea.value === selection && mounted && selection.studio === props.studio && selection.workbookId === props.workbook.id && selection.sheetId === props.activeSheetId) {
      const context = currentTitleContext(selection.chartId)
      if (context) selectedPlotArea.value = { ...selection, context, snapshot: context.snapshot }
      else closePlotAreaSelection()
      publishOverlayCommands()
    }
  } catch (error) {
    if (mounted && selection.studio === props.studio && selection.workbookId === props.workbook.id && selection.sheetId === props.activeSheetId
      && (!chartElementSelectionIsCurrent(selection) || selectedPlotArea.value !== selection || !plotAreaMenu.value)) pointerError.value = error instanceof Error ? error.message : '绘图区修改失败。'
    throw error
  } finally {
    if (plotAreaCommit.value === selection) plotAreaCommit.value = null
    if (selectedPlotArea.value && !chartElementSelectionIsCurrent(selectedPlotArea.value)) closePlotAreaSelection()
  }
}
function plotAreaOutlineCommands(): DrawCommand[] {
  const commands: DrawCommand[] = [], selected = selectedPlotArea.value?.chartId
  for (const chartId of new Set([hoveredPlotArea.value, selected])) {
    if (!chartId) continue
    const item = displayCharts.value.find(item => item.chart.id === chartId), target = chartRenderer.getPlotAreaTarget(chartId)
    if (item && target && commandIndexes.has(chartId)) commands.push(...chartDrawingOutlineCommands(item, chartId,
      chartPlotAreaOutlineCommands(target, { x: 0, y: 0 }, token.value.colorPrimary, chartId === selected ? 1.5 : 1)))
    if (item && target && chartId === selected && canEditChart(item.chart) && !titleCommit.value && !plotAreaCommit.value && !elementEditing.busy.value) {
      for (const handle of chartPlotAreaHandlePoints(target)) {
        const point = chartDrawingWorldPoint(item, handle)
        commands.push({ kind: 'rect', x: point.x - 3 / zoomScale(), y: point.y - 3 / zoomScale(), width: 6 / zoomScale(), height: 6 / zoomScale(),
          fill: token.value.colorBgContainer, stroke: token.value.colorPrimary, strokeWidth: 1 / zoomScale(),
          drawingRef: { kind: 'chart', id: chartId }, drawingOverlay: true })
      }
    }
  }
  return commands
}
function titleOutlineCommands(): DrawCommand[] {
  const commands: DrawCommand[] = []
  for (const selection of [hoveredTitle.value, selectedTitle.value]) {
    if (!selection || selection === hoveredTitle.value && selectedTitle.value?.chartId === selection.chartId && selectedTitle.value.kind === selection.kind) continue
    const item = displayCharts.value.find(item => item.chart.id === selection.chartId), target = chartRenderer.getTitleTarget(selection.chartId, selection.kind)
    if (item && target && commandIndexes.has(selection.chartId)) commands.push(...chartDrawingOutlineCommands(item, selection.chartId,
      chartTitleOutlineCommands(target, { x: 0, y: 0 }, token.value.colorPrimary, selection === selectedTitle.value ? 1.5 : 1)))
  }
  return commands
}
function closeElementSelection(restoreFocus = false): void {
  if (drag.value?.mode === 'element') onPointerCancel()
  hoveredElement.value = undefined
  elementEditing.close(restoreFocus)
}
function selectPivotCategory(item: ChartDisplay, target: PivotChartCanvasAxisTarget, openMenu: boolean, point?: ChartVector): boolean {
  if (target.label.level >= target.label.depth - 1) return false
  closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); clearChartPointer()
  const selection = pivotAxis.select(item.chart.id, target, openMenu)
  if (!selection) return false
  const bounds = chartDrawingViewportBounds(item, target.bounds, chartView(item.chart.id))
  const anchor = point ?? { x: bounds.x, y: bounds.y + bounds.height }
  pivotAxisMenuPoint.value = { x: Math.max(0, Math.min(anchor.x, (root.value?.clientWidth ?? anchor.x + 4) - 4)),
    y: Math.max(0, Math.min(anchor.y, (root.value?.clientHeight ?? anchor.y + 4) - 4)) }
  selectedId.value = item.chart.id; emit('select-chart', item.chart.id); queueOverlayFrame()
  return true
}
function openPivotAxisContext(event: MouseEvent, item: ChartDisplay): boolean {
  if (!mounted || !props.active || !commandIndexes.has(item.chart.id) || drag.value || pivotAxis.busy.value || anyPivotFieldBusy.value) return false
  const point = pointerPoint(event), local = chartDrawingLocalPoint(item, point, chartView(item.chart.id))
  const target = chartRenderer.getPivotAxisHit(item.chart.id, local)
  return !!target && selectPivotCategory(item, target, true, point)
}
function pivotAxisKeyboard(event: KeyboardEvent, item: ChartDisplay): boolean {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return false
  const selection = selectedPivotAxis.value
  if (selection?.pivotChartId !== item.chart.id) return false
  if (event.key === 'ContextMenu' || event.key === 'F10' && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey
    || event.key === 'Enter' && !event.ctrlKey && !event.metaKey && !event.altKey) {
    if (!pivotAxis.openSelectedMenu()) return false
    event.preventDefault(); event.stopPropagation(); return true
  }
  if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); pivotAxis.close(true); return true }
  if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key) && !event.ctrlKey && !event.metaKey && !event.altKey) {
    const targets = chartRenderer.getPivotAxisTargets(item.chart.id).filter(target => target.label.level < target.label.depth - 1)
    const index = targets.findIndex(target => target.key === selection.key)
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? targets.length - 1
      : Math.max(0, Math.min(targets.length - 1, index + (event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? -1 : 1)))
    const target = targets[next]
    if (target) selectPivotCategory(item, target, false)
    event.preventDefault(); event.stopPropagation(); return true
  }
  return false
}
function pivotAxisOutlineCommands(): DrawCommand[] {
  const targets = [...(hoveredPivotAxis.value ? [{ chartId: hoveredPivotAxis.value.chartId, target: hoveredPivotAxis.value.target, selected: false }] : []),
    ...(selectedPivotAxis.value ? [{ chartId: selectedPivotAxis.value.pivotChartId, target: selectedPivotAxis.value.target, selected: true }] : [])]
  return targets.flatMap(({ chartId, target, selected }) => {
    const item = displayCharts.value.find(item => item.chart.id === chartId), live = chartRenderer.getPivotAxisTarget(chartId, target.key)
    if (!item || !live || !commandIndexes.has(chartId)) return []
    return chartDrawingOutlineCommands(item, chartId, chartCanvasElementOutline(live, { x: 0, y: 0 }, token.value.colorPrimary, selected ? 1.5 : 1))
  })
}
function selectCanvasElement(item: ChartDisplay, target: ChartCanvasElementTarget): ChartCanvasElementSelection | undefined {
  pivotAxis.close()
  closeTitleSelection(); closePlotAreaSelection()
  const selection = elementEditing.select(item.chart.id, target)
  if (selection) { selectedId.value = item.chart.id; emit('select-chart', item.chart.id); queueOverlayFrame() }
  return selection
}
async function openCanvasElement(item: ChartDisplay, target: ChartCanvasElementTarget): Promise<void> {
  if (drag.value) onPointerCancel()
  const selection = selectCanvasElement(item, target)
  await nextTick()
  if (selection && mounted && selectedElement.value === selection && elementEditing.isCurrent(selection)) await elementMenu.value?.focusFirst()
}
function elementOutlineCommands(): DrawCommand[] {
  const commands: DrawCommand[] = [], selected = selectedElement.value, hovered = hoveredElement.value
  const elements = [...(hovered && !(selected?.chartId === hovered.chartId && selected.key === chartCanvasElementKey(hovered.element))
    ? [{ ...hovered, selected: false }] : []), ...(selected ? [{ chartId: selected.chartId, element: selected.element, selected: true }] : [])]
  for (const selection of elements) {
    const item = displayCharts.value.find(item => item.chart.id === selection.chartId), target = chartRenderer.getElementTarget(selection.chartId, selection.element)
    if (item && target && commandIndexes.has(selection.chartId)) commands.push(...chartDrawingOutlineCommands(item, selection.chartId,
      chartCanvasElementOutline(target, { x: 0, y: 0 }, token.value.colorPrimary, selection.selected ? 1.5 : 1)))
  }
  return commands
}
function hoverChart(event: PointerEvent, item: ChartDisplay): void {
  if (!mounted || drag.value || !root.value || !commandIndexes.has(item.chart.id)) return
  const local = chartAxisPointerLocalPoint(pointerPoint(event), item.position, defaultChartView())
  const pivotTarget = chartRenderer.getPivotAxisHit(item.chart.id, local)
  if (pivotTarget) {
    clearChartPointer(); hoveredPivotAxis.value = { chartId: item.chart.id, target: pivotTarget }; queueOverlayFrame(); return
  }
  if (hoveredPivotAxis.value) { hoveredPivotAxis.value = undefined; queueOverlayFrame() }
  const titleHit = chartRenderer.getTitleHit(item.chart.id, local)
  if (titleHit) {
    clearChartPointer()
    hoveredTitle.value = { chartId: item.chart.id, kind: titleHit.kind }
    queueOverlayFrame()
    return
  }
  if (hoveredTitle.value) { hoveredTitle.value = undefined; queueOverlayFrame() }
  const element = chartRenderer.getElementHit(item.chart.id, local, event.shiftKey)
  if (element) {
    clearChartPointer()
    hoveredElement.value = { chartId: item.chart.id, element: element.element }
    queueOverlayFrame(); return
  }
  if (hoveredElement.value) { hoveredElement.value = undefined; queueOverlayFrame() }
  const plotAreaHit = !chartRenderer.getLegendHit(item.chart.id, local) ? chartRenderer.getPlotAreaHit(item.chart.id, local, 5 / zoomScale(), selectedPlotArea.value?.chartId === item.chart.id) : undefined
  if (plotAreaHit) {
    clearChartPointer(); hoveredPlotArea.value = item.chart.id; queueOverlayFrame(); return
  }
  if (hoveredPlotArea.value) { hoveredPlotArea.value = undefined; queueOverlayFrame() }
  desiredPointer = { chartId: item.chart.id, point: local }
  pointerDirty = true
  queueOverlayFrame()
}

function clearChartPointer(chartId?: string): void {
  if (hoveredPivotAxis.value && (chartId === undefined || hoveredPivotAxis.value.chartId === chartId)) { hoveredPivotAxis.value = undefined; queueOverlayFrame() }
  if (hoveredTitle.value && (chartId === undefined || hoveredTitle.value.chartId === chartId)) { hoveredTitle.value = undefined; queueOverlayFrame() }
  if (hoveredPlotArea.value && (chartId === undefined || hoveredPlotArea.value === chartId)) { hoveredPlotArea.value = undefined; queueOverlayFrame() }
  if (hoveredElement.value && (chartId === undefined || hoveredElement.value.chartId === chartId)) { hoveredElement.value = undefined; queueOverlayFrame() }
  if (chartId !== undefined && desiredPointer?.chartId !== chartId && activePointerId !== chartId && contentTooltip.value?.chartId !== chartId) return
  contentTooltip.value = undefined
  if (!desiredPointer && !activePointerId && !pointerDirty) return
  desiredPointer = undefined
  pointerDirty = true
  queueOverlayFrame()
}

function updatePointerFrame(chartId: string, frame: EChartsCanvasFrame | undefined): void {
  const index = commandIndexes.get(chartId), command = index === undefined ? undefined : contentCommands[index]
  if (!frame || index === undefined || command?.kind !== 'image') return
  contentCommands[index] = { ...command, source: frame.source, sourceVersion: frame.sourceVersion }
  renderedFrames.value = { ...renderedFrames.value, [chartId]: { version: frame.sourceVersion, width: frame.source.width, height: frame.source.height } }
}

function updatePointerFrames(): void {
  if (desiredPointer && !commandIndexes.has(desiredPointer.chartId)) desiredPointer = undefined
  if (activePointerId && activePointerId !== desiredPointer?.chartId) {
    updatePointerFrame(activePointerId, chartRenderer.updatePointer(activePointerId))
    activePointerId = undefined
  }
  if (desiredPointer && commandIndexes.has(desiredPointer.chartId)) {
    updatePointerFrame(desiredPointer.chartId, chartRenderer.updatePointer(desiredPointer.chartId, desiredPointer.point))
    activePointerId = desiredPointer.chartId
    const content = chartRenderer.getContentTooltip(desiredPointer.chartId, desiredPointer.point)
    contentTooltip.value = content ? { chartId: desiredPointer.chartId, point: { ...desiredPointer.point }, content } : undefined
    pointerError.value = ''
  } else contentTooltip.value = undefined
}

function removeChart(event: KeyboardEvent, chartId: string): void {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || event.ctrlKey || event.metaKey || event.altKey) return
  if (event.key !== 'Delete' && event.key !== 'Backspace') return
  event.preventDefault(); event.stopPropagation()
  const chart = sheetCharts.value.find(chart => chart.id === chartId)
  if (!chart || !canEditChart(chart)) return
  if (selectedElement.value?.chartId === chartId) { void elementEditing.remove(); return }
  if (selectedTitle.value?.chartId === chartId) {
    event.preventDefault(); event.stopPropagation()
    void applyTitleSpec(selectedTitle.value, undefined).catch(error => { pointerError.value = error instanceof Error ? error.message : '删除标题失败。' })
    return
  }
  if (selectedPlotArea.value?.chartId === chartId) {
    const selection = { ...selectedPlotArea.value, generation: ++plotAreaGeneration }
    selectedPlotArea.value = selection
    void applyPlotArea(selection, undefined).catch(error => { pointerError.value = error instanceof Error ? error.message : '恢复绘图区失败。' })
    return
  }
  if (props.studio) void objectActions.execute('delete', objectActions.capture(chartId))
  else { emit('remove-chart', chartId); clearChartPointer(chartId); selectedId.value = null; emit('select-chart', null) }
}

function beginDrag(event: PointerEvent, item: ChartDisplay, mode: 'move' | 'resize' | 'plot-area', handle?: ChartPlotAreaResizeHandle): void {
  if (event.button !== 0 || !mounted || !props.active || drag.value || titleCommit.value || plotAreaCommit.value || elementEditing.busy.value || objectActions.busy.value || pivotAxis.busy.value || anyPivotFieldBusy.value) return
  if (mode !== 'move' && !canEditChart(item.chart)) return
  event.preventDefault(); event.stopPropagation()
  clearChartPointer()
  selectedId.value = item.chart.id
  emit('select-chart', item.chart.id)
  const point = pointerPoint(event)
  const target = event.currentTarget instanceof HTMLElement ? event.currentTarget : null
  let dragMode: DragState['mode'] = mode
  let legendHit: ChartCanvasLegendHit | undefined
  let titleHit: ChartCanvasTitleTarget | undefined
  let titleSelection: TitleSelection | undefined
  let plotAreaHit: ChartCanvasPlotAreaTarget | undefined
  let plotAreaSelection: PlotAreaSelection | undefined
  let elementSelection: ChartCanvasElementSelection | undefined
  if (mode === 'plot-area') {
    plotAreaHit = chartRenderer.getPlotAreaTarget(item.chart.id)
    plotAreaSelection = selectPlotArea(item.chart)
  }
  if (mode === 'move') {
    try {
      // Remove transient pointer graphics before querying native z-order.
      // The pointer is view state and must not shield an authored mark.
      updatePointerFrames()
      const local = chartAxisPointerLocalPoint(point, item.position,
        defaultChartView())
      const pivotTarget = chartRenderer.getPivotAxisHit(item.chart.id, local)
      if (pivotTarget && selectPivotCategory(item, pivotTarget, false)) return
      pivotAxis.close()
      titleHit = chartRenderer.getTitleHit(item.chart.id, local)
      if (titleHit) {
        dragMode = 'title'
        titleSelection = selectTitle(item.chart, titleHit.kind)
      } else {
        closeTitleSelection(); legendHit = chartRenderer.getLegendHit(item.chart.id, local)
        if (!legendHit) {
          const element = chartRenderer.getElementHit(item.chart.id, local, event.shiftKey)
          if (element) { dragMode = 'element'; elementSelection = selectCanvasElement(item, element) }
          else {
            closeElementSelection()
            plotAreaHit = chartRenderer.getPlotAreaHit(item.chart.id, local, 5 / zoomScale(), selectedPlotArea.value?.chartId === item.chart.id)
            if (plotAreaHit) { dragMode = 'plot-area'; plotAreaSelection = selectPlotArea(item.chart) }
            else closePlotAreaSelection()
          }
        }
      }
      if (legendHit) { closePlotAreaSelection(); closeElementSelection(); dragMode = 'legend' }
      else if (!titleHit && !plotAreaHit && dragMode !== 'element') {
        closeTitleSelection()
        const content = chartRenderer.beginContentDrag(item.chart.id, local)
        if (content) { dragMode = content.kind; updatePointerFrame(item.chart.id, content.frame); queueOverlayFrame() }
      }
    } catch (error) {
      pointerError.value = error instanceof Error ? error.message : '图表内容拖动失败'
      return
    }
  }
  if (mode === 'resize') { closeTitleSelection(); closePlotAreaSelection(); closeElementSelection() }
  if (dragMode === 'title' && (!titleSelection?.context || titleCommit.value || plotAreaCommit.value)) return
  if (dragMode === 'plot-area' && (!plotAreaHit || !plotAreaSelection?.context || titleCommit.value || plotAreaCommit.value)) return
  if (dragMode === 'element' && !elementSelection) return
  if (dragMode !== 'sankey-node' && dragMode !== 'legend' && dragMode !== 'element' && !canEditChart(item.chart)) return
  const workbookId = props.workbook.id
  const frame = item.drawingGeometry ? drawingAffineFrame(item.drawingGeometry) : undefined
  drag.value = {
    chartId: item.chart.id,
    mode: dragMode,
    ...(legendHit ? { legendHit } : {}),
    ...(titleHit ? { titleHit, titleSelection } : {}),
    ...(plotAreaHit ? { plotAreaHit, plotAreaSelection, plotAreaHandle: handle } : {}),
    ...(elementSelection ? { elementSelection } : {}),
    pointerId: event.pointerId,
    target,
    workbookId,
    studio: props.studio,
    sheetId: props.activeSheetId,
    ...(props.studio && workbookId ? { modelRevision: props.studio.getWorkbookRevision(workbookId) } : {}),
    chartSnapshot: JSON.stringify(item.chart),
    originX: point.x,
    originY: point.y,
    position: frame ? { x: GRID_HEADER_WIDTH + frame.left, y: GRID_HEADER_HEIGHT + frame.top } : { ...item.position },
    size: frame ? { width: frame.width, height: frame.height } : { ...item.size },
    contentSize: { ...item.size },
    pane: { ...chartView(item.chart.id) },
  }
  if (target && typeof target.setPointerCapture === 'function') {
    try { target.setPointerCapture(event.pointerId) }
    catch (error) {
      onPointerCancel()
      pointerError.value = error instanceof Error ? error.message : '图表拖动指针捕获失败'
    }
  }
}

function dragIsCurrent(active: DragState): boolean {
  const chart = sheetCharts.value.find(chart => chart.id === active.chartId)
  if (active.modelRevision !== undefined && (!props.studio || !active.workbookId || !props.studio.getWorkbookModel(active.workbookId))) return false
  return props.active && active.studio === props.studio && active.workbookId === props.workbook.id && active.sheetId === props.activeSheetId
    && (active.modelRevision === undefined || !props.studio || !active.workbookId || props.studio.getWorkbookRevision(active.workbookId) === active.modelRevision)
    && JSON.stringify(chart) === active.chartSnapshot
    && !!chart && (active.mode === 'sankey-node' || active.mode === 'legend' || active.mode === 'element' || canEditChart(chart))
}

function releaseDragPointer(active: DragState): void {
  if (!active.target?.hasPointerCapture?.(active.pointerId)) return
  try { active.target.releasePointerCapture(active.pointerId) }
  catch (error) { pointerError.value = error instanceof Error ? error.message : '图表拖动指针释放失败' }
}

function flushContentDrag(): void {
  const desired = desiredContentDrag, active = drag.value
  desiredContentDrag = undefined
  if (!desired || !active || active.mode !== 'sankey-node' || desired.chartId !== active.chartId || desired.pointerId !== active.pointerId) return
  if (!dragIsCurrent(active)) { onPointerCancel(); return }
  updatePointerFrame(active.chartId, chartRenderer.updateContentDrag(active.chartId, desired.point))
}

function onPointerMove(event: PointerEvent): void {
  if (!mounted) return
  const active = drag.value
  if (!active || event.pointerId !== active.pointerId) return
  if (!dragIsCurrent(active)) { onPointerCancel(); return }
  const point = pointerPoint(event)
  if (active.mode === 'legend' || active.mode === 'element') return
  if (active.mode === 'sankey-node') {
    desiredContentDrag = { chartId: active.chartId, pointerId: active.pointerId, point: chartAxisPointerLocalPoint(point, active.position,
      defaultChartView()) }
    queueOverlayFrame()
    return
  }
  const scale = zoomScale()
  const dx = (point.x - active.originX) / scale
  const dy = (point.y - active.originY) / scale
  const item = displayCharts.value.find(item => item.chart.id === active.chartId)
  const contentDelta = item ? chartDrawingLocalDelta({ ...item, size: active.contentSize, drawingGeometry: item.chart.drawingGeometry },
    { x: active.originX, y: active.originY }, point, active.pane) : { x: dx, y: dy }
  if (active.mode === 'title') {
    if (!active.titleHit || !active.titleSelection || Math.hypot(point.x - active.originX, point.y - active.originY) <= 4 && !titleDrafts.value[active.chartId]) return
    if (!titleDrafts.value[active.chartId]) {
      // Beginning a real drag ends the floating input session, even for the
      // same title. Delayed input events keep the retired selection token.
      const selection = { ...active.titleSelection, generation: ++titleGeneration }
      selectedTitle.value = selection
      active.titleSelection = selection
    }
    titleDrafts.value = { ...titleDrafts.value, [active.chartId]: { kind: active.titleHit.kind,
      value: { manualLayout: chartTitleDragLayout(active.titleHit, active.contentSize, contentDelta) } } }
    publishOverlayCommands()
    return
  }
  if (active.mode === 'plot-area') {
    if (!active.plotAreaHit || !active.plotAreaSelection || Math.hypot(point.x - active.originX, point.y - active.originY) <= 4 && !plotAreaDrafts.value[active.chartId]) return
    if (!plotAreaDrafts.value[active.chartId]) {
      const selection = { ...active.plotAreaSelection, generation: ++plotAreaGeneration }
      selectedPlotArea.value = selection; active.plotAreaSelection = selection
    }
    plotAreaDrafts.value = { ...plotAreaDrafts.value, [active.chartId]: chartPlotAreaDragLayout(active.plotAreaHit, active.contentSize, contentDelta, active.plotAreaHandle) }
    publishOverlayCommands()
    return
  }
  const position = active.mode === 'move'
    ? { x: Math.max(0, active.position.x + dx), y: Math.max(0, active.position.y + dy) }
    : active.position
  const size = active.mode === 'resize'
    ? { width: Math.max(160, active.size.width + contentDelta.x * active.size.width / active.contentSize.width),
      height: Math.max(120, active.size.height + contentDelta.y * active.size.height / active.contentSize.height) }
    : active.contentSize
  const sourceGeometry = item?.chart.drawingGeometry, sourceFrame = sourceGeometry ? drawingAffineFrame(sourceGeometry) : undefined
  const drawingGeometry = sourceGeometry && sourceFrame ? drawingAffineWithFrame(sourceGeometry, {
    left: position.x - GRID_HEADER_WIDTH, top: position.y - GRID_HEADER_HEIGHT,
    width: active.mode === 'resize' ? size.width : sourceFrame.width, height: active.mode === 'resize' ? size.height : sourceFrame.height,
  }) : undefined
  drafts.value = { ...drafts.value, [active.chartId]: { position, size, drawingGeometry } }
}

function onPointerUp(event: PointerEvent): void {
  const active = drag.value
  if (!active || event.pointerId !== active.pointerId) return
  if (!dragIsCurrent(active)) { onPointerCancel(); return }
  if (active.mode === 'element') {
    drag.value = null; releaseDragPointer(active)
    const selection = active.elementSelection, point = pointerPoint(event)
    if (selection && selectedElement.value === selection && elementEditing.isCurrent(selection)
      && Math.hypot(point.x - active.originX, point.y - active.originY) <= 4) void elementMenu.value?.focusFirst()
    queueOverlayFrame(); return
  }
  if (active.mode === 'legend') {
    const point = pointerPoint(event), local = chartAxisPointerLocalPoint(point, active.position,
      defaultChartView())
    drag.value = null
    releaseDragPointer(active)
    if (Math.hypot(point.x - active.originX, point.y - active.originY) > 4
      || JSON.stringify(chartRenderer.getLegendHit(active.chartId, local)) !== JSON.stringify(active.legendHit)) return
    try {
      updatePointerFrame(active.chartId, chartRenderer.activateLegend(active.chartId, local))
      desiredPointer = { chartId: active.chartId, point: local }; pointerDirty = true
      queueOverlayFrame()
    } catch (error) {
      publishOverlayCommands()
      pointerError.value = error instanceof Error ? error.message : '图例交互失败'
    }
    return
  }
  if (active.mode === 'sankey-node') {
    desiredContentDrag = { chartId: active.chartId, pointerId: active.pointerId, point: chartAxisPointerLocalPoint(pointerPoint(event), active.position,
      defaultChartView()) }
    try { flushContentDrag() }
    catch (error) { onPointerCancel(); pointerError.value = error instanceof Error ? error.message : '图表内容拖动失败'; return }
    if (drag.value !== active) return
  } else onPointerMove(event)
  const draft = drafts.value[active.chartId]
  const titleDraft = titleDrafts.value[active.chartId]
  const plotAreaDraft = plotAreaDrafts.value[active.chartId]
  const next = { ...drafts.value }
  delete next[active.chartId]
  drafts.value = next
  if (titleDraft && active.mode !== 'title') { const titles = { ...titleDrafts.value }; delete titles[active.chartId]; titleDrafts.value = titles; publishOverlayCommands() }
  if (plotAreaDraft && active.mode !== 'plot-area') { const plots = { ...plotAreaDrafts.value }; delete plots[active.chartId]; plotAreaDrafts.value = plots; publishOverlayCommands() }
  drag.value = null
  releaseDragPointer(active)
  if (active.mode === 'title') {
    if (titleDraft && active.titleSelection) {
      const original = sheetCharts.value.find(chart => chart.id === active.chartId)?.options?.[titleDraft.kind === 'title' ? 'titleStyle' : 'subtitleStyle']?.manualLayout
      const clearDraft = () => {
        if (titleDrafts.value[active.chartId] === titleDraft) { const titles = { ...titleDrafts.value }; delete titles[active.chartId]; titleDrafts.value = titles }
        publishOverlayCommands()
      }
      if (JSON.stringify(original) !== JSON.stringify(titleDraft.value.manualLayout)) void applyTitleSpec(active.titleSelection, titleDraft.value)
        .catch(error => { pointerError.value = error instanceof Error ? error.message : '标题拖动提交失败。' }).finally(clearDraft)
      else clearDraft()
    }
  } else if (active.mode === 'plot-area') {
    if (plotAreaDraft && active.plotAreaSelection) {
      const original = sheetCharts.value.find(chart => chart.id === active.chartId)?.options?.plotArea
      const clearDraft = () => {
        if (plotAreaDrafts.value[active.chartId] === plotAreaDraft) { const plots = { ...plotAreaDrafts.value }; delete plots[active.chartId]; plotAreaDrafts.value = plots }
        publishOverlayCommands()
      }
      if (!chartPlotAreasEqual(original, plotAreaDraft)) void applyPlotArea(active.plotAreaSelection, plotAreaDraft)
        .catch(error => { pointerError.value = error instanceof Error ? error.message : '绘图区拖动提交失败。' }).finally(clearDraft)
      else clearDraft()
    }
  } else if (active.mode === 'sankey-node') {
    try { updatePointerFrame(active.chartId, chartRenderer.endContentDrag(active.chartId)); queueOverlayFrame() }
    catch (error) { pointerError.value = error instanceof Error ? error.message : '图表内容拖动结束失败' }
  } else if (draft) emit('update-chart', { chartId: active.chartId,
    patch: active.mode === 'resize' ? { size: draft.size } : { position: draft.position } })
}

function onPointerCancel(event?: Event): void {
  const active = drag.value
  if (event instanceof PointerEvent && active && event.pointerId !== active.pointerId) return
  desiredContentDrag = undefined
  clearChartPointer()
  if (!active) return
  const next = { ...drafts.value }
  delete next[active.chartId]
  drafts.value = next
  if (titleDrafts.value[active.chartId]) { const titles = { ...titleDrafts.value }; delete titles[active.chartId]; titleDrafts.value = titles; publishOverlayCommands() }
  if (plotAreaDrafts.value[active.chartId]) { const plots = { ...plotAreaDrafts.value }; delete plots[active.chartId]; plotAreaDrafts.value = plots; publishOverlayCommands() }
  drag.value = null
  releaseDragPointer(active)
  if (active.mode === 'sankey-node') {
    try { updatePointerFrame(active.chartId, chartRenderer.endContentDrag(active.chartId, true)); queueOverlayFrame() }
    catch (error) { pointerError.value = error instanceof Error ? error.message : '图表内容拖动取消失败' }
  }
}

function onDragKeyDown(event: KeyboardEvent): void {
  if (event.key !== 'Escape' || event.isComposing || event.keyCode === 229 || event.defaultPrevented) return
  if (!drag.value) {
    if (contentTooltip.value) clearChartPointer()
    if (selectedTitle.value) { event.preventDefault(); event.stopPropagation(); closeTitleSelection(true) }
    if (selectedPlotArea.value) { event.preventDefault(); event.stopPropagation(); closePlotAreaSelection(true) }
    if (selectedElement.value) { event.preventDefault(); event.stopPropagation(); closeElementSelection(true) }
    return
  }
  event.preventDefault(); event.stopPropagation()
  onPointerCancel()
  if (selectedTitle.value) closeTitleSelection(true)
  if (selectedPlotArea.value) closePlotAreaSelection(true)
  if (selectedElement.value) closeElementSelection(true)
}

function resizeChartWithKeyboard(event: KeyboardEvent, item: ChartDisplay): void {
  if (!mounted || !props.active || event.defaultPrevented || event.isComposing || event.keyCode === 229
    || event.ctrlKey || event.metaKey || event.altKey) return
  if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return
  event.preventDefault(); event.stopPropagation()
  if (!canEditChart(item.chart) || titleCommit.value || plotAreaCommit.value || elementEditing.busy.value || objectActions.busy.value || anyPivotFieldBusy.value) return
  if (drag.value) onPointerCancel()
  closeTitleSelection()
  closePlotAreaSelection()
  closeElementSelection()
  const step = event.shiftKey ? 10 : 1
  const width = Math.max(160, item.size.width + (event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0))
  const height = Math.max(120, item.size.height + (event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0))
  if (width !== item.chart.size.width || height !== item.chart.size.height) emit('update-chart', { chartId: item.chart.id, patch: { size: { width, height } } })
}
function editPlotAreaWithKeyboard(event: KeyboardEvent, item: ChartDisplay, handle?: ChartPlotAreaResizeHandle): void {
  const selection = selectedPlotArea.value
  if (!mounted || !props.active || event.defaultPrevented || event.isComposing || event.keyCode === 229
    || event.ctrlKey || event.metaKey || event.altKey || selection?.chartId !== item.chart.id
    || !['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return
  event.preventDefault(); event.stopPropagation()
  const target = chartRenderer.getPlotAreaTarget(item.chart.id)
  if (!selection || !target || !canEditChart(item.chart) || !selection.context || drag.value || titleCommit.value || plotAreaCommit.value || elementEditing.busy.value || objectActions.busy.value || anyPivotFieldBusy.value) return
  const step = event.shiftKey ? 10 : 1
  const area = chartPlotAreaDragLayout({ ...target, area: item.chart.options?.plotArea ?? target.area }, item.size, { x: event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0,
    y: event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0 }, handle)
  if (!chartPlotAreasEqual(item.chart.options?.plotArea, area)) {
    const next = { ...selection, generation: ++plotAreaGeneration }
    selectedPlotArea.value = next
    void applyPlotArea(next, area).catch(error => { pointerError.value = error instanceof Error ? error.message : '绘图区键盘修改失败。' })
  }
}
function chartKeyDown(event: KeyboardEvent, item: ChartDisplay): void {
  if (!mounted || !props.active || event.defaultPrevented || event.isComposing || event.keyCode === 229 || isTextInput(event.target)) return
  if (pivotAxisKeyboard(event, item)) return
  selectChart(item.chart.id)
  removeChart(event, item.chart.id)
  if (selectedElement.value?.chartId === item.chart.id) {
    if (event.key === 'F2' && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey && !event.defaultPrevented) {
      event.preventDefault(); event.stopPropagation(); void elementEditing.more()
    }
    return
  }
  editPlotAreaWithKeyboard(event, item)
  if (event.defaultPrevented || selectedTitle.value || selectedPlotArea.value || drag.value || titleCommit.value || plotAreaCommit.value || elementEditing.busy.value || anyPivotFieldBusy.value) return
  if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key) && !event.ctrlKey && !event.metaKey && !event.altKey) {
    event.preventDefault(); event.stopPropagation()
    if (!canEditChart(item.chart) || objectActions.busy.value) return
    const direction = sheetDrawingMoveDirectionForKey(event.key)
    if (direction !== undefined) emit('drawing-command', SHEET_DRAWING_GROUP_COMMAND_IDS.move, {
      sheetId: props.activeSheetId, selection: [{ kind: 'chart', id: item.chart.id }],
      expectedSelection: props.selectedDrawingRefs ? [...props.selectedDrawingRefs] : undefined,
      expectedRevision: props.studio && props.workbook.id ? props.studio.getWorkbookRevision(props.workbook.id) : undefined,
      direction, step: event.shiftKey ? 10 : 1,
    })
    return
  }
  const command = chartObjectShortcut(event)
  if (!command) return
  event.preventDefault(); event.stopPropagation()
  if (command === PASTE_COPIED_CHART_COMMAND) void objectActions.paste()
  else {
    const action = command === CHART_OBJECT_COMMANDS.copy ? 'copy' : 'edit'
    selectChart(item.chart.id); void objectActions.execute(action, objectActions.capture(item.chart.id))
  }
}
function isTextInput(target: EventTarget | null): boolean {
  return target instanceof Element && !!target.closest('input, textarea, [contenteditable="true"], [role="textbox"]')
}
function routeNativeDrawingEvent(event: MouseEvent, item: ChartDisplay): void {
  if (routedDispatch || !(event.target instanceof Element) || !event.target.closest('.chart-hitbox, .chart-resize-handle')) return
  if (event.type === 'pointerdown') routedClickChartId = undefined
  if (event.type === 'click' && routedClickChartId === item.chart.id) {
    routedClickChartId = undefined; event.preventDefault(); event.stopPropagation(); return
  }
  clearChartPointer(item.chart.id)
  if (props.routeDrawingEvent?.(event, item.chart.id)) {
    if (event.type === 'pointerdown') routedClickChartId = item.chart.id
    return
  }
  authorPanes.set(item.chart.id, defaultChartView())
}

/** Frozen pane hits use the same author controller as the library-owned input bridge. */
function routeCanvasEvent(event: MouseEvent, chartId: string, pane: Pick<ChartDrawingViewport, 'scrollLeft' | 'scrollTop' | 'headerWidth' | 'headerHeight'>): void {
  const item = displayCharts.value.find(item => item.chart.id === chartId)
  if (!mounted || !props.active || !item) return
  routedPane = { ...defaultChartView(), ...pane }; authorPanes.set(chartId, { ...routedPane })
  try {
    if (event.type === 'pointerdown' && event instanceof PointerEvent) {
      const target = selectedPlotArea.value?.chartId === chartId ? chartRenderer.getPlotAreaTarget(chartId) : undefined,
        point = pointerPoint(event), handle = target ? chartPlotAreaHandlePoints(target).find(handle => {
          const projected = chartDrawingViewportPoint(item, handle, routedPane!)
          return Math.hypot(point.x - projected.x, point.y - projected.y) <= 6
        }) : undefined
      beginDrag(event, item, handle ? 'plot-area' : 'move', handle?.handle)
    }
    else if (event.type === 'pointermove' && event instanceof PointerEvent) hoverChart(event, item)
    else if (event.type === 'dblclick') editChartObject(event, item)
    else if (event.type === 'wheel' && event instanceof WheelEvent) scaleRelationContent(event, item)
    else if (event.type === 'contextmenu') {
      const hitbox = Array.from(root.value?.querySelectorAll<HTMLElement>('.chart-hitbox') ?? []).find(node => node.dataset.chartId === chartId)
      if (hitbox) {
        routedDispatch = true
        try { hitbox.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true,
          clientX: event.clientX, clientY: event.clientY, button: event.button, buttons: event.buttons,
          ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, altKey: event.altKey })) }
        finally { routedDispatch = false }
      }
    }
  } finally { routedPane = undefined }
}
function routeCanvasKeydown(event: KeyboardEvent, chartId: string): boolean {
  const item = displayCharts.value.find(item => item.chart.id === chartId)
  if (!mounted || !props.active || !item || event.defaultPrevented || event.isComposing || event.keyCode === 229) return false
  if (event.key === 'ContextMenu' || event.key === 'F10' && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey) {
    const hitbox = Array.from(root.value?.querySelectorAll<HTMLElement>('.chart-hitbox') ?? []).find(node => node.dataset.chartId === chartId)
    if (hitbox) {
      hitbox.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: event.key, code: event.code,
        ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, altKey: event.altKey }))
      event.preventDefault(); event.stopPropagation()
    }
  } else if (event.key === 'Enter' && !event.ctrlKey && !event.metaKey && !event.altKey) {
    event.preventDefault(); event.stopPropagation(); enterChart(chartId)
  } else chartKeyDown(event, item)
  return event.defaultPrevented
}
function onChartObjectPaste(event: ClipboardEvent): void {
  const overlay = root.value
  if (!objectHostActive.value || !props.active || event.defaultPrevented || isTextInput(event.target)
    || !overlay || !(event.target instanceof Node) || !overlay.parentElement?.contains(event.target)) return
  if (drag.value || titleCommit.value || plotAreaCommit.value || elementEditing.busy.value || anyPivotFieldBusy.value) return
  if (overlay.contains(event.target) && (selectedTitle.value || selectedPlotArea.value || selectedElement.value)) return
  closeTitleSelection(); closePlotAreaSelection(); closeElementSelection()
  objectActions.handlePasteEvent(event)
}

function onScroll(): void {
  pivotAxis.closeMenu()
  if (!scroller) return
  onPointerCancel()
  closeTitleSelection()
  closePlotAreaSelection()
  closeElementSelection()
  authorPanes.clear()
  scrollLeft.value = scroller.scrollLeft
  scrollTop.value = scroller.scrollTop
}
function onWindowBlur(): void { onPointerCancel(); closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); pivotAxis.close(); objectActions.cancel() }
function onViewportResize(): void { onPointerCancel(); closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); pivotAxis.close(); publishOverlayCommands() }
function onOutsidePointerDown(event: PointerEvent): void {
  if (props.selectedDrawingRefs !== undefined && event.target instanceof Element && event.target.closest('.canvas-scroll') && root.value?.parentElement?.contains(event.target)) return
  if (selectedTitle.value && event.target instanceof Node && !root.value?.contains(event.target)) closeTitleSelection()
  if (selectedPlotArea.value && event.target instanceof Node && !root.value?.contains(event.target)) closePlotAreaSelection()
  if (selectedElement.value && event.target instanceof Node && !root.value?.contains(event.target)) closeElementSelection()
  if (selectedPivotAxis.value && !pivotAxis.menuOpen.value && event.target instanceof Node && !root.value?.contains(event.target)) pivotAxis.close()
}

function publishOverlayCommands(): void {
  contentDirty = true
  queueOverlayFrame()
}

function queueOverlayFrame(): void {
  if (!mounted || frameRequest !== undefined) return
  frameRequest = window.requestAnimationFrame(() => {
    frameRequest = undefined
    if (!mounted) return
    const renderContent = contentDirty
    contentDirty = false
    if (renderContent) contentCommands = renderCommands()
    try { flushContentDrag() }
    catch (error) { onPointerCancel(); pointerError.value = error instanceof Error ? error.message : '图表内容拖动失败' }
    if (pointerDirty || renderContent) {
      pointerDirty = false
      try { updatePointerFrames() }
      catch (reason) {
        pointerError.value = reason instanceof Error ? reason.message : '图表指示器绘制失败'
        desiredPointer = undefined
        activePointerId = undefined
        contentTooltip.value = undefined
        chartRenderer.invalidate()
        publishOverlayCommands()
      }
    }
    if (selectedElement.value && !elementEditing.busy.value && !chartRenderer.getElementTarget(selectedElement.value.chartId, selectedElement.value.element)) closeElementSelection()
    pivotAxis.reconcile()
    emit('overlay-commands', [...contentCommands, ...plotAreaOutlineCommands(), ...titleOutlineCommands(), ...elementOutlineCommands(), ...pivotAxisOutlineCommands()])
  })
}
function watchPixelRatio(): void {
  onPointerCancel()
  closeTitleSelection()
  closePlotAreaSelection()
  closeElementSelection()
  pivotAxis.close()
  pixelRatioQuery?.removeEventListener('change', watchPixelRatio)
  pixelRatioQuery = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`)
  pixelRatioQuery.addEventListener('change', watchPixelRatio)
  publishOverlayCommands()
}
function fontsChanged(): void {
  onPointerCancel()
  closeTitleSelection()
  closePlotAreaSelection()
  closeElementSelection()
  pivotAxis.close()
  clearChartPointer()
  chartRenderer.invalidate()
  publishOverlayCommands()
}

watch([displayCharts, titleDrafts, plotAreaDrafts, chartTheme, () => props.zoom], publishOverlayCommands, { deep: true, flush: 'post' })
watch(statistics, publishOverlayCommands, { flush: 'post' })
watch([() => props.workbook, () => props.studio, () => props.zoom, chartTheme, statistics], () => {
  if (drag.value) onPointerCancel()
  clearChartPointer()
  if (selectedTitle.value && !titleCommit.value && !titleSelectionIsCurrent(selectedTitle.value)) closeTitleSelection()
  if (selectedPlotArea.value && !plotAreaCommit.value && !chartElementSelectionIsCurrent(selectedPlotArea.value)) closePlotAreaSelection()
  pivotAxis.reconcile()
}, { deep: true, flush: 'sync' })
watch([() => props.workbook.id, () => props.studio, () => props.zoom, chartTheme], () => { closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); pivotAxis.close() }, { deep: true, flush: 'sync' })
watch(() => props.active, active => {
  if (!active) { onPointerCancel(); closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); pivotAxis.close(); objectActions.cancel(); authorPanes.clear() }
}, { flush: 'sync' })
watch(() => props.selectedDrawingRefs, refs => {
  if (refs === undefined) return
  const next = refs.length === 1 && refs[0]?.kind === 'chart' ? refs[0].id : null
  if (next !== selectedId.value) {
    onPointerCancel(); closeTitleSelection(); closePlotAreaSelection(); closeElementSelection(); pivotAxis.close(); objectActions.cancel()
    selectedId.value = next
  }
}, { deep: true, immediate: true, flush: 'sync' })
watch(() => props.activeSheetId, () => {
  pivotFieldBusy.value = {}
  onPointerCancel()
  closeTitleSelection()
  closePlotAreaSelection()
  closeElementSelection()
  pivotAxis.close()
  pointerError.value = ''
  authorPanes.clear(); routedClickChartId = undefined
  const refs = props.selectedDrawingRefs
  selectedId.value = refs?.length === 1 && refs[0]?.kind === 'chart' ? refs[0].id : null
  if (refs === undefined) emit('select-chart', null)
  drag.value = null
  drafts.value = {}
  publishOverlayCommands()
})
watch(sheetCharts, charts => {
  pivotFieldBusy.value = Object.fromEntries(Object.entries(pivotFieldBusy.value).filter(([id, busy]) => busy && charts.some(chart => chart.id === id)))
  if (selectedId.value && !charts.some(chart => chart.id === selectedId.value)) {
    selectedId.value = null; if (props.selectedDrawingRefs === undefined) emit('select-chart', null)
  }
}, { flush: 'post' })

onMounted(() => {
  mounted = true
  objectHostActive.value = true
  scroller = root.value?.parentElement?.querySelector<HTMLElement>('.canvas-scroll') ?? null
  scroller?.addEventListener('scroll', onScroll, { passive: true })
  if (typeof ResizeObserver !== 'undefined' && root.value) {
    resizeObserver = new ResizeObserver(onViewportResize)
    resizeObserver.observe(root.value)
  }
  window.addEventListener('pointermove', onPointerMove)
  window.addEventListener('pointerup', onPointerUp)
  window.addEventListener('pointercancel', onPointerCancel)
  window.addEventListener('keydown', onDragKeyDown)
  window.addEventListener('blur', onWindowBlur)
  document.addEventListener('pointerdown', onOutsidePointerDown)
  document.addEventListener('paste', onChartObjectPaste, true)
  window.addEventListener('resize', onViewportResize)
  document.fonts?.addEventListener('loadingdone', fontsChanged)
  watchPixelRatio()
  onScroll()
  publishOverlayCommands()
})

onActivated(() => {
  mounted = true
  objectHostActive.value = true
  watchPixelRatio()
  onScroll()
  publishOverlayCommands()
})

onDeactivated(() => {
  objectHostActive.value = false
  onPointerCancel()
  closeTitleSelection()
  closePlotAreaSelection()
  closeElementSelection()
  pivotAxis.close()
  mounted = false
  if (frameRequest !== undefined) window.cancelAnimationFrame(frameRequest)
  frameRequest = undefined
  pixelRatioQuery?.removeEventListener('change', watchPixelRatio)
  desiredPointer = undefined
  activePointerId = undefined
  contentTooltip.value = undefined
  pointerError.value = ''
  pointerDirty = false
  contentDirty = true
  contentCommands = []
  commandIndexes.clear()
  drag.value = null
  drafts.value = {}
  chartRenderer.dispose()
  renderedFrames.value = {}
  emit('overlay-commands', [])
})

onUnmounted(() => {
  objectHostActive.value = false
  onPointerCancel()
  closeTitleSelection()
  closePlotAreaSelection()
  closeElementSelection()
  pivotAxis.close()
  mounted = false
  desiredPointer = undefined
  activePointerId = undefined
  contentTooltip.value = undefined
  contentCommands = []
  commandIndexes.clear()
  authorPanes.clear()
  if (frameRequest !== undefined) window.cancelAnimationFrame(frameRequest)
  pixelRatioQuery?.removeEventListener('change', watchPixelRatio)
  document.fonts?.removeEventListener('loadingdone', fontsChanged)
  window.removeEventListener('resize', onViewportResize)
  resizeObserver?.disconnect()
  chartRenderer.dispose()
  scroller?.removeEventListener('scroll', onScroll)
  window.removeEventListener('pointermove', onPointerMove); window.removeEventListener('pointerup', onPointerUp)
  window.removeEventListener('pointercancel', onPointerCancel); window.removeEventListener('blur', onWindowBlur)
  document.removeEventListener('pointerdown', onOutsidePointerDown)
  document.removeEventListener('paste', onChartObjectPaste, true)
  window.removeEventListener('keydown', onDragKeyDown)
  emit('overlay-commands', [])
})
watch(() => [props.headerWidth, props.headerHeight] as const, () => {
  onPointerCancel(); authorPanes.clear(); routedPane = undefined; queueOverlayFrame()
})
defineExpose({ routeCanvasEvent, routeCanvasKeydown })
</script>

<template>
  <div ref="root" class="chart-overlay" :style="overlayStyle" aria-label="工作表图表图层">
    <Alert v-if="errorMessage" class="chart-render-error" type="error" show-icon :title="errorMessage" />
    <div
      v-for="item in displayCharts"
      :key="item.chart.id"
      class="chart-object"
      :class="{ 'chart-object-has-pivot-fields': (item.chart.ownedPivot || item.chart.pivotTableId) && pivotChartControls(item.chart).showFieldButtons }"
      :style="chartStyle(item)"
      @pointerdown.capture="routeNativeDrawingEvent($event, item)"
      @pointermove.capture="routeNativeDrawingEvent($event, item)"
      @click.capture="routeNativeDrawingEvent($event, item)"
      @dblclick.capture="routeNativeDrawingEvent($event, item)"
      @contextmenu.capture="routeNativeDrawingEvent($event, item)"
      @wheel.capture="routeNativeDrawingEvent($event, item)"
    >
      <ChartObjectContextMenu :chart-id="item.chart.id" :selected="selectedId === item.chart.id"
        :capture="() => objectActions.capture(item.chart.id)" :current="objectActions.isCurrent" :capabilities="objectActions.capabilities"
        :execute="objectActions.execute" :restore-focus="objectActions.restoreFocus"
        :intercept-context="event => openPivotAxisContext(event, item)" :intercept-keyboard="event => pivotAxisKeyboard(event, item)"
        @select="selectWholeChart(item.chart.id)">
      <Button
      type="text"
      class="chart-hitbox"
      :class="{ selected: selectedId === item.chart.id, 'chart-view-only': !canEditChart(item.chart), 'chart-title-hover': hoveredTitle?.chartId === item.chart.id,
        'chart-plot-area-hover': hoveredPlotArea === item.chart.id, 'chart-element-hover': hoveredElement?.chartId === item.chart.id,
        'chart-pivot-axis-hover': hoveredPivotAxis?.chartId === item.chart.id }"
      :aria-label="item.chart.title || '图表'"
      :aria-pressed="selectedId === item.chart.id"
      :data-chart-id="item.chart.id"
      :data-title-selected="selectedTitle?.chartId === item.chart.id ? selectedTitle.kind : undefined"
      :data-plot-area-selected="selectedPlotArea?.chartId === item.chart.id ? 'true' : undefined"
      :data-element-selected="selectedElement?.chartId === item.chart.id ? selectedElement.key : undefined"
      :data-element-kind="selectedElement?.chartId === item.chart.id ? selectedElement.element.kind : undefined"
      :data-pivot-axis-selected="selectedPivotAxis?.pivotChartId === item.chart.id ? selectedPivotAxis.fieldId : undefined"
      :data-rendered="renderedFrames[item.chart.id] ? 'true' : 'false'"
      :data-renderer="renderedFrames[item.chart.id] ? 'echarts-canvas' : undefined"
      :data-source-version="renderedFrames[item.chart.id]?.version"
      :data-bitmap-width="renderedFrames[item.chart.id]?.width"
      :data-bitmap-height="renderedFrames[item.chart.id]?.height"
      :data-compute-revision="snapshots[item.chart.id]?.modelRevision"
      :data-source-diagnostic="diagnostics[item.chart.id]?.code"
      :aria-busy="Boolean(pending[item.chart.id]) || titleCommit?.chartId === item.chart.id || plotAreaCommit?.chartId === item.chart.id || elementCommit?.chartId === item.chart.id || objectActions.busyChartId.value === item.chart.id || pivotAxis.commit.value?.chartId === item.chart.id || pivotFieldBusy[item.chart.id]"
      tabindex="0"
      @pointerdown="beginDrag($event, item, 'move')"
      @pointermove="hoverChart($event, item)"
      @wheel="scaleRelationContent($event, item)"
      @pointerleave="clearChartPointer(item.chart.id)"
      @blur="clearChartPointer(item.chart.id)"
      @lostpointercapture="onPointerCancel"
      @click.stop="selectChart(item.chart.id)"
      @dblclick="editChartObject($event, item)"
      @keydown.enter.prevent="enterChart(item.chart.id)"
      @keydown="chartKeyDown($event, item)"
    >
      <Spin v-if="pending[item.chart.id] || titleCommit?.chartId === item.chart.id || plotAreaCommit?.chartId === item.chart.id || elementCommit?.chartId === item.chart.id || objectActions.busyChartId.value === item.chart.id || pivotAxis.commit.value?.chartId === item.chart.id || pivotFieldBusy[item.chart.id]" size="small" class="chart-compute-spinner" />
      </Button>
      </ChartObjectContextMenu>
      <PivotChartFieldButtons v-if="studio && workbook.id && (item.chart.ownedPivot || item.chart.pivotTableId)"
        :key="`${workbook.id}:${activeSheetId}:${item.chart.id}`" :studio="studio" :unit-id="workbook.id" :sheet-id="activeSheetId" :chart-id="item.chart.id"
        :active="active && objectHostActive" :busy="Boolean(drag) || Boolean(titleCommit) || Boolean(plotAreaCommit) || elementEditing.busy.value || objectActions.busy.value || pivotAxis.busy.value || Boolean(pending[item.chart.id])"
        :open-settings="fieldId => openPivotFieldSettings(item.chart.id, fieldId)" :restore-focus="() => restorePivotFieldFocus(item.chart.id)"
        @pointerdown.capture="activatePivotFieldButtons(item.chart.id)" @focusin="activatePivotFieldButtons(item.chart.id)"
        @busy-change="setPivotFieldBusy(item.chart.id, $event)" />
      <PivotChartHierarchyControls v-if="pivotAxisState[item.chart.id] && pivotAxisState[item.chart.id]!.fields.length > 1 && pivotChartControls(item.chart).showExpandCollapseButtons"
        :state="pivotAxisState[item.chart.id]!" :disabled="!pivotAxis.editable(item.chart.id)" :busy="pivotAxis.commit.value?.chartId === item.chart.id"
        :step="direction => pivotAxis.step(item.chart.id, direction)" />
      <Tooltip v-if="selectedId === item.chart.id && selectedPlotArea?.chartId !== item.chart.id" title="调整图表大小；方向键微调，Shift 加速">
        <Button type="text" size="small" class="chart-resize-handle" aria-label="调整图表大小" :disabled="!canEditChart(item.chart) || anyPivotFieldBusy"
        @pointerdown="beginDrag($event, item, 'resize')" @lostpointercapture="onPointerCancel"
        @keydown="resizeChartWithKeyboard($event, item)" />
      </Tooltip>
      <Tooltip v-if="selectedId === item.chart.id && supportsChartPlotArea(item.chart.type)" title="选择绘图区，拖动其边框或使用浮动工具栏">
        <Button type="default" size="small" class="chart-plot-area-entry" aria-label="编辑图表绘图区" :aria-pressed="selectedPlotArea?.chartId === item.chart.id"
          @pointerdown.stop @click.stop="openPlotArea(item)">绘图区</Button>
      </Tooltip>
      <Tooltip v-for="handle in plotAreaHandles(item)" :key="handle.handle" :title="handle.label">
        <Button type="text" size="small" class="chart-plot-area-handle" :style="handle.style" :aria-label="handle.label"
          :data-plot-area-handle="handle.handle" :disabled="!canEditChart(item.chart) || !selectedPlotArea?.context || Boolean(titleCommit) || Boolean(plotAreaCommit) || Boolean(elementCommit) || Boolean(objectActions.busy.value) || anyPivotFieldBusy"
          @pointerdown="beginDrag($event, item, 'plot-area', handle.handle)" @lostpointercapture="onPointerCancel"
          @keydown="editPlotAreaWithKeyboard($event, item, handle.handle)" />
      </Tooltip>
    </div>
    <div v-if="selectedTitle && selectedTitleChart && !titleDrafts[selectedTitle.chartId]" class="chart-title-menu-host" :style="titleMenuStyle">
      <ChartTitleFloatMenu :key="selectedTitle.generation" ref="titleMenu" :chart="selectedTitleChart" :kind="selectedTitle.kind"
        :disabled="!canEditChart(selectedTitleChart) || !selectedTitle.context || Boolean(titleCommit) || Boolean(plotAreaCommit) || Boolean(elementCommit) || Boolean(objectActions.busy.value) || anyPivotFieldBusy" :apply="titleMenuApply"
        @close="closeTitleSelection(true)" />
    </div>
    <div v-if="selectedPlotArea && selectedPlotAreaChart && selectedPlotAreaTarget && !plotAreaDrafts[selectedPlotArea.chartId]" class="chart-plot-area-menu-host" :style="plotAreaMenuStyle">
      <ChartPlotAreaFloatMenu :key="selectedPlotArea.generation" ref="plotAreaMenu" :chart="selectedPlotAreaChart" :target="selectedPlotAreaTarget"
        :disabled="!canEditChart(selectedPlotAreaChart) || !selectedPlotArea.context || Boolean(titleCommit) || Boolean(plotAreaCommit) || Boolean(elementCommit) || Boolean(objectActions.busy.value) || anyPivotFieldBusy" :apply="plotAreaMenuApply"
        @close="closePlotAreaSelection(true)" />
    </div>
    <div v-if="selectedElement && selectedElementChart && selectedElementTarget" class="chart-element-menu-host" :style="elementMenuStyle">
      <ChartElementFloatMenu :key="selectedElement.generation" ref="elementMenu" :chart="selectedElementChart" :target="selectedElementTarget"
        :disabled="elementDisabled" :more-available="Boolean(openChartEditor)" :apply="elementMenuApply" :more="elementMenuMore" :promote="elementMenuPromote"
        @close="closeElementSelection(true)" />
    </div>
    <div v-if="selectedPivotAxis && pivotAxis.menuOpen.value" class="pivot-chart-axis-menu-host" :style="pivotAxisMenuStyle">
      <PivotChartAxisMenu :key="selectedPivotAxis.generation" :open="pivotAxis.menuOpen.value" :entries="pivotAxis.entries.value" :execute="pivotAxis.execute"
        @close="restore => pivotAxis.closeMenu(restore)" />
    </div>
    <Tooltip :open="Boolean(hoveredPivotAxis)" :title="hoveredPivotAxis ? `${hoveredPivotAxis.target.label.fieldName}：${hoveredPivotAxis.target.label.text}` : ''"
      :trigger="[]" :unique="false" :arrow="false" :styles="{ root: { pointerEvents: 'none' } }">
      <span class="chart-content-tooltip-anchor" :style="pivotAxisTooltipStyle" aria-hidden="true" />
    </Tooltip>
    <Tooltip :open="Boolean(contentTooltip)" :title="contentTooltipText" :trigger="[]" :unique="false" placement="top" :arrow="false" fresh
      :mouse-enter-delay="0" :mouse-leave-delay="0" :styles="{ root: { pointerEvents: 'none' }, container: { whiteSpace: 'pre-line' } }">
      <span class="chart-content-tooltip-anchor" :style="contentTooltipStyle" aria-hidden="true" />
    </Tooltip>
  </div>
</template>

<style scoped>
.chart-overlay { position: absolute; z-index: 2; inset: 0; overflow: hidden; clip-path: inset(26px 0 0 44px); pointer-events: none; }
.chart-object { position: absolute; pointer-events: none; }
.chart-object-has-pivot-fields :deep(.field-button-area-column) { padding-right: 64px; }
.chart-content-tooltip-anchor { position: absolute; width: 1px; height: 1px; pointer-events: none; }
.chart-title-menu-host { position: absolute; z-index: var(--chart-author-layer); pointer-events: auto; }
.chart-plot-area-menu-host { position: absolute; z-index: var(--chart-author-layer); pointer-events: auto; }
.chart-element-menu-host { position: absolute; z-index: var(--chart-author-layer); pointer-events: auto; }
.pivot-chart-axis-menu-host { position: absolute; z-index: var(--chart-author-layer); pointer-events: none; }
.chart-hitbox { position: absolute; inset: 0; width: 100%; height: 100%; display: block; box-sizing: border-box; margin: 0; border: 1px solid transparent; padding: 0; color: inherit; background: transparent; pointer-events: auto; cursor: move; outline: none; text-align: left; }
.chart-hitbox:hover, .chart-hitbox:focus-visible { background: transparent; }
.chart-hitbox.chart-view-only { cursor: default; }
.chart-hitbox.chart-title-hover { cursor: text; }
.chart-hitbox.chart-plot-area-hover { cursor: move; }
.chart-hitbox.chart-view-only.chart-plot-area-hover { cursor: default; }
.chart-hitbox.chart-element-hover { cursor: pointer; }
.chart-hitbox.chart-view-only.chart-element-hover { cursor: default; }
.chart-hitbox.chart-pivot-axis-hover, .chart-hitbox.chart-view-only.chart-pivot-axis-hover { cursor: pointer; }
.chart-hitbox:hover, .chart-hitbox.selected { border-color: var(--chart-primary); box-shadow: 0 0 0 1px var(--chart-focus-ring); }
.chart-render-error { position: absolute; z-index: var(--chart-author-layer); inset: 30px 4px auto 48px; pointer-events: auto; }
.chart-compute-spinner { position: absolute; inset: 50% auto auto 50%; transform: translate(-50%, -50%); pointer-events: none; }
.chart-resize-handle { position: absolute; right: -5px; bottom: -5px; width: 9px; min-width: 9px; height: 9px; padding: 0; border: 1px solid var(--chart-surface); border-radius: 2px; background: var(--chart-primary); pointer-events: auto; cursor: nwse-resize; }
.chart-resize-handle:hover, .chart-resize-handle:focus-visible { background: var(--chart-primary); box-shadow: 0 0 0 2px var(--chart-focus-ring); }
.chart-resize-handle:disabled { background: var(--chart-surface); border-color: var(--chart-focus-ring); cursor: default; }
.chart-plot-area-entry { position: absolute; right: 48px; top: 8px; pointer-events: auto; }
.chart-plot-area-handle { position: absolute; z-index: 2; width: 9px; min-width: 9px; height: 9px; padding: 0; transform: translate(-50%, -50%);
  border: 1px solid var(--chart-surface); border-radius: 2px; background: var(--chart-primary); pointer-events: auto; }
.chart-plot-area-handle:hover, .chart-plot-area-handle:focus-visible { background: var(--chart-primary); box-shadow: 0 0 0 2px var(--chart-focus-ring); }
.chart-plot-area-handle:disabled { background: var(--chart-surface); border-color: var(--chart-focus-ring); cursor: default !important; }
</style>
