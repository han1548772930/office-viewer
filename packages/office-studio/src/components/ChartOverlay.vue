<script setup lang="ts">
import { computed, onActivated, onDeactivated, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue'
import { Alert, Button, Spin, Tooltip, theme } from 'antdv-next'
import type { OfficeStudio } from '../api/office-studio'
import type { ChartModel } from '../domain/model'
import { chartEChartsOption } from '../features/charts/echarts-option'
import { EChartsCanvasRenderer, type EChartsCanvasFrame } from '../features/charts/echarts-renderer'
import { chartAxisPointerLocalPoint } from '../features/charts/chart-axis-pointer'
import type { ChartVector } from '../features/charts/chart-geometry'
import type { ChartComputedStatistics } from '../features/charts/chart-derived'
import { useChartStatistics } from '../features/charts/views/use-chart-statistics'
import { canEditWorksheetProtection } from '../features/sheet-protection/range-protection'
import type { DrawCommand } from '../rendering/contracts'
import type { ExcelWorkbookChartView } from './chart-overlay-model'
import { cloneChart, toChartWorkbook } from './chart-overlay-model'

const props = withDefaults(defineProps<{
  workbook: ExcelWorkbookChartView
  activeSheetId: string
  zoom: number
  computedStatistics?: Readonly<Record<string, ChartComputedStatistics>>
  studio?: OfficeStudio
}>(), { zoom: 100 })

const emit = defineEmits<{
  (event: 'update-chart', payload: { chartId: string; patch: Partial<Pick<ChartModel, 'position' | 'size'>> }): void
  (event: 'remove-chart', chartId: string): void
  (event: 'overlay-commands', commands: DrawCommand[]): void
}>()

interface ChartDisplay {
  chart: ChartModel
  position: { x: number; y: number }
  size: { width: number; height: number }
}
interface DragState {
  chartId: string
  mode: 'move' | 'resize' | 'sankey-node'
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
}

const root = ref<HTMLElement | null>(null)
const { token } = theme.useToken()
const scrollLeft = ref(0)
const scrollTop = ref(0)
const selectedId = ref<string | null>(null)
const renderError = ref('')
const pointerError = ref('')
const renderedFrames = ref<Record<string, { version: number; width: number; height: number }>>({})
const drafts = ref<Record<string, Pick<ChartDisplay, 'position' | 'size'>>>({})
const drag = shallowRef<DragState | null>(null)
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
let contentCommands: DrawCommand[] = []
const commandIndexes = new Map<string, number>()
const chartRenderer = new EChartsCanvasRenderer()

const sheetCharts = computed(() => (props.workbook.charts ?? []).filter((chart) => chart.sheetId === props.activeSheetId))
const domainWorkbook = computed(() => props.studio ? props.studio.getWorkbookModel(props.workbook.id ?? '') : toChartWorkbook(props.workbook))
const { statistics, snapshots, pending, errors: computeErrors } = useChartStatistics({
  studio: () => props.studio, workbook: () => props.studio ? props.studio.getWorkbookModel(props.workbook.id ?? '') : domainWorkbook.value, activeSheetId: () => props.activeSheetId,
  externalStatistics: () => props.computedStatistics,
})
const errorMessage = computed(() => renderError.value || pointerError.value || Object.values(computeErrors.value).join('；'))
const displayCharts = computed<ChartDisplay[]>(() => sheetCharts.value.map((chart) => {
  const draft = drafts.value[chart.id]
  const position = draft?.position ?? chart.position
  const size = draft?.size ?? chart.size
  return { chart, position, size }
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
}))
const renderedChartIds = computed(() => new Set(sheetCharts.value.map(chart => chart.id)))

function renderCommands(): DrawCommand[] {
  const host = root.value
  commandIndexes.clear()
  if (!host || host.clientWidth <= 0 || host.clientHeight <= 0) {
    chartRenderer.retain(new Set())
    renderedFrames.value = {}
    activePointerId = undefined
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
      const draft = drafts.value[item.chart.id]
      const model = draft ? { ...cloneChart(item.chart), position: { ...item.position }, size: { ...item.size } } : item.chart
      const option = chartEChartsOption(workbook, model, chartTheme.value, computed)
      const frame = chartRenderer.render(item.chart.id, option, {
        width: Math.max(1, item.size.width),
        height: Math.max(1, item.size.height),
        devicePixelRatio: dpr,
      }, chartTheme.value)
      frames[item.chart.id] = { version: frame.sourceVersion, width: frame.source.width, height: frame.source.height }
      commandIndexes.set(item.chart.id, commands.length)
      commands.push({
        kind: 'image',
        source: frame.source,
        sourceVersion: frame.sourceVersion,
        x: item.position.x,
        y: item.position.y,
        width: item.size.width,
        height: item.size.height,
      })
    } catch (error) {
      errorMessage = error instanceof Error ? error.message : '图表绘制失败'
    }
  }
  renderedFrames.value = frames
  renderError.value = errorMessage
  chartRenderer.retain(renderedChartIds.value)
  return commands
}

function zoomScale(): number { return Math.max(0.1, props.zoom / 100) }
function viewportX(value: number): number { return value * zoomScale() - scrollLeft.value }
function viewportY(value: number): number { return value * zoomScale() - scrollTop.value }
function pointerPoint(event: PointerEvent): { x: number; y: number } {
  const rect = root.value?.getBoundingClientRect()
  return rect ? { x: event.clientX - rect.left, y: event.clientY - rect.top } : { x: 0, y: 0 }
}

function chartStyle(item: ChartDisplay): Record<string, string> {
  return {
    left: `${viewportX(item.position.x)}px`,
    top: `${viewportY(item.position.y)}px`,
    width: `${item.size.width * zoomScale()}px`,
    height: `${item.size.height * zoomScale()}px`,
  }
}

function selectChart(chartId: string): void { selectedId.value = chartId }
function canEditChart(chart: ChartModel): boolean {
  const workbook = props.studio ? props.studio.getWorkbookModel(props.workbook.id ?? '') : domainWorkbook.value
  const sheet = workbook?.sheets.find(sheet => sheet.id === chart.sheetId)
  return !!sheet && canEditWorksheetProtection(sheet, props.studio?.currentUserId ?? 'local-user')
}
function hoverChart(event: PointerEvent, item: ChartDisplay): void {
  if (!mounted || drag.value || !root.value || !commandIndexes.has(item.chart.id)) return
  desiredPointer = { chartId: item.chart.id, point: chartAxisPointerLocalPoint(pointerPoint(event), item.position,
    { zoom: props.zoom, scrollLeft: scrollLeft.value, scrollTop: scrollTop.value }) }
  pointerDirty = true
  queueOverlayFrame()
}

function clearChartPointer(chartId?: string): void {
  if (chartId !== undefined && desiredPointer?.chartId !== chartId) return
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
  if (activePointerId && activePointerId !== desiredPointer?.chartId) {
    updatePointerFrame(activePointerId, chartRenderer.updatePointer(activePointerId))
    activePointerId = undefined
  }
  if (desiredPointer && commandIndexes.has(desiredPointer.chartId)) {
    updatePointerFrame(desiredPointer.chartId, chartRenderer.updatePointer(desiredPointer.chartId, desiredPointer.point))
    activePointerId = desiredPointer.chartId
    pointerError.value = ''
  }
}

function removeChart(event: KeyboardEvent, chartId: string): void {
  if (event.key !== 'Delete' && event.key !== 'Backspace') return
  event.preventDefault(); event.stopPropagation()
  const chart = sheetCharts.value.find(chart => chart.id === chartId)
  if (!chart || !canEditChart(chart)) return
  emit('remove-chart', chartId)
  clearChartPointer(chartId)
  selectedId.value = null
}

function beginDrag(event: PointerEvent, item: ChartDisplay, mode: 'move' | 'resize'): void {
  if (event.button !== 0 || !mounted || drag.value) return
  if (mode === 'resize' && !canEditChart(item.chart)) return
  event.preventDefault(); event.stopPropagation()
  clearChartPointer()
  selectedId.value = item.chart.id
  const point = pointerPoint(event)
  const target = event.currentTarget instanceof HTMLElement ? event.currentTarget : null
  let dragMode: DragState['mode'] = mode
  if (mode === 'move') {
    try {
      const content = chartRenderer.beginContentDrag(item.chart.id, chartAxisPointerLocalPoint(point, item.position,
        { zoom: props.zoom, scrollLeft: scrollLeft.value, scrollTop: scrollTop.value }))
      if (content) { dragMode = content.kind; updatePointerFrame(item.chart.id, content.frame); queueOverlayFrame() }
    } catch (error) {
      pointerError.value = error instanceof Error ? error.message : '图表内容拖动失败'
      return
    }
  }
  if (dragMode !== 'sankey-node' && !canEditChart(item.chart)) return
  const workbookId = props.workbook.id
  drag.value = {
    chartId: item.chart.id,
    mode: dragMode,
    pointerId: event.pointerId,
    target,
    workbookId,
    studio: props.studio,
    sheetId: props.activeSheetId,
    ...(props.studio && workbookId ? { modelRevision: props.studio.getWorkbookRevision(workbookId) } : {}),
    chartSnapshot: JSON.stringify(item.chart),
    originX: point.x,
    originY: point.y,
    position: { ...item.position },
    size: { ...item.size },
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
  return active.studio === props.studio && active.workbookId === props.workbook.id && active.sheetId === props.activeSheetId
    && (active.modelRevision === undefined || !props.studio || !active.workbookId || props.studio.getWorkbookRevision(active.workbookId) === active.modelRevision)
    && JSON.stringify(chart) === active.chartSnapshot
    && !!chart && (active.mode === 'sankey-node' || canEditChart(chart))
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
  if (active.mode === 'sankey-node') {
    desiredContentDrag = { chartId: active.chartId, pointerId: active.pointerId, point: chartAxisPointerLocalPoint(point, active.position,
      { zoom: props.zoom, scrollLeft: scrollLeft.value, scrollTop: scrollTop.value }) }
    queueOverlayFrame()
    return
  }
  const scale = zoomScale()
  const dx = (point.x - active.originX) / scale
  const dy = (point.y - active.originY) / scale
  const position = active.mode === 'move'
    ? { x: Math.max(0, active.position.x + dx), y: Math.max(0, active.position.y + dy) }
    : active.position
  const size = active.mode === 'resize'
    ? { width: Math.max(160, active.size.width + dx), height: Math.max(120, active.size.height + dy) }
    : active.size
  drafts.value = { ...drafts.value, [active.chartId]: { position, size } }
}

function onPointerUp(event: PointerEvent): void {
  const active = drag.value
  if (!active || event.pointerId !== active.pointerId) return
  if (!dragIsCurrent(active)) { onPointerCancel(); return }
  if (active.mode === 'sankey-node') {
    desiredContentDrag = { chartId: active.chartId, pointerId: active.pointerId, point: chartAxisPointerLocalPoint(pointerPoint(event), active.position,
      { zoom: props.zoom, scrollLeft: scrollLeft.value, scrollTop: scrollTop.value }) }
    try { flushContentDrag() }
    catch (error) { onPointerCancel(); pointerError.value = error instanceof Error ? error.message : '图表内容拖动失败'; return }
    if (drag.value !== active) return
  } else onPointerMove(event)
  const draft = drafts.value[active.chartId]
  const next = { ...drafts.value }
  delete next[active.chartId]
  drafts.value = next
  drag.value = null
  releaseDragPointer(active)
  if (active.mode === 'sankey-node') {
    try { updatePointerFrame(active.chartId, chartRenderer.endContentDrag(active.chartId)); queueOverlayFrame() }
    catch (error) { pointerError.value = error instanceof Error ? error.message : '图表内容拖动结束失败' }
  } else if (draft) emit('update-chart', { chartId: active.chartId, patch: { position: draft.position, size: draft.size } })
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
  drag.value = null
  releaseDragPointer(active)
  if (active.mode === 'sankey-node') {
    try { updatePointerFrame(active.chartId, chartRenderer.endContentDrag(active.chartId, true)); queueOverlayFrame() }
    catch (error) { pointerError.value = error instanceof Error ? error.message : '图表内容拖动取消失败' }
  }
}

function onDragKeyDown(event: KeyboardEvent): void {
  if (event.key !== 'Escape' || !drag.value) return
  event.preventDefault(); event.stopPropagation()
  onPointerCancel()
}

function resizeChartWithKeyboard(event: KeyboardEvent, item: ChartDisplay): void {
  if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return
  event.preventDefault(); event.stopPropagation()
  if (!canEditChart(item.chart)) return
  if (drag.value) onPointerCancel()
  const step = event.shiftKey ? 10 : 1
  const width = Math.max(160, item.size.width + (event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0))
  const height = Math.max(120, item.size.height + (event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0))
  if (width !== item.chart.size.width || height !== item.chart.size.height) emit('update-chart', { chartId: item.chart.id, patch: { size: { width, height } } })
}

function onScroll(): void {
  if (!scroller) return
  onPointerCancel()
  scrollLeft.value = scroller.scrollLeft
  scrollTop.value = scroller.scrollTop
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
        chartRenderer.invalidate()
        publishOverlayCommands()
      }
    }
    emit('overlay-commands', [...contentCommands])
  })
}
function watchPixelRatio(): void {
  pixelRatioQuery?.removeEventListener('change', watchPixelRatio)
  pixelRatioQuery = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`)
  pixelRatioQuery.addEventListener('change', watchPixelRatio)
  publishOverlayCommands()
}
function fontsChanged(): void {
  onPointerCancel()
  clearChartPointer()
  chartRenderer.invalidate()
  publishOverlayCommands()
}

watch([displayCharts, chartTheme, () => props.zoom], publishOverlayCommands, { deep: true, flush: 'post' })
watch(statistics, publishOverlayCommands, { flush: 'post' })
watch([() => props.workbook, () => props.studio, () => props.zoom, chartTheme, statistics], () => {
  if (drag.value) onPointerCancel()
}, { deep: true, flush: 'sync' })
watch(() => props.activeSheetId, () => {
  onPointerCancel()
  pointerError.value = ''
  selectedId.value = null
  drag.value = null
  drafts.value = {}
  publishOverlayCommands()
})

onMounted(() => {
  mounted = true
  scroller = root.value?.parentElement?.querySelector<HTMLElement>('.canvas-scroll') ?? null
  scroller?.addEventListener('scroll', onScroll, { passive: true })
  if (typeof ResizeObserver !== 'undefined' && root.value) {
    resizeObserver = new ResizeObserver(publishOverlayCommands)
    resizeObserver.observe(root.value)
  }
  window.addEventListener('pointermove', onPointerMove)
  window.addEventListener('pointerup', onPointerUp)
  window.addEventListener('pointercancel', onPointerCancel)
  window.addEventListener('keydown', onDragKeyDown)
  window.addEventListener('blur', onPointerCancel)
  window.addEventListener('resize', publishOverlayCommands)
  document.fonts?.addEventListener('loadingdone', fontsChanged)
  watchPixelRatio()
  onScroll()
  publishOverlayCommands()
})

onActivated(() => {
  mounted = true
  watchPixelRatio()
  onScroll()
  publishOverlayCommands()
})

onDeactivated(() => {
  onPointerCancel()
  mounted = false
  if (frameRequest !== undefined) window.cancelAnimationFrame(frameRequest)
  frameRequest = undefined
  pixelRatioQuery?.removeEventListener('change', watchPixelRatio)
  desiredPointer = undefined
  activePointerId = undefined
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
  onPointerCancel()
  mounted = false
  desiredPointer = undefined
  activePointerId = undefined
  contentCommands = []
  commandIndexes.clear()
  if (frameRequest !== undefined) window.cancelAnimationFrame(frameRequest)
  pixelRatioQuery?.removeEventListener('change', watchPixelRatio)
  document.fonts?.removeEventListener('loadingdone', fontsChanged)
  window.removeEventListener('resize', publishOverlayCommands)
  resizeObserver?.disconnect()
  chartRenderer.dispose()
  scroller?.removeEventListener('scroll', onScroll)
  window.removeEventListener('pointermove', onPointerMove); window.removeEventListener('pointerup', onPointerUp)
  window.removeEventListener('pointercancel', onPointerCancel); window.removeEventListener('blur', onPointerCancel)
  window.removeEventListener('keydown', onDragKeyDown)
  emit('overlay-commands', [])
})
</script>

<template>
  <div ref="root" class="chart-overlay" :style="overlayStyle" aria-label="工作表图表图层">
    <Alert v-if="errorMessage" class="chart-render-error" type="error" show-icon :message="errorMessage" />
    <div
      v-for="item in displayCharts"
      :key="item.chart.id"
      class="chart-object"
      :style="chartStyle(item)"
    >
      <Button
      type="text"
      class="chart-hitbox"
      :class="{ selected: selectedId === item.chart.id, 'chart-view-only': !canEditChart(item.chart) }"
      :aria-label="item.chart.title || '图表'"
      :aria-pressed="selectedId === item.chart.id"
      :data-rendered="renderedFrames[item.chart.id] ? 'true' : 'false'"
      :data-renderer="renderedFrames[item.chart.id] ? 'echarts-canvas' : undefined"
      :data-source-version="renderedFrames[item.chart.id]?.version"
      :data-bitmap-width="renderedFrames[item.chart.id]?.width"
      :data-bitmap-height="renderedFrames[item.chart.id]?.height"
      :data-compute-revision="snapshots[item.chart.id]?.modelRevision"
      :aria-busy="Boolean(pending[item.chart.id])"
      tabindex="0"
      @pointerdown="beginDrag($event, item, 'move')"
      @pointermove="hoverChart($event, item)"
      @pointerleave="clearChartPointer(item.chart.id)"
      @blur="clearChartPointer(item.chart.id)"
      @lostpointercapture="onPointerCancel"
      @click.stop="selectChart(item.chart.id)"
      @keydown.enter.prevent="selectChart(item.chart.id)"
      @keydown="removeChart($event, item.chart.id)"
    >
      <Spin v-if="pending[item.chart.id]" size="small" class="chart-compute-spinner" />
      </Button>
      <Tooltip v-if="selectedId === item.chart.id" title="调整图表大小；方向键微调，Shift 加速">
        <Button type="text" size="small" class="chart-resize-handle" aria-label="调整图表大小" :disabled="!canEditChart(item.chart)"
        @pointerdown="beginDrag($event, item, 'resize')" @lostpointercapture="onPointerCancel"
        @keydown="resizeChartWithKeyboard($event, item)" />
      </Tooltip>
    </div>
  </div>
</template>

<style scoped>
.chart-overlay { position: absolute; z-index: 2; inset: 0; overflow: hidden; clip-path: inset(26px 0 0 44px); pointer-events: none; }
.chart-object { position: absolute; pointer-events: none; }
.chart-hitbox { position: absolute; inset: 0; width: 100%; height: 100%; display: block; box-sizing: border-box; margin: 0; border: 1px solid transparent; padding: 0; color: inherit; background: transparent; pointer-events: auto; cursor: move; outline: none; text-align: left; }
.chart-hitbox:hover, .chart-hitbox:focus-visible { background: transparent; }
.chart-hitbox.chart-view-only { cursor: default; }
.chart-hitbox:hover, .chart-hitbox.selected { border-color: var(--chart-primary); box-shadow: 0 0 0 1px var(--chart-focus-ring); }
.chart-render-error { position: absolute; z-index: 3; inset: 30px 4px auto 48px; pointer-events: auto; }
.chart-compute-spinner { position: absolute; inset: 50% auto auto 50%; transform: translate(-50%, -50%); pointer-events: none; }
.chart-resize-handle { position: absolute; right: -5px; bottom: -5px; width: 9px; min-width: 9px; height: 9px; padding: 0; border: 1px solid var(--chart-surface); border-radius: 2px; background: var(--chart-primary); pointer-events: auto; cursor: nwse-resize; }
.chart-resize-handle:hover, .chart-resize-handle:focus-visible { background: var(--chart-primary); box-shadow: 0 0 0 2px var(--chart-focus-ring); }
.chart-resize-handle:disabled { background: var(--chart-surface); border-color: var(--chart-focus-ring); cursor: default; }
</style>
