<script setup lang="ts">
import { computed, onActivated, onDeactivated, onMounted, onUnmounted, ref, watch } from 'vue'
import { Alert, Button, Spin, theme } from 'antdv-next'
import type { OfficeStudio } from '../api/office-studio'
import type { ChartModel } from '../domain/model'
import { chartEChartsOption } from '../features/charts/echarts-option'
import { EChartsCanvasRenderer, type EChartsCanvasFrame } from '../features/charts/echarts-renderer'
import { chartAxisPointerLocalPoint } from '../features/charts/chart-axis-pointer'
import type { ChartVector } from '../features/charts/chart-geometry'
import type { ChartComputedStatistics } from '../features/charts/chart-derived'
import { useChartStatistics } from '../features/charts/views/use-chart-statistics'
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
  mode: 'move' | 'resize'
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
const drag = ref<DragState | null>(null)
let scroller: HTMLElement | null = null
let resizeObserver: ResizeObserver | undefined
let frameRequest: number | undefined
let pixelRatioQuery: MediaQueryList | undefined
let mounted = false
let contentDirty = true
let pointerDirty = false
let desiredPointer: { chartId: string; point: ChartVector } | undefined
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
  emit('remove-chart', chartId)
  clearChartPointer(chartId)
  selectedId.value = null
}

function beginDrag(event: PointerEvent, item: ChartDisplay, mode: DragState['mode']): void {
  if (event.button !== 0) return
  event.preventDefault(); event.stopPropagation()
  clearChartPointer()
  selectedId.value = item.chart.id
  const point = pointerPoint(event)
  drag.value = {
    chartId: item.chart.id,
    mode,
    originX: point.x,
    originY: point.y,
    position: { ...item.position },
    size: { ...item.size },
  }
}

function onPointerMove(event: PointerEvent): void {
  if (!mounted) return
  const active = drag.value
  if (!active) return
  const point = pointerPoint(event)
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

function onPointerUp(): void {
  const active = drag.value
  if (!active) return
  const draft = drafts.value[active.chartId]
  if (draft) emit('update-chart', { chartId: active.chartId, patch: { position: draft.position, size: draft.size } })
  const next = { ...drafts.value }
  delete next[active.chartId]
  drafts.value = next
  drag.value = null
}

function onPointerCancel(): void {
  clearChartPointer()
  const active = drag.value
  if (!active) return
  const next = { ...drafts.value }
  delete next[active.chartId]
  drafts.value = next
  drag.value = null
}

function onScroll(): void {
  if (!scroller) return
  clearChartPointer()
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
  clearChartPointer()
  chartRenderer.invalidate()
  publishOverlayCommands()
}

watch([displayCharts, chartTheme, () => props.zoom], publishOverlayCommands, { deep: true, flush: 'post' })
watch(statistics, publishOverlayCommands, { flush: 'post' })
watch(() => props.activeSheetId, () => {
  clearChartPointer()
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
  emit('overlay-commands', [])
})
</script>

<template>
  <div ref="root" class="chart-overlay" :style="overlayStyle" aria-label="工作表图表图层">
    <Alert v-if="errorMessage" class="chart-render-error" type="error" show-icon :message="errorMessage" />
    <Button
      v-for="item in displayCharts"
      :key="item.chart.id"
      type="text"
      class="chart-hitbox"
      :class="{ selected: selectedId === item.chart.id }"
      :style="chartStyle(item)"
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
      @click.stop="selectChart(item.chart.id)"
      @keydown.enter.prevent="selectChart(item.chart.id)"
      @keydown="removeChart($event, item.chart.id)"
    >
      <Spin v-if="pending[item.chart.id]" size="small" class="chart-compute-spinner" />
      <span v-if="selectedId === item.chart.id" class="chart-resize-handle" aria-label="调整图表大小" @pointerdown="beginDrag($event, item, 'resize')" />
    </Button>
  </div>
</template>

<style scoped>
.chart-overlay { position: absolute; z-index: 2; inset: 0; overflow: hidden; clip-path: inset(26px 0 0 44px); pointer-events: none; }
.chart-hitbox { position: absolute; display: block; box-sizing: border-box; margin: 0; border: 1px solid transparent; padding: 0; color: inherit; background: transparent; pointer-events: auto; cursor: move; outline: none; text-align: left; }
.chart-hitbox:hover, .chart-hitbox:focus-visible { background: transparent; }
.chart-hitbox:hover, .chart-hitbox.selected { border-color: var(--chart-primary); box-shadow: 0 0 0 1px var(--chart-focus-ring); }
.chart-render-error { position: absolute; z-index: 3; inset: 30px 4px auto 48px; pointer-events: auto; }
.chart-compute-spinner { position: absolute; inset: 50% auto auto 50%; transform: translate(-50%, -50%); pointer-events: none; }
.chart-resize-handle { position: absolute; right: -5px; bottom: -5px; width: 9px; height: 9px; border: 1px solid var(--chart-surface); border-radius: 2px; background: var(--chart-primary); cursor: nwse-resize; }
</style>
