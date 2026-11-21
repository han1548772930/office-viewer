<script setup lang="ts">
import { computed, onMounted, onUnmounted, ref, watch } from 'vue'
import { Button } from 'antdv-next'
import type { ChartModel } from '../domain/model'
import { chartRenderCommands } from '../features/charts/chart-engine'
import type { DrawCommand } from '../rendering/contracts'
import type { ExcelWorkbookChartView } from './chart-overlay-model'
import { cloneChart, toChartWorkbook } from './chart-overlay-model'

const props = withDefaults(defineProps<{
  workbook: ExcelWorkbookChartView
  activeSheetId: string
  zoom: number
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
const scrollLeft = ref(0)
const scrollTop = ref(0)
const selectedId = ref<string | null>(null)
const drafts = ref<Record<string, Pick<ChartDisplay, 'position' | 'size'>>>({})
const drag = ref<DragState | null>(null)
let scroller: HTMLElement | null = null

const sheetCharts = computed(() => (props.workbook.charts ?? []).filter((chart) => chart.sheetId === props.activeSheetId))
const domainWorkbook = computed(() => toChartWorkbook(props.workbook))
const displayCharts = computed<Array<ChartDisplay & { commands: DrawCommand[] }>>(() => sheetCharts.value.flatMap((chart) => {
  const draft = drafts.value[chart.id]
  const position = draft?.position ?? chart.position
  const size = draft?.size ?? chart.size
  const model = draft ? { ...cloneChart(chart), position: { ...position }, size: { ...size } } : chart
  try { return [{ chart, position, size, commands: chartRenderCommands(domainWorkbook.value, model) }] }
  catch { return [] }
}))

const overlayCommands = computed(() => displayCharts.value.flatMap((item) => item.commands))

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
function removeChart(event: KeyboardEvent, chartId: string): void {
  if (event.key !== 'Delete' && event.key !== 'Backspace') return
  event.preventDefault(); event.stopPropagation()
  emit('remove-chart', chartId)
  selectedId.value = null
}

function beginDrag(event: PointerEvent, item: ChartDisplay, mode: DragState['mode']): void {
  if (event.button !== 0) return
  event.preventDefault(); event.stopPropagation()
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

function onScroll(): void {
  if (!scroller) return
  scrollLeft.value = scroller.scrollLeft
  scrollTop.value = scroller.scrollTop
}

function publishOverlayCommands(): void { emit('overlay-commands', overlayCommands.value) }

watch(overlayCommands, publishOverlayCommands, { deep: true })

onMounted(() => {
  scroller = root.value?.parentElement?.querySelector<HTMLElement>('.canvas-scroll') ?? null
  scroller?.addEventListener('scroll', onScroll, { passive: true })
  window.addEventListener('pointermove', onPointerMove)
  window.addEventListener('pointerup', onPointerUp)
  publishOverlayCommands()
})

onUnmounted(() => {
  scroller?.removeEventListener('scroll', onScroll)
  window.removeEventListener('pointermove', onPointerMove); window.removeEventListener('pointerup', onPointerUp)
})
</script>

<template>
  <div ref="root" class="chart-overlay" aria-label="工作表图表图层">
    <Button
      v-for="item in displayCharts"
      :key="item.chart.id"
      type="text"
      class="chart-hitbox"
      :class="{ selected: selectedId === item.chart.id }"
      :style="chartStyle(item)"
      :aria-label="item.chart.title || '图表'"
      :aria-pressed="selectedId === item.chart.id"
      tabindex="0"
      @pointerdown="beginDrag($event, item, 'move')"
      @click.stop="selectChart(item.chart.id)"
      @keydown.enter.prevent="selectChart(item.chart.id)"
      @keydown="removeChart($event, item.chart.id)"
    >
      <span v-if="selectedId === item.chart.id" class="chart-resize-handle" aria-label="调整图表大小" @pointerdown="beginDrag($event, item, 'resize')" />
    </Button>
  </div>
</template>

<style scoped>
.chart-overlay { position: absolute; z-index: 2; inset: 0; overflow: hidden; clip-path: inset(26px 0 0 44px); pointer-events: none; }
.chart-hitbox { position: absolute; display: block; box-sizing: border-box; margin: 0; border: 1px solid transparent; padding: 0; color: inherit; background: transparent; pointer-events: auto; cursor: move; outline: none; text-align: left; }
.chart-hitbox:hover, .chart-hitbox:focus-visible { background: transparent; }
.chart-hitbox:hover, .chart-hitbox.selected { border-color: #246bfd; box-shadow: 0 0 0 1px rgb(36 107 253 / 18%); }
.chart-resize-handle { position: absolute; right: -5px; bottom: -5px; width: 9px; height: 9px; border: 1px solid #fff; border-radius: 2px; background: #246bfd; cursor: nwse-resize; }
</style>
