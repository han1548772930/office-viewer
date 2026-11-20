<script setup lang="ts">
import { computed, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, watch } from 'vue'
import type { RenderBackend, RenderSurface } from '../../rendering/contracts'
import type { DrawCommand } from '../../rendering/contracts'
import { createRenderSurface } from '../../rendering/createRenderSurface'
import type { GridHit, GridRange, GridViewport, SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, SpreadsheetGeometry } from '../../rendering/spreadsheet/layout'
import { createSpreadsheetScene, displayCell } from '../../rendering/spreadsheet/scene'
import { createCellEditorLayout } from '../../rendering/spreadsheet/editor-layout'

const props = defineProps<{ sheet: SpreadsheetRenderSheet; selectedKey: string; selectedRange: string[]; selectedRanges?: GridRange[]; editingKey: string | null; editValue: string; zoom: number; searchText?: string; overlayCommands?: readonly DrawCommand[] }>()
const emit = defineEmits<{
  (event: 'select-cell', key: string, pointer: PointerEvent): void
  (event: 'extend-selection', key: string): void
  (event: 'select-header', hit: GridHit, pointer: PointerEvent): void
  (event: 'edit-cell', key: string): void
  (event: 'update:editValue', value: string): void
  (event: 'finish-edit'): void
  (event: 'editor-keydown', eventValue: KeyboardEvent): void
  (event: 'paste', eventValue: ClipboardEvent): void
  (event: 'resize-axis', axis: 'row' | 'column', index: number, size: number): void
  (event: 'backend-change', backend: RenderBackend, reason?: string): void
}>()
const root = ref<HTMLElement | null>(null), canvasHost = ref<HTMLElement | null>(null), scroller = ref<HTMLElement | null>(null), input = ref<HTMLInputElement | null>(null)
const viewport = ref<GridViewport>({ width: 1, height: 1, scrollLeft: 0, scrollTop: 0, zoom: props.zoom / 100 })
const geometry = computed(() => new SpreadsheetGeometry(props.sheet))
const scene = computed(() => createSpreadsheetScene(geometry.value, viewport.value, { activeCell: props.selectedKey, selectedCells: props.selectedRange, selectedRanges: props.selectedRanges, searchText: props.searchText, overlayCommands: props.overlayCommands }))
const fullSize = computed(() => ({ width: `${GRID_HEADER_WIDTH + geometry.value.columns.offset(props.sheet.columnCount) * props.zoom / 100}px`, height: `${GRID_HEADER_HEIGHT + geometry.value.rows.offset(props.sheet.rowCount) * props.zoom / 100}px` }))
const editLayout = computed(() => props.editingKey ? createCellEditorLayout(geometry.value, props.editingKey, viewport.value) : undefined)
const editStyle = computed(() => editLayout.value ? { left: `${editLayout.value.rect.x}px`, top: `${editLayout.value.rect.y}px`, width: `${editLayout.value.rect.width}px`, height: `${editLayout.value.rect.height}px`, clipPath: editLayout.value.clipPath, fontFamily: props.sheet.cells[props.editingKey ?? '']?.fontFamily ?? 'Aptos, sans-serif', fontSize: `${(props.sheet.cells[props.editingKey ?? '']?.fontSize ?? 11) * 4 / 3 * viewport.value.zoom}px` } : { clipPath: 'inset(50%)' })
const composing = ref(false)
const surfaceError = ref('')
const sheetViews = new Map<string, { left: number; top: number }>()
let surface: RenderSurface | undefined, observer: ResizeObserver | undefined, frame = 0, disposed = false, active = true, dragging = false
let dprQuery: MediaQueryList | undefined
let resizeDrag: { axis: 'row' | 'column'; index: number; size: number; origin: number } | undefined
let lastPointer: { x: number; y: number } | undefined
function queueRender() { if (active && !disposed && !frame) frame = requestAnimationFrame(draw) }
function draw() {
  frame = 0
  if (disposed || !active || !surface) return
  surface.render(scene.value)
  if (dragging && lastPointer && scroller.value) {
    const dx = lastPointer.x < GRID_HEADER_WIDTH ? -16 : lastPointer.x > viewport.value.width - 16 ? 16 : 0
    const dy = lastPointer.y < GRID_HEADER_HEIGHT ? -16 : lastPointer.y > viewport.value.height - 16 ? 16 : 0
    if (dx || dy) {
      const oldLeft = scroller.value.scrollLeft, oldTop = scroller.value.scrollTop
      scroller.value.scrollBy(dx, dy); onScroll()
      const hit = geometry.value.hitTest(Math.max(GRID_HEADER_WIDTH + 1, Math.min(viewport.value.width - 1, lastPointer.x)), Math.max(GRID_HEADER_HEIGHT + 1, Math.min(viewport.value.height - 1, lastPointer.y)), viewport.value)
      if (hit?.kind === 'cell') emit('extend-selection', hit.key)
      if (oldLeft !== scroller.value.scrollLeft || oldTop !== scroller.value.scrollTop) queueRender()
    }
  }
}
function resize() {
  if (!active || !scroller.value) return
  const width = scroller.value.clientWidth, height = scroller.value.clientHeight
  viewport.value = { ...viewport.value, width, height, scrollLeft: scroller.value.scrollLeft, scrollTop: scroller.value.scrollTop }
  surface?.resize(width, height, window.devicePixelRatio || 1)
  queueRender()
}
function watchDevicePixelRatio() {
  dprQuery?.removeEventListener('change', onDevicePixelRatioChange)
  dprQuery = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`)
  dprQuery.addEventListener('change', onDevicePixelRatioChange, { once: true })
}
function onDevicePixelRatioChange() {
  watchDevicePixelRatio()
  resize()
}
function onScroll() {
  if (!scroller.value) return
  const left = scroller.value.scrollLeft, top = scroller.value.scrollTop
  if (left !== viewport.value.scrollLeft || top !== viewport.value.scrollTop) viewport.value = { ...viewport.value, scrollLeft: left, scrollTop: top }
  sheetViews.set(props.sheet.id, { left, top }); queueRender()
}
function point(event: MouseEvent) { const rect = root.value!.getBoundingClientRect(); return { x: event.clientX - rect.left, y: event.clientY - rect.top } }
function onPointerDown(event: PointerEvent) {
  if ((event.target as HTMLElement).closest('input')) return
  const position = point(event), hit = geometry.value.hitTest(position.x, position.y, viewport.value)
  if (!hit) return
  if (event.button !== 0 && event.button !== 2) return
  event.preventDefault(); root.value?.focus({ preventScroll: true })
  if (hit.kind === 'resize-column' || hit.kind === 'resize-row') {
    resizeDrag = { axis: hit.kind === 'resize-column' ? 'column' : 'row', index: hit.kind === 'resize-column' ? hit.column : hit.row, size: hit.size, origin: hit.kind === 'resize-column' ? position.x : position.y }
  } else if (hit.kind === 'cell') {
    dragging = event.button === 0; lastPointer = position
    emit('select-cell', hit.key, event)
  } else emit('select-header', hit, event)
  root.value?.setPointerCapture(event.pointerId)
}
function onPointerMove(event: PointerEvent) {
  const position = point(event), hit = geometry.value.hitTest(Math.max(GRID_HEADER_WIDTH + 1, Math.min(viewport.value.width - 1, position.x)), Math.max(GRID_HEADER_HEIGHT + 1, Math.min(viewport.value.height - 1, position.y)), viewport.value)
  lastPointer = position
  if (dragging && hit?.kind === 'cell') { emit('extend-selection', hit.key); queueRender() }
  if (root.value && !dragging) {
    const hover = geometry.value.hitTest(position.x, position.y, viewport.value)
    root.value.style.cursor = resizeDrag?.axis === 'column' || hover?.kind === 'resize-column' ? 'col-resize' : resizeDrag?.axis === 'row' || hover?.kind === 'resize-row' ? 'row-resize' : 'cell'
  }
}
function onPointerUp(event: PointerEvent) {
  if (resizeDrag && event.type !== 'pointercancel') {
    const position = point(event), delta = ((resizeDrag.axis === 'column' ? position.x : position.y) - resizeDrag.origin) / viewport.value.zoom
    emit('resize-axis', resizeDrag.axis, resizeDrag.index, Math.max(resizeDrag.axis === 'column' ? 24 : 16, Math.round(resizeDrag.size + delta)))
  }
  resizeDrag = undefined
  dragging = false; lastPointer = undefined
  if (root.value?.hasPointerCapture(event.pointerId)) root.value.releasePointerCapture(event.pointerId)
}
function onDoubleClick(event: MouseEvent) { const p = point(event), hit = geometry.value.hitTest(p.x, p.y, viewport.value); if (hit?.kind === 'cell') emit('edit-cell', hit.key) }
function onInputKeydown(event: KeyboardEvent) { if (!composing.value && !event.isComposing) emit('editor-keydown', event) }
function onPaste(event: ClipboardEvent) { if (!(event.target as HTMLElement).closest('input')) emit('paste', event) }
function focusEditor() { void nextTick(() => input.value?.focus({ preventScroll: true })) }
function ensureVisible(address: string) {
  const cell = parseAddress(address), scroll = scroller.value
  if (!cell || !scroll) return
  const next = geometry.value.scrollForCell(cell.row, cell.column, viewport.value)
  scroll.scrollLeft = next.scrollLeft; scroll.scrollTop = next.scrollTop
  onScroll()
}
function focus() { root.value?.focus({ preventScroll: true }) }
defineExpose({ focusEditor, ensureVisible, focus, getBackend: () => surface?.backend })
watch(scene, queueRender)
watch(fullSize, () => { void nextTick(resize) })
watch(() => props.zoom, () => { viewport.value = { ...viewport.value, zoom: props.zoom / 100 }; void nextTick(() => { resize(); ensureVisible(props.selectedKey) }) })
watch(() => props.editingKey, key => { if (key) ensureVisible(key) })
watch(() => props.sheet.id, () => {
  void nextTick(() => {
    const view = sheetViews.get(props.sheet.id)
    if (scroller.value) { scroller.value.scrollLeft = view?.left ?? 0; scroller.value.scrollTop = view?.top ?? 0 }
    onScroll()
  })
})
onMounted(async () => {
  if (!canvasHost.value) return
  // Layout and navigation must be ready while the GPU adapter/pipeline is
  // still loading. Otherwise early name-box/keyboard navigation uses a 1px
  // viewport and silently fails to scroll to the selected cell.
  observer = new ResizeObserver(resize); observer.observe(root.value!); watchDevicePixelRatio(); window.addEventListener('resize', resize); window.visualViewport?.addEventListener('resize', resize); document.fonts?.addEventListener('loadingdone', queueRender); resize()
  try {
    const created = await createRenderSurface(canvasHost.value, { onBackendChange: (backend, reason) => emit('backend-change', backend, reason) })
    if (disposed) { created.dispose(); return }
    surface = created
    if (active) resize()
    else surface.resize(0, 0, window.devicePixelRatio || 1)
  } catch (reason) {
    surfaceError.value = reason instanceof Error ? reason.message : String(reason)
  }
})
onActivated(() => { active = true; void nextTick(resize) })
onDeactivated(() => {
  active = false
  cancelAnimationFrame(frame); frame = 0
  dragging = false; lastPointer = undefined; resizeDrag = undefined
  surface?.resize(0, 0, window.devicePixelRatio || 1)
})
onUnmounted(() => { disposed = true; cancelAnimationFrame(frame); observer?.disconnect(); dprQuery?.removeEventListener('change', onDevicePixelRatioChange); window.removeEventListener('resize', resize); window.visualViewport?.removeEventListener('resize', resize); document.fonts?.removeEventListener('loadingdone', queueRender); surface?.dispose() })
</script>

<template>
  <div ref="root" class="spreadsheet-canvas" role="grid" aria-label="工作表画布" :aria-rowcount="sheet.rowCount" :aria-colcount="sheet.columnCount" tabindex="0" @pointerdown="onPointerDown" @pointermove="onPointerMove" @pointerup="onPointerUp" @pointercancel="onPointerUp" @dblclick="onDoubleClick" @paste="onPaste">
    <div ref="scroller" class="canvas-scroll" @scroll="onScroll"><div :style="fullSize" /></div>
    <div ref="canvasHost" class="canvas-surface" :style="{ width: `${viewport.width}px`, height: `${viewport.height}px` }" />
    <input v-if="editingKey" ref="input" class="canvas-cell-input" :style="editStyle" :value="editValue" aria-label="单元格编辑" @input="emit('update:editValue', ($event.target as HTMLInputElement).value)" @blur="!composing && emit('finish-edit')" @keydown.stop="onInputKeydown" @compositionstart="composing = true" @compositionend="composing = false">
    <div class="a11y-cell" role="row"><span role="gridcell" aria-selected="true">{{ selectedKey }} {{ displayCell(sheet.cells[selectedKey]) }}</span></div>
    <div v-if="surfaceError" class="canvas-surface-error" role="alert">画布初始化失败：{{ surfaceError }}</div>
  </div>
</template>

<style scoped>
.spreadsheet-canvas { position: relative; min-width: 0; min-height: 0; flex: 1; overflow: hidden; outline: none; background: #fff; touch-action: none; user-select: none; }
.spreadsheet-canvas:focus-visible { box-shadow: inset 0 0 0 1px #246bfd; }
.canvas-scroll { position: absolute; inset: 0; overflow: auto; overscroll-behavior: contain; }
.canvas-surface { position: absolute; inset: 0; pointer-events: none; overflow: hidden; }
.canvas-cell-input { position: absolute; z-index: 3; box-sizing: border-box; margin: 0; padding: 2px 5px; border: 2px solid #246bfd; border-radius: 0; outline: none; color: #243247; background: #fff; user-select: text; }
.a11y-cell { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; pointer-events: none; }
.canvas-surface-error { position: absolute; z-index: 4; top: 12px; left: 12px; max-width: calc(100% - 24px); padding: 8px 12px; border: 1px solid #e3b3b3; color: #922f2f; background: #fff2f2; font-size: 12px; }
</style>
