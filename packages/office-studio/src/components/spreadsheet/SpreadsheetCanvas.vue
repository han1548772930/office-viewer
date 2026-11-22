<script setup lang="ts">
import { computed, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, watch } from 'vue'
import { AutoComplete, Button, theme } from 'antdv-next'
import type { RenderBackend, RenderSurface } from '../../rendering/contracts'
import type { DrawCommand } from '../../rendering/contracts'
import type { TimelineLevel } from '../../domain/model'
import { createRenderSurface } from '../../rendering/createRenderSurface'
import type { GridHit, GridRange, GridViewport, SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, SpreadsheetGeometry } from '../../rendering/spreadsheet/layout'
import { createSpreadsheetScene, displayCell } from '../../rendering/spreadsheet/scene'
import { createCellEditorLayout } from '../../rendering/spreadsheet/editor-layout'
import { fillPreviewRange, hitSelectionFillHandle, SHEET_SELECTION_STROKE_WIDTH } from '../../rendering/spreadsheet/selection'
import { detectSheetFillDownEndRow, type SheetFillApplyType } from '../../features/sheet-workbench/sheet-fill.ts'
import { canEditProtectedRanges } from '../../features/sheet-protection/range-protection.ts'
import { convertTextToNumber } from '../../features/sheet-workbench/text-to-number.ts'
import { Minus, Plus } from 'lucide-vue-next'
import { hitTestSlicerObjects, slicerObjectContentPoint, slicerObjectDrawCommands, slicerObjectLayouts, timelineSerialAtOffset, type SlicerCanvasControl, type SlicerCanvasHit } from '../../features/slicer/slicer-renderer'
import SheetContextMenu from './SheetContextMenu.vue'

const MAX_SHEET_ROWS = 1_048_576
const MAX_SHEET_COLUMNS = 16_384

const props = defineProps<{ sheet: SpreadsheetRenderSheet; selectedKey: string; selectedRange: string[]; selectedRanges?: GridRange[]; editingKey: string | null; editValue: string; zoom: number; currentUserId?: string; searchText?: string; validationOptions?: readonly string[]; overlayCommands?: readonly DrawCommand[]; slicers?: readonly SlicerCanvasControl[]; selectedSlicerId?: string; crosshairHighlight?: { enabled: boolean; color: string } }>()
const emit = defineEmits<{
  (event: 'select-cell', key: string, pointer: PointerEvent): void
  (event: 'extend-selection', key: string): void
  (event: 'fill-selection', source: GridRange, target: { row: number; column: number }, applyType?: SheetFillApplyType): void
  (event: 'select-header', hit: GridHit, pointer: PointerEvent): void
  (event: 'edit-cell', key: string): void
  (event: 'update:editValue', value: string): void
  (event: 'finish-edit'): void
  (event: 'editor-keydown', eventValue: KeyboardEvent): void
  (event: 'paste', eventValue: ClipboardEvent): void
  (event: 'resize-axis', axis: 'row' | 'column', index: number, size: number): void
  (event: 'context-command', commandId: string, params?: unknown): void
  (event: 'open-hyperlink', key: string): void
  (event: 'backend-change', backend: RenderBackend, reason?: string): void
  (event: 'select-slicer', slicerId: string): void
  (event: 'slicer-item', payload: { slicerId: string; value: string; index: number; anchorIndex?: number; ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; multiSelect: boolean }): void
  (event: 'slicer-clear', slicerId: string): void
  (event: 'slicer-range', payload: { slicerId: string; start: number; end: number; level: TimelineLevel }): void
  (event: 'slicer-scroll-to', payload: { slicerId: string; position: number }): void
  (event: 'slicer-level', payload: { slicerId: string; level: TimelineLevel }): void
  (event: 'slicer-settings', slicerId: string): void
  (event: 'slicer-move', payload: { slicerId: string; left: number; top: number }): void
  (event: 'slicer-resize', payload: { slicerId: string; width: number; height: number }): void
}>()
const root = ref<HTMLElement | null>(null), canvasHost = ref<HTMLElement | null>(null), scroller = ref<HTMLElement | null>(null), input = ref<{ focus: (options?: { preventScroll?: boolean; cursor?: 'start' | 'end' | 'all' }) => void } | null>(null)
const viewport = ref<GridViewport>({ width: 1, height: 1, scrollLeft: 0, scrollTop: 0, zoom: props.zoom / 100 })
const fillPreview = ref<GridRange>()
const geometry = computed(() => new SpreadsheetGeometry(props.sheet))
const slicerLayouts = computed(() => slicerObjectLayouts(geometry.value, props.slicers ?? [], props.sheet.date1904 ?? false))
const timelineMenuId = ref('')
const focusedSlicerItems = ref(new Map<string, number>())
const focusedTimelinePeriods = ref(new Map<string, number>())
const scene = computed(() => {
  const selectionCommands: DrawCommand[] = []
  const selected = slicerLayouts.value.find(({ id }) => id === props.selectedSlicerId)
  if (selected) selectionCommands.push({ kind: 'rect', ...selected.rect, fill: 'transparent', stroke: '#246bfd', strokeWidth: 2 })
  return createSpreadsheetScene(geometry.value, viewport.value, { activeCell: props.selectedKey, editing: Boolean(props.editingKey), selectedCells: props.selectedRange, selectedRanges: props.selectedRanges, fillPreviewRange: fillPreview.value, searchText: props.searchText, overlayCommands: [...(props.overlayCommands ?? []), ...slicerObjectDrawCommands(slicerLayouts.value, { timelineMenuId: timelineMenuId.value, selectedSlicerId: props.selectedSlicerId, multiSelect: multiSelect.value, focusedItems: focusedSlicerItems.value, focusedPeriods: focusedTimelinePeriods.value, onIconReady: queueRender }), ...selectionCommands], crosshairHighlight: props.crosshairHighlight })
})
const outlineToggleRects = computed(() => geometry.value.outlineToggleRects(viewport.value))
const outlineLevelButtons = computed(() => {
  const outline = props.sheet.outline
  if (!outline) return []
  return (['row', 'column'] as const).flatMap(axis => {
    const groups = outline.groups.filter(group => group.axis === axis)
    if (!groups.length) return []
    const maxDepth = Math.max(...groups.map(group => groups.filter(parent =>
      parent.start <= group.start && parent.end >= group.end && (parent.start < group.start || parent.end > group.end),
    ).length)) + 1
    return Array.from({ length: maxDepth }, (_, index) => ({
      axis, level: index + 1, left: 1 + index * 12, top: axis === 'row' ? 1 : 13,
      label: `${axis === 'row' ? '行' : '列'}大纲级别 ${index + 1}`,
    }))
  })
})
const fullSize = computed(() => ({ width: `${GRID_HEADER_WIDTH + geometry.value.columns.offset(props.sheet.columnCount) * props.zoom / 100}px`, height: `${GRID_HEADER_HEIGHT + geometry.value.rows.offset(props.sheet.rowCount) * props.zoom / 100}px` }))
const editLayout = computed(() => props.editingKey ? createCellEditorLayout(geometry.value, props.editingKey, viewport.value) : undefined)
const { token } = theme.useToken()
const editInputStyles = computed(() => ({
  root: {
    boxSizing: 'border-box' as const,
    height: '100%',
    minHeight: 0,
    border: `${SHEET_SELECTION_STROKE_WIDTH}px solid ${token.value.colorPrimary}`,
    borderRadius: 0,
    background: props.sheet.cells[props.editingKey ?? '']?.background ?? token.value.colorBgContainer,
    boxShadow: 'none',
  },
  input: { boxSizing: 'border-box' as const, height: '100%', padding: 0, color: 'inherit', font: 'inherit' },
}))
const editStyle = computed(() => editLayout.value ? {
  left: `${editLayout.value.rect.x}px`, top: `${editLayout.value.rect.y}px`,
  width: `${editLayout.value.rect.width}px`, height: `${editLayout.value.rect.height}px`,
  clipPath: editLayout.value.clipPath,
  borderWidth: `${SHEET_SELECTION_STROKE_WIDTH}px`,
  paddingLeft: `${Math.max(0, 6 * viewport.value.zoom - SHEET_SELECTION_STROKE_WIDTH / 2)}px`,
  fontFamily: props.sheet.cells[props.editingKey ?? '']?.fontFamily ?? 'Aptos, sans-serif',
  fontSize: `${(props.sheet.cells[props.editingKey ?? '']?.fontSize ?? 11) * 4 / 3 * viewport.value.zoom}px`,
} : { clipPath: 'inset(50%)' })
const fillHandleStyle = computed(() => editLayout.value ? {
  left: `${editLayout.value.fillHandle.x}px`, top: `${editLayout.value.fillHandle.y}px`,
  width: `${editLayout.value.fillHandle.width}px`, height: `${editLayout.value.fillHandle.height}px`,
  clipPath: editLayout.value.fillHandleClipPath,
} : { clipPath: 'inset(50%)' })
const composing = ref(false)
const surfaceError = ref('')
const contextMenuOpen = ref(false)
const contextMenuHit = ref<GridHit | null>(null)
const contextMenuPosition = ref({ x: 0, y: 0 })
const contextMenuSelection = computed(() => props.selectedRanges?.at(-1))
const canShiftHorizontally = computed(() => {
  const range = contextMenuSelection.value
  if (!range) return false
  return canEditProtectedRanges(props.sheet as SpreadsheetRenderSheet, [{
    startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: MAX_SHEET_COLUMNS - 1,
  }], props.currentUserId ?? 'local-user')
})
const canShiftVertically = computed(() => {
  const range = contextMenuSelection.value
  if (!range) return false
  return canEditProtectedRanges(props.sheet as SpreadsheetRenderSheet, [{
    startRow: range.startRow, endRow: MAX_SHEET_ROWS - 1, startCol: range.startColumn, endCol: range.endColumn,
  }], props.currentUserId ?? 'local-user')
})
const hasContextHyperlink = computed(() => contextMenuHit.value?.kind === 'cell' && Boolean(props.sheet.cells[contextMenuHit.value.key]?.hyperlink))
const canConvertSelectedText = computed(() => contextMenuOpen.value && Object.entries(props.sheet.cells).some(([key, cell]) => {
  const address = parseAddress(key)
  if (!address) return false
  const selected = props.selectedRanges?.length
    ? props.selectedRanges.some(range => address.row >= range.startRow && address.row <= range.endRow && address.column >= range.startColumn && address.column <= range.endColumn)
    : key === props.selectedKey
  return selected && Boolean(convertTextToNumber(cell))
}))
const sheetViews = new Map<string, { left: number; top: number }>()
let surface: RenderSurface | undefined, observer: ResizeObserver | undefined, frame = 0, disposed = false, active = true, dragging = false
let dprQuery: MediaQueryList | undefined
let resizeDrag: { axis: 'row' | 'column'; index: number; size: number; origin: number } | undefined
let fillDrag: { source: GridRange; target?: { row: number; column: number } } | undefined
let slicerPress: { hit: SlicerCanvasHit; start: SlicerCanvasHit; startPoint: { x: number; y: number } } | undefined
const slicerAnchor = new Map<string, number>()
const multiSelect = ref(new Set<string>())
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
  // KeepAlive detaches the DOM and can reset its native scroll offsets. Those
  // inactive events must not replace the worksheet's last visible viewport.
  if (!active || !scroller.value) return
  const left = scroller.value.scrollLeft, top = scroller.value.scrollTop
  if (left !== viewport.value.scrollLeft || top !== viewport.value.scrollTop) viewport.value = { ...viewport.value, scrollLeft: left, scrollTop: top }
  sheetViews.set(props.sheet.id, { left, top }); queueRender()
}
function restoreViewport() {
  if (!active || !scroller.value) return
  const view = sheetViews.get(props.sheet.id)
  scroller.value.scrollLeft = view?.left ?? 0
  scroller.value.scrollTop = view?.top ?? 0
  resize()
  onScroll()
}
function point(event: MouseEvent) { const rect = root.value!.getBoundingClientRect(); return { x: event.clientX - rect.left, y: event.clientY - rect.top } }
function onPointerDown(event: PointerEvent) {
  if ((event.target as HTMLElement).closest('input')) return
  const position = point(event)
  const content = slicerObjectContentPoint(position.x, position.y, viewport.value)
  const slicerHit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
  if (slicerHit) {
    if (event.button !== 0) return
    event.preventDefault()
    event.stopPropagation()
    root.value?.focus({ preventScroll: true })
    emit('select-slicer', slicerHit.id)
    if (slicerHit.type === 'item') {
      const next = new Map(focusedSlicerItems.value)
      next.set(slicerHit.id, slicerHit.index)
      focusedSlicerItems.value = next
    }
    if (slicerHit.type === 'period') {
      const next = new Map(focusedTimelinePeriods.value)
      next.set(slicerHit.id, slicerHit.index)
      focusedTimelinePeriods.value = next
    }
    slicerPress = { hit: slicerHit, start: slicerHit, startPoint: content }
    root.value?.setPointerCapture(event.pointerId)
    return
  }
  const fillRange = event.button === 0 && !props.editingKey ? hitSelectionFillHandle(geometry.value, props.selectedRanges ?? [], viewport.value, position.x, position.y) : undefined
  const hit: GridHit | undefined = fillRange ? { kind: 'fill-handle', range: fillRange } : geometry.value.hitTest(position.x, position.y, viewport.value)
  if (!hit) return
  if (event.button !== 0 && event.button !== 2) return
  event.preventDefault(); root.value?.focus({ preventScroll: true })
  if (hit.kind === 'fill-handle') {
    fillDrag = { source: hit.range }
    fillPreview.value = undefined
  } else if (hit.kind === 'outline-toggle') {
    emit('context-command', 'sheet.operation.outline-toggle-group', { axis: hit.axis, groupId: hit.groupId })
  } else if (hit.kind === 'outline-level') {
    emit('context-command', 'sheet.operation.outline-set-level', { axis: hit.axis, level: hit.level })
  } else if (hit.kind === 'resize-column' || hit.kind === 'resize-row') {
    resizeDrag = { axis: hit.kind === 'resize-column' ? 'column' : 'row', index: hit.kind === 'resize-column' ? hit.column : hit.row, size: hit.size, origin: hit.kind === 'resize-column' ? position.x : position.y }
  } else if (hit.kind === 'cell') {
    if (event.button === 0 && (event.ctrlKey || event.metaKey) && !props.editingKey && props.sheet.cells[hit.key]?.hyperlink) {
      emit('open-hyperlink', hit.key)
      return
    }
    dragging = event.button === 0; lastPointer = position
    emit('select-cell', hit.key, event)
  } else emit('select-header', hit, event)
  root.value?.setPointerCapture(event.pointerId)
}
function onPointerMove(event: PointerEvent) {
  const position = point(event), hit = geometry.value.hitTest(Math.max(GRID_HEADER_WIDTH + 1, Math.min(viewport.value.width - 1, position.x)), Math.max(GRID_HEADER_HEIGHT + 1, Math.min(viewport.value.height - 1, position.y)), viewport.value)
  lastPointer = position
  if (slicerPress && ['period', 'timelineStart', 'timelineEnd', 'scroll'].includes(slicerPress.start.type)) {
    const content = slicerObjectContentPoint(position.x, position.y, viewport.value)
    const hit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
    if (hit?.id === slicerPress.hit.id) slicerPress.hit = hit
    return
  }
  if (fillDrag && hit?.kind === 'cell') {
    fillDrag.target = { row: hit.row, column: hit.column }
    fillPreview.value = fillPreviewRange(fillDrag.source, fillDrag.target)
    queueRender()
  }
  if (dragging && hit?.kind === 'cell') { emit('extend-selection', hit.key); queueRender() }
  if (root.value && !dragging && !fillDrag) {
    const hover = geometry.value.hitTest(position.x, position.y, viewport.value)
    const fillHandle = !props.editingKey && hitSelectionFillHandle(geometry.value, props.selectedRanges ?? [], viewport.value, position.x, position.y)
    const hasHyperlink = hover?.kind === 'cell' && Boolean(props.sheet.cells[hover.key]?.hyperlink)
    root.value.style.cursor = fillHandle ? 'crosshair' : hasHyperlink || hover?.kind === 'outline-toggle' || hover?.kind === 'outline-level' ? 'pointer' : resizeDrag?.axis === 'column' || hover?.kind === 'resize-column' ? 'col-resize' : resizeDrag?.axis === 'row' || hover?.kind === 'resize-row' ? 'row-resize' : 'cell'
  }
}
function onPointerUp(event: PointerEvent) {
  if (slicerPress) {
    const press = slicerPress
    slicerPress = undefined
    if (event.type !== 'pointercancel') {
      const position = point(event)
      const content = slicerObjectContentPoint(position.x, position.y, viewport.value)
      const hit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
      if (timelineMenuId.value && hit?.type !== 'timelineMenu' && hit?.type !== 'timelineLevel') timelineMenuId.value = ''
      if (press.start.type === 'move' || press.start.type === 'resize') {
        const layout = slicerLayouts.value.find(({ id }) => id === press.start.id)
        const dx = content.x - press.startPoint.x
        const dy = content.y - press.startPoint.y
        if (layout && !layout.model.lockedPosition && (Math.abs(dx) >= 1 || Math.abs(dy) >= 1)) {
          if (press.start.type === 'move') emit('slicer-move', { slicerId: layout.id, left: Math.max(0, layout.rect.x + dx), top: Math.max(0, layout.rect.y + dy) })
          else emit('slicer-resize', { slicerId: layout.id, width: Math.max(80, layout.rect.width + dx), height: Math.max(48, layout.rect.height + dy) })
        }
      } else if (press.start.type === 'scroll' && press.start.part === 'thumb') {
        const layout = slicerLayouts.value.find(({ id }) => id === press.start.id)
        if (layout && layout.scrollbar && layout.scrollThumb) {
          const delta = layout.model.type === 'timeline' ? content.x - press.startPoint.x : content.y - press.startPoint.y
          const trackLength = layout.model.type === 'timeline' ? layout.scrollbar.width - layout.scrollThumb.width : layout.scrollbar.height - layout.scrollThumb.height
          const position = trackLength > 0 ? clampOffset(layout.scrollOffset + delta / trackLength * layout.maxScrollOffset, layout.maxScrollOffset) : 0
          if (layout.model.type === 'timeline') {
            const items = props.slicers?.find(control => control.model.id === layout.id)?.items ?? []
            const serial = timelineSerialAtOffset(items, layout.model.level ?? 'month', props.sheet.date1904 ?? false, position)
            if (serial !== undefined) emit('slicer-scroll-to', { slicerId: layout.id, position: serial })
          } else emit('slicer-scroll-to', { slicerId: layout.id, position })
        }
      } else if (hit?.id === press.start.id) {
        if (press.start.type === 'clear' && hit.type === 'clear') emit('slicer-clear', hit.id)
        else if (press.start.type === 'multiSelect' && hit.type === 'multiSelect') {
          const next = new Set(multiSelect.value)
          next.has(hit.id) ? next.delete(hit.id) : next.add(hit.id)
          multiSelect.value = next
        } else if (press.start.type === 'item' && hit.type === 'item') {
          emit('slicer-item', { slicerId: hit.id, value: hit.value, index: hit.index, anchorIndex: slicerAnchor.get(hit.id), ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, multiSelect: multiSelect.value.has(hit.id) })
          if (!event.shiftKey) slicerAnchor.set(hit.id, hit.index)
        } else if (press.start.type === 'period' && hit.type === 'period') {
          const layout = slicerLayouts.value.find(({ id }) => id === hit.id)
          const first = Math.min(press.start.index, hit.index), last = Math.max(press.start.index, hit.index)
          const periods = layout?.periods.filter(({ period }) => period.key >= first && period.key <= last).sort((a, b) => a.period.key - b.period.key)
          if (periods?.length && layout) emit('slicer-range', { slicerId: hit.id, start: periods[0]!.period.start, end: periods[periods.length - 1]!.period.end, level: layout.model.selectionLevel ?? layout.model.level ?? 'month' })
        } else if ((press.start.type === 'timelineStart' || press.start.type === 'timelineEnd') && (hit.type === 'period' || hit.type === 'timelineStart' || hit.type === 'timelineEnd')) {
          const layout = slicerLayouts.value.find(({ id }) => id === press.start.id)
          const period = hit.type === 'period' ? hit.period : layout?.periods.find(({ selected }) => selected)?.period
          if (layout && period) {
            const current = layout.model.timelineRange
            const start = press.start.type === 'timelineStart' ? period.start : Number(current?.start ?? period.start)
            const end = press.start.type === 'timelineEnd' ? period.end : Number(current?.end ?? period.end)
            if (start <= end) emit('slicer-range', { slicerId: layout.id, start, end, level: layout.model.selectionLevel ?? layout.model.level ?? 'month' })
          }
        } else if (press.start.type === 'timelineLevel' && hit.type === 'timelineLevel') {
          timelineMenuId.value = timelineMenuId.value === hit.id ? '' : hit.id
          queueRender()
        } else if (press.start.type === 'timelineMenu' && hit.type === 'timelineMenu') {
          const layout = slicerLayouts.value.find(({ id }) => id === hit.id)
          const index = layout?.menu?.findIndex(rect => content.y >= rect.y && content.y <= rect.y + rect.height) ?? -1
          const levels: TimelineLevel[] = ['year', 'quarter', 'month', 'day']
          if (index >= 0 && layout) emit('slicer-level', { slicerId: hit.id, level: levels[index]! })
          timelineMenuId.value = ''
          queueRender()
        } else if (press.start.type === 'scroll' && hit.type === 'scroll' && press.start.part === 'track') {
          const layout = slicerLayouts.value.find(({ id }) => id === hit.id)
          if (layout && layout.scrollbar) {
            const coordinate = layout.model.type === 'timeline' ? content.x - layout.scrollbar.x : content.y - layout.scrollbar.y
            const extent = layout.model.type === 'timeline' ? layout.scrollbar.width : layout.scrollbar.height
            const position = clampOffset(Math.round(coordinate / Math.max(1, extent) * layout.maxScrollOffset), layout.maxScrollOffset)
            if (layout.model.type === 'timeline') {
              const items = props.slicers?.find(control => control.model.id === layout.id)?.items ?? []
              const serial = timelineSerialAtOffset(items, layout.model.level ?? 'month', props.sheet.date1904 ?? false, position)
              if (serial !== undefined) emit('slicer-scroll-to', { slicerId: layout.id, position: serial })
            } else emit('slicer-scroll-to', { slicerId: layout.id, position })
          }
        }
        else if (press.start.type === 'timelineMenu' && hit.type === 'timelineMenu') emit('slicer-settings', hit.id)
      }
    }
    if (root.value?.hasPointerCapture(event.pointerId)) root.value.releasePointerCapture(event.pointerId)
    return
  }
  if (fillDrag && event.type !== 'pointercancel' && fillDrag.target) {
    const { source, target } = fillDrag
    const insideSource = target.row >= source.startRow && target.row <= source.endRow && target.column >= source.startColumn && target.column <= source.endColumn
    if (!insideSource) emit('fill-selection', source, target)
  }
  fillDrag = undefined
  fillPreview.value = undefined
  if (resizeDrag && event.type !== 'pointercancel') {
    const position = point(event), delta = ((resizeDrag.axis === 'column' ? position.x : position.y) - resizeDrag.origin) / viewport.value.zoom
    emit('resize-axis', resizeDrag.axis, resizeDrag.index, Math.max(resizeDrag.axis === 'column' ? 24 : 16, Math.round(resizeDrag.size + delta)))
  }
  resizeDrag = undefined
  dragging = false; lastPointer = undefined
  if (root.value?.hasPointerCapture(event.pointerId)) root.value.releasePointerCapture(event.pointerId)
}
function onDoubleClick(event: MouseEvent) {
  const p = point(event)
  const content = slicerObjectContentPoint(p.x, p.y, viewport.value)
  const slicerHit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
  if (slicerHit) { emit('slicer-settings', slicerHit.id); return }
  const source = hitSelectionFillHandle(geometry.value, props.selectedRanges ?? [], viewport.value, p.x, p.y)
  if (source) {
    const targetRow = detectSheetFillDownEndRow(source, props.sheet)
    if (targetRow !== undefined) emit('fill-selection', source, { row: targetRow, column: source.endColumn }, 'SERIES')
    return
  }
  const hit = geometry.value.hitTest(p.x, p.y, viewport.value)
  if (hit?.kind === 'cell') emit('edit-cell', hit.key)
}

function onContextMenu(event: MouseEvent) {
  event.preventDefault()
  const bounds = root.value?.getBoundingClientRect()
  if (!bounds) return
  const localX = event.clientX - bounds.left
  const localY = event.clientY - bounds.top
  const content = slicerObjectContentPoint(localX, localY, viewport.value)
  const slicerHit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
  if (slicerHit) { emit('select-slicer', slicerHit.id); emit('slicer-settings', slicerHit.id); return }
  const hit = geometry.value.hitTest(localX, localY, viewport.value)
  if (!hit || hit.kind === 'resize-row' || hit.kind === 'resize-column') { contextMenuOpen.value = false; return }
  contextMenuHit.value = hit
  contextMenuPosition.value = { x: event.clientX, y: event.clientY }
  contextMenuOpen.value = true
}
function onInputKeydown(event: KeyboardEvent) { if (!composing.value && !event.isComposing) emit('editor-keydown', event) }
function onPaste(event: ClipboardEvent) { if (!(event.target as HTMLElement).closest('input')) emit('paste', event) }
function onWheel(event: WheelEvent) {
  const position = point(event)
  const content = slicerObjectContentPoint(position.x, position.y, viewport.value)
  const hit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
  if (!hit) return
  event.preventDefault()
  const layout = slicerLayouts.value.find(item => item.id === hit.id)
  if (!layout || !event.deltaY) return
  if (layout.model.type === 'timeline') {
    const step = Math.max(1, Math.round(Math.abs(event.deltaY) / 100))
    const items = props.slicers?.find(control => control.model.id === layout.id)?.items ?? []
    const position = timelineSerialAtOffset(items, layout.model.level ?? 'month', props.sheet.date1904 ?? false, layout.scrollOffset + Math.sign(event.deltaY) * step)
    if (position !== undefined) emit('slicer-scroll-to', { slicerId: layout.id, position })
  } else {
    const columns = Math.max(1, layout.model.columnCount ?? 1)
    const step = Math.max(1, Math.round(Math.abs(event.deltaY) / Math.max(1, layout.model.rowHeight ?? 24))) * columns
    emit('slicer-scroll-to', { slicerId: layout.id, position: clampOffset(layout.scrollOffset + Math.sign(event.deltaY) * step, layout.maxScrollOffset) })
  }
}
function clampOffset(value: number, max: number): number { return Math.min(max, Math.max(0, Math.round(value))) }
function onKeydown(event: KeyboardEvent) {
  if (!props.selectedSlicerId || props.editingKey || event.isComposing) return
  const layout = slicerLayouts.value.find(({ id }) => id === props.selectedSlicerId)
  if (!layout) return
  if (layout.model.type === 'slicer') {
    const visible = layout.items.map(item => item.index)
    if (!visible.length) return
    const current = focusedSlicerItems.value.get(layout.id) ?? visible[0]!
    const columns = Math.max(1, layout.model.columnCount ?? 1)
    let next = current
    if (event.key === 'ArrowDown') next += columns
    else if (event.key === 'ArrowUp') next -= columns
    else if (event.key === 'ArrowRight') next += 1
    else if (event.key === 'ArrowLeft') next -= 1
    else if (event.key === 'Home') next = 0
    else if (event.key === 'End') next = layout.items.at(-1)!.index
    else if (event.key === 'Enter' || event.key === ' ') {
      const item = layout.items.find(candidate => candidate.index === current)
      if (item) {
        emit('slicer-item', { slicerId: layout.id, value: item.value, index: item.index, anchorIndex: slicerAnchor.get(layout.id), ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, multiSelect: multiSelect.value.has(layout.id) })
        if (!event.shiftKey) slicerAnchor.set(layout.id, item.index)
      }
      event.preventDefault(); event.stopPropagation(); return
    } else if (event.key === 'Escape') { multiSelect.value.delete(layout.id); event.preventDefault(); event.stopPropagation(); queueRender(); return }
    else return
    const bounded = Math.min(layout.items.at(-1)!.index, Math.max(0, next))
    if (bounded < visible[0]! || bounded > visible.at(-1)!) {
      const startItem = bounded > visible.at(-1)! ? Math.max(0, bounded - layout.visibleCount + columns) : Math.floor(bounded / columns) * columns
      emit('slicer-scroll-to', { slicerId: layout.id, position: startItem })
    }
    focusedSlicerItems.value = new Map(focusedSlicerItems.value).set(layout.id, bounded)
  } else {
    const periods = layout.periods
    if (event.key === 'Escape' && timelineMenuId.value === layout.id) { timelineMenuId.value = ''; event.preventDefault(); queueRender(); return }
    if (!periods.length) return
    const current = focusedTimelinePeriods.value.get(layout.id) ?? periods[0]!.period.key
    if (timelineMenuId.value === layout.id) {
      const levels: TimelineLevel[] = ['year', 'quarter', 'month', 'day']
      const index = levels.indexOf(layout.model.level ?? 'month')
      if (event.key === 'Enter' || event.key === ' ') emit('slicer-level', { slicerId: layout.id, level: levels[index]! })
      else if (event.key === 'ArrowDown' || event.key === 'ArrowRight') emit('slicer-level', { slicerId: layout.id, level: levels[Math.min(3, index + 1)]! })
      else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') emit('slicer-level', { slicerId: layout.id, level: levels[Math.max(0, index - 1)]! })
      else return
      timelineMenuId.value = ''; event.preventDefault(); event.stopPropagation(); queueRender(); return
    }
    if (event.key === 'ArrowRight' || event.key === 'ArrowLeft') {
      const delta = event.key === 'ArrowRight' ? 1 : -1
      const period = periods.find(item => item.period.key === current + delta) ?? periods[delta > 0 ? periods.length - 1 : 0]!
      focusedTimelinePeriods.value = new Map(focusedTimelinePeriods.value).set(layout.id, period.period.key)
    } else if (event.key === 'Home' || event.key === 'End') {
      const period = event.key === 'Home' ? periods[0]!.period : periods.at(-1)!.period
      focusedTimelinePeriods.value = new Map(focusedTimelinePeriods.value).set(layout.id, period.key)
    } else if (event.key === 'Enter' || event.key === ' ') {
      const period = periods.find(item => item.period.key === current)?.period
      if (period) emit('slicer-range', { slicerId: layout.id, start: period.start, end: period.end, level: layout.model.selectionLevel ?? layout.model.level ?? 'month' })
    } else return
  }
  event.preventDefault(); event.stopPropagation(); queueRender()
}
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
watch(() => props.sheet.id, () => { void nextTick(restoreViewport) })
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
onActivated(() => { active = true; void nextTick(restoreViewport) })
onDeactivated(() => {
  sheetViews.set(props.sheet.id, { left: viewport.value.scrollLeft, top: viewport.value.scrollTop })
  active = false
  cancelAnimationFrame(frame); frame = 0
  dragging = false; fillDrag = undefined; fillPreview.value = undefined; lastPointer = undefined; resizeDrag = undefined
  surface?.resize(0, 0, window.devicePixelRatio || 1)
})
onUnmounted(() => { disposed = true; cancelAnimationFrame(frame); observer?.disconnect(); dprQuery?.removeEventListener('change', onDevicePixelRatioChange); window.removeEventListener('resize', resize); window.visualViewport?.removeEventListener('resize', resize); document.fonts?.removeEventListener('loadingdone', queueRender); surface?.dispose() })
</script>

<template>
  <div ref="root" class="spreadsheet-canvas" role="grid" aria-label="工作表画布" :aria-rowcount="sheet.rowCount" :aria-colcount="sheet.columnCount" tabindex="0" @keydown="onKeydown" @pointerdown="onPointerDown" @pointermove="onPointerMove" @pointerup="onPointerUp" @pointercancel="onPointerUp" @dblclick="onDoubleClick" @contextmenu="onContextMenu" @paste="onPaste" @wheel="onWheel">
    <div ref="scroller" class="canvas-scroll" @scroll="onScroll"><div :style="fullSize" /></div>
    <div ref="canvasHost" class="canvas-surface" :style="{ width: `${viewport.width}px`, height: `${viewport.height}px` }" />
    <Button
      v-for="control in outlineToggleRects"
      :key="`outline-${control.axis}-${control.groupId}`"
      class="outline-control outline-toggle"
      :style="{ left: `${control.x}px`, top: `${control.y}px` }"
      type="default"
      size="small"
      :aria-label="`${control.axis === 'row' ? '行' : '列'}分组${control.collapsed ? '展开' : '折叠'}`"
      :title="`${control.axis === 'row' ? '行' : '列'}分组${control.collapsed ? '展开' : '折叠'}`"
      @pointerdown.stop
      @click.stop="emit('context-command', 'sheet.operation.outline-toggle-group', { axis: control.axis, groupId: control.groupId })"
    >
      <component :is="control.collapsed ? Plus : Minus" :size="10" aria-hidden="true" />
    </Button>
    <Button
      v-for="control in outlineLevelButtons"
      :key="`outline-${control.axis}-level-${control.level}`"
      class="outline-control outline-level"
      :style="{ left: `${control.left}px`, top: `${control.top}px` }"
      type="default"
      size="small"
      :aria-label="control.label"
      :title="control.label"
      @pointerdown.stop
      @click.stop="emit('context-command', 'sheet.operation.outline-set-level', { axis: control.axis, level: control.level })"
    >{{ control.level }}</Button>
    <AutoComplete v-if="editingKey" ref="input" class="canvas-cell-input" :style="editStyle" :styles="editInputStyles" variant="borderless" :value="editValue" :options="(validationOptions ?? []).map(value => ({ value, label: value }))" aria-label="单元格编辑" @change="value => emit('update:editValue', String(value ?? ''))" @blur="!composing && emit('finish-edit')" @keydown.stop="onInputKeydown" @compositionstart="composing = true" @compositionend="composing = false" />
    <span v-if="editingKey && editLayout" class="canvas-cell-fill-handle" :style="fillHandleStyle" aria-hidden="true" />
    <div class="a11y-cell" role="row"><span role="gridcell" aria-selected="true">{{ selectedKey }} {{ displayCell(sheet.cells[selectedKey]) }}</span></div>
    <div v-if="surfaceError" class="canvas-surface-error" role="alert">画布初始化失败：{{ surfaceError }}</div>
    <SheetContextMenu
      v-model:open="contextMenuOpen"
      :hit="contextMenuHit"
      :sheet-id="sheet.id"
      :x="contextMenuPosition.x"
      :y="contextMenuPosition.y"
      :can-delete-row="sheet.rowCount > 1"
      :can-delete-column="sheet.columnCount > 1"
      :frozen="Boolean(sheet.freeze?.rows || sheet.freeze?.columns)"
      :has-hidden-rows="Boolean(sheet.hiddenRows?.length)"
      :has-hidden-columns="Boolean(sheet.hiddenColumns?.length)"
      :can-convert-text="canConvertSelectedText"
      :has-hyperlink="hasContextHyperlink"
      :can-shift-horizontally="canShiftHorizontally"
      :can-shift-vertically="canShiftVertically"
      @command="(commandId, params) => emit('context-command', commandId, params)"
    />
  </div>
</template>

<style scoped>
.spreadsheet-canvas { position: relative; min-width: 0; min-height: 0; flex: 1; overflow: hidden; outline: none; background: #fff; touch-action: none; user-select: none; }
.spreadsheet-canvas:focus-visible { box-shadow: inset 0 0 0 1px #246bfd; }
.canvas-scroll { position: absolute; inset: 0; overflow: auto; overscroll-behavior: contain; }
.canvas-surface { position: absolute; inset: 0; pointer-events: none; overflow: hidden; }
.outline-control { position: absolute; z-index: 3; box-sizing: border-box; width: 12px; height: 12px; min-width: 12px; min-height: 12px; margin: 0; padding: 0; border: 1px solid #7c899a; border-radius: 1px; color: #465568; background: #fff; box-shadow: none; line-height: 10px; }
.outline-control :deep(.ant-btn-icon) { display: inline-flex; align-items: center; justify-content: center; }
.outline-level { font-size: 8px; }
.canvas-cell-input { position: absolute; z-index: 3; display: block; box-sizing: border-box; margin: 0; padding: 0; border-radius: 0; background: transparent; user-select: text; }
.canvas-cell-input :deep(.ant-select-input) { height: 100%; }
.canvas-cell-fill-handle { position: absolute; z-index: 4; box-sizing: border-box; border: 1px solid #fff; background: #246bfd; pointer-events: none; }
.a11y-cell { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; pointer-events: none; }
.canvas-surface-error { position: absolute; z-index: 4; top: 12px; left: 12px; max-width: calc(100% - 24px); padding: 8px 12px; border: 1px solid #e3b3b3; color: #922f2f; background: #fff2f2; font-size: 12px; }
</style>
