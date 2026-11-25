<script setup lang="ts">
import { computed, h, inject, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, useId, watch, type VNode } from 'vue'
import { Alert, AutoComplete, Button, Dropdown, Input, Tooltip, theme } from 'antdv-next'
import type { RenderBackend, RenderSurface } from '../../rendering/contracts'
import type { DrawCommand } from '../../rendering/contracts'
import type { CellRange, TimelineLevel } from '../../domain/model'
import { createRenderSurface } from '../../rendering/createRenderSurface'
import type { GridHit, GridRange, GridViewport, SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout'
import { cellAddress, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, SpreadsheetGeometry } from '../../rendering/spreadsheet/layout'
import { createSpreadsheetScene, displayCell, sheetNumberFormatPreviewPatternAt } from '../../rendering/spreadsheet/scene'
import type { SheetNumberFormatPreview } from '../../rendering/spreadsheet/scene'
import { createCellEditorLayout, hitCellEditorFillHandle } from '../../rendering/spreadsheet/editor-layout'
import { createSheetRichTextEditorScene, hitSheetRichTextEditor, hitSheetRichTextEditorPosition, moveSheetRichTextEditorCaret,
  sheetRichTextEditorCaret, sheetRichTextEditorWord, type SheetRichTextEditorCaretAffinity } from '../../rendering/spreadsheet/rich-text-editor.ts'
import { clampSheetRichTextEditorScroll, stepSheetRichTextEditorDragScroll } from '../../rendering/spreadsheet/rich-text-editor-scroll.ts'
import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetViewCellAt } from '../../services/excelViewAdapter.ts'
import { fillPreviewRange, hitSelectionFillHandle, SHEET_SELECTION_STROKE_WIDTH, sheetRangeDashedOutlineCommands } from '../../rendering/spreadsheet/selection'
import { detectSheetFillDownEndRow, type SheetFillApplyType } from '../../features/sheet-workbench/sheet-fill.ts'
import { resolveSheetFillTarget, sheetFillPointerTarget } from '../../features/sheet-workbench/sheet-fill-target.ts'
import type { SheetFillOptionsView } from '../../features/sheet-workbench/sheet-fill-store.ts'
import { canResumeSheetFillGesture, sheetFillGestureSelectionMatches, sheetFillGestureStillCurrent, snapshotSheetFillGesture,
  type SheetFillCommitContext, type SheetFillGestureContext } from '../../features/sheet-workbench/sheet-fill-gesture.ts'
import type { SheetFormatPainterCanvasState } from '../../features/sheet-workbench/sheet-format-painter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { sheetAxisHeaderAffectedRange, sheetAxisHeaderDestination, sheetAxisHeaderMoveSource,
  type SheetAxisHeaderMoveSource } from '../../features/sheet-workbench/sheet-axis-header-gesture.ts'
import { sheetHeaderAxisIndexAt, sheetHeaderDragCommands } from '../../rendering/spreadsheet/header-drag.ts'
import { convertTextToNumber } from '../../features/sheet-workbench/text-to-number.ts'
import { sheetCellTextAlignment } from '../../features/sheet-workbench/sheet-cell-alignment.ts'
import { replaceSheetCellInputSelection, sheetCellInputActionForKey, sheetCellNeedsMultilineEditor } from '../../features/sheet-workbench/sheet-cell-edit-input.ts'
import { createSheetCellNativeInputBridge, type SheetCellInputBridgeContext } from '../../features/sheet-workbench/sheet-cell-native-input.ts'
import type { SheetCellNativeInputEvent } from '../../features/sheet-workbench/sheet-cell-input-trace.ts'
import { sheetCellWrapStrategy } from '../../features/sheet-workbench/sheet-cell-layout.ts'
import { layoutSheetCellText, sheetCellTextFont } from '../../features/sheet-workbench/sheet-cell-text-layout.ts'
import { layoutSheetCellRichText } from '../../features/sheet-workbench/sheet-cell-rich-text-layout.ts'
import { measureSpreadsheetText } from '../../rendering/spreadsheet/text-metrics.ts'
import { ChevronDown, ClipboardPaste, Minus, Plus } from 'lucide-vue-next'
import { hitTestSlicerObjects, slicerObjectContentPoint, slicerObjectDrawCommands, slicerObjectLayouts, timelineSerialAtOffset, type SlicerCanvasControl, type SlicerCanvasHit } from '../../features/slicer/slicer-renderer'
import SheetContextMenu from './SheetContextMenu.vue'
import SheetFillOptions from './SheetFillOptions.vue'
import { sheetContextStructureTarget } from './sheet-context-menu.ts'
import { sheetAxisVisibilityTargetHasHidden } from '../../features/sheet-workbench/sheet-axis-visibility.ts'
import { OfficeTranslatorKey, useOfficeRegion } from '../../i18n/office-i18n'
import { SHEET_CLIPBOARD_COMMANDS } from '../../features/sheet-workbench/sheet-clipboard.ts'
import { sheetClipboardCommandCapability, type SheetClipboardCapabilities } from '../../features/sheet-workbench/sheet-clipboard-capabilities.ts'
import { SHEET_CLIPBOARD_PASTE_OPTIONS, sheetClipboardGridRange, sheetClipboardOverlayPosition, type SheetClipboardPasteOptionsView } from './sheet-clipboard-overlay.ts'
import { useSheetImageCanvas } from './use-sheet-image-canvas.ts'
import { SHEET_IMAGE_COMMANDS, type SheetImageCommandInput } from '../../features/sheet-images/sheet-image-commands.ts'
import { sheetImagePlacement } from '../../features/sheet-images/sheet-image-placement.ts'
import { SHEET_IMAGE_CROP_PRESETS } from '../../features/sheet-images/sheet-image-crop.ts'
import { SHEET_IMAGE_ALIGN_OPERATIONS, SHEET_IMAGE_ALIGN_PRESETS } from '../../features/sheet-images/sheet-image-selection.ts'
import type { SheetDrawingAffineGeometry, SheetDrawingLeafRef, SheetDrawingRef } from '../../domain/sheet-drawings.ts'
import { drawingPaintOrder, DRAWING_ARRANGE_MENU, DRAWING_GROUP_MENU, SHEET_DRAWING_GROUP_COMMAND_IDS, type SheetDrawingScene } from '../../features/drawing-groups/index.ts'
import { sheetDrawingGeometryRange } from '../../features/sheet-drawings/sheet-drawing-viewport.ts'
import { useSheetDrawingCanvas } from './use-sheet-drawing-canvas.ts'
import type { SheetDrawingCommandInput } from '../../api/sheet-drawing-commands.ts'

const MAX_SHEET_ROWS = 1_048_576
const MAX_SHEET_COLUMNS = 16_384

const props = defineProps<{ sheet: SpreadsheetRenderSheet; drawingScene?: SheetDrawingScene; selectedDrawingRefs?: readonly SheetDrawingRef[]; selectedImageId?: string; selectedImageIds?: readonly string[]; imageInteractionBlocked?: boolean; selectedKey: string; selectedRange: string[]; selectedRanges?: GridRange[]; editingKey: string | null; editValue: string; editSessionId?: string; richTextEditorCell?: ExcelCellView; editSelection?: { start: number; end: number; host?: 'cell' | 'formula' }; inputContext?: () => SheetCellInputBridgeContext; keepEditFocus?: (target: Element | null) => boolean; prepareFill?: () => SheetFillCommitContext | undefined; zoom: number; currentUserId?: string; workbookRevision?: number; canFormat?: boolean; searchText?: string; numberFormatPreview?: SheetNumberFormatPreview; validationOptions?: readonly string[]; overlayCommands?: readonly DrawCommand[]; slicers?: readonly SlicerCanvasControl[]; selectedSlicerId?: string; crosshairHighlight?: { enabled: boolean; color: string }; formatPainter?: SheetFormatPainterCanvasState | null; copyRange?: CellRange; pasteOptions?: SheetClipboardPasteOptionsView; fillOptions?: SheetFillOptionsView; clipboardCapabilities?: SheetClipboardCapabilities; pasteOptionsCapabilities?: SheetClipboardCapabilities }>()
const emit = defineEmits<{
  (event: 'select-cell', key: string, pointer: PointerEvent): void
  (event: 'extend-selection', key: string): void
  (event: 'selection-end', value: { canceled: boolean; ranges: GridRange[]; key: string }): void
  (event: 'fill-selection', source: GridRange, target: { row: number; column: number }, applyType?: SheetFillApplyType): void
  (event: 'refill-selection', value: { sessionId: string; applyType: SheetFillApplyType }): void
  (event: 'select-header', hit: GridHit, pointer: PointerEvent): void
  (event: 'extend-header-selection', axis: 'row' | 'column', index: number): void
  (event: 'edit-cell', key: string): void
  (event: 'update:editValue', value: string): void
  (event: 'input-event', value: SheetCellNativeInputEvent): void
  (event: 'finish-edit'): void
  (event: 'composition-change', composing: boolean): void
  (event: 'editor-keydown', eventValue: KeyboardEvent): void
  (event: 'copy', eventValue: ClipboardEvent): void
  (event: 'cut', eventValue: ClipboardEvent): void
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
  (event: 'select-image', imageId?: string, additive?: boolean, preserve?: boolean): void
  (event: 'edit-image', imageId: string, row?: number, column?: number): void
  (event: 'save-image', imageId: string): void
  (event: 'image-command', commandId: string, input: SheetImageCommandInput): void
  (event: 'select-drawing', ref?: SheetDrawingRef, additive?: boolean, preserve?: boolean): void
  (event: 'drawing-command', commandId: string, input: SheetDrawingCommandInput): void
  (event: 'chart-event', eventValue: MouseEvent | KeyboardEvent, chartId: string, view: { scrollLeft: number; scrollTop: number }): void
}>()
const root = ref<HTMLElement | null>(null), canvasHost = ref<HTMLElement | null>(null), scroller = ref<HTMLElement | null>(null), input = ref<{ focus: () => void } | null>(null)
// The synchronous callback also owns "no session" while Vue is still flushing the previous props.
const inputSessionId = () => props.inputContext ? props.inputContext().sessionId : props.editSessionId
const nativeInput = createSheetCellNativeInputBridge({ host: 'cell', sessionId: inputSessionId,
  value: () => props.inputContext?.().value ?? props.editValue, contains: element => Boolean(root.value?.contains(element) && element.closest('.canvas-cell-input')),
  emit: event => emit('input-event', event) })
const viewport = ref<GridViewport>({ width: 1, height: 1, scrollLeft: 0, scrollTop: 0, zoom: props.zoom / 100 })
const fillPreview = ref<GridRange>()
const fillHandleHovered = ref(false)
const headerDrag = ref<(SheetAxisHeaderMoveSource & { pointerStart: number; moved: boolean; copy: boolean; destination?: number; allowed: boolean; revision?: number })>()
const geometry = computed(() => new SpreadsheetGeometry(props.sheet))
const region = useOfficeRegion()
const slicerLayouts = computed(() => slicerObjectLayouts(geometry.value, props.slicers ?? [], props.sheet.date1904 ?? false, region.value))
const timelineMenuId = ref('')
const focusedSlicerItems = ref(new Map<string, number>())
const focusedTimelinePeriods = ref(new Map<string, number>())
const scene = computed(() => {
  const selectionCommands: DrawCommand[] = []
  const selected = slicerLayouts.value.find(({ id }) => id === props.selectedSlicerId)
  if (selected) selectionCommands.push({ kind: 'rect', ...selected.rect, fill: 'transparent', stroke: '#246bfd', strokeWidth: 2 })
  const overlays = (props.overlayCommands ?? []).flatMap(command => {
    if (!command.drawingRef) return [command]
    const object = props.drawingScene?.objects.find(object => object.ref.kind === command.drawingRef?.kind && object.ref.id === command.drawingRef.id)
    if (object && !drawingVisible(object.ref, object.geometry)) return []
    const preview = drawingCanvas.previewGeometry.value.get(`${command.drawingRef.kind}:${command.drawingRef.id}`)
    return command.kind === 'image' && command.drawingRef.kind === 'chart' && preview ? [{ ...command, x: 0, y: 0,
      width: preview.width, height: preview.height, matrix: { ...preview.matrix, e: preview.matrix.e + GRID_HEADER_WIDTH, f: preview.matrix.f + GRID_HEADER_HEIGHT } }] : [command]
  })
  const result = createSpreadsheetScene(geometry.value, viewport.value, { region: region.value, activeCell: props.selectedKey, editing: Boolean(props.editingKey), selectedCells: props.selectedRange, selectedRanges: props.selectedRanges, fillPreviewRange: fillPreview.value, formatPainterSource: props.formatPainter?.sourceSheetId === props.sheet.id ? props.formatPainter.sourceRange : undefined, formatPainterColor: token.value.colorPrimary, suppressFillHandle: Boolean(props.formatPainter || props.selectedImageId || props.selectedDrawingRefs?.length), searchText: props.searchText, numberFormatPreview: props.numberFormatPreview, drawingOrder: drawingOrder.value, overlayCommands: [...overlays, ...slicerObjectDrawCommands(slicerLayouts.value, { timelineMenuId: timelineMenuId.value, selectedSlicerId: props.selectedSlicerId, multiSelect: multiSelect.value, focusedItems: focusedSlicerItems.value, focusedPeriods: focusedTimelinePeriods.value, onIconReady: queueRender }), ...selectionCommands], crosshairHighlight: props.crosshairHighlight, ...imageCanvas.sceneOptions.value })
  result.commands.push(...drawingCanvas.commands.value)
  const drag = headerDrag.value
  result.commands.push(...sheetHeaderDragCommands(geometry.value, viewport.value, drag?.moved ? {
    axis: drag.axis, source: drag.interval, destination: drag.destination, allowed: drag.allowed,
  } : undefined, { source: token.value.colorFillSecondary, preview: token.value.colorFillTertiary,
    target: token.value.colorPrimary, invalid: token.value.colorError }))
  if (props.copyRange && !props.editingKey) result.commands.push(...sheetRangeDashedOutlineCommands(
    geometry.value, sheetClipboardGridRange(props.copyRange), viewport.value, token.value.colorPrimary,
  ))
  result.commands.push(...(richTextEditorScene.value?.commands ?? []))
  return result
})
const selectedCellText = computed(() => {
  const position = parseAddress(props.selectedKey)
  const preview = position ? sheetNumberFormatPreviewPatternAt(props.numberFormatPreview, position.row, position.column) : undefined
  return displayCell(position ? sheetViewCellAt(props.sheet, position.row, position.column) : undefined, props.sheet.date1904, region.value, preview)
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
const editedCell = computed(() => {
  const position = props.editingKey ? parseAddress(props.editingKey) : undefined
  return position ? sheetViewCellAt(props.sheet, position.row, position.column) : undefined
})
const editorScroll = ref({ left: 0, top: 0 })
const editorFocused = ref(false)
const editorCaretAffinity = ref<SheetRichTextEditorCaretAffinity>('downstream')
const editorRichLayout = computed(() => {
  const cell = props.richTextEditorCell, position = props.editingKey ? parseAddress(props.editingKey) : undefined
  if (!cell || cell.value !== props.editValue || !position) return undefined
  const rect = geometry.value.cellRect(position.row, position.column, viewport.value)
  return layoutSheetCellRichText(cell, props.editValue, { width: Math.max(1, rect.width - 12 * viewport.value.zoom),
    height: Infinity, wrap: sheetCellWrapStrategy(cell) === 'wrap', zoom: viewport.value.zoom, measure: measureSpreadsheetText })
})
const useRichTextEditor = computed(() => Boolean(editorRichLayout.value))
const retainedMultiline = ref(false)
const useMultilineEditor = computed(() => useRichTextEditor.value || sheetCellNeedsMultilineEditor(editedCell.value, props.editValue, retainedMultiline.value))
const editFont = computed(() => sheetCellTextFont(editedCell.value, viewport.value.zoom))
const editLayout = computed(() => {
  if (!props.editingKey) return undefined
  const position = parseAddress(props.editingKey)
  if (!useMultilineEditor.value || !position) return createCellEditorLayout(geometry.value, props.editingKey, viewport.value)
  const rect = geometry.value.cellRect(position.row, position.column, viewport.value)
  const layout = editorRichLayout.value ?? layoutSheetCellText(props.editValue, { width: Math.max(1, rect.width - 12 * viewport.value.zoom),
    lineHeight: editFont.value.lineHeight, wrap: sheetCellWrapStrategy(editedCell.value) === 'wrap',
    measure: text => measureSpreadsheetText(text, editFont.value.font) })
  return createCellEditorLayout(geometry.value, props.editingKey, viewport.value, { contentHeight: layout.height + 6 * viewport.value.zoom })
})
const { token } = theme.useToken()
const richTextEditorScene = computed(() => {
  const layout = editLayout.value, cell = props.richTextEditorCell
  if (!layout || !cell || !useRichTextEditor.value) return undefined
  const inset = SHEET_SELECTION_STROKE_WIDTH / 2
  const rect = { x: layout.rect.x + inset, y: layout.rect.y + inset,
    width: Math.max(0, layout.rect.width - 2 * inset), height: Math.max(0, layout.rect.height - 2 * inset) }
  return createSheetRichTextEditorScene(cell, rect, layout.clip, viewport.value.zoom, props.editSelection ?? { start: 0, end: 0 }, {
    background: cell.background ?? token.value.colorBgContainer, color: token.value.colorText, accent: token.value.colorPrimary,
    selectionColor: token.value.colorPrimaryBgHover, scrollLeft: editorScroll.value.left, scrollTop: editorScroll.value.top,
    composing: composing.value, focused: editorFocused.value, caretAffinity: editorCaretAffinity.value,
  })
})
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const pasteOptionsOpen = ref(false)
const pasteOptionsOwner = useId().replace(/:/g, '-')
const pasteOptionsPopupClass = `sheet-paste-options-popup-${pasteOptionsOwner}`
let pasteOptionsFocusRequest = 0, pasteOptionsFocusFrame = 0
const pasteOptionsPosition = computed(() => props.editingKey ? undefined : sheetClipboardOverlayPosition(geometry.value, viewport.value, props.pasteOptions?.range))
const pasteOptionsStyle = computed(() => pasteOptionsPosition.value
  ? { left: `${pasteOptionsPosition.value.left}px`, top: `${pasteOptionsPosition.value.top}px` } : undefined)
const fillOptionsPosition = computed(() => props.editingKey || props.formatPainter || props.fillOptions?.sheetId !== props.sheet.id
  ? undefined : sheetClipboardOverlayPosition(geometry.value, viewport.value, props.fillOptions?.range))
const fillOptionsStyle = computed(() => fillOptionsPosition.value
  ? { left: `${fillOptionsPosition.value.left}px`, top: `${fillOptionsPosition.value.top}px` } : undefined)
const pasteOptionsMenu = computed(() => ({
  items: SHEET_CLIPBOARD_PASTE_OPTIONS.flatMap(option => {
    const entry = props.pasteOptionsCapabilities && sheetClipboardCommandCapability(props.pasteOptionsCapabilities, SHEET_CLIPBOARD_COMMANDS.optionalPaste, option.mode)
    if (entry?.hidden) return []
    const label = translate(option.translationKey, option.label)
    const reason = entry?.disabled ? entry.reason : undefined
    return [{ key: option.mode, disabled: entry?.disabled, label: reason
      ? h(Tooltip, { title: `${label}：${reason}` }, { default: () => h('span', { 'aria-description': reason }, label) }) : label }]
  }),
  selectedKeys: props.pasteOptions ? [props.pasteOptions.mode] : [],
  onClick: ({ key }: { key: string | number }) => {
    const option = SHEET_CLIPBOARD_PASTE_OPTIONS.find(candidate => candidate.mode === key)
    if (!option || !props.pasteOptions) return
    const entry = props.pasteOptionsCapabilities && sheetClipboardCommandCapability(props.pasteOptionsCapabilities, SHEET_CLIPBOARD_COMMANDS.optionalPaste, option.mode)
    if (entry?.hidden || entry?.disabled) return
    pasteOptionsOpen.value = false
    emit('context-command', SHEET_CLIPBOARD_COMMANDS.optionalPaste, { mode: option.mode })
  },
}))
function ownsPasteOptions(element: Element): boolean {
  return element.closest('.sheet-paste-options')?.getAttribute('data-sheet-paste-options-owner') === pasteOptionsOwner
    || !!element.closest(`.${pasteOptionsPopupClass}`)
}
function focusPasteOptions(): void {
  const request = ++pasteOptionsFocusRequest, sheet = props.sheet, revision = props.workbookRevision
  const options = props.pasteOptions, selection = JSON.stringify(props.selectedRanges), host = root.value
  cancelAnimationFrame(pasteOptionsFocusFrame)
  void nextTick(() => {
    if (request !== pasteOptionsFocusRequest || !pasteOptionsOpen.value) return
    const attempt = (remaining: number): void => {
      pasteOptionsFocusFrame = 0
      if (request !== pasteOptionsFocusRequest || !pasteOptionsOpen.value || disposed || !active || props.editingKey
        || root.value !== host || props.sheet !== sheet || props.workbookRevision !== revision
        || props.pasteOptions !== options || JSON.stringify(props.selectedRanges) !== selection) return
      const focused = document.activeElement
      if (focused !== host && (!(focused instanceof Element) || !ownsPasteOptions(focused))) return
      const menu = document.querySelector<HTMLElement>(`.${pasteOptionsPopupClass} [role="menu"]`)
      if (menu && menu.getClientRects().length) menu.focus({ preventScroll: true })
      else if (remaining > 0) pasteOptionsFocusFrame = requestAnimationFrame(() => attempt(remaining - 1))
    }
    pasteOptionsFocusFrame = requestAnimationFrame(() => attempt(2))
  })
}
function onPasteOptionsKeydown(event: KeyboardEvent): void {
  if (event.isComposing || event.keyCode === 229) return
  if (event.key === 'Escape' && pasteOptionsOpen.value) {
    event.preventDefault(); event.stopPropagation(); onPasteOptionsOpenChange(false)
  } else if (event.key === 'ArrowDown' && !pasteOptionsOpen.value) {
    event.preventDefault(); event.stopPropagation(); onPasteOptionsOpenChange(true)
  }
}
const pasteOptionsPopup = (node: VNode) => h('div', { onKeydownCapture: onPasteOptionsKeydown }, [node])
function onPasteOptionsFocusIn(event: FocusEvent): void {
  if (pasteOptionsOpen.value && event.target instanceof Element && !ownsPasteOptions(event.target)) pasteOptionsOpen.value = false
}
function onPasteOptionsOpenChange(open: boolean): void {
  pasteOptionsOpen.value = open
  if (open) { focusPasteOptions(); return }
  const focused = document.activeElement, sheet = props.sheet, revision = props.workbookRevision, host = root.value
  const options = props.pasteOptions, selection = JSON.stringify(props.selectedRanges), request = pasteOptionsFocusRequest
  if (!(focused instanceof Element) || !ownsPasteOptions(focused)) return
  void nextTick(() => {
    if (pasteOptionsOpen.value || request !== pasteOptionsFocusRequest || disposed || !active || props.editingKey
      || root.value !== host || props.sheet !== sheet || props.workbookRevision !== revision
      || props.pasteOptions !== options || JSON.stringify(props.selectedRanges) !== selection) return
    const actual = document.activeElement
    if (actual === document.body || actual instanceof Element && ownsPasteOptions(actual)) focus()
  })
}
watch(pasteOptionsOpen, open => {
  if (open) return
  pasteOptionsFocusRequest++; cancelAnimationFrame(pasteOptionsFocusFrame); pasteOptionsFocusFrame = 0
}, { flush: 'sync' })
const themeStyle = computed(() => ({
  '--sheet-canvas-background': token.value.colorBgContainer,
  '--sheet-canvas-text': token.value.colorText,
  '--sheet-canvas-primary': token.value.colorPrimary,
  '--sheet-canvas-border': token.value.colorBorder,
}))
const editInputStyles = computed(() => ({
  root: {
    boxSizing: 'border-box' as const,
    height: '100%',
    minHeight: 0,
    border: `${SHEET_SELECTION_STROKE_WIDTH}px solid ${token.value.colorPrimary}`,
    borderRadius: 0,
    background: useRichTextEditor.value ? 'transparent' : editedCell.value?.background ?? token.value.colorBgContainer,
    boxShadow: 'none',
  },
  input: { boxSizing: 'border-box' as const, height: '100%', padding: 0,
    color: useRichTextEditor.value && !composing.value ? 'transparent' : editedCell.value?.color ?? token.value.colorText,
    caretColor: useRichTextEditor.value && !composing.value ? 'transparent' : token.value.colorPrimary,
    font: 'inherit', textAlign: sheetCellTextAlignment(editedCell.value) },
}))
const editTextAreaStyles = computed(() => ({
  root: { boxSizing: 'border-box' as const, height: '100%', minHeight: 0, border: 0, borderRadius: 0, background: 'transparent', boxShadow: 'none' },
  textarea: { ...editInputStyles.value.input, minHeight: 0, border: 0, borderRadius: 0, resize: 'none' as const,
    lineHeight: `${editFont.value.lineHeight}px`, whiteSpace: sheetCellWrapStrategy(editedCell.value) === 'wrap' ? 'pre-wrap' as const : 'pre' as const,
    overflow: useRichTextEditor.value && !composing.value ? 'hidden' : 'auto', background: 'transparent', boxShadow: 'none',
    paddingTop: `${Math.max(0, 3 * viewport.value.zoom - SHEET_SELECTION_STROKE_WIDTH / 2)}px`,
    paddingBottom: `${Math.max(0, 3 * viewport.value.zoom - SHEET_SELECTION_STROKE_WIDTH / 2)}px` },
}))
const editStyle = computed(() => editLayout.value ? {
  left: `${editLayout.value.rect.x}px`, top: `${editLayout.value.rect.y}px`,
  width: `${editLayout.value.rect.width}px`, height: `${editLayout.value.rect.height}px`,
  clipPath: editLayout.value.clipPath,
  borderWidth: `${SHEET_SELECTION_STROKE_WIDTH}px`,
  paddingLeft: `${Math.max(0, 6 * viewport.value.zoom - SHEET_SELECTION_STROKE_WIDTH / 2)}px`,
  paddingRight: `${Math.max(0, 6 * viewport.value.zoom - SHEET_SELECTION_STROKE_WIDTH / 2)}px`,
  fontFamily: editedCell.value?.fontFamily ?? 'Aptos, "Microsoft YaHei", sans-serif',
  fontSize: `${editFont.value.fontSize}px`, lineHeight: `${editFont.value.lineHeight}px`,
  fontWeight: editedCell.value?.bold ? 600 : 400,
  fontStyle: editedCell.value?.italic ? 'italic' as const : 'normal' as const,
} : { clipPath: 'inset(50%)' })
const fillHandleStyle = computed(() => editLayout.value ? {
  left: `${editLayout.value.fillHandle.x}px`, top: `${editLayout.value.fillHandle.y}px`,
  width: `${editLayout.value.fillHandle.width}px`, height: `${editLayout.value.fillHandle.height}px`,
  clipPath: editLayout.value.fillHandleClipPath,
} : { clipPath: 'inset(50%)' })
const composing = ref(false)
const drawingOrder = computed(() => props.drawingScene ? drawingPaintOrder(props.drawingScene) : undefined)
function drawingVisible(_ref: SheetDrawingLeafRef, frame: SheetDrawingAffineGeometry): boolean {
  return canViewWorksheetProtection(props.sheet, props.currentUserId ?? 'local-user')
    && canViewProtectedRanges(props.sheet, [sheetDrawingGeometryRange(geometry.value, frame)], props.currentUserId ?? 'local-user')
}
const drawingCanvas = useSheetDrawingCanvas({
  scene: () => props.drawingScene, geometry: () => geometry.value, viewport: () => viewport.value,
  selection: () => props.selectedDrawingRefs ?? [], revision: () => props.workbookRevision, userId: () => props.currentUserId ?? 'local-user',
  blocked: () => Boolean(props.editingKey || composing.value || props.formatPainter || props.imageInteractionBlocked),
  color: () => token.value.colorPrimary, background: () => token.value.colorBgContainer, visible: drawingVisible,
  select: (ref, additive, preserve) => emit('select-drawing', ref, additive, preserve), command: (id, input) => emit('drawing-command', id, input),
  focus: () => root.value?.focus({ preventScroll: true }), capture: id => root.value?.setPointerCapture(id),
  release: id => { if (root.value?.hasPointerCapture(id)) root.value.releasePointerCapture(id) }, render: queueRender,
})
const imageCanvas = useSheetImageCanvas({
  geometry: () => geometry.value, viewport: () => viewport.value, sheet: () => props.sheet,
  selectedId: () => props.selectedImageId, revision: () => props.workbookRevision, userId: () => props.currentUserId ?? 'local-user',
  selectedIds: () => props.selectedImageIds ?? (props.selectedImageId ? [props.selectedImageId] : []),
  drawingOrder: () => drawingOrder.value?.filter(ref => ref.kind === 'image').map(ref => ref.id) ?? (props.sheet.images ?? []).map(image => image.id),
  drawingPreview: id => drawingCanvas.previewGeometry.value.get(`image:${id}`),
  selectionControls: () => !drawingCanvas.sharedControls.value,
  hitAllowed: (id, position) => !props.drawingScene || drawingCanvas.pick(position)?.leaf.kind === 'image' && drawingCanvas.pick(position)?.leaf.id === id,
  editing: () => Boolean(props.editingKey || props.imageInteractionBlocked), composing: () => composing.value, painter: () => Boolean(props.formatPainter),
  color: () => token.value.colorPrimary, background: () => token.value.colorBgContainer,
  select: (id, additive, preserve) => emit('select-image', id, additive, preserve), edit: (id, row, column) => emit('edit-image', id, row, column),
  update: (id, image, expectedRevision) => emit('image-command', SHEET_IMAGE_COMMANDS.update, {
    sheetId: props.sheet.id, imageId: id, patch: { placement: sheetImagePlacement(image, geometry.value), crop: image.crop,
      rotation: image.rotation, flipX: image.flipX, flipY: image.flipY, drawingGeometry: image.drawingGeometry }, expectedRevision,
  }),
  remove: id => emit('image-command', SHEET_IMAGE_COMMANDS.remove, { sheetId: props.sheet.id, imageId: id, expectedRevision: props.workbookRevision }),
  removeMany: imageIds => emit('image-command', SHEET_IMAGE_COMMANDS.remove, { sheetId: props.sheet.id, imageIds, expectedRevision: props.workbookRevision }),
  updateBatch: (images, expectedRevision) => emit('image-command', SHEET_IMAGE_COMMANDS.updateBatch, { sheetId: props.sheet.id, images, expectedRevision }),
  crop: (id, input) => emit('image-command', id, { ...input, sheetId: props.sheet.id, expectedRevision: props.workbookRevision }),
  render: queueRender,
})
const imageSelectionIds = computed(() => props.selectedImageIds ?? (props.selectedImageId ? [props.selectedImageId] : []))
const multipleImagesSelected = computed(() => imageSelectionIds.value.length > 1)
const drawingContextOpen = ref(false), drawingContextPosition = ref({ x: 0, y: 0 })
const drawingGroupMenu = computed(() => ({ items: DRAWING_GROUP_MENU.items.map(item => ({ key: item.contentCommandId, label: item.label,
  disabled: item.action === 'group' ? !drawingCanvas.canGroup.value : !drawingCanvas.canUngroup.value })),
  onClick: ({ key }: { key: string | number }) => chooseDrawingMenu(String(key)) }))
const drawingArrangeMenu = computed(() => ({ items: DRAWING_ARRANGE_MENU.items.map(item => ({ key: item.action, label: item.label, disabled: !drawingCanvas.editable.value })),
  onClick: ({ key }: { key: string | number }) => chooseDrawingMenu(`arrange:${key}`) }))
const drawingContextMenu = computed(() => ({ selectable: false, items: [
  { key: 'group', label: '组合', children: drawingGroupMenu.value.items },
  { key: 'arrange', label: '排列', children: DRAWING_ARRANGE_MENU.items.map(item => ({ key: `arrange:${item.action}`, label: item.label, disabled: !drawingCanvas.editable.value })) },
  { key: 'remove', label: '删除', danger: true, disabled: !drawingCanvas.editable.value },
], onClick: ({ key }: { key: string | number }) => chooseDrawingMenu(String(key)) }))
function chooseDrawingMenu(key: string): void {
  drawingContextOpen.value = false
  if (key === 'remove') drawingCanvas.run('sheet.command.remove-sheet-drawings')
  else if (key === SHEET_DRAWING_GROUP_COMMAND_IDS.group && drawingCanvas.canGroup.value) drawingCanvas.run(key)
  else if (key === SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup && drawingCanvas.canUngroup.value) drawingCanvas.run(key)
  else if (key.startsWith('arrange:')) {
    const item = DRAWING_ARRANGE_MENU.items.find(item => `arrange:${item.action}` === key)
    if (item) drawingCanvas.run(SHEET_DRAWING_GROUP_COMMAND_IDS.arrange, { arrangeType: item.action })
  }
}
watch([() => props.sheet, () => props.workbookRevision, () => props.selectedDrawingRefs, () => props.imageInteractionBlocked], () => { drawingContextOpen.value = false }, { flush: 'sync' })
const imageArrangeMenu = computed(() => ({ items: [{ key: '0', label: '上移一层' }, { key: '1', label: '下移一层' },
  { key: '2', label: '置于顶层' }, { key: '3', label: '置于底层' }], onClick: (info: { key: string | number }) => {
  const arrangeType = Number(info.key)
  if (!props.selectedImageId || !imageCanvas.selectedEditable.value || ![0, 1, 2, 3].includes(arrangeType)) return
  emit('image-command', SHEET_IMAGE_COMMANDS.arrange, { sheetId: props.sheet.id, drawingIds: imageSelectionIds.value,
    arrangeType: arrangeType as 0 | 1 | 2 | 3, expectedRevision: props.workbookRevision })
} }))
const imageCropMenu = computed(() => ({ items: SHEET_IMAGE_CROP_PRESETS.map(preset => ({ key: preset.value, label: preset.label })),
  onClick: (info: { key: string | number }) => {
    if (!props.selectedImageId || multipleImagesSelected.value || !imageCanvas.selectedEditable.value) return
    const preset = SHEET_IMAGE_CROP_PRESETS.find(preset => preset.value === String(info.key))
    if (!preset) return
    emit('image-command', SHEET_IMAGE_COMMANDS.cropAuto, { sheetId: props.sheet.id, imageId: props.selectedImageId,
      cropType: preset.value, expectedRevision: props.workbookRevision })
  } }))
const imageAlignMenu = computed(() => ({ items: SHEET_IMAGE_ALIGN_PRESETS.map(preset => ({ key: preset.value, label: preset.label })),
  onClick: (info: { key: string | number }) => {
    const preset = SHEET_IMAGE_ALIGN_PRESETS.find(preset => preset.value === String(info.key))
    if (!preset || !imageCanvas.selectedEditable.value || !multipleImagesSelected.value) return
    emit('image-command', SHEET_IMAGE_ALIGN_OPERATIONS.align, { sheetId: props.sheet.id, imageIds: imageSelectionIds.value,
      alignType: preset.value, expectedRevision: props.workbookRevision })
  } }))
const surfaceError = ref('')
const contextMenuOpen = ref(false)
const contextMenuHit = ref<GridHit | null>(null)
const contextMenuPosition = ref({ x: 0, y: 0 })
const contextMenuSelection = computed(() => props.selectedRanges?.at(-1))
const contextImageRanges = computed<CellRange[]>(() => {
  const hit = contextMenuHit.value, ranges = props.selectedRanges ?? []
  if (!hit || !['cell', 'row', 'column', 'all'].includes(hit.kind)) return []
  const inside = ranges.some(range => hit.kind === 'cell' ? hit.row >= range.startRow && hit.row <= range.endRow && hit.column >= range.startColumn && hit.column <= range.endColumn
    : hit.kind === 'row' ? range.startColumn === 0 && range.endColumn === props.sheet.columnCount - 1 && hit.row >= range.startRow && hit.row <= range.endRow
      : hit.kind === 'column' && range.startRow === 0 && range.endRow === props.sheet.rowCount - 1 && hit.column >= range.startColumn && hit.column <= range.endColumn)
  if (inside) return ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }))
  return [hit.kind === 'cell' ? { startRow: hit.row, endRow: hit.row, startCol: hit.column, endCol: hit.column }
    : hit.kind === 'row' ? { startRow: hit.row, endRow: hit.row, startCol: 0, endCol: props.sheet.columnCount - 1 }
      : hit.kind === 'column' ? { startRow: 0, endRow: props.sheet.rowCount - 1, startCol: hit.column, endCol: hit.column }
        : { startRow: 0, endRow: props.sheet.rowCount - 1, startCol: 0, endCol: props.sheet.columnCount - 1 }]
})
const contextCellImage = computed(() => {
  const hit = contextMenuHit.value
  if (hit?.kind !== 'cell') return undefined
  const merge = geometry.value.mergeAt(hit.row, hit.column), row = merge?.startRow ?? hit.row, column = merge?.startColumn ?? hit.column
  const image = props.sheet.cells[cellAddress(row, column)]?.image
  return image && canViewProtectedRanges(props.sheet, [{ startRow: row, endRow: row, startCol: column, endCol: column }], props.currentUserId ?? 'local-user') ? { id: image.id, row, column } : undefined
})
const contextCellImagePositions = computed(() => Object.entries(props.sheet.cells).flatMap(([key, cell]) => {
  const point = cell.image && parseAddress(key)
  return point && contextImageRanges.value.some(range => point.row >= range.startRow && point.row <= range.endRow && point.column >= range.startCol && point.column <= range.endCol)
    ? [{ startRow: point.row, endRow: point.row, startCol: point.column, endCol: point.column }] : []
}))
const canSaveContextImages = computed(() => canViewProtectedRanges(props.sheet, contextCellImagePositions.value, props.currentUserId ?? 'local-user'))
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
function canChangeContextAxis(axis: 'row' | 'column'): boolean {
  const hit = contextMenuHit.value
  if (!hit || !canEditWorksheetProtection(props.sheet, props.currentUserId ?? 'local-user')) return false
  const target = sheetContextStructureTarget(hit, { sheetId: props.sheet.id, rowCount: props.sheet.rowCount,
    columnCount: props.sheet.columnCount, selectedRanges: props.selectedRanges }, axis)
  const first = target.params.ranges[0]
  if (!first) return false
  const range = axis === 'row' ? { ...first, endRow: MAX_SHEET_ROWS - 1 } : { ...first, endCol: MAX_SHEET_COLUMNS - 1 }
  return canEditProtectedRanges(props.sheet, [range], props.currentUserId ?? 'local-user')
}
const canStructureRows = computed(() => canChangeContextAxis('row'))
const canStructureColumns = computed(() => canChangeContextAxis('column'))
function contextAxisRanges(axis: 'row' | 'column') {
  return contextMenuHit.value ? sheetContextStructureTarget(contextMenuHit.value, { sheetId: props.sheet.id,
    rowCount: props.sheet.rowCount, columnCount: props.sheet.columnCount, selectedRanges: props.selectedRanges }, axis).params.ranges : []
}
function canChangeContextVisibility(axis: 'row' | 'column'): boolean {
  const ranges = contextAxisRanges(axis)
  return ranges.length > 0 && canEditWorksheetProtection(props.sheet, props.currentUserId ?? 'local-user')
    && canEditProtectedRanges(props.sheet, ranges, props.currentUserId ?? 'local-user')
}
const canVisibilityRows = computed(() => canChangeContextVisibility('row'))
const canVisibilityColumns = computed(() => canChangeContextVisibility('column'))
const hasContextHiddenRows = computed(() => sheetAxisVisibilityTargetHasHidden(props.sheet, 'row', contextAxisRanges('row')))
const hasContextHiddenColumns = computed(() => sheetAxisVisibilityTargetHasHidden(props.sheet, 'column', contextAxisRanges('column')))
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
let fillDrag: { source: GridRange; context: SheetFillGestureContext; startPosition: { x: number; y: number }; moved: boolean; target?: { row: number; column: number } } | undefined
let pendingFillPress: { pointerId: number; source: GridRange; origin: SheetFillGestureContext; startPosition: { x: number; y: number }; position: { x: number; y: number } } | undefined
let selectionPress: { pointerId: number; sheetId: string; painter: boolean } | undefined
let headerSelectionDrag: { axis: 'row' | 'column' } | undefined
let gesturePointerId: number | undefined
let richTextPress: { pointerId: number; sessionId?: string; sheetId: string; key: string; value: string; anchor: number } | undefined
let richTextPointer: { x: number; y: number } | undefined
let richTextScrollFrame = 0, richTextScrollTime: number | undefined
let editorAffinityOffset: number | undefined
let preferredEditorX: number | undefined
let painterDoubleClickUntil = 0
let slicerPress: { hit: SlicerCanvasHit; start: SlicerCanvasHit; startPoint: { x: number; y: number }; pane?: { frozenX: boolean; frozenY: boolean } } | undefined
const slicerAnchor = new Map<string, number>()
const multiSelect = ref(new Set<string>())
let lastPointer: { x: number; y: number } | undefined
let swappingEditor = false, editorHasPendingFocus = false, editorFocusRequest = 0, editorSwapRequest = 0
function queueRender() { if (active && !disposed && !frame) frame = requestAnimationFrame(draw) }
function draw() {
  frame = 0
  if (disposed || !active || !surface) return
  surface.render(scene.value)
  if ((dragging || fillDrag?.moved || headerDrag.value?.moved || headerSelectionDrag) && lastPointer && scroller.value) {
    const axis = headerDrag.value?.axis ?? headerSelectionDrag?.axis
    const dx = axis === 'row' ? 0 : lastPointer.x < GRID_HEADER_WIDTH ? -16 : lastPointer.x > viewport.value.width - 16 ? 16 : 0
    const dy = axis === 'column' ? 0 : lastPointer.y < GRID_HEADER_HEIGHT ? -16 : lastPointer.y > viewport.value.height - 16 ? 16 : 0
    if (dx || dy) {
      const oldLeft = scroller.value.scrollLeft, oldTop = scroller.value.scrollTop
      scroller.value.scrollBy(dx, dy)
      if (oldLeft !== scroller.value.scrollLeft || oldTop !== scroller.value.scrollTop) {
        onScroll()
        if (axis) updateHeaderGesture(lastPointer)
        else if (fillDrag) updateFillDragTarget(lastPointer)
        else {
          const hit = geometry.value.hitTest(Math.max(GRID_HEADER_WIDTH + 1, Math.min(viewport.value.width - 1, lastPointer.x)), Math.max(GRID_HEADER_HEIGHT + 1, Math.min(viewport.value.height - 1, lastPointer.y)), viewport.value)
          if (hit?.kind === 'cell') emit('extend-selection', hit.key)
        }
        queueRender()
      }
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
type DrawingHit = NonNullable<ReturnType<typeof drawingCanvas.pick>>
function drawingContentPoint(position: { x: number; y: number }, hit?: DrawingHit, capturedPane?: { frozenX: boolean; frozenY: boolean }): { x: number; y: number } {
  const pane = capturedPane ?? hit?.layout
  return slicerObjectContentPoint(position.x, position.y, pane ? { ...viewport.value,
    scrollLeft: pane.frozenX ? 0 : viewport.value.scrollLeft, scrollTop: pane.frozenY ? 0 : viewport.value.scrollTop } : viewport.value)
}
function emitChartEvent(event: MouseEvent, hit: DrawingHit): void {
  emit('chart-event', event, hit.leaf.id, { scrollLeft: hit.layout.frozenX ? 0 : viewport.value.scrollLeft, scrollTop: hit.layout.frozenY ? 0 : viewport.value.scrollTop })
}
/** Native chart hitboxes are an input bridge. Common Canvas order decides which drawing receives the event. */
function routeDrawingEvent(event: MouseEvent, chartId: string): boolean {
  if (!props.drawingScene || !root.value || !active || disposed) return false
  const position = point(event), hit = drawingCanvas.pick(position), controls = drawingCanvas.hasControlAt(position) || imageCanvas.hasControlAt(position),
    additive = event.type === 'pointerdown' && (event.ctrlKey || event.metaKey || event.shiftKey
      && !(props.selectedDrawingRefs?.length === 1 && props.selectedDrawingRefs[0]?.kind === 'chart' && props.selectedDrawingRefs[0].id === chartId))
  if (!controls && !additive && hit?.leaf.kind === 'chart' && hit.leaf.id === chartId && hit.target.kind === 'chart') return false
  event.preventDefault(); event.stopPropagation()
  if (event.type === 'pointerdown' && event instanceof PointerEvent) onPointerDown(event)
  else if (event.type === 'pointermove' && event instanceof PointerEvent) onPointerMove(event)
  else if (event.type === 'dblclick') onDoubleClick(event)
  else if (event.type === 'contextmenu') onContextMenu(event)
  else if (event.type === 'wheel' && event instanceof WheelEvent) onWheel(event)
  return true
}
function editorElement(): HTMLInputElement | HTMLTextAreaElement | undefined {
  const element = root.value?.querySelector('.canvas-cell-input input, .canvas-cell-input textarea')
  return element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement ? element : undefined
}
function selectEditorText(start: number, end = start, affinity: SheetRichTextEditorCaretAffinity = 'downstream'): void {
  const element = editorElement()
  if (!element || composing.value) return
  element.focus({ preventScroll: true })
  editorCaretAffinity.value = affinity; editorAffinityOffset = end
  element.setSelectionRange(Math.min(start, end), Math.max(start, end), start > end ? 'backward' : 'forward')
  // A library-owned input remains the authority for native offsets and the shared edit trace.
  element.dispatchEvent(new Event('select', { bubbles: true }))
}
function scrollRichEditor(left: number, top: number): void {
  const scene = richTextEditorScene.value, layout = editLayout.value
  if (!scene || !layout) return
  const next = clampSheetRichTextEditorScroll(scene, { left, top })
  if (next.left !== editorScroll.value.left || next.top !== editorScroll.value.top) editorScroll.value = next
}
function ensureRichEditorCaretVisible(): void {
  const scene = richTextEditorScene.value, layout = editLayout.value
  if (!scene || !layout || composing.value || richTextPress) return
  const element = editorElement()
  const selection = props.editSelection ?? { start: 0, end: 0 }
  const offset = element === document.activeElement && element?.selectionDirection === 'backward' ? selection.start : selection.end
  const caret = sheetRichTextEditorCaret(scene, offset)
  if (!caret) return
  const margin = 3 * viewport.value.zoom, clip = layout.clip
  const left = editorScroll.value.left + (caret.x < clip.x + margin ? caret.x - clip.x - margin
    : caret.x > clip.x + clip.width - margin ? caret.x - clip.x - clip.width + margin : 0)
  const top = editorScroll.value.top + (caret.top < clip.y + margin ? caret.top - clip.y - margin
    : caret.top + caret.height > clip.y + clip.height - margin ? caret.top + caret.height - clip.y - clip.height + margin : 0)
  scrollRichEditor(left, top)
}
function onEditorSelectionCapture(event: Event): void {
  nativeInput.capture(event, 'selection')
  const target = event.target
  if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) || !target.closest('.canvas-cell-input')) return
  editorFocused.value = target === document.activeElement
  const focus = target.selectionDirection === 'backward' ? target.selectionStart : target.selectionEnd
  const nativeMove = event.type === 'keyup' && !['ArrowUp', 'ArrowDown', 'Home', 'End'].includes((event as KeyboardEvent).key)
  if (focus !== editorAffinityOffset || nativeMove) { editorCaretAffinity.value = 'downstream'; editorAffinityOffset = undefined }
  if (event.type === 'keyup' && !['ArrowUp', 'ArrowDown'].includes((event as KeyboardEvent).key)) preferredEditorX = undefined
  void nextTick(ensureRichEditorCaretVisible)
}
function onEditorScrollCapture(event: Event): void {
  const target = event.target
  if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) || !target.closest('.canvas-cell-input')) return
  // Canvas owns rich-text scrolling because native uniform-font scroll extents can differ.
  if (!useRichTextEditor.value || composing.value) editorScroll.value = { left: target.scrollLeft, top: target.scrollTop }
}
function richPressIsCurrent(): boolean {
  return !!richTextPress && richTextPress.sessionId === inputSessionId() && richTextPress.sheetId === props.sheet.id
    && richTextPress.key === props.editingKey && richTextPress.value === props.editValue
}
function stopRichTextDragScroll(): void {
  cancelAnimationFrame(richTextScrollFrame)
  richTextScrollFrame = 0; richTextScrollTime = undefined
}
function continueRichTextDragScroll(time: number): void {
  richTextScrollFrame = 0
  if (!active || disposed || composing.value || !richPressIsCurrent()) { clearPointerGesture(); return }
  const scene = richTextEditorScene.value, position = richTextPointer
  if (!scene || !position || editorElement() !== document.activeElement) { stopRichTextDragScroll(); return }
  const elapsed = richTextScrollTime === undefined ? 1000 / 60 : time - richTextScrollTime
  richTextScrollTime = time
  const next = stepSheetRichTextEditorDragScroll(scene, editorScroll.value, position, elapsed)
  if (next.left === editorScroll.value.left && next.top === editorScroll.value.top) { stopRichTextDragScroll(); return }
  editorScroll.value = next
  const hit = hitSheetRichTextEditorPosition(richTextEditorScene.value ?? scene, position.x, position.y)
  if (hit && richTextPress) selectEditorText(richTextPress.anchor, hit.offset, hit.affinity)
  if (richPressIsCurrent()) richTextScrollFrame = requestAnimationFrame(continueRichTextDragScroll)
}
function updateRichTextPointer(position: { x: number; y: number }): void {
  if (!richPressIsCurrent() || composing.value) { clearPointerGesture(); return }
  const scene = richTextEditorScene.value, layout = editLayout.value
  if (!scene || !layout || !richTextPress) return
  richTextPointer = position
  const hit = hitSheetRichTextEditorPosition(scene, position.x, position.y)
  if (hit) selectEditorText(richTextPress.anchor, hit.offset, hit.affinity)
  const clip = layout.clip
  const outside = position.x < clip.x || position.x > clip.x + clip.width || position.y < clip.y || position.y > clip.y + clip.height
  if (outside && !richTextScrollFrame) richTextScrollFrame = requestAnimationFrame(continueRichTextDragScroll)
  else if (!outside) stopRichTextDragScroll()
}
function onEditorPointerDownCapture(event: PointerEvent): void {
  onEditorFillPointerDown(event)
  if (event.defaultPrevented || event.button !== 0 || !active || disposed || composing.value || gesturePointerId !== undefined) return
  const scene = richTextEditorScene.value, element = editorElement(), key = props.editingKey
  if (!scene || !element || !key) return
  const position = point(event), clip = scene.clip
  if (position.x < clip.x || position.x > clip.x + clip.width || position.y < clip.y || position.y > clip.y + clip.height) return
  const hit = hitSheetRichTextEditorPosition(scene, position.x, position.y)
  if (!hit) return
  const offset = hit.offset
  event.preventDefault(); event.stopPropagation()
  const anchor = event.shiftKey ? element.selectionDirection === 'backward' ? element.selectionEnd ?? offset : element.selectionStart ?? offset : offset
  preferredEditorX = undefined
  selectEditorText(anchor, offset, hit.affinity)
  richTextPress = { pointerId: event.pointerId, sessionId: inputSessionId(), sheetId: props.sheet.id, key, value: props.editValue, anchor }
  richTextPointer = position
  gesturePointerId = event.pointerId
  root.value?.setPointerCapture(event.pointerId)
}
function selectedAxisMoveSource(hit: GridHit): SheetAxisHeaderMoveSource | undefined {
  if (props.editingKey || props.formatPainter || !canEditWorksheetProtection(props.sheet, props.currentUserId ?? 'local-user')) return undefined
  const axis = hit.kind === 'row' ? 'row' : hit.kind === 'column' ? 'column' : undefined
  if (!axis) return undefined
  const ranges = (props.selectedRanges ?? []).map(range => ({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }))
  const source = sheetAxisHeaderMoveSource(props.sheet, ranges, axis, hit.kind === 'row' ? hit.row : hit.kind === 'column' ? hit.column : -1)
  return source && canEditProtectedRanges(props.sheet, [source.range], props.currentUserId ?? 'local-user') ? source : undefined
}
function updateHeaderGesture(position: { x: number; y: number }, copy?: boolean): void {
  const drag = headerDrag.value, axis = drag?.axis ?? headerSelectionDrag?.axis
  if (!axis) return
  const coordinate = axis === 'row' ? position.y : position.x
  const index = sheetHeaderAxisIndexAt(geometry.value, viewport.value, axis, coordinate)
  if (drag) {
    const moved = drag.moved || Math.abs(coordinate - drag.pointerStart) >= 3
    const destination = index === undefined ? undefined : sheetAxisHeaderDestination(drag.interval, index)
    const copying = copy ?? drag.copy
    const allowed = destination !== undefined && canEditWorksheetProtection(props.sheet, props.currentUserId ?? 'local-user')
      && canEditProtectedRanges(props.sheet, [sheetAxisHeaderAffectedRange(props.sheet, drag, destination, copying)], props.currentUserId ?? 'local-user')
    headerDrag.value = { ...drag, moved, copy: copying, destination, allowed }
    if (root.value) root.value.style.cursor = !moved ? 'grab' : allowed || destination === undefined ? copying ? 'copy' : 'grabbing' : 'not-allowed'
  } else if (index !== undefined) emit('extend-header-selection', axis, index)
  queueRender()
}
function fillGestureContext(): SheetFillGestureContext {
  return { sheetId: props.sheet.id, workbookRevision: props.workbookRevision, selectedKey: props.selectedKey,
    selectedRanges: props.selectedRanges ?? [], editingKey: props.editingKey, editSessionId: inputSessionId() }
}
function updateFillDragTarget(position: { x: number; y: number }): void {
  if (!fillDrag) return
  if (!sheetFillGestureStillCurrent(fillDrag.context, fillGestureContext()) || props.formatPainter) {
    clearPointerGesture(); return
  }
  fillDrag.moved ||= Math.hypot(position.x - fillDrag.startPosition.x, position.y - fillDrag.startPosition.y) >= 3
  if (!fillDrag.moved) return
  const hit = geometry.value.hitTest(Math.max(GRID_HEADER_WIDTH + 1, Math.min(viewport.value.width - 1, position.x)),
    Math.max(GRID_HEADER_HEIGHT + 1, Math.min(viewport.value.height - 1, position.y)), viewport.value)
  if (hit?.kind !== 'cell') return
  fillDrag.target = { row: hit.row, column: hit.column }
  fillPreview.value = fillPreviewRange(fillDrag.source, fillDrag.target)
  if (root.value) root.value.style.cursor = 'crosshair'
  queueRender()
}
function onEditorFillPointerDown(event: PointerEvent): void {
  if (event.button !== 0 || !props.editingKey || props.formatPainter || !active || disposed) return
  const position = point(event)
  const source = hitCellEditorFillHandle(geometry.value, props.editingKey, viewport.value, position.x, position.y)
  if (!source) return
  // Capture before antdv-next's input handlers and before native blur can schedule another commit.
  event.preventDefault(); event.stopPropagation()
  if (composing.value || gesturePointerId !== undefined || !props.prepareFill) return
  const press = { pointerId: event.pointerId, source: { ...source }, origin: snapshotSheetFillGesture(fillGestureContext()), startPosition: position, position }
  pendingFillPress = press
  gesturePointerId = event.pointerId
  root.value?.setPointerCapture(event.pointerId)
  let committed: SheetFillCommitContext | undefined
  try { committed = props.prepareFill() }
  catch (error) { clearPointerGesture(); throw error }
  if (!committed) { clearPointerGesture(); return }
  const accepted = committed
  void nextTick(() => {
    if (pendingFillPress !== press || gesturePointerId !== press.pointerId) return
    pendingFillPress = undefined
    if (!active || disposed || composing.value || props.formatPainter || !root.value?.hasPointerCapture(press.pointerId)
      || !canResumeSheetFillGesture(press.origin, accepted, fillGestureContext())) {
      clearPointerGesture(); return
    }
    fillDrag = { source: press.source, context: snapshotSheetFillGesture(fillGestureContext()), startPosition: press.startPosition, moved: false }
    lastPointer = press.position
    root.value.focus({ preventScroll: true })
    updateFillDragTarget(press.position)
  })
}
function onPointerDown(event: PointerEvent) {
  if (gesturePointerId !== undefined || !active || disposed) return
  if ((event.target as HTMLElement).closest('input, textarea')) return
  const position = point(event)
  if (drawingCanvas.begin(position, event)) return
  if (imageCanvas.begin(position, event)) {
    event.preventDefault(); event.stopPropagation(); root.value?.focus({ preventScroll: true })
    if (imageCanvas.hasGesture()) { gesturePointerId = event.pointerId; root.value?.setPointerCapture(event.pointerId) }
    return
  }
  const drawingHit = drawingCanvas.pick(position)
  if (drawingHit?.leaf.kind === 'chart') {
    root.value?.focus({ preventScroll: true }); emitChartEvent(event, drawingHit); return
  }
  const content = drawingContentPoint(position, drawingHit)
  const slicerHit = hitTestSlicerObjects(drawingHit ? slicerLayouts.value.filter(layout => layout.id === drawingHit.leaf.id) : slicerLayouts.value, content.x, content.y, timelineMenuId.value)
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
    slicerPress = { hit: slicerHit, start: slicerHit, startPoint: content, pane: drawingHit?.layout }
    gesturePointerId = event.pointerId
    root.value?.setPointerCapture(event.pointerId)
    return
  }
  const fillRange = event.button === 0 && !props.editingKey && !props.formatPainter ? hitSelectionFillHandle(geometry.value, props.selectedRanges ?? [], viewport.value, position.x, position.y) : undefined
  let hit: GridHit | undefined = fillRange ? { kind: 'fill-handle', range: fillRange } : geometry.value.hitTest(position.x, position.y, viewport.value)
  if (event.button === 2 && hit?.kind === 'resize-row') hit = { kind: 'row', row: hit.row }
  else if (event.button === 2 && hit?.kind === 'resize-column') hit = { kind: 'column', column: hit.column }
  if (!hit) return
  if (event.button !== 0 && event.button !== 2) return
  event.preventDefault(); root.value?.focus({ preventScroll: true })
  const headerSource = event.button === 0 && !event.shiftKey ? selectedAxisMoveSource(hit) : undefined
  if (headerSource) {
    headerDrag.value = { ...headerSource, pointerStart: headerSource.axis === 'row' ? position.y : position.x,
      moved: false, copy: event.ctrlKey || event.metaKey, allowed: true, revision: props.workbookRevision }
    lastPointer = position
    gesturePointerId = event.pointerId
    root.value?.setPointerCapture(event.pointerId)
    if (root.value) root.value.style.cursor = 'grab'
    return
  }
  if (hit.kind === 'fill-handle') {
    fillDrag = { source: { ...hit.range }, context: snapshotSheetFillGesture(fillGestureContext()), startPosition: position, moved: false }
    fillPreview.value = undefined
    lastPointer = position
    if (root.value) root.value.style.cursor = 'crosshair'
  } else if (hit.kind === 'outline-toggle') {
    emit('context-command', 'sheet.operation.outline-toggle-group', { axis: hit.axis, groupId: hit.groupId })
  } else if (hit.kind === 'outline-level') {
    emit('context-command', 'sheet.operation.outline-set-level', { axis: hit.axis, level: hit.level })
  } else if (hit.kind === 'resize-column' || hit.kind === 'resize-row') {
    resizeDrag = { axis: hit.kind === 'resize-column' ? 'column' : 'row', index: hit.kind === 'resize-column' ? hit.column : hit.row, size: hit.size, origin: hit.kind === 'resize-column' ? position.x : position.y }
  } else if (hit.kind === 'cell') {
    if (event.button === 0 && (event.ctrlKey || event.metaKey) && !props.editingKey && !props.formatPainter && props.sheet.cells[hit.key]?.hyperlink) {
      emit('open-hyperlink', hit.key)
      return
    }
    dragging = event.button === 0; lastPointer = position
    emit('select-cell', hit.key, event)
  } else {
    if (event.button === 0 && (hit.kind === 'row' || hit.kind === 'column')) {
      headerSelectionDrag = { axis: hit.kind === 'row' ? 'row' : 'column' }; lastPointer = position
    }
    emit('select-header', hit, event)
  }
  if (event.button === 0 && (hit.kind === 'cell' || hit.kind === 'row' || hit.kind === 'column' || hit.kind === 'all')) {
    selectionPress = { pointerId: event.pointerId, sheetId: props.sheet.id, painter: Boolean(props.formatPainter) }
  }
  gesturePointerId = event.pointerId
  root.value?.setPointerCapture(event.pointerId)
}
function onPointerMove(event: PointerEvent) {
  if (drawingCanvas.move(point(event), event)) return
  if (imageCanvas.move(point(event), event)) return
  if (!active || disposed || gesturePointerId !== undefined && gesturePointerId !== event.pointerId) return
  if (richTextPress) { updateRichTextPointer(point(event)); return }
  if (gesturePointerId === undefined) {
    const drawingHit = drawingCanvas.pick(point(event))
    if (drawingHit?.leaf.kind === 'chart' && drawingHit.target.kind === 'chart') { emitChartEvent(event, drawingHit); return }
    if (drawingHit?.target.kind === 'group' || drawingHit?.target.kind === 'shape') { if (root.value) root.value.style.cursor = 'move'; return }
  }
  const position = point(event), hit = geometry.value.hitTest(Math.max(GRID_HEADER_WIDTH + 1, Math.min(viewport.value.width - 1, position.x)), Math.max(GRID_HEADER_HEIGHT + 1, Math.min(viewport.value.height - 1, position.y)), viewport.value)
  lastPointer = position
  if (pendingFillPress) { pendingFillPress.position = position; return }
  if (headerDrag.value || headerSelectionDrag) { updateHeaderGesture(position, event.ctrlKey || event.metaKey); return }
  if (slicerPress && ['period', 'timelineStart', 'timelineEnd', 'scroll'].includes(slicerPress.start.type)) {
    const content = drawingContentPoint(position, undefined, slicerPress.pane)
    const hit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
    if (hit?.id === slicerPress.hit.id) slicerPress.hit = hit
    return
  }
  if (fillDrag) updateFillDragTarget(position)
  if (dragging && hit?.kind === 'cell') { emit('extend-selection', hit.key); queueRender() }
  if (root.value && !dragging && !fillDrag) {
    const hover = geometry.value.hitTest(position.x, position.y, viewport.value)
    const fillHandle = !props.formatPainter && (props.editingKey
      ? hitCellEditorFillHandle(geometry.value, props.editingKey, viewport.value, position.x, position.y)
      : hitSelectionFillHandle(geometry.value, props.selectedRanges ?? [], viewport.value, position.x, position.y))
    fillHandleHovered.value = Boolean(fillHandle)
    const hasHyperlink = hover?.kind === 'cell' && Boolean(props.sheet.cells[hover.key]?.hyperlink)
    root.value.style.cursor = fillHandle ? 'crosshair' : hasHyperlink || hover?.kind === 'outline-toggle' || hover?.kind === 'outline-level' ? 'pointer' : resizeDrag?.axis === 'column' || hover?.kind === 'resize-column' ? 'col-resize' : resizeDrag?.axis === 'row' || hover?.kind === 'resize-row' ? 'row-resize' : hover && selectedAxisMoveSource(hover) ? 'grab' : 'cell'
  }
}
function onPointerUp(event: PointerEvent) {
  if (drawingCanvas.end(point(event), event)) { clearPointerGesture(); return }
  if (imageCanvas.end(point(event), event)) { clearPointerGesture(); return }
  if (gesturePointerId !== event.pointerId) return
  if (richTextPress) {
    if (event.type === 'pointerup') updateRichTextPointer(point(event))
    clearPointerGesture(); return
  }
  if (pendingFillPress) { clearPointerGesture(); return }
  gesturePointerId = undefined
  const canceled = event.type !== 'pointerup'
  if (headerDrag.value) {
    const drag = headerDrag.value
    headerDrag.value = undefined; lastPointer = undefined
    if (root.value) root.value.style.cursor = ''
    if (root.value?.hasPointerCapture(event.pointerId)) root.value.releasePointerCapture(event.pointerId)
    queueRender()
    if (!canceled && drag.moved && drag.allowed && drag.destination !== undefined && !props.editingKey) {
      emit('context-command', `sheet.command.${drag.copy ? 'copy' : 'move'}-${drag.axis === 'row' ? 'rows' : 'cols'}`, {
        sheetId: props.sheet.id, fromRange: drag.range, destinationIndex: drag.destination, expectedRevision: drag.revision,
      })
    }
    return
  }
  if (slicerPress) {
    const press = slicerPress
    slicerPress = undefined
    if (!canceled) {
      const position = point(event)
      const content = drawingContentPoint(position, undefined, press.pane)
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
  if (fillDrag && !canceled) updateFillDragTarget(point(event))
  const completedFill = fillDrag
  fillDrag = undefined
  fillPreview.value = undefined
  if (completedFill?.moved && !canceled && completedFill.target) {
    const { source, target } = completedFill
    const range = { startRow: source.startRow, endRow: source.endRow, startCol: source.startColumn, endCol: source.endColumn }
    if (resolveSheetFillTarget(range, sheetFillPointerTarget(range, target.row, target.column))) emit('fill-selection', source, target)
  }
  if (resizeDrag && !canceled) {
    const position = point(event), delta = ((resizeDrag.axis === 'column' ? position.x : position.y) - resizeDrag.origin) / viewport.value.zoom
    emit('resize-axis', resizeDrag.axis, resizeDrag.index, Math.max(0, Math.round(resizeDrag.size + delta)))
  }
  resizeDrag = undefined
  headerSelectionDrag = undefined
  dragging = false; lastPointer = undefined
  const press = selectionPress
  if (press?.pointerId === event.pointerId) {
    selectionPress = undefined
    if (press.painter && !canceled) painterDoubleClickUntil = event.timeStamp + 600
    // Parent selection handlers update props after the native pointer event; defer application until those updates arrive.
    void nextTick(() => {
      if (!active || disposed || props.sheet.id !== press.sheetId) return
      const cell = parseAddress(props.selectedKey)
      const ranges = props.selectedRanges?.length ? props.selectedRanges.map(range => ({ ...range }))
        : cell ? [{ startRow: cell.row, endRow: cell.row, startColumn: cell.column, endColumn: cell.column }] : []
      emit('selection-end', { canceled: canceled || Boolean(props.editingKey), ranges, key: props.selectedKey })
    })
  }
  if (root.value?.hasPointerCapture(event.pointerId)) root.value.releasePointerCapture(event.pointerId)
}
function clearPointerGesture() {
  drawingCanvas.cancel()
  imageCanvas.cancel()
  const pointerId = gesturePointerId
  gesturePointerId = undefined
  richTextPress = undefined
  richTextPointer = undefined; stopRichTextDragScroll()
  selectionPress = undefined; dragging = false; fillDrag = undefined; pendingFillPress = undefined; fillPreview.value = undefined; fillHandleHovered.value = false; lastPointer = undefined; resizeDrag = undefined; slicerPress = undefined
  headerDrag.value = undefined; headerSelectionDrag = undefined
  if (root.value) root.value.style.cursor = ''
  if (pointerId !== undefined && root.value?.hasPointerCapture(pointerId)) root.value.releasePointerCapture(pointerId)
  queueRender()
}
function onDoubleClick(event: MouseEvent) {
  if (richTextEditorScene.value && props.editingKey && !composing.value) {
    const position = point(event), scene = richTextEditorScene.value, clip = scene.clip
    if (position.x >= clip.x && position.x <= clip.x + clip.width && position.y >= clip.y && position.y <= clip.y + clip.height) {
      const offset = hitSheetRichTextEditor(scene, position.x, position.y)
      if (offset !== undefined) {
        const word = sheetRichTextEditorWord(props.editValue, Math.min(offset, Math.max(0, props.editValue.length - 1)))
        event.preventDefault(); event.stopPropagation(); preferredEditorX = undefined
        selectEditorText(word.start, word.end)
      }
      return
    }
  }
  if ((event.target as HTMLElement).closest('input, textarea')) return
  if (props.formatPainter || event.timeStamp < painterDoubleClickUntil) { event.preventDefault(); return }
  const p = point(event)
  if (drawingCanvas.enterGroup(p)) { event.preventDefault(); event.stopPropagation(); return }
  const drawingHit = drawingCanvas.pick(p)
  if (drawingHit?.leaf.kind === 'chart') { emitChartEvent(event, drawingHit); return }
  if (imageCanvas.editAt(p)) { event.preventDefault(); event.stopPropagation(); return }
  const content = drawingContentPoint(p, drawingHit)
  const slicerHit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
  if (slicerHit) { emit('slicer-settings', slicerHit.id); return }
  const source = hitSelectionFillHandle(geometry.value, props.selectedRanges ?? [], viewport.value, p.x, p.y)
  if (source) {
    const targetRow = detectSheetFillDownEndRow(source, props.sheet)
    if (targetRow !== undefined) emit('fill-selection', source, { row: targetRow, column: source.endColumn })
    return
  }
  const hit = geometry.value.hitTest(p.x, p.y, viewport.value)
  if (hit?.kind === 'cell') emit('edit-cell', hit.key)
  else if (hit?.kind === 'resize-row' || hit?.kind === 'resize-column') emit('context-command',
    hit.kind === 'resize-row' ? 'sheet.command.auto-fit-row' : 'sheet.command.auto-fit-column', {
      sheetId: props.sheet.id, index: hit.kind === 'resize-row' ? hit.row : hit.column,
    })
}

function onContextMenu(event: MouseEvent) {
  if ((event.target as HTMLElement).closest('input, textarea')) return
  event.preventDefault()
  const bounds = root.value?.getBoundingClientRect()
  if (!bounds) return
  const localX = event.clientX - bounds.left
  const localY = event.clientY - bounds.top
  const drawingHit = drawingCanvas.pick({ x: localX, y: localY })
  if (drawingCanvas.selectAt({ x: localX, y: localY })) {
    contextMenuOpen.value = false; drawingContextPosition.value = { x: localX, y: localY }
    const sheetId = props.sheet.id, revision = props.workbookRevision
    void nextTick(() => { if (active && !disposed && props.sheet.id === sheetId && props.workbookRevision === revision
      && drawingCanvas.sharedControls.value && !props.imageInteractionBlocked) drawingContextOpen.value = true })
    return
  }
  if (drawingHit?.leaf.kind === 'chart') { emitChartEvent(event, drawingHit); return }
  if (imageCanvas.selectAt({ x: localX, y: localY })) { contextMenuOpen.value = false; return }
  const content = drawingContentPoint({ x: localX, y: localY }, drawingHit)
  const slicerHit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y, timelineMenuId.value)
  if (slicerHit) { emit('select-slicer', slicerHit.id); emit('slicer-settings', slicerHit.id); return }
  const actual = geometry.value.hitTest(localX, localY, viewport.value)
  const hit = actual?.kind === 'resize-row' ? { kind: 'row' as const, row: actual.row }
    : actual?.kind === 'resize-column' ? { kind: 'column' as const, column: actual.column } : actual
  if (!hit) { contextMenuOpen.value = false; return }
  contextMenuHit.value = hit
  contextMenuPosition.value = { x: event.clientX, y: event.clientY }
  contextMenuOpen.value = true
}
function changeInputComposition(value: boolean): void {
  composing.value = value
  emit('composition-change', value)
}
function onInputKeydown(event: KeyboardEvent) {
  if (!(event.target instanceof HTMLElement) || !root.value?.contains(event.target)) return
  if (!event.defaultPrevented && !composing.value && !event.isComposing && event.keyCode !== 229) emit('editor-keydown', event)
}
function onEditorKeydownCapture(event: KeyboardEvent): void {
  if (!(event.target as HTMLElement).closest('.canvas-cell-input')) return
  nativeInput.capture(event, 'selection')
  // Capture before AutoComplete's option list can interpret an IME Enter or Alt+Enter as a choice.
  if (composing.value || event.isComposing || event.keyCode === 229) { event.stopPropagation(); return }
  const scene = richTextEditorScene.value, element = editorElement()
  if (scene && element && !event.ctrlKey && !event.metaKey && !event.altKey
    && (event.key === 'ArrowUp' || event.key === 'ArrowDown' || event.key === 'Home' || event.key === 'End')) {
    const backward = element.selectionDirection === 'backward', start = element.selectionStart ?? 0, end = element.selectionEnd ?? start
    const current = backward ? start : end
    const next = moveSheetRichTextEditorCaret(scene, current, event.key, preferredEditorX === undefined ? undefined : preferredEditorX - editorScroll.value.left)
    if (next) {
      event.preventDefault(); event.stopPropagation()
      preferredEditorX = next.preferredX === undefined ? undefined : next.preferredX + editorScroll.value.left
      selectEditorText(event.shiftKey ? backward ? end : start : next.offset, next.offset, next.affinity)
    }
    return
  }
  preferredEditorX = undefined
  if (sheetCellInputActionForKey(event) === 'line-break') { event.stopPropagation(); emit('editor-keydown', event) }
}
function onEditValueChange(value: unknown): void {
  // Native events are captured before AutoComplete/Input emits change. Only the old isolated bridge uses value-only updates.
  if (inputSessionId()) return
  const draft = String(value ?? '')
  if (/[\r\n]/u.test(draft)) retainedMultiline.value = true
  emit('update:editValue', draft)
}
function onTextAreaChange(event: Event): void {
  if (event.target instanceof HTMLTextAreaElement) onEditValueChange(event.target.value)
}
function onNativeInputCapture(event: Event): void {
  const target = event.target
  if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
    if (target.closest('.canvas-cell-input') && /[\r\n]/u.test(target.value)) retainedMultiline.value = true
  }
  nativeInput.capture(event, 'input')
  editorCaretAffinity.value = 'downstream'; editorAffinityOffset = undefined
  preferredEditorX = undefined
  void nextTick(ensureRichEditorCaretVisible)
}
function onCompositionCapture(event: CompositionEvent, value: boolean): void {
  const target = event.target
  if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) || !target.closest('.canvas-cell-input')) return
  nativeInput.capture(event, value ? 'compositionstart' : 'compositionend')
  if (value && richTextPress) clearPointerGesture()
  changeInputComposition(value)
  if (value) editorScroll.value = { left: target.scrollLeft, top: target.scrollTop }
  else void nextTick(ensureRichEditorCaretVisible)
}
function onAutoCompleteSelect(value: unknown): void {
  if (!inputSessionId()) { onEditValueChange(value); return }
  const element = root.value?.querySelector('.canvas-cell-input input, .canvas-cell-input textarea')
  if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) {
    nativeInput.replace(element, String(value ?? ''), undefined, 'insertReplacementText')
  }
}
function onEditorPasteCapture(event: ClipboardEvent): void {
  const target = event.target
  if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) || !target.closest('.canvas-cell-input')) return
  nativeInput.capture(event, 'selection')
  const text = event.clipboardData?.getData('text/plain')
  if (!text || !/[\r\n]/u.test(text) || composing.value) return
  event.preventDefault(); event.stopPropagation()
  const selection = { start: target.selectionStart ?? target.value.length, end: target.selectionEnd ?? target.value.length }
  const replacement = inputSessionId() ? nativeInput.replace(target, text, selection, 'insertFromPaste', target.value)
    : replaceSheetCellInputSelection(target.value, text, selection)
  if (!replacement) return
  if (!inputSessionId()) onEditValueChange(replacement.value)
  retainedMultiline.value = true
  focusEditor(replacement.selection)
}
function onEditorBlur(event: FocusEvent): void {
  if (!(event.target instanceof HTMLElement) || !root.value?.contains(event.target)) return
  editorFocused.value = false
  if (richTextPress) clearPointerGesture()
  const next = event.relatedTarget
  if (next instanceof HTMLElement && next.closest('.canvas-cell-input, .formula-toolbar')) return
  if (props.keepEditFocus?.(next instanceof Element ? next : null)) return
  const key = props.editingKey, sessionId = inputSessionId()
  void nextTick(() => {
    if (!key || props.editingKey !== key || inputSessionId() !== sessionId || composing.value || swappingEditor) return
    const focused = document.activeElement
    if (focused instanceof HTMLElement && focused.closest('.canvas-cell-input, .formula-toolbar')) return
    if (props.keepEditFocus?.(focused)) return
    emit('finish-edit')
  })
}
function onClipboard(event: ClipboardEvent, kind: 'copy' | 'cut' | 'paste'): void {
  if (event.defaultPrevented || !active || disposed || !(event.target instanceof Element)
    || event.target.closest('input, textarea, [contenteditable]:not([contenteditable="false"])')) return
  if (kind === 'copy') emit('copy', event)
  else if (kind === 'cut') emit('cut', event)
  else emit('paste', event)
}
function onWheel(event: WheelEvent) {
  if (richTextEditorScene.value && !composing.value && (event.target as HTMLElement).closest('.canvas-cell-input')) {
    event.preventDefault()
    const scale = event.deltaMode === 1 ? editFont.value.lineHeight : event.deltaMode === 2 ? editLayout.value?.clip.height ?? 1 : 1
    scrollRichEditor(editorScroll.value.left + (event.deltaX || (event.shiftKey ? event.deltaY : 0)) * scale,
      editorScroll.value.top + (event.shiftKey ? 0 : event.deltaY * scale))
    return
  }
  if ((event.target as HTMLElement).closest('input, textarea')) return
  const position = point(event)
  const drawingHit = drawingCanvas.pick(position)
  if (drawingHit?.leaf.kind === 'chart') { emitChartEvent(event, drawingHit); return }
  const content = drawingContentPoint(position, drawingHit)
  const hit = hitTestSlicerObjects(drawingHit ? slicerLayouts.value.filter(layout => layout.id === drawingHit.leaf.id) : slicerLayouts.value, content.x, content.y, timelineMenuId.value)
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
  if ((event.key === 'ContextMenu' || event.key === 'F10' && event.shiftKey) && drawingCanvas.sharedControls.value && !event.isComposing && event.keyCode !== 229) {
    const menu = drawingCanvas.menuPosition.value
    if (menu) { event.preventDefault(); event.stopPropagation(); drawingContextPosition.value = { x: menu.left, y: menu.top + 34 }; drawingContextOpen.value = true; return }
  }
  if (drawingCanvas.keydown(event)) return
  if (imageCanvas.keydown(event)) return
  const selectedChart = props.selectedDrawingRefs?.length === 1 && props.selectedDrawingRefs[0]?.kind === 'chart' ? props.selectedDrawingRefs[0] : undefined
  if (selectedChart && !props.editingKey && !props.imageInteractionBlocked) {
    emit('chart-event', event, selectedChart.id, { scrollLeft: viewport.value.scrollLeft, scrollTop: viewport.value.scrollTop })
    if (event.defaultPrevented) return
  }
  if (event.key === 'Escape' && !event.isComposing && event.keyCode !== 229 && (fillDrag || pendingFillPress || headerDrag.value || headerSelectionDrag)) {
    event.preventDefault(); event.stopPropagation(); clearPointerGesture(); return
  }
  if (!props.selectedSlicerId || props.editingKey || event.isComposing || event.keyCode === 229) return
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
function focusEditor(selection?: { start: number; end: number }) {
  const key = props.editingKey, sessionId = inputSessionId(), request = ++editorFocusRequest
  editorHasPendingFocus = true
  void nextTick(() => {
    if (!key || props.editingKey !== key || inputSessionId() !== sessionId || request !== editorFocusRequest || !active || disposed) return
    editorHasPendingFocus = false
    input.value?.focus()
    // AutoComplete exposes focus but no selection API. Both native elements are owned by antdv-next.
    const element = root.value?.querySelector('.canvas-cell-input input, .canvas-cell-input textarea')
    if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) {
      element.focus({ preventScroll: true })
      if (selection && !composing.value) element.setSelectionRange(selection.start, selection.end)
      element.dispatchEvent(new Event('select', { bubbles: true }))
    }
  })
}
function insertEditorText(text: string): void {
  const element = root.value?.querySelector('.canvas-cell-input input, .canvas-cell-input textarea')
  if (!(element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement)) return
  const selection = { start: element.selectionStart ?? element.value.length, end: element.selectionEnd ?? element.value.length }
  const replacement = inputSessionId() ? nativeInput.replace(element, text, selection, 'insertLineBreak', element.value)
    : replaceSheetCellInputSelection(element.value, text, selection)
  if (!replacement) return
  if (!inputSessionId()) onEditValueChange(replacement.value)
  if (/[\r\n]/u.test(replacement.value)) retainedMultiline.value = true
  focusEditor(replacement.selection)
}
function ensureVisible(address: string) {
  const cell = parseAddress(address), scroll = scroller.value
  if (!cell || !scroll) return
  const next = geometry.value.scrollForCell(cell.row, cell.column, viewport.value)
  scroll.scrollLeft = next.scrollLeft; scroll.scrollTop = next.scrollTop
  onScroll()
}
function focus() { root.value?.focus({ preventScroll: true }) }
defineExpose({ focusEditor, insertEditorText, ensureVisible, focus, getBackend: () => surface?.backend, operateImageCrop: imageCanvas.operateImageCrop, routeDrawingEvent })
watch(scene, queueRender)
watch(fullSize, () => { void nextTick(resize) })
watch(() => props.zoom, () => { viewport.value = { ...viewport.value, zoom: props.zoom / 100 }; void nextTick(() => { resize(); ensureVisible(props.selectedKey) }) })
watch([() => props.editingKey, () => props.editSessionId], ([key]) => {
  changeInputComposition(false); retainedMultiline.value = false; editorHasPendingFocus = false; editorFocusRequest++
  editorScroll.value = { left: 0, top: 0 }; editorFocused.value = false; preferredEditorX = undefined
  editorCaretAffinity.value = 'downstream'; editorAffinityOffset = undefined
  if (richTextPress) clearPointerGesture()
  if (key) ensureVisible(key)
}, { flush: 'sync' })
watch(() => props.editValue, draft => {
  if (props.editingKey && /[\r\n]/u.test(draft)) retainedMultiline.value = true
  if (richTextPress && !richPressIsCurrent()) clearPointerGesture()
}, { flush: 'sync' })
watch([() => props.richTextEditorCell, () => props.editSelection, () => props.zoom], () => {
  void nextTick(ensureRichEditorCaretVisible)
})
watch(useMultilineEditor, () => {
  if (!props.editingKey) return
  const element = root.value?.querySelector('.canvas-cell-input input, .canvas-cell-input textarea')
  const focused = element === document.activeElement
  const selection = element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement
    ? { start: element.selectionStart ?? element.value.length, end: element.selectionEnd ?? element.value.length } : undefined
  const key = props.editingKey, sessionId = inputSessionId(), request = ++editorSwapRequest
  const focusRequest = editorFocusRequest, preserveSelection = !editorHasPendingFocus
  swappingEditor = true
  void nextTick(() => {
    if (key === props.editingKey && sessionId === inputSessionId() && focused && preserveSelection && focusRequest === editorFocusRequest) focusEditor(selection)
    void nextTick(() => { if (request === editorSwapRequest) swappingEditor = false })
  })
})
watch(() => props.sheet.id, () => {
  pasteOptionsOpen.value = false
  clearPointerGesture()
  void nextTick(restoreViewport)
}, { flush: 'sync' })
watch([() => props.sheet, () => props.workbookRevision, () => props.editingKey, () => props.editSessionId], () => {
  // The synchronous commit updates these props together. Validate its acknowledgement after that flush.
  if (pendingFillPress && pendingFillPress.origin.sheetId === props.sheet.id) return
  clearPointerGesture()
}, { flush: 'sync' })
watch([() => props.selectedKey, () => props.selectedRanges], () => {
  const current = fillGestureContext()
  if (pendingFillPress && !sheetFillGestureSelectionMatches(pendingFillPress.origin, current)
    || fillDrag && !sheetFillGestureStillCurrent(fillDrag.context, current)) clearPointerGesture()
}, { flush: 'sync' })
watch([() => props.zoom, () => props.formatPainter], () => {
  if (fillDrag || pendingFillPress || richTextPress) clearPointerGesture()
}, { flush: 'sync' })
watch([pasteOptionsPosition, () => props.pasteOptions], () => { pasteOptionsOpen.value = false })
onMounted(async () => {
  if (!canvasHost.value) return
  document.addEventListener('focusin', onPasteOptionsFocusIn, true)
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
  pasteOptionsOpen.value = false
  sheetViews.set(props.sheet.id, { left: viewport.value.scrollLeft, top: viewport.value.scrollTop })
  active = false
  cancelAnimationFrame(frame); frame = 0
  clearPointerGesture()
  surface?.resize(0, 0, window.devicePixelRatio || 1)
})
onUnmounted(() => { disposed = true; pasteOptionsOpen.value = false; document.removeEventListener('focusin', onPasteOptionsFocusIn, true); clearPointerGesture(); cancelAnimationFrame(frame); observer?.disconnect(); dprQuery?.removeEventListener('change', onDevicePixelRatioChange); window.removeEventListener('resize', resize); window.visualViewport?.removeEventListener('resize', resize); document.fonts?.removeEventListener('loadingdone', queueRender); surface?.dispose() })
</script>

<template>
  <div ref="root" class="spreadsheet-canvas"
    :class="{ 'format-painter-active': Boolean(formatPainter), 'fill-handle-hovered': fillHandleHovered, 'rich-text-editor-active': useRichTextEditor && !composing }"
    :style="themeStyle" role="grid" aria-label="工作表画布" :aria-rowcount="sheet.rowCount" :aria-colcount="sheet.columnCount" tabindex="0"
    @keydown.capture="onEditorKeydownCapture" @keydown="onKeydown" @keyup.capture="onEditorSelectionCapture"
    @beforeinput.capture="nativeInput.capture($event, 'beforeinput')" @input.capture="onNativeInputCapture"
    @compositionstart.capture="onCompositionCapture($event, true)" @compositionend.capture="onCompositionCapture($event, false)"
    @focusin.capture="onEditorSelectionCapture" @select.capture="onEditorSelectionCapture" @scroll.capture="onEditorScrollCapture"
    @pointerdown.capture="onEditorPointerDownCapture" @pointerdown="onPointerDown" @pointermove="onPointerMove" @pointerup="onPointerUp"
    @pointercancel="onPointerUp" @lostpointercapture="onPointerUp" @dblclick="onDoubleClick" @contextmenu="onContextMenu"
    @copy="event => onClipboard(event, 'copy')" @cut="event => onClipboard(event, 'cut')"
    @paste.capture="onEditorPasteCapture" @paste="event => onClipboard(event, 'paste')" @wheel="onWheel">
    <div ref="scroller" class="canvas-scroll" @scroll="onScroll"><div :style="fullSize" /></div>
    <div ref="canvasHost" class="canvas-surface" :style="{ width: `${viewport.width}px`, height: `${viewport.height}px` }" />
    <div v-if="drawingCanvas.menuPosition.value" class="sheet-drawing-float-menu" role="toolbar" aria-label="绘图组合工具"
      :style="{ left: `${drawingCanvas.menuPosition.value.left}px`, top: `${drawingCanvas.menuPosition.value.top}px`, background: token.colorBgElevated, borderColor: token.colorBorder }"
      @pointerdown.stop @dblclick.stop @contextmenu.prevent.stop @keydown.stop>
      <Dropdown :trigger="['click']" :menu="drawingGroupMenu" :auto-focus="false" destroy-on-hidden>
        <Button size="small" :disabled="!drawingCanvas.canGroup.value && !drawingCanvas.canUngroup.value" aria-label="组合或取消组合">组合</Button>
      </Dropdown>
      <Dropdown :trigger="['click']" :menu="drawingArrangeMenu" :auto-focus="false" destroy-on-hidden>
        <Button size="small" :disabled="!drawingCanvas.editable.value" aria-label="排列绘图对象">排列</Button>
      </Dropdown>
      <Dropdown v-if="multipleImagesSelected && imageSelectionIds.length === selectedDrawingRefs?.length" :trigger="['click']" :menu="imageAlignMenu" :auto-focus="false" destroy-on-hidden>
        <Button size="small" :disabled="!imageCanvas.selectedEditable.value" aria-label="对齐图片">对齐</Button>
      </Dropdown>
      <Tooltip title="删除当前选择的绘图对象"><Button size="small" danger :disabled="!drawingCanvas.editable.value" @click="chooseDrawingMenu('remove')">删除</Button></Tooltip>
    </div>
    <div v-if="drawingContextOpen" class="sheet-drawing-context-anchor" :style="{ left: `${drawingContextPosition.x}px`, top: `${drawingContextPosition.y}px` }">
      <Dropdown :trigger="[]" :open="drawingContextOpen" :menu="drawingContextMenu" auto-focus destroy-on-hidden @update:open="drawingContextOpen = $event">
        <span aria-hidden="true" />
      </Dropdown>
    </div>
    <Alert v-if="drawingCanvas.error.value" class="sheet-drawing-error" type="error" :title="drawingCanvas.error.value" show-icon closable @close="drawingCanvas.error.value = ''" @pointerdown.stop />
    <div v-if="imageCanvas.menuPosition.value && selectedImageId" class="sheet-image-float-menu" role="toolbar" aria-label="图片工具"
      :style="{ left: `${imageCanvas.menuPosition.value.left}px`, top: `${imageCanvas.menuPosition.value.top}px`, background: token.colorBgElevated, borderColor: token.colorBorder }"
      @pointerdown.stop @dblclick.stop @contextmenu.prevent.stop @keydown.stop>
      <Tooltip v-if="!imageCanvas.isCropping.value && !multipleImagesSelected" :title="translate('sheet.image.properties', '图片属性')"><Button size="small" :disabled="!imageCanvas.selectedEditable.value" @click="emit('edit-image', selectedImageId)">{{ translate('sheet.image.properties', '图片属性') }}</Button></Tooltip>
      <Tooltip v-if="!imageCanvas.isCropping.value && !multipleImagesSelected" :title="translate('sheet.image.save', '保存图片')"><Button size="small" @click="emit('save-image', selectedImageId)">{{ translate('sheet.image.save', '保存图片') }}</Button></Tooltip>
      <Dropdown v-if="!multipleImagesSelected" :trigger="['click']" :menu="imageCropMenu" :disabled="!imageCanvas.selectedEditable.value" :auto-focus="false" :destroy-on-hidden="true">
        <Button size="small" :disabled="!imageCanvas.selectedEditable.value" aria-label="裁剪图片">裁剪</Button>
      </Dropdown>
      <Button v-if="imageCanvas.isCropping.value" size="small" type="primary" @click="emit('image-command', SHEET_IMAGE_COMMANDS.cropClose, { sheetId: sheet.id, expectedRevision: workbookRevision })">完成裁剪</Button>
      <Button v-if="imageCanvas.isCropping.value" size="small" @click="emit('image-command', SHEET_IMAGE_COMMANDS.cropClose, { sheetId: sheet.id, isCancel: true })">取消</Button>
      <Dropdown v-if="!imageCanvas.isCropping.value" :trigger="['click']" :menu="imageArrangeMenu" :disabled="!imageCanvas.selectedEditable.value" :auto-focus="false" :destroy-on-hidden="true">
        <Button size="small" :disabled="!imageCanvas.selectedEditable.value" aria-label="排列图片">排列</Button>
      </Dropdown>
      <Dropdown v-if="!imageCanvas.isCropping.value && multipleImagesSelected" :trigger="['click']" :menu="imageAlignMenu" :disabled="!imageCanvas.selectedEditable.value" :auto-focus="false" :destroy-on-hidden="true">
        <Button size="small" :disabled="!imageCanvas.selectedEditable.value" aria-label="对齐图片">对齐</Button>
      </Dropdown>
      <Tooltip v-if="!imageCanvas.isCropping.value" :title="translate('sheet.image.delete', '删除图片')"><Button size="small" danger :disabled="!imageCanvas.selectedEditable.value" @click="emit('image-command', SHEET_IMAGE_COMMANDS.remove, { sheetId: sheet.id, imageIds: imageSelectionIds, expectedRevision: workbookRevision })">{{ translate('sheet.image.delete', '删除图片') }}</Button></Tooltip>
    </div>
    <Alert v-if="imageCanvas.imageError.value" class="sheet-image-error" type="error" :title="imageCanvas.imageError.value" show-icon closable @close="imageCanvas.imageError.value = ''" @pointerdown.stop />
    <SheetFillOptions v-if="fillOptionsStyle && fillOptions" :options="fillOptions" :position="fillOptionsStyle"
      :revision="workbookRevision" :selection-key="JSON.stringify(selectedRanges)" :label="translate('sheet.fill-options', '自动填充选项')"
      :labels="{ COPY: translate('sheet.fill.copy', '复制单元格'), SERIES: translate('sheet.fill.series', '填充序列'),
        ONLY_FORMAT: translate('sheet.fill.only-format', '仅填充格式'), NO_FORMAT: translate('sheet.fill.no-format', '不带格式填充') }"
      @refill="emit('refill-selection', $event)" @restore-focus="root?.focus({ preventScroll: true })" />
    <Dropdown v-if="pasteOptionsPosition && pasteOptions" :open="pasteOptionsOpen" :trigger="['click']" :menu="pasteOptionsMenu" :popup-render="pasteOptionsPopup" :classes="{ root: `sheet-paste-options-popup ${pasteOptionsPopupClass}` }" :auto-focus="false" :auto-adjust-overflow="true" :destroy-on-hidden="true" placement="bottomLeft" @open-change="onPasteOptionsOpenChange">
      <Tooltip :title="translate('sheet.paste-options', '粘贴选项')" :open="pasteOptionsOpen ? false : undefined">
        <Button class="sheet-paste-options" :data-sheet-paste-options-owner="pasteOptionsOwner" :style="pasteOptionsStyle" size="small" :aria-label="translate('sheet.paste-options', '粘贴选项')" :aria-expanded="pasteOptionsOpen" aria-haspopup="menu" @pointerdown.stop @keydown.stop="onPasteOptionsKeydown" @dblclick.stop @contextmenu.stop>
          <ClipboardPaste :size="14" aria-hidden="true" /><ChevronDown :size="10" aria-hidden="true" />
        </Button>
      </Tooltip>
    </Dropdown>
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
    <AutoComplete v-if="editingKey" :key="editSessionId ?? editingKey" ref="input" class="canvas-cell-input" :style="editStyle" :styles="editInputStyles" variant="borderless" :value="editValue" :options="(validationOptions ?? []).map(value => ({ value, label: value }))" aria-label="单元格编辑" @change="onEditValueChange" @select="onAutoCompleteSelect" @blur="onEditorBlur" @keydown.stop="onInputKeydown">
      <Input.TextArea v-if="useMultilineEditor" :rows="1" :auto-size="false" :change-on-composing="true" :styles="editTextAreaStyles" :wrap="sheetCellWrapStrategy(editedCell) === 'wrap' ? 'soft' : 'off'" variant="borderless" spellcheck="false" aria-label="单元格编辑" @change="onTextAreaChange" />
    </AutoComplete>
    <span v-if="editingKey && editLayout && !formatPainter" class="canvas-cell-fill-handle" :style="fillHandleStyle" aria-hidden="true" />
    <div class="a11y-cell" role="row"><span role="gridcell" aria-selected="true">{{ selectedKey }} {{ selectedCellText }}</span></div>
    <Alert v-if="surfaceError" class="canvas-surface-error" type="error" :title="`画布初始化失败：${surfaceError}`" show-icon />
    <SheetContextMenu
      v-model:open="contextMenuOpen"
      :hit="contextMenuHit"
      :sheet-id="sheet.id"
      :x="contextMenuPosition.x"
      :y="contextMenuPosition.y"
      :can-delete-row="sheet.rowCount > 1"
      :can-delete-column="sheet.columnCount > 1"
      :frozen="Boolean(sheet.freeze?.rows || sheet.freeze?.columns)"
      :has-hidden-rows="hasContextHiddenRows"
      :has-hidden-columns="hasContextHiddenColumns"
      :can-convert-text="canConvertSelectedText"
      :has-hyperlink="hasContextHyperlink"
      :has-cell-images="Boolean(contextCellImagePositions.length)" :can-save-cell-images="canSaveContextImages" :context-image="contextCellImage"
      :can-format="canFormat"
      :clipboard-capabilities="clipboardCapabilities"
      :can-shift-horizontally="canShiftHorizontally"
      :can-shift-vertically="canShiftVertically"
      :selected-ranges="selectedRanges"
      :row-count="sheet.rowCount"
      :column-count="sheet.columnCount"
      :can-structure-rows="canStructureRows"
      :can-structure-columns="canStructureColumns"
      :can-visibility-rows="canVisibilityRows"
      :can-visibility-columns="canVisibilityColumns"
      @command="(commandId, params) => emit('context-command', commandId, params)"
    />
  </div>
</template>

<style scoped>
.spreadsheet-canvas { position: relative; min-width: 0; min-height: 0; flex: 1; overflow: hidden; outline: none; color: var(--sheet-canvas-text); background: var(--sheet-canvas-background); touch-action: none; user-select: none; }
.spreadsheet-canvas:focus-visible { box-shadow: inset 0 0 0 1px var(--sheet-canvas-primary); }
.spreadsheet-canvas.format-painter-active { cursor: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='28' height='28' viewBox='0 0 28 28'%3E%3Cpath d='M2 1v18l4-4 3 7 3-2-3-6h6z' fill='white' stroke='black'/%3E%3Cpath d='m17 13 8-10 2 2-8 10z' fill='white' stroke='black'/%3E%3Cpath d='M17 14c-4 0-2 5-6 6 5 2 10-1 8-5z' fill='black' stroke='white'/%3E%3C/svg%3E") 2 1, crosshair !important; }
.canvas-scroll { position: absolute; inset: 0; overflow: auto; overscroll-behavior: contain; }
.canvas-surface { position: absolute; inset: 0; pointer-events: none; overflow: hidden; }
.sheet-paste-options { position: absolute; z-index: 4; display: inline-flex; align-items: center; justify-content: center; gap: 2px; width: 44px; height: 24px; padding: 0 4px; color: var(--sheet-canvas-text); background: var(--sheet-canvas-background); border-color: var(--sheet-canvas-border); }
.sheet-paste-options:focus-visible { outline: 2px solid var(--sheet-canvas-primary); outline-offset: 1px; }
.outline-control { position: absolute; z-index: 3; box-sizing: border-box; width: 12px; height: 12px; min-width: 12px; min-height: 12px; margin: 0; padding: 0; border: 1px solid var(--sheet-canvas-border); border-radius: 1px; color: var(--sheet-canvas-text); background: var(--sheet-canvas-background); box-shadow: none; line-height: 10px; }
.outline-control :deep(.ant-btn-icon) { display: inline-flex; align-items: center; justify-content: center; }
.outline-level { font-size: 8px; }
.canvas-cell-input { position: absolute; z-index: 3; display: block; box-sizing: border-box; margin: 0; padding: 0; border-radius: 0; background: transparent; user-select: text; touch-action: auto; }
.canvas-cell-input :deep(.ant-select-input) { height: 100%; }
.canvas-cell-input :deep(.ant-select-content) { height: 100%; min-height: 0; }
.rich-text-editor-active :deep(.canvas-cell-input .ant-select-content),
.rich-text-editor-active :deep(.canvas-cell-input .ant-select-selector) { background: transparent; }
.rich-text-editor-active :deep(.canvas-cell-input input), .rich-text-editor-active :deep(.canvas-cell-input textarea) { -webkit-text-fill-color: transparent; }
.rich-text-editor-active :deep(.canvas-cell-input input::selection), .rich-text-editor-active :deep(.canvas-cell-input textarea::selection) { color: transparent; background: transparent; }
.canvas-cell-fill-handle { position: absolute; z-index: 4; box-sizing: border-box; border: 1px solid var(--sheet-canvas-background); background: var(--sheet-canvas-primary); pointer-events: none; }
.fill-handle-hovered :deep(.canvas-cell-input input), .fill-handle-hovered :deep(.canvas-cell-input textarea) { cursor: crosshair; }
.a11y-cell { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; pointer-events: none; }
.canvas-surface-error { position: absolute; z-index: 4; top: 12px; inset-inline-start: 12px; max-width: calc(100% - 24px); }
.sheet-image-float-menu { position: absolute; z-index: 5; display: flex; gap: 4px; padding: 3px; border: 1px solid; border-radius: 4px; }
.sheet-drawing-float-menu { position: absolute; z-index: 6; display: flex; gap: 4px; padding: 3px; border: 1px solid; border-radius: 4px; }
.sheet-drawing-context-anchor { position: absolute; width: 1px; height: 1px; pointer-events: none; }
.sheet-drawing-context-anchor > span { display: block; width: 1px; height: 1px; }
.sheet-drawing-error { position: absolute; z-index: 7; top: 32px; left: 48px; max-width: calc(100% - 70px); }
.sheet-image-error { position: absolute; z-index: 6; top: 32px; left: 48px; max-width: calc(100% - 70px); }
</style>
