<script setup lang="ts">
import { computed, h, inject, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, ref, useId, watch, type VNode } from 'vue'
import { Alert, AutoComplete, Button, Dropdown, Input, Tooltip, theme } from 'antdv-next'
import type { RenderBackend, RenderSurface } from '../../rendering/contracts'
import type { DrawCommand } from '../../rendering/contracts'
import type { CellRange, TimelineLevel } from '../../domain/model'
import type { ITimelineRange } from '../../domain/timeline-filters.ts'
import { createRenderSurface } from '../../rendering/createRenderSurface'
import type { GridHit, GridRange, GridViewport, SpreadsheetRenderSheet } from '../../rendering/spreadsheet/layout'
import { cellAddress, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, SpreadsheetGeometry } from '../../rendering/spreadsheet/layout'
import { createSpreadsheetScene, displayCell, sheetNumberFormatPreviewPatternAt } from '../../rendering/spreadsheet/scene'
import type { SheetNumberFormatPreview } from '../../rendering/spreadsheet/scene'
import { createCellEditorLayout, hitCellEditorFillHandle } from '../../rendering/spreadsheet/editor-layout'
import { createSheetRichTextEditorScene, hitSheetRichTextEditor, hitSheetRichTextEditorPosition, moveSheetRichTextEditorCaret,
  sheetRichTextEditorCaret, sheetRichTextEditorWord, type SheetRichTextEditorCaretAffinity } from '../../rendering/spreadsheet/rich-text-editor.ts'
import { clampSheetRichTextEditorScroll, stepSheetRichTextEditorDragScroll } from '../../rendering/spreadsheet/rich-text-editor-scroll.ts'
import type { ExcelCellView, ExcelSheetView } from '../../services/excelViewAdapter.ts'
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
import { ChevronDown, ClipboardPaste, Filter, ListFilter, Minus, Plus } from 'lucide-vue-next'
import { sheetFilterHeaderControls } from '../../features/sheet-sort-filter/sheet-filter-header.ts'
import { sheetFilterPermissionRanges } from '../../features/sheet-sort-filter/sheet-filter-column.ts'
import { hitTestSlicerObjects, slicerObjectContentPoint, slicerObjectDrawCommands, slicerObjectLayouts, slicerObjectLocalPoint,
  timelineSerialAtOffset, type SlicerCanvasControl, type SlicerCanvasHit, type SlicerObjectLayout } from '../../features/slicer/slicer-renderer'
import { captureSlicerCanvasGesture, canResumeSlicerCanvasGesture, slicerCanvasGestureDelta, slicerCanvasGestureGeometry,
  slicerCanvasScrollOffset, timelineCanvasSelectionRange, timelineCanvasHandleRange, type SlicerCanvasGesture } from '../../features/slicer/slicer-canvas-gesture.ts'
import { slicerControlPaneLayouts, slicerControlTabIndex, type SlicerControlPaneLayout } from '../../features/slicer/slicer-control-overlay.ts'
import SlicerControlOverlay from './SlicerControlOverlay.vue'
import { resolveSlicerRenderProfile, type SlicerRenderProfile } from '../../features/slicer/slicer-style-rendering.ts'
import SheetContextMenu from './SheetContextMenu.vue'
import SheetFillOptions from './SheetFillOptions.vue'
import { sheetContextMenuFeatureSearchContext, sheetContextMenuSelectionHit, sheetContextStructureTarget, type SheetContextMenuContext } from './sheet-context-menu.ts'
import type { RibbonGroup } from '../../platform/ribbon/ribbon-schema.ts'
import type { SheetOutlineCommandInput } from '../../features/sheet-outline/sheet-outline-command.ts'
import { sparklineContextCapability } from '../../features/sparkline/sparkline-context-menu.ts'
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
import { drawingPaintOrder, DRAWING_ALIGN_MENU, DRAWING_ARRANGE_MENU, DRAWING_GROUP_MENU, SHEET_DRAWING_GROUP_COMMAND_IDS, type SheetDrawingScene } from '../../features/drawing-groups/index.ts'
import { sheetDrawingGeometryRange } from '../../features/sheet-drawings/sheet-drawing-viewport.ts'
import { useSheetDrawingCanvas } from './use-sheet-drawing-canvas.ts'
import type { SheetDrawingCommandInput } from '../../api/sheet-drawing-commands.ts'
import SheetValidationCellControl from './SheetValidationCellControl.vue'
import type { SheetValidationCellEntry, SheetValidationCellPopup } from './use-sheet-validation-cell.ts'
import type { SheetValidationDropdownUiOptions } from '../../features/sheet-validation/validation-dropdown-ui.ts'
import { sheetValidationCellControlRect, sheetValidationCellVisibleRect } from '../../rendering/spreadsheet/validation-cell-layout.ts'
import SheetCheckboxValidationOverlay from './SheetCheckboxValidationOverlay.vue'
import type { SheetCheckboxCellEntry } from './use-sheet-validation-checkbox.ts'
import { sheetCheckboxCellPositions } from '../../rendering/spreadsheet/validation-checkbox-layout.ts'
import SheetValidationFeedbackPopup from './SheetValidationFeedbackPopup.vue'
import type { SheetValidationFeedbackReader } from '../../features/sheet-validation/validation-cell-feedback.ts'
import type { SheetValidationTagOptionsReader } from '../../features/sheet-validation/validation-cell-options.ts'

const MAX_SHEET_ROWS = 1_048_576
const MAX_SHEET_COLUMNS = 16_384

const props = defineProps<{ filterHeaderBlocked?: boolean; activeRangeIndex?: number; workbookId?: string; workbookHistoryGeneration?: number; sheet: SpreadsheetRenderSheet; sparklineSheets?: readonly ExcelSheetView[]; sparklineMetadata?: Record<string, unknown>; drawingScene?: SheetDrawingScene; selectedDrawingRefs?: readonly SheetDrawingRef[]; selectedImageId?: string; selectedImageIds?: readonly string[]; imageInteractionBlocked?: boolean; selectedKey: string; selectedRange: string[]; selectedRanges?: GridRange[]; editingKey: string | null; editValue: string; editSessionId?: string; richTextEditorCell?: ExcelCellView; editSelection?: { start: number; end: number; host?: 'cell' | 'formula' }; inputContext?: () => SheetCellInputBridgeContext; keepEditFocus?: (target: Element | null) => boolean; prepareFill?: () => SheetFillCommitContext | undefined; zoom: number; currentUserId?: string; workbookRevision?: number; canFormat?: boolean; searchText?: string; numberFormatPreview?: SheetNumberFormatPreview; conditionalFormatPreview?: readonly CellRange[]; validationOptions?: readonly string[]; validationCellEntry?: SheetValidationCellEntry; validationCellPopup?: SheetValidationCellPopup; validationDropdownUi?: SheetValidationDropdownUiOptions; validationCurrentYear?: number; validationCheckboxAt?: (row: number, column: number) => SheetCheckboxCellEntry | undefined; validationCheckboxError?: string; validationFeedback?: SheetValidationFeedbackReader; validationTagOptions?: SheetValidationTagOptionsReader; validationInputOwner?: string; overlayCommands?: readonly DrawCommand[]; slicers?: readonly SlicerCanvasControl[]; selectedSlicerId?: string; crosshairHighlight?: { enabled: boolean; color: string }; formatPainter?: SheetFormatPainterCanvasState | null; copyRange?: CellRange; pasteOptions?: SheetClipboardPasteOptionsView; fillOptions?: SheetFillOptionsView; clipboardCapabilities?: SheetClipboardCapabilities; pasteOptionsCapabilities?: SheetClipboardCapabilities }>()
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
  (event: 'slicer-item', payload: { slicerId: string; value: string; index: number; anchorIndex?: number; ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; multiSelect: boolean; expectedRevision?: number }): void
  (event: 'slicer-clear', slicerId: string, expectedRevision?: number): void
  (event: 'slicer-range', payload: { slicerId: string; range: ITimelineRange; expectedRevision?: number }): void
  (event: 'slicer-window', payload: { slicerId: string; start: number; count: number; expectedRevision?: number }): void
  (event: 'slicer-retry', slicerId: string): void
  (event: 'slicer-scroll-to', payload: { slicerId: string; position: number; expectedRevision?: number }): void
  (event: 'slicer-level', payload: { slicerId: string; level: TimelineLevel; expectedRevision?: number }): void
  (event: 'slicer-settings', slicerId: string): void
  (event: 'slicer-connections', slicerId: string): void
  (event: 'slicer-move', payload: { slicerId: string; left: number; top: number }): void
  (event: 'slicer-resize', payload: { slicerId: string; width: number; height: number }): void
  (event: 'slicer-geometry', payload: { slicerId: string; geometry: SheetDrawingAffineGeometry; expectedRevision?: number }): void
  (event: 'select-image', imageId?: string, additive?: boolean, preserve?: boolean): void
  (event: 'edit-image', imageId: string, row?: number, column?: number): void
  (event: 'save-image', imageId: string): void
  (event: 'image-command', commandId: string, input: SheetImageCommandInput): void
  (event: 'select-drawing', ref?: SheetDrawingRef, additive?: boolean, preserve?: boolean): void
  (event: 'drawing-command', commandId: string, input: SheetDrawingCommandInput): void
  (event: 'chart-event', eventValue: MouseEvent | KeyboardEvent, chartId: string, view: { scrollLeft: number; scrollTop: number; headerWidth?: number; headerHeight?: number }): void
  (event: 'clear-checkbox-error'): void
  (event: 'open-filter-column', column: number): void
}>()
const root = ref<HTMLElement | null>(null), canvasHost = ref<HTMLElement | null>(null), scroller = ref<HTMLElement | null>(null), input = ref<{ focus: () => void } | null>(null)
// The synchronous callback also owns "no session" while Vue is still flushing the previous props.
const inputSessionId = () => props.inputContext ? props.inputContext().sessionId : props.editSessionId
const nativeInput = createSheetCellNativeInputBridge({ host: 'cell', sessionId: inputSessionId,
  value: () => props.inputContext?.().value ?? props.editValue, contains: element => Boolean(root.value?.contains(element) && element.closest('.canvas-cell-input')),
  emit: event => emit('input-event', event) })
const geometry = computed(() => new SpreadsheetGeometry(props.sheet))
const viewport = ref<GridViewport>({ width: 1, height: 1, scrollLeft: 0, scrollTop: 0, zoom: props.zoom / 100,
  headerWidth: geometry.value.headerWidth, headerHeight: geometry.value.headerHeight })
const fillPreview = ref<GridRange>()
const fillHandleHovered = ref(false)
const filterHeaderPermissionRanges = computed(() => sheetFilterPermissionRanges({ mergedRanges: geometry.value.merges.map(range => ({
  startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
})) }, props.sheet.autoFilter?.range))
const filterHeaderControls = computed(() => sheetFilterHeaderControls(geometry.value, viewport.value, props.sheet.autoFilter,
  () => !props.filterHeaderBlocked && !composing.value && canEditWorksheetProtection(props.sheet, props.currentUserId ?? '')
    && !!props.sheet.autoFilter && canEditProtectedRanges(props.sheet, filterHeaderPermissionRanges.value, props.currentUserId ?? '')))
function openFilterHeader(column: number): void {
  const control = filterHeaderControls.value.find(control => control.column === column && !control.disabled)
  if (control) emit('open-filter-column', column)
}
const headerDrag = ref<(SheetAxisHeaderMoveSource & { pointerStart: number; moved: boolean; copy: boolean; destination?: number; allowed: boolean; revision?: number })>()
const region = useOfficeRegion()
const slicerScrollPreviews = ref(new Map<string, number>())
const slicerLayouts = computed(() => slicerObjectLayouts(geometry.value, props.slicers ?? [], props.sheet.date1904 ?? false,
  region.value, drawingCanvas.previewGeometry.value, slicerScrollPreviews.value))
const slicerControlPanes = computed(() => slicerControlPaneLayouts(geometry.value, viewport.value, slicerLayouts.value,
  props.drawingScene, drawingCanvas.previewGeometry.value, drawingVisible))
const slicerTheme = computed(() => ({ background: token.value.colorBgContainer, text: token.value.colorText,
  secondaryFill: token.value.colorFillSecondary, mutedText: token.value.colorTextTertiary }))
const slicerStyleProfiles = computed(() => new Map<string, SlicerRenderProfile>(slicerLayouts.value.map(layout => [`${layout.model.type}:${layout.id}`,
  resolveSlicerRenderProfile(layout.styleOwner ?? {}, layout.model.type, layout.model.style, slicerTheme.value)])))
function slicerProfile(layout: SlicerObjectLayout): SlicerRenderProfile {
  return slicerStyleProfiles.value.get(`${layout.model.type}:${layout.id}`)
    ?? resolveSlicerRenderProfile(layout.styleOwner ?? {}, layout.model.type, layout.model.style, slicerTheme.value)
}
const timelineMenuId = ref('')
const slicerMenuPaneKey = ref('')
const slicerContentId = ref('')
const focusedSlicerItems = ref(new Map<string, number>())
const focusedTimelinePeriods = ref(new Map<string, number>())
const scene = computed(() => {
  const selectionCommands: DrawCommand[] = []
  const selected = slicerLayouts.value.find(({ id }) => id === props.selectedSlicerId)
  if (selected && !drawingCanvas.sharedControls.value) selectionCommands.push({ kind: 'rect', ...selected.rect,
    matrix: selected.matrix, localClip: selected.rect, drawingRef: { kind: selected.model.type, id: selected.id },
    fill: 'transparent', stroke: token.value.colorPrimary, strokeWidth: 2 })
  const overlays = (props.overlayCommands ?? []).flatMap(command => {
    if (!command.drawingRef) return [command]
    const object = props.drawingScene?.objects.find(object => object.ref.kind === command.drawingRef?.kind && object.ref.id === command.drawingRef.id)
    if (object && !drawingVisible(object.ref, object.geometry)) return []
    const preview = drawingCanvas.previewGeometry.value.get(`${command.drawingRef.kind}:${command.drawingRef.id}`)
    return command.kind === 'image' && command.drawingRef.kind === 'chart' && preview ? [{ ...command, x: 0, y: 0,
      width: preview.width, height: preview.height, matrix: { ...preview.matrix, e: preview.matrix.e + GRID_HEADER_WIDTH, f: preview.matrix.f + GRID_HEADER_HEIGHT } }] : [command]
  })
  const result = createSpreadsheetScene(geometry.value, viewport.value, {
    region: region.value, activeCell: props.selectedKey, editing: Boolean(props.editingKey), selectedCells: props.selectedRange,
    selectedRanges: props.selectedRanges, fillPreviewRange: fillPreview.value,
    formatPainterSource: props.formatPainter?.sourceSheetId === props.sheet.id ? props.formatPainter.sourceRange : undefined,
    conditionalFormatPreviewRanges: !props.editingKey && canViewWorksheetProtection(props.sheet, props.currentUserId ?? 'local-user')
      && canViewProtectedRanges(props.sheet, props.conditionalFormatPreview ?? [], props.currentUserId ?? 'local-user')
      ? props.conditionalFormatPreview?.map(sheetClipboardGridRange) : undefined,
    conditionalFormatPreviewColor: token.value.colorPrimary,
    validationTagTheme: { fill: token.value.colorFillSecondary, text: token.value.colorText },
    validationTagOptions: (row, column) => validationTagOptionCells.value.get(cellAddress(row, column)),
    checkboxRenderable: (row, column) => checkboxControls.value.entries.has(cellAddress(row, column)),
    validationInvalid: (row, column) => !!validationFeedbackCells.value.get(cellAddress(row, column))?.issues.length,
    validationInvalidColor: token.value.colorError,
    canViewCell: (row, column) => {
      const merge = geometry.value.mergeAt(row, column)
      const range = merge ? { startRow: merge.startRow, endRow: merge.endRow, startCol: merge.startColumn, endCol: merge.endColumn }
        : { startRow: row, endRow: row, startCol: column, endCol: column }
      return canViewWorksheetProtection(props.sheet, props.currentUserId ?? 'local-user')
        && canViewProtectedRanges(props.sheet, [range], props.currentUserId ?? 'local-user')
    },
    formatPainterColor: token.value.colorPrimary, suppressFillHandle: Boolean(props.formatPainter || props.selectedImageId || props.selectedDrawingRefs?.length),
    searchText: props.searchText, numberFormatPreview: props.numberFormatPreview, drawingOrder: drawingOrder.value,
    sparklineContext: { sheets: props.sparklineSheets, metadata: props.sparklineMetadata, canReadSource: (sheet, range) =>
      canViewWorksheetProtection(sheet, props.currentUserId ?? 'local-user') && canViewProtectedRanges(sheet, [range], props.currentUserId ?? 'local-user') },
    overlayCommands: [...overlays, ...slicerObjectDrawCommands(slicerLayouts.value.filter(layout =>
      drawingVisible({ kind: layout.model.type, id: layout.id }, layout.geometry)), { controlUI: 'antd', theme: slicerTheme.value,
      styleProfiles: slicerStyleProfiles.value, selectedSlicerId: props.selectedSlicerId, multiSelect: multiSelect.value,
      focusedItems: focusedSlicerItems.value, focusedPeriods: focusedTimelinePeriods.value }), ...selectionCommands],
    crosshairHighlight: props.crosshairHighlight, ...imageCanvas.sceneOptions.value,
  })
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
function outlineOwner(): SheetOutlineCommandInput {
  return { unitId: props.workbookId ?? 'standalone-workbook', subUnitId: props.sheet.id, host: 'workbench',
    expectedRevision: props.workbookRevision, expectedHistoryGeneration: props.workbookHistoryGeneration,
    expectedUserId: props.currentUserId ?? 'local-user' }
}
function outlineOwnerCurrent(owner: SheetOutlineCommandInput): boolean {
  return owner.unitId === (props.workbookId ?? 'standalone-workbook') && owner.subUnitId === props.sheet.id
    && owner.expectedRevision === props.workbookRevision && owner.expectedHistoryGeneration === props.workbookHistoryGeneration
    && owner.expectedUserId === (props.currentUserId ?? 'local-user')
}
function canChangeOutline(axis: 'row' | 'column', groupId?: string): boolean {
  const userId = props.currentUserId ?? 'local-user'
  if (!active || disposed || props.editingKey || composing.value || props.imageInteractionBlocked || props.formatPainter
    || !canEditWorksheetProtection(props.sheet, userId)) return false
  const groups = props.sheet.outline?.groups.filter(group => group.axis === axis && (!groupId || group.id === groupId)) ?? []
  return groups.length > 0 && canEditProtectedRanges(props.sheet, groups.map(group => axis === 'row'
    ? { startRow: group.start, endRow: group.end, startCol: 0, endCol: props.sheet.columnCount - 1 }
    : { startRow: 0, endRow: props.sheet.rowCount - 1, startCol: group.start, endCol: group.end }), userId)
}
const outlineToggleRects = computed(() => geometry.value.outlineToggleRects(viewport.value).map(control => ({
  ...control, owner: outlineOwner(), disabled: !canChangeOutline(control.axis, control.groupId),
  label: `${control.axis === 'row' ? '行' : '列'}分组${control.collapsed ? '展开' : '折叠'}`,
})))
const outlineLevelButtons = computed(() => geometry.value.outlineLevelRects().map(control => ({
  ...control, owner: outlineOwner(), disabled: !canChangeOutline(control.axis),
  label: `${control.axis === 'row' ? '行' : '列'}大纲级别 ${control.level}`,
})))
function changeOutlineControl(control: (typeof outlineToggleRects.value)[number]): void {
  if (control.disabled || !outlineOwnerCurrent(control.owner) || !canChangeOutline(control.axis, control.groupId)) return
  emit('context-command', 'sheet.command.set-dimension-outline-collapsed', {
    ...control.owner, outlineId: control.groupId, collapsed: !control.collapsed,
  })
}
function changeOutlineLevel(control: (typeof outlineLevelButtons.value)[number]): void {
  if (control.disabled || !outlineOwnerCurrent(control.owner) || !canChangeOutline(control.axis)) return
  emit('context-command', 'sheet.command.set-outline-level', { ...control.owner, axis: control.axis, level: control.level })
}
const fullSize = computed(() => ({ width: `${geometry.value.headerWidth + geometry.value.columns.offset(props.sheet.columnCount) * props.zoom / 100}px`,
  height: `${geometry.value.headerHeight + geometry.value.rows.offset(props.sheet.rowCount) * props.zoom / 100}px` }))
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
const checkboxControls = computed(() => {
  const entries = new Map<string, SheetCheckboxCellEntry>()
  const positions = sheetCheckboxCellPositions(geometry.value, viewport.value, (row, column) => {
    const entry = props.validationCheckboxAt?.(row, column)
    if (entry) entries.set(entry.key, entry)
    return !!entry
  })
  return { entries, controls: positions.map(position => ({ position, entry: entries.get(position.key)! })) }
})
const validationCellPositions = computed(() => {
  const positions = new Map<string, { row: number; column: number }>()
  for (const pane of geometry.value.panes(viewport.value)) for (const row of pane.rows) for (const column of pane.columns) {
    const merge = geometry.value.mergeAt(row, column), position = { row: merge?.startRow ?? row, column: merge?.startColumn ?? column }
    positions.set(cellAddress(position.row, position.column), position)
  }
  const selected = parseAddress(props.selectedKey)
  if (selected) positions.set(props.selectedKey, selected)
  return [...positions.values()]
})
const validationFeedbackCells = computed(() => props.validationFeedback?.(validationCellPositions.value) ?? new Map())
const validationTagOptionCells = computed(() => props.validationTagOptions?.(validationCellPositions.value) ?? new Map())
const selectedValidationFeedback = computed(() => {
  const position = parseAddress(props.selectedKey), merge = position && geometry.value.mergeAt(position.row, position.column)
  return position ? validationFeedbackCells.value.get(cellAddress(merge?.startRow ?? position.row, merge?.startColumn ?? position.column)) : undefined
})
const hoveredValidationKey = ref<string>(), dismissedValidationPrompt = ref<string>()
let validationHoverCandidate: string | undefined, validationHoverTimer: ReturnType<typeof setTimeout> | undefined
const hoveredValidationFeedback = computed(() => hoveredValidationKey.value ? validationFeedbackCells.value.get(hoveredValidationKey.value) : undefined)
const hoveredValidationPosition = computed(() => hoveredValidationFeedback.value?.issues.length
  ? sheetValidationCellVisibleRect(geometry.value, viewport.value, hoveredValidationFeedback.value.key) : undefined)
const validationPromptOwner = computed(() => selectedValidationFeedback.value?.prompt ? JSON.stringify([
  props.workbookId, props.sheet.id, props.currentUserId, selectedValidationFeedback.value.key, selectedValidationFeedback.value.prompt]) : undefined)
const validationPromptPosition = computed(() => validationPromptOwner.value && validationPromptOwner.value !== dismissedValidationPrompt.value
  && !(hoveredValidationPosition.value && hoveredValidationFeedback.value?.key === selectedValidationFeedback.value?.key)
  ? sheetValidationCellVisibleRect(geometry.value, viewport.value, selectedValidationFeedback.value?.key) : undefined)
const dismissValidationPrompt = computed(() => {
  const owner = validationPromptOwner.value
  return () => { if (owner && !composing.value && validationPromptOwner.value === owner) dismissedValidationPrompt.value = owner }
})
function clearValidationHover(): void {
  if (validationHoverTimer) clearTimeout(validationHoverTimer)
  validationHoverTimer = undefined; validationHoverCandidate = undefined; hoveredValidationKey.value = undefined
}
function scheduleValidationHover(key?: string): void {
  if (key === validationHoverCandidate) return
  clearValidationHover()
  validationHoverCandidate = key
  if (key) validationHoverTimer = setTimeout(() => {
    validationHoverTimer = undefined
    if (active && !disposed && validationHoverCandidate === key) hoveredValidationKey.value = key
  }, 100)
}
watch(() => [props.workbookId, props.sheet.id, props.currentUserId, props.selectedKey], () => {
  dismissedValidationPrompt.value = undefined; clearValidationHover()
})
watch(() => [viewport.value.scrollLeft, viewport.value.scrollTop, viewport.value.zoom], clearValidationHover)
const validationCellControlPosition = computed(() => {
  const rect = sheetValidationCellControlRect(geometry.value, viewport.value, props.validationCellEntry?.key)
  return rect && !props.formatPainter && !props.selectedDrawingRefs?.length && !composing.value
    ? { left: `${rect.x}px`, top: `${rect.y}px`, width: `${rect.width}px`, height: `${rect.height}px` } : undefined
})
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
  contentEditing: () => props.selectedDrawingRefs?.length === 1 && ['slicer', 'timeline'].includes(props.selectedDrawingRefs[0]!.kind)
    && slicerContentId.value === props.selectedDrawingRefs[0]?.id,
  contentControlAt: (ref, position, pane) => {
    const layout = slicerLayouts.value.find(layout => layout.id === ref.id && layout.model.type === ref.kind)
    if (!layout) return false
    const content = drawingContentPoint(position, undefined, pane), hit = hitTestSlicerObjects([layout], content.x, content.y)
    return !!hit && !['move', 'resize', 'body', 'timelineMenu'].includes(hit.type)
  },
  keepRatio: () => props.selectedDrawingRefs?.length === 1 && !!props.slicers?.find(control => control.model.id === props.selectedDrawingRefs[0]?.id
    && control.model.type === props.selectedDrawingRefs[0]?.kind)?.model.drawing?.keepRatio,
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
  nudge: (ids, direction, step, expectedRevision) => {
    if (!props.drawingScene) return false
    emit('drawing-command', SHEET_DRAWING_GROUP_COMMAND_IDS.move, { sheetId: props.sheet.id,
      selection: ids.map(id => ({ kind: 'image', id })), expectedSelection: props.selectedDrawingRefs ? [...props.selectedDrawingRefs] : undefined,
      expectedRevision, direction, step })
    return true
  },
  crop: (id, input) => emit('image-command', id, { ...input, sheetId: props.sheet.id, expectedRevision: props.workbookRevision }),
  render: queueRender,
})
const imageSelectionIds = computed(() => props.selectedImageIds ?? (props.selectedImageId ? [props.selectedImageId] : []))
const multipleImagesSelected = computed(() => imageSelectionIds.value.length > 1)
const drawingContextOpen = ref(false), drawingContextPosition = ref({ x: 0, y: 0 })
const drawingContextPopupClass = `sheet-drawing-context-popup-${pasteOptionsOwner}`
function closeDrawingContextMenu(): void {
  const focused = document.activeElement, host = root.value, sheet = props.sheet,
    revision = props.workbookRevision, selection = JSON.stringify(props.selectedDrawingRefs)
  drawingContextOpen.value = false
  if (!(focused instanceof Element) || !focused.closest(`.${drawingContextPopupClass}`)) return
  void nextTick(() => {
    if (drawingContextOpen.value || !active || disposed || props.editingKey || props.imageInteractionBlocked
      || root.value !== host || props.sheet !== sheet || props.workbookRevision !== revision
      || JSON.stringify(props.selectedDrawingRefs) !== selection) return
    const current = document.activeElement
    if (current === document.body || current instanceof Element && current.closest(`.${drawingContextPopupClass}`)) focus()
  })
}
function onDrawingContextOpenChange(open: boolean): void {
  if (open) drawingContextOpen.value = true
  else closeDrawingContextMenu()
}
function onDrawingContextKeydown(event: KeyboardEvent): void {
  if (!drawingContextOpen.value || event.key !== 'Escape' || event.isComposing || event.keyCode === 229) return
  event.preventDefault(); event.stopPropagation(); closeDrawingContextMenu()
}
const drawingContextPopup = (node: VNode) => h('div', { onKeydownCapture: onDrawingContextKeydown }, [node])
const drawingGroupMenu = computed(() => ({ items: DRAWING_GROUP_MENU.items.map(item => ({ key: item.contentCommandId, label: item.label,
  disabled: item.action === 'group' ? !drawingCanvas.canGroup.value : !drawingCanvas.canUngroup.value })),
  onClick: ({ key }: { key: string | number }) => chooseDrawingMenu(String(key)) }))
const drawingArrangeMenu = computed(() => ({ items: DRAWING_ARRANGE_MENU.items.map(item => ({ key: item.action, label: item.label, disabled: !drawingCanvas.editable.value })),
  onClick: ({ key }: { key: string | number }) => chooseDrawingMenu(`arrange:${key}`) }))
const drawingAlignEnabled = computed(() => drawingCanvas.transformable.value && (props.selectedDrawingRefs?.length ?? 0) > 1)
const drawingAlignMenu = computed(() => ({ items: DRAWING_ALIGN_MENU.items.map(item => ({ key: `align:${item.alignType}`, label: item.label, disabled: !drawingAlignEnabled.value })),
  onClick: ({ key }: { key: string | number }) => chooseDrawingMenu(String(key)) }))
const drawingMenuSlicer = computed(() => {
  const refs = props.selectedDrawingRefs ?? [], ref = refs.length === 1 ? refs[0] : undefined
  if (!ref || (ref.kind !== 'slicer' && ref.kind !== 'timeline')) return undefined
  return props.slicers?.find(control => control.model.id === ref.id && control.model.type === ref.kind)?.model
})
const drawingContextMenu = computed(() => ({ selectable: false, items: [
  ...(drawingMenuSlicer.value ? [{ key: 'slicer:settings', label: drawingMenuSlicer.value.type === 'timeline' ? '时间线设置' : '切片器设置', disabled: !drawingCanvas.editable.value },
    ...(!drawingMenuSlicer.value.tableId ? [{ key: 'slicer:connections', label: '报表连接', disabled: !drawingCanvas.editable.value }] : [])] : []),
  { key: 'group', label: '组合', children: drawingGroupMenu.value.items },
  { key: 'arrange', label: '排列', children: DRAWING_ARRANGE_MENU.items.map(item => ({ key: `arrange:${item.action}`, label: item.label, disabled: !drawingCanvas.editable.value })) },
  { key: 'align', label: '对齐', children: drawingAlignMenu.value.items },
  { key: 'remove', label: '删除', danger: true, disabled: !drawingCanvas.editable.value },
], onClick: ({ key }: { key: string | number }) => chooseDrawingMenu(String(key)) }))
function chooseDrawingMenu(key: string): void {
  closeDrawingContextMenu()
  if (key === 'slicer:settings' && drawingMenuSlicer.value && drawingCanvas.editable.value) emit('slicer-settings', drawingMenuSlicer.value.id)
  else if (key === 'slicer:connections' && drawingMenuSlicer.value && !drawingMenuSlicer.value.tableId && drawingCanvas.editable.value) emit('slicer-connections', drawingMenuSlicer.value.id)
  else if (key === 'remove') drawingCanvas.run('sheet.command.remove-sheet-drawings')
  else if (key === SHEET_DRAWING_GROUP_COMMAND_IDS.group && drawingCanvas.canGroup.value) drawingCanvas.run(key)
  else if (key === SHEET_DRAWING_GROUP_COMMAND_IDS.ungroup && drawingCanvas.canUngroup.value) drawingCanvas.run(key)
  else if (key.startsWith('arrange:')) {
    const item = DRAWING_ARRANGE_MENU.items.find(item => `arrange:${item.action}` === key)
    if (item) drawingCanvas.run(SHEET_DRAWING_GROUP_COMMAND_IDS.arrange, { arrangeType: item.action })
  } else if (key.startsWith('align:') && drawingAlignEnabled.value) {
    const item = DRAWING_ALIGN_MENU.items.find(item => `align:${item.alignType}` === key)
    if (item) drawingCanvas.run(SHEET_DRAWING_GROUP_COMMAND_IDS.align, { alignType: item.alignType })
  }
}
watch([() => props.sheet, () => props.workbookRevision, () => props.selectedDrawingRefs, () => props.imageInteractionBlocked,
  () => props.editingKey, () => props.currentUserId, viewport], () => { drawingContextOpen.value = false }, { flush: 'sync' })
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
const contextMenuContext = computed(() => contextMenuOpen.value && contextMenuHit.value ? createSheetContextMenuContext(contextMenuHit.value) : undefined)
const contextSelection = computed(() => {
  const ranges = props.selectedRanges ?? []
  return ranges[props.activeRangeIndex ?? ranges.length - 1]
})
function createSheetContextMenuContext(hit: GridHit): SheetContextMenuContext {
  const userId = props.currentUserId ?? 'local-user'
  const selectedRanges = props.selectedRanges ?? []
  const selection = contextSelection.value
  const inside = selectedRanges.some(range => hit.kind === 'cell' ? hit.row >= range.startRow && hit.row <= range.endRow && hit.column >= range.startColumn && hit.column <= range.endColumn
    : hit.kind === 'row' ? range.startColumn === 0 && range.endColumn === props.sheet.columnCount - 1 && hit.row >= range.startRow && hit.row <= range.endRow
      : hit.kind === 'column' && range.startRow === 0 && range.endRow === props.sheet.rowCount - 1 && hit.column >= range.startColumn && hit.column <= range.endColumn)
  const ranges: CellRange[] = inside ? selectedRanges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }))
    : [hit.kind === 'cell' ? { startRow: hit.row, endRow: hit.row, startCol: hit.column, endCol: hit.column }
      : hit.kind === 'row' ? { startRow: hit.row, endRow: hit.row, startCol: 0, endCol: props.sheet.columnCount - 1 }
        : hit.kind === 'column' ? { startRow: 0, endRow: props.sheet.rowCount - 1, startCol: hit.column, endCol: hit.column }
          : { startRow: 0, endRow: props.sheet.rowCount - 1, startCol: 0, endCol: props.sheet.columnCount - 1 }]
  const merge = hit.kind === 'cell' ? geometry.value.mergeAt(hit.row, hit.column) : undefined
  const imagePosition = hit.kind === 'cell' ? { row: merge?.startRow ?? hit.row, column: merge?.startColumn ?? hit.column } : undefined
  const image = imagePosition ? props.sheet.cells[cellAddress(imagePosition.row, imagePosition.column)]?.image : undefined
  const imageRanges = Object.entries(props.sheet.cells).flatMap(([key, cell]) => {
    const point = cell.image && parseAddress(key)
    return point && ranges.some(range => point.row >= range.startRow && point.row <= range.endRow && point.column >= range.startCol && point.column <= range.endCol)
      ? [{ startRow: point.row, endRow: point.row, startCol: point.column, endCol: point.column }] : []
  })
  const structure = { sheetId: props.sheet.id, rowCount: props.sheet.rowCount, columnCount: props.sheet.columnCount, selectedRanges }
  const rowRanges = sheetContextStructureTarget(hit, structure, 'row').params.ranges
  const columnRanges = sheetContextStructureTarget(hit, structure, 'column').params.ranges
  const canEditSheet = canEditWorksheetProtection(props.sheet, userId)
  function canStructureAxis(axis: 'row' | 'column'): boolean {
    const first = (axis === 'row' ? rowRanges : columnRanges)[0]
    if (!first || !canEditSheet) return false
    return canEditProtectedRanges(props.sheet, [axis === 'row' ? { ...first, endRow: MAX_SHEET_ROWS - 1 } : { ...first, endCol: MAX_SHEET_COLUMNS - 1 }], userId)
  }
  const primary = hit.kind === 'cell' ? hit.key : props.selectedKey
  const sparkline = hit.kind === 'cell' || hit.kind === 'row' || hit.kind === 'column'
    ? sparklineContextCapability(props.sheet, ranges, primary, userId, props.sparklineSheets) : undefined
  const canConvertText = Object.entries(props.sheet.cells).some(([key, cell]) => {
    const address = parseAddress(key)
    if (!address) return false
    const selected = selectedRanges.length ? selectedRanges.some(range => address.row >= range.startRow && address.row <= range.endRow
      && address.column >= range.startColumn && address.column <= range.endColumn) : key === props.selectedKey
    return selected && Boolean(convertTextToNumber(cell))
  })
  return { ...structure, clipboardCapabilities: props.clipboardCapabilities, canFormat: props.canFormat, canConvertText,
    canDeleteRow: props.sheet.rowCount > 1, canDeleteColumn: props.sheet.columnCount > 1,
    frozen: Boolean(props.sheet.freeze?.rows || props.sheet.freeze?.columns),
    hasHiddenRows: sheetAxisVisibilityTargetHasHidden(props.sheet, 'row', rowRanges),
    hasHiddenColumns: sheetAxisVisibilityTargetHasHidden(props.sheet, 'column', columnRanges),
    hasHyperlink: hit.kind === 'cell' && Boolean(props.sheet.cells[hit.key]?.hyperlink),
    hasCellImages: imageRanges.length > 0, canSaveCellImages: canViewProtectedRanges(props.sheet, imageRanges, userId),
    contextImage: imagePosition && image && canViewProtectedRanges(props.sheet, [{ startRow: imagePosition.row, endRow: imagePosition.row,
      startCol: imagePosition.column, endCol: imagePosition.column }], userId) ? { id: image.id, ...imagePosition } : undefined,
    canShiftHorizontally: Boolean(selection && canEditProtectedRanges(props.sheet, [{ startRow: selection.startRow, endRow: selection.endRow,
      startCol: selection.startColumn, endCol: MAX_SHEET_COLUMNS - 1 }], userId)),
    canShiftVertically: Boolean(selection && canEditProtectedRanges(props.sheet, [{ startRow: selection.startRow, endRow: MAX_SHEET_ROWS - 1,
      startCol: selection.startColumn, endCol: selection.endColumn }], userId)),
    canStructureRows: canStructureAxis('row'), canStructureColumns: canStructureAxis('column'),
    canVisibilityRows: rowRanges.length > 0 && canEditSheet && canEditProtectedRanges(props.sheet, rowRanges, userId),
    canVisibilityColumns: columnRanges.length > 0 && canEditSheet && canEditProtectedRanges(props.sheet, columnRanges, userId),
    sparkline: sparkline?.visible ? { ranges: ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow,
      startColumn: range.startCol, endColumn: range.endCol, sheetId: props.sheet.id })), primary, canEdit: sparkline.canEdit,
      canDelete: sparkline.canDelete, expectedRevision: props.workbookRevision } : undefined,
    outline: hit.kind === 'row' || hit.kind === 'column' ? { context: { unitId: props.workbookId ?? 'standalone-workbook',
      sheetId: props.sheet.id, sheet: props.sheet, userId, selections: (hit.kind === 'row' ? rowRanges : columnRanges).map(range => ({
        ...range, rangeType: hit.kind === 'row' ? 1 as const : 2 as const })) }, owner: outlineOwner() } : undefined,
  }
}
function featureSearchContext(): RibbonGroup[] {
  if (!active || disposed) return []
  const position = parseAddress(props.selectedKey)
  if (!position) return []
  const hit = sheetContextMenuSelectionHit({ kind: 'cell', key: props.selectedKey, ...position }, contextSelection.value, props.sheet)
  return sheetContextMenuFeatureSearchContext(hit, createSheetContextMenuContext(hit), translate)
}
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
let slicerPress: { hit: SlicerCanvasHit; start: SlicerCanvasHit; gesture: SlicerCanvasGesture;
  pane?: { frozenX: boolean; frozenY: boolean } } | undefined
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
    const dx = axis === 'row' ? 0 : lastPointer.x < geometry.value.headerWidth ? -16 : lastPointer.x > viewport.value.width - 16 ? 16 : 0
    const dy = axis === 'column' ? 0 : lastPointer.y < geometry.value.headerHeight ? -16 : lastPointer.y > viewport.value.height - 16 ? 16 : 0
    if (dx || dy) {
      const oldLeft = scroller.value.scrollLeft, oldTop = scroller.value.scrollTop
      scroller.value.scrollBy(dx, dy)
      if (oldLeft !== scroller.value.scrollLeft || oldTop !== scroller.value.scrollTop) {
        onScroll()
        if (axis) updateHeaderGesture(lastPointer)
        else if (fillDrag) updateFillDragTarget(lastPointer)
        else {
          const hit = geometry.value.hitTest(Math.max(geometry.value.headerWidth + 1, Math.min(viewport.value.width - 1, lastPointer.x)), Math.max(geometry.value.headerHeight + 1, Math.min(viewport.value.height - 1, lastPointer.y)), viewport.value)
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
function slicerGestureContext(controlId = props.selectedSlicerId) {
  return { unitId: props.slicers?.find(control => control.model.id === controlId)?.model.unitId,
    sheetId: props.sheet.id, revision: props.workbookRevision, userId: props.currentUserId ?? 'local-user', viewport: viewport.value }
}
function currentSlicerGesture(press: NonNullable<typeof slicerPress>): boolean {
  return active && !disposed && !props.editingKey && !composing.value && !props.imageInteractionBlocked && !props.formatPainter
    && props.selectedSlicerId === press.start.id
    && (!props.drawingScene || props.selectedDrawingRefs?.length === 1 && props.selectedDrawingRefs[0]?.id === press.start.id
      && props.selectedDrawingRefs[0]?.kind === press.gesture.layout.model.type)
    && canResumeSlicerCanvasGesture(press.gesture, slicerLayouts.value.find(layout => layout.id === press.start.id), slicerGestureContext())
}
let slicerLevelPress: SlicerCanvasGesture | undefined
let slicerLevelPaneKey = ''
let slicerMenuGesture: SlicerCanvasGesture | undefined
let slicerAntScrollPress: { gesture: SlicerCanvasGesture; pane?: SlicerControlPaneLayout; pointerId?: number; part: 'track' | 'thumb' } | undefined
function slicerControlEditable(layout: SlicerObjectLayout): boolean {
  const userId = props.currentUserId ?? 'local-user'
  return active && !disposed && !props.editingKey && !composing.value && !props.imageInteractionBlocked && !props.formatPainter
    && drawingVisible({ kind: layout.model.type, id: layout.id }, layout.geometry)
    && canEditWorksheetProtection(props.sheet, userId) && canEditProtectedRanges(props.sheet, [sheetDrawingGeometryRange(geometry.value, layout.geometry)], userId)
}
function activateSlicerControl(id: string): SlicerObjectLayout | undefined {
  const layout = slicerLayouts.value.find(layout => layout.id === id)
  if (!layout || !slicerControlEditable(layout)) return undefined
  if (props.selectedSlicerId !== id) emit('select-slicer', id)
  slicerContentId.value = id
  return layout
}
function currentSlicerControl(gesture: SlicerCanvasGesture): boolean {
  const layout = slicerLayouts.value.find(layout => layout.id === gesture.layout.id)
  return !!layout && slicerControlEditable(layout) && props.selectedSlicerId === layout.id
    && (!props.drawingScene || props.selectedDrawingRefs?.length === 1 && props.selectedDrawingRefs[0]?.kind === layout.model.type
      && props.selectedDrawingRefs[0]?.id === layout.id)
    && canResumeSlicerCanvasGesture(gesture, layout, slicerGestureContext(layout.id))
}
function restoreSlicerControlFocus(gesture: SlicerCanvasGesture): void {
  void nextTick(() => {
    const current = slicerGestureContext(gesture.layout.id)
    if (active && !disposed && !props.editingKey && !props.imageInteractionBlocked && !props.formatPainter
      && current.unitId === gesture.unitId && current.sheetId === gesture.sheetId && current.userId === gesture.userId
      && props.selectedSlicerId === gesture.layout.id && (current.revision === gesture.revision
        || gesture.revision !== undefined && current.revision === gesture.revision + 1)) root.value?.focus({ preventScroll: true })
  })
}
function beginSlicerAntControl(payload: { id: string; event: PointerEvent; pane: SlicerControlPaneLayout }): SlicerCanvasGesture | undefined {
  if (payload.event.button !== 0 || gesturePointerId !== undefined || drawingCanvas.hasGesture() || imageCanvas.hasGesture() || !root.value || !active || disposed) return undefined
  const position = point(payload.event)
  if (drawingCanvas.begin(position, payload.event)) return undefined
  const hit = drawingCanvas.pick(position)
  if (props.drawingScene && (!hit || hit.leaf.id !== payload.id || hit.leaf.kind !== payload.pane.layout.model.type)) return undefined
  const layout = activateSlicerControl(payload.id)
  if (!layout) return undefined
  return captureSlicerCanvasGesture(layout, drawingContentPoint(position, undefined, payload.pane.pane), slicerGestureContext(layout.id))
}
function beginSlicerLevel(payload: { id: string; event: PointerEvent; pane: SlicerControlPaneLayout }): void {
  if (timelineMenuId.value && timelineMenuId.value !== payload.id) closeSlicerLevel()
  slicerLevelPress = beginSlicerAntControl(payload)
  slicerLevelPaneKey = slicerLevelPress ? payload.pane.key : ''
}
function toggleSlicerLevel(payload: { id: string; event: MouseEvent }): void {
  const layout = slicerLayouts.value.find(layout => layout.id === payload.id), pointer = slicerLevelPress
  slicerLevelPress = undefined
  if (!layout || !slicerControlEditable(layout)) return
  const gesture = payload.event.detail === 0
    ? captureSlicerCanvasGesture(layout, { x: layout.bounds.x, y: layout.bounds.y }, slicerGestureContext(layout.id)) : pointer
  if (!gesture || !currentSlicerControl(gesture)) return
  if (timelineMenuId.value === layout.id) closeSlicerLevel(true)
  else {
    const pane = payload.event.detail === 0 ? slicerControlPanes.value.find(pane => pane.layout.id === layout.id
      && slicerControlTabIndex(slicerControlPanes.value, pane, layout.level) === 0) : slicerControlPanes.value.find(pane => pane.key === slicerLevelPaneKey)
    if (!pane) return
    slicerMenuGesture = gesture; slicerMenuPaneKey.value = pane.key; timelineMenuId.value = layout.id
  }
}
function closeSlicerLevel(restoreFocus = false): void {
  const gesture = slicerMenuGesture
  slicerMenuGesture = undefined; timelineMenuId.value = ''; slicerMenuPaneKey.value = ''
  if (restoreFocus && gesture) restoreSlicerControlFocus(gesture)
}
function onSlicerMenuOutsidePointer(event: PointerEvent): void {
  if (!active || disposed || !timelineMenuId.value) return
  const target = event.composedPath().find(node => node instanceof Element)
  const trigger = target instanceof Element ? target.closest('[data-slicer-control="level"]') : null
  const owned = event.composedPath().some(node => node instanceof Element && (
    node.getAttribute('data-slicer-popup') === timelineMenuId.value
    || !!trigger && node.getAttribute('data-slicer-id') === timelineMenuId.value
      && node.getAttribute('data-slicer-pane') === slicerMenuPaneKey.value))
  if (!owned) closeSlicerLevel()
}
function chooseSlicerLevel(payload: { id: string; level: TimelineLevel }): void {
  const gesture = slicerMenuGesture, current = gesture?.layout.id === payload.id && currentSlicerControl(gesture)
  closeSlicerLevel()
  if (!gesture || !current) return
  emit('slicer-level', { slicerId: payload.id, level: payload.level, expectedRevision: gesture.revision })
  restoreSlicerControlFocus(gesture)
}
function cancelSlicerAntScroll(): void {
  slicerAntScrollPress = undefined; slicerScrollPreviews.value = new Map()
}
function leaveSlicerAntScroll(id: string): void {
  const press = slicerAntScrollPress
  if (press?.gesture.layout.id === id && press.pointerId === undefined) cancelSlicerAntScroll()
}
function beginSlicerAntScroll(payload: { id: string; event: PointerEvent; pane: SlicerControlPaneLayout }): void {
  closeSlicerLevel()
  const gesture = beginSlicerAntControl(payload)
  if (!gesture || !root.value || !gesture.layout.scrollbar || gesture.layout.maxScrollOffset <= 0) return
  const local = gesture.localStartPoint, thumb = gesture.layout.scrollThumb,
    part = thumb && local.x >= thumb.x && local.x <= thumb.x + thumb.width && local.y >= thumb.y && local.y <= thumb.y + thumb.height ? 'thumb' : 'track'
  slicerAntScrollPress = { gesture, pane: payload.pane, pointerId: payload.event.pointerId, part }
  gesturePointerId = payload.event.pointerId
  root.value.focus({ preventScroll: true })
  root.value.setPointerCapture(payload.event.pointerId)
  slicerScrollPreviews.value = new Map(slicerScrollPreviews.value).set(`${gesture.layout.model.type}:${gesture.layout.id}`,
    slicerCanvasScrollOffset(gesture, gesture.startPoint, part))
}
function previewSlicerAntScroll(payload: { id: string; value: number }): void {
  if (!Number.isFinite(payload.value)) return
  let press = slicerAntScrollPress
  if (press?.gesture.layout.id !== payload.id) {
    const layout = activateSlicerControl(payload.id)
    if (!layout) return
    press = { gesture: captureSlicerCanvasGesture(layout, { x: layout.bounds.x, y: layout.bounds.y }, slicerGestureContext(layout.id)), part: 'track' }
    slicerAntScrollPress = press
  }
  if (!currentSlicerControl(press.gesture)) { clearPointerGesture(); return }
  const layout = press.gesture.layout, step = layout.model.type === 'timeline' ? 1 : Math.max(1, layout.model.columnCount ?? 1),
    offset = Math.min(layout.maxScrollOffset, Math.max(0, Math.round(payload.value / step) * step))
  slicerScrollPreviews.value = new Map(slicerScrollPreviews.value).set(`${layout.model.type}:${layout.id}`, offset)
}
function finishSlicerAntScroll(payload: { id: string; value: number }): void {
  const press = slicerAntScrollPress
  if (!press || press.gesture.layout.id !== payload.id) return
  const current = currentSlicerControl(press.gesture), layout = press.gesture.layout,
    step = layout.model.type === 'timeline' ? 1 : Math.max(1, layout.model.columnCount ?? 1),
    offset = Math.min(layout.maxScrollOffset, Math.max(0, Math.round(payload.value / step) * step))
  cancelSlicerAntScroll()
  if (!current || !Number.isFinite(offset)) return
  const items = props.slicers?.find(control => control.model.id === layout.id)?.items ?? [], position = layout.model.type === 'timeline'
    ? timelineSerialAtOffset(items, layout.model.level ?? 'month', props.sheet.date1904 ?? false, offset, layout.timeline) : offset
  if (position !== undefined) emit('slicer-scroll-to', { slicerId: layout.id, position, expectedRevision: press.gesture.revision })
  if (press.pane) restoreSlicerControlFocus(press.gesture)
}
function onSlicerAntScrollMove(event: PointerEvent): boolean {
  const press = slicerAntScrollPress
  if (!press || press.pointerId !== event.pointerId || !press.pane) return false
  if (!currentSlicerControl(press.gesture)) { clearPointerGesture(); return true }
  const content = drawingContentPoint(point(event), undefined, press.pane.pane), offset = slicerCanvasScrollOffset(press.gesture, content, press.part)
  slicerScrollPreviews.value = new Map(slicerScrollPreviews.value).set(`${press.gesture.layout.model.type}:${press.gesture.layout.id}`, offset)
  event.preventDefault(); return true
}
function onSlicerAntScrollUp(event: PointerEvent): boolean {
  const press = slicerAntScrollPress
  if (!press || press.pointerId !== event.pointerId) return false
  if (event.type === 'pointerup') onSlicerAntScrollMove(event)
  const offset = slicerScrollPreviews.value.get(`${press.gesture.layout.model.type}:${press.gesture.layout.id}`) ?? press.gesture.layout.scrollOffset
  gesturePointerId = undefined
  press.pointerId = undefined
  if (root.value?.hasPointerCapture(event.pointerId)) root.value.releasePointerCapture(event.pointerId)
  if (event.type === 'pointerup') finishSlicerAntScroll({ id: press.gesture.layout.id, value: offset })
  else cancelSlicerAntScroll()
  return true
}
function slicerKeyboardAction(payload: { id: string; action: 'clear' | 'multiSelect' | 'item' | 'period'; event: MouseEvent; value?: string; index?: number }): void {
  const layout = activateSlicerControl(payload.id)
  if (!layout) return
  if (payload.action === 'clear') emit('slicer-clear', layout.id, props.workbookRevision)
  else if (payload.action === 'multiSelect') {
    const next = new Set(multiSelect.value); next.has(layout.id) ? next.delete(layout.id) : next.add(layout.id); multiSelect.value = next
  } else if (payload.action === 'item' && payload.value !== undefined && payload.index !== undefined) {
    focusedSlicerItems.value = new Map(focusedSlicerItems.value).set(layout.id, payload.index)
    emit('slicer-item', { slicerId: layout.id, value: payload.value, index: payload.index, anchorIndex: slicerAnchor.get(layout.id),
      ctrlKey: payload.event.ctrlKey, metaKey: payload.event.metaKey, shiftKey: payload.event.shiftKey, multiSelect: multiSelect.value.has(layout.id), expectedRevision: props.workbookRevision })
    if (!payload.event.shiftKey) slicerAnchor.set(layout.id, payload.index)
  } else if (payload.action === 'period') {
    const period = layout.periods.find(period => period.period.key === payload.index)?.period
    if (period) {
      focusedTimelinePeriods.value = new Map(focusedTimelinePeriods.value).set(layout.id, period.key)
      emit('slicer-range', { slicerId: layout.id, range: timelineCanvasSelectionRange(layout.model, period, period, props.sheet.date1904 ?? false), expectedRevision: props.workbookRevision })
    }
  }
}
function cancelSlicerControl(id: string): void {
  const menuOnly = timelineMenuId.value === id
  const gesture = slicerMenuGesture ?? slicerAntScrollPress?.gesture
  closeSlicerLevel(); clearPointerGesture()
  if (!menuOnly && props.selectedSlicerId === id) slicerContentId.value = ''
  if (gesture) restoreSlicerControlFocus(gesture)
  else root.value?.focus({ preventScroll: true })
}
function resetSlicerInteraction() {
  slicerLevelPress = undefined; slicerLevelPaneKey = ''; slicerMenuGesture = undefined; slicerMenuPaneKey.value = ''; cancelSlicerAntScroll()
  slicerContentId.value = ''; timelineMenuId.value = ''
  focusedSlicerItems.value = new Map(); focusedTimelinePeriods.value = new Map()
  slicerAnchor.clear(); multiSelect.value = new Set()
}
function emitChartEvent(event: MouseEvent, hit: DrawingHit): void {
  emit('chart-event', event, hit.leaf.id, { scrollLeft: hit.layout.frozenX ? 0 : viewport.value.scrollLeft, scrollTop: hit.layout.frozenY ? 0 : viewport.value.scrollTop,
    headerWidth: geometry.value.headerWidth, headerHeight: geometry.value.headerHeight })
}
/** Native chart hitboxes are an input bridge. Common Canvas order decides which drawing receives the event. */
function routeDrawingEvent(event: MouseEvent, chartId: string): boolean {
  if (!props.drawingScene || !root.value || !active || disposed) return false
  const position = point(event), hit = drawingCanvas.pick(position), controls = drawingCanvas.hasControlAt(position) || imageCanvas.hasControlAt(position),
    sharedSelection = drawingCanvas.sharedControls.value && !!hit && props.selectedDrawingRefs?.some(ref => ref.kind === hit.target.kind && ref.id === hit.target.id),
    additive = event.type === 'pointerdown' && (event.ctrlKey || event.metaKey || event.shiftKey
      && !(props.selectedDrawingRefs?.length === 1 && props.selectedDrawingRefs[0]?.kind === 'chart' && props.selectedDrawingRefs[0].id === chartId))
  if (!controls && !additive && !sharedSelection && hit?.leaf.kind === 'chart' && hit.leaf.id === chartId && hit.target.kind === 'chart'
    && !hit.layout.frozenX && !hit.layout.frozenY) return false
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
  if (event.target instanceof Element && event.target.closest('.sheet-filter-header')) return
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
  const hit = geometry.value.hitTest(Math.max(geometry.value.headerWidth + 1, Math.min(viewport.value.width - 1, position.x)),
    Math.max(geometry.value.headerHeight + 1, Math.min(viewport.value.height - 1, position.y)), viewport.value)
  if (hit?.kind !== 'cell') return
  fillDrag.target = { row: hit.row, column: hit.column }
  fillPreview.value = fillPreviewRange(fillDrag.source, fillDrag.target)
  if (root.value) root.value.style.cursor = 'crosshair'
  queueRender()
}
function captureFillPointer(pointerId: number): boolean {
  const element = root.value
  if (!element?.isConnected) { clearPointerGesture(); return false }
  try {
    if (!element.hasPointerCapture(pointerId)) element.setPointerCapture(pointerId)
    return true
  } catch (error) {
    clearPointerGesture()
    const OwnerDOMException = element.ownerDocument.defaultView?.DOMException
    if (OwnerDOMException && error instanceof OwnerDOMException && (error.name === 'NotFoundError' || error.name === 'InvalidStateError')) return false
    throw error
  }
}
function onEditorFillPointerDown(event: PointerEvent): void {
  if (event.button !== 0 || !props.editingKey || props.formatPainter || !active || disposed) return
  const position = point(event)
  const handle = event.target instanceof Element ? event.target.closest('.canvas-cell-fill-handle') : null
  const edited = parseAddress(props.editingKey)
  // The visible handle owns its pointer hit, including its inner half over Ant Input.
  // Geometric hit testing also covers hosts forwarding the native input event.
  const source = handle && root.value?.contains(handle) && edited
    ? geometry.value.mergeAt(edited.row, edited.column)
      ?? { startRow: edited.row, endRow: edited.row, startColumn: edited.column, endColumn: edited.column }
    : hitCellEditorFillHandle(geometry.value, props.editingKey, viewport.value, position.x, position.y)
  if (!source) return
  // Capture before antdv-next's input handlers and before native blur can schedule another commit.
  event.preventDefault(); event.stopPropagation()
  if (composing.value || gesturePointerId !== undefined || !props.prepareFill) return
  const press = { pointerId: event.pointerId, source: { ...source }, origin: snapshotSheetFillGesture(fillGestureContext()), startPosition: position, position }
  pendingFillPress = press
  gesturePointerId = event.pointerId
  if (!captureFillPointer(event.pointerId)) return
  let committed: SheetFillCommitContext | undefined
  try { committed = props.prepareFill() }
  catch (error) { clearPointerGesture(); throw error }
  if (!committed) { clearPointerGesture(); return }
  const accepted = committed
  void nextTick(() => {
    if (pendingFillPress !== press || gesturePointerId !== press.pointerId) return
    pendingFillPress = undefined
    if (!active || disposed || composing.value || props.formatPainter || !root.value
      || !canResumeSheetFillGesture(press.origin, accepted, fillGestureContext())) {
      clearPointerGesture(); return
    }
    // The input/handle can be removed by the accepted commit before the browser
    // applies pending capture. Keep ownership on the stable Canvas host.
    if (!captureFillPointer(press.pointerId)) return
    fillDrag = { source: press.source, context: snapshotSheetFillGesture(fillGestureContext()), startPosition: press.startPosition, moved: false }
    lastPointer = press.position
    root.value.focus({ preventScroll: true })
    updateFillDragTarget(press.position)
  })
}
function onPointerDown(event: PointerEvent) {
  if (gesturePointerId !== undefined || drawingCanvas.hasGesture() || imageCanvas.hasGesture() || !active || disposed) return
  if ((event.target as HTMLElement).closest('input, textarea')) return
  closeSlicerLevel()
  const position = point(event)
  const selectedControl = props.selectedDrawingRefs?.length === 1 ? props.selectedDrawingRefs[0] : undefined
  if (selectedControl && drawingCanvas.hasControlAt(position)) { slicerContentId.value = ''; timelineMenuId.value = '' }
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
  const slicerHit = hitTestSlicerObjects(drawingHit ? slicerLayouts.value.filter(layout => layout.id === drawingHit.leaf.id) : slicerLayouts.value, content.x, content.y)
  if (slicerHit) {
    if (event.button !== 0) return
    event.preventDefault()
    event.stopPropagation()
    root.value?.focus({ preventScroll: true })
    emit('select-slicer', slicerHit.id)
    const layout = slicerLayouts.value.find(layout => layout.id === slicerHit.id)
    if (!layout) return
    if (!slicerControlEditable(layout)) return
    slicerContentId.value = slicerHit.type === 'move' || slicerHit.type === 'resize' ? '' : slicerHit.id
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
    slicerPress = { hit: slicerHit, start: slicerHit, gesture: captureSlicerCanvasGesture(layout, content, slicerGestureContext(layout.id)), pane: drawingHit?.layout }
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
    const control = outlineToggleRects.value.find(control => control.axis === hit.axis && control.groupId === hit.groupId)
    if (control) changeOutlineControl(control)
  } else if (hit.kind === 'outline-level') {
    const control = outlineLevelButtons.value.find(control => control.axis === hit.axis && control.level === hit.level)
    if (control) changeOutlineLevel(control)
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
  if (gesturePointerId !== undefined) clearValidationHover()
  if (onSlicerAntScrollMove(event)) { clearValidationHover(); return }
  if (drawingCanvas.move(point(event), event)) { clearValidationHover(); return }
  if (imageCanvas.move(point(event), event)) { clearValidationHover(); return }
  if (!active || disposed || gesturePointerId !== undefined && gesturePointerId !== event.pointerId) return
  if (richTextPress) { updateRichTextPointer(point(event)); return }
  if (gesturePointerId === undefined) {
    const drawingHit = drawingCanvas.pick(point(event))
    if (drawingHit?.leaf.kind === 'chart' && drawingHit.target.kind === 'chart') { clearValidationHover(); emitChartEvent(event, drawingHit); return }
    if (drawingHit?.target.kind === 'group' || drawingHit?.target.kind === 'shape') { clearValidationHover(); if (root.value) root.value.style.cursor = 'move'; return }
  }
  const position = point(event), hit = geometry.value.hitTest(Math.max(geometry.value.headerWidth + 1, Math.min(viewport.value.width - 1, position.x)), Math.max(geometry.value.headerHeight + 1, Math.min(viewport.value.height - 1, position.y)), viewport.value)
  lastPointer = position
  if (pendingFillPress) { pendingFillPress.position = position; return }
  if (headerDrag.value || headerSelectionDrag) { updateHeaderGesture(position, event.ctrlKey || event.metaKey); return }
  if (slicerPress && ['period', 'timelineStart', 'timelineEnd', 'scroll'].includes(slicerPress.start.type)) {
    if (!currentSlicerGesture(slicerPress)) { clearPointerGesture(); return }
    const content = drawingContentPoint(position, undefined, slicerPress.pane)
    const hit = hitTestSlicerObjects([slicerPress.gesture.layout], content.x, content.y)
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
    scheduleValidationHover(!props.editingKey && !fillHandle && hover?.kind === 'cell' ? hover.key : undefined)
    const hasHyperlink = hover?.kind === 'cell' && Boolean(props.sheet.cells[hover.key]?.hyperlink)
    root.value.style.cursor = fillHandle ? 'crosshair' : hasHyperlink || hover?.kind === 'outline-toggle' || hover?.kind === 'outline-level' ? 'pointer' : resizeDrag?.axis === 'column' || hover?.kind === 'resize-column' ? 'col-resize' : resizeDrag?.axis === 'row' || hover?.kind === 'resize-row' ? 'row-resize' : hover && selectedAxisMoveSource(hover) ? 'grab' : 'cell'
  }
}
function onPointerUp(event: PointerEvent) {
  // Child editors and Ant controls may release their own capture while the root
  // still owns the active fill. A bubbled child loss cannot end that gesture.
  if (event.type === 'lostpointercapture' && event.target !== root.value) return
  if (onSlicerAntScrollUp(event)) return
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
    if (!canceled && currentSlicerGesture(press)) {
      const position = point(event)
      const content = drawingContentPoint(position, undefined, press.pane)
      const localContent = slicerObjectLocalPoint(press.gesture.layout, content)
      const hit = hitTestSlicerObjects([press.gesture.layout], content.x, content.y)
      if (timelineMenuId.value && hit?.type !== 'timelineLevel') closeSlicerLevel()
      if (press.start.type === 'move' || press.start.type === 'resize') {
        const layout = press.gesture.layout
        const dx = content.x - press.gesture.startPoint.x
        const dy = content.y - press.gesture.startPoint.y
        if (layout && !layout.model.lockedPosition && (Math.abs(dx) >= 1 || Math.abs(dy) >= 1)) {
          const next = slicerCanvasGestureGeometry(press.gesture, content, press.start.type)
          if (props.drawingScene) emit('drawing-command', SHEET_DRAWING_GROUP_COMMAND_IDS.transform, { sheetId: props.sheet.id,
            selection: [{ kind: layout.model.type, id: layout.id }], expectedSelection: props.selectedDrawingRefs ? [...props.selectedDrawingRefs] : undefined,
            expectedRevision: press.gesture.revision, delta: slicerCanvasGestureDelta(layout.geometry, next) })
          else emit('slicer-geometry', { slicerId: layout.id, geometry: next, expectedRevision: press.gesture.revision })
        }
      } else if (press.start.type === 'scroll' && press.start.part === 'thumb') {
        const layout = slicerLayouts.value.find(({ id }) => id === press.start.id)
        if (layout && layout.scrollbar && layout.scrollThumb) {
          const delta = layout.model.type === 'timeline' ? localContent.x - press.gesture.localStartPoint.x : localContent.y - press.gesture.localStartPoint.y
          const trackLength = layout.model.type === 'timeline' ? layout.scrollbar.width - layout.scrollThumb.width : layout.scrollbar.height - layout.scrollThumb.height
          const position = trackLength > 0 ? clampOffset(layout.scrollOffset + delta / trackLength * layout.maxScrollOffset, layout.maxScrollOffset) : 0
          if (layout.model.type === 'timeline') {
            const items = props.slicers?.find(control => control.model.id === layout.id)?.items ?? []
            const serial = timelineSerialAtOffset(items, layout.model.level ?? 'month', props.sheet.date1904 ?? false, position, layout.timeline)
            if (serial !== undefined) emit('slicer-scroll-to', { slicerId: layout.id, position: serial, expectedRevision: press.gesture.revision })
          } else emit('slicer-scroll-to', { slicerId: layout.id, position, expectedRevision: press.gesture.revision })
        }
      } else if (hit?.id === press.start.id) {
        if (press.start.type === 'clear' && hit.type === 'clear') emit('slicer-clear', hit.id, press.gesture.revision)
        else if (press.start.type === 'multiSelect' && hit.type === 'multiSelect') {
          const next = new Set(multiSelect.value)
          next.has(hit.id) ? next.delete(hit.id) : next.add(hit.id)
          multiSelect.value = next
        } else if (press.start.type === 'item' && hit.type === 'item') {
          emit('slicer-item', { slicerId: hit.id, value: hit.value, index: hit.index, anchorIndex: slicerAnchor.get(hit.id), ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, multiSelect: multiSelect.value.has(hit.id), expectedRevision: press.gesture.revision })
          if (!event.shiftKey) slicerAnchor.set(hit.id, hit.index)
        } else if (press.start.type === 'period' && hit.type === 'period') {
          const layout = slicerLayouts.value.find(({ id }) => id === hit.id)
          const first = Math.min(press.start.index, hit.index), last = Math.max(press.start.index, hit.index)
          const periods = layout?.periods.filter(({ period }) => period.key >= first && period.key <= last).sort((a, b) => a.period.key - b.period.key)
          if (periods?.length && layout) emit('slicer-range', { slicerId: hit.id,
            range: timelineCanvasSelectionRange(layout.model, periods[0]!.period, periods[periods.length - 1]!.period, props.sheet.date1904 ?? false), expectedRevision: press.gesture.revision })
        } else if ((press.start.type === 'timelineStart' || press.start.type === 'timelineEnd') && (hit.type === 'period' || hit.type === 'timelineStart' || hit.type === 'timelineEnd')) {
          const layout = slicerLayouts.value.find(({ id }) => id === press.start.id)
          const period = hit.type === 'period' ? hit.period : layout?.periods.find(({ rect }) => localContent.x >= rect.x && localContent.x < rect.x + rect.width)?.period
          if (layout && period && Math.abs(localContent.x - press.gesture.localStartPoint.x) >= 1) {
            const range = timelineCanvasHandleRange(layout.model, period, press.start.type === 'timelineStart' ? 'start' : 'end', props.sheet.date1904 ?? false)
            if (range) emit('slicer-range', { slicerId: layout.id, range, expectedRevision: press.gesture.revision })
          }
        } else if (press.start.type === 'timelineLevel' && hit.type === 'timelineLevel') {
          if (timelineMenuId.value === hit.id) closeSlicerLevel(true)
          else {
            const pane = slicerControlPanes.value.find(pane => pane.layout.id === hit.id
              && (!press.pane || pane.pane.frozenX === press.pane.frozenX && pane.pane.frozenY === press.pane.frozenY))
            if (pane) { slicerMenuGesture = press.gesture; slicerMenuPaneKey.value = pane.key; timelineMenuId.value = hit.id }
          }
          queueRender()
        } else if (press.start.type === 'scroll' && hit.type === 'scroll' && press.start.part === 'track') {
          const layout = slicerLayouts.value.find(({ id }) => id === hit.id)
          if (layout && layout.scrollbar) {
            const coordinate = layout.model.type === 'timeline' ? localContent.x - layout.scrollbar.x : localContent.y - layout.scrollbar.y
            const extent = layout.model.type === 'timeline' ? layout.scrollbar.width : layout.scrollbar.height
            const position = clampOffset(Math.round(coordinate / Math.max(1, extent) * layout.maxScrollOffset), layout.maxScrollOffset)
            if (layout.model.type === 'timeline') {
              const items = props.slicers?.find(control => control.model.id === layout.id)?.items ?? []
              const serial = timelineSerialAtOffset(items, layout.model.level ?? 'month', props.sheet.date1904 ?? false, position, layout.timeline)
              if (serial !== undefined) emit('slicer-scroll-to', { slicerId: layout.id, position: serial, expectedRevision: press.gesture.revision })
            } else emit('slicer-scroll-to', { slicerId: layout.id, position, expectedRevision: press.gesture.revision })
          }
        }
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
  slicerPress = undefined
  slicerLevelPress = undefined
  cancelSlicerAntScroll()
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
  const slicerHit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y)
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
  const slicerHit = hitTestSlicerObjects(slicerLayouts.value, content.x, content.y)
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
function openValidationFromKey(event: KeyboardEvent): boolean {
  if (!active || disposed || composing.value || event.isComposing || event.keyCode === 229 || event.defaultPrevented
    || !event.altKey || event.ctrlKey || event.metaKey || event.key !== 'ArrowDown' || props.formatPainter
    || props.selectedDrawingRefs?.length || !props.validationCellEntry || props.validationCellEntry.disabled) return false
  if (!props.validationCellEntry.open()) return false
  event.preventDefault(); event.stopPropagation(); return true
}
function onEditorKeydownCapture(event: KeyboardEvent): void {
  if (!(event.target as HTMLElement).closest('.canvas-cell-input')) return
  nativeInput.capture(event, 'selection')
  // Capture before AutoComplete's option list can interpret an IME Enter or Alt+Enter as a choice.
  if (composing.value || event.isComposing || event.keyCode === 229) { event.stopPropagation(); return }
  if (openValidationFromKey(event)) return
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
  const hit = hitTestSlicerObjects(drawingHit ? slicerLayouts.value.filter(layout => layout.id === drawingHit.leaf.id) : slicerLayouts.value, content.x, content.y)
  if (!hit) return
  event.preventDefault()
  const layout = slicerLayouts.value.find(item => item.id === hit.id)
  if (!layout || !event.deltaY) return
  if (layout.model.type === 'timeline') {
    const step = Math.max(1, Math.round(Math.abs(event.deltaY) / 100))
    const items = props.slicers?.find(control => control.model.id === layout.id)?.items ?? []
    const position = timelineSerialAtOffset(items, layout.model.level ?? 'month', props.sheet.date1904 ?? false, layout.scrollOffset + Math.sign(event.deltaY) * step, layout.timeline)
    if (position !== undefined) emit('slicer-scroll-to', { slicerId: layout.id, position, expectedRevision: props.workbookRevision })
  } else {
    const columns = Math.max(1, layout.model.columnCount ?? 1)
    const step = Math.max(1, Math.round(Math.abs(event.deltaY) / Math.max(1, layout.model.rowHeight ?? 24))) * columns
    emit('slicer-scroll-to', { slicerId: layout.id, position: clampOffset(layout.scrollOffset + Math.sign(event.deltaY) * step, layout.maxScrollOffset), expectedRevision: props.workbookRevision })
  }
}
function clampOffset(value: number, max: number): number { return Math.min(max, Math.max(0, Math.round(value))) }
function onKeydown(event: KeyboardEvent) {
  if (!active || disposed || event.defaultPrevented) return
  if (!props.editingKey && event.target === root.value && event.key === ' ' && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey
    && !composing.value && !event.isComposing && event.keyCode !== 229 && !props.formatPainter && !props.selectedDrawingRefs?.length) {
    const checkbox = checkboxControls.value.entries.get(props.selectedKey)
    if (checkbox && !checkbox.disabled) { event.preventDefault(); event.stopPropagation(); checkbox.activate(); return }
  }
  if (!props.editingKey && openValidationFromKey(event)) return
  if (event.key === 'Escape' && !event.isComposing && event.keyCode !== 229 && (slicerAntScrollPress || slicerMenuGesture || slicerLevelPress)) {
    event.preventDefault(); event.stopPropagation(); cancelSlicerControl(props.selectedSlicerId ?? ''); return
  }
  if (event.key === 'Escape' && !event.isComposing && event.keyCode !== 229 && slicerPress) {
    slicerContentId.value = ''; timelineMenuId.value = ''
    event.preventDefault(); event.stopPropagation(); clearPointerGesture(); return
  }
  const selectedControl = props.selectedDrawingRefs?.length === 1 && ['slicer', 'timeline'].includes(props.selectedDrawingRefs[0]!.kind)
    ? props.selectedDrawingRefs[0] : undefined
  if (selectedControl && !props.editingKey && !event.isComposing && event.keyCode !== 229 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    if ((event.key === 'Enter' || event.key === ' ') && slicerContentId.value !== selectedControl.id) {
      slicerContentId.value = selectedControl.id
      event.preventDefault(); event.stopPropagation(); queueRender(); return
    }
    if (event.key === 'Escape' && slicerContentId.value === selectedControl.id && timelineMenuId.value !== selectedControl.id) {
      slicerContentId.value = ''; event.preventDefault(); event.stopPropagation(); queueRender(); return
    }
  }
  if ((event.key === 'ContextMenu' || event.key === 'F10' && event.shiftKey) && !event.ctrlKey && !event.metaKey && !event.altKey
    && drawingCanvas.sharedControls.value && !event.isComposing && event.keyCode !== 229) {
    const menu = drawingCanvas.menuPosition.value
    if (menu) { event.preventDefault(); event.stopPropagation(); drawingContextPosition.value = { x: menu.left, y: menu.top + 34 }; drawingContextOpen.value = true; return }
  }
  if (drawingCanvas.keydown(event)) return
  if (imageCanvas.keydown(event)) return
  const selectedChart = props.selectedDrawingRefs?.length === 1 && props.selectedDrawingRefs[0]?.kind === 'chart' ? props.selectedDrawingRefs[0] : undefined
  if (selectedChart && !props.editingKey && !props.imageInteractionBlocked) {
    emit('chart-event', event, selectedChart.id, { scrollLeft: viewport.value.scrollLeft, scrollTop: viewport.value.scrollTop,
      headerWidth: geometry.value.headerWidth, headerHeight: geometry.value.headerHeight })
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
    else if (event.key === 'End') next = layout.totalItemCount - 1
    else if (event.key === 'Enter' || event.key === ' ') {
      const item = layout.items.find(candidate => candidate.index === current)
      if (item) {
        emit('slicer-item', { slicerId: layout.id, value: item.value, index: item.index, anchorIndex: slicerAnchor.get(layout.id), ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, multiSelect: multiSelect.value.has(layout.id), expectedRevision: props.workbookRevision })
        if (!event.shiftKey) slicerAnchor.set(layout.id, item.index)
      }
      event.preventDefault(); event.stopPropagation(); return
    } else if (event.key === 'Escape') { multiSelect.value.delete(layout.id); event.preventDefault(); event.stopPropagation(); queueRender(); return }
    else return
    const bounded = Math.min(layout.totalItemCount - 1, Math.max(0, next))
    if (bounded < visible[0]! || bounded > visible.at(-1)!) {
      const startItem = bounded > visible.at(-1)! ? Math.max(0, bounded - layout.visibleCount + columns) : Math.floor(bounded / columns) * columns
      emit('slicer-scroll-to', { slicerId: layout.id, position: startItem, expectedRevision: props.workbookRevision })
    }
    focusedSlicerItems.value = new Map(focusedSlicerItems.value).set(layout.id, bounded)
  } else {
    const periods = layout.periods
    if (event.key === 'Escape' && timelineMenuId.value === layout.id) { closeSlicerLevel(true); event.preventDefault(); queueRender(); return }
    if (!periods.length) return
    const current = focusedTimelinePeriods.value.get(layout.id) ?? periods[0]!.period.key
    if (timelineMenuId.value === layout.id) {
      const levels: TimelineLevel[] = ['year', 'quarter', 'month', 'day']
      const index = levels.indexOf(layout.model.level ?? 'month')
      if (event.key === 'Enter' || event.key === ' ') chooseSlicerLevel({ id: layout.id, level: levels[index]! })
      else if (event.key === 'ArrowDown' || event.key === 'ArrowRight') chooseSlicerLevel({ id: layout.id, level: levels[Math.min(3, index + 1)]! })
      else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') chooseSlicerLevel({ id: layout.id, level: levels[Math.max(0, index - 1)]! })
      else return
      closeSlicerLevel(); event.preventDefault(); event.stopPropagation(); queueRender(); return
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
      if (period) emit('slicer-range', { slicerId: layout.id, range: timelineCanvasSelectionRange(layout.model, period, period, props.sheet.date1904 ?? false), expectedRevision: props.workbookRevision })
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
defineExpose({ focusEditor, insertEditorText, ensureVisible, focus, featureSearchContext, getBackend: () => surface?.backend, operateImageCrop: imageCanvas.operateImageCrop, routeDrawingEvent })
watch(scene, queueRender)
watch([() => props.workbookId, () => props.sheet.id, () => props.workbookRevision, () => props.workbookHistoryGeneration, () => props.currentUserId], () => { contextMenuOpen.value = false })
watch(() => [geometry.value.headerWidth, geometry.value.headerHeight] as const, ([headerWidth, headerHeight]) => {
  viewport.value = { ...viewport.value, headerWidth, headerHeight }
  void nextTick(() => { resize(); ensureVisible(props.selectedKey) })
}, { flush: 'sync' })
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
  resetSlicerInteraction()
  clearPointerGesture()
  void nextTick(restoreViewport)
}, { flush: 'sync' })
watch([() => props.slicers?.[0]?.model.unitId, () => props.currentUserId], () => {
  resetSlicerInteraction()
  clearPointerGesture()
}, { flush: 'sync' })
watch(() => slicerLayouts.value.filter(layout => layout.model.type === 'slicer' && !layout.dataStatus?.error
  && layout.items.length < Math.min(layout.visibleCount, Math.max(0, layout.totalItemCount - layout.scrollOffset)))
  .map(layout => ({ slicerId: layout.id, start: layout.scrollOffset, count: layout.visibleCount, expectedRevision: props.workbookRevision })),
  requests => { for (const request of requests) emit('slicer-window', request) }, { immediate: true, flush: 'post' })
watch([() => props.selectedSlicerId, () => props.selectedDrawingRefs?.map(ref => `${ref.kind}:${ref.id}`).join('|'),
  () => props.editingKey, () => props.imageInteractionBlocked, () => props.formatPainter], () => {
  const ref = props.selectedDrawingRefs?.length === 1 ? props.selectedDrawingRefs[0] : undefined
  if (props.editingKey || props.imageInteractionBlocked || props.formatPainter || slicerContentId.value !== props.selectedSlicerId
    || props.selectedDrawingRefs && (!ref || !['slicer', 'timeline'].includes(ref.kind) || ref.id !== slicerContentId.value)) {
    slicerContentId.value = ''; closeSlicerLevel()
  }
  if (slicerPress && !currentSlicerGesture(slicerPress)) clearPointerGesture()
  if (slicerAntScrollPress && !currentSlicerControl(slicerAntScrollPress.gesture)) clearPointerGesture()
  if (slicerMenuGesture && !currentSlicerControl(slicerMenuGesture)) closeSlicerLevel()
}, { flush: 'sync' })
watch(viewport, () => {
  if (slicerPress && !currentSlicerGesture(slicerPress) || slicerAntScrollPress && !currentSlicerControl(slicerAntScrollPress.gesture)) clearPointerGesture()
  if (slicerMenuGesture && !currentSlicerControl(slicerMenuGesture)) closeSlicerLevel()
  if (slicerLevelPress && !currentSlicerControl(slicerLevelPress)) slicerLevelPress = undefined
}, { flush: 'sync' })
watch([() => props.sheet, () => props.workbookRevision, () => props.editingKey, () => props.editSessionId], () => {
  // The synchronous commit updates these props together. Validate its acknowledgement after that flush.
  if (pendingFillPress && pendingFillPress.origin.sheetId === props.sheet.id) return
  clearPointerGesture()
  if (slicerMenuGesture && !currentSlicerControl(slicerMenuGesture)) closeSlicerLevel()
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
  document.addEventListener('pointerdown', onSlicerMenuOutsidePointer, true)
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
  clearValidationHover()
  pasteOptionsOpen.value = false
  drawingContextOpen.value = false
  resetSlicerInteraction()
  sheetViews.set(props.sheet.id, { left: viewport.value.scrollLeft, top: viewport.value.scrollTop })
  active = false
  cancelAnimationFrame(frame); frame = 0
  clearPointerGesture()
  surface?.resize(0, 0, window.devicePixelRatio || 1)
})
onUnmounted(() => { disposed = true; clearValidationHover(); pasteOptionsOpen.value = false; drawingContextOpen.value = false; resetSlicerInteraction(); document.removeEventListener('focusin', onPasteOptionsFocusIn, true); document.removeEventListener('pointerdown', onSlicerMenuOutsidePointer, true); clearPointerGesture(); cancelAnimationFrame(frame); observer?.disconnect(); dprQuery?.removeEventListener('change', onDevicePixelRatioChange); window.removeEventListener('resize', resize); window.visualViewport?.removeEventListener('resize', resize); document.fonts?.removeEventListener('loadingdone', queueRender); surface?.dispose() })
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
    @pointerleave="clearValidationHover"
    @copy="event => onClipboard(event, 'copy')" @cut="event => onClipboard(event, 'cut')"
    @paste.capture="onEditorPasteCapture" @paste="event => onClipboard(event, 'paste')" @wheel="onWheel">
    <div ref="scroller" class="canvas-scroll" @scroll="onScroll"><div :style="fullSize" /></div>
    <div ref="canvasHost" class="canvas-surface" :style="{ width: `${viewport.width}px`, height: `${viewport.height}px` }" />
    <SlicerControlOverlay v-for="pane in slicerControlPanes" :key="pane.key" :pane="pane" :panes="slicerControlPanes"
      :profile="slicerProfile(pane.layout)" :colors="slicerProfile(pane.layout).palette" :disabled="!slicerControlEditable(pane.layout)"
      :menu-open="timelineMenuId === pane.layout.id && slicerMenuPaneKey === pane.key" :multi-selected="multiSelect.has(pane.layout.id)"
      :focused-item="focusedSlicerItems.get(pane.layout.id)" :focused-period="focusedTimelinePeriods.get(pane.layout.id)"
      @native-pointer="onPointerDown" @activate="activateSlicerControl" @keyboard-action="slicerKeyboardAction"
      @level-pointer="beginSlicerLevel" @level-toggle="toggleSlicerLevel" @level-change="chooseSlicerLevel"
      @level-close="id => { if (timelineMenuId === id && slicerMenuPaneKey === pane.key) closeSlicerLevel(true) }"
      @scroll-pointer="beginSlicerAntScroll" @scroll-change="previewSlicerAntScroll" @scroll-complete="finishSlicerAntScroll"
      @scroll-leave="leaveSlicerAntScroll" @cancel="cancelSlicerControl" @retry="id => emit('slicer-retry', id)" />
    <div v-if="drawingCanvas.menuPosition.value" class="sheet-drawing-float-menu" role="toolbar" aria-label="绘图组合工具"
      :style="{ left: `${drawingCanvas.menuPosition.value.left}px`, top: `${drawingCanvas.menuPosition.value.top}px`, background: token.colorBgElevated, borderColor: token.colorBorder }"
      @pointerdown.stop @dblclick.stop @contextmenu.prevent.stop @keydown.stop>
      <Tooltip v-if="drawingMenuSlicer" :title="drawingMenuSlicer.type === 'timeline' ? '时间线设置' : '切片器设置'"><Button size="small" :disabled="!drawingCanvas.editable.value" @click="chooseDrawingMenu('slicer:settings')">设置</Button></Tooltip>
      <Tooltip v-if="drawingMenuSlicer && !drawingMenuSlicer.tableId" title="报表连接"><Button size="small" :disabled="!drawingCanvas.editable.value" @click="chooseDrawingMenu('slicer:connections')">报表连接</Button></Tooltip>
      <Dropdown :trigger="['click']" :menu="drawingGroupMenu" :auto-focus="false" destroy-on-hidden>
        <Button size="small" :disabled="!drawingCanvas.canGroup.value && !drawingCanvas.canUngroup.value" aria-label="组合或取消组合">组合</Button>
      </Dropdown>
      <Dropdown :trigger="['click']" :menu="drawingArrangeMenu" :auto-focus="false" destroy-on-hidden>
        <Button size="small" :disabled="!drawingCanvas.editable.value" aria-label="排列绘图对象">排列</Button>
      </Dropdown>
      <Dropdown :trigger="['click']" :menu="drawingAlignMenu" :auto-focus="false" destroy-on-hidden>
        <Button size="small" :disabled="!drawingAlignEnabled" aria-label="对齐绘图对象">对齐</Button>
      </Dropdown>
      <Tooltip title="删除当前选择的绘图对象"><Button size="small" danger :disabled="!drawingCanvas.editable.value" @click="chooseDrawingMenu('remove')">删除</Button></Tooltip>
    </div>
    <div v-if="drawingContextOpen" class="sheet-drawing-context-anchor" :style="{ left: `${drawingContextPosition.x}px`, top: `${drawingContextPosition.y}px` }">
      <Dropdown :trigger="[]" :open="drawingContextOpen" :menu="drawingContextMenu" :classes="{ root: drawingContextPopupClass }"
        :popup-render="drawingContextPopup" auto-focus destroy-on-hidden @update:open="onDrawingContextOpenChange">
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
    <Tooltip v-for="control in outlineToggleRects" :key="`outline-${control.axis}-${control.groupId}`" :title="control.label">
      <Button class="outline-control outline-toggle" :style="{ left: `${control.x}px`, top: `${control.y}px` }"
        type="default" size="small" :aria-label="control.label" :aria-expanded="!control.collapsed" :disabled="control.disabled"
        @pointerdown.stop @keydown.stop @dblclick.stop @click.stop="changeOutlineControl(control)">
        <component :is="control.collapsed ? Plus : Minus" :size="10" aria-hidden="true" />
      </Button>
    </Tooltip>
    <Tooltip v-for="control in filterHeaderControls" :key="`filter-${control.key}`" :title="control.active ? '筛选条件已应用' : '筛选'">
      <Button class="sheet-filter-header" type="default" size="small" :data-filter-column="control.column"
        :style="{ left: `${control.rect.x}px`, top: `${control.rect.y}px`, width: `${control.rect.width}px`, height: `${control.rect.height}px`, clipPath: control.clipPath }"
        :aria-label="`筛选第 ${control.column + 1} 列`" :aria-pressed="control.active" :disabled="control.disabled"
        @pointerdown.stop.prevent @keydown.stop @dblclick.stop @contextmenu.prevent.stop @click.stop="openFilterHeader(control.column)">
        <component :is="control.active ? Filter : ListFilter" :size="Math.max(8, 12 * viewport.zoom)" aria-hidden="true" />
      </Button>
    </Tooltip>
    <Tooltip v-for="control in outlineLevelButtons" :key="`outline-${control.axis}-level-${control.level}`" :title="control.label">
      <Button class="outline-control outline-level" :style="{ left: `${control.x}px`, top: `${control.y}px` }"
        :type="control.active ? 'primary' : 'default'" size="small" :aria-label="control.label" :aria-pressed="control.active" :disabled="control.disabled"
        @pointerdown.stop @keydown.stop @dblclick.stop @click.stop="changeOutlineLevel(control)">{{ control.level }}</Button>
    </Tooltip>
    <SheetValidationCellControl v-if="validationCellEntry && validationCellControlPosition" :key="validationCellEntry.hostKey"
      :entry="validationCellEntry" :popup="validationCellPopup" :position="validationCellControlPosition" :region="region"
      :current-year="validationCurrentYear ?? new Date().getFullYear()" :dropdown-ui="validationDropdownUi" />
    <SheetCheckboxValidationOverlay v-for="control in checkboxControls.controls" :key="control.entry.hostKey"
      :entry="control.entry" :position="control.position" :selected-key="selectedKey" />
    <Alert v-if="validationCheckboxError" class="sheet-validation-checkbox-error" type="warning" :title="validationCheckboxError"
      show-icon closable @pointerdown.stop @close="emit('clear-checkbox-error')" />
    <SheetValidationFeedbackPopup v-if="validationPromptPosition && selectedValidationFeedback" :key="validationPromptOwner"
      :feedback="selectedValidationFeedback" :position="validationPromptPosition" mode="prompt" :input-owner="validationInputOwner ?? ''"
      :composing="composing" :dismiss="dismissValidationPrompt" />
    <SheetValidationFeedbackPopup v-if="hoveredValidationPosition && hoveredValidationFeedback" :key="`invalid:${hoveredValidationFeedback.key}`"
      :feedback="hoveredValidationFeedback" :position="hoveredValidationPosition" mode="invalid" :input-owner="validationInputOwner ?? ''" :composing="composing" />
    <AutoComplete v-if="editingKey" :key="editSessionId ?? editingKey" ref="input" class="canvas-cell-input" :style="editStyle" :styles="editInputStyles" variant="borderless" :value="editValue" :options="(validationOptions ?? []).map(value => ({ value, label: value }))" aria-label="单元格编辑" @change="onEditValueChange" @select="onAutoCompleteSelect" @blur="onEditorBlur" @keydown.stop="onInputKeydown">
      <Input.TextArea v-if="useMultilineEditor" :rows="1" :auto-size="false" :change-on-composing="true" :styles="editTextAreaStyles" :wrap="sheetCellWrapStrategy(editedCell) === 'wrap' ? 'soft' : 'off'" variant="borderless" spellcheck="false" aria-label="单元格编辑" @change="onTextAreaChange" />
    </AutoComplete>
    <span v-if="editingKey && editLayout && !formatPainter" class="canvas-cell-fill-handle" :style="fillHandleStyle" aria-hidden="true" @pointerdown="onEditorFillPointerDown" />
    <div class="a11y-cell" role="row"><span role="gridcell" aria-selected="true" :aria-invalid="selectedValidationFeedback?.issues.length ? 'true' : undefined">{{ selectedKey }} {{ selectedCellText }}</span></div>
    <Alert v-if="surfaceError" class="canvas-surface-error" type="error" :title="`画布初始化失败：${surfaceError}`" show-icon />
    <SheetContextMenu
      v-if="contextMenuContext"
      v-bind="contextMenuContext"
      v-model:open="contextMenuOpen"
      :hit="contextMenuHit"
      :x="contextMenuPosition.x"
      :y="contextMenuPosition.y"
      @command="(commandId, params) => emit('context-command', commandId, params)"
    />
  </div>
</template>

<style scoped>
.spreadsheet-canvas { position: relative; min-width: 0; min-height: 0; flex: 1; overflow: hidden; outline: none; color: var(--sheet-canvas-text); background: var(--sheet-canvas-background); touch-action: none; user-select: none; }
.spreadsheet-canvas:focus-visible { box-shadow: inset 0 0 0 1px var(--sheet-canvas-primary); }
.spreadsheet-canvas.format-painter-active { cursor: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='28' height='28' viewBox='0 0 28 28'%3E%3Cpath d='M2 1v18l4-4 3 7 3-2-3-6h6z' fill='white' stroke='black'/%3E%3Cpath d='m17 13 8-10 2 2-8 10z' fill='white' stroke='black'/%3E%3Cpath d='M17 14c-4 0-2 5-6 6 5 2 10-1 8-5z' fill='black' stroke='white'/%3E%3C/svg%3E") 2 1, crosshair !important; }
.canvas-scroll { position: absolute; inset: 0; overflow: auto; overscroll-behavior: contain; direction: ltr; }
.canvas-surface { position: absolute; inset: 0; pointer-events: none; overflow: hidden; }
.sheet-paste-options { position: absolute; z-index: 4; display: inline-flex; align-items: center; justify-content: center; gap: 2px; width: 44px; height: 24px; padding: 0 4px; color: var(--sheet-canvas-text); background: var(--sheet-canvas-background); border-color: var(--sheet-canvas-border); }
.sheet-paste-options:focus-visible { outline: 2px solid var(--sheet-canvas-primary); outline-offset: 1px; }
.outline-control { position: absolute; z-index: 3; box-sizing: border-box; width: 14px; height: 14px; min-width: 14px; min-height: 14px; margin: 0; padding: 0; border: 1px solid var(--sheet-canvas-border); border-radius: 1px; color: var(--sheet-canvas-text); background: var(--sheet-canvas-background); box-shadow: none; line-height: 10px; }
.sheet-filter-header { position: absolute; z-index: 4; min-width: 0; min-height: 0; padding: 0; border-radius: 1px; box-shadow: none; }
.sheet-filter-header:not(:disabled) { color: var(--sheet-canvas-text); background: var(--sheet-canvas-background); border-color: var(--sheet-canvas-border); }
.sheet-filter-header[aria-pressed="true"]:not(:disabled) { color: var(--sheet-canvas-primary); }
.sheet-filter-header :deep(.ant-btn-icon) { display: inline-flex; align-items: center; justify-content: center; }
.outline-control :deep(.ant-btn-icon) { display: inline-flex; align-items: center; justify-content: center; }
.outline-level { font-size: 9px; }
.outline-level.ant-btn-primary { color: var(--sheet-canvas-background); background: var(--sheet-canvas-primary); border-color: var(--sheet-canvas-primary); }
.canvas-cell-input { position: absolute; z-index: 3; display: block; box-sizing: border-box; margin: 0; padding: 0; border-radius: 0; background: transparent; user-select: text; touch-action: auto; }
.canvas-cell-input :deep(.ant-select-input) { height: 100%; }
.canvas-cell-input :deep(.ant-select-content) { height: 100%; min-height: 0; }
.rich-text-editor-active :deep(.canvas-cell-input .ant-select-content),
.rich-text-editor-active :deep(.canvas-cell-input .ant-select-selector) { background: transparent; }
.rich-text-editor-active :deep(.canvas-cell-input input), .rich-text-editor-active :deep(.canvas-cell-input textarea) { -webkit-text-fill-color: transparent; }
.rich-text-editor-active :deep(.canvas-cell-input input::selection), .rich-text-editor-active :deep(.canvas-cell-input textarea::selection) { color: transparent; background: transparent; }
.canvas-cell-fill-handle { position: absolute; z-index: 4; box-sizing: border-box; border: 1px solid var(--sheet-canvas-background); background: var(--sheet-canvas-primary); pointer-events: auto; cursor: crosshair; touch-action: none; }
.sheet-validation-checkbox-error { position: absolute; z-index: 6; left: 12px; right: 12px; bottom: 12px; }
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
