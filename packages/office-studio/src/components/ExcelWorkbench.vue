<script setup lang="ts">
import { computed, defineAsyncComponent, inject, nextTick, onActivated, onDeactivated, onMounted, onUnmounted, provide, ref, shallowRef, toRaw, useId, watch } from 'vue'
import type { CSSProperties } from 'vue'
import { OfficeFontListKey } from '../platform/ribbon/ribbon-fonts'
import { RibbonEditorFocusScopeKey } from '../platform/ribbon/ribbon-popup-scope'
import { WORKBENCH_CONTROL_THEME } from '../platform/workbench-theme.ts'
import { Alert, Button, Checkbox, ConfigProvider, Form, FormItem, Input, InputNumber, Modal, Select, Splitter, SplitterPanel, Tooltip, Upload, theme } from 'antdv-next'
import { FOffice, type FChart } from '../api/facade'
import type { OfficeStudio } from '../api/office-studio'
import type { ApiSheetSelection } from '../api/events.ts'
import { SPARKLINE_PUBLIC_COMMANDS } from '../api/sparkline-public-commands.ts'
import { sparklineApiRangeReference } from '../api/sparklines.ts'
import { sparklinePanelOwnerCurrent, type SparklinePanelOwner } from '../features/sparkline/sparkline-panel-state.ts'
import { MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND } from '../api/sheet-selection-commands.ts'
import { COPY_WORKSHEET_COMMAND, planWorksheetCopy, type CopyWorksheetCommandInput } from '../features/sheet-workbench/sheet-worksheet-copy.ts'
import { WORKSHEET_LIFECYCLE_COMMANDS, planWorksheetLifecycle, type WorksheetLifecycleCommandId, type WorksheetLifecycleCommandInput } from '../features/sheet-workbench/sheet-worksheet-lifecycle.ts'
import { COMMIT_SHEET_CELL_INPUT_COMMAND, planSheetCellInput, sheetCellInputConflictMessage, type SheetCellInputCommandInput,
  type SheetCellInputCommandResult } from '../features/sheet-workbench/sheet-cell-input-command.ts'
import type { ExcelCellView, ExcelWorkbookView as DomainExcelWorkbookView } from '../services/excelViewAdapter'
import { applyExcelWorkbookView, parseCellKey, sheetViewCellAt, toExcelCellView, toExcelWorkbookView } from '../services/excelViewAdapter'
import { createBlankWorkbook } from '../domain/model'
import type { ChartModel } from '../domain/model'
import { useSheetChartPanels } from '../features/charts/views/use-sheet-chart-panels'
import type { SlicerModel } from '../domain/model'
import SpreadsheetCanvas from './spreadsheet/SpreadsheetCanvas.vue'
import ChartOverlay from './ChartOverlay.vue'
import { CHART_OBJECT_COMMANDS, type ChartObjectTarget } from '../features/charts/chart-object-commands'
import type { ChartObjectChange } from '../features/charts/views/use-chart-object-actions'
import type { ChartEditPanelRoute } from '../features/charts/chart-edit-panel-route'
import { chartPositionForRange, chartSourceFromGridRange, type ExcelWorkbookChartView } from './chart-overlay-model'
import type { DrawCommand } from '../rendering/contracts'
import { SpreadsheetGeometry, type GridHit, type GridRange, type SpreadsheetRenderSheet } from '../rendering/spreadsheet/layout'
import { addRange, containsCell, selectionBounds, selectionCount, subtractRange } from '../rendering/spreadsheet/selection'
import { ArrowDown, ArrowUp, Download, FileSpreadsheet, FolderOpen, ListRestart, Replace, Save, Search, X } from 'lucide-vue-next'
import WorkbenchRibbon from '../platform/ribbon/WorkbenchRibbon.vue'
import type { RibbonVariant } from '../platform/ribbon/ribbon-schema'
import { useWorkbenchDisplay, workbenchVisibility } from '../platform/workbench-display'
import SheetFormulaBar from './spreadsheet/SheetFormulaBar.vue'
import SheetFooter from './spreadsheet/SheetFooter.vue'
import { createSheetRibbon, type SlicerRibbonControlState } from './spreadsheet/sheet-ribbon'
import { canConfigurePanelSlicer } from '../features/slicer/slicer-panel-state'
import { slicerPanelOwnerCurrent, type SlicerPanelOwner } from '../features/slicer/slicer-panel-owner'
import { calculateSelectionStatistics, frozenPaneCountsForSelection, normalizeSheetViewState, parseSheetAddress, validateSheetName, type SheetViewState } from '../features/sheet-workbench/sheet-workbench'
import { normalizeSheetSelection, type SheetSelectionState } from '../features/sheet-workbench/sheet-selection.ts'
import { planSheetSelectionShortcut, SheetSelectionDirections, type SheetSelectionShortcutKey } from '../features/sheet-workbench/sheet-selection-navigation.ts'
import { recalculateSheetViewFormulas } from '../features/sheet-workbench/sheet-view-formulas'
import { type SheetBorderOptions } from '../features/sheet-workbench/sheet-toolbar-format'
import { classifySheetNumberFormat, SHEET_NUMBER_FORMATS } from '../features/sheet-workbench/sheet-number-format'
import { applySheetViewNumberFormatCommand, numberFormatActionForCommand, SHEET_NUMBER_FORMAT_COMMANDS, type SheetNumberFormatCommandInput } from '../features/sheet-workbench/sheet-number-format-command.ts'
import { applySheetViewFontFormatCommand, fontFormatActionForCommand, sheetFontFormatCommandForShortcut, SHEET_FONT_FORMAT_COMMAND_IDS, type SheetFontFormatCommandInput } from '../features/sheet-workbench/sheet-font-format-command.ts'
import { sheetStyleSelectionProbes } from '../features/sheet-workbench/sheet-style-selection.ts'
import { applySheetViewBorderFormatCommand, borderFormatActionForCommand, createSheetBorderFormatState, SHEET_BORDER_FORMAT_COMMAND_IDS, SHEET_BORDER_FORMAT_COMMANDS, sheetBorderFormatCommandInput, sheetBorderFormatOptions, type SheetBorderFormatCommandInput, type SheetBorderFormatState } from '../features/sheet-workbench/sheet-border-format-command.ts'
import { rememberSheetColor, sheetBorderColor, sheetOpaqueColorHex, type SheetColorField, type SheetColorMemory } from '../features/sheet-workbench/sheet-colors.ts'
import { sheetCellNumberPattern, sheetCellNumberValue } from '../features/sheet-workbench/sheet-number-presentation.ts'
import { createSheetCellEditorStart, editSheetCellInlineText, editSheetCellText, sheetCellContentKey, sheetCellInputText } from '../features/sheet-workbench/sheet-cell-edit.ts'
import { sheetCellInputActionForKey } from '../features/sheet-workbench/sheet-cell-edit-input.ts'
import { createSheetCellEditSession, getSheetCellEditConflict, resolveSheetDate1904, type SheetCellEditSession } from '../features/sheet-workbench/sheet-cell-edit-session.ts'
import { SheetCellInputEditRecorder, type SheetCellInputEditTrace, type SheetCellNativeInputEvent } from '../features/sheet-workbench/sheet-cell-input-trace.ts'
import { sheetCellFontValues, type SheetRichTextFontField } from '../features/sheet-workbench/sheet-cell-rich-text-style.ts'
import { replaySheetCellInlineFormats, SheetCellInlineFormatDraft, type SheetCellInlineFormatEdit } from '../features/sheet-workbench/sheet-cell-inline-format.ts'
import { FORMAT_SHEET_CELL_INPUT_OPERATION, planSheetCellInlineFontFormat,
  type SheetCellInlineFontFormatInput, type SheetCellInlineFontFormatResult } from '../features/sheet-workbench/sheet-cell-inline-format-command.ts'
import { applySheetViewLayoutFormatCommand, getSheetViewLayoutFormatState, layoutFormatActionForCommand, SHEET_LAYOUT_FORMAT_COMMAND_IDS, type SheetLayoutFormatCommandInput } from '../features/sheet-workbench/sheet-layout-format-command.ts'
import { clearModeForCommand, SHEET_CLEAR_COMMAND_IDS, SHEET_CLEAR_COMMANDS } from '../features/sheet-workbench/sheet-clear-command.ts'
import { SHEET_MERGE_COMMAND_IDS } from '../features/sheet-workbench/sheet-merge-command.ts'
import { useSheetRangeEditor } from './spreadsheet/use-sheet-range-editor.ts'
import SheetMergeConfirmationDialog from './spreadsheet/SheetMergeConfirmationDialog.vue'
import { SHEET_AXIS_DIMENSION_COMMAND_IDS, SHEET_AXIS_DIMENSION_COMMANDS, refreshWorkbookViewAutoRowHeights, sheetAxisDimensionActionForCommand, sheetAxisDimensionCommandInput } from '../features/sheet-workbench/sheet-axis-dimensions.ts'
import { sheetColumnWidth, sheetRowHeight } from '../domain/sheet-axis.ts'
import { useSheetAxisDimensions } from './spreadsheet/use-sheet-axis-dimensions.ts'
import { SHEET_AXIS_STRUCTURE_COMMAND_IDS } from '../features/sheet-workbench/sheet-axis-structure-command.ts'
import { SHEET_AXIS_VISIBILITY_COMMAND_IDS, SheetAxisVisibilityError } from '../features/sheet-workbench/sheet-axis-visibility.ts'
import { useSheetAxisVisibility } from './spreadsheet/use-sheet-axis-visibility.ts'
import { useSheetAxisStructure } from './spreadsheet/use-sheet-axis-structure.ts'
import { useSheetCellShift } from './spreadsheet/use-sheet-cell-shift.ts'
import { SHEET_CELL_SHIFT_COMMAND_IDS } from '../features/sheet-workbench/sheet-cell-shift-command.ts'
import SheetCellShiftConfirmationDialog from './spreadsheet/SheetCellShiftConfirmationDialog.vue'
import { sheetAxisHeaderRange } from '../features/sheet-workbench/sheet-axis-header-gesture.ts'
import SheetAxisSizeDialog from './spreadsheet/SheetAxisSizeDialog.vue'
import SheetRangeThemeDialog from './spreadsheet/SheetRangeThemeDialog.vue'
import { useSheetRangeTheme } from './spreadsheet/use-sheet-range-theme.ts'
import { OPEN_SHEET_RANGE_THEME_OPERATION, SHEET_RANGE_THEME_COMMANDS } from '../features/sheet-workbench/sheet-range-theme-command.ts'
import { useSheetImages } from './spreadsheet/use-sheet-images.ts'
import { useSheetDrawings } from './spreadsheet/use-sheet-drawings.ts'
import { useSheetDrawingSelection } from './spreadsheet/use-sheet-drawing-selection.ts'
import { reconcileWorkbookDrawingStates } from '../features/sheet-drawings/sheet-drawing-model-lifecycle.ts'
import type { SheetDrawingAffineGeometry, SheetDrawingRef } from '../domain/sheet-drawings.ts'
import { createSheetDrawingScene } from '../features/sheet-drawings/sheet-drawing-host.ts'
import { drawingPaintOrder } from '../features/drawing-groups/index.ts'
import { drawingAffineBounds, type DrawingAffineGeometry } from '../domain/drawing-affine.ts'
import { sheetClipboardObjectRange } from '../features/sheet-workbench/sheet-clipboard-object-geometry.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, sheetOutlineHeaderExtents } from '../rendering/spreadsheet/layout.ts'
import { canViewProtectedRanges } from '../features/sheet-protection/range-protection.ts'
import SheetImagePickerDialog from './spreadsheet/SheetImagePickerDialog.vue'
import SheetImageEditorDialog from './spreadsheet/SheetImageEditorDialog.vue'
import SheetImageSaveDialog from './spreadsheet/SheetImageSaveDialog.vue'
import { SHEET_IMAGE_COMMAND_IDS } from '../features/sheet-images/sheet-image-commands.ts'
import { measureSpreadsheetText } from '../rendering/spreadsheet/text-metrics.ts'
import type { SheetNumberFormatPreview } from '../rendering/spreadsheet/scene.ts'
import { buildSheetFunctionFormula, SHEET_FUNCTION_CATALOG } from '../features/sheet-workbench/sheet-function-catalog'
import { findVisibleSheetRow, findVisibleSheetColumn } from '../features/sheet-workbench/sheet-row-visibility'
import { validateSheetDataValidationChanges } from '../features/sheet-validation/validation-engine.ts'
import { planQuickValidation } from '../features/sheet-validation/quick-validation.ts'
import { formatValidationRange } from '../features/sheet-validation/validation-panel-state.ts'
import { resolveSheetClipboardCommand } from '../features/sheet-workbench/sheet-clipboard.ts'
import { useSheetClipboard } from './spreadsheet/use-sheet-clipboard.ts'
import type { SheetClipboardCapabilitiesInput, SheetClipboardPermissionEvaluator } from '../features/sheet-workbench/sheet-clipboard-capabilities.ts'
import { SHEET_FILL_COMMAND_IDS } from '../features/sheet-workbench/sheet-auto-fill.ts'
import { useSheetAutoFill } from './spreadsheet/use-sheet-auto-fill.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewWorksheetProtection } from '../features/sheet-protection/range-protection.ts'
import { SHEET_FORMAT_PAINTER_COMMAND_IDS } from '../features/sheet-workbench/sheet-format-painter-session.ts'
import { useSheetFormatPainter } from './spreadsheet/use-sheet-format-painter.ts'
import { applySheetConditionalFormatsCommand, SET_SHEET_CONDITIONAL_FORMATS_COMMAND, type SetSheetConditionalFormatsInput,
  type SheetConditionalFormatSubmitResult } from '../features/conditional-formatting/conditional-format-command.ts'
import { CHANGE_SHEET_OUTLINE_COMMAND, getSheetOutlineHiddenSpans, getSheetOutlineUpdate, type ChangeSheetOutlineInput, type SheetOutlineAction, type SheetOutlineOperation } from '../features/sheet-outline/sheet-outline.ts'
import { SHEET_OUTLINE_COMMAND_IDS, sheetOutlineCommandInput } from '../features/sheet-outline/sheet-outline-command.ts'
import SheetOutlineActionDialog from '../features/sheet-outline/SheetOutlineActionDialog.vue'
import SheetOutlineSettingsDialog from '../features/sheet-outline/SheetOutlineSettingsDialog.vue'
import { useSheetOutlineWorkbench, sheetOutlineCommandForShortcut, OPEN_SHEET_OUTLINE_SETTINGS_OPERATION } from './spreadsheet/use-sheet-outline-workbench.ts'
import { createOfficeTranslator, createRibbonTranslator, provideWorkbenchRegionalSettings } from '../i18n/office-i18n.ts'
import { createSheetReplacePatches, findSheetMatches, nextSheetMatchIndex, type SheetFindBy, type SheetFindDirection, type SheetFindScope } from '../features/sheet-find-replace/sheet-find-replace.ts'
import { convertTextToNumber } from '../features/sheet-workbench/text-to-number.ts'
import { resolveSlicerBounds } from '../features/slicer/slicer-drawing.ts'
import { slicerControlItemStates } from '../features/slicer/slicer-control-selection.ts'
import { useSlicerCanvasData } from '../features/slicer/use-slicer-canvas-data.ts'
import type { SlicerCommandContext } from '../api/slicer-command-context.ts'
import type { SheetFillCommitContext } from '../features/sheet-workbench/sheet-fill-gesture.ts'
import { canSetSheetHyperlink, normalizeExternalSheetHyperlink, parseSheetHyperlinkTarget, SET_SHEET_HYPERLINK_COMMAND, REMOVE_SHEET_HYPERLINK_COMMAND } from '../features/sheet-workbench/sheet-hyperlink.ts'

defineOptions({ inheritAttrs: false })

// Splitter recreates its DOM inside ResizeObserver; pass sizing through its semantic styles.
const sheetSplitterStyles: { root: CSSProperties; panel: CSSProperties } = {
  root: { flex: '1 1 0', height: 'auto', minWidth: 0, minHeight: 0, overflow: 'hidden', alignItems: 'stretch' },
  panel: { position: 'relative', display: 'flex', height: 'auto', minWidth: 0, minHeight: 0, overflow: 'hidden' },
}

const PivotTablePanel = defineAsyncComponent(() => import('./spreadsheet/PivotTablePanel.vue'))
const PivotChartPanel = defineAsyncComponent(() => import('./spreadsheet/PivotChartPanel.vue'))
const ChartEditorPanel = defineAsyncComponent(() => import('./spreadsheet/ChartEditorPanel.vue'))
const ChartInsertDialog = defineAsyncComponent(() => import('../features/charts/views/ChartInsertDialog.vue'))
const SparklinePanel = defineAsyncComponent(() => import('./spreadsheet/SparklinePanel.vue'))
const SlicerPanel = defineAsyncComponent(() => import('./spreadsheet/SlicerPanel.vue'))
const SheetDataPanel = defineAsyncComponent(() => import('./spreadsheet/SheetDataPanel.vue'))
const SheetDataValidationPanel = defineAsyncComponent(() => import('./spreadsheet/SheetDataValidationPanel.vue'))
const SheetRangeProtectionPanel = defineAsyncComponent(() => import('./spreadsheet/SheetRangeProtectionPanel.vue'))
const SheetProtectionPanel = defineAsyncComponent(() => import('./spreadsheet/SheetProtectionPanel.vue'))
const CustomNumberFormatPanel = defineAsyncComponent(() => import('./spreadsheet/CustomNumberFormatPanel.vue'))
const ConditionalFormattingPanel = defineAsyncComponent(() => import('./spreadsheet/ConditionalFormattingPanel.vue'))

export type ExcelCell = ExcelCellView
export type ExcelSheet = SpreadsheetRenderSheet
export type ExcelWorkbook = Omit<DomainExcelWorkbookView, 'sheets'> & { sheets: ExcelSheet[]; charts?: ChartModel[] }

const props = withDefaults(defineProps<{ modelValue?: ExcelWorkbook; model?: ExcelWorkbook; canUndo?: boolean; canRedo?: boolean; showHeader?: boolean; studio?: OfficeStudio; ribbonVariant?: RibbonVariant;
  hidePasteOptions?: boolean; clipboardPermission?: SheetClipboardPermissionEvaluator;
  clipboardContext?: Pick<SheetClipboardCapabilitiesInput, 'exclusiveRange' | 'shapeTextEditing'> }>(), {
  modelValue: undefined,
  model: undefined,
  canUndo: undefined,
  canRedo: undefined,
  showHeader: false,
  studio: undefined,
  ribbonVariant: 'classic',
  hidePasteOptions: false,
})
const { locale: officeLocale, region: officeRegion, direction: officeDirection } = provideWorkbenchRegionalSettings(() => props.studio?.regionalSettings)

const emit = defineEmits<{
  (event: 'update:modelValue', value: ExcelWorkbook): void
  (event: 'update:model', value: ExcelWorkbook): void
  (event: 'change', value: ExcelWorkbook): void
  (event: 'edit-cell', payload: { sheetId: string; key: string; cell: ExcelCell }): void
  (event: 'import', file: File): void
  (event: 'save', value: ExcelWorkbook): void
  (event: 'export', value: ExcelWorkbook): void
  (event: 'undo', source?: { workbookId: string }): void
  (event: 'redo', source?: { workbookId: string }): void
  (event: 'notify', payload: { type: 'info' | 'success' | 'warning' | 'error'; message: string }): void
}>()

const columnCount = 12

function columnName(index: number) {
  let result = ''
  let value = index
  while (value >= 0) {
    result = String.fromCharCode((value % 26) + 65) + result
    value = Math.floor(value / 26) - 1
  }
  return result
}

function createDefaultWorkbook(): ExcelWorkbook {
  const cells: Record<string, ExcelCell> = {
    A1: { value: '季度销售汇总', bold: true },
    A2: { value: '产品 / 月份', bold: true, background: '#f3f6fa' },
    B2: { value: '一月', bold: true, background: '#f3f6fa', align: 'center' },
    C2: { value: '二月', bold: true, background: '#f3f6fa', align: 'center' },
    D2: { value: '三月', bold: true, background: '#f3f6fa', align: 'center' },
    E2: { value: '第一季度', bold: true, background: '#f3f6fa', align: 'center' },
    A3: { value: '云服务器' },
    B3: { value: '120' },
    C3: { value: '132' },
    D3: { value: '145' },
    E3: { value: '', cachedValue: '397', formula: '=SUM(B3:D3)', bold: true },
    A4: { value: '数据服务' },
    B4: { value: '98' },
    C4: { value: '106' },
    D4: { value: '118' },
    E4: { value: '', cachedValue: '322', formula: '=SUM(B4:D4)', bold: true },
    A5: { value: '安全产品' },
    B5: { value: '76' },
    C5: { value: '89' },
    D5: { value: '94' },
    E5: { value: '', cachedValue: '259', formula: '=SUM(B5:D5)', bold: true },
    A6: { value: '合计', bold: true, background: '#eaf3ff' },
    B6: { value: '', cachedValue: '294', formula: '=SUM(B3:B5)', bold: true, background: '#eaf3ff' },
    C6: { value: '', cachedValue: '327', formula: '=SUM(C3:C5)', bold: true, background: '#eaf3ff' },
    D6: { value: '', cachedValue: '357', formula: '=SUM(D3:D5)', bold: true, background: '#eaf3ff' },
    E6: { value: '', cachedValue: '978', formula: '=SUM(E3:E5)', bold: true, background: '#eaf3ff' },
    A8: { value: '说明', bold: true },
    A9: { value: '数据更新时间：2025-03-31' },
    B9: { value: '单位：万元', italic: true },
  }
  return {
    name: '季度销售汇总.xlsx',
    sheets: [
      { id: 'sheet-1', name: '销售汇总', cells, rowCount: 100, columnCount: columnCount },
      { id: 'sheet-2', name: '产品明细', cells: {}, rowCount: 100, columnCount: columnCount },
      { id: 'sheet-3', name: '数据透视', cells: {}, rowCount: 100, columnCount: columnCount },
    ],
  }
}

const workbook = ref<ExcelWorkbook>(cloneWorkbook(props.modelValue ?? props.model ?? createDefaultWorkbook()))
let autoHeightBaseline = cloneWorkbook(workbook.value)
const userId = computed(() => props.studio?.currentUserId ?? 'local-user')
const availableSheets = computed(() => workbook.value.sheets.filter(sheet => canViewWorksheetProtection(sheet, userId.value)))
const activeSheetId = ref(availableSheets.value.find(sheet => sheet.id === workbook.value.activeSheetId && sheet.visible !== false)?.id ?? availableSheets.value.find(sheet => sheet.visible !== false)?.id ?? '')
const selectedKey = ref('A1')
const selectionAnchor = ref('A1')
let headerSelectionAnchor: { axis: 'row' | 'column'; index: number } | undefined
let headerSelectionBase: GridRange[] = []
let applyingHeaderSelection = false
let applyingApiSelection = false
let syncingUiSelection = false
let applyingShortcutSelection = false
let standaloneSelectionTabStart: SheetSelectionState | undefined
const selectedRanges = ref<GridRange[]>([{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }])
const activeSelectionIndex = ref(-1)
const emptyCellSelection: string[] = []
const editingKey = ref<string | null>(null)
const editValue = ref('')
const editSession = shallowRef<SheetCellEditSession | null>(null)
let editRecorder: SheetCellInputEditRecorder | undefined
let editInlineDraft: SheetCellInlineFormatDraft | undefined
const editFontRevision = ref(0)
const editTextSelection = ref<{ start: number; end: number; host: 'cell' | 'formula' }>({ start: 0, end: 0, host: 'cell' })
const ribbonEditorOwner = useId()
provide(RibbonEditorFocusScopeKey, { ownerId: ribbonEditorOwner })
let editSessionStudio: OfficeStudio | undefined
const cellInputContext = () => ({ sessionId: editSession.value?.sessionId, value: editRecorder?.value ?? editValue.value })
const cellEditorComposing = ref(false), formulaEditorComposing = ref(false)
const editComposing = computed(() => cellEditorComposing.value || formulaEditorComposing.value)
watch(() => props.studio, () => { cancelEdit() }, { flush: 'sync' })
const selecting = ref(false)
const formulaInput = ref<InstanceType<typeof SheetFormulaBar> | null>(null)
const spreadsheetCanvas = ref<InstanceType<typeof SpreadsheetCanvas> | null>(null)
const chartOverlay = ref<InstanceType<typeof ChartOverlay> | null>(null)
const displayPreference = useWorkbenchDisplay()
const chrome = computed(() => workbenchVisibility(displayPreference?.value.chrome ?? 'full'))
const zoom = ref<number>(displayPreference?.value.zoom ?? 100)
const searchText = ref('')
const replaceText = ref('')
const findBy = ref<SheetFindBy>('value')
const findDirection = ref<SheetFindDirection>('row')
const findScope = ref<SheetFindScope>('sheet')
const findCaseSensitive = ref(false)
const findWholeCell = ref(false)
const findSelectionOnly = ref(false)
const findSelectionRanges = ref<GridRange[]>([])
const activeTab = ref('home')
const nameError = ref('')
const searchOpen = ref(false)
const searchInput = ref<{ focus: () => void } | null>(null)
const replaceInput = ref<{ focus: () => void } | null>(null)
const renameSheetId = ref<string | null>(null)
const renameValue = ref('')
const renameError = ref('')
const viewSessions = new Map<string, SheetViewState>()
watch(() => displayPreference?.value.zoom, value => {
  if (value === undefined) return
  for (const [id, session] of viewSessions) viewSessions.set(id, { ...session, zoom: value })
  zoom.value = value
})
const chartCommands = ref<DrawCommand[]>([])
const focusedChartId = ref<string | null>(null)
const hyperlinkDialogOpen = ref(false)
const hyperlinkTarget = ref('')
const hyperlinkDisplay = ref('')
const hyperlinkError = ref('')
const hyperlinkCell = ref<{ unitId: string; sheetId: string; key: string } | null>(null)
const editingHyperlink = computed(() => {
  const target = hyperlinkCell.value
  return target && target.unitId === workbook.value.id ? workbook.value.sheets.find(sheet => sheet.id === target.sheetId)?.cells[target.key]?.hyperlink : undefined
})
const workbenchRef = ref<HTMLElement | null>(null)
const isFullscreen = ref(false)
const pivotPanelOpen = ref(false)
const pivotSource = ref('')
const slicerPanelOpen = ref(false)
const slicerKind = ref<'slicer' | 'timeline'>('slicer')
const slicerPanelSection = ref<'settings' | 'styles' | 'connections'>('settings')
const selectedSlicerId = ref('')
const dataPanelOpen = ref(false)
const dataPanelMode = ref<'filter' | 'sort'>('filter')
const dataPanelRange = ref('A1')
const functionPanelOpen = ref(false)
const functionCategory = ref<string | null>(null)
const functionName = ref('SUM')
const functionArguments = ref('')
const functionError = ref('')
const validationPanelOpen = ref(false)
const validationPanelRange = ref('A1')
const validationPanelRuleId = ref<string>()
const validationPanelInitialType = ref<'checkbox' | 'list' | 'date' | 'time' | undefined>()
const protectionPanelOpen = ref(false)
const worksheetProtectionSheetId = ref<string | null>(null)
const protectionStartNew = ref(false)
const sparklinePanelOpen = ref(false)
const sparklineTargetRange = ref('A1')
interface SparklineWorkbenchContext {
  key: number; studio: OfficeStudio; workbookId: string; sheetId: string; targetRange: string; targetRanges: string[]
  setOpen(open: boolean): void; applied(): void; closed(owner?: SparklinePanelOwner): void
}
let sparklineContextSequence = 0
const sparklineContext = shallowRef<SparklineWorkbenchContext>()
const customNumberFormatOpen = ref(false)
const numberFormatPreview = ref<SheetNumberFormatPreview>()
const sheetSidebarWidth = ref(384)
const workbenchActive = ref(true)
const conditionalFormattingOpen = ref(false)
const conditionalFormattingRange = ref('A1')
const conditionalFormattingContext = shallowRef<{ unitId: string | undefined; sheetId: string; expectedRevision?: number; standaloneSnapshot?: string }>()
const functionEntry = computed(() => SHEET_FUNCTION_CATALOG.find(entry => entry.name === functionName.value))
const functionOptions = computed(() => SHEET_FUNCTION_CATALOG
  .filter(entry => functionCategory.value === null || entry.category === functionCategory.value)
  .map(entry => ({ label: `${entry.name} · ${entry.label}`, value: entry.name })))
const sortBusy = ref(false)
const sortProgress = ref('')
let sortController: AbortController | undefined
const chartWorkbook = computed(() => workbook.value as ExcelWorkbookChartView)
const sheetCharts = computed(() => (workbook.value.charts ?? []).filter(chart => chart.sheetId === activeSheetId.value))
const focusedChart = computed(() => sheetCharts.value.find(chart => chart.id === focusedChartId.value) ?? null)
const activeSheet = computed(() => availableSheets.value.find((sheet) => sheet.id === activeSheetId.value && sheet.visible !== false) ?? availableSheets.value.find(sheet => sheet.visible !== false))
const { context: outlineContext, actionRequest: outlineActionRequest, settingsRequest: outlineSettingsRequest,
  actionOpen: outlineActionOpen, settingsOpen: outlineSettingsOpen, uiOpen: outlineUiOpen, openSettings: openOutlineSettings, run: runOutlineAlias } = useSheetOutlineWorkbench({
  state: () => ({ studio: props.studio, unitId: workbook.value.id ?? 'standalone-workbook', sheetId: activeSheetId.value, sheet: activeSheet.value,
    ranges: selectedRanges.value.map(range => ({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
      ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })), userId: userId.value, active: workbenchActive.value }),
  apply: input => runOutlineCommand(input.action, input),
  refresh: () => { const model = workbook.value.id && props.studio?.getWorkbookModel(workbook.value.id); if (model) { workbook.value = toExcelWorkbookView(model); commitHistory() } },
  focus: () => { if (!editingKey.value) spreadsheetCanvas.value?.focus() },
})
const { chartModalOpen, chartEditorOpen, pivotChartPanelOpen, chartInsertBusy, chartEditorBusy, chartError, chartEditorError,
  chartInsertContext, chartEditorContext, chartEditorRoute, chartEditorRouteKey, pivotChartContext, chartEditorChart, chartSourceRevision, chartEditorSheet, chartEditorSheets, chartInsertSheets,
  openInsertChart, openChartEditor, openPivotChart } = useSheetChartPanels({
  studio: () => props.studio, unitId: () => workbook.value.id, sheetId: () => activeSheet.value?.id,
  sheets: () => availableSheets.value, active: () => workbenchActive.value,
  canRestoreFocus: () => !editingKey.value && !editComposing.value && !rangeThemePending.value && !cellShiftPending.value
    && (!props.studio || props.studio.activeWorkbookId === workbook.value.id && props.studio.editor.state.activeKind === 'workbook'),
  refresh: (studio, unitId) => {
    const model = studio.getWorkbookModel(unitId)
    if (model && props.studio === studio && workbook.value.id === unitId) applyExternalWorkbook(toExcelWorkbookView(model))
  },
  focus: () => spreadsheetCanvas.value?.focus(), notify: message => emit('notify', { type: 'error', message }),
})
const activeDate1904 = computed(() => resolveSheetDate1904(activeSheet.value, workbook.value))
const activeRenderSheet = computed(() => activeSheet.value
  ? { ...activeSheet.value, date1904: activeDate1904.value } : undefined)
const sheetViewportHeaders = computed(() => sheetOutlineHeaderExtents(activeRenderSheet.value?.outline))
const slicerCanvasData = useSlicerCanvasData({ studio: () => props.studio, unitId: () => workbook.value.id,
  sheetId: () => activeSheetId.value, sheet: () => activeRenderSheet.value, view: () => workbook.value,
  active: () => workbenchActive.value && (!props.studio || props.studio.activeWorkbookId === workbook.value.id && props.studio.editor.state.activeKind === 'workbook') })
const slicerControls = computed(() => slicerCanvasData.controls.value.map(control => ({ ...control, styleOwner: workbook.value })))
const translate = computed(() => createOfficeTranslator(officeLocale.value))
function t(key: string, fallback: string) { return translate.value(`sheet.${key}`, fallback) }
const activeCell = computed(() => {
  const sheet = activeSheet.value
  const position = sheet ? parseCellKey(sheetCellContentKey(sheet, selectedKey.value)) : undefined
  return sheet && position ? sheetViewCellAt(sheet, position.row, position.col) : { value: '' }
})
const numberFormatSource = computed(() => {
  const range = selectedRanges.value[activeSelectionIndex.value] ?? selectedRanges.value.at(-1)
  const key = range ? `${columnName(range.startColumn)}${range.startRow + 1}` : selectedKey.value
  const position = parseCellKey(key), sheet = activeSheet.value
  const cell = sheet && position ? sheetViewCellAt(sheet, position.row, position.col) : undefined
  return { key, pattern: sheetCellNumberPattern(cell), sample: sheetCellNumberValue(cell) }
})
const numberFormatSelectionKey = computed(() => JSON.stringify([workbook.value.id, activeSheet.value?.id, selectedRanges.value]))
const sheetSidebarOpen = computed(() => Boolean(activeSheet.value && (customNumberFormatOpen.value
  || validationPanelOpen.value && props.studio && workbook.value.id)))
const selectedAddress = computed(() => {
  const range = selectedRanges.value[0]
  return selectedRanges.value.length === 1 && range && (range.startRow !== range.endRow || range.startColumn !== range.endColumn) ? rangeNotation(range) : selectedKey.value
})
const slicerSourceRange = computed(() => {
  if (selectedRanges.value.length !== 1) return undefined
  const range = selectedRanges.value[0]
  return range ? {
    startRow: range.startRow,
    endRow: range.endRow,
    startCol: range.startColumn,
    endCol: range.endColumn,
  } : undefined
})
const selectedCount = computed(() => selectionCount(selectedRanges.value))
const sheetCells = computed(() => activeSheet.value?.cells ?? {})
const totalRows = computed(() => activeSheet.value?.rowCount ?? 100)
const freezeRows = computed(() => activeSheet.value?.freeze?.rows ?? 0)
const freezeColumnsCount = computed(() => activeSheet.value?.freeze?.columns ?? 0)
const isFrozen = computed(() => freezeRows.value > 0 || freezeColumnsCount.value > 0)
const canUndo = computed(() => {
  editFontRevision.value
  return editingKey.value && editRecorder ? !editComposing.value && (editRecorder.canUndoDraft || editInlineDraft?.canUndo(editRecorder.trace) === true) : props.canUndo ?? false
})
const canRedo = computed(() => {
  editFontRevision.value
  return editingKey.value && editRecorder ? !editComposing.value && (editRecorder.canRedoDraft || editInlineDraft?.canRedo(editRecorder.trace) === true) : props.canRedo ?? false
})
const selectedStatistics = computed(() => calculateSelectionStatistics(activeSheet.value ?? { id: '', name: '', cells: {}, rowCount: 0, columnCount: 0 }, selectedRanges.value))
const inlineDraftReplay = computed(() => {
  editFontRevision.value
  const session = editSession.value, recorder = editRecorder, draft = editInlineDraft
  if (!session || !recorder || !draft || editingKey.value !== session.key || session.sheetId !== activeSheetId.value) return undefined
  const trace = recorder.trace
  return replaySheetCellInlineFormats(getCell(session.key), trace, draft.formats(trace),
    { numberFormat: session.source.numberFormat, overwrite: session.mode === 'overwrite' })
})
const inlineFontEditable = computed(() => {
  const session = editSession.value
  if (!session || !inlineDraftReplay.value || !activeSheet.value) return false
  const cell = editSheetCellInlineText(getCell(session.key), editValue.value, session.date1904, session.region, session.currentYear)
  return !cell.formula && (typeof cell.rawValue === 'string' || cell.rawValue === null)
})
const richTextEditorCell = computed<ExcelCellView | undefined>(() => {
  const replay = inlineFontEditable.value ? inlineDraftReplay.value : undefined
  if (!replay?.cell.richText?.length) return undefined
  const cell = replay.cell, raw = String(cell.rawValue ?? '')
  const prefix = editValue.value.length - raw.length
  if (prefix !== 0 && !(prefix === 1 && editValue.value === "'" + raw)) return undefined
  // The editor escape is a view-only run; it never becomes a stored run or acquires a source ID.
  return { ...cell, value: editValue.value, rawValue: editValue.value,
    richText: prefix ? [{ text: "'" }, ...cell.richText] : cell.richText }
})
const selectionFormat = computed(() => {
  const probes = activeSheet.value ? sheetStyleSelectionProbes(activeSheet.value, selectedRanges.value.map(range => ({
    startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
  }))) : { cells: [], complete: true }
  function common<K extends keyof ExcelCell>(property: K, fallback: ExcelCell[K]): ExcelCell[K] | undefined {
    if (!probes.complete) return undefined
    const values = probes.cells.map(cell => cell[property] ?? fallback)
    if (!values.length) values.push(fallback)
    return values.every(value => value === values[0]) ? values[0] : undefined
  }
  function commonFont<K extends SheetRichTextFontField>(field: K, fallback: import('../domain/model.ts').FontStyle[K]) {
    const replay = inlineFontEditable.value ? inlineDraftReplay.value : undefined
    if (!replay && !probes.complete) return undefined
    const prefix = replay ? editValue.value.length - String(replay.cell.rawValue ?? '').length : 0
    const selection = { start: Math.max(0, editTextSelection.value.start - prefix), end: Math.max(0, editTextSelection.value.end - prefix) }
    const typing = replay?.typing && selection.start === selection.end && replay.typing.position === selection.start ? replay.typing.font[field] : undefined
    const values = replay ? typing === undefined ? sheetCellFontValues(replay.cell, field, selection) : [typing]
      : probes.cells.flatMap(cell => sheetCellFontValues(cell, field))
    if (!values.length) values.push(fallback)
    return values.every(value => value === values[0]) ? values[0] : undefined
  }
  const formats = probes.cells.map(cell => cell.rawNumberFormat ?? SHEET_NUMBER_FORMATS[cell.numberFormat ?? 'general'])
  if (!formats.length) formats.push('General')
  const numberFormatPattern = probes.complete && formats.every(format => format === formats[0]) ? formats[0]! : null
  const numberFormatKind = numberFormatPattern === null ? 'custom' : classifySheetNumberFormat(numberFormatPattern)
  const borderColors = probes.cells.map(cell => sheetBorderColor(cell.border))
  if (!borderColors.length) borderColors.push(null)
  const borderColor = probes.complete && borderColors.every(color => color === borderColors[0]) ? borderColors[0] : undefined
  return {
    fontFamily: commonFont('family', 'Aptos') ?? null,
    fontSize: commonFont('size', 11) ?? null,
    numberFormat: numberFormatKind === 'custom' ? null : numberFormatKind,
    numberFormatPattern,
    bold: commonFont('bold', false), italic: commonFont('italic', false), underline: commonFont('underline', false), strike: commonFont('strike', false),
    ...getSheetViewLayoutFormatState(activeSheet.value, selectedModelRanges()),
    color: commonFont('color', null), background: common('background', null),
    borderColor,
  }
})
const validationOptions = computed(() => {
  const workbookId = workbook.value.id
  const sheetId = activeSheet.value?.id
  if (!props.studio || !workbookId || !sheetId) return []
  const [column, row] = splitKey(editingKey.value ?? selectedKey.value)
  return FOffice.newAPI(props.studio).getWorkbook(workbookId)?.getSheetBySheetId(sheetId)?.getCellValidationOptions(row - 1, column) ?? []
})
const { canMerge, canUnmerge, mergeRequest, mergeError, runClear, runMerge, confirmMerge, cancelMerge, afterMergeClose } = useSheetRangeEditor({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  ranges: selectedModelRanges, userId: () => userId.value, region: () => officeRegion.value,
  active: () => workbenchActive.value, applyView: value => { workbook.value = value },
  changed: recalculate => { if (recalculate) recalculateFormulas(); commitHistory() },
  focus: () => { if (!editingKey.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() },
})
const { pending: axisSizePending, axisSizeError, openAxisSizeDialog, runAxisDimensionCommand,
  confirmAxisSize, cancelAxisSize, afterAxisSizeClose } = useSheetAxisDimensions({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  ranges: selectedModelRanges, userId: () => userId.value, region: () => officeRegion.value,
  active: () => workbenchActive.value, finishEdit,
  prepareView: prepareAxisWorkbookView, applyView: applyExternalWorkbook,
  changed: options => { clearPainterSession(); commitHistory({ ...options, prepared: true }) },
  focus: () => { if (!editingKey.value && !mergeRequest.value && !painterRequest.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() },
})
const axisSizeRequest = computed(() => axisSizePending.value?.request ?? null)
const { pending: rangeThemePending, request: rangeThemeRequest, rangeThemeError, openRangeThemeDialog, runRangeThemeCommand,
  confirmRangeTheme, cancelRangeTheme, afterRangeThemeClose } = useSheetRangeTheme({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  ranges: selectedModelRanges, userId: () => userId.value, active: () => workbenchActive.value, finishEdit,
  prepareView: (view, previous) => {
    refreshWorkbookViewAutoRowHeights(view, previous, { measure: props.studio?.sheetTextMeasure ?? measureSpreadsheetText, region: officeRegion.value })
    return view
  },
  applyView: applyExternalWorkbook,
  changed: options => {
    if (options?.current && !options.current()) return
    clearPainterSession()
    if (!options?.current || options.current()) commitHistory({ ...options, prepared: true })
  },
  focus: () => { if (!editingKey.value && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() },
})
const { runAxisStructureCommand } = useSheetAxisStructure({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  ranges: selectedModelRanges, userId: () => userId.value, region: () => officeRegion.value,
  active: () => workbenchActive.value, finishEdit,
  prepareView: prepareAxisWorkbookView,
  applyView: applyExternalWorkbook,
  selection: range => {
    selectedRanges.value = [{ startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol }]
    selectedKey.value = `${columnName(range.startCol)}${range.startRow + 1}`
    selectionAnchor.value = selectedKey.value
    syncApiSelection()
  },
  changed: options => { clearPainterSession(); commitHistory({ ...options, prepared: true }) },
})
const { pending: cellShiftPending, request: cellShiftRequest, cellShiftError, runCellShiftCommand,
  confirmCellShift, cancelCellShift, afterCellShiftClose } = useSheetCellShift({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  ranges: selectedModelRanges, activeRangeIndex: () => activeSelectionIndex.value,
  userId: () => userId.value, region: () => officeRegion.value,
  active: () => workbenchActive.value, finishEdit,
  prepareView: prepareAxisWorkbookView, applyView: applyExternalWorkbook,
  selection: range => {
    selectedRanges.value = [{ startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol }]
    selectedKey.value = `${columnName(range.startCol)}${range.startRow + 1}`
    selectionAnchor.value = selectedKey.value
    activeSelectionIndex.value = 0
    syncApiSelection()
  },
  changed: options => {
    if (options?.current && !options.current()) return
    clearPainterSession()
    if (!options?.current || options.current()) commitHistory({ ...options, prepared: true })
  },
  focus: () => {
    if (!editingKey.value && !editComposing.value && !mergeRequest.value && !painterRequest.value
      && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus()
  },
})
const { pending: imagePickerPending, editor: imageEditorPending, saving: imageSavePending, pickerRequest: imagePickerRequest,
  editorRequest: imageEditorRequest, saveRequest: imageSaveRequest, selectedImageId, selectedImageIds, imageError, busy: imageBusy, progress: imageProgress,
  runImageCommand, confirmPicker: confirmImagePicker, cancelPicker: cancelImagePicker, openImageEditor, confirmEditor: confirmImageEditor,
  cancelEditor: cancelImageEditor, confirmSave: confirmImageSave, cancelSave: cancelImageSave, afterClose: afterImageClose, saveImage } = useSheetImages({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value, selectedKey: () => selectedKey.value,
  ranges: selectedModelRanges, userId: () => userId.value, active: () => workbenchActive.value,
  finishEdit: () => !editComposing.value && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && finishEdit(),
  prepareView: prepareAxisWorkbookView, applyView: applyExternalWorkbook,
  changed: options => { if (options?.current && !options.current()) return; clearPainterSession(); if (!options?.current || options.current()) commitHistory({ ...options, prepared: true }) },
  notify: message => emit('notify', { type: 'error', message }),
  focus: () => { if (!editingKey.value && !editComposing.value && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() },
  crop: (id, input) => spreadsheetCanvas.value?.operateImageCrop(id, input) ?? { handled: false },
})
const imageUiOpen = computed(() => Boolean(imagePickerPending.value || imageEditorPending.value || imageSavePending.value))
const drawingWorkbook = computed(() => applyExcelWorkbookView(createBlankWorkbook(workbook.value.name), workbook.value))
function drawingPermissionRange(frame: DrawingAffineGeometry) {
  const bounds = drawingAffineBounds(frame)
  return sheetClipboardObjectRange(drawingWorkbook.value, activeSheetId.value,
    { ...bounds, left: bounds.left + GRID_HEADER_WIDTH, top: bounds.top + GRID_HEADER_HEIGHT })
}
function canViewDrawing(ref: SheetDrawingRef): boolean {
  const object = drawingScene.value?.objects.find(object => object.ref.kind === ref.kind && object.ref.id === ref.id), sheet = activeSheet.value
  return !!object && object.available !== false && !!sheet && canViewWorksheetProtection(sheet, userId.value)
    && canViewProtectedRanges(sheet, [drawingPermissionRange(object.geometry)], userId.value)
}
const drawingSelection = useSheetDrawingSelection({
  studio: () => props.studio, unitId: () => workbook.value.id, sheetId: () => activeSheetId.value, active: () => workbenchActive.value,
  scene: () => activeSheet.value ? createSheetDrawingScene(drawingWorkbook.value, activeSheetId.value, {
    canEdit: (_ref, frame) => canEditWorksheetProtection(activeSheet.value!, userId.value)
      && canViewProtectedRanges(activeSheet.value!, [drawingPermissionRange(frame)], userId.value)
      && canEditProtectedRanges(activeSheet.value!, [drawingPermissionRange(frame)], userId.value),
  }) : undefined,
  nativeSelection: () => [...selectedImageIds.value.map(id => ({ kind: 'image' as const, id })),
    ...(focusedChartId.value ? [{ kind: 'chart' as const, id: focusedChartId.value }] : []),
    ...(selectedSlicerId.value ? [{ kind: workbook.value.slicers?.find(slicer => slicer.id === selectedSlicerId.value)?.type ?? 'slicer', id: selectedSlicerId.value }] : [])],
  applyNative: refs => {
    selectedImageIds.value = refs.every(ref => ref.kind === 'image') ? refs.map(ref => ref.id) : []
    focusedChartId.value = refs.length === 1 && refs[0]?.kind === 'chart' ? refs[0].id : null
    selectedSlicerId.value = refs.length === 1 && (refs[0]?.kind === 'slicer' || refs[0]?.kind === 'timeline') ? refs[0].id : ''
  },
})
const selectedDrawingRefs = drawingSelection.refs, drawingScene = drawingSelection.scene
const drawingOrder = computed(() => drawingScene.value ? drawingPaintOrder(drawingScene.value) : undefined)
function selectDrawing(ref?: SheetDrawingRef, additive = false, preserve = false): void { drawingSelection.select(ref, additive, preserve) }
function routeChartDrawingEvent(event: MouseEvent, chartId: string): boolean { return spreadsheetCanvas.value?.routeDrawingEvent(event, chartId) ?? false }
function routeCanvasChartEvent(event: MouseEvent | KeyboardEvent, chartId: string, pane: { scrollLeft: number; scrollTop: number; headerWidth?: number; headerHeight?: number }): void {
  if (event instanceof KeyboardEvent) chartOverlay.value?.routeCanvasKeydown(event, chartId)
  else chartOverlay.value?.routeCanvasEvent(event, chartId, pane)
}
const { runDrawingCommand, error: drawingError } = useSheetDrawings({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  selection: () => selectedDrawingRefs.value, select: refs => drawingSelection.set(refs), userId: () => userId.value,
  active: () => workbenchActive.value, finishEdit,
  blocked: () => Boolean(editComposing.value || imageUiOpen.value || rangeThemePending.value || cellShiftPending.value || axisSizePending.value || mergeRequest.value || painterRequest.value),
  prepareView: prepareAxisWorkbookView, applyView: applyExternalWorkbook,
  changed: options => { if (options?.current && !options.current()) return; clearPainterSession(); if (!options?.current || options.current()) commitHistory({ ...options, prepared: true }) },
  notify: message => emit('notify', { type: 'error', message }), focus: () => spreadsheetCanvas.value?.focus(),
})
function selectImage(id?: string, additive = false, preserve = false): void {
  selectDrawing(id ? { kind: 'image', id } : undefined, additive, preserve)
}
function selectSlicer(id: string): void { selectDrawing(id ? { kind: workbook.value.slicers?.find(slicer => slicer.id === id)?.type ?? 'slicer', id } : undefined) }
function selectChart(id: string | null): void { selectDrawing(id ? { kind: 'chart', id } : undefined) }
function prepareAxisWorkbookView(view: DomainExcelWorkbookView, previous: DomainExcelWorkbookView): DomainExcelWorkbookView {
  const next = view.calculationMode === 'manual' ? view
    : recalculateSheetViewFormulas(view, undefined, { numberFormatLocale: officeRegion.value }).workbook
  refreshWorkbookViewAutoRowHeights(next, previous, { measure: props.studio?.sheetTextMeasure ?? measureSpreadsheetText, region: officeRegion.value })
  const baseline = applyExcelWorkbookView(createBlankWorkbook(previous.name), previous),
    candidate = applyExcelWorkbookView(createBlankWorkbook(next.name), next)
  reconcileWorkbookDrawingStates(baseline, candidate)
  return toExcelWorkbookView(candidate)
}
const { runAxisVisibilityCommand } = useSheetAxisVisibility({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  ranges: selectedModelRanges, userId: () => userId.value, region: () => officeRegion.value,
  active: () => workbenchActive.value, finishEdit, prepareView: prepareAxisWorkbookView,
  applyView: applyExternalWorkbook,
  selection: (ranges, primary) => {
    const key = primary ? `${columnName(primary.column)}${primary.row + 1}` : selectedKey.value
    const next = normalizeSheetViewState(activeSheet.value ?? { rowCount: 0, columnCount: 0 }, { key, anchor: key,
      ranges: ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol })), zoom: zoom.value })
    selectedRanges.value = next.ranges; selectedKey.value = next.key; selectionAnchor.value = next.anchor
    syncApiSelection()
  },
  changed: options => { clearPainterSession(); commitHistory({ ...options, prepared: true }) },
})
const { capabilities: clipboardCapabilities, copyRange: clipboardCopyRange, pasteOptions: clipboardPasteOptions,
  pasteOptionsCapabilities: clipboardPasteOptionsCapabilities, runClipboardCommand,
  handleCopy: handleSheetCopy, handlePaste: handleSheetPaste, handleClipboardKeydown } = useSheetClipboard({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  range: () => {
    const range = selectedRanges.value[activeSelectionIndex.value] ?? selectedRanges.value.at(-1)
    return range ? { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn } : undefined
  },
  userId: () => userId.value, region: () => officeRegion.value,
  capabilitySelection: () => ({ ranges: selectedModelRanges(), activeRangeIndex: activeSelectionIndex.value }),
  selectedImageIds: () => [...selectedImageIds.value],
  selectImageIds: ids => { selectedImageIds.value = [...ids]; if (ids.length) { focusedChartId.value = null; selectedSlicerId.value = '' } },
  pasteImageFiles: (files, input) => runImageCommand('sheet.command.insert-float-image', { ...input, files, row: input.range?.startRow,
    column: input.range && ('startCol' in input.range ? input.range.startCol : input.range.startColumn) }),
  hidePasteOptions: () => props.hidePasteOptions,
  exclusiveRange: () => props.clipboardContext?.exclusiveRange ?? false,
  shapeTextEditing: () => props.clipboardContext?.shapeTextEditing ?? false,
  get permission() { return props.clipboardPermission },
  active: () => workbenchActive.value, editing: () => Boolean(editingKey.value), finishEdit,
  prepareView: prepareAxisWorkbookView, applyView: applyExternalWorkbook,
  selection: ranges => {
    const range = ranges.at(-1)
    if (!range) return
    selectedRanges.value = ranges.map(item => ({ startRow: item.startRow, endRow: item.endRow, startColumn: item.startCol, endColumn: item.endCol }))
    selectedKey.value = `${columnName(range.startCol)}${range.startRow + 1}`
    selectionAnchor.value = selectedKey.value
    syncApiSelection()
  },
  changed: options => {
    if (!options.current()) return
    clearPainterSession()
    if (options.current()) commitHistory({ ...options, prepared: true })
  },
  focus: () => { if (!editingKey.value && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() },
  notify: value => emit('notify', value),
})
const { fillOptions: sheetFillOptions, fillSelection, refill: refillSelection, runFillCommand, handleFillKeydown } = useSheetAutoFill({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  range: () => selectedModelRanges()[activeSelectionIndex.value] ?? selectedModelRanges().at(-1),
  userId: () => userId.value, region: () => officeRegion.value,
  active: () => workbenchActive.value, editing: () => Boolean(editingKey.value), finishEdit,
  prepareView: prepareAxisWorkbookView, applyView: applyExternalWorkbook,
  selection: range => {
    selectedRanges.value = [{ startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol }]
    selectedKey.value = `${columnName(range.startCol)}${range.startRow + 1}`
    selectionAnchor.value = selectedKey.value
    activeSelectionIndex.value = 0
    syncApiSelection()
  },
  changed: options => {
    if (!options.current()) return
    clearPainterSession()
    if (options.current()) commitHistory({ ...options, prepared: true })
  },
  focus: () => { if (!editingKey.value && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() },
  notify: value => emit('notify', value),
})
const canEditValidationSelection = computed(() => Boolean(props.studio && workbook.value.id && activeSheet.value && selectedRanges.value.length
  && canEditWorksheetProtection(activeSheet.value, userId.value) && canEditProtectedRanges(activeSheet.value, selectedModelRanges(), userId.value)))
const borderFormatState = ref<SheetBorderFormatState>(createSheetBorderFormatState())
const borderOptions = computed(() => sheetBorderFormatOptions(borderFormatState.value))
const borderPreset = computed(() => borderFormatState.value.type)
const standaloneBorderStates = ref(new Map<string, SheetBorderFormatState>())
const colorMemory = ref(new Map<string, SheetColorMemory>())
const colorMemoryId = computed(() => workbook.value.id ?? 'standalone:' + workbook.value.name)
const activeColorMemory = computed(() => colorMemory.value.get(colorMemoryId.value) ?? {})
const { token: sheetThemeToken } = theme.useToken()
function rememberColor(field: SheetColorField, value: unknown, unitId = colorMemoryId.value): void {
  const previous = colorMemory.value.get(unitId) ?? {}
  const next = rememberSheetColor(previous, field, value)
  if (next !== previous) colorMemory.value.set(unitId, next)
}
watch(() => props.studio, (studio, _previous, onCleanup) => {
  if (!studio) return
  const listener = studio.events.addEvent('CommandExecuted', event => {
    if (!SHEET_FONT_FORMAT_COMMAND_IDS.includes(event.id)) return
    const action = fontFormatActionForCommand(event.id, event.params)
    const unitId = event.unitId ?? studio.activeWorkbookId
    if (action?.kind === 'color' && action.value !== null && studio.getWorkbookModel(unitId)) rememberColor(action.field, action.value, unitId)
  })
  const disposedListener = studio.events.addEvent('WorkbookDisposed', event => { colorMemory.value.delete(event.workbookId) })
  onCleanup(() => { listener.dispose(); disposedListener.dispose() })
}, { immediate: true, flush: 'sync' })
function syncBorderFormatState(): void {
  const sheet = activeSheet.value, unitId = workbook.value.id, studio = props.studio
  if (studio && !studio.isDisposed && unitId && sheet && studio.getWorkbookModel(unitId)?.sheets.some(modelSheet => modelSheet.id === sheet.id)
    && canViewWorksheetProtection(sheet, userId.value)) {
    borderFormatState.value = studio.getSheetBorderFormatState(unitId, sheet.id)
  } else borderFormatState.value = { ...(standaloneBorderStates.value.get(`${colorMemoryId.value}\0${sheet?.id ?? ''}`) ?? createSheetBorderFormatState()) }
}
watch([() => props.studio, colorMemoryId, activeSheetId], syncBorderFormatState, { immediate: true, flush: 'sync' })
watch(() => props.studio, (studio, _previous, onCleanup) => {
  if (!studio) return
  const listener = studio.events.addEvent('SheetBorderFormatStateChanged', event => {
    if (event.workbookId === workbook.value.id && event.sheetId === activeSheetId.value) borderFormatState.value = { ...event.state }
  })
  const disposed = studio.events.addEvent('WorkbookDisposed', event => {
    for (const key of standaloneBorderStates.value.keys()) if (key.startsWith(event.workbookId + '\0')) standaloneBorderStates.value.delete(key)
  })
  onCleanup(() => { listener.dispose(); disposed.dispose() })
}, { immediate: true, flush: 'sync' })
const officeFonts = inject(OfficeFontListKey, undefined)
const canFormatSelection = computed(() => !!activeSheet.value && !!selectedRanges.value.length
  && canEditWorksheetProtection(activeSheet.value, userId.value) && canEditProtectedRanges(activeSheet.value, selectedModelRanges(), userId.value))
const { painterActive, painterUnavailable, formatPainterCanvas, painterRequest, painterError,
  runFormatPainterCommand, cancelFormatPainter, clearPainterSession, finishFormatPainterSelection,
  confirmPainter, cancelPainterConfirmation, afterPainterClose } = useSheetFormatPainter({
  studio: () => props.studio, workbook: () => workbook.value, sheetId: () => activeSheetId.value,
  ranges: selectedModelRanges, userId: () => userId.value, region: () => officeRegion.value,
  active: () => workbenchActive.value,
  canStart: () => {
    const sheet = activeSheet.value, source = selectedModelRanges().at(-1)
    return !!sheet && !!source && canEditWorksheetProtection(sheet, userId.value) && canEditProtectedRanges(sheet, [source], userId.value)
  },
  hasObjectSelection: () => Boolean(selectedImageId.value || focusedChart.value || slicerControls.value.some(({ model }) => model.id === selectedSlicerId.value)),
  applyView: value => { workbook.value = value },
  selection: ranges => {
    const next = ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol }))
    selectedRanges.value = next
    const anchor = activeSheet.value ? parseSheetAddress(selectedKey.value, activeSheet.value)?.anchor : undefined
    const [column, row] = splitKey(anchor ?? '')
    if (!anchor || !containsCell(next, row - 1, column)) {
      const last = next.at(-1)
      if (last) selectedKey.value = `${columnName(last.startColumn)}${last.startRow + 1}`
    }
    selectionAnchor.value = selectedKey.value
    syncApiSelection()
  },
  changed: recalculate => { if (recalculate) recalculateFormulas(); commitHistory() },
  focus: () => { if (!editingKey.value && !mergeRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() },
  notify: message => emit('notify', { type: 'error', message }),
})
const slicerRibbonControl = computed<SlicerRibbonControlState | undefined>(() => {
  const refs = selectedDrawingRefs.value, studio = props.studio, unitId = workbook.value.id, sheetId = activeSheetId.value
  const ref = refs.length === 1 ? refs[0] : undefined
  if (!ref || (ref.kind !== 'slicer' && ref.kind !== 'timeline') || !studio || !unitId || !workbenchActive.value) return undefined
  const control = workbook.value.slicers?.find(control => control.id === ref.id && control.type === ref.kind && control.hostSheetId === sheetId)
  if (!control) return undefined
  return { id: control.id, type: control.type, style: control.style, styleOwner: workbook.value, supportsConnections: !control.tableId,
    revision: studio.getWorkbookRevision(unitId), canConfigure: canConfigurePanelSlicer(studio, { workbookId: unitId, sheetId }, control.id) }
})
let previousSlicerRibbonTab = 'home'
watch(() => slicerRibbonControl.value?.id, (id, previous) => {
  if (id && !previous) { if (activeTab.value !== 'slicer-format') previousSlicerRibbonTab = activeTab.value; activeTab.value = 'slicer-format' }
  if (!id && activeTab.value === 'slicer-format') activeTab.value = previousSlicerRibbonTab
})
watch(activeTab, id => { if (id !== 'slicer-format') previousSlicerRibbonTab = id })
const ribbonTabs = computed(() => createSheetRibbon({
  clipboardCapabilities: clipboardCapabilities.value,
  region: officeRegion.value,
  fontList: officeFonts?.value, fontEditable: canFormatSelection.value && !editComposing.value,
  rangeThemeEditable: canFormatSelection.value && !editComposing.value && !rangeThemePending.value && !cellShiftPending.value,
  imageEditable: canFormatSelection.value && !editComposing.value && !imageUiOpen.value,
  backgroundImageEditable: Boolean(activeSheet.value && canEditWorksheetProtection(activeSheet.value, userId.value) && !editComposing.value && !imageUiOpen.value),
  hasBackgroundImage: Boolean(activeSheet.value?.backgroundImage),
  painterActive: painterActive.value, painterUnavailable: painterUnavailable.value,
  borderOptions: borderOptions.value, borderPreset: borderPreset.value, onBorderOptions: changeBorderOptions,
  lastTextColor: activeColorMemory.value.color, lastBackgroundColor: activeColorMemory.value.background,
  defaultTextColor: sheetOpaqueColorHex(sheetThemeToken.value.colorText), defaultBackgroundColor: sheetOpaqueColorHex(sheetThemeToken.value.colorPrimary),
  ...selectionFormat.value, canUndo: canUndo.value, canRedo: canRedo.value, hasSheet: Boolean(activeSheet.value),
  canMerge: canMerge.value, canUnmerge: canUnmerge.value,
  canDeleteRow: totalRows.value > 1, canDeleteColumn: (activeSheet.value?.columnCount ?? 0) > 1,
  chartAvailable: Boolean(props.studio && workbook.value.id), gridlines: activeSheet.value?.showGridlines !== false,
  pivotChartAvailable: Boolean(props.studio && workbook.value.id && activeSheet.value),
  dataValidationAvailable: Boolean(props.studio && workbook.value.id),
  validationPanelOpen: validationPanelOpen.value,
  quickValidationEditable: canEditValidationSelection.value,
  hyperlinkAvailable: canOpenHyperlinkDialog(),
  canSort: Boolean(!sortBusy.value && props.studio && workbook.value.id && selectedRanges.value.length === 1 && selectedRanges.value[0] && selectedRanges.value[0].endRow > selectedRanges.value[0].startRow),
  filterEnabled: Boolean(activeSheet.value?.autoFilter),
  filterCanToggle: Boolean(props.studio && workbook.value.id && (activeSheet.value?.autoFilter || (!sortBusy.value && selectedRanges.value.length === 1 && selectedRanges.value[0] && selectedRanges.value[0].endRow > selectedRanges.value[0].startRow))),
  frozen: isFrozen.value, zoom: zoom.value, fullscreen: isFullscreen.value,
  calculationMode: workbook.value.calculationMode ?? 'automatic',
  wrapStrategy: selectionFormat.value.wrapStrategy ?? null,
  outlineSummaryAbove: activeSheet.value?.outline?.summaryAbove ?? false,
  outlineSummaryLeft: activeSheet.value?.outline?.summaryLeft ?? false,
  outlineContext: outlineContext.value,
  outlineSettingsAvailable: Boolean(props.studio && workbook.value.id),
  slicerControl: slicerRibbonControl.value,
}, createRibbonTranslator(officeLocale.value, 'sheet')))
const formulaValue = computed(() => editingKey.value ? editValue.value : sheetCellInputText(activeCell.value, activeDate1904.value, officeRegion.value))
const searchMatches = computed(() => findSheetMatches(availableSheets.value, activeSheetId.value, {
  text: searchText.value, replacement: replaceText.value, findBy: findBy.value, direction: findDirection.value,
  scope: findScope.value, caseSensitive: findCaseSensitive.value, wholeCell: findWholeCell.value,
}, findScope.value === 'sheet' && findSelectionOnly.value ? findSelectionRanges.value : [], findScope.value === 'sheet' && findSelectionOnly.value))

watch(
  () => props.modelValue,
  (value) => {
    if (value) applyExternalWorkbook(value)
  },
)

watch(
  () => props.model,
  (value) => {
    if (!value || props.modelValue) return
    applyExternalWorkbook(value)
  },
)

function cloneWorkbook(value: ExcelWorkbook): ExcelWorkbook {
  return JSON.parse(JSON.stringify(value)) as ExcelWorkbook
}

function applyExternalWorkbook(value: ExcelWorkbook) {
  if (JSON.stringify(workbook.value) === JSON.stringify(value)) return
  const previousSheetId = activeSheetId.value
  const sameDocument = value.id !== undefined || workbook.value.id !== undefined ? value.id === workbook.value.id : value.sheets.some(sheet => sheet.id === previousSheetId)
  const viewableSheets = value.sheets.filter(sheet => canViewWorksheetProtection(sheet, userId.value))
  const requested = viewableSheets.find(sheet => sheet.id === value.activeSheetId && sheet.visible !== false)
  const nextSheetId = requested?.id ?? viewableSheets.find(sheet => sheet.id === previousSheetId && sheet.visible !== false)?.id ?? viewableSheets.find(sheet => sheet.visible !== false)?.id ?? ''
  if (sameDocument) saveViewSession()
  else { viewSessions.clear(); sortController?.abort(); cancelEdit(); closeNumberFormatPanel() }
  const session = sameDocument ? viewSessions.get(nextSheetId) : undefined
  workbook.value = cloneWorkbook(value)
  autoHeightBaseline = value
  activeSheetId.value = nextSheetId
  for (const id of viewSessions.keys()) if (!value.sheets.some(sheet => sheet.id === id)) viewSessions.delete(id)
  const sheet = activeSheet.value
  const apiSelection = value.id ? props.studio?.getSelectionState(value.id, nextSheetId) : undefined
  const next = normalizeSheetViewState(sheet ?? { rowCount: 0, columnCount: 0 }, apiSelection
    ? { ...session, ...sheetSelectionViewState(apiSelection), anchor: session?.anchor ?? sheetSelectionViewState(apiSelection).key }
    : session)
  selectedKey.value = next.key
  selectedRanges.value = next.ranges
  selectionAnchor.value = next.anchor
  activeSelectionIndex.value = next.activeRangeIndex ?? -1
  zoom.value = session?.zoom ?? displayPreference?.value.zoom ?? next.zoom
  if (nextSheetId !== previousSheetId || (editingKey.value && !parseSheetAddress(editingKey.value, sheet ?? { rowCount: 0, columnCount: 0 }))) cancelEdit()
}

function commitHistory(options: { prepared?: boolean; syncSelection?: boolean; current?: () => boolean } = {}) {
  if (options.current && !options.current()) return
  workbook.value.activeSheetId = activeSheetId.value
  if (!options.prepared) refreshWorkbookViewAutoRowHeights(workbook.value, autoHeightBaseline, {
    measure: props.studio?.sheetTextMeasure ?? measureSpreadsheetText, region: officeRegion.value,
  })
  const snapshot = cloneWorkbook(workbook.value)
  // The emitted snapshot is separate from the mutable view and doubles as the next measurement baseline.
  autoHeightBaseline = snapshot
  emit('update:modelValue', snapshot)
  if (options.current && !options.current()) return
  emit('update:model', snapshot)
  if (options.current && !options.current()) return
  emit('change', snapshot)
  if ((!options.current || options.current()) && options.syncSelection !== false) syncApiSelection(true)
}

function validateIncomingValues(entries: readonly { key: string; value: string }[], region = officeRegion.value, date1904 = activeDate1904.value, currentYear?: number): boolean {
  const sheet = activeSheet.value
  if (!sheet || !entries.length) return true
  if (!canEditWorksheetProtection(sheet, props.studio?.currentUserId ?? 'local-user')) {
    emit('notify', { type: 'warning', message: '此工作表受保护，当前用户没有编辑权限。' })
    return false
  }
  const positions = entries.map(({ key }) => {
    const [column, row] = splitKey(key)
    return { row: row - 1, column }
  })
  if (!canEditProtectedRanges(sheet, positions.map(({ row, column }) => ({ startRow: row, endRow: row, startCol: column, endCol: column })), userId.value)) {
    emit('notify', { type: 'warning', message: '目标单元格受保护，当前用户没有编辑权限。' })
    return false
  }
  const candidateView = cloneWorkbook(workbook.value)
  const candidateSheet = candidateView.sheets.find(item => item.id === sheet.id)
  if (!candidateSheet) return false
  for (const entry of entries) {
    candidateSheet.cells[entry.key] = editSheetCellText(candidateSheet.cells[entry.key], entry.value, date1904, region, currentYear)
  }
  const model = applyExcelWorkbookView(props.studio?.getWorkbookModel(workbook.value.id ?? '') ?? createBlankWorkbook(), candidateView, region)
  const result = validateSheetDataValidationChanges(model, sheet.id, positions)
  if (result.blocked) {
    const issue = result.issues.find(item => item.severity === 'stop')!
    emit('notify', { type: 'error', message: `${issue.title}：${issue.message}` })
    return false
  }
  if (result.issues.length) {
    const issue = result.issues[0]!
    emit('notify', { type: issue.severity === 'warning' ? 'warning' : 'info', message: `${issue.title}：${issue.message}` })
  }
  return true
}

function syncApiSelection(preserveTabStart = false) {
  const studio = props.studio
  const id = workbook.value.id
  if (applyingApiSelection || !studio || !id || !studio.getWorkbookModel(id)?.sheets.some(sheet => sheet.id === activeSheetId.value)) return
  const sheet = activeSheet.value, [column, row] = splitKey(selectedKey.value)
  const primary = containsCell(selectedRanges.value, row - 1, column) ? { row: row - 1, column } : undefined
  const active = selectedRanges.value[activeSelectionIndex.value]
  const activeRangeIndex = active && primary && containsCell([active], primary.row, primary.column) ? activeSelectionIndex.value : undefined
  const ranges: ApiSheetSelection['ranges'] = selectedRanges.value.map(range => ({ ...range,
    rangeType: range.rangeType ?? (range.startRow === 0 && range.endRow === (sheet?.rowCount ?? 0) - 1 && range.startColumn === 0 && range.endColumn === (sheet?.columnCount ?? 0) - 1 ? 3
      : range.startColumn === 0 && range.endColumn === (sheet?.columnCount ?? 0) - 1 ? 1
        : range.startRow === 0 && range.endRow === (sheet?.rowCount ?? 0) - 1 ? 2 : 0),
  }))
  syncingUiSelection = true
  try {
    const requested = normalizeSheetSelection(ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow,
      startCol: range.startColumn, endCol: range.endColumn, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })),
      sheet ?? {}, { primary, activeRangeIndex })
    const saved = studio.getSelectionState(id, activeSheetId.value)
    const same = saved && requested.activeRangeIndex === saved.activeRangeIndex
      && requested.primary?.row === saved.primary?.row && requested.primary?.column === saved.primary?.column
      && saved.ranges.length === ranges.length && saved.ranges.every((range, index) => {
        const next = ranges[index]!
        return range.startRow === next.startRow && range.endRow === next.endRow && range.startColumn === next.startColumn
          && range.endColumn === next.endColumn && range.rangeType === next.rangeType
      })
    if (preserveTabStart && saved && same && studio.activeWorkbookId === id && studio.editor.state.activeKind === 'workbook'
      && studio.getWorkbookModel(id)?.activeSheetId === activeSheetId.value) {
      activeSelectionIndex.value = saved.activeRangeIndex
      return
    }
    studio.setSelections(id, activeSheetId.value, ranges, { primary, activeRangeIndex: requested.activeRangeIndex })
  }
  catch (error) {
    const saved = studio.getSelectionState(id, activeSheetId.value)
    if (saved) {
      const previous = sheetSelectionViewState(saved)
      applyingApiSelection = true
      try {
        selectedRanges.value = previous.ranges; selectedKey.value = previous.key; selectionAnchor.value = previous.key
        activeSelectionIndex.value = previous.activeRangeIndex ?? -1
      } finally { applyingApiSelection = false }
    }
    selecting.value = false; resetHeaderSelection()
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '无法设置选区。' })
  } finally { syncingUiSelection = false }
  activeSelectionIndex.value = studio.getSelectionState(id, activeSheetId.value)?.activeRangeIndex ?? -1
}

function sheetSelectionViewState(selection: ApiSheetSelection): Pick<SheetViewState, 'key' | 'ranges' | 'activeRangeIndex'> {
  const primary = selection.primary
  return { key: primary ? `${columnName(primary.column)}${primary.row + 1}` : 'A1', activeRangeIndex: selection.activeRangeIndex,
    ranges: selection.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startColumn: range.startColumn,
      endColumn: range.endColumn, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) })) }
}

watch(() => props.studio, (studio, _previous, onCleanup) => {
  if (!studio) return
  const apply = (unitId: string, sheetId: string, selection: ApiSheetSelection) => {
    if (!workbenchActive.value || unitId !== workbook.value.id) return
    const { key, ranges, activeRangeIndex } = sheetSelectionViewState(selection)
    const sameCoordinates = ranges.length === selectedRanges.value.length && ranges.every((range, index) => {
      const current = selectedRanges.value[index]!
      return range.startRow === current.startRow && range.endRow === current.endRow && range.startColumn === current.startColumn && range.endColumn === current.endColumn
    })
    // Acknowledging the current header gesture must keep its drag/Shift anchor alive.
    if (syncingUiSelection && sheetId === activeSheetId.value && sameCoordinates) {
      if (selectedKey.value !== key) { selectedKey.value = key; selectionAnchor.value = key }
      activeSelectionIndex.value = activeRangeIndex ?? -1
      return
    }
    if (sheetId === activeSheetId.value && activeSelectionIndex.value === activeRangeIndex && selectedKey.value === key
      && JSON.stringify(ranges) === JSON.stringify(selectedRanges.value)) return
    const model = studio.getWorkbookModel(unitId)
    if (!model) return
    applyingApiSelection = true
    try {
      applyExternalWorkbook(toExcelWorkbookView(model))
      activeSheetId.value = sheetId
      selectedRanges.value = ranges; selectedKey.value = key; selectionAnchor.value = key
      activeSelectionIndex.value = activeRangeIndex ?? -1
      saveViewSession()
    } finally { applyingApiSelection = false }
  }
  const listener = studio.events.addEvent('SelectionsChanged', event => apply(event.workbookId, event.sheetId, event.selection))
  onCleanup(() => listener.dispose())
  const current = workbook.value.id ? studio.getSelectionState(workbook.value.id) : undefined
  const sheetId = workbook.value.id ? studio.getWorkbookModel(workbook.value.id)?.activeSheetId : undefined
  if (current && sheetId && workbook.value.id) apply(workbook.value.id, sheetId, current)
}, { immediate: true, flush: 'sync' })

function rangeNotation(range: GridRange): string {
  return `${columnName(range.startColumn)}${range.startRow + 1}:${columnName(range.endColumn)}${range.endRow + 1}`
}

function openChartDialog() {
  const sheet = activeSheet.value, studio = props.studio, unitId = workbook.value.id
  if (!sheet || !studio || !unitId) {
    emit('notify', { type: 'error', message: '图表命令上下文尚未绑定到活动工作簿。' })
    return
  }
  const bounds = selectionBounds(selectedRanges.value) ?? { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }
  openInsertChart({ type: 'bar', source: { range: chartSourceFromGridRange(bounds), orientation: 'columns', sheetId: sheet.id },
    position: chartPositionForRange(bounds, sheet), size: { width: 480, height: 300 }, categoryField: 0, valueFields: [] })
}

function openHyperlinkDialog(): void {
  const sheet = activeSheet.value
  const unitId = workbook.value.id
  const key = hyperlinkSelectionKey()
  if (!sheet || !unitId || !canOpenHyperlinkDialog()) return
  const cell = sheet.cells[key]
  hyperlinkCell.value = { unitId, sheetId: sheet.id, key }
  hyperlinkTarget.value = cell?.hyperlink ?? ''
  hyperlinkDisplay.value = cell?.displayValue ?? cell?.value ?? ''
  hyperlinkError.value = ''
  hyperlinkDialogOpen.value = true
}

function confirmHyperlinkDialog(): void {
  runHyperlinkCommand(SET_SHEET_HYPERLINK_COMMAND)
}

function deleteSelectedHyperlink(): void {
  runHyperlinkCommand(REMOVE_SHEET_HYPERLINK_COMMAND)
}

function openSelectedHyperlink(): void {
  const key = hyperlinkCell.value?.key ?? selectedKey.value
  const sheetId = hyperlinkCell.value?.sheetId ?? activeSheetId.value
  const target = hyperlinkCell.value?.unitId === workbook.value.id
    ? workbook.value.sheets.find(sheet => sheet.id === sheetId)?.cells[key]?.hyperlink
    : activeSheet.value?.cells[key]?.hyperlink
  if (target) navigateSheetHyperlink(target)
}

function openCellHyperlink(key: string): void {
  const target = activeSheet.value?.cells[key]?.hyperlink
  if (target) navigateSheetHyperlink(target)
}

function executeSheetContextCommand(commandId: string, params?: unknown): void {
  if (commandId === 'sheet.operation.open-cell-hyperlink' && typeof params === 'object' && params !== null && 'key' in params && typeof params.key === 'string') {
    openCellHyperlink(params.key)
    return
  }
  void executeRibbon(commandId, params)
}

function navigateSheetHyperlink(target: string): void {
  const resolved = parseSheetHyperlinkTarget(target, workbook.value)
  if (!resolved) {
    emit('notify', { type: 'warning', message: '此超链接目标不存在或无效。' })
    return
  }
  if (resolved.type === 'external') {
    const safeUrl = normalizeExternalSheetHyperlink(resolved.url)
    if (safeUrl) window.open(safeUrl, '_blank', 'noopener,noreferrer')
    return
  }
  const sheet = workbook.value.sheets.find(candidate => candidate.id === resolved.sheetId)
  if (!sheet || sheet.visible === false) {
    emit('notify', { type: 'warning', message: '超链接目标工作表当前不可用。' })
    return
  }
  if (!canViewWorksheetProtection(sheet, userId.value)) {
    emit('notify', { type: 'warning', message: '当前用户没有查看目标工作表的权限。' })
    return
  }
  activateSheet(sheet.id)
  if (activeSheetId.value !== sheet.id) return
  const { startRow, endRow, startCol, endCol } = resolved.range
  selectedKey.value = `${columnName(startCol)}${startRow + 1}`
  selectionAnchor.value = selectedKey.value
  selectedRanges.value = [{ startRow, endRow, startColumn: startCol, endColumn: endCol }]
  syncApiSelection()
  void nextTick(() => spreadsheetCanvas.value?.ensureVisible(selectedKey.value))
}

function hyperlinkSelectionKey(): string {
  const range = selectedRanges.value[0]
  return range ? `${columnName(range.startColumn)}${range.startRow + 1}` : selectedKey.value
}

function canOpenHyperlinkDialog(): boolean {
  const sheet = activeSheet.value
  if (!props.studio || !workbook.value.id || !sheet) return false
  const key = hyperlinkSelectionKey()
  const [col, row] = splitKey(key)
  const range = { startRow: row - 1, endRow: row - 1, startCol: col, endCol: col }
  return canSetSheetHyperlink(sheet, row - 1, col, sheet.cells[key]?.formula)
    && canEditWorksheetProtection(sheet, userId.value) && canEditProtectedRanges(sheet, [range], userId.value)
}

function runHyperlinkCommand(commandId: typeof SET_SHEET_HYPERLINK_COMMAND | typeof REMOVE_SHEET_HYPERLINK_COMMAND): void {
  const target = hyperlinkCell.value
  if (!props.studio || !target || target.unitId !== workbook.value.id) return
  const [col, row] = splitKey(target.key)
  try {
    props.studio.executeCommandSync(commandId, {
      unitId: target.unitId, sheetId: target.sheetId, row: row - 1, col,
      ...(commandId === SET_SHEET_HYPERLINK_COMMAND ? { draft: { target: hyperlinkTarget.value, display: hyperlinkDisplay.value } } : {}),
    })
    const model = props.studio.getWorkbookModel(target.unitId)
    if (model) workbook.value = toExcelWorkbookView(model)
    commitHistory()
    hyperlinkDialogOpen.value = false
  } catch (error) {
    hyperlinkError.value = error instanceof Error ? error.message : '无法修改超链接。'
  }
}

function getChartWorksheet(sheetId: string) {
  if (!props.studio || !workbook.value.id) return null
  return FOffice.newAPI(props.studio).getWorkbook(workbook.value.id)?.getSheetBySheetId(sheetId) ?? null
}

function getSlicer(id: string) { return getChartWorksheet(activeSheet.value?.id ?? '')?.getSlicerControl(id) ?? null }

async function refreshSlicerView() {
  if (!props.studio || !workbook.value.id) return
  const model = props.studio.getWorkbookModel(workbook.value.id)
  if (model) {
    workbook.value = toExcelWorkbookView(model)
    commitHistory({ prepared: true })
  }
}

async function applySlicerChange(id: string, change: (control: NonNullable<ReturnType<typeof getSlicer>>, context: SlicerCommandContext) => Promise<unknown>, expectedRevision?: number): Promise<void> {
  const studio = props.studio, unitId = workbook.value.id, sheetId = activeSheet.value?.id, user = userId.value
  if (!studio || !unitId || !sheetId || !workbenchActive.value) return
  const current = () => props.studio === studio && workbook.value.id === unitId && activeSheet.value?.id === sheetId && userId.value === user
    && workbenchActive.value && studio.currentUserId === user && studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook'
    && studio.getWorkbookModel(unitId)?.activeSheetId === sheetId
  if (!current()) return
  try {
    const control = FOffice.newAPI(studio).getWorkbook(unitId)?.getSheetBySheetId(sheetId)?.getSlicerControl(id)
    if (!control) throw new Error('切片器所属的工作簿、工作表或对象已被删除。')
    const accepted = await change(control, { host: 'workbench', expectedRevision: expectedRevision ?? studio.getWorkbookRevision(unitId) })
    if (accepted === false) throw new Error('切片器操作未能提交。')
    if (!current()) return
    const model = studio.getWorkbookModel(unitId)
    if (model) {
      applyExternalWorkbook(toExcelWorkbookView(model))
      commitHistory({ prepared: true, current })
    }
  } catch (error) { if (current()) emit('notify', { type: 'error', message: error instanceof Error ? error.message : '切片器更新失败' }) }
}
async function updateSlicerPresentation(id: string, patch: import('../features/slicer/slicer-engine').SlicerPresentationPatch, expectedRevision?: number) {
  await applySlicerChange(id, (control, context) => control.updatePresentation(patch, context), expectedRevision)
}

async function handleSlicerItem(payload: { slicerId: string; value: string; index: number; anchorIndex?: number; ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; multiSelect: boolean; expectedRevision?: number }) {
  await applySlicerChange(payload.slicerId, async (control, context) => {
    const state = slicerControlItemStates(control.getItemLabels(), control.getItemState(), payload)
    if (!await control.setItemStates(state.checkedItems, state.partialItems, context)) throw new Error('切片器成员选择未能提交。')
  }, payload.expectedRevision)
}

async function clearSlicer(id: string, expectedRevision?: number) {
  await applySlicerChange(id, (control, context) => control.clearFilter(context), expectedRevision)
}

function requestSlicerWindow(payload: { slicerId: string; start: number; count: number; expectedRevision?: number }): void {
  const studio = props.studio, unitId = workbook.value.id, sheetId = activeSheet.value?.id
  if (!studio || !unitId || !sheetId || !workbenchActive.value || studio.isDisposed || studio.activeWorkbookId !== unitId
    || studio.editor.state.activeKind !== 'workbook' || studio.getWorkbookModel(unitId)?.activeSheetId !== sheetId
    || payload.expectedRevision !== undefined && payload.expectedRevision !== studio.getWorkbookRevision(unitId)) return
  slicerCanvasData.request(payload.slicerId, payload.start, payload.count)
}

async function setSlicerTimelineRange(payload: { slicerId: string; range: import('../domain/timeline-filters.ts').ITimelineRange; expectedRevision?: number }) {
  await applySlicerChange(payload.slicerId, (control, context) => getChartWorksheet(control.subUnitId)?.getTimeline(control.id)?.setDateRange(payload.range, context)
    ?? Promise.resolve(false), payload.expectedRevision)
}

function openSlicerSettings(id: string) { selectSlicer(id); slicerPanelSection.value = 'settings'; slicerPanelOpen.value = true }
function openSlicerConnections(id: string) {
  const control = getSlicer(id)?.getModel()
  if (!control || control.tableId) return
  selectSlicer(id); slicerPanelSection.value = 'connections'; slicerPanelOpen.value = true
}
function openSlicerPanel(kind: 'slicer' | 'timeline') { selectedSlicerId.value = ''; slicerPanelSection.value = 'settings'; slicerKind.value = kind; slicerPanelOpen.value = true }
function slicerRibbonIntentId(params: unknown): string | undefined {
  if (typeof params !== 'object' || !params || !('slicerId' in params) || typeof params.slicerId !== 'string') return undefined
  return slicerRibbonControl.value?.id === params.slicerId ? params.slicerId : undefined
}
async function handleSlicerPanelChanged(id: string) { selectedSlicerId.value = id; await refreshSlicerView() }
function handleSlicerPanelClosed(owner?: SlicerPanelOwner): void {
  if (!owner) return
  void nextTick(() => {
    if (props.studio !== owner.studio || workbook.value.id !== owner.target.workbookId || activeSheetId.value !== owner.target.sheetId
      || !slicerPanelOwnerCurrent(owner, true) || !workbenchActive.value || slicerPanelOpen.value || editingKey.value || editComposing.value
      || imageUiOpen.value || rangeThemePending.value || cellShiftPending.value || mergeRequest.value || painterRequest.value) return
    spreadsheetCanvas.value?.focus()
  })
}
function moveSlicer(payload: { slicerId: string; left: number; top: number }) {
  const info = getSlicer(payload.slicerId)?.getModel()
  if (!info || info.lockedPosition || !activeSheet.value) return
  const bounds = resolveSlicerBounds(info, new SpreadsheetGeometry(activeSheet.value))
  void updateSlicerPresentation(payload.slicerId, { layout: { ...bounds, left: payload.left, top: payload.top } })
}
function resizeSlicer(payload: { slicerId: string; width: number; height: number }) {
  const info = getSlicer(payload.slicerId)?.getModel()
  if (!info || info.lockedPosition || !activeSheet.value) return
  const bounds = resolveSlicerBounds(info, new SpreadsheetGeometry(activeSheet.value))
  const height = info.drawing?.keepRatio ? payload.width * bounds.height / bounds.width : payload.height
  void updateSlicerPresentation(payload.slicerId, { layout: { ...bounds, width: payload.width, height } })
}
function transformSlicer(payload: { slicerId: string; geometry: SheetDrawingAffineGeometry; expectedRevision?: number }) {
  const control = getSlicer(payload.slicerId)
  if (!control || control.getModel().lockedPosition) return
  void updateSlicerPresentation(payload.slicerId, { drawingGeometry: payload.geometry }, payload.expectedRevision)
}
function scrollSlicerTo(payload: { slicerId: string; position: number; expectedRevision?: number }) {
  const info = getSlicer(payload.slicerId)?.getModel()
  if (!info) return
  void updateSlicerPresentation(payload.slicerId, info.type === 'timeline' ? { scrollPosition: payload.position } : { startItem: payload.position }, payload.expectedRevision)
}

async function updateChartLayout(payload: { chartId: string; patch: Partial<Pick<ChartModel, 'position' | 'size'>> }) {
  await applyChartChange(payload.chartId, async chart => { await chart.update(payload.patch); return true })
}

async function applyChartChange(chartId: string, change: (chart: FChart) => Promise<boolean>, applied?: () => void): Promise<void> {
  const studio = props.studio, unitId = workbook.value.id, sheetId = activeSheet.value?.id
  if (!studio || !unitId || !sheetId || !workbenchActive.value) return
  const current = (): boolean => props.studio === studio && workbook.value.id === unitId && activeSheet.value?.id === sheetId
    && workbenchActive.value && studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook'
    && studio.getWorkbookModel(unitId)?.activeSheetId === sheetId
  try {
    const chart = FOffice.newAPI(studio).getWorkbook(unitId)?.getSheetBySheetId(sheetId)?.getChartById(chartId)
    if (!chart) throw new Error('图表所属的工作簿、工作表或图表已被删除。')
    if (!await change(chart)) throw new Error('图表操作未能提交。')
    if (!current()) return
    const model = studio.getWorkbookModel(unitId)
    if (model) applyExternalWorkbook(toExcelWorkbookView(model))
    if (current()) applied?.()
  } catch (reason) {
    if (current()) emit('notify', { type: 'error', message: reason instanceof Error ? reason.message : String(reason) })
  }
}

function openFocusedChartPanel(target: ChartObjectTarget, route?: ChartEditPanelRoute): boolean {
  const studio = props.studio
  if (!studio || !workbenchActive.value || workbook.value.id !== target.unitId || activeSheet.value?.id !== target.sheetId
    || studio.activeWorkbookId !== target.unitId || studio.editor.state.activeKind !== 'workbook'
    || studio.getWorkbookRevision(target.unitId) !== target.expectedRevision) return false
  const chart = studio.getWorkbookModel(target.unitId)?.charts?.find(chart => chart.id === target.chartId && chart.sheetId === target.sheetId)
  if (!chart || JSON.stringify(chart) !== target.snapshot) return false
  selectChart(target.chartId)
  return openChartEditor(chart, route) && chartEditorOpen.value && chartEditorContext.value?.chartId === target.chartId
}
function handleChartObjectChange(change: ChartObjectChange): void {
  const studio = props.studio
  if (!studio || !workbenchActive.value || workbook.value.id !== change.unitId || activeSheet.value?.id !== change.sheetId) return
  const model = studio.getWorkbookModel(change.unitId)
  if (model) applyExternalWorkbook(toExcelWorkbookView(model))
  if (change.action === 'paste') selectChart(change.chartId)
  if (change.action === 'delete') {
    if (focusedChartId.value === change.chartId) focusedChartId.value = null
    if (chartEditorContext.value?.chartId === change.chartId) chartEditorContext.value.setOpen(false)
  }
}

async function removeChart(chartId: string) {
  const studio = props.studio, unitId = workbook.value.id, sheetId = activeSheet.value?.id
  if (!studio || !unitId || !sheetId) return
  try {
    if (await studio.executeCommand(CHART_OBJECT_COMMANDS.delete, { unitId, sheetId, chartId, expectedRevision: studio.getWorkbookRevision(unitId) })) {
      handleChartObjectChange({ unitId, sheetId, chartId, action: 'delete' })
    }
  } catch (reason) { emit('notify', { type: 'error', message: reason instanceof Error ? reason.message : String(reason) }) }
}

function setCellValue(key: string, value: string, region = officeRegion.value, date1904 = activeDate1904.value,
  currentYear = new Date().getFullYear(), session?: SheetCellEditSession, editTrace?: SheetCellInputEditTrace, inlineFormats?: SheetCellInlineFormatEdit[]): boolean {
  const sheet = activeSheet.value
  if (!sheet) return false
  const unitId = workbook.value.id, [column, row] = splitKey(key)
  const input: SheetCellInputCommandInput = { unitId: unitId ?? 'standalone-input', sheetId: sheet.id,
    row: row - 1, column, input: value, region, date1904, currentYear, ...(session ? { session } : {}), ...(editTrace ? { editTrace } : {}),
    ...(inlineFormats?.length ? { inlineFormats } : {}) }
  const studio = props.studio
  const sameTarget = () => props.studio === studio && workbook.value.id === unitId && activeSheetId.value === sheet.id
    && workbenchActive.value && (!studio || !unitId || studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook'
      && studio.getWorkbookModel(unitId)?.activeSheetId === sheet.id)
  try {
    let result: SheetCellInputCommandResult
    if (studio && unitId) {
      result = studio.executeCommandSync(COMMIT_SHEET_CELL_INPUT_COMMAND, input)
      if (!result.accepted) {
        emit('notify', { type: result.reason === 'conflict' ? 'warning' : 'error', message: result.message ?? '单元格编辑未能提交。' })
        return false
      }
      // Observers may have switched units or started another edit after the committed transaction.
      if (!sameTarget()) return true
      const model = studio.getWorkbookModel(unitId)
      if (!model) return false
      applyExternalWorkbook(toExcelWorkbookView(model))
    } else {
      const before = applyExcelWorkbookView(createBlankWorkbook(), workbook.value, region)
      const plan = planSheetCellInput({ ...before, date1904 }, input, { region, currentYear, userId: userId.value })
      result = plan.result
      if (!result.accepted) {
        emit('notify', { type: result.reason === 'conflict' ? 'warning' : 'error', message: result.message ?? '单元格编辑未能提交。' })
        return false
      }
      if (result.changed && plan.cell) {
        let next = cloneWorkbook(workbook.value)
        const target = next.sheets.find(item => item.id === sheet.id)!
        target.cells[`${columnName(plan.position.column)}${plan.position.row + 1}`] = toExcelCellView(plan.cell)
        if (before.calculationMode !== 'manual') next = recalculateSheetViewFormulas(next, undefined, { numberFormatLocale: region }).workbook
        refreshWorkbookViewAutoRowHeights(next, autoHeightBaseline, { measure: props.studio?.sheetTextMeasure ?? measureSpreadsheetText, region })
        applyExternalWorkbook(next)
      }
    }
    const publishedView = workbook.value, revision = studio && unitId ? studio.getWorkbookRevision(unitId) : undefined
    const current = () => sameTarget() && workbook.value === publishedView
      && (!studio || !unitId || studio.getWorkbookRevision(unitId) === revision)
    if (result.issues.length) {
      const issue = result.issues[0]!
      emit('notify', { type: issue.severity === 'warning' ? 'warning' : 'info', message: `${issue.title}：${issue.message}` })
    }
    if (result.changed && current() && activeSheet.value) {
      const contentKey = sheetCellContentKey(activeSheet.value, key), cell = activeSheet.value.cells[contentKey]
      if (cell) emit('edit-cell', { sheetId: sheet.id, key: contentKey, cell: structuredClone(toRaw(cell)) })
      commitHistory({ prepared: true, syncSelection: !studio || !unitId, current })
    }
    return true
  } catch (error) {
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '单元格编辑未能提交' })
    return false
  }
}

function recalculateFormulas() {
  if (workbook.value.calculationMode === 'manual') return
  const previous = workbook.value.id ? props.studio?.getWorkbookModel(workbook.value.id) : undefined
  const result = recalculateSheetViewFormulas(workbook.value, previous, { numberFormatLocale: officeRegion.value })
  workbook.value = result.workbook
}

function recalculateWorkbook() {
  const workbookId = workbook.value.id
  if (!workbookId || !props.studio) {
    const result = recalculateSheetViewFormulas(workbook.value, undefined, { numberFormatLocale: officeRegion.value })
    workbook.value = result.workbook
    return
  }
  const book = FOffice.newAPI(props.studio).getWorkbook(workbookId)
  if (book) {
    book.recalculate()
    const model = props.studio?.getWorkbookModel(workbookId)
    if (model) workbook.value = toExcelWorkbookView(model)
    return
  }
}

function splitKey(key: string): [number, number] {
  const match = key.match(/^([A-Z]+)(\d+)$/i)
  if (!match) return [0, 1]
  let column = 0
  for (const character of match[1].toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64
  return [column - 1, Number(match[2])]
}

function getCell(key: string): ExcelCell {
  const sheet = activeSheet.value, position = parseCellKey(key)
  return sheet && position ? sheetViewCellAt(sheet, position.row, position.col) : { value: '' }
}

function selectCell(key: string, extend = false, keepHidden = false) {
  standaloneSelectionTabStart = undefined
  const [requestedColumn, requestedRow] = splitKey(key)
  const visibleRow = keepHidden ? requestedRow - 1 : activeSheet.value ? findVisibleSheetRow(activeSheet.value, requestedRow - 1, 1) : undefined
  const column = keepHidden ? requestedColumn : activeSheet.value ? findVisibleSheetColumn(activeSheet.value, requestedColumn, 1) : undefined
  if (visibleRow === undefined || column === undefined) return
  const row = visibleRow + 1
  key = `${columnName(column)}${row}`
  if (extend) {
    const [anchorColumn, anchorRow] = splitKey(selectionAnchor.value)
    selectedRanges.value = [{ startRow: Math.min(anchorRow, row) - 1, endRow: Math.max(anchorRow, row) - 1, startColumn: Math.min(anchorColumn, column), endColumn: Math.max(anchorColumn, column) }]
  } else {
    selectionAnchor.value = key
    selectedRanges.value = [{ startRow: row - 1, endRow: row - 1, startColumn: column, endColumn: column }]
  }
  selectedKey.value = key
  syncApiSelection()
}

function beginEdit(key = selectedKey.value, initialValue?: string, target: 'cell' | 'formula' = 'cell', focusEditor = true) {
  if (imageUiOpen.value) return
  drawingSelection.set([])
  const sheet = activeSheet.value
  const unitId = workbook.value.id
  const address = sheet ? parseSheetAddress(key, sheet) : undefined
  if (!sheet || !address) return
  key = sheetCellContentKey(sheet, address.anchor)
  if (editingKey.value === key && editSession.value?.documentId === unitId && editSession.value.sheetId === sheet.id && initialValue === undefined) {
    const session = editSession.value
    if (focusEditor) nextTick(() => {
      if (editSession.value !== session || !workbenchActive.value) return
      target === 'cell' ? spreadsheetCanvas.value?.focusEditor() : formulaInput.value?.focus()
    })
    return
  }
  if (editingKey.value && !finishEdit()) return
  const currentSheet = activeSheet.value
  if (!currentSheet || currentSheet.id !== sheet.id || workbook.value.id !== unitId || !workbenchActive.value
    || props.studio && unitId && (props.studio.activeWorkbookId !== unitId || props.studio.editor.state.activeKind !== 'workbook')) return
  key = sheetCellContentKey(currentSheet, key)
  const sourceCell = getCell(key)
  const session = createSheetCellEditSession({ documentId: workbook.value.id, sheetId: currentSheet.id, key, date1904: activeDate1904.value, cell: sourceCell },
    { region: officeRegion.value, currentYear: new Date().getFullYear(), mode: initialValue === undefined ? 'edit' : 'overwrite' })
  const start = createSheetCellEditorStart(sourceCell, initialValue, session.date1904, session.region)
  const source = sheetCellInputText(sourceCell, session.date1904, session.region)
  editRecorder = new SheetCellInputEditRecorder(session.sessionId!, source)
  editInlineDraft = new SheetCellInlineFormatDraft()
  editTextSelection.value = { ...start.selection, host: target }
  editFontRevision.value++
  if (initialValue !== undefined) {
    const event = { sessionId: session.sessionId!, hostId: `initial-${session.sessionId}`, host: 'cell' as const,
      value: source, start: 0, end: source.length }
    editRecorder.handle({ ...event, phase: 'bind' })
    editRecorder.handle({ ...event, phase: 'replace', text: start.value, inputType: 'insertText' })
  }
  editSessionStudio = props.studio
  cellEditorComposing.value = false
  formulaEditorComposing.value = false
  editSession.value = session
  editingKey.value = key
  editValue.value = start.value
  if (focusEditor) nextTick(() => {
    if (editSession.value !== session || editingKey.value !== key || !workbenchActive.value) return
    if (target === 'cell') spreadsheetCanvas.value?.focusEditor(start.selection)
    else formulaInput.value?.focus(start.selection)
  })
}

function beginFormulaEdit() {
  // A pointer focus already has a caret chosen by the user; keep it when the formula bar starts its shared session.
  if (!editingKey.value) beginEdit(selectedKey.value, undefined, 'formula', false)
}

function handleCellNativeInput(event: SheetCellNativeInputEvent): void {
  const session = editSession.value, recorder = editRecorder
  if (!session || !recorder || event.sessionId !== session.sessionId || props.studio !== editSessionStudio
    || editingKey.value !== session.key || workbook.value.id !== session.documentId || activeSheetId.value !== session.sheetId
    || !workbenchActive.value || props.studio && session.documentId && (props.studio.activeWorkbookId !== session.documentId
      || props.studio.editor.state.activeKind !== 'workbook')) return
  if (!recorder.handle(event)) return
  editInlineDraft?.reconcile(recorder.trace)
  const caret = event.phase === 'replace' ? event.start + (event.text?.length ?? 0) : undefined
  editTextSelection.value = { start: caret ?? event.start, end: caret ?? event.end, host: event.host }
  editFontRevision.value++
  editValue.value = recorder.value
  cellEditorComposing.value = recorder.compositionHost === 'cell'
  formulaEditorComposing.value = recorder.compositionHost === 'formula'
}

function finishFormulaEdit(navigation?: { key: 'Enter' | 'Tab'; shift: boolean }) {
  if (!finishEdit()) return
  if (navigation) moveSelectionForShortcut(navigation.key, navigation.shift)
  if (workbenchActive.value && (!props.studio || props.studio.activeWorkbookId === workbook.value.id
    && props.studio.editor.state.activeKind === 'workbook')) spreadsheetCanvas.value?.focus()
}

function cancelFormulaEdit() { cancelEdit(); spreadsheetCanvas.value?.focus() }

function navigateToAddress(input: string) {
  const sheet = activeSheet.value
  if (!sheet) return
  const parsed = parseSheetAddress(input, sheet)
  if (!parsed) { nameError.value = '请输入当前工作表内的有效地址或范围，例如 A1、A1:C8'; return }
  if (!finishEdit()) return
  nameError.value = ''
  const next = normalizeSheetViewState(sheet, { key: parsed.anchor, anchor: parsed.anchor, ranges: [parsed.range], zoom: zoom.value })
  selectedKey.value = next.key
  selectionAnchor.value = next.anchor
  selectedRanges.value = next.ranges
  syncApiSelection()
  spreadsheetCanvas.value?.ensureVisible(next.key)
  spreadsheetCanvas.value?.focus()
}

function inputNavigationState(): string {
  const studio = props.studio, unitId = workbook.value.id, sheetId = activeSheetId.value
  return JSON.stringify({ unitId, sheetId, active: workbenchActive.value,
    activeUnit: studio?.activeWorkbookId, activeKind: studio?.editor.state.activeKind,
    selection: studio && unitId ? studio.getSelectionState(unitId, sheetId)
      : { key: selectedKey.value, ranges: selectedRanges.value, activeRangeIndex: activeSelectionIndex.value } })
}

function finishEdit(): boolean {
  if (!editingKey.value) return true
  if (editComposing.value) return false
  const session = editSession.value
  const sheet = activeSheet.value
  const conflict = session && sheet ? getSheetCellEditConflict(session, { documentId: workbook.value.id, sheetId: sheet.id,
    key: editingKey.value, date1904: activeDate1904.value, cell: getCell(editingKey.value) }) : 'cell'
  if (conflict) {
    emit('notify', { type: 'warning', message: sheetCellInputConflictMessage(conflict) })
    return false
  }
  const navigation = inputNavigationState()
  const trace = editRecorder?.trace
  if (!session || !setCellValue(editingKey.value, editValue.value, session.region, session.date1904, session.currentYear, session, trace,
    trace ? editInlineDraft?.formats(trace) : undefined)) return false
  if (editSession.value !== session || workbook.value.id !== session.documentId || activeSheetId.value !== session.sheetId) return false
  const continueAction = navigation === inputNavigationState()
  cancelEdit()
  if (!continueAction) void nextTick(() => {
    if (editSession.value || rangeThemePending.value || cellShiftPending.value || !workbenchActive.value || workbook.value.id !== session.documentId || activeSheetId.value !== session.sheetId
      || props.studio && (props.studio.activeWorkbookId !== session.documentId || props.studio.editor.state.activeKind !== 'workbook')) return
    if (document.activeElement === document.body) spreadsheetCanvas.value?.focus()
  })
  return continueAction
}

function prepareCanvasFillGesture(): SheetFillCommitContext | undefined {
  const studio = props.studio, unitId = workbook.value.id, sheetId = activeSheetId.value
  if (!finishEdit() || !workbenchActive.value || props.studio !== studio
    || workbook.value.id !== unitId || activeSheetId.value !== sheetId || !activeSheet.value
    || studio && unitId && (studio.activeWorkbookId !== unitId || studio.editor.state.activeKind !== 'workbook')) return undefined
  return { sheetId, workbookRevision: studio && unitId ? studio.getWorkbookRevision(unitId) : undefined }
}

function cancelEdit() {
  editingKey.value = null
  editSession.value = null
  editRecorder = undefined
  editInlineDraft = undefined
  editFontRevision.value++
  editSessionStudio = undefined
  cellEditorComposing.value = false
  formulaEditorComposing.value = false
}

function moveSelection(rowOffset: number, columnOffset: number, extend = false) {
  const [column, row] = splitKey(selectedKey.value)
  const nextColumn = activeSheet.value ? findVisibleSheetColumn(activeSheet.value, column + columnOffset, columnOffset < 0 ? -1 : 1, false) : undefined
  const nextRow = activeSheet.value ? findVisibleSheetRow(activeSheet.value, row - 1 + rowOffset, rowOffset < 0 ? -1 : 1, false) : undefined
  if (nextRow === undefined || nextColumn === undefined) return
  selectCell(`${columnName(nextColumn)}${nextRow + 1}`, extend)
}

function moveSelectionForShortcut(key: SheetSelectionShortcutKey, shift = false): void {
  const sheet = activeSheet.value
  if (!sheet || !selectedRanges.value.length) return
  const direction = key === 'Tab' ? shift ? SheetSelectionDirections.LEFT : SheetSelectionDirections.RIGHT
    : shift ? SheetSelectionDirections.UP : SheetSelectionDirections.DOWN
  try {
    const studio = props.studio, unitId = workbook.value.id
    if (studio && unitId) {
      studio.executeCommandSync(MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND, { unitId, sheetId: sheet.id, direction, key })
      return
    }
    const [column, row] = splitKey(selectedKey.value)
    const state = normalizeSheetSelection(selectedModelRanges().map((range, index) => ({ ...range,
      ...(selectedRanges.value[index]!.rangeType === undefined ? {} : { rangeType: selectedRanges.value[index]!.rangeType }) })), sheet,
      { primary: { row: row - 1, column }, ...(activeSelectionIndex.value < 0 ? {} : { activeRangeIndex: activeSelectionIndex.value }) })
    const mergedRanges = (sheet.merges ?? []).flatMap(value => {
      const parsed = parseSheetAddress(value, sheet)?.range
      return parsed ? [{ startRow: parsed.startRow, endRow: parsed.endRow, startCol: parsed.startColumn, endCol: parsed.endColumn }] : []
    })
    const plan = planSheetSelectionShortcut(state, direction, key, { ...sheet, mergedRanges }, standaloneSelectionTabStart)
    if (!plan?.selection.primary) return
    applyingShortcutSelection = true
    try {
      selectedRanges.value = plan.selection.ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow,
        startColumn: range.startCol, endColumn: range.endCol, ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) }))
      selectedKey.value = `${columnName(plan.selection.primary.column)}${plan.selection.primary.row + 1}`
      selectionAnchor.value = selectedKey.value
      activeSelectionIndex.value = plan.selection.activeRangeIndex
      standaloneSelectionTabStart = plan.tabStart
    } finally { applyingShortcutSelection = false }
  } catch (error) {
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '无法移动选区。' })
  }
}

function handleGridKeydown(event: KeyboardEvent) {
  if (imageUiOpen.value || outlineUiOpen.value) return
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return
  const target = event.target as HTMLElement | null
  if (!target?.closest('.spreadsheet-canvas')) return
  if (target?.matches('input, textarea, select')
    && !target.closest('.canvas-cell-input, .formula-input')) return
  if (event.ctrlKey && event.key === 'PageDown') {
    event.preventDefault()
    switchNextSheet()
    return
  }
  if (handleInlineFormatHistoryKey(event)) return
  const fontCommand = sheetFontFormatCommandForShortcut(event)
  if (fontCommand) {
    event.preventDefault(); void executeRibbon(fontCommand); return
  }
  if (editingKey.value) {
    const action = sheetCellInputActionForKey(event, editComposing.value)
    if (action === 'line-break') {
      event.preventDefault()
      spreadsheetCanvas.value?.insertEditorText('\n')
    } else if (action === 'cancel') {
      event.preventDefault(); cancelEdit(); spreadsheetCanvas.value?.focus()
    } else if (action !== 'native') {
      event.preventDefault()
      if (!finishEdit()) return
      if (action !== 'commit') moveSelectionForShortcut(action === 'next-column' || action === 'previous-column' ? 'Tab' : 'Enter',
        action === 'previous-column' || action === 'previous-row')
      spreadsheetCanvas.value?.focus()
    }
    return
  }
  if (handleClipboardKeydown(event)) return
  if (handleFillKeydown(event)) return
  const outlineShortcut = sheetOutlineCommandForShortcut(event, typeof navigator !== 'undefined' && /Mac|iPhone|iPad/u.test(navigator.platform))
  if (outlineShortcut) { event.preventDefault(); void executeRibbon(outlineShortcut); return }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') {
    event.preventDefault()
    void executeRibbon('sheet.operation.insert-hyperlink')
    return
  }
  if (event.key === 'Enter' && !event.shiftKey && !event.altKey && !event.ctrlKey && !event.metaKey && activeCell.value.hyperlink) {
    event.preventDefault()
    openCellHyperlink(selectedKey.value)
    return
  }
  if (event.key === 'F2') {
    event.preventDefault()
    beginEdit()
    return
  }
  if ((event.key === 'Enter' || event.key === 'Tab') && !event.altKey && !event.ctrlKey && !event.metaKey) {
    event.preventDefault()
    moveSelectionForShortcut(event.key, event.shiftKey)
    return
  }
  if (event.key === 'ArrowDown') {
    event.preventDefault()
    moveSelection(1, 0, event.shiftKey)
  } else if (event.key === 'ArrowUp') {
    event.preventDefault()
    moveSelection(-1, 0, event.shiftKey)
  } else if (event.key === 'ArrowLeft') {
    event.preventDefault()
    moveSelection(0, -1, event.shiftKey)
  } else if (event.key === 'ArrowRight') {
    event.preventDefault()
    moveSelection(0, 1, event.shiftKey)
  } else if (event.key === 'Delete' || event.key === 'Backspace') {
    event.preventDefault()
    void executeRibbon(SHEET_CLEAR_COMMANDS.content)
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') {
    event.preventDefault()
    if (event.shiftKey) redo(); else undo()
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') {
    event.preventDefault()
    redo()
  } else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey) {
    event.preventDefault()
    beginEdit(selectedKey.value, event.key)
  }
}

const maximumDirectCells = 100_000
function checkDirectSelection(count = selectedCount.value): boolean {
  if (count <= maximumDirectCells) return true
  emit('notify', { type: 'warning', message: `此操作最多处理 ${maximumDirectCells.toLocaleString()} 个单元格，请缩小当前选区。` })
  return false
}
function applyToSelection(action: (key: string) => void): boolean {
  const sheet = activeSheet.value
  if (sheet && !canEditWorksheetProtection(sheet, props.studio?.currentUserId ?? 'local-user')) {
    emit('notify', { type: 'warning', message: '此工作表受保护，当前用户没有编辑权限。' })
    return false
  }
  if (!checkDirectSelection()) return false
  if (!sheet || !selectedRanges.value.length || !canEditProtectedRanges(sheet, selectedModelRanges(), userId.value)) return false
  for (const range of selectedRanges.value) for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) action(`${columnName(column)}${row + 1}`)
  return true
}
function selectedExistingKeys(): string[] {
  return Object.keys(sheetCells.value).filter(key => { const [column, row] = splitKey(key); return containsCell(selectedRanges.value, row - 1, column) })
}
function undo() {
  if (editingKey.value) { runCellDraftHistory(-1, true); return }
  emit('undo', workbook.value.id ? { workbookId: workbook.value.id } : undefined)
}

function redo() {
  if (editingKey.value) { runCellDraftHistory(1, true); return }
  emit('redo', workbook.value.id ? { workbookId: workbook.value.id } : undefined)
}

function keepCellEditFocus(target: Element | null): boolean {
  const session = editSession.value
  if (!target || !session || !workbenchActive.value || workbook.value.id !== session.documentId || activeSheetId.value !== session.sheetId) return false
  if (target.closest('[data-ribbon-editor-owner]')?.getAttribute('data-ribbon-editor-owner') === ribbonEditorOwner) return true
  const popup = target.closest('.ant-select-dropdown, .ant-popover, .ant-dropdown, .ant-color-picker')
  return [...(popup?.querySelectorAll('[data-ribbon-editor-owner]') ?? [])]
    .some(element => element.getAttribute('data-ribbon-editor-owner') === ribbonEditorOwner)
}

function restoreCellInputFocus(session: SheetCellEditSession, selection = { ...editTextSelection.value }): void {
  const studio = props.studio
  // Popup controls restore their trigger after closing; restore the input after that flush.
  void nextTick(() => nextTick(() => {
    if (editSession.value !== session || props.studio !== studio || !workbenchActive.value || editComposing.value
      || editingKey.value !== session.key || workbook.value.id !== session.documentId || activeSheetId.value !== session.sheetId
      || studio && session.documentId && (studio.activeWorkbookId !== session.documentId || studio.editor.state.activeKind !== 'workbook')) return
    const focused = document.activeElement
    if (focused instanceof Element && focused !== document.body && !keepCellEditFocus(focused)
      && !focused.closest('.canvas-cell-input, .formula-toolbar')) return
    if (selection.host === 'formula') formulaInput.value?.focus(selection)
    else spreadsheetCanvas.value?.focusEditor(selection)
  }))
}

function runCellDraftHistory(direction: -1 | 1, includeText = false): boolean {
  const session = editSession.value, recorder = editRecorder, draft = editInlineDraft
  if (!session || !recorder || !draft || editComposing.value || !workbenchActive.value) return false
  const formatted = direction < 0 ? draft.undo(recorder.trace) : draft.redo(recorder.trace)
  if (!formatted) {
    if (!includeText) return false
    const restored = recorder.moveDraftHistory(direction)
    if (!restored) return false
    editValue.value = restored.value
    editTextSelection.value = { ...restored.selection, host: editTextSelection.value.host }
    draft.reconcile(recorder.trace)
  }
  editFontRevision.value++
  restoreCellInputFocus(session)
  return true
}

function handleInlineFormatHistoryKey(event: KeyboardEvent): boolean {
  if (event.defaultPrevented || event.isComposing || event.keyCode === 229 || event.altKey || !editingKey.value || (!event.ctrlKey && !event.metaKey)) return false
  const key = event.key.toLowerCase()
  const direction = key === 'z' ? event.shiftKey ? 1 : -1 : key === 'y' && !event.shiftKey ? 1 : undefined
  if (!direction || !runCellDraftHistory(direction)) return false
  event.preventDefault(); event.stopPropagation(); return true
}

function runInlineFontFormatCommand(commandId: string, params?: unknown): boolean {
  const session = editSession.value, recorder = editRecorder, draft = editInlineDraft, sheet = activeSheet.value
  if (!editingKey.value || !session || !recorder || !draft || !sheet) return false
  try {
    const action = fontFormatActionForCommand(commandId, params)
    if (!action || action.kind === 'color' && action.field !== 'color') return false
    const trace = recorder.trace
    if (!inlineFontEditable.value) {
      if (!draft.formats(trace).length) return false
      emit('notify', { type: 'warning', message: '当前草稿包含局部文字格式，请保持文本类型后继续；草稿已保留。' })
      return true
    }
    if (editComposing.value || !canFormatSelection.value || !workbenchActive.value) return true
    const unitId = workbook.value.id, studio = props.studio, [column, row] = splitKey(session.key)
    const input: SheetCellInlineFontFormatInput = { unitId: unitId ?? 'standalone-input', sheetId: session.sheetId,
      row: row - 1, column, input: recorder.value, session, editTrace: trace, inlineFormats: draft.formats(trace),
      commandId, params, selection: { start: editTextSelection.value.start, end: editTextSelection.value.end } }
    const result = studio && unitId ? studio.executeCommandSync<SheetCellInlineFontFormatInput, SheetCellInlineFontFormatResult>(FORMAT_SHEET_CELL_INPUT_OPERATION, input)
      : planSheetCellInlineFontFormat({ ...applyExcelWorkbookView(createBlankWorkbook(), workbook.value, session.region), date1904: session.date1904 }, input,
        { region: session.region, currentYear: session.currentYear, userId: userId.value })
    if (editSession.value !== session || editInlineDraft !== draft || props.studio !== studio || workbook.value.id !== unitId
      || activeSheetId.value !== session.sheetId || !workbenchActive.value) return true
    if (!result.accepted || !result.edit) {
      if (result.message) emit('notify', { type: 'warning', message: result.message })
      return true
    }
    const currentCell = getCell(session.key)
    if (getSheetCellEditConflict(session, { documentId: workbook.value.id, sheetId: session.sheetId, key: session.key, date1904: activeDate1904.value, cell: currentCell })) return true
    draft.append(trace, result.edit)
    editFontRevision.value++
    if (action.kind === 'color' && action.value !== null) rememberColor(action.field, action.value)
    restoreCellInputFocus(session)
    return true
  } catch (error) {
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '无法应用局部文字格式。' })
    return true
  }
}

function runFontFormatCommand(commandId: string, params?: unknown): boolean {
  const sheet = activeSheet.value
  if (!sheet || !canFormatSelection.value) return false
  const action = fontFormatActionForCommand(commandId, params)
  if (!action) return false
  const sparseReset = (action.kind === 'color' || action.kind === 'family' || action.kind === 'size') && action.value === null
  if (!sparseReset && !checkDirectSelection()) return false
  const ranges = selectedModelRanges()
  const [column, row] = splitKey(selectedKey.value)
  const primary = { row: row - 1, col: column }
  let changed: boolean
  if (props.studio && workbook.value.id) {
    // Preserve enum values for font-weight/style commands; the action uses booleans internally.
    const value = params && typeof params === 'object' ? 'value' in params ? params.value : undefined : params
    const input: SheetFontFormatCommandInput = { unitId: workbook.value.id, sheetId: sheet.id, ranges,
      ...(action.kind === 'size-step' ? { primary } : {}),
      ...(value !== undefined ? { value } : {}) }
    changed = props.studio.executeCommandSync<SheetFontFormatCommandInput, boolean>(commandId, input)
    if (changed) {
      const model = props.studio.getWorkbookModel(workbook.value.id)
      if (model) workbook.value = toExcelWorkbookView(model)
    }
  } else {
    changed = applySheetViewFontFormatCommand(sheet, ranges, action, userId.value, primary)
    if (action.kind === 'color' && action.value !== null) rememberColor(action.field, action.value)
  }
  if (changed) commitHistory()
  return true
}

function runLayoutFormatCommand(commandId: string, params?: unknown): boolean {
  const sheet = activeSheet.value
  if (!sheet || !canFormatSelection.value || !checkDirectSelection()) return false
  const action = layoutFormatActionForCommand(commandId, params)
  if (!action) return false
  const ranges = selectedModelRanges()
  let changed = false
  if (props.studio && workbook.value.id) {
    const input: SheetLayoutFormatCommandInput = { unitId: workbook.value.id, sheetId: sheet.id, ranges,
      ...('value' in action ? { value: commandId === 'sheet.command.set-wrap' && action.kind === 'wrap' ? action.value === 'wrap' : action.value } : {}) }
    changed = props.studio.executeCommandSync<SheetLayoutFormatCommandInput, boolean>(commandId, input)
    if (changed) {
      const model = props.studio.getWorkbookModel(workbook.value.id)
      if (model) workbook.value = toExcelWorkbookView(model)
    }
  } else changed = applySheetViewLayoutFormatCommand(sheet, ranges, action, userId.value)
  if (changed) commitHistory()
  return true
}
function runNumberFormatCommand(commandId: string, params?: unknown): boolean {
  const sheet = activeSheet.value
  if (!sheet || !canFormatSelection.value || !checkDirectSelection()) return false
  const action = numberFormatActionForCommand(commandId, params, officeRegion.value)
  if (!action) return false
  const ranges = selectedModelRanges()
  let changed = false
  if (props.studio && workbook.value.id) {
    const input: SheetNumberFormatCommandInput = { unitId: workbook.value.id, sheetId: sheet.id, ranges,
      ...(action.kind === 'set' ? { pattern: action.pattern } : action.kind === 'cells' ? { values: action.values } : {}) }
    changed = props.studio.executeCommandSync<SheetNumberFormatCommandInput, boolean>(commandId, input)
    if (changed) {
      const model = props.studio.getWorkbookModel(workbook.value.id)
      if (model) workbook.value = toExcelWorkbookView(model)
    }
  } else changed = applySheetViewNumberFormatCommand(sheet, ranges, action, officeRegion.value, userId.value)
  if (changed) commitHistory()
  return true
}
function openNumberFormatPanel(): void {
  if (!canFormatSelection.value) return
  numberFormatPreview.value = undefined
  validationPanelOpen.value = false
  customNumberFormatOpen.value = true
}
function closeNumberFormatPanel(): void {
  numberFormatPreview.value = undefined
  customNumberFormatOpen.value = false
}
function previewNumberFormat(pattern: string | undefined): void {
  numberFormatPreview.value = customNumberFormatOpen.value && canFormatSelection.value && pattern !== undefined
    ? { pattern, ranges: selectedRanges.value.map(range => ({ ...range })) } : undefined
}
function applyCustomNumberFormat(pattern: string, complete: (applied: boolean) => void): void {
  try { complete(finishEdit() && runNumberFormatCommand(SHEET_NUMBER_FORMAT_COMMANDS.set, { pattern })) }
  catch (error: unknown) {
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '数字格式操作失败。' })
    complete(false)
  }
}
function resizeSheetSidebar(sizes: number[]): void {
  const size = sizes[1]
  if (sheetSidebarOpen.value && size !== undefined && Number.isFinite(size)) sheetSidebarWidth.value = Math.max(280, Math.min(800, size))
}
function runBorderFormatCommand(commandId: string, params?: unknown): boolean {
  const sheet = activeSheet.value
  if (!sheet || !canFormatSelection.value) return false
  const ranges = selectedModelRanges()
  const input: SheetBorderFormatCommandInput = { ...sheetBorderFormatCommandInput(params), ranges,
    ...(commandId === SHEET_BORDER_FORMAT_COMMANDS.color || commandId === SHEET_BORDER_FORMAT_COMMANDS.style ? { activate: true } : {}) }
  const action = borderFormatActionForCommand(commandId, input, borderFormatState.value)
  if (!action || action.apply && action.state.type !== 'none' && !checkDirectSelection()) return false
  let changed = false, accepted = action.apply
  if (props.studio && workbook.value.id) {
    input.unitId = workbook.value.id
    input.sheetId = sheet.id
    const beforeRevision = props.studio.getWorkbookRevision(input.unitId)
    accepted = props.studio.executeCommandSync<SheetBorderFormatCommandInput, boolean>(commandId, input)
    changed = props.studio.getWorkbookRevision(input.unitId) !== beforeRevision
    borderFormatState.value = props.studio.getSheetBorderFormatState(input.unitId, sheet.id)
    if (changed) {
      const model = props.studio.getWorkbookModel(input.unitId)
      if (model) workbook.value = toExcelWorkbookView(model)
    }
  } else {
    changed = applySheetViewBorderFormatCommand(sheet, ranges, action, userId.value)
    borderFormatState.value = { ...action.state }
    standaloneBorderStates.value.set(`${colorMemoryId.value}\0${sheet.id}`, { ...action.state })
  }
  if (changed) commitHistory()
  return accepted
}

function changeBorderOptions(value: SheetBorderOptions): void {
  if (!canFormatSelection.value) return
  void executeRibbon(SHEET_BORDER_FORMAT_COMMANDS.basic, { value: { type: borderPreset.value, style: value.style, color: value.color } })
}

function convertSelectedTextCells() {
  if (!activeSheet.value || !checkDirectSelection()) return
  const sheet = activeSheet.value
  const keys = selectedExistingKeys()
  const targets = keys.map(key => {
    const [column, row] = splitKey(key)
    return { startRow: row - 1, endRow: row - 1, startCol: column, endCol: column }
  })
  if (!canEditProtectedRanges(sheet, targets, props.studio?.currentUserId ?? 'local-user')) {
    emit('notify', { type: 'warning', message: '选区包含受保护的单元格，未转换数据。' })
    return
  }
  let changed = false
  for (const key of keys) {
    const cell = sheet.cells[key]
    if (!cell) continue
    const converted = convertTextToNumber(cell)
    if (!converted) continue
    sheet.cells[key] = converted
    changed = true
  }
  if (changed) { recalculateFormulas(); commitHistory() }
}

function saveViewSession() {
  if (activeSheetId.value) viewSessions.set(activeSheetId.value, { key: selectedKey.value, anchor: selectionAnchor.value,
    ranges: selectedRanges.value.map(range => ({ ...range })), zoom: zoom.value,
    ...(activeSelectionIndex.value < 0 && selectedRanges.value.length ? {} : { activeRangeIndex: activeSelectionIndex.value }) })
}

function activateSheet(id: string) {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id && candidate.visible !== false)
  if (!sheet) return
  if (!canViewWorksheetProtection(sheet, userId.value)) {
    emit('notify', { type: 'warning', message: '此工作表受保护，当前用户没有查看权限。' })
    return
  }
  if (id === activeSheetId.value) return
  if (!finishEdit()) return
  saveViewSession()
  activeSheetId.value = id
  workbook.value.activeSheetId = id
  nameError.value = ''
  const session = viewSessions.get(id)
  const next = normalizeSheetViewState(sheet, session)
  selectedKey.value = next.key
  selectionAnchor.value = next.anchor
  selectedRanges.value = next.ranges
  activeSelectionIndex.value = next.activeRangeIndex ?? -1
  zoom.value = session?.zoom ?? displayPreference?.value.zoom ?? next.zoom
  if (props.studio && workbook.value.id && props.studio.getWorkbookModel(workbook.value.id)?.sheets.some(candidate => candidate.id === id)) {
    FOffice.newAPI(props.studio).getWorkbook(workbook.value.id)?.setActiveSheet(id)
  }
  syncApiSelection()
}

function runWorksheetLifecycle(command: WorksheetLifecycleCommandId, input: WorksheetLifecycleCommandInput): boolean {
  if (!finishEdit()) return false
  const studio = props.studio, unitId = workbook.value.id, hostSheetId = activeSheetId.value
  try {
    if (studio && unitId) {
      const changed = studio.executeCommandSync<WorksheetLifecycleCommandInput, boolean>(command, {
        ...input, unitId, expectedRevision: input.expectedRevision ?? studio.getWorkbookRevision(unitId),
        expectedHistoryGeneration: input.expectedHistoryGeneration ?? studio.editor.history.captureSavepoint(unitId).generation,
        host: input.host ?? { unitId, sheetId: hostSheetId },
      })
      const latest = studio.getWorkbookModel(unitId)
      if (!changed || !latest || props.studio !== studio || workbook.value.id !== unitId
        || studio.activeWorkbookId !== unitId || studio.editor.state.activeKind !== 'workbook') return false
      const revision = studio.getWorkbookRevision(unitId)
      applyExternalWorkbook(toExcelWorkbookView(latest))
      commitHistory({ prepared: true, syncSelection: false, current: () => props.studio === studio && workbook.value.id === unitId
        && studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook' && studio.getWorkbookRevision(unitId) === revision })
    } else {
      const model = applyExcelWorkbookView(createBlankWorkbook(), workbook.value, officeRegion.value)
      const plan = planWorksheetLifecycle(model, command, input, userId.value)
      if (!plan.changed) return false
      applyExternalWorkbook(toExcelWorkbookView(plan.workbook))
      commitHistory()
    }
    const active = activeSheetId.value
    nextTick(() => {
      if (props.studio === studio && workbook.value.id === unitId && activeSheetId.value === active && !editingKey.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus()
    })
    return true
  } catch (reason) {
    emit('notify', { type: 'error', message: reason instanceof Error ? reason.message : '工作表操作未能提交。' })
    return false
  }
}

function addSheet() {
  runWorksheetLifecycle(WORKSHEET_LIFECYCLE_COMMANDS.insert, { sheet: { id: `sheet-${crypto.randomUUID()}`, rowCount: 100, columnCount } })
}

function copySheet(id: string) {
  if (!finishEdit()) return
  const studio = props.studio, unitId = workbook.value.id, targetSheetId = `sheet-${crypto.randomUUID()}`
  const hostSheetId = activeSheetId.value
  try {
    if (studio && unitId) {
      const changed = studio.executeCommandSync<CopyWorksheetCommandInput, boolean>(COPY_WORKSHEET_COMMAND, {
        unitId, subUnitId: id, targetSheetId, expectedRevision: studio.getWorkbookRevision(unitId), host: { unitId, sheetId: hostSheetId },
      })
      const latest = studio.getWorkbookModel(unitId)
      if (!changed || !latest || props.studio !== studio || workbook.value.id !== unitId
        || studio.activeWorkbookId !== unitId || studio.editor.state.activeKind !== 'workbook') return
      applyExternalWorkbook(toExcelWorkbookView(latest))
      const revision = studio.getWorkbookRevision(unitId)
      commitHistory({ prepared: true, syncSelection: false, current: () => props.studio === studio && workbook.value.id === unitId
        && studio.activeWorkbookId === unitId && studio.editor.state.activeKind === 'workbook' && studio.getWorkbookRevision(unitId) === revision })
    } else {
      const model = applyExcelWorkbookView(createBlankWorkbook(), workbook.value, officeRegion.value)
      const plan = planWorksheetCopy(model, id, { userId: userId.value, targetSheetId })
      applyExternalWorkbook(toExcelWorkbookView(plan.workbook))
      commitHistory()
    }
    nextTick(() => {
      if (workbook.value.id === unitId && activeSheetId.value === targetSheetId && !editingKey.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus()
    })
  } catch (reason) {
    emit('notify', { type: 'error', message: reason instanceof Error ? reason.message : '无法复制工作表。' })
  }
}

function openWorksheetProtection(id: string) {
  if (!workbook.value.sheets.some(sheet => sheet.id === id)) return
  activateSheet(id)
  worksheetProtectionSheetId.value = id
}

function removeWorksheetProtection(id: string) {
  if (!workbook.value.id || !props.studio) return
  try {
    props.studio.executeCommandSync('sheet.command.delete-worksheet-protection', { unitId: workbook.value.id, sheetId: id })
    refreshWorksheetProtection()
  } catch (reason) {
    emit('notify', { type: 'error', message: reason instanceof Error ? reason.message : '无法移除工作表保护。' })
  }
}

function refreshWorksheetProtection() {
  const id = workbook.value.id
  if (!id || !props.studio) return
  const model = props.studio.getWorkbookModel(id)
  if (!model) return
  workbook.value = toExcelWorkbookView(model)
  commitHistory()
}

function setSheetTabColor(id: string, color: string, context: WorksheetLifecycleCommandInput = {}) {
  if (context.unitId !== undefined && context.unitId !== workbook.value.id) return
  runWorksheetLifecycle(WORKSHEET_LIFECYCLE_COMMANDS.tabColor, { ...context, subUnitId: id, value: color })
}

function moveSheet(id: string, order: number) {
  runWorksheetLifecycle(WORKSHEET_LIFECYCLE_COMMANDS.order, { subUnitId: id, order })
}

function switchNextSheet() {
  const sheets = availableSheets.value.filter(sheet => sheet.visible !== false)
  const index = sheets.findIndex(sheet => sheet.id === activeSheetId.value)
  const next = sheets[(index + 1) % sheets.length]
  if (next) activateSheet(next.id)
}

function renameSheet(id: string) {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id)
  if (!sheet) return
  renameSheetId.value = id
  renameValue.value = sheet.name
  renameError.value = ''
}

function confirmRenameSheet() {
  const id = renameSheetId.value
  if (!id) return
  const validation = validateSheetName(renameValue.value, workbook.value.sheets, id)
  if (!validation.valid) { renameError.value = validation.message; return }
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id)
  if (!sheet) return
  const name = renameValue.value.trim()
  if (name === sheet.name || runWorksheetLifecycle(WORKSHEET_LIFECYCLE_COMMANDS.rename, { subUnitId: id, name })) renameSheetId.value = null
}

function deleteSheet(id: string) {
  if (runWorksheetLifecycle(WORKSHEET_LIFECYCLE_COMMANDS.remove, { subUnitId: id })) viewSessions.delete(id)
}

function hideSheet(id: string) {
  runWorksheetLifecycle(WORKSHEET_LIFECYCLE_COMMANDS.hidden, { subUnitId: id })
}

function showSheet(id: string) {
  runWorksheetLifecycle(WORKSHEET_LIFECYCLE_COMMANDS.show, { subUnitId: id })
}

function importWorkbook(file: File): false {
  emit('import', file)
  return false
}

function saveWorkbook() {
  if (!finishEdit()) return
  if (!canReadEntireWorkbook()) return
  emit('save', cloneWorkbook(workbook.value))
}

function exportWorkbook() {
  if (!finishEdit()) return
  if (!canReadEntireWorkbook()) return
  emit('export', cloneWorkbook(workbook.value))
}

function canReadEntireWorkbook(): boolean {
  if (availableSheets.value.length === workbook.value.sheets.length) return true
  emit('notify', { type: 'warning', message: '工作簿包含当前用户无权查看的工作表，无法保存或导出完整内容。' })
  return false
}

function canChangeWorksheet(id: string): boolean {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id)
  if (sheet && canEditWorksheetProtection(sheet, userId.value)) return true
  emit('notify', { type: 'warning', message: '此工作表受保护，当前用户没有编辑权限。' })
  return false
}

function scrollToSelected() {
  spreadsheetCanvas.value?.ensureVisible(selectedKey.value)
}

function startCellSelection(key: string, event: MouseEvent) {
  if (imageUiOpen.value) return
  drawingSelection.set([])
  if (!finishEdit()) return
  resetHeaderSelection()
  const [column, row] = splitKey(key)
  if (event.button === 2 && containsCell(selectedRanges.value, row - 1, column)) return
  focusedChartId.value = null
  selectedSlicerId.value = ''
  selecting.value = event.button === 0
  if (event.ctrlKey || event.metaKey) {
    selectionAnchor.value = key
    selectedKey.value = key
    const point = { startRow: row - 1, endRow: row - 1, startColumn: column, endColumn: column }
    selectedRanges.value = containsCell(selectedRanges.value, row - 1, column) ? selectedRanges.value.flatMap(range => subtractRange(range, point)) : addRange(selectedRanges.value, point)
    if (!selectedRanges.value.length) selectedRanges.value = [point]
  } else selectCell(key, event.shiftKey)
  syncApiSelection()
}

function selectHeader(hit: GridHit, event: PointerEvent) {
  if (imageUiOpen.value) return
  drawingSelection.set([])
  if (!finishEdit()) return
  const sheet = activeSheet.value
  if (!sheet) return
  const axis = hit.kind === 'row' ? 'row' : hit.kind === 'column' ? 'column' : undefined
  const index = hit.kind === 'row' ? hit.row : hit.kind === 'column' ? hit.column : undefined
  if (event.button === 2 && index !== undefined && selectedRanges.value.some(range => axis === 'row'
    ? range.startColumn === 0 && range.endColumn === sheet.columnCount - 1 && index >= range.startRow && index <= range.endRow
    : range.startRow === 0 && range.endRow === sheet.rowCount - 1 && index >= range.startColumn && index <= range.endColumn)) return
  const anchor = axis && event.shiftKey && headerSelectionAnchor?.axis === axis ? headerSelectionAnchor.index : index
  const target = axis && index !== undefined && anchor !== undefined ? sheetAxisHeaderRange(sheet, axis, anchor, index) : undefined
  const range = hit.kind === 'all' ? { startRow: 0, endRow: sheet.rowCount - 1, startColumn: 0, endColumn: sheet.columnCount - 1 }
    : target ? { startRow: target.startRow, endRow: target.endRow, startColumn: target.startCol, endColumn: target.endCol } : undefined
  if (!range) return
  focusedChartId.value = null
  selectedSlicerId.value = ''
  selecting.value = event.button === 0
  headerSelectionBase = event.ctrlKey || event.metaKey ? selectedRanges.value.map(item => ({ ...item })) : []
  headerSelectionAnchor = axis && anchor !== undefined ? { axis, index: anchor } : undefined
  setHeaderSelectionRanges(headerSelectionBase.length ? addRange(headerSelectionBase, range) : [range])
  selectedKey.value = `${columnName(range.startColumn)}${range.startRow + 1}`
  selectionAnchor.value = selectedKey.value
  syncApiSelection()
}

function extendHeaderSelection(axis: 'row' | 'column', index: number): void {
  const sheet = activeSheet.value, anchor = headerSelectionAnchor
  if (!sheet || !selecting.value || !anchor || anchor.axis !== axis) return
  const target = sheetAxisHeaderRange(sheet, axis, anchor.index, index)
  const range = { startRow: target.startRow, endRow: target.endRow, startColumn: target.startCol, endColumn: target.endCol }
  setHeaderSelectionRanges(headerSelectionBase.length ? addRange(headerSelectionBase, range) : [range])
  syncApiSelection()
}

function resetHeaderSelection(): void { headerSelectionAnchor = undefined; headerSelectionBase = [] }
function setHeaderSelectionRanges(ranges: GridRange[]): void {
  applyingHeaderSelection = true
  try { selectedRanges.value = ranges } finally { applyingHeaderSelection = false }
}

function resizeAxis(axis: 'row' | 'column', index: number, size: number) {
  const sheet = activeSheet.value
  if (!sheet) return
  const previous = axis === 'row' ? sheetRowHeight(sheet, index) : sheetColumnWidth(sheet, index)
  void executeRibbon(axis === 'row' ? SHEET_AXIS_DIMENSION_COMMANDS.deltaRowHeight : SHEET_AXIS_DIMENSION_COMMANDS.deltaColumnWidth,
    { sheetId: sheet.id, ...(axis === 'row' ? { anchorRow: index, deltaY: size - previous } : { anchorCol: index, deltaX: size - previous }) })
}

function extendCellSelection(key: string) {
  if (selecting.value) selectCell(key, true)
}

function stopSelecting() {
  selecting.value = false
}
function finishCanvasSelection(event: { canceled: boolean; ranges: GridRange[]; key: string }): void {
  selecting.value = false
  finishFormatPainterSelection(event)
}

function freeze(rows: number, columns: number) {
  if (!activeSheet.value) return
  activeSheet.value.freeze = { rows, columns }
  commitHistory()
}

function openSearch(focusReplacement = false) {
  if (!searchOpen.value) {
    findSelectionRanges.value = selectedRanges.value.map(range => ({ ...range }))
    findSelectionOnly.value = selectedRanges.value.some(range => range.startRow !== range.endRow || range.startColumn !== range.endColumn)
  }
  searchOpen.value = true
  void nextTick(() => (focusReplacement ? replaceInput.value : searchInput.value)?.focus())
}
function focusFindMatch(index: number) {
  const match = searchMatches.value[index]
  if (!match || !finishEdit()) return
  activeSheetId.value = match.sheetId
  selectCell(match.key, false, true)
  spreadsheetCanvas.value?.ensureVisible(match.key)
}
function moveFindMatch(direction: 1 | -1) {
  const sheetId = activeSheetId.value
  const index = nextSheetMatchIndex(searchMatches.value, sheetId, selectedKey.value, direction, findDirection.value, workbook.value.sheets.map(sheet => sheet.id))
  focusFindMatch(index)
}
function findNext() { moveFindMatch(1) }
function findPrevious() { moveFindMatch(-1) }
function applyFindReplace(replaceAll: boolean) {
  const matches = searchMatches.value
  const selectedIndex = matches.findIndex(match => match.sheetId === activeSheetId.value && match.key === selectedKey.value)
  const targetIndex = selectedIndex >= 0 ? selectedIndex : nextSheetMatchIndex(matches, activeSheetId.value, selectedKey.value, 1)
  const patches = createSheetReplacePatches(workbook.value.sheets, matches, {
    text: searchText.value, replacement: replaceText.value, findBy: findBy.value, direction: findDirection.value,
    scope: findScope.value, caseSensitive: findCaseSensitive.value, wholeCell: findWholeCell.value,
  }, replaceAll, targetIndex)
  if (!patches.length) return

  if (props.studio && workbook.value.id) {
    try {
      for (const patch of patches) {
        const [column, row] = splitKey(patch.key)
        props.studio.assertCanEditRange(workbook.value.id, patch.sheetId, {
          startRow: row - 1, endRow: row - 1, startCol: column, endCol: column,
        })
      }
    } catch (error) {
      emit('notify', { type: 'error', message: error instanceof Error ? error.message : t('find.replace-error', 'Replace failed') })
      return
    }
  }

  const candidate = cloneWorkbook(workbook.value)
  const grouped = new Map<string, typeof patches>()
  for (const patch of patches) {
    const sheet = candidate.sheets.find(item => item.id === patch.sheetId)
    if (!sheet) continue
    sheet.cells[patch.key] = patch.cell
    const entries = grouped.get(patch.sheetId) ?? []
    entries.push(patch)
    grouped.set(patch.sheetId, entries)
  }

  const baseModel = props.studio?.getWorkbookModel(workbook.value.id ?? '') ?? createBlankWorkbook()
  const candidateModel = applyExcelWorkbookView(baseModel, candidate, officeRegion.value)
  for (const [sheetId, entries] of grouped) {
    const positions = entries.map(({ key }) => {
      const [column, row] = splitKey(key)
      return { row: row - 1, column }
    })
    const result = validateSheetDataValidationChanges(candidateModel, sheetId, positions)
    const issue = result.issues.find(item => item.severity === 'stop')
    if (issue) {
      emit('notify', { type: 'error', message: `${issue.title}: ${issue.message}` })
      return
    }
    if (result.issues.length) {
      const warning = result.issues[0]!
      emit('notify', { type: warning.severity === 'warning' ? 'warning' : 'info', message: `${warning.title}: ${warning.message}` })
    }
  }

  const previous = workbook.value
  workbook.value = candidate
  try {
    recalculateFormulas()
  } catch (error) {
    workbook.value = previous
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : t('find.replace-error', 'Replace failed') })
    return
  }
  commitHistory()
  const replaced = patches.length
  emit('notify', { type: 'success', message: replaceAll
    ? t('find.replaced-all', `Replaced ${replaced} matches`).replace('{count}', String(replaced))
    : t('find.replaced-one', 'Replaced 1 match') })
  if (replaceAll) moveFindMatch(1)
}
function closeSearch() { searchOpen.value = false; searchText.value = ''; replaceText.value = ''; findSelectionOnly.value = false; findSelectionRanges.value = []; if (!rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() }
function syncFullscreen() { isFullscreen.value = document.fullscreenElement === workbenchRef.value }
async function toggleFullscreen() {
  const workbench = workbenchRef.value
  if (!workbench) return
  try {
    if (document.fullscreenElement === workbench) await document.exitFullscreen()
    else await workbench.requestFullscreen()
  } catch (error) {
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '无法切换全屏' })
  }
}

function openPivotPanel() {
  const bounds = selectionBounds(selectedRanges.value)
  pivotSource.value = bounds ? rangeNotation(bounds) : 'A1:C4'
  pivotPanelOpen.value = true
}

function openPivotChartPanel(params?: Record<string, unknown>): boolean {
  const studio = props.studio, sheetId = activeSheet.value?.id, unitId = workbook.value.id
  if (!studio || !sheetId || !unitId || params?.unitId && params.unitId !== unitId
    || params?.subUnitId && params.subUnitId !== sheetId || params?.sheetId && params.sheetId !== sheetId) return false
  const model = studio.getWorkbookModel(unitId)
  if (!model) return false
  const requestedChartId = typeof params?.pivotChartId === 'string' ? params.pivotChartId : typeof params?.chartId === 'string' ? params.chartId : undefined
  const ref = selectedDrawingRefs.value.length === 1 ? selectedDrawingRefs.value[0] : undefined
  const selectedChart = model.charts?.find(chart => chart.id === (requestedChartId ?? (ref?.kind === 'chart' ? ref.id : undefined)) && chart.sheetId === sheetId && (chart.pivotTableId || chart.ownedPivot))
  if (requestedChartId && !selectedChart) return false
  const initialFieldId = typeof params?.fieldId === 'string' ? params.fieldId : undefined
  if (initialFieldId && (!selectedChart || !FOffice.newAPI(studio).getWorkbook(unitId)?.getPivotChartById(selectedChart.id)?.getFieldSetting(initialFieldId))) return false
  const range = selectedRanges.value.length === 1 ? selectedRanges.value[0] : undefined
  const sourceRange = range ? { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn } : undefined
  const pivotTableId = typeof params?.pivotTableId === 'string' ? params.pivotTableId : model.pivotTables?.find(pivot => {
    const output = pivot.outputRange
    return pivot.anchor.subUnitId === sheetId && sourceRange && output && sourceRange.startRow >= output.startRow && sourceRange.startRow <= output.endRow
      && sourceRange.startCol >= output.startCol && sourceRange.startCol <= output.endCol
  })?.id
  return openPivotChart({ chartId: selectedChart?.id, pivotTableId, initialFieldId,
    sourceKind: params?.mode === 'chartOnly' ? 'range' : pivotTableId ? 'pivotTable' : 'range', sourceSheetId: sheetId, sourceRange })
}

function openPivotChartFieldSettings(chartId: string, fieldId: string): boolean {
  return openPivotChartPanel({ pivotChartId: chartId, fieldId })
}

function openDataPanel(mode: 'filter' | 'sort'): void {
  const bounds = selectionBounds(selectedRanges.value)
  dataPanelMode.value = mode
  dataPanelRange.value = bounds ? rangeNotation(bounds) : selectedKey.value
  dataPanelOpen.value = true
}
function openDataValidationPanel(initialType?: 'checkbox' | 'list' | 'date' | 'time', ruleId?: string): void {
  closeNumberFormatPanel()
  validationPanelRuleId.value = ruleId
  validationPanelRange.value = selectedModelRanges().map(range => formatValidationRange(range)).join(', ') || selectedKey.value
  validationPanelInitialType.value = initialType
  validationPanelOpen.value = true
}

function selectedModelRanges() {
  return selectedRanges.value.map(range => ({
    startRow: range.startRow,
    endRow: range.endRow,
    startCol: range.startColumn,
    endCol: range.endColumn,
  }))
}

function refreshValidationView(): void {
  const model = workbook.value.id && props.studio?.getWorkbookModel(workbook.value.id)
  if (!model) return
  workbook.value = toExcelWorkbookView(model)
  commitHistory()
}

function applyQuickDataValidation(action: unknown): void {
  if (!props.studio || !workbook.value.id || !activeSheet.value) return
  try {
    const plan = planQuickValidation(action === 'dropdown' ? 'edit-dropdown' : action, selectedModelRanges(), createRibbonTranslator(officeLocale.value, 'sheet'))
    if (!plan) return
    const target = { unitId: workbook.value.id, sheetId: activeSheet.value.id }
    const changed = plan.kind === 'set'
      ? props.studio.executeCommandSync('sheet.command.set-data-validation', { ...target, rule: plan.rule })
      : props.studio.executeCommandSync('sheet.command.clear-quick-data-validation', { ...target, ranges: plan.ranges, types: plan.types })
    if (!changed) return
    refreshValidationView()
    if (plan.kind === 'set' && plan.edit) openDataValidationPanel(undefined, plan.rule.id)
  } catch (error) {
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : t('sheet.validation.save-failed', 'Could not apply data validation') })
  }
}
function openConditionalFormattingPanel(): void {
  const sheet = activeSheet.value
  if (!sheet) return
  try {
    if (props.studio) {
      if (!workbook.value.id) throw new Error('找不到目标工作簿。')
      props.studio.assertCanViewWorksheet(workbook.value.id, sheet.id)
    } else if (!canViewWorksheetProtection(sheet, userId.value)) throw new Error('该工作表受保护，当前用户没有查看权限。')
    conditionalFormattingContext.value = { unitId: workbook.value.id, sheetId: sheet.id,
      expectedRevision: props.studio && workbook.value.id ? props.studio.getWorkbookRevision(workbook.value.id) : undefined,
      standaloneSnapshot: props.studio ? undefined : JSON.stringify(workbook.value) }
  } catch (reason) {
    emit('notify', { type: 'error', message: reason instanceof Error ? reason.message : '无法打开条件格式面板。' })
    return
  }
  const bounds = selectionBounds(selectedRanges.value)
  conditionalFormattingRange.value = bounds ? rangeNotation(bounds) : selectedKey.value
  conditionalFormattingOpen.value = true
}
function applyConditionalFormatting(rules: SetSheetConditionalFormatsInput['rules']): SheetConditionalFormatSubmitResult {
  const context = conditionalFormattingContext.value
  try {
    if (!context || context.unitId !== workbook.value.id || context.sheetId !== activeSheetId.value) throw new Error('条件格式目标已切换，请重新打开面板后编辑。')
    let changed: boolean
    if (props.studio) {
      if (!context.unitId) throw new Error('找不到目标工作簿。')
      changed = props.studio.executeCommandSync<SetSheetConditionalFormatsInput, boolean>(SET_SHEET_CONDITIONAL_FORMATS_COMMAND,
        { unitId: context.unitId, sheetId: context.sheetId, rules, expectedRevision: context.expectedRevision })
      if (changed) {
        const model = props.studio.getWorkbookModel(context.unitId)
        if (model) workbook.value = toExcelWorkbookView(model)
      }
    } else {
      if (context.standaloneSnapshot !== JSON.stringify(workbook.value)) throw new Error('工作簿已更改，请重新打开条件格式面板后编辑。')
      const candidate = cloneWorkbook(workbook.value), sheet = candidate.sheets.find(item => item.id === context.sheetId)
      if (!sheet) throw new Error('找不到目标工作表。')
      changed = applySheetConditionalFormatsCommand(sheet, rules, userId.value).changed
      if (changed) workbook.value = candidate
    }
    if (changed) commitHistory()
    return { accepted: true, changed }
  } catch (reason) {
    const message = reason instanceof Error ? reason.message : '无法应用条件格式，草稿已保留。'
    return { accepted: false, error: message.startsWith('Command was canceled:') ? '条件格式操作已取消，草稿已保留。' : message }
  }
}
function onConditionalFormattingClosed(): void {
  if (conditionalFormattingOpen.value) return
  conditionalFormattingContext.value = undefined
  void nextTick(() => {
    if (workbenchActive.value && !conditionalFormattingOpen.value && !editingKey.value && !customNumberFormatOpen.value
      && !validationPanelOpen.value && !hyperlinkDialogOpen.value && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus()
  })
}
function openSparklinePanel(params?: unknown): void {
  const studio = props.studio, sheet = activeSheet.value, unitId = workbook.value.id
  if (!studio || !sheet || !unitId || !workbenchActive.value) return
  if (params !== undefined && (!params || typeof params !== 'object' || Array.isArray(params))) throw new TypeError('迷你图面板上下文无效。')
  const input = params as { sheetId?: string; ranges?: import('../api/events.ts').ApiRange[]; expectedRevision?: number } | undefined
  if (input?.sheetId !== undefined && input.sheetId !== sheet.id || input?.expectedRevision !== undefined && input.expectedRevision !== studio.getWorkbookRevision(unitId)) throw new Error('迷你图右键目标已更改，请重新打开。')
  const ranges = input?.ranges ? input.ranges.map(range => sparklineApiRangeReference(studio, unitId, sheet.id, range, true)) : selectedRanges.value.map(rangeNotation)
  const targets = ranges.length ? ranges : [selectedKey.value]
  sparklineTargetRange.value = targets[0]!
  const context: SparklineWorkbenchContext = { key: ++sparklineContextSequence, studio, workbookId: unitId, sheetId: sheet.id,
    targetRange: targets[0]!, targetRanges: [...targets],
    setOpen(open) { if (sparklineContext.value === context) sparklinePanelOpen.value = open },
    applied() {
      if (sparklineContext.value !== context || props.studio !== studio || !workbenchActive.value || activeSheetId.value !== context.sheetId) return
      const model = studio.getWorkbookModel(unitId)
      if (model) { workbook.value = toExcelWorkbookView(model); commitHistory({ prepared: true, syncSelection: false }) }
    },
    closed(owner) {
      if (sparklineContext.value !== context || sparklinePanelOpen.value) return
      sparklineContext.value = undefined
      if (!owner || !sparklinePanelOwnerCurrent(owner, true)) return
      void nextTick(() => { if (!sparklineContext.value && workbenchActive.value && props.studio === studio && !editingKey.value
        && !imageUiOpen.value && !conditionalFormattingOpen.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value
        && sparklinePanelOwnerCurrent(owner, true)) spreadsheetCanvas.value?.focus() })
    },
  }
  sparklineContext.value = context; sparklinePanelOpen.value = true
}
function removeContextSparklines(params?: unknown): void {
  const studio = props.studio, sheet = activeSheet.value, unitId = workbook.value.id
  if (!studio || !sheet || !unitId || !workbenchActive.value) return
  if (params !== undefined && (!params || typeof params !== 'object' || Array.isArray(params))) throw new TypeError('迷你图删除上下文无效。')
  const input = params as { sheetId?: string; ranges?: import('../api/events.ts').ApiRange[]; expectedRevision?: number; isSingle?: boolean } | undefined
  if (input?.sheetId !== undefined && input.sheetId !== sheet.id) throw new Error('迷你图右键目标已更改。')
  const changed = studio.executeCommandSync(SPARKLINE_PUBLIC_COMMANDS.remove, { targetInfo: { unitId, subUnitId: sheet.id }, ranges: input?.ranges,
    isSingle: input?.isSingle ?? false, expectedRevision: input?.expectedRevision ?? studio.getWorkbookRevision(unitId),
    historyGeneration: studio.editor.history.captureSavepoint(unitId).generation, expectedUserId: studio.currentUserId, host: 'workbench' })
  if (changed) { const model = studio.getWorkbookModel(unitId); if (model) { workbook.value = toExcelWorkbookView(model); commitHistory({ prepared: true, syncSelection: false }) } }
}
function changeOutline(axis: 'row' | 'column', operation: SheetOutlineOperation): void {
  const range = selectedRanges.value.length === 1 ? selectedRanges.value[0] : undefined
  if (!range) return
  const start = axis === 'row' ? range.startRow : range.startColumn
  const end = axis === 'row' ? range.endRow : range.endColumn
  runOutlineCommand({ kind: operation, axis, start, end })
}
function runOutlineCommand(action: SheetOutlineAction, context?: Pick<ChangeSheetOutlineInput, 'expectedRevision' | 'expectedHistoryGeneration' | 'expectedUserId' | 'host'>): void {
  const sheet = activeSheet.value
  if (!sheet) return
  try {
    const unitId = workbook.value.id
    if (props.studio && unitId && props.studio.getWorkbookModel(unitId)) {
      if (!props.studio.executeCommandSync<ChangeSheetOutlineInput, boolean>(CHANGE_SHEET_OUTLINE_COMMAND, { unitId, sheetId: sheet.id, action,
        expectedRevision: context?.expectedRevision ?? props.studio.getWorkbookRevision(unitId),
        expectedHistoryGeneration: context?.expectedHistoryGeneration ?? props.studio.editor.history.captureSavepoint(unitId).generation,
        expectedUserId: context?.expectedUserId ?? userId.value, host: context?.host ?? 'workbench' })) return
      const model = props.studio.getWorkbookModel(unitId)
      if (model) workbook.value = toExcelWorkbookView(model)
    } else {
      const outline = getSheetOutlineUpdate(sheet, action, userId.value)
      if (!outline) return
      const outlineHiddenRowSpans = getSheetOutlineHiddenSpans(outline, 'row', sheet.rowCount)
      const outlineHiddenColumnSpans = getSheetOutlineHiddenSpans(outline, 'column', sheet.columnCount)
      workbook.value = { ...workbook.value, sheets: workbook.value.sheets.map(candidate => candidate.id === sheet.id
        ? { ...candidate, outline, outlineHiddenRowSpans, outlineHiddenColumnSpans, outlineHiddenRows: undefined, outlineHiddenColumns: undefined }
        : candidate) }
    }
    commitHistory()
  } catch (error) {
    emit('notify', { type: 'warning', message: error instanceof Error ? error.message : '大纲操作失败。' })
  }
}
function openFunctionPanel(category?: string): void {
  if (category !== undefined && !SHEET_FUNCTION_CATALOG.some(entry => entry.category === category)) return
  functionCategory.value = category ?? null
  if (category && !SHEET_FUNCTION_CATALOG.some(entry => entry.category === category && entry.name === functionName.value)) {
    const firstFunction = SHEET_FUNCTION_CATALOG.find(entry => entry.category === category)
    if (firstFunction) functionName.value = firstFunction.name
  }
  functionArguments.value = ''
  functionError.value = ''
  functionPanelOpen.value = true
}
function insertFunction(): void {
  try {
    const formula = buildSheetFunctionFormula(functionName.value, functionArguments.value)
    functionPanelOpen.value = false
    beginEdit(selectedKey.value, formula, 'formula')
  } catch (error) { functionError.value = error instanceof Error ? error.message : '无法插入函数。' }
}

async function sortSelection(ascending: boolean, _extendSelection = false) {
  const bounds = selectedRanges.value.length === 1 ? selectedRanges.value[0] : undefined
  const worksheet = getChartWorksheet(activeSheet.value?.id ?? '')
  if (sortBusy.value || !bounds || bounds.endRow <= bounds.startRow || !worksheet) return
  sortController = new AbortController()
  sortBusy.value = true
  sortProgress.value = '正在排序…'
  try {
    await worksheet.sortRange(rangeNotation(bounds), [{ column: bounds.startColumn, direction: ascending ? 'ascending' : 'descending' }], {
      signal: sortController.signal,
      onProgress: progress => { sortProgress.value = progress.phase === 'initializing' ? '正在准备排序…' : `排序：${progress.processed.toLocaleString()} / ${progress.total.toLocaleString()}` },
    })
  } catch (error) {
    if (!sortController.signal.aborted) throw error
  } finally {
    sortBusy.value = false
    sortController = undefined
  }
}

function runAutoFilterCommand(commandId: 'sheet.command.toggle-auto-filter' | 'sheet.command.clear-filter' | 'sheet.command.reapply-filter'): void {
  const sheet = activeSheet.value
  const unitId = workbook.value.id
  if (!props.studio || !unitId || !sheet) return
  const bounds = selectedRanges.value.length === 1 ? selectedRanges.value[0] : undefined
  const range = sheet.autoFilter?.range ?? (bounds && bounds.endRow > bounds.startRow ? {
    startRow: bounds.startRow,
    endRow: bounds.endRow,
    startCol: bounds.startColumn,
    endCol: bounds.endColumn,
  } : undefined)
  if (commandId === 'sheet.command.toggle-auto-filter' && !range) {
    emit('notify', { type: 'info', message: '请选择包含标题行和数据行的区域后再启用筛选。' })
    return
  }
  try {
    props.studio.executeCommandSync(commandId, {
      unitId,
      sheetId: sheet.id,
      ...(commandId === 'sheet.command.toggle-auto-filter' ? { range } : {}),
    })
    const model = props.studio.getWorkbookModel(unitId)
    if (model) workbook.value = toExcelWorkbookView(model)
    commitHistory()
  } catch (error) {
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '无法更新工作表筛选。' })
  }
}

async function executeRibbon(commandId: string, params?: unknown) {
  if (SHEET_IMAGE_COMMAND_IDS.includes(commandId)) { await runImageCommand(commandId, params); return }
  if (resolveSheetClipboardCommand(commandId)) { await runClipboardCommand(commandId, params); return }
  if (SHEET_FILL_COMMAND_IDS.some(id => id === commandId)) { runFillCommand(commandId, params); return }
  if (SHEET_FONT_FORMAT_COMMAND_IDS.includes(commandId) && runInlineFontFormatCommand(commandId, params)) return
  if (commandId !== 'sheet.command.undo' && commandId !== 'sheet.command.redo' && !finishEdit()) return
  if ((SHEET_OUTLINE_COMMAND_IDS as readonly string[]).includes(commandId)) {
    try { runOutlineAlias(commandId, params) } catch (error) { emit('notify', { type: 'warning', message: error instanceof Error ? error.message : '大纲操作失败。' }) }
    return
  }
  const commands: Record<string, () => void | Promise<void>> = {
    'sheet.command.undo': undo, 'sheet.command.redo': redo,
    'sheet.command.set-calculation-mode': () => {
      if ((params !== 'automatic' && params !== 'manual') || !props.studio || !workbook.value.id) return
      const book = FOffice.newAPI(props.studio).getWorkbook(workbook.value.id)
      if (!book) return
      book.setCalculationMode(params)
      const model = props.studio.getWorkbookModel(workbook.value.id)
      if (model) workbook.value = toExcelWorkbookView(model)
      commitHistory()
    },
    'sheet.command.set-number-format': () => { runNumberFormatCommand(SHEET_NUMBER_FORMAT_COMMANDS.set, params) },
    'sheet.command.set-currency-format': () => { runNumberFormatCommand(SHEET_NUMBER_FORMAT_COMMANDS.currency) },
    'sheet.command.set-percent-format': () => { runNumberFormatCommand(SHEET_NUMBER_FORMAT_COMMANDS.percent) },
    'sheet.command.increase-decimals': () => { runNumberFormatCommand(SHEET_NUMBER_FORMAT_COMMANDS.increase) },
    'sheet.command.decrease-decimals': () => { runNumberFormatCommand(SHEET_NUMBER_FORMAT_COMMANDS.decrease) },
    'sheet.command.text-to-number': convertSelectedTextCells,
    'sheet.command.toggle-auto-filter': () => runAutoFilterCommand('sheet.command.toggle-auto-filter'),
    'sheet.command.clear-filter': () => runAutoFilterCommand('sheet.command.clear-filter'),
    'sheet.command.reapply-filter': () => runAutoFilterCommand('sheet.command.reapply-filter'),
    'sheet.operation.insert-chart': openChartDialog,
    'sheet.operation.insert-hyperlink': openHyperlinkDialog,
    'sheet.operation.open-pivot': openPivotPanel,
    'sheet.operation.open-pivot-chart': openPivotChartPanel,
    'sheet.operation.open-pivot-chart-panel': openPivotChartPanel,
    'sheet.operation.open-slicer': () => openSlicerPanel('slicer'),
    'sheet.operation.open-timeline': () => openSlicerPanel('timeline'),
    'sheet.operation.open-slicer-settings-dialog': () => { const id = slicerRibbonIntentId(params); if (id) openSlicerSettings(id) },
    'sheet.operation.open-slicer-connections-dialog': () => { const id = slicerRibbonIntentId(params); if (id) openSlicerConnections(id) },
    'sheet.command.update-slicer-presentation': async () => {
      const id = slicerRibbonIntentId(params)
      if (!id || typeof params !== 'object' || !params || !('patch' in params) || !('expectedRevision' in params)) return
      const patch = params.patch, revision = params.expectedRevision
      if (typeof patch !== 'object' || !patch || Object.keys(patch).some(key => key !== 'style') || !('style' in patch)
        || patch.style !== undefined && typeof patch.style !== 'string' || typeof revision !== 'number' || !Number.isSafeInteger(revision) || revision < 0) return
      await updateSlicerPresentation(id, { style: patch.style }, revision)
    },
    'sheet.operation.open-data-panel': () => openDataPanel(params === 'sort' ? 'sort' : 'filter'),
    'sheet.operation.open-data-validation': openDataValidationPanel,
    'sheet.operation.add-data-validation': () => applyQuickDataValidation('add-default'),
    'sheet.operation.quick-data-validation': () => applyQuickDataValidation(params),
    'sheet.operation.open-conditional-formatting': openConditionalFormattingPanel,
    'sheet.operation.open-sparkline': () => openSparklinePanel(params),
    'sheet.operation.open-sparkline-selector': () => openSparklinePanel(params),
    'sheet.operation.open-sparkline-panel': () => openSparklinePanel(params),
    [SPARKLINE_PUBLIC_COMMANDS.remove]: () => removeContextSparklines(params),
    [OPEN_SHEET_OUTLINE_SETTINGS_OPERATION]: () => { openOutlineSettings() },
    'sheet.operation.open-custom-number-format': openNumberFormatPanel,
    [OPEN_SHEET_RANGE_THEME_OPERATION]: () => openRangeThemeDialog(params),
    'sheet.operation.open.numfmt.panel': openNumberFormatPanel,
    'sheet.operation.close.numfmt.panel': closeNumberFormatPanel,
    'sheet.operation.open-range-protection': () => {
      if (typeof params === 'object' && params !== null && 'sheetId' in params && typeof params.sheetId === 'string') {
        const sheet = workbook.value.sheets.find(candidate => candidate.id === params.sheetId)
        if (sheet) {
          activateSheet(sheet.id)
          if ('range' in params && typeof params.range === 'string') validationPanelRange.value = params.range
          else { const bounds = selectionBounds(selectedRanges.value); if (bounds) validationPanelRange.value = rangeNotation(bounds) }
          protectionStartNew.value = 'mode' in params && params.mode === 'add'
          protectionPanelOpen.value = true
        }
        return
      }
      const bounds = selectionBounds(selectedRanges.value)
      if (bounds) { validationPanelRange.value = rangeNotation(bounds); protectionStartNew.value = false; protectionPanelOpen.value = true }
    },
    'sheet.operation.outline-group-rows': () => changeOutline('row', 'group'),
    'sheet.operation.outline-toggle-rows': () => changeOutline('row', 'toggle'),
    'sheet.operation.outline-remove-rows': () => changeOutline('row', 'remove'),
    'sheet.operation.outline-group-columns': () => changeOutline('column', 'group'),
    'sheet.operation.outline-toggle-columns': () => changeOutline('column', 'toggle'),
    'sheet.operation.outline-remove-columns': () => changeOutline('column', 'remove'),
    'sheet.operation.outline-collapse-rows': () => changeOutline('row', 'collapse'),
    'sheet.operation.outline-expand-rows': () => changeOutline('row', 'expand'),
    'sheet.operation.outline-collapse-columns': () => changeOutline('column', 'collapse'),
    'sheet.operation.outline-expand-columns': () => changeOutline('column', 'expand'),
    'sheet.operation.outline-summary-position': () => {
      if (typeof params !== 'object' || params === null || !('axis' in params) || !('atStart' in params)) return
      const value = params as { axis?: unknown; atStart?: unknown }
      if ((value.axis === 'row' || value.axis === 'column') && typeof value.atStart === 'boolean') runOutlineCommand({ kind: 'summary-position', axis: value.axis, atStart: value.atStart }, sheetOutlineCommandInput(params))
    },
    'sheet.operation.outline-toggle-group': () => {
      if (typeof params !== 'object' || params === null || !('axis' in params) || !('groupId' in params)) return
      const value = params as { axis?: unknown; groupId?: unknown }
      if ((value.axis === 'row' || value.axis === 'column') && typeof value.groupId === 'string') runOutlineCommand({ kind: 'toggle-group', axis: value.axis, groupId: value.groupId }, sheetOutlineCommandInput(params))
    },
    'sheet.operation.outline-set-level': () => {
      if (typeof params !== 'object' || params === null || !('axis' in params) || !('level' in params)) return
      const value = params as { axis?: unknown; level?: unknown }
      if ((value.axis === 'row' || value.axis === 'column') && typeof value.level === 'number' && Number.isSafeInteger(value.level)) runOutlineCommand({ kind: 'set-level', axis: value.axis, level: value.level }, sheetOutlineCommandInput(params))
    },
    'sheet.operation.open-function': () => openFunctionPanel(typeof params === 'string' ? params : undefined),
    'sheet.operation.insert-function': () => {
      if (typeof params !== 'string' || !SHEET_FUNCTION_CATALOG.some(entry => entry.name === params)) return
      try {
        beginEdit(selectedKey.value, buildSheetFunctionFormula(params, ''), 'formula')
      } catch (error) {
        functionError.value = error instanceof Error ? error.message : '无法插入函数。'
      }
    },
    'sheet.command.sort-ascending': () => sortSelection(true, typeof params === 'object' && params !== null && 'extendSelection' in params && params.extendSelection === true),
    'sheet.command.sort-descending': () => sortSelection(false, typeof params === 'object' && params !== null && 'extendSelection' in params && params.extendSelection === true),
    'sheet.command.recalculate': () => { recalculateWorkbook(); commitHistory() },
    'sheet.operation.start-formula': () => { beginEdit(selectedKey.value, activeCell.value.formula ?? '=', 'formula') },
    'sheet.command.toggle-gridlines': () => { if (activeSheet.value) { activeSheet.value.showGridlines = activeSheet.value.showGridlines === false; commitHistory() } },
    'sheet.command.freeze-selection': () => {
      if (typeof params === 'object' && params !== null && 'axis' in params && 'index' in params) {
        const value = params as { axis?: unknown; index?: unknown }
        if (value.axis === 'row' && Number.isSafeInteger(value.index)) freeze(Number(value.index) + 1, 0)
        else if (value.axis === 'column' && Number.isSafeInteger(value.index)) freeze(0, Number(value.index) + 1)
        return
      }
      const range = selectedRanges.value.at(-1)
      if (!range) return
      const counts = frozenPaneCountsForSelection(range)
      freeze(counts.rows, counts.columns)
    },
    'sheet.command.freeze-first-row': () => freeze(1, 0), 'sheet.command.freeze-first-column': () => freeze(0, 1),
    'sheet.command.unfreeze': () => freeze(0, 0),
    'sheet.operation.zoom-in': () => { zoom.value = Math.min(400, zoom.value + 10) },
    'sheet.operation.zoom-out': () => { zoom.value = Math.max(10, zoom.value - 10) },
    'sheet.operation.reset-zoom': () => { zoom.value = 100 },
    'sheet.operation.set-zoom': () => { if (typeof params === 'number' && Number.isFinite(params)) zoom.value = Math.min(400, Math.max(10, params)) },
    'sheet.operation.toggle-fullscreen': toggleFullscreen,
    'sheet.operation.find': () => openSearch(),
    'sheet.operation.find-replace': () => openSearch(true),
    'sheet.operation.find-next': findNext,
    'sheet.operation.find-previous': findPrevious,
    'sheet.operation.replace': () => applyFindReplace(false),
    'sheet.operation.replace-all': () => applyFindReplace(true),
  }
  const command = SHEET_AXIS_DIMENSION_COMMAND_IDS.includes(commandId) ? () => {
    const action = sheetAxisDimensionActionForCommand(commandId), input = sheetAxisDimensionCommandInput(params)
    if (action?.kind === 'set' && input.value === undefined) openAxisSizeDialog(action.axis, input)
    else runAxisDimensionCommand(commandId, input)
  } : (SHEET_AXIS_STRUCTURE_COMMAND_IDS as readonly string[]).includes(commandId) ? () => { runAxisStructureCommand(commandId, params) }
    : SHEET_CELL_SHIFT_COMMAND_IDS.includes(commandId) ? () => { runCellShiftCommand(commandId, params) }
    : commandId === SHEET_RANGE_THEME_COMMANDS.set || commandId === SHEET_RANGE_THEME_COMMANDS.remove ? () => { runRangeThemeCommand(commandId, params) }
    : (SHEET_AXIS_VISIBILITY_COMMAND_IDS as readonly string[]).includes(commandId) ? () => { runAxisVisibilityCommand(commandId, params) }
    : SHEET_FONT_FORMAT_COMMAND_IDS.includes(commandId) ? () => { runFontFormatCommand(commandId, params) }
    : SHEET_LAYOUT_FORMAT_COMMAND_IDS.includes(commandId) ? () => { runLayoutFormatCommand(commandId, params) }
      : SHEET_BORDER_FORMAT_COMMAND_IDS.includes(commandId) ? () => { runBorderFormatCommand(commandId, params) }
        : SHEET_CLEAR_COMMAND_IDS.includes(commandId) ? () => {
          if (runClear(commandId, params) && clearModeForCommand(commandId) !== 'content' && !(props.studio && workbook.value.id)) clearPainterSession()
        } : SHEET_MERGE_COMMAND_IDS.includes(commandId) ? () => { runMerge(commandId, params) }
          : SHEET_FORMAT_PAINTER_COMMAND_IDS.includes(commandId) ? () => { runFormatPainterCommand(commandId, params) } : commands[commandId]
  if (!command) return
  try { await command() } catch (error) {
    emit('notify', { type: error instanceof SheetAxisVisibilityError ? 'warning' : 'error', message: error instanceof Error ? error.message : '操作失败' })
    void nextTick(() => { if (workbenchActive.value && !editingKey.value && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() })
    return
  }
  if ((!commandId.startsWith('sheet.operation.') || commandId.startsWith('sheet.operation.zoom') || commandId === 'sheet.operation.set-zoom'
    || SHEET_FORMAT_PAINTER_COMMAND_IDS.includes(commandId)) && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value) {
    void nextTick(() => { if (workbenchActive.value && !mergeRequest.value && !painterRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() })
  }
}

function handlePainterEscape(event: KeyboardEvent): void {
  if (event.key !== 'Escape' || event.isComposing || event.keyCode === 229 || editComposing.value || editingKey.value
    || !workbenchActive.value || !painterActive.value || mergeRequest.value || painterRequest.value || axisSizePending.value || rangeThemePending.value || cellShiftPending.value) return
  const target = event.target instanceof HTMLElement ? event.target : null
  const painterButton = target?.closest('[data-ribbon-item="sheet.operation.format-painter"], [data-ribbon-item="sheet.operation.continuous-format-painter"]')
  if (!target || !workbenchRef.value?.contains(target) && !painterButton) return
  event.preventDefault()
  cancelFormatPainter()
  if (!painterActive.value) void nextTick().then(() => nextTick(() => {
    if (workbenchActive.value && !painterRequest.value && !mergeRequest.value && !axisSizePending.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus()
  }))
}

function onWindowKeydown(event: KeyboardEvent) {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
    // The application shell owns global shortcuts when it supplies the header.
    if (!props.showHeader) return
    event.preventDefault()
    saveWorkbook()
  }
  if ((event.ctrlKey || event.metaKey) && ['f', 'h'].includes(event.key.toLowerCase()) && document.activeElement?.closest('.excel-workbench')) {
    event.preventDefault()
    void executeRibbon(event.key.toLowerCase() === 'h' ? 'sheet.operation.find-replace' : 'sheet.operation.find')
  }
}

watch(selectedKey, () => {
  if (!applyingShortcutSelection) standaloneSelectionTabStart = undefined
  nameError.value = ''; void nextTick(scrollToSelected)
}, { flush: 'sync' })
watch(selectedRanges, (next, previous) => {
  if (!applyingShortcutSelection && JSON.stringify(next) !== JSON.stringify(previous)) standaloneSelectionTabStart = undefined
  if (!applyingApiSelection && !applyingShortcutSelection) activeSelectionIndex.value = -1
  if (!applyingHeaderSelection) resetHeaderSelection()
}, { flush: 'sync' })
watch([() => workbook.value.id, activeSheetId, () => activeSheet.value?.rowCount, () => activeSheet.value?.columnCount], resetHeaderSelection, { flush: 'sync' })
watch([() => workbook.value.id, activeSheetId], cancelEdit, { flush: 'sync' })
watch([() => workbook.value.id, activeSheetId], () => { validationPanelOpen.value = false })
watch([() => workbook.value.id, activeSheetId], () => { conditionalFormattingOpen.value = false; conditionalFormattingContext.value = undefined }, { flush: 'sync' })
watch([() => workbook.value.id, activeSheetId], closeNumberFormatPanel, { flush: 'sync' })
watch(numberFormatSelectionKey, () => { numberFormatPreview.value = undefined }, { flush: 'sync' })
watch(editingKey, key => { if (key) closeNumberFormatPanel() }, { flush: 'sync' })
watch(canFormatSelection, editable => { if (!editable) closeNumberFormatPanel() }, { flush: 'sync' })
watch([customNumberFormatOpen, validationPanelOpen], async ([numberFormatOpen, validationOpen]) => {
  if (!numberFormatOpen) numberFormatPreview.value = undefined
  if (numberFormatOpen || validationOpen) return
  await nextTick()
  if (workbenchActive.value && !customNumberFormatOpen.value && !validationPanelOpen.value && !editingKey.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus()
})
watch(hyperlinkDialogOpen, open => { if (!open) void nextTick(() => { if (workbenchActive.value && !rangeThemePending.value && !cellShiftPending.value && !imageUiOpen.value) spreadsheetCanvas.value?.focus() }) })
watch(() => workbook.value.id, id => { if (hyperlinkCell.value?.unitId !== id) hyperlinkDialogOpen.value = false })
defineExpose({ save: saveWorkbook, focus: () => spreadsheetCanvas.value?.focus() })

onMounted(() => {
  document.addEventListener('fullscreenchange', syncFullscreen)
  window.addEventListener('mouseup', stopSelecting)
  window.addEventListener('keydown', handlePainterEscape, true)
  window.addEventListener('keydown', onWindowKeydown)
})

onUnmounted(() => {
  workbenchActive.value = false
  cancelMerge()
  clearPainterSession()
  sortController?.abort()
  document.removeEventListener('fullscreenchange', syncFullscreen)
  window.removeEventListener('mouseup', stopSelecting)
  window.removeEventListener('keydown', handlePainterEscape, true)
  window.removeEventListener('keydown', onWindowKeydown)
})
onDeactivated(() => {
  workbenchActive.value = false
  cancelMerge()
  clearPainterSession()
  validationPanelOpen.value = false
  closeNumberFormatPanel()
  sortController?.abort()
  if (document.fullscreenElement === workbenchRef.value) {
    void document.exitFullscreen().catch((error: unknown) => emit('notify', { type: 'error', message: error instanceof Error ? error.message : '无法退出全屏' }))
  }
})
onActivated(() => { workbenchActive.value = true })

</script>

<template>
  <ConfigProvider :direction="officeDirection" :theme="WORKBENCH_CONTROL_THEME">
  <section v-bind="$attrs" ref="workbenchRef" class="excel-workbench" :dir="officeDirection" @keydown="handleGridKeydown">
    <header v-if="showHeader && chrome.header" class="excel-header">
      <div class="document-brand">
        <div class="excel-logo"><FileSpreadsheet :size="18" /></div>
        <div class="document-title-wrap">
          <div class="document-title">{{ workbook.name }}</div>
          <div class="document-status">本地工作簿</div>
        </div>
      </div>
      <div class="header-actions">
        <Upload accept=".xlsx,.xls,.xlsb,.ods,.csv,.tsv" :show-upload-list="false" :before-upload="importWorkbook"><Tooltip title="打开文件"><Button type="text" size="small" aria-label="打开文件"><FolderOpen :size="16" /></Button></Tooltip></Upload>
        <Tooltip title="保存"><Button type="text" size="small" aria-label="保存工作簿" @click="saveWorkbook"><Save :size="16" /></Button></Tooltip>
        <Tooltip title="导出"><Button type="text" size="small" aria-label="导出工作簿" @click="exportWorkbook"><Download :size="16" /></Button></Tooltip>
      </div>
    </header>

    <WorkbenchRibbon v-if="chrome.toolbar" v-model="activeTab" :data-ribbon-editor-owner="ribbonEditorOwner" :tabs="ribbonTabs" :variant="props.ribbonVariant" @execute="executeRibbon" />
    <SheetFormulaBar v-if="chrome.header" ref="formulaInput" :address="selectedAddress" :value="formulaValue" :editing="Boolean(editingKey)" :edit-session-id="editSession?.sessionId" :input-context="cellInputContext" :composing="editComposing" :error="nameError" @input-event="handleCellNativeInput" @format-command="executeRibbon" @format-keydown="handleInlineFormatHistoryKey" @begin="beginFormulaEdit" @confirm="finishFormulaEdit" @cancel="cancelFormulaEdit" @navigate="navigateToAddress" @focus-grid="nameError = ''; spreadsheetCanvas?.focus()" />
    <div v-if="sortBusy" class="sheet-find-panel" role="status"><span>{{ sortProgress }}</span><Button size="small" aria-label="取消排序" @click="sortController?.abort()">取消</Button></div>
    <div v-if="searchOpen" class="sheet-find-panel" role="search" @keydown.stop @keydown.esc.prevent="closeSearch">
      <div class="find-input-row">
        <Search :size="15" />
        <Input ref="searchInput" v-model:value="searchText" size="small" :placeholder="t('find.find-placeholder', 'Find')" :aria-label="t('find.find', 'Find')" @press-enter="executeRibbon('sheet.operation.find-next')" />
        <span class="search-count">{{ searchText ? t('find.match-count', '{count} matches').replace('{count}', String(searchMatches.length)) : t('find.enter-query', 'Enter search text') }}</span>
        <Button size="small" :disabled="!searchMatches.length" :title="t('find.previous-match', 'Previous match')" :aria-label="t('find.previous-match', 'Previous match')" @click="executeRibbon('sheet.operation.find-previous')"><ArrowUp :size="14" /></Button>
        <Button size="small" :disabled="!searchMatches.length" :title="t('find.next-match', 'Next match')" :aria-label="t('find.next-match', 'Next match')" @click="executeRibbon('sheet.operation.find-next')"><ArrowDown :size="14" /></Button>
        <Button type="text" size="small" :aria-label="t('find.close', 'Close find and replace')" @click="closeSearch"><X :size="15" /></Button>
      </div>
      <div class="find-options-row">
        <Input ref="replaceInput" v-model:value="replaceText" size="small" :placeholder="t('find.replace-placeholder', 'Replace with')" :aria-label="t('find.replace', 'Replace')" @press-enter="executeRibbon('sheet.operation.replace')" />
        <Button size="small" :disabled="!searchMatches.some(match => match.replaceable)" @click="executeRibbon('sheet.operation.replace')"><Replace :size="14" />{{ t('find.replace', 'Replace') }}</Button>
        <Button size="small" :disabled="!searchMatches.some(match => match.replaceable)" @click="executeRibbon('sheet.operation.replace-all')"><ListRestart :size="14" />{{ t('find.replace-all', 'Replace All') }}</Button>
        <Select v-model:value="findScope" size="small" :aria-label="t('find.scope', 'Find range')" :options="[{ value: 'sheet', label: t('find.current-sheet', 'Current Sheet') }, { value: 'workbook', label: t('find.workbook', 'Workbook') }]" />
        <Select v-model:value="findDirection" size="small" :aria-label="t('find.direction', 'Find direction')" :options="[{ value: 'row', label: t('find.by-row', 'By row') }, { value: 'column', label: t('find.by-column', 'By column') }]" />
        <Select v-model:value="findBy" size="small" :aria-label="t('find.find-by', 'Find by')" :options="[{ value: 'value', label: t('find.value', 'Value') }, { value: 'formula', label: t('find.formula', 'Formula') }]" />
        <Checkbox v-model:checked="findCaseSensitive">{{ t('find.case-sensitive', 'Match case') }}</Checkbox>
        <Checkbox v-model:checked="findWholeCell">{{ t('find.whole-cell', 'Match entire cell') }}</Checkbox>
      </div>
    </div>

    <Splitter class="sheet-canvas-layout" :class="{ 'has-sidebar': sheetSidebarOpen }" :styles="sheetSplitterStyles" orientation="horizontal" @resize="resizeSheetSidebar">
      <SplitterPanel class="sheet-canvas-stack" :min="0" :resizable="sheetSidebarOpen" :destroy-on-hidden="false">
    <SpreadsheetCanvas v-if="activeRenderSheet" ref="spreadsheetCanvas" :edit-value="editValue" :edit-session-id="editSession?.sessionId" :input-context="cellInputContext" :keep-edit-focus="keepCellEditFocus" :prepare-fill="prepareCanvasFillGesture"
      :rich-text-editor-cell="richTextEditorCell" :edit-selection="editTextSelection"
      :sheet="activeRenderSheet" :sparkline-sheets="availableSheets" :selected-key="selectedKey" :selected-range="emptyCellSelection" :selected-ranges="selectedRanges"
      :editing-key="editingKey" :zoom="zoom" :current-user-id="studio?.currentUserId ?? 'local-user'" :workbook-id="workbook.id"
      :workbook-revision="studio && workbook.id ? studio.getWorkbookRevision(workbook.id) : undefined"
      :workbook-history-generation="studio && workbook.id ? studio.editor.history.captureSavepoint(workbook.id).generation : undefined"
      :can-format="canFormatSelection" :format-painter="formatPainterCanvas" :search-text="searchText" :number-format-preview="numberFormatPreview"
      :copy-range="clipboardCopyRange" :paste-options="clipboardPasteOptions" :fill-options="sheetFillOptions" :clipboard-capabilities="clipboardCapabilities"
      :paste-options-capabilities="clipboardPasteOptionsCapabilities"
      :validation-options="validationOptions" :overlay-commands="chartCommands" :slicers="slicerControls" :selected-slicer-id="selectedSlicerId"
      :drawing-scene="drawingScene" :selected-drawing-refs="selectedDrawingRefs"
      @select-drawing="selectDrawing" @drawing-command="runDrawingCommand" @chart-event="routeCanvasChartEvent"
      :selected-image-id="selectedImageId" :selected-image-ids="selectedImageIds" :image-interaction-blocked="imageUiOpen || outlineUiOpen || Boolean(rangeThemePending || cellShiftPending || axisSizePending || mergeRequest || painterRequest)"
      @select-image="selectImage" @edit-image="openImageEditor" @save-image="saveImage" @image-command="runImageCommand"
      @select-cell="startCellSelection" @extend-selection="extendCellSelection" @extend-header-selection="extendHeaderSelection"
      @selection-end="finishCanvasSelection" @fill-selection="fillSelection" @refill-selection="refillSelection" @select-header="selectHeader" @edit-cell="beginEdit" @finish-edit="finishEdit"
      @input-event="handleCellNativeInput" @editor-keydown="handleGridKeydown" @copy="handleSheetCopy" @cut="event => handleSheetCopy(event, true)" @paste="handleSheetPaste" @resize-axis="resizeAxis"
      @context-command="executeSheetContextCommand" @open-hyperlink="openCellHyperlink" @select-slicer="selectSlicer"
      @slicer-item="handleSlicerItem" @slicer-clear="clearSlicer" @slicer-range="setSlicerTimelineRange" @slicer-scroll-to="scrollSlicerTo"
      @slicer-window="requestSlicerWindow" @slicer-retry="slicerCanvasData.retry"
      @slicer-level="({ slicerId, level, expectedRevision }) => updateSlicerPresentation(slicerId, { level }, expectedRevision)" @slicer-settings="openSlicerSettings" @slicer-connections="openSlicerConnections"
      @slicer-move="moveSlicer" @slicer-resize="resizeSlicer" @slicer-geometry="transformSlicer" />
        <ChartOverlay v-if="activeSheet" ref="chartOverlay" :studio="props.studio" :workbook="chartWorkbook" :active-sheet-id="activeSheet.id" :zoom="zoom" :active="workbenchActive" :header-width="sheetViewportHeaders.width" :header-height="sheetViewportHeaders.height" :restore-focus-allowed="!rangeThemePending && !cellShiftPending && !imageUiOpen && !outlineUiOpen && !pivotChartPanelOpen && !chartEditorOpen && !chartModalOpen" :open-chart-editor="openFocusedChartPanel" :open-pivot-chart-settings="openPivotChartFieldSettings"
          :selected-drawing-refs="selectedDrawingRefs" :drawing-order="drawingOrder" :drawing-visible="canViewDrawing" :route-drawing-event="routeChartDrawingEvent"
          @drawing-command="runDrawingCommand"
          @select-chart="selectChart" @object-change="handleChartObjectChange"
          @overlay-commands="chartCommands = $event" @update-chart="updateChartLayout" @remove-chart="removeChart" />
      </SplitterPanel>
      <SplitterPanel class="sheet-context-panel" :size="sheetSidebarOpen ? sheetSidebarWidth : 0" :min="sheetSidebarOpen ? 280 : 0" :max="800" :resizable="sheetSidebarOpen">
        <CustomNumberFormatPanel v-if="customNumberFormatOpen" v-model:open="customNumberFormatOpen" :disabled="!canFormatSelection" :date1904="activeDate1904" :current-pattern="numberFormatSource.pattern" :sample-value="numberFormatSource.sample" :selection-key="numberFormatSelectionKey" @preview="previewNumberFormat" @apply="applyCustomNumberFormat" @preference-error="(message) => emit('notify', { type: 'warning', message })" />
        <SheetDataValidationPanel v-else-if="validationPanelOpen && activeSheet && studio && workbook.id" v-model:open="validationPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="validationPanelRange" :initial-type="validationPanelInitialType" :initial-rule-id="validationPanelRuleId" :canvas-ranges="selectedModelRanges()" :can-add="canEditValidationSelection" @add="applyQuickDataValidation('add-default')" @applied="refreshValidationView" />
      </SplitterPanel>
    </Splitter>
    <Alert v-if="drawingError" type="error" :title="drawingError" show-icon closable @close="drawingError = ''" />

    <ChartEditorPanel v-if="chartEditorContext" :key="chartEditorContext.key" :open="chartEditorOpen" :chart="chartEditorChart" :sheet="chartEditorSheet" :sheets="chartEditorSheets" :source-revision="chartSourceRevision" :load-source="chartEditorContext.loadSource" :load-statistics="chartEditorContext.loadStatistics" :route="chartEditorRoute" :route-key="chartEditorRouteKey" :navigation-current="chartEditorContext.navigationCurrent" :date1904="activeDate1904" :busy="chartEditorBusy" :error="chartEditorError" @update:open="chartEditorContext.setOpen" @apply="chartEditorContext.apply" @closed="chartEditorContext.closed" />
    <ChartInsertDialog v-if="chartInsertContext" :key="chartInsertContext.key" :open="chartModalOpen" :config="chartInsertContext.config" :sheets="chartInsertSheets" :host-sheet-id="chartInsertContext.sheetId" :source-revision="chartSourceRevision" :load-source="chartInsertContext.loadSource" :load-statistics="chartInsertContext.loadStatistics" :date1904="activeDate1904" :busy="chartInsertBusy" :error="chartError" @update:open="chartInsertContext.setOpen" @apply="chartInsertContext.apply" @closed="chartInsertContext.closed" />

    <Modal v-model:open="hyperlinkDialogOpen" :title="editingHyperlink ? '编辑超链接' : '插入超链接'" ok-text="确定" cancel-text="取消" :width="440" @ok="confirmHyperlinkDialog">
      <Form layout="vertical" @submit.prevent>
        <FormItem label="链接地址" :validate-status="hyperlinkError ? 'error' : undefined" :help="hyperlinkError || undefined">
          <Input v-model:value="hyperlinkTarget" aria-label="链接地址" placeholder="https://example.com" allow-clear @press-enter="confirmHyperlinkDialog" />
        </FormItem>
        <FormItem label="显示文字">
          <Input v-model:value="hyperlinkDisplay" aria-label="显示文字" placeholder="留空时保留单元格内容" @press-enter="confirmHyperlinkDialog" />
        </FormItem>
      </Form>
      <div class="hyperlink-actions">
        <Button v-if="editingHyperlink" danger @click="deleteSelectedHyperlink">移除超链接</Button>
        <Button v-if="editingHyperlink" type="text" @click="openSelectedHyperlink">打开链接</Button>
      </div>
    </Modal>

    <Modal :open="Boolean(renameSheetId)" title="重命名工作表" ok-text="确定" cancel-text="取消" :width="360" @ok="confirmRenameSheet" @cancel="renameSheetId = null">
      <Form layout="vertical" @submit.prevent>
        <FormItem label="工作表名称" :validate-status="renameError ? 'error' : undefined">
          <Input v-model:value="renameValue" aria-label="工作表名称" :status="renameError ? 'error' : undefined" :maxlength="31" @press-enter="confirmRenameSheet" />
          <template #help><span v-if="renameError" role="alert">{{ renameError }}</span></template>
        </FormItem>
      </Form>
    </Modal>
    <Modal v-model:open="functionPanelOpen" title="插入函数" ok-text="插入到公式栏" cancel-text="取消" :width="480" @ok="insertFunction">
      <Form layout="vertical" class="chart-insert-form" @submit.prevent>
        <FormItem label="函数"><Select v-model:value="functionName" :options="functionOptions" aria-label="函数名称" show-search option-filter-prop="label" /></FormItem>
        <p>{{ functionEntry?.category }} · {{ functionEntry?.label }}</p>
        <FormItem label="参数或区域"><Input v-model:value="functionArguments" aria-label="函数参数" :placeholder="functionEntry?.example" @press-enter="insertFunction" /></FormItem>
        <p class="chart-insert-hint">例如 {{ functionEntry?.example }}。插入后可在公式栏修改并确认。目录列出当前引擎支持的函数。</p>
        <p v-if="functionError" role="alert" class="chart-insert-error">{{ functionError }}</p>
      </Form>
    </Modal>
    <SheetAxisSizeDialog :request="axisSizeRequest" :error="axisSizeError" @confirm="confirmAxisSize" @cancel="cancelAxisSize" @closed="afterAxisSizeClose" />
    <SheetRangeThemeDialog :request="rangeThemeRequest" :error="rangeThemeError" @confirm="confirmRangeTheme" @cancel="cancelRangeTheme" @closed="afterRangeThemeClose" />
    <SheetImagePickerDialog :request="imagePickerRequest" :busy="imageBusy" :error="imageError" :progress="imageProgress" @confirm="confirmImagePicker" @cancel="cancelImagePicker" @closed="afterImageClose" />
    <SheetImageEditorDialog :request="imageEditorRequest" :busy="imageBusy" :error="imageError" @confirm="confirmImageEditor" @cancel="cancelImageEditor" @closed="afterImageClose" />
    <SheetImageSaveDialog :request="imageSaveRequest" :busy="imageBusy" :error="imageError" @confirm="confirmImageSave" @cancel="cancelImageSave" @closed="afterImageClose" />
    <SheetCellShiftConfirmationDialog :request="cellShiftRequest" :error="cellShiftError" @confirm="confirmCellShift" @cancel="cancelCellShift" @closed="afterCellShiftClose" />
    <SheetMergeConfirmationDialog :request="mergeRequest" :error="mergeError" @confirm="confirmMerge" @cancel="cancelMerge" @closed="afterMergeClose" />
    <SheetMergeConfirmationDialog :request="painterRequest" :error="painterError" @confirm="confirmPainter" @cancel="cancelPainterConfirmation" @closed="afterPainterClose" />
    <SheetFooter v-if="chrome.footer" :sheets="availableSheets" :sheet-order="workbook.sheets.map(sheet => sheet.id)" :workbook-id="workbook.id" :revision="studio?.getWorkbookRevision(workbook.id)" :history-generation="studio?.editor.history.captureSavepoint(workbook.id).generation" :active-sheet-id="activeSheetId" :zoom="zoom" :statistics="selectedStatistics" :number-format="activeCell.rawNumberFormat ?? activeCell.numberFormat" @activate="activateSheet" @add="addSheet" @rename="renameSheet" @copy="copySheet" @tab-color="setSheetTabColor" @move="moveSheet" @protect-sheet="openWorksheetProtection" @unprotect-sheet="removeWorksheetProtection" @delete="deleteSheet" @hide="hideSheet" @show="showSheet" @update:zoom="zoom = $event" />
    <PivotTablePanel v-if="pivotPanelOpen && activeSheet" v-model:open="pivotPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="pivotSource" />
    <PivotChartPanel v-if="pivotChartContext" :key="pivotChartContext.key" :open="pivotChartPanelOpen" :studio="pivotChartContext.studio" :workbook-id="pivotChartContext.unitId" :sheet-id="pivotChartContext.sheetId"
      :chart-id="pivotChartContext.chartId" :pivot-table-id="pivotChartContext.pivotTableId" :source-kind="pivotChartContext.sourceKind" :source-sheet-id="pivotChartContext.sourceSheetId" :source-range="pivotChartContext.sourceRange" :initial-field-id="pivotChartContext.initialFieldId"
      @update:open="pivotChartContext.setOpen" @busy-change="pivotChartContext.setBusy" @owner-change="pivotChartContext.ownerChanged" @applied="pivotChartContext.applied" @closed="pivotChartContext.closed" />
    <SlicerPanel v-if="slicerPanelOpen && activeSheet && studio && workbook.id" v-model:open="slicerPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="slicerSourceRange" :initial-kind="slicerKind" :initial-section="slicerPanelSection" :selected-slicer-id="selectedSlicerId" @changed="handleSlicerPanelChanged" @closed="handleSlicerPanelClosed" />
    <SheetDataPanel v-if="dataPanelOpen && activeSheet && studio && workbook.id" v-model:open="dataPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="dataPanelRange" :initial-mode="dataPanelMode" />
    <ConditionalFormattingPanel v-if="conditionalFormattingContext && activeSheet" v-model:open="conditionalFormattingOpen" :rules="activeSheet.conditionalFormats ?? []" :default-range="conditionalFormattingRange" :submit="applyConditionalFormatting" @closed="onConditionalFormattingClosed" />
    <SheetRangeProtectionPanel v-if="protectionPanelOpen && activeSheet && studio && workbook.id" v-model:open="protectionPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="validationPanelRange" :start-new="protectionStartNew" />
    <SheetProtectionPanel v-if="worksheetProtectionSheetId && studio && workbook.id" :open="Boolean(worksheetProtectionSheetId)" :studio="studio" :workbook-id="workbook.id" :sheet-id="worksheetProtectionSheetId" :rule="workbook.sheets.find(sheet => sheet.id === worksheetProtectionSheetId)?.worksheetProtection" @update:open="open => { if (!open) worksheetProtectionSheetId = null }" @applied="refreshWorksheetProtection" />
    <SheetOutlineActionDialog v-if="outlineActionRequest" :key="outlineActionRequest.key" :open="outlineActionOpen" :studio="outlineActionRequest.studio" :workbook-id="outlineActionRequest.workbookId" :sheet-id="outlineActionRequest.sheetId" :ranges="outlineActionRequest.ranges" :operation="outlineActionRequest.operation"
      @update:open="outlineActionRequest.setOpen" @changed="outlineActionRequest.changed" />
    <SheetOutlineSettingsDialog v-if="outlineSettingsRequest" :key="outlineSettingsRequest.key" :open="outlineSettingsOpen" :studio="outlineSettingsRequest.studio" :workbook-id="outlineSettingsRequest.workbookId" :sheet-id="outlineSettingsRequest.sheetId"
      @update:open="outlineSettingsRequest.setOpen" @changed="outlineSettingsRequest.changed" />
    <SparklinePanel v-if="sparklineContext" :key="sparklineContext.key" :open="sparklinePanelOpen" :studio="sparklineContext.studio" :workbook-id="sparklineContext.workbookId" :sheet-id="sparklineContext.sheetId" :target-range="sparklineContext.targetRange" :target-ranges="sparklineContext.targetRanges"
      @update:open="sparklineContext.setOpen" @applied="sparklineContext.applied" @closed="sparklineContext.closed" />
  </section>
  </ConfigProvider>
</template>

<style scoped>
.excel-workbench { --ink: #172033; --muted: #718096; --line: #dfe5ee; --blue: #246bfd; display: flex; flex: 1; min-width: 0; min-height: 0; flex-direction: column; overflow: hidden; color: var(--ink); background: #fff; font-family: Inter, "Segoe UI", "Microsoft YaHei", sans-serif; font-size: 13px; }
.excel-header { display: flex; flex: 0 0 53px; align-items: center; justify-content: space-between; border-bottom: 1px solid var(--line); padding: 0 18px; background: #fff; }
.document-brand, .header-actions { display: flex; align-items: center; gap: 10px; }
.excel-logo { display: grid; width: 30px; height: 30px; place-items: center; color: #fff; border-radius: 6px; background: #1f9d68; }
.document-title-wrap { display: grid; gap: 1px; }
.document-title { color: #1c2535; font-size: 14px; font-weight: 600; line-height: 18px; }
.document-status { color: #8994a6; font-size: 11px; }
.header-actions { gap: 3px; }
.sheet-canvas-layout { display: flex; flex: 1; min-width: 0; min-height: 0; overflow: hidden; }
.sheet-canvas-layout:not(.has-sidebar) :deep(.ant-splitter-bar) { display: none; }
/* Splitter recreates Panel vnodes, so panel layout must target its generated DOM. */
.sheet-canvas-layout :deep(.sheet-canvas-stack) { position: relative; display: flex; height: 100%; min-width: 0; min-height: 0; overflow: hidden; }
.sheet-canvas-layout :deep(.sheet-canvas-stack .spreadsheet-canvas) { flex: 1; }
.sheet-canvas-layout :deep(.sheet-context-panel) { display: flex; height: 100%; min-width: 0; min-height: 0; overflow: hidden; }
.sheet-canvas-layout :deep(.sheet-context-panel .validation-sidebar) { width: 100%; min-width: 0; flex: 1; }
.chart-insert-form { display: flex; flex-direction: column; gap: 14px; padding: 12px 0 4px; }
.chart-insert-form :deep(.ant-form-item) { margin-bottom: 0; }
.hyperlink-actions { display: flex; justify-content: space-between; gap: 8px; }
.chart-insert-error { margin: 8px 0 0; color: #c2413a; }
.chart-insert-hint { margin: 0; color: #8490a0; font-size: 11px; }
.sheet-find-panel { display: flex; flex: 0 0 auto; flex-direction: column; gap: 5px; box-sizing: border-box; min-width: 0; border-bottom: 1px solid var(--line); padding: 6px 12px; color: #526071; background: #f8faff; }
.find-input-row, .find-options-row { display: flex; min-width: 0; align-items: center; gap: 6px; }
.find-input-row :deep(.ant-input) { width: min(300px, 32vw); min-width: 100px; }
.find-options-row :deep(.ant-input) { width: min(240px, 25vw); min-width: 100px; }
.find-options-row :deep(.ant-select) { min-width: 105px; }
.find-options-row :deep(.ant-checkbox-wrapper) { flex-shrink: 0; font-size: 12px; }
.search-count { flex-shrink: 0; font-size: 12px; }
@media (max-width: 760px) { .find-options-row { flex-wrap: wrap; } .find-options-row :deep(.ant-input) { flex: 1 1 160px; width: auto; } }
@media (max-width: 600px) { .sheet-find-panel { gap: 6px; padding: 6px; } .find-input-row :deep(.ant-input) { flex: 1; width: auto; } .find-options-row :deep(.ant-select) { flex: 1 1 105px; } .search-count { display: none; } }
</style>
