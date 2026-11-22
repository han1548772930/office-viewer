<script setup lang="ts">
import { computed, defineAsyncComponent, inject, nextTick, onDeactivated, onMounted, onUnmounted, ref, watch } from 'vue'
import { Button, Checkbox, ConfigProvider, Form, FormItem, Input, InputNumber, Modal, Select, Tooltip, Upload } from 'antdv-next'
import { FOffice } from '../api/facade'
import type { OfficeStudio } from '../api/office-studio'
import type { ExcelCellView, ExcelWorkbookView as DomainExcelWorkbookView } from '../services/excelViewAdapter'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../services/excelViewAdapter'
import { createBlankWorkbook } from '../domain/model'
import type { ChartModel, ChartSourceOrientation, ChartType } from '../domain/model'
import type { SlicerModel } from '../domain/model'
import SpreadsheetCanvas from './spreadsheet/SpreadsheetCanvas.vue'
import ChartOverlay from './ChartOverlay.vue'
import { chartPositionForRange, type ExcelWorkbookChartView } from './chart-overlay-model'
import type { DrawCommand } from '../rendering/contracts'
import { SpreadsheetGeometry, type GridHit, type GridRange, type SpreadsheetRenderSheet } from '../rendering/spreadsheet/layout'
import { addRange, containsCell, selectionBounds, selectionCount, subtractRange } from '../rendering/spreadsheet/selection'
import { ArrowDown, ArrowUp, Download, FileSpreadsheet, FolderOpen, ListRestart, Replace, Save, Search, X } from 'lucide-vue-next'
import WorkbenchRibbon from '../platform/ribbon/WorkbenchRibbon.vue'
import type { RibbonVariant } from '../platform/ribbon/ribbon-schema'
import { useWorkbenchDisplay, workbenchVisibility } from '../platform/workbench-display'
import SheetFormulaBar from './spreadsheet/SheetFormulaBar.vue'
import SheetFooter from './spreadsheet/SheetFooter.vue'
import { createSheetRibbon } from './spreadsheet/sheet-ribbon'
import { calculateSelectionStatistics, createUniqueSheetName, frozenPaneCountsForSelection, normalizeSheetViewState, parseSheetAddress, validateSheetName, type SheetViewState } from '../features/sheet-workbench/sheet-workbench'
import { recalculateSheetViewFormulas } from '../features/sheet-workbench/sheet-view-formulas'
import { adjustSheetDecimals, applySheetBorders, applySheetNumberFormat, clearSheetCellFormat, isSheetBorderPreset, nextSheetFontSize, setSheetTextFit, sheetBorderStyleLabels, type SheetBorderOptions, type SheetBorderPreset } from '../features/sheet-workbench/sheet-toolbar-format'
import { classifySheetNumberFormat, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from '../features/sheet-workbench/sheet-number-format'
import { buildSheetFunctionFormula, SHEET_FUNCTION_CATALOG } from '../features/sheet-workbench/sheet-function-catalog'
import { findVisibleSheetRow } from '../features/sheet-workbench/sheet-row-visibility'
import { validateSheetDataValidationChanges } from '../features/sheet-validation/validation-engine.ts'
import { planQuickValidation } from '../features/sheet-validation/quick-validation.ts'
import { formatValidationRange } from '../features/sheet-validation/validation-panel-state.ts'
import { createSheetPastePlan, type SheetPasteMode } from '../features/sheet-workbench/sheet-paste.ts'
import { createSheetFillPlan, type SheetFillApplyType } from '../features/sheet-workbench/sheet-fill.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewWorksheetProtection } from '../features/sheet-protection/range-protection.ts'
import { createFormatPainterPatches, type FormatPainterSource } from '../features/sheet-workbench/sheet-format-painter.ts'
import { CHANGE_SHEET_OUTLINE_COMMAND, getCollapsedOutlineIndexes, getSheetOutlineUpdate, type ChangeSheetOutlineInput, type SheetOutlineAction, type SheetOutlineOperation } from '../features/sheet-outline/sheet-outline.ts'
import { createOfficeTranslator, createRibbonTranslator, provideWorkbenchLocales } from '../i18n/office-i18n.ts'
import { createSheetReplacePatches, findSheetMatches, nextSheetMatchIndex, type SheetFindBy, type SheetFindDirection, type SheetFindScope } from '../features/sheet-find-replace/sheet-find-replace.ts'
import { convertTextToNumber } from '../features/sheet-workbench/text-to-number.ts'
import { moveSheetCellRange, structureEdit, transformWorkbookReferencesForSheetEdit, type SheetCellMoveDirection } from '../features/sheet-workbench/sheet-cell-range-move.ts'
import { resolveSlicerBounds } from '../features/slicer/slicer-drawing.ts'
import { canSetSheetHyperlink, normalizeExternalSheetHyperlink, parseSheetHyperlinkTarget, SET_SHEET_HYPERLINK_COMMAND, REMOVE_SHEET_HYPERLINK_COMMAND } from '../features/sheet-workbench/sheet-hyperlink.ts'

const PivotTablePanel = defineAsyncComponent(() => import('./spreadsheet/PivotTablePanel.vue'))
const PivotChartPanel = defineAsyncComponent(() => import('./spreadsheet/PivotChartPanel.vue'))
const ChartEditorPanel = defineAsyncComponent(() => import('./spreadsheet/ChartEditorPanel.vue'))
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

const props = withDefaults(defineProps<{ modelValue?: ExcelWorkbook; model?: ExcelWorkbook; canUndo?: boolean; canRedo?: boolean; showHeader?: boolean; studio?: OfficeStudio; ribbonVariant?: RibbonVariant }>(), {
  modelValue: undefined,
  model: undefined,
  canUndo: undefined,
  canRedo: undefined,
  showHeader: false,
  studio: undefined,
  ribbonVariant: 'classic',
})
const officeLocale = provideWorkbenchLocales(() => props.studio?.regionalSettings)

const emit = defineEmits<{
  (event: 'update:modelValue', value: ExcelWorkbook): void
  (event: 'update:model', value: ExcelWorkbook): void
  (event: 'change', value: ExcelWorkbook): void
  (event: 'edit-cell', payload: { sheetId: string; key: string; cell: ExcelCell }): void
  (event: 'import', file: File): void
  (event: 'save', value: ExcelWorkbook): void
  (event: 'export', value: ExcelWorkbook): void
  (event: 'undo'): void
  (event: 'redo'): void
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
const userId = computed(() => props.studio?.currentUserId ?? 'local-user')
const availableSheets = computed(() => workbook.value.sheets.filter(sheet => canViewWorksheetProtection(sheet, userId.value)))
const activeSheetId = ref(availableSheets.value.find(sheet => sheet.id === workbook.value.activeSheetId && sheet.visible !== false)?.id ?? availableSheets.value.find(sheet => sheet.visible !== false)?.id ?? '')
const selectedKey = ref('A2')
const selectionAnchor = ref('A2')
const selectedRanges = ref<GridRange[]>([{ startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }])
const emptyCellSelection: string[] = []
const editingKey = ref<string | null>(null)
const editValue = ref('')
const selecting = ref(false)
const formulaInput = ref<InstanceType<typeof SheetFormulaBar> | null>(null)
const spreadsheetCanvas = ref<InstanceType<typeof SpreadsheetCanvas> | null>(null)
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
const copiedCells = ref<Array<Array<ExcelCell | undefined>> | null>(null)
const copiedFormulaCells = ref<Array<Array<ExcelCell | undefined>> | null>(null)
const chartCommands = ref<DrawCommand[]>([])
const chartEditorOpen = ref(false)
const focusedChartId = ref<string | null>(null)
const chartPointerOrigin = ref<{ x: number; y: number } | null>(null)
const chartModalOpen = ref(false)
const chartSource = ref('')
const chartTitle = ref('')
const chartType = ref<ChartType>('bar')
const chartTypeOptions: Array<{ label: string; value: ChartType }> = [
  { label: '簇状柱形图', value: 'bar' },
  { label: '折线图', value: 'line' },
  { label: '面积图', value: 'area' },
  { label: '饼图', value: 'pie' },
  { label: '圆环图', value: 'doughnut' },
  { label: '堆积柱形图', value: 'stacked-column' },
  { label: '散点图', value: 'scatter' },
  { label: '雷达图', value: 'radar' },
]
const chartOrientation = ref<ChartSourceOrientation>('columns')
const chartCategoryField = ref(0)
const chartValueFields = ref<number[]>([])
const chartSeriesXFields = ref<Record<number, number>>({})
const chartXScale = ref<'linear' | 'log'>('linear')
const chartYScale = ref<'linear' | 'log'>('linear')
const chartFieldOptions = computed(() => {
  const range = parseChartRange(chartSource.value)
  const sheet = activeSheet.value
  if (!range || !sheet) return []
  const count = chartOrientation.value === 'columns' ? range.endColumn - range.startColumn + 1 : range.endRow - range.startRow + 1
  return Array.from({ length: count }, (_unused, value) => {
    const row = range.startRow + (chartOrientation.value === 'rows' ? value : 0)
    const column = range.startColumn + (chartOrientation.value === 'columns' ? value : 0)
    const address = `${columnName(column)}${row + 1}`
    const title = sheet.cells[address]?.value
    return { value, label: title == null || title === '' ? address : `${address} · ${title}` }
  })
})
watch([chartSource, chartOrientation], () => {
  chartCategoryField.value = 0
  chartValueFields.value = []
  chartSeriesXFields.value = {}
})
const chartError = ref('')
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
const pivotChartPanelOpen = ref(false)
const pivotSource = ref('')
const slicerPanelOpen = ref(false)
const slicerKind = ref<'slicer' | 'timeline'>('slicer')
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
const axisSizeDialogOpen = ref(false)
const axisSizeDialogKind = ref<'row' | 'column'>('row')
const axisSizeTarget = ref<SheetAxisTarget | null>(null)
const axisSizeValue = ref(20)
const axisSizeError = ref('')
const cellMoveConfirmOpen = ref(false)
const pendingCellMove = ref<{ direction: SheetCellMoveDirection; range: GridRange; sheetId: string } | null>(null)
const copiedRange = ref<GridRange | null>(null)
const copiedFormulaRange = ref<GridRange | null>(null)
const copiedColumnWidths = ref<Record<number, number> | undefined>()
const copiedDefaultColumnWidth = ref<number>()
const sparklinePanelOpen = ref(false)
const sparklineTargetRange = ref('A1')
const customNumberFormatOpen = ref(false)
const conditionalFormattingOpen = ref(false)
const conditionalFormattingRange = ref('A1')
const formatPainterSource = ref<FormatPainterSource | null>(null)
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
const slicerControls = computed(() => {
  const sheet = getChartWorksheet(activeSheetId.value)
  return (workbook.value.slicers ?? []).filter(model => model.hostSheetId === activeSheetId.value).flatMap((model): Array<{ model: SlicerModel; items: string[] }> => {
    try {
      const control = sheet?.getSlicer(model.id)
      if (!control) return []
      const pivot = workbook.value.pivotTables?.find(candidate => candidate.id === model.pivotTableId)
      const table = workbook.value.tables?.find(candidate => candidate.id === model.tableId)
      const tableColumn = table ? table.range.startCol + table.columns.findIndex(column => column.id === model.sourceFieldId) : -1
      const selectedItems = model.tableId
        ? table?.autoFilter?.columns[tableColumn]
        : pivot?.fields.find(field => field.id === model.sourceFieldId && field.area !== 'value')?.selectedItems
      const normalizedSelection = Array.isArray(selectedItems)
        ? selectedItems
        : selectedItems?.kind === 'values'
          ? selectedItems.selectedItems.map(value => value == null ? '' : String(value))
          : undefined
      return [{ model: { ...model, selectedItems: normalizedSelection ? [...normalizedSelection] : undefined }, items: control.getItems() }]
    } catch { return [] }
  })
})

const activeSheet = computed(() => availableSheets.value.find((sheet) => sheet.id === activeSheetId.value && sheet.visible !== false) ?? availableSheets.value.find(sheet => sheet.visible !== false))
const translate = computed(() => createOfficeTranslator(officeLocale.value))
function t(key: string, fallback: string) { return translate.value(`sheet.${key}`, fallback) }
const activeCell = computed(() => activeSheet.value?.cells[selectedKey.value] ?? { value: '' })
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
const canUndo = computed(() => props.canUndo ?? false)
const canRedo = computed(() => props.canRedo ?? false)
const selectedStatistics = computed(() => calculateSelectionStatistics(activeSheet.value ?? { id: '', name: '', cells: {}, rowCount: 0, columnCount: 0 }, selectedRanges.value))
const selectionFormat = computed(() => {
  const keys = selectedExistingKeys()
  function common<K extends keyof ExcelCell>(property: K, fallback: ExcelCell[K]): ExcelCell[K] | undefined {
    const values = keys.map(key => sheetCells.value[key]?.[property] ?? fallback)
    if (selectedCount.value > keys.length || !values.length) values.push(fallback)
    return values.every(value => value === values[0]) ? values[0] : undefined
  }
  return {
    fontFamily: common('fontFamily', 'Aptos') ?? null,
    fontSize: common('fontSize', 11) ?? null,
    numberFormat: (() => {
      const formats = keys.map(key => sheetCells.value[key]?.rawNumberFormat ?? SHEET_NUMBER_FORMATS[sheetCells.value[key]?.numberFormat ?? 'general'])
      if (selectedCount.value > keys.length || !formats.length) formats.push('General')
      const value = formats.every(format => format === formats[0]) ? classifySheetNumberFormat(formats[0]!) : 'custom'
      return value === 'custom' ? null : value
    })(),
    bold: common('bold', false), italic: common('italic', false), underline: common('underline', false), strike: common('strike', false),
    align: common('align', 'left'), vertical: common('vertical', 'middle'), wrap: common('wrap', false),
    wrapStrategy: common('wrapStrategy', 'overflow'),
    color: common('color', null), background: common('background', null),
    rotation: common('rotation', 0) ?? null, rotationVertical: common('rotationVertical', false),
    shrinkToFit: common('shrinkToFit', false),
  }
})
const validationOptions = computed(() => {
  const workbookId = workbook.value.id
  const sheetId = activeSheet.value?.id
  if (!props.studio || !workbookId || !sheetId) return []
  const [column, row] = splitKey(editingKey.value ?? selectedKey.value)
  return FOffice.newAPI(props.studio).getWorkbook(workbookId)?.getSheetBySheetId(sheetId)?.getCellValidationOptions(row - 1, column) ?? []
})
const canUnmerge = computed(() => {
  const [column, row] = splitKey(selectedKey.value)
  return activeSheet.value?.merges?.some(merge => { const range = parseChartRange(merge); return range && containsCell([range], row - 1, column) }) ?? false
})
const canEditValidationSelection = computed(() => Boolean(props.studio && workbook.value.id && activeSheet.value && selectedRanges.value.length
  && canEditWorksheetProtection(activeSheet.value, userId.value) && canEditProtectedRanges(activeSheet.value, selectedModelRanges(), userId.value)))
const borderOptions = ref<SheetBorderOptions>({ style: 'thin', color: '#1f2937' })
const borderPreset = ref<SheetBorderPreset>('all')
const ribbonTabs = computed(() => createSheetRibbon({
  borderOptions: borderOptions.value, borderPreset: borderPreset.value, onBorderOptions: changeBorderOptions,
  ...selectionFormat.value, canUndo: canUndo.value, canRedo: canRedo.value, hasSheet: Boolean(activeSheet.value),
  canMerge: selectedCount.value > 1, canUnmerge: canUnmerge.value,
  canDeleteRow: totalRows.value > 1, canDeleteColumn: (activeSheet.value?.columnCount ?? 0) > 1,
  chartAvailable: Boolean(props.studio && workbook.value.id), gridlines: activeSheet.value?.showGridlines !== false,
  pivotChartAvailable: Boolean(props.studio && workbook.value.pivotTables?.length),
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
}, createRibbonTranslator(officeLocale.value, 'sheet')))
const formulaValue = computed(() => editingKey.value ? editValue.value : activeCell.value.formula ?? activeCell.value.value)
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
  else { viewSessions.clear(); sortController?.abort(); editingKey.value = null }
  const session = sameDocument ? viewSessions.get(nextSheetId) : undefined
  workbook.value = cloneWorkbook(value)
  activeSheetId.value = nextSheetId
  for (const id of viewSessions.keys()) if (!value.sheets.some(sheet => sheet.id === id)) viewSessions.delete(id)
  const sheet = activeSheet.value
  const next = normalizeSheetViewState(sheet ?? { rowCount: 0, columnCount: 0 }, session)
  selectedKey.value = next.key
  selectedRanges.value = next.ranges
  selectionAnchor.value = next.anchor
  zoom.value = session?.zoom ?? displayPreference?.value.zoom ?? next.zoom
  if (nextSheetId !== previousSheetId || (editingKey.value && !parseSheetAddress(editingKey.value, sheet ?? { rowCount: 0, columnCount: 0 }))) editingKey.value = null
}

function commitHistory() {
  workbook.value.activeSheetId = activeSheetId.value
  const snapshot = cloneWorkbook(workbook.value)
  emit('update:modelValue', snapshot)
  emit('update:model', snapshot)
  emit('change', snapshot)
  syncApiSelection()
}

function validateIncomingValues(entries: readonly { key: string; value: string }[]): boolean {
  const sheet = activeSheet.value
  if (!sheet || !entries.length) return true
  if (!canEditWorksheetProtection(sheet, props.studio?.currentUserId ?? 'local-user')) {
    emit('notify', { type: 'warning', message: '此工作表受保护，当前用户没有编辑权限。' })
    return false
  }
  const candidateView = cloneWorkbook(workbook.value)
  const candidateSheet = candidateView.sheets.find(item => item.id === sheet.id)
  if (!candidateSheet) return false
  for (const entry of entries) {
    const current = candidateSheet.cells[entry.key] ?? { value: '' }
    const formula = entry.value.startsWith('=')
    candidateSheet.cells[entry.key] = {
      ...current, value: formula ? '' : entry.value, formula: formula ? entry.value : undefined,
      cachedValue: formula ? '' : undefined, rawValue: undefined, displayValue: undefined,
    }
  }
  const model = applyExcelWorkbookView(props.studio?.getWorkbookModel(workbook.value.id ?? '') ?? createBlankWorkbook(), candidateView)
  const positions = entries.map(({ key }) => {
    const [column, row] = splitKey(key)
    return { row: row - 1, column }
  })
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

function syncApiSelection() {
  const studio = props.studio
  const id = workbook.value.id
  const bounds = selectionBounds(selectedRanges.value)
  if (!studio || !id || !bounds || !studio.getWorkbookModel(id)?.sheets.some(sheet => sheet.id === activeSheetId.value)) return
  studio.setSelection(id, activeSheetId.value, bounds)
}

function rangeNotation(range: GridRange): string {
  return `${columnName(range.startColumn)}${range.startRow + 1}:${columnName(range.endColumn)}${range.endRow + 1}`
}

function parseChartRange(value: string): GridRange | undefined {
  const parts = value.trim().split(':')
  const first = splitKey(parts[0] ?? '')
  const last = splitKey(parts[1] ?? parts[0] ?? '')
  if (!parts[0] || !/^[A-Z]+\d+$/i.test(parts[0]) || !/^[A-Z]+\d+$/i.test(parts[1] ?? parts[0])) return undefined
  return {
    startRow: Math.min(first[1], last[1]) - 1,
    endRow: Math.max(first[1], last[1]) - 1,
    startColumn: Math.min(first[0], last[0]),
    endColumn: Math.max(first[0], last[0]),
  }
}

function openChartDialog() {
  const bounds = selectionBounds(selectedRanges.value)
  chartSource.value = bounds ? rangeNotation(bounds) : 'A1:C4'
  chartTitle.value = ''
  chartType.value = 'bar'
  chartOrientation.value = 'columns'
  chartCategoryField.value = 0
  chartValueFields.value = []
  chartSeriesXFields.value = {}
  chartXScale.value = 'linear'
  chartYScale.value = 'linear'
  chartError.value = ''
  chartModalOpen.value = true
}

async function confirmChartDialog() {
  const sheet = activeSheet.value
  const range = parseChartRange(chartSource.value)
  if (!sheet || !range) { chartError.value = '请输入有效的数据范围，例如 A1:C4'; return }
  try {
    const target = getChartWorksheet(sheet.id)
    if (!target) throw new Error('图表命令上下文尚未绑定到活动工作簿')
    const position = chartPositionForRange(range, sheet)
    await target.insertChart(target.newChart(chartType.value)
      .setSource(chartSource.value, chartOrientation.value)
      .setCategoryField(chartCategoryField.value)
      .setValueFields(chartValueFields.value)
      .setSeriesToCategory(chartType.value === 'scatter' ? Object.fromEntries(chartValueFields.value
        .filter((field) => chartSeriesXFields.value[field] !== undefined)
        .map((field) => [field, [chartSeriesXFields.value[field]!]])) : {})
      .setXAxis({ scaleType: chartType.value === 'scatter' ? chartXScale.value : 'linear' })
      .setYAxis({ scaleType: chartType.value === 'scatter' ? chartYScale.value : 'linear' })
      .setAbsolutePosition(position.x, position.y)
      .setSize(480, 300)
      .setTitle(chartTitle.value.trim())
      .build())
    chartModalOpen.value = false
  } catch (error) {
    chartError.value = error instanceof Error ? error.message : '图表数据范围不可用'
  }
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

function getSlicer(id: string) { return getChartWorksheet(activeSheet.value?.id ?? '')?.getSlicer(id) ?? null }

async function refreshSlicerView() {
  if (!props.studio || !workbook.value.id) return
  const model = props.studio.getWorkbookModel(workbook.value.id)
  if (model) {
    workbook.value = toExcelWorkbookView(model)
    commitHistory()
  }
}

async function updateSlicerPresentation(id: string, patch: import('../features/slicer/slicer-engine').SlicerPresentationPatch) {
  try { await getSlicer(id)?.updatePresentation(patch); await refreshSlicerView() }
  catch (error) { emit('notify', { type: 'error', message: error instanceof Error ? error.message : '切片器更新失败' }) }
}

async function handleSlicerItem(payload: { slicerId: string; value: string; index: number; anchorIndex?: number; ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; multiSelect: boolean }) {
  const control = getSlicer(payload.slicerId)
  if (!control) return
  const items = control.getItems()
  const info = control.getInfo()
  let selected = info.selectedItems == null ? [...items] : [...info.selectedItems]
  if (payload.shiftKey && payload.anchorIndex !== undefined) {
    const first = Math.max(0, Math.min(payload.anchorIndex, payload.index))
    const last = Math.min(items.length - 1, Math.max(payload.anchorIndex, payload.index))
    selected = items.slice(first, last + 1)
  } else if (payload.ctrlKey || payload.metaKey || payload.multiSelect) {
    selected = selected.includes(payload.value) ? selected.filter(value => value !== payload.value) : [...selected, payload.value]
  } else selected = [payload.value]
  try {
    await control.setItems(selected.length === items.length ? null : selected)
    await refreshSlicerView()
  } catch (error) { emit('notify', { type: 'error', message: error instanceof Error ? error.message : '切片器筛选失败' }) }
}

async function clearSlicer(id: string) {
  try { await getSlicer(id)?.clearFilter(); await refreshSlicerView() }
  catch (error) { emit('notify', { type: 'error', message: error instanceof Error ? error.message : '清除筛选失败' }) }
}

async function setSlicerTimelineRange(payload: { slicerId: string; start: number; end: number; level: import('../domain/model').TimelineLevel }) {
  try { await getSlicer(payload.slicerId)?.setTimelineRange({ start: payload.start, end: payload.end, level: payload.level }); await refreshSlicerView() }
  catch (error) { emit('notify', { type: 'error', message: error instanceof Error ? error.message : '时间线筛选失败' }) }
}

function openSlicerSettings(id: string) { selectedSlicerId.value = id; slicerPanelOpen.value = true }
function openSlicerPanel(kind: 'slicer' | 'timeline') { selectedSlicerId.value = ''; slicerKind.value = kind; slicerPanelOpen.value = true }
async function handleSlicerPanelChanged(id: string) { selectedSlicerId.value = id; await refreshSlicerView() }
function moveSlicer(payload: { slicerId: string; left: number; top: number }) {
  const info = getSlicer(payload.slicerId)?.getInfo()
  if (!info || info.lockedPosition || !activeSheet.value) return
  const bounds = resolveSlicerBounds(info, new SpreadsheetGeometry(activeSheet.value))
  void updateSlicerPresentation(payload.slicerId, { layout: { ...bounds, left: payload.left, top: payload.top } })
}
function resizeSlicer(payload: { slicerId: string; width: number; height: number }) {
  const info = getSlicer(payload.slicerId)?.getInfo()
  if (!info || info.lockedPosition || !activeSheet.value) return
  const bounds = resolveSlicerBounds(info, new SpreadsheetGeometry(activeSheet.value))
  const height = info.drawing?.keepRatio ? payload.width * bounds.height / bounds.width : payload.height
  void updateSlicerPresentation(payload.slicerId, { layout: { ...bounds, width: payload.width, height } })
}
function scrollSlicerTo(payload: { slicerId: string; position: number }) {
  const info = getSlicer(payload.slicerId)?.getInfo()
  if (!info) return
  void updateSlicerPresentation(payload.slicerId, info.type === 'timeline' ? { scrollPosition: payload.position } : { startItem: payload.position })
}

async function updateChartLayout(payload: { chartId: string; patch: Partial<Pick<ChartModel, 'position' | 'size'>> }) {
  const chart = getChartWorksheet(activeSheet.value?.id ?? '')?.getChartById(payload.chartId)
  if (chart) await chart.update(payload.patch)
}

function rememberChartPointer(event: PointerEvent): void {
  if (!(event.target instanceof Element) || !event.target.closest('.chart-hitbox')) {
    chartPointerOrigin.value = null
    return
  }
  chartPointerOrigin.value = { x: event.clientX, y: event.clientY }
}

function focusChartFromOverlay(event: MouseEvent): void {
  if (!(event.target instanceof Element) || !event.target.closest('.chart-hitbox')) return
  const origin = chartPointerOrigin.value
  chartPointerOrigin.value = null
  if (origin && Math.hypot(event.clientX - origin.x, event.clientY - origin.y) > 5) return
  const root = event.currentTarget
  if (!(root instanceof HTMLElement)) return
  const hitboxes = [...root.querySelectorAll('.chart-hitbox')]
  const index = hitboxes.indexOf(event.target.closest('.chart-hitbox')!)
  const chart = sheetCharts.value[index]
  if (!chart) return
  focusedChartId.value = chart.id
  chartEditorOpen.value = true
}

async function applyChartSettings(patch: Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>): Promise<void> {
  const chart = focusedChart.value
  const target = chart && getChartWorksheet(chart.sheetId)?.getChartById(chart.id)
  if (!chart || !target) return
  try {
    await target.update(patch)
  } catch (error) {
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '图表设置未能应用' })
  }
}

async function removeChart(chartId: string) {
  const chart = getChartWorksheet(activeSheet.value?.id ?? '')?.getChartById(chartId)
  if (chart) await chart.remove()
}

function setCellValue(key: string, value: string): boolean {
  const sheet = activeSheet.value
  if (!sheet) return false
  if (!canEditWorksheetProtection(sheet, props.studio?.currentUserId ?? 'local-user')) {
    emit('notify', { type: 'warning', message: '此工作表受保护，当前用户没有编辑权限。' })
    return false
  }
  if (!validateIncomingValues([{ key, value }])) return false
  const previous = sheet.cells[key]
  const current = previous ?? { value: '' }
  if ((current.formula ?? current.value) === value) return true
  const formula = value.startsWith('=') ? value : undefined
  sheet.cells[key] = { ...current, value: formula ? '' : value, formula, cachedValue: formula ? '' : undefined, rawValue: undefined, displayValue: undefined }
  try {
    recalculateFormulas()
  } catch (error) {
    if (previous) sheet.cells[key] = previous
    else delete sheet.cells[key]
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '单元格编辑未能提交' })
    return false
  }
  emit('edit-cell', { sheetId: sheet.id, key, cell: activeSheet.value?.cells[key] ?? sheet.cells[key]! })
  commitHistory()
  return true
}

function recalculateFormulas() {
  if (workbook.value.calculationMode === 'manual') return
  const previous = workbook.value.id ? props.studio?.getWorkbookModel(workbook.value.id) : undefined
  const result = recalculateSheetViewFormulas(workbook.value, previous)
  workbook.value = result.workbook
}

function recalculateWorkbook() {
  const workbookId = workbook.value.id
  if (!workbookId || !props.studio) {
    const result = recalculateSheetViewFormulas(workbook.value)
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
  return sheetCells.value[key] ?? { value: '' }
}

function selectCell(key: string, extend = false, keepHidden = false) {
  const [column, requestedRow] = splitKey(key)
  const visibleRow = keepHidden ? requestedRow - 1 : activeSheet.value ? findVisibleSheetRow(activeSheet.value, requestedRow - 1, 1) : undefined
  if (visibleRow === undefined) return
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
  if (!selecting.value) applyFormatPainterToSelection()
}

function armFormatPainter(): void {
  const range = selectionBounds(selectedRanges.value)
  if (!range || !activeSheet.value) return
  const cells = Array.from({ length: range.endRow - range.startRow + 1 }, (_, rowOffset) =>
    Array.from({ length: range.endColumn - range.startColumn + 1 }, (_, columnOffset) =>
      ({ ...getCell(`${columnName(range.startColumn + columnOffset)}${range.startRow + rowOffset + 1}`) })),
  )
  formatPainterSource.value = { range, cells }
  emit('notify', { type: 'info', message: '格式刷已就绪，请选择目标单元格范围。' })
}

function applyFormatPainterToSelection(): void {
  const source = formatPainterSource.value
  const sheet = activeSheet.value
  const targets = selectedRanges.value
  if (!source || !sheet || !targets.length) return
  const patches = createFormatPainterPatches(source, targets, sheet.cells, (row, column) => `${columnName(column)}${row + 1}`)
  if (!checkDirectSelection(patches.length)) return
  for (const patch of patches) sheet.cells[`${columnName(patch.column)}${patch.row + 1}`] = patch.cell
  formatPainterSource.value = null
  commitHistory()
}

function beginEdit(key = selectedKey.value, initialValue?: string, target: 'cell' | 'formula' = 'cell') {
  editingKey.value = key
  editValue.value = initialValue ?? getCell(key).formula ?? getCell(key).value
  nextTick(() => target === 'cell' ? spreadsheetCanvas.value?.focusEditor() : formulaInput.value?.focus())
}

function beginFormulaEdit() {
  if (!editingKey.value) beginEdit(selectedKey.value, undefined, 'formula')
}

function finishFormulaEdit(navigation?: { key: 'Enter' | 'Tab'; shift: boolean }) {
  if (!finishEdit()) return
  if (navigation) moveSelection(navigation.key === 'Enter' ? (navigation.shift ? -1 : 1) : 0, navigation.key === 'Tab' ? (navigation.shift ? -1 : 1) : 0)
  spreadsheetCanvas.value?.focus()
}

function cancelFormulaEdit() { cancelEdit(); spreadsheetCanvas.value?.focus() }

function navigateToAddress(input: string) {
  const sheet = activeSheet.value
  if (!sheet) return
  const parsed = parseSheetAddress(input, sheet)
  if (!parsed) { nameError.value = '请输入当前工作表内的有效地址或范围，例如 A1、A1:C8'; return }
  finishEdit()
  nameError.value = ''
  const next = normalizeSheetViewState(sheet, { key: parsed.anchor, anchor: parsed.anchor, ranges: [parsed.range], zoom: zoom.value })
  selectedKey.value = next.key
  selectionAnchor.value = next.anchor
  selectedRanges.value = next.ranges
  syncApiSelection()
  spreadsheetCanvas.value?.ensureVisible(next.key)
  spreadsheetCanvas.value?.focus()
}

function finishEdit(): boolean {
  if (!editingKey.value) return true
  if (!setCellValue(editingKey.value, editValue.value)) return false
  editingKey.value = null
  return true
}

function cancelEdit() {
  editingKey.value = null
}

function moveSelection(rowOffset: number, columnOffset: number, extend = false) {
  const [column, row] = splitKey(selectedKey.value)
  const nextColumn = Math.max(0, Math.min((activeSheet.value?.columnCount ?? columnCount) - 1, column + columnOffset))
  const nextRow = activeSheet.value ? findVisibleSheetRow(activeSheet.value, row - 1 + rowOffset, rowOffset < 0 ? -1 : 1, false) : undefined
  if (nextRow === undefined) return
  selectCell(`${columnName(nextColumn)}${nextRow + 1}`, extend)
}

function handleGridKeydown(event: KeyboardEvent) {
  if (event.isComposing) return
  const target = event.target as HTMLElement | null
  if (!target?.closest('.spreadsheet-canvas')) return
  if (target?.matches('input, textarea, select')
    && !target.closest('.canvas-cell-input, .formula-input')) return
  if (event.ctrlKey && event.key === 'PageDown') {
    event.preventDefault()
    switchNextSheet()
    return
  }
  if (editingKey.value) {
    if (event.key === 'Enter' || event.key === 'Tab') {
      event.preventDefault()
      if (!finishEdit()) return
      moveSelection(event.key === 'Enter' ? (event.shiftKey ? -1 : 1) : 0, event.key === 'Tab' ? (event.shiftKey ? -1 : 1) : 0)
      spreadsheetCanvas.value?.focus()
    }
    if (event.key === 'Escape') {
      event.preventDefault()
      cancelEdit()
      spreadsheetCanvas.value?.focus()
    }
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') {
    event.preventDefault()
    void executeRibbon('sheet.operation.insert-hyperlink')
    return
  }
  if (event.key === 'Enter' && activeCell.value.hyperlink) {
    event.preventDefault()
    openCellHyperlink(selectedKey.value)
    return
  }
  if (event.key === 'Enter' || event.key === 'F2') {
    event.preventDefault()
    beginEdit()
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
  } else if (event.key === 'ArrowRight' || event.key === 'Tab') {
    event.preventDefault()
    moveSelection(0, event.key === 'Tab' && event.shiftKey ? -1 : 1, event.key === 'ArrowRight' && event.shiftKey)
  } else if (event.key === 'Delete' || event.key === 'Backspace') {
    event.preventDefault()
    clearSelection()
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c') {
    event.preventDefault()
    copySelection()
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'v') {
    event.preventDefault()
    pasteSelection()
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'x') {
    event.preventDefault()
    cutSelection()
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
  for (const range of selectedRanges.value) for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) action(`${columnName(column)}${row + 1}`)
  return true
}
function selectedExistingKeys(): string[] {
  return Object.keys(sheetCells.value).filter(key => { const [column, row] = splitKey(key); return containsCell(selectedRanges.value, row - 1, column) })
}
function copySelection(formulaOnly = false): boolean {
  const bounds = selectionBounds(selectedRanges.value)
  if (!activeSheet.value || !bounds || !checkDirectSelection(selectionCount([bounds]))) return false
  if (formulaOnly) copiedFormulaRange.value = { ...bounds }
  else copiedRange.value = { ...bounds }
  if (!formulaOnly) {
    copiedColumnWidths.value = Object.fromEntries(Array.from({ length: bounds.endColumn - bounds.startColumn + 1 }, (_, offset) => {
      const column = bounds.startColumn + offset
      return [column, activeSheet.value?.columnWidths?.[column] ?? activeSheet.value?.defaultColumnWidth ?? 96]
    }))
    copiedDefaultColumnWidth.value = activeSheet.value.defaultColumnWidth ?? 96
  }
  const copiedRangeCells = Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, rowOffset) =>
    Array.from({ length: bounds.endColumn - bounds.startColumn + 1 }, (_, columnOffset) => ({ ...getCell(`${columnName(bounds.startColumn + columnOffset)}${bounds.startRow + rowOffset + 1}`) })),
  )
  copiedFormulaCells.value = copiedRangeCells.map(row => row.map(cell => cell.formula ? { ...cell } : undefined))
  if (!formulaOnly) copiedCells.value = copiedRangeCells
  const text = formulaOnly
    ? copiedFormulaCells.value.map(row => row.map(cell => cell?.formula ?? '').join('\t')).join('\n')
    : copiedRangeCells.map(row => row.map(cell => cell.formula ?? cell.value).join('\t')).join('\n')
  if (navigator.clipboard?.writeText) void navigator.clipboard.writeText(text).catch(() => emit('notify', { type: 'info', message: '系统剪贴板不可用，内容已保留在工作台内，可继续粘贴。' }))
  return true
}

function cutSelection() {
  if (!copySelection()) return
  if (!activeSheet.value) return
  selectedExistingKeys().forEach((key) => delete activeSheet.value?.cells[key])
  recalculateFormulas()
  commitHistory()
}

function clearSelection() {
  if (!activeSheet.value) return
  const keys = selectedExistingKeys()
  if (!validateIncomingValues(keys.map(key => ({ key, value: '' })))) return
  selectedExistingKeys().forEach((key) => { const cell = activeSheet.value?.cells[key]; if (cell) activeSheet.value!.cells[key] = { ...cell, value: '', formula: undefined, cachedValue: undefined, rawValue: '', displayValue: '' } })
  recalculateFormulas()
  commitHistory()
}

function applyPasteText(text: string) {
  if (!text || !activeSheet.value) return
  const [startColumn, startRow] = splitKey(selectedKey.value)
  const entries: Array<{ key: string; value: string }> = []
  text.replace(/\r/g, '').split('\n').forEach((rowText, rowOffset) => {
    rowText.split('\t').forEach((value, columnOffset) => {
      const column = startColumn + columnOffset
      const row = startRow + rowOffset
      if (column >= (activeSheet.value?.columnCount ?? columnCount) || row > totalRows.value) return
      const key = `${columnName(column)}${row}`
      entries.push({ key, value })
    })
  })
  if (!checkDirectSelection(entries.length) || !validateIncomingValues(entries)) return
  entries.forEach(({ key, value }) => {
    const current = activeSheet.value?.cells[key] ?? { value: '' }
    const formula = value.startsWith('=') ? value : undefined
    activeSheet.value!.cells[key] = { ...current, value: formula ? '' : value, formula, cachedValue: formula ? '' : undefined, rawValue: undefined, displayValue: undefined }
  })
  recalculateFormulas()
  commitHistory()
}

function pasteInternal(mode: SheetPasteMode) {
  const sheet = activeSheet.value
  const source = mode === 'formulas' ? copiedFormulaCells.value : copiedCells.value
  const sourceRange = mode === 'formulas' ? copiedFormulaRange.value : copiedRange.value
  const targetRange = selectionBounds(selectedRanges.value)
  if (!sheet || !source || !sourceRange || !targetRange) {
    if (mode === 'all' || mode === 'values') void pasteSelection()
    return
  }
  const targetCells = source.map((row, rowOffset) => row.map((_, columnOffset) => {
    const rowIndex = targetRange.startRow + rowOffset
    const columnIndex = targetRange.startColumn + columnOffset
    return sheet.cells[`${columnName(columnIndex)}${rowIndex + 1}`] ?? { value: '' }
  }))
  const plan = createSheetPastePlan({
    mode,
    sourceCells: source,
    targetCells,
    sourceStartRow: sourceRange.startRow,
    sourceStartColumn: sourceRange.startColumn,
    targetStartRow: targetRange.startRow,
    targetStartColumn: targetRange.startColumn,
    sourceColumnWidths: copiedColumnWidths.value,
    sourceDefaultColumnWidth: copiedDefaultColumnWidth.value,
  })
  const patches = plan.cellPatches.map(({ rowOffset, columnOffset, cell }) => ({
    key: `${columnName(targetRange.startColumn + columnOffset)}${targetRange.startRow + rowOffset + 1}`,
    cell,
  }))
  if (!patches.length && !plan.columnWidthPatches.length) return
  if (!checkDirectSelection(patches.length) || !validateIncomingValues(patches.map(({ key, cell }) => ({ key, value: cell.formula ?? cell.value })))) return
  patches.forEach(({ key, cell }) => { sheet.cells[key] = cell })
  if (plan.columnWidthPatches.length) {
    sheet.columnWidths ??= {}
    plan.columnWidthPatches.forEach(({ column, width }) => { sheet.columnWidths![column] = width })
  }
  recalculateFormulas()
  commitHistory()
}

function fillSelection(source: GridRange, target: { row: number; column: number }, applyType: SheetFillApplyType = 'SERIES') {
  const sheet = activeSheet.value
  if (!sheet) return
  const sourceCells = Array.from({ length: source.endRow - source.startRow + 1 }, (_, rowOffset) =>
    Array.from({ length: source.endColumn - source.startColumn + 1 }, (_, columnOffset) => {
      const key = `${columnName(source.startColumn + columnOffset)}${source.startRow + rowOffset + 1}`
      return { ...getCell(key) }
    }),
  )
  const targetRange = {
    startRow: Math.min(source.startRow, target.row),
    endRow: Math.max(source.endRow, target.row),
    startColumn: Math.min(source.startColumn, target.column),
    endColumn: Math.max(source.endColumn, target.column),
  }
  const targetCells = Array.from({ length: targetRange.endRow - targetRange.startRow + 1 }, (_, rowOffset) =>
    Array.from({ length: targetRange.endColumn - targetRange.startColumn + 1 }, (_, columnOffset) => {
      const key = `${columnName(targetRange.startColumn + columnOffset)}${targetRange.startRow + rowOffset + 1}`
      return sheet.cells[key]
    }),
  )
  const plan = createSheetFillPlan(source, target, sourceCells, { applyType, targetCells })
  if (!plan.cellPatches.length || !checkDirectSelection(plan.cellPatches.length)) return
  const destinations = plan.cellPatches.map(({ row, column }) => ({ startRow: row, endRow: row, startCol: column, endCol: column }))
  if (!canEditProtectedRanges(sheet, destinations, props.studio?.currentUserId ?? 'local-user')) {
    emit('notify', { type: 'warning', message: '填充范围包含受保护单元格，未修改数据。' })
    return
  }
  const entries = plan.cellPatches.map(({ row, column, cell }) => ({
    key: `${columnName(column)}${row + 1}`,
    value: cell.formula ?? String(cell.rawValue ?? cell.value ?? ''),
  }))
  if (applyType !== 'ONLY_FORMAT' && !validateIncomingValues(entries)) return
  const before = cloneWorkbook(workbook.value)
  try {
    for (const { row, column, cell } of plan.cellPatches) {
      sheet.cells[`${columnName(column)}${row + 1}`] = cell
    }
    recalculateFormulas()
    selectedRanges.value = [plan.range]
    selectedKey.value = `${columnName(target.column)}${target.row + 1}`
    selectionAnchor.value = selectedKey.value
    commitHistory()
  } catch (error) {
    workbook.value = before
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '填充操作未能提交' })
  }
}

function handlePaste(event: ClipboardEvent) {
  const text = event.clipboardData?.getData('text/plain')
  if (text) {
    event.preventDefault()
    applyPasteText(text)
  }
}

async function pasteSelection() {
  let text = ''
  try {
    text = (await navigator.clipboard?.readText()) ?? ''
  } catch {
    text = ''
  }
  if (!text && copiedCells.value) text = copiedCells.value.map((row) => row.map((cell) => cell?.formula ?? cell?.value ?? '').join('\t')).join('\n')
  applyPasteText(text)
}

function undo() {
  emit('undo')
}

function redo() {
  emit('redo')
}

function toggleStyle(style: 'bold' | 'italic' | 'underline' | 'strike') {
  if (!activeSheet.value) return
  const value = selectionFormat.value[style] !== true
  if (!applyToSelection((key) => {
    const cell = activeSheet.value?.cells[key] ?? { value: '' }
    activeSheet.value!.cells[key] = { ...cell, [style]: value }
  })) return
  commitHistory()
}

function setCellFormat<K extends keyof ExcelCell>(property: K, value: ExcelCell[K]) {
  if (!activeSheet.value) return
  if (!applyToSelection(key => {
    const cell = activeSheet.value?.cells[key] ?? { value: '' }
    activeSheet.value!.cells[key] = { ...cell, [property]: value }
  })) return
  commitHistory()
}

function changeSelectedCells(change: (cell: ExcelCell) => ExcelCell): void {
  if (!activeSheet.value || !applyToSelection(key => { activeSheet.value!.cells[key] = change(activeSheet.value!.cells[key] ?? { value: '' }) })) return
  commitHistory()
}
function setTextRotation(value: number | 'v'): void {
  changeSelectedCells(cell => typeof value === 'number'
    ? { ...cell, rotation: value, rotationVertical: false }
    : { ...cell, rotation: 0, rotationVertical: true })
}
function setNumberFormat(format: SheetNumberFormat): void { changeSelectedCells(cell => applySheetNumberFormat(cell, format)) }
function changeFontSize(direction: -1 | 1): void { changeSelectedCells(cell => ({ ...cell, fontSize: nextSheetFontSize(cell.fontSize ?? 11, direction) })) }
function setBorders(preset: SheetBorderPreset): void {
  if (!activeSheet.value || !checkDirectSelection()) return
  if (!canEditWorksheetProtection(activeSheet.value, userId.value) || !canEditProtectedRanges(activeSheet.value, selectedModelRanges(), userId.value)) {
    emit('notify', { type: 'warning', message: '当前选区受保护。' }); return
  }
  borderPreset.value = preset
  applySheetBorders(activeSheet.value.cells, selectedRanges.value, preset, borderOptions.value)
  commitHistory()
}

function changeBorderOptions(value: SheetBorderOptions): void {
  if (value.style === borderOptions.value.style && value.color === borderOptions.value.color) return
  borderOptions.value = { ...value }
  setBorders(borderPreset.value)
}

function setAlign(align: 'left' | 'center' | 'right') {
  if (!activeSheet.value) return
  if (!applyToSelection((key) => {
    const cell = activeSheet.value?.cells[key] ?? { value: '' }
    activeSheet.value!.cells[key] = { ...cell, align }
  })) return
  commitHistory()
}

function setBackground(color: string) {
  if (!activeSheet.value) return
  if (!applyToSelection((key) => {
    const cell = activeSheet.value?.cells[key] ?? { value: '' }
    activeSheet.value!.cells[key] = { ...cell, background: color }
  })) return
  commitHistory()
}

function resetSelectionColor(command: 'sheet.command.reset-text-color' | 'sheet.command.reset-background-color') {
  const sheet = activeSheet.value
  if (!sheet) return
  const ranges = selectedModelRanges()
  if (props.studio && workbook.value.id) {
    const unitId = workbook.value.id
    if (!props.studio.executeCommandSync(command, { unitId, sheetId: sheet.id, ranges })) return
    const model = props.studio.getWorkbookModel(unitId)
    if (model) workbook.value = toExcelWorkbookView(model)
  } else {
    if (!canEditWorksheetProtection(sheet, userId.value) || !canEditProtectedRanges(sheet, ranges, userId.value)) throw new Error('当前选区受保护。')
    const field = command === 'sheet.command.reset-text-color' ? 'color' : 'background'
    const keys = selectedExistingKeys().filter(key => sheet.cells[key]?.[field] != null)
    if (!keys.length) return
    for (const key of keys) sheet.cells[key]![field] = null
  }
  commitHistory()
}

function mergeSelection(mode: 'all' | 'vertical' | 'horizontal' = 'all') {
  if (!activeSheet.value) return
  const ranges = selectedRanges.value.flatMap(selection => {
    if (mode === 'horizontal') return Array.from({ length: selection.endRow - selection.startRow + 1 }, (_, offset) => ({
      startRow: selection.startRow + offset, endRow: selection.startRow + offset,
      startColumn: selection.startColumn, endColumn: selection.endColumn,
    }))
    if (mode === 'vertical') return Array.from({ length: selection.endColumn - selection.startColumn + 1 }, (_, offset) => ({
      startRow: selection.startRow, endRow: selection.endRow,
      startColumn: selection.startColumn + offset, endColumn: selection.startColumn + offset,
    }))
    return [selection]
  }).filter(range => range.startRow !== range.endRow || range.startColumn !== range.endColumn)
  if (!ranges.length) return
  const existing = activeSheet.value.merges ?? []
  const overlaps = ranges.some(range => existing.some(merge => {
    const parsed = parseChartRange(merge)
    return parsed && range.startRow <= parsed.endRow && range.endRow >= parsed.startRow
      && range.startColumn <= parsed.endColumn && range.endColumn >= parsed.startColumn
  }))
  if (overlaps) { emit('notify', { type: 'warning', message: '所选区域与现有合并单元格重叠。' }); return }
  activeSheet.value.merges = [...existing, ...ranges.map(range => `${columnName(range.startColumn)}${range.startRow + 1}:${columnName(range.endColumn)}${range.endRow + 1}`)]
  const cells = selectedExistingKeys()
  for (const range of ranges) {
    const master = `${columnName(range.startColumn)}${range.startRow + 1}`
    for (const key of cells) {
      const [column, row] = splitKey(key)
      if (key !== master && row - 1 >= range.startRow && row - 1 <= range.endRow && column >= range.startColumn && column <= range.endColumn) delete activeSheet.value.cells[key]
    }
  }
  const first = ranges[0]!
  selectCell(`${columnName(first.startColumn)}${first.startRow + 1}`)
  commitHistory()
}

function transformMerges(axis: 'row' | 'column', index: number, delta: 1 | -1, targetSheet = activeSheet.value) {
  if (!targetSheet?.merges) return
  targetSheet.merges = targetSheet.merges.flatMap((range) => {
    const [start, end] = range.split(':')
    const [startColumn, startRow] = splitKey(start)
    const [endColumn, endRow] = splitKey(end)
    const startAxis = axis === 'row' ? startRow : startColumn
    const endAxis = axis === 'row' ? endRow : endColumn
    if (delta < 0 && index >= startAxis && index <= endAxis) return []
    const move = (key: string) => {
      const [column, row] = splitKey(key)
      const next = axis === 'row'
        ? `${columnName(column)}${row > index || (delta > 0 && row === index) ? row + delta : row}`
        : `${columnName(column > index || (delta > 0 && column === index) ? column + delta : column)}${row}`
      return next
    }
    return [`${move(start)}:${move(end)}`]
  })
}

function unmergeSelection() {
  if (!activeSheet.value) return
  const merges = activeSheet.value.merges ?? []
  const key = selectedKey.value
  const merge = merges.find((range) => {
    const [start, end] = range.split(':')
    const [startColumn, startRow] = splitKey(start)
    const [endColumn, endRow] = splitKey(end)
    const [column, row] = splitKey(key)
    return column >= startColumn && column <= endColumn && row >= startRow && row <= endRow
  })
  if (merge) {
    activeSheet.value.merges = merges.filter((range) => range !== merge)
    commitHistory()
  }
}

interface SheetAxisTarget { sheetId: string; index: number }

function readAxisTarget(params: unknown): SheetAxisTarget | undefined | null {
  if (params === undefined) return undefined
  if (typeof params !== 'object' || params === null || !('sheetId' in params) || !('index' in params)) return null
  const value = params as { sheetId?: unknown; index?: unknown }
  if (typeof value.sheetId !== 'string' || !Number.isSafeInteger(value.index) || Number(value.index) < 0) return null
  return { sheetId: value.sheetId, index: Number(value.index) }
}

function axisTargetSheet(target: SheetAxisTarget | undefined) {
  return target ? workbook.value.sheets.find(sheet => sheet.id === target.sheetId) : activeSheet.value
}

function setActiveCellForAxisTarget(sheet: ExcelSheet, row: number, column: number) {
  if (sheet.id !== activeSheetId.value) return
  selectedKey.value = `${columnName(Math.max(0, column))}${Math.max(1, row + 1)}`
  selectCell(selectedKey.value)
}

function insertRow(params?: unknown) {
  const requestedTarget = readAxisTarget(params)
  if (requestedTarget === null) return
  const sheet = axisTargetSheet(requestedTarget)
  if (!sheet) return
  const [, selectedRow] = splitKey(selectedKey.value)
  const insertionRow = requestedTarget?.index ?? selectedRow - 1
  if (insertionRow >= sheet.rowCount) return
  const updated: Record<string, ExcelCell> = {}
  Object.entries(sheet.cells).forEach(([key, cell]) => {
    const [column, row] = splitKey(key)
    updated[`${columnName(column)}${row >= insertionRow + 1 ? row + 1 : row}`] = cell
  })
  sheet.cells = updated
  transformMerges('row', insertionRow + 1, 1, sheet)
  if (sheet.hiddenRows) sheet.hiddenRows = sheet.hiddenRows.map(row => row >= insertionRow ? row + 1 : row)
  sheet.rowCount += 1
  setActiveCellForAxisTarget(sheet, insertionRow, splitKey(selectedKey.value)[0] - 1)
  commitHistory()
}

function deleteRow(params?: unknown) {
  const requestedTarget = readAxisTarget(params)
  if (requestedTarget === null) return
  const sheet = axisTargetSheet(requestedTarget)
  if (!sheet || sheet.rowCount <= 1) return
  const [, selectedRow] = splitKey(selectedKey.value)
  const deletedRow = requestedTarget?.index ?? selectedRow - 1
  if (deletedRow >= sheet.rowCount) return
  const updated: Record<string, ExcelCell> = {}
  Object.entries(sheet.cells).forEach(([key, cell]) => {
    const [column, row] = splitKey(key)
    if (row < deletedRow + 1) updated[key] = cell
    else if (row > deletedRow + 1) updated[`${columnName(column)}${row - 1}`] = cell
  })
  sheet.cells = updated
  transformMerges('row', deletedRow + 1, -1, sheet)
  sheet.rowCount -= 1
  if (sheet.hiddenRows) sheet.hiddenRows = sheet.hiddenRows.filter(row => row !== deletedRow).map(row => row > deletedRow ? row - 1 : row)
  setActiveCellForAxisTarget(sheet, Math.min(deletedRow, sheet.rowCount - 1), splitKey(selectedKey.value)[0] - 1)
  commitHistory()
}

function insertColumn(params?: unknown) {
  const requestedTarget = readAxisTarget(params)
  if (requestedTarget === null) return
  const sheet = axisTargetSheet(requestedTarget)
  if (!sheet) return
  const [selectedColumn] = splitKey(selectedKey.value)
  const insertionColumn = requestedTarget?.index ?? selectedColumn - 1
  if (insertionColumn >= sheet.columnCount) return
  const updated: Record<string, ExcelCell> = {}
  Object.entries(sheet.cells).forEach(([key, cell]) => {
    const [column, row] = splitKey(key)
    updated[`${columnName(column >= insertionColumn ? column + 1 : column)}${row}`] = cell
  })
  sheet.cells = updated
  transformMerges('column', insertionColumn, 1, sheet)
  sheet.columnCount += 1
  if (sheet.hiddenColumns) sheet.hiddenColumns = sheet.hiddenColumns.map(column => column >= insertionColumn ? column + 1 : column)
  setActiveCellForAxisTarget(sheet, splitKey(selectedKey.value)[1] - 1, insertionColumn)
  commitHistory()
}

function deleteColumn(params?: unknown) {
  const requestedTarget = readAxisTarget(params)
  if (requestedTarget === null) return
  const sheet = axisTargetSheet(requestedTarget)
  if (!sheet || sheet.columnCount <= 1) return
  const [selectedColumn] = splitKey(selectedKey.value)
  const deletedColumn = requestedTarget?.index ?? selectedColumn - 1
  if (deletedColumn >= sheet.columnCount) return
  const updated: Record<string, ExcelCell> = {}
  Object.entries(sheet.cells).forEach(([key, cell]) => {
    const [column, row] = splitKey(key)
    if (column < deletedColumn) updated[key] = cell
    else if (column > deletedColumn) updated[`${columnName(column - 1)}${row}`] = cell
  })
  sheet.cells = updated
  transformMerges('column', deletedColumn, -1, sheet)
  sheet.columnCount -= 1
  if (sheet.hiddenColumns) sheet.hiddenColumns = sheet.hiddenColumns.filter(column => column !== deletedColumn).map(column => column > deletedColumn ? column - 1 : column)
  setActiveCellForAxisTarget(sheet, splitKey(selectedKey.value)[1] - 1, Math.min(deletedColumn, sheet.columnCount - 1))
  commitHistory()
}

function requestCellRangeMove(direction: SheetCellMoveDirection, confirmMerges = false): void {
  const sheet = activeSheet.value
  const range = selectedRanges.value.at(-1)
  if (!sheet || !range || !checkDirectSelection()) return
  if ((direction === 'right' || direction === 'down') && selectedRanges.value.some((first, index) =>
    selectedRanges.value.slice(index + 1).some(second => first.startRow <= second.endRow && first.endRow >= second.startRow
      && first.startColumn <= second.endColumn && first.endColumn >= second.startColumn))) {
    emit('notify', { type: 'warning', message: '选区相互重叠，无法插入并移动单元格。' })
    return
  }
  if (!finishEdit()) return
  const result = moveSheetCellRange(sheet, range, direction, confirmMerges)
  if (result.status === 'invalid') {
    emit('notify', { type: 'warning', message: result.reason })
    return
  }
  if (result.status === 'confirm-merge') {
    pendingCellMove.value = { direction, range: { ...range }, sheetId: sheet.id }
    cellMoveConfirmOpen.value = true
    return
  }
  applyCellRangeMove(sheet.id, range, direction, result.sheet)
}

function confirmCellRangeMove(): void {
  const pending = pendingCellMove.value
  cellMoveConfirmOpen.value = false
  pendingCellMove.value = null
  if (!pending) return
  const sheet = workbook.value.sheets.find(candidate => candidate.id === pending.sheetId)
  if (!sheet || activeSheetId.value !== pending.sheetId || !finishEdit()) return
  const result = moveSheetCellRange(sheet, pending.range, pending.direction, true)
  if (result.status !== 'applied') {
    emit('notify', { type: 'warning', message: result.status === 'invalid' ? result.reason : '当前选区无法移位。' })
    return
  }
  applyCellRangeMove(sheet.id, pending.range, pending.direction, result.sheet)
}

function applyCellRangeMove(sheetId: string, range: GridRange, direction: SheetCellMoveDirection, movedSheet: ExcelSheet): void {
  const previousWorkbook = cloneWorkbook(workbook.value)
  try {
    const nextWorkbook = {
      ...workbook.value,
      sheets: workbook.value.sheets.map(sheet => sheet.id === sheetId ? movedSheet : sheet),
    }
    workbook.value = transformWorkbookReferencesForSheetEdit(nextWorkbook, sheetId, structureEdit(range, direction))
    recalculateFormulas()
    const anchor = `${columnName(range.startColumn)}${range.startRow + 1}`
    selectedKey.value = anchor
    selectionAnchor.value = anchor
    selectedRanges.value = [{ ...range }]
    syncApiSelection()
    commitHistory()
  } catch (error) {
    workbook.value = previousWorkbook
    emit('notify', { type: 'error', message: error instanceof Error ? error.message : '单元格移位未能提交。' })
  }
}

function hideAxisTarget(axis: 'row' | 'column', params?: unknown) {
  const target = readAxisTarget(params)
  if (!target) return
  const sheet = axisTargetSheet(target)
  if (!sheet) return
  const hidden = axis === 'row' ? (sheet.hiddenRows ??= []) : (sheet.hiddenColumns ??= [])
  if (target.index >= (axis === 'row' ? sheet.rowCount : sheet.columnCount) || hidden.includes(target.index)) return
  hidden.push(target.index)
  hidden.sort((left, right) => left - right)
  if (sheet.id === activeSheetId.value) {
    const [column, row] = splitKey(selectedKey.value)
    if (axis === 'row' && row - 1 === target.index) {
      const next = Array.from({ length: sheet.rowCount }, (_, index) => index).find(index => !hidden.includes(index))
      if (next !== undefined) setActiveCellForAxisTarget(sheet, next, column - 1)
    }
    if (axis === 'column' && column === target.index) {
      const next = Array.from({ length: sheet.columnCount }, (_, index) => index).find(index => !hidden.includes(index))
      if (next !== undefined) setActiveCellForAxisTarget(sheet, row - 1, next)
    }
  }
  commitHistory()
}

function showHiddenAxisTargets(axis: 'row' | 'column', params?: unknown) {
  if (typeof params !== 'object' || params === null || !('sheetId' in params) || typeof params.sheetId !== 'string') return
  const sheet = workbook.value.sheets.find(candidate => candidate.id === params.sheetId)
  if (!sheet) return
  if (axis === 'row') sheet.hiddenRows = []
  else sheet.hiddenColumns = []
  commitHistory()
}

function sheetTextWidth(text: string, fontSize: number, fontFamily: string) {
  const context = document.createElement('canvas').getContext('2d')
  if (!context) return text.length * fontSize
  context.font = `${fontSize}px ${fontFamily}`
  return context.measureText(text).width
}

function autoFitAxisTarget(axis: 'row' | 'column', params?: unknown) {
  const target = readAxisTarget(params)
  if (!target) return
  const sheet = axisTargetSheet(target)
  if (!sheet || target.index >= (axis === 'row' ? sheet.rowCount : sheet.columnCount)) return
  const baseFontSize = 14
  if (axis === 'column') {
    const width = Object.entries(sheet.cells).reduce((maximum, [key, cell]) => {
      const [column] = splitKey(key)
      if (column !== target.index) return maximum
      const text = cell.displayValue ?? cell.value
      const fontSize = cell.fontSize ?? baseFontSize
      return Math.max(maximum, sheetTextWidth(text, fontSize, cell.fontFamily ?? 'Arial') + 16 + (cell.indent ?? 0) * fontSize)
    }, 24)
    sheet.columnWidths ??= {}
    sheet.columnWidths[target.index] = Math.min(400, Math.ceil(width))
  } else {
    const height = Object.entries(sheet.cells).reduce((maximum, [key, cell]) => {
      const [column, row] = splitKey(key)
      if (row !== target.index + 1) return maximum
      const text = cell.displayValue ?? cell.value
      const fontSize = cell.fontSize ?? baseFontSize
      const columnWidth = sheet.columnWidths?.[column] ?? sheet.defaultColumnWidth ?? 96
      const lines = cell.wrap ? Math.max(1, Math.ceil(sheetTextWidth(text, fontSize, cell.fontFamily ?? 'Arial') / Math.max(1, columnWidth - 12))) : 1
      return Math.max(maximum, lines * Math.max(fontSize + 5, 18) + 4)
    }, 20)
    sheet.rowHeights ??= {}
    sheet.rowHeights[target.index] = Math.min(409, Math.ceil(height))
  }
  commitHistory()
}

function openAxisSizeDialog(axis: 'row' | 'column', params?: unknown) {
  const target = readAxisTarget(params)
  if (!target) return
  const sheet = axisTargetSheet(target)
  if (!sheet || target.index >= (axis === 'row' ? sheet.rowCount : sheet.columnCount)) return
  axisSizeTarget.value = target
  axisSizeValue.value = axis === 'row' ? sheet.rowHeights?.[target.index] ?? 20 : sheet.columnWidths?.[target.index] ?? sheet.defaultColumnWidth ?? 96
  axisSizeError.value = ''
  axisSizeDialogOpen.value = true
}

function confirmAxisSize() {
  const target = axisSizeTarget.value
  const value = axisSizeValue.value
  const axis = axisSizeDialogKind.value
  if (!target || !Number.isFinite(value) || value < 0 || value > 1000) {
    axisSizeError.value = axis === 'row' ? '行高必须介于 0 到 1000 像素。' : '列宽必须介于 0 到 1000 像素。'
    return
  }
  const sheet = axisTargetSheet(target)
  if (!sheet || target.index >= (axis === 'row' ? sheet.rowCount : sheet.columnCount)) return
  if (axis === 'row') { sheet.rowHeights ??= {}; sheet.rowHeights[target.index] = value }
  else { sheet.columnWidths ??= {}; sheet.columnWidths[target.index] = value }
  axisSizeDialogOpen.value = false
  commitHistory()
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
  if (activeSheetId.value) viewSessions.set(activeSheetId.value, { key: selectedKey.value, anchor: selectionAnchor.value, ranges: selectedRanges.value.map(range => ({ ...range })), zoom: zoom.value })
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
  zoom.value = session?.zoom ?? displayPreference?.value.zoom ?? next.zoom
  if (props.studio && workbook.value.id && props.studio.getWorkbookModel(workbook.value.id)?.sheets.some(candidate => candidate.id === id)) {
    FOffice.newAPI(props.studio).getWorkbook(workbook.value.id)?.setActiveSheet(id)
  }
  syncApiSelection()
}

function addSheet() {
  if (!finishEdit()) return
  const id = `sheet-${crypto.randomUUID()}`
  workbook.value.sheets.push({ id, name: createUniqueSheetName(workbook.value.sheets), cells: {}, rowCount: 100, columnCount })
  activateSheet(id)
  commitHistory()
}

function copySheet(id: string) {
  if (!finishEdit()) return
  const sourceIndex = workbook.value.sheets.findIndex(sheet => sheet.id === id)
  const source = workbook.value.sheets[sourceIndex]
  if (!source) return
  if (!canChangeWorksheet(id)) return
  let suffix = 2
  let name = `${source.name} (${suffix})`
  while (workbook.value.sheets.some(sheet => sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())) {
    suffix += 1
    name = `${source.name} (${suffix})`
  }
  const duplicate = cloneWorkbook({ ...workbook.value, sheets: [source] }).sheets[0]
  if (!duplicate) return
  duplicate.id = `sheet-${crypto.randomUUID()}`
  duplicate.name = name
  duplicate.visible = true
  workbook.value.sheets.splice(sourceIndex + 1, 0, duplicate)
  viewSessions.set(duplicate.id, { key: selectedKey.value, anchor: selectionAnchor.value, ranges: selectedRanges.value.map(range => ({ ...range })), zoom: zoom.value })
  activateSheet(duplicate.id)
  commitHistory()
  emit('notify', { type: 'info', message: `已复制工作表“${source.name}”的单元格和工作表设置。图表、透视表与结构化表格关联不会作为副本创建。` })
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

function setSheetTabColor(id: string, color: string) {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id)
  if (!sheet || (color !== '' && !/^#[0-9a-f]{6}$/i.test(color))) return
  if (!canChangeWorksheet(id)) return
  sheet.tabColor = color || undefined
  commitHistory()
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
  if (!canChangeWorksheet(id)) return
  sheet.name = renameValue.value.trim()
  renameSheetId.value = null
  commitHistory()
}

function deleteSheet(id: string) {
  const sheets = workbook.value.sheets
  const sheet = sheets.find(candidate => candidate.id === id)
  if (!sheet || (sheet.visible !== false && sheets.filter(candidate => candidate.visible !== false).length <= 1)) return
  if (!canChangeWorksheet(id)) return
  finishEdit()
  const remaining = sheets.filter(candidate => candidate.id !== id)
  if (activeSheetId.value === id) {
    const next = remaining.find(candidate => candidate.visible !== false && canViewWorksheetProtection(candidate, userId.value))
    if (next) activateSheet(next.id)
  }
  workbook.value.sheets = remaining
  viewSessions.delete(id)
  commitHistory()
}

function hideSheet(id: string) {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id)
  const next = availableSheets.value.find(candidate => candidate.id !== id && candidate.visible !== false)
  if (!sheet || !next) return
  if (!canChangeWorksheet(id)) return
  if (id === activeSheetId.value) activateSheet(next.id)
  sheet.visible = false
  commitHistory()
}

function showSheet(id: string) {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id)
  if (!sheet) return
  if (!canChangeWorksheet(id)) return
  sheet.visible = true
  activateSheet(id)
  commitHistory()
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
  if (!finishEdit()) return
  const [column, row] = splitKey(key)
  if (event.button === 2 && containsCell(selectedRanges.value, row - 1, column)) return
  selecting.value = true
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
  const sheet = activeSheet.value
  if (!sheet) return
  const range = hit.kind === 'all' ? { startRow: 0, endRow: sheet.rowCount - 1, startColumn: 0, endColumn: sheet.columnCount - 1 }
    : hit.kind === 'row' ? { startRow: hit.row, endRow: hit.row, startColumn: 0, endColumn: sheet.columnCount - 1 }
    : hit.kind === 'column' ? { startRow: 0, endRow: sheet.rowCount - 1, startColumn: hit.column, endColumn: hit.column } : undefined
  if (!range) return
  selecting.value = true
  selectedRanges.value = event.ctrlKey || event.metaKey ? addRange(selectedRanges.value, range) : [range]
  selectedKey.value = `${columnName(range.startColumn)}${range.startRow + 1}`
  selectionAnchor.value = selectedKey.value
  syncApiSelection()
}

function resizeAxis(axis: 'row' | 'column', index: number, size: number) {
  const sheet = activeSheet.value
  if (!sheet) return
  if (axis === 'row') { sheet.rowHeights ??= {}; sheet.rowHeights[index] = size }
  else { sheet.columnWidths ??= {}; sheet.columnWidths[index] = size }
  commitHistory()
}

function extendCellSelection(key: string) {
  if (selecting.value) selectCell(key, true)
}

function stopSelecting() {
  const shouldApplyPainter = selecting.value && formatPainterSource.value !== null
  selecting.value = false
  if (shouldApplyPainter) applyFormatPainterToSelection()
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
  if (!match) return
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
  const candidateModel = applyExcelWorkbookView(baseModel, candidate)
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
function closeSearch() { searchOpen.value = false; searchText.value = ''; replaceText.value = ''; findSelectionOnly.value = false; findSelectionRanges.value = []; spreadsheetCanvas.value?.focus() }
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

function openPivotChartPanel(): void {
  if (!workbook.value.pivotTables?.length) {
    emit('notify', { type: 'info', message: '请先创建数据透视表。' })
    return
  }
  pivotChartPanelOpen.value = true
}

function openDataPanel(mode: 'filter' | 'sort'): void {
  const bounds = selectionBounds(selectedRanges.value)
  dataPanelMode.value = mode
  dataPanelRange.value = bounds ? rangeNotation(bounds) : selectedKey.value
  dataPanelOpen.value = true
}
function openDataValidationPanel(initialType?: 'checkbox' | 'list' | 'date' | 'time', ruleId?: string): void {
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
  const bounds = selectionBounds(selectedRanges.value)
  conditionalFormattingRange.value = bounds ? rangeNotation(bounds) : selectedKey.value
  conditionalFormattingOpen.value = true
}
function applyConditionalFormatting(rules: NonNullable<ExcelSheet['conditionalFormats']>): void {
  const sheet = activeSheet.value
  if (!sheet) return
  const cloned = rules.map(rule => ({
    ...rule,
    ranges: rule.ranges.map(range => ({ ...range })),
    style: { ...rule.style },
    config: rule.config ? {
      ...rule.config,
      min: rule.config.min ? { ...rule.config.min } : undefined,
      max: rule.config.max ? { ...rule.config.max } : undefined,
      thresholds: rule.config.thresholds?.map(item => ({ ...item, value: { ...item.value } })),
      icons: rule.config.icons?.map(item => ({ ...item, value: { ...item.value } })),
    } : undefined,
    metadata: rule.metadata ? { ...rule.metadata } : undefined,
  }))
  workbook.value = { ...workbook.value, sheets: workbook.value.sheets.map(candidate => candidate.id === sheet.id ? { ...candidate, conditionalFormats: cloned } : candidate) }
  commitHistory()
}
function openSparklinePanel(): void {
  const bounds = selectionBounds(selectedRanges.value)
  sparklineTargetRange.value = bounds ? rangeNotation(bounds) : selectedKey.value
  sparklinePanelOpen.value = true
}
function changeOutline(axis: 'row' | 'column', operation: SheetOutlineOperation): void {
  const range = selectedRanges.value.length === 1 ? selectedRanges.value[0] : undefined
  if (!range) return
  const start = axis === 'row' ? range.startRow : range.startColumn
  const end = axis === 'row' ? range.endRow : range.endColumn
  runOutlineCommand({ kind: operation, axis, start, end })
}
function runOutlineCommand(action: SheetOutlineAction): void {
  const sheet = activeSheet.value
  if (!sheet) return
  try {
    const unitId = workbook.value.id
    if (props.studio && unitId && props.studio.getWorkbookModel(unitId)) {
      if (!props.studio.executeCommandSync<ChangeSheetOutlineInput, boolean>(CHANGE_SHEET_OUTLINE_COMMAND, { unitId, sheetId: sheet.id, action })) return
      const model = props.studio.getWorkbookModel(unitId)
      if (model) workbook.value = toExcelWorkbookView(model)
    } else {
      const outline = getSheetOutlineUpdate(sheet, action, userId.value)
      if (!outline) return
      const outlineHiddenRows = [...getCollapsedOutlineIndexes({ count: sheet.rowCount, groups: outline.groups.filter(group => group.axis === 'row') })]
      const outlineHiddenColumns = [...getCollapsedOutlineIndexes({ count: sheet.columnCount, groups: outline.groups.filter(group => group.axis === 'column') })]
      workbook.value = { ...workbook.value, sheets: workbook.value.sheets.map(candidate => candidate.id === sheet.id
        ? { ...candidate, outline, outlineHiddenRows, outlineHiddenColumns }
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
  if (commandId !== 'sheet.command.undo' && commandId !== 'sheet.command.redo' && !finishEdit()) return
  const commands: Record<string, () => void | Promise<void>> = {
    'sheet.command.undo': undo, 'sheet.command.redo': redo,
    'sheet.command.copy': () => { copySelection() }, 'sheet.command.copy-formulas': () => { copySelection(true) }, 'sheet.command.cut': cutSelection,
    'sheet.command.paste': () => pasteInternal('all'),
    'sheet.command.paste-formulas': () => pasteInternal('formulas'),
    'sheet.command.paste-values': () => pasteInternal('values'),
    'sheet.command.paste-formats': () => pasteInternal('formats'),
    'sheet.command.paste-column-widths': () => pasteInternal('columnWidths'),
    'sheet.command.paste-except-borders': () => pasteInternal('exceptBorders'),
    'sheet.command.clear-content': clearSelection,
    'sheet.command.clear-all': () => {
      if (!activeSheet.value || !checkDirectSelection()) return
      if (!validateIncomingValues(selectedExistingKeys().map(key => ({ key, value: '' })))) return
      selectedExistingKeys().forEach(key => { delete activeSheet.value?.cells[key] })
      recalculateFormulas()
      commitHistory()
    },
    'sheet.command.set-calculation-mode': () => {
      if ((params !== 'automatic' && params !== 'manual') || !props.studio || !workbook.value.id) return
      const book = FOffice.newAPI(props.studio).getWorkbook(workbook.value.id)
      if (!book) return
      book.setCalculationMode(params)
      const model = props.studio.getWorkbookModel(workbook.value.id)
      if (model) workbook.value = toExcelWorkbookView(model)
      commitHistory()
    },
    'sheet.command.toggle-bold': () => toggleStyle('bold'), 'sheet.command.toggle-italic': () => toggleStyle('italic'),
    'sheet.command.toggle-underline': () => toggleStyle('underline'),
    'sheet.command.toggle-strike': () => toggleStyle('strike'),
    'sheet.command.set-font-family': () => { if (typeof params === 'string') setCellFormat('fontFamily', params) },
    'sheet.command.set-font-size': () => { if (typeof params === 'number' && params > 0) setCellFormat('fontSize', params) },
    'sheet.command.increase-font-size': () => changeFontSize(1),
    'sheet.command.decrease-font-size': () => changeFontSize(-1),
    'sheet.command.set-number-format': () => { if (typeof params === 'string' && Object.hasOwn(SHEET_NUMBER_FORMATS, params)) setNumberFormat(params as SheetNumberFormat) },
    'sheet.command.increase-decimals': () => changeSelectedCells(cell => adjustSheetDecimals(cell, 1)),
    'sheet.command.decrease-decimals': () => changeSelectedCells(cell => adjustSheetDecimals(cell, -1)),
    'sheet.command.clear-format': () => changeSelectedCells(clearSheetCellFormat),
    'sheet.command.set-vertical-align': () => { if (params === 'top' || params === 'middle' || params === 'bottom') setCellFormat('vertical', params) },
    'sheet.command.set-borders': () => { if (isSheetBorderPreset(params)) setBorders(params) },
    'sheet.command.set-border-color': () => { if (typeof params === 'string' && /^#[0-9a-f]{6}$/i.test(params)) changeBorderOptions({ ...borderOptions.value, color: params }) },
    'sheet.command.set-border-style': () => { if (typeof params === 'string' && Object.hasOwn(sheetBorderStyleLabels, params)) changeBorderOptions({ ...borderOptions.value, style: params as SheetBorderOptions['style'] }) },
    'sheet.command.set-color': () => { if (typeof params === 'string') setCellFormat('color', params) },
    'sheet.command.reset-text-color': () => resetSelectionColor('sheet.command.reset-text-color'),
    'sheet.command.reset-background-color': () => resetSelectionColor('sheet.command.reset-background-color'),
    'sheet.command.set-background': () => { if (typeof params === 'string') setBackground(params) },
    'sheet.command.set-align': () => { if (params === 'left' || params === 'center' || params === 'right') setAlign(params) },
    'sheet.command.toggle-wrap': () => changeSelectedCells(cell => setSheetTextFit(cell, 'wrap', selectionFormat.value.wrap !== true)),
    'sheet.command.set-wrap-strategy': () => {
      if (params !== 'overflow' && params !== 'wrap' && params !== 'clip') return
      changeSelectedCells(cell => ({ ...cell, wrap: params === 'wrap', wrapStrategy: params }))
    },
    'sheet.command.set-shrink-to-fit': () => {
      if (typeof params === 'boolean') changeSelectedCells(cell => setSheetTextFit(cell, 'shrinkToFit', params))
    },
    'sheet.command.set-text-rotation': () => {
      if (params === 'v') setTextRotation(params)
      else if (typeof params === 'number' && [-90, -45, 0, 45, 90].includes(params)) setTextRotation(params)
    },
    'sheet.command.merge': () => mergeSelection(params === 'vertical' || params === 'horizontal' ? params : 'all'), 'sheet.command.unmerge': unmergeSelection,
    'sheet.command.insert-row': () => insertRow(params), 'sheet.command.insert-column': () => insertColumn(params),
    'sheet.command.delete-row': () => deleteRow(params), 'sheet.command.delete-column': () => deleteColumn(params),
    'sheet.command.insert-range-move-right': () => requestCellRangeMove('right'),
    'sheet.command.insert-range-move-down': () => requestCellRangeMove('down'),
    'sheet.command.delete-range-move-left': () => requestCellRangeMove('left'),
    'sheet.command.delete-range-move-up': () => requestCellRangeMove('up'),
    'sheet.command.hide-row': () => hideAxisTarget('row', params),
    'sheet.command.hide-column': () => hideAxisTarget('column', params),
    'sheet.command.show-hidden-rows': () => showHiddenAxisTargets('row', params),
    'sheet.command.show-hidden-columns': () => showHiddenAxisTargets('column', params),
    'sheet.command.auto-fit-row': () => autoFitAxisTarget('row', params),
    'sheet.command.auto-fit-column': () => autoFitAxisTarget('column', params),
    'sheet.command.set-row-height': () => { axisSizeDialogKind.value = 'row'; openAxisSizeDialog('row', params) },
    'sheet.command.set-worksheet-col-width': () => { axisSizeDialogKind.value = 'column'; openAxisSizeDialog('column', params) },
    'sheet.command.text-to-number': convertSelectedTextCells,
    'sheet.command.toggle-auto-filter': () => runAutoFilterCommand('sheet.command.toggle-auto-filter'),
    'sheet.command.clear-filter': () => runAutoFilterCommand('sheet.command.clear-filter'),
    'sheet.command.reapply-filter': () => runAutoFilterCommand('sheet.command.reapply-filter'),
    'sheet.operation.insert-chart': openChartDialog,
    'sheet.operation.insert-hyperlink': openHyperlinkDialog,
    'sheet.operation.open-pivot': openPivotPanel,
    'sheet.operation.open-pivot-chart': openPivotChartPanel,
    'sheet.operation.open-slicer': () => openSlicerPanel('slicer'),
    'sheet.operation.open-timeline': () => openSlicerPanel('timeline'),
    'sheet.operation.open-data-panel': () => openDataPanel(params === 'sort' ? 'sort' : 'filter'),
    'sheet.operation.open-data-validation': openDataValidationPanel,
    'sheet.operation.add-data-validation': () => applyQuickDataValidation('add-default'),
    'sheet.operation.quick-data-validation': () => applyQuickDataValidation(params),
    'sheet.operation.open-conditional-formatting': openConditionalFormattingPanel,
    'sheet.operation.open-sparkline': openSparklinePanel,
    'sheet.operation.format-painter': armFormatPainter,
    'sheet.operation.open-custom-number-format': () => { customNumberFormatOpen.value = true },
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
    'sheet.operation.outline-summary-position': () => {
      if (typeof params !== 'object' || params === null || !('axis' in params) || !('atStart' in params)) return
      const value = params as { axis?: unknown; atStart?: unknown }
      if ((value.axis === 'row' || value.axis === 'column') && typeof value.atStart === 'boolean') runOutlineCommand({ kind: 'summary-position', axis: value.axis, atStart: value.atStart })
    },
    'sheet.operation.outline-toggle-group': () => {
      if (typeof params !== 'object' || params === null || !('axis' in params) || !('groupId' in params)) return
      const value = params as { axis?: unknown; groupId?: unknown }
      if ((value.axis === 'row' || value.axis === 'column') && typeof value.groupId === 'string') runOutlineCommand({ kind: 'toggle-group', axis: value.axis, groupId: value.groupId })
    },
    'sheet.operation.outline-set-level': () => {
      if (typeof params !== 'object' || params === null || !('axis' in params) || !('level' in params)) return
      const value = params as { axis?: unknown; level?: unknown }
      if ((value.axis === 'row' || value.axis === 'column') && typeof value.level === 'number' && Number.isSafeInteger(value.level)) runOutlineCommand({ kind: 'set-level', axis: value.axis, level: value.level })
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
  const command = commands[commandId]
  if (!command) return
  try { await command() } catch (error) { emit('notify', { type: 'error', message: error instanceof Error ? error.message : '操作失败' }); return }
  if (!commandId.startsWith('sheet.operation.') || commandId.startsWith('sheet.operation.zoom') || commandId === 'sheet.operation.set-zoom') void nextTick(() => spreadsheetCanvas.value?.focus())
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

watch(selectedKey, () => { nameError.value = ''; void nextTick(scrollToSelected) })
watch([() => workbook.value.id, activeSheetId], () => { validationPanelOpen.value = false })
watch(hyperlinkDialogOpen, open => { if (!open) void nextTick(() => spreadsheetCanvas.value?.focus()) })
watch(() => workbook.value.id, id => { if (hyperlinkCell.value?.unitId !== id) hyperlinkDialogOpen.value = false })
defineExpose({ save: saveWorkbook, focus: () => spreadsheetCanvas.value?.focus() })

onMounted(() => {
  document.addEventListener('fullscreenchange', syncFullscreen)
  window.addEventListener('mouseup', stopSelecting)
  window.addEventListener('keydown', onWindowKeydown)
})

onUnmounted(() => {
  sortController?.abort()
  document.removeEventListener('fullscreenchange', syncFullscreen)
  window.removeEventListener('mouseup', stopSelecting)
  window.removeEventListener('keydown', onWindowKeydown)
})
onDeactivated(() => {
  validationPanelOpen.value = false
  sortController?.abort()
  if (document.fullscreenElement === workbenchRef.value) {
    void document.exitFullscreen().catch((error: unknown) => emit('notify', { type: 'error', message: error instanceof Error ? error.message : '无法退出全屏' }))
  }
})

</script>

<template>
  <ConfigProvider>
  <section ref="workbenchRef" class="excel-workbench" @keydown="handleGridKeydown">
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

    <WorkbenchRibbon v-if="chrome.toolbar" v-model="activeTab" :tabs="ribbonTabs" :variant="props.ribbonVariant" @execute="executeRibbon" />
    <SheetFormulaBar v-if="chrome.header" ref="formulaInput" :address="selectedAddress" :value="formulaValue" :editing="Boolean(editingKey)" :error="nameError" @update:value="editValue = $event" @begin="beginFormulaEdit" @confirm="finishFormulaEdit" @cancel="cancelFormulaEdit" @navigate="navigateToAddress" @focus-grid="nameError = ''; spreadsheetCanvas?.focus()" />
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

    <div class="sheet-canvas-layout">
      <div class="sheet-canvas-stack">
    <SpreadsheetCanvas v-if="activeSheet" ref="spreadsheetCanvas" v-model:edit-value="editValue" :sheet="activeSheet" :selected-key="selectedKey" :selected-range="emptyCellSelection" :selected-ranges="selectedRanges" :editing-key="editingKey" :zoom="zoom" :current-user-id="studio?.currentUserId ?? 'local-user'" :search-text="searchText" :validation-options="validationOptions" :overlay-commands="chartCommands" :slicers="slicerControls" :selected-slicer-id="selectedSlicerId" @select-cell="startCellSelection" @extend-selection="extendCellSelection" @fill-selection="fillSelection" @select-header="selectHeader" @edit-cell="beginEdit" @finish-edit="finishEdit" @editor-keydown="handleGridKeydown" @paste="handlePaste" @resize-axis="resizeAxis" @context-command="executeSheetContextCommand" @open-hyperlink="openCellHyperlink" @select-slicer="selectedSlicerId = $event" @slicer-item="handleSlicerItem" @slicer-clear="clearSlicer" @slicer-range="setSlicerTimelineRange" @slicer-scroll-to="scrollSlicerTo" @slicer-level="({ slicerId, level }) => updateSlicerPresentation(slicerId, { level })" @slicer-settings="openSlicerSettings" @slicer-move="moveSlicer" @slicer-resize="resizeSlicer" />
        <ChartOverlay v-if="activeSheet" :workbook="chartWorkbook" :active-sheet-id="activeSheet.id" :zoom="zoom" @pointerdown.capture="rememberChartPointer" @click.capture="focusChartFromOverlay" @overlay-commands="chartCommands = $event" @update-chart="updateChartLayout" @remove-chart="removeChart" />
      </div>
      <SheetDataValidationPanel v-if="validationPanelOpen && activeSheet && studio && workbook.id" v-model:open="validationPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="validationPanelRange" :initial-type="validationPanelInitialType" :initial-rule-id="validationPanelRuleId" :canvas-ranges="selectedModelRanges()" :can-add="canEditValidationSelection" @add="applyQuickDataValidation('add-default')" @applied="refreshValidationView" />
    </div>

    <ChartEditorPanel v-model:open="chartEditorOpen" :chart="focusedChart" :sheet="activeSheet" @apply="applyChartSettings" />

    <Modal v-model:open="chartModalOpen" title="插入图表" ok-text="插入" cancel-text="取消" :width="440" @ok="confirmChartDialog">
      <Form layout="vertical" class="chart-insert-form" @submit.prevent>
        <FormItem label="数据区域"><Input v-model:value="chartSource" placeholder="A1:C4" /></FormItem>
        <FormItem label="图表类型"><Select v-model:value="chartType" :options="chartTypeOptions" /></FormItem>
        <FormItem label="系列方向"><Select v-model:value="chartOrientation" :options="[{ label: '按列', value: 'columns' }, { label: '按行', value: 'rows' }]" /></FormItem>
        <FormItem :label="chartType === 'scatter' ? '共享 X 字段' : '分类字段'"><Select v-model:value="chartCategoryField" :options="chartFieldOptions" /></FormItem>
        <FormItem :label="chartType === 'scatter' ? 'Y 系列字段' : '数值系列字段'"><Select v-model:value="chartValueFields" mode="multiple" :options="chartFieldOptions" placeholder="留空时使用其余字段" /></FormItem>
        <template v-if="chartType === 'scatter'">
          <FormItem v-for="field in chartValueFields" :key="field" :label="`${chartFieldOptions.find(option => option.value === field)?.label ?? field} 的 X 字段`"><Select v-model:value="chartSeriesXFields[field]" :options="chartFieldOptions" allow-clear placeholder="使用共享 X 字段" /></FormItem>
          <FormItem label="X 轴刻度"><Select v-model:value="chartXScale" :options="[{ label: '线性', value: 'linear' }, { label: '对数', value: 'log' }]" /></FormItem>
          <FormItem label="Y 轴刻度"><Select v-model:value="chartYScale" :options="[{ label: '线性', value: 'linear' }, { label: '对数', value: 'log' }]" /></FormItem>
        </template>
        <FormItem label="图表标题"><Input v-model:value="chartTitle" placeholder="可选" /></FormItem>
        <p v-if="chartError" role="alert" class="chart-insert-error">{{ chartError }}</p>
        <p class="chart-insert-hint">{{ chartOrientation === 'columns' ? '数据区域首行作为字段名。' : '数据区域首列作为字段名。' }}{{ chartType === 'scatter' ? 'X 与 Y 按数据位置配对，空值和无效数值不绘制。对数轴仅绘制正数。' : chartType === 'pie' || chartType === 'doughnut' ? '请选择一个数值系列。' : '' }}</p>
      </Form>
    </Modal>

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
    <Modal v-model:open="axisSizeDialogOpen" :title="axisSizeDialogKind === 'row' ? '行高' : '列宽'" ok-text="确定" cancel-text="取消" :width="360" @ok="confirmAxisSize">
      <Form layout="vertical" @submit.prevent>
        <FormItem :label="axisSizeDialogKind === 'row' ? '行高（像素）' : '列宽（像素）'">
        <InputNumber v-model:value="axisSizeValue" :min="0" :max="1000" :precision="0" :aria-label="axisSizeDialogKind === 'row' ? '行高' : '列宽'" @press-enter="confirmAxisSize" />
        </FormItem>
      </Form>
      <p v-if="axisSizeError" class="chart-insert-error" role="alert">{{ axisSizeError }}</p>
    </Modal>
    <Modal v-model:open="cellMoveConfirmOpen" title="合并单元格" ok-text="继续" cancel-text="取消" :width="420" @ok="confirmCellRangeMove" @cancel="pendingCellMove = null">
      <p>此操作会拆分与移位区域交叉的合并单元格，是否继续？</p>
    </Modal>
    <SheetFooter v-if="chrome.footer" :sheets="availableSheets" :active-sheet-id="activeSheetId" :zoom="zoom" :statistics="selectedStatistics" :number-format="activeCell.rawNumberFormat ?? activeCell.numberFormat" @activate="activateSheet" @add="addSheet" @rename="renameSheet" @copy="copySheet" @tab-color="setSheetTabColor" @protect-sheet="openWorksheetProtection" @unprotect-sheet="removeWorksheetProtection" @delete="deleteSheet" @hide="hideSheet" @show="showSheet" @update:zoom="zoom = $event" />
    <PivotTablePanel v-if="pivotPanelOpen && activeSheet" v-model:open="pivotPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="pivotSource" />
    <PivotChartPanel v-if="pivotChartPanelOpen && activeSheet && studio && workbook.id" v-model:open="pivotChartPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" />
    <SlicerPanel v-if="slicerPanelOpen && activeSheet && studio && workbook.id" v-model:open="slicerPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="slicerSourceRange" :initial-kind="slicerKind" :selected-slicer-id="selectedSlicerId" @changed="handleSlicerPanelChanged" />
    <SheetDataPanel v-if="dataPanelOpen && activeSheet && studio && workbook.id" v-model:open="dataPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="dataPanelRange" :initial-mode="dataPanelMode" />
    <ConditionalFormattingPanel v-if="conditionalFormattingOpen && activeSheet" v-model:open="conditionalFormattingOpen" :rules="activeSheet.conditionalFormats ?? []" :default-range="conditionalFormattingRange" @apply="applyConditionalFormatting" />
    <SheetRangeProtectionPanel v-if="protectionPanelOpen && activeSheet && studio && workbook.id" v-model:open="protectionPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="validationPanelRange" :start-new="protectionStartNew" />
    <SheetProtectionPanel v-if="worksheetProtectionSheetId && studio && workbook.id" :open="Boolean(worksheetProtectionSheetId)" :studio="studio" :workbook-id="workbook.id" :sheet-id="worksheetProtectionSheetId" :rule="workbook.sheets.find(sheet => sheet.id === worksheetProtectionSheetId)?.worksheetProtection" @update:open="open => { if (!open) worksheetProtectionSheetId = null }" @applied="refreshWorksheetProtection" />
    <SparklinePanel v-if="sparklinePanelOpen && activeSheet && studio && workbook.id" :open="sparklinePanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :target-range="sparklineTargetRange" @update:open="sparklinePanelOpen = $event" />
      <CustomNumberFormatPanel v-model:open="customNumberFormatOpen" :date1904="activeSheet?.date1904 ?? workbook.date1904" :current-pattern="activeCell.rawNumberFormat ?? 'General'" @apply="(pattern) => changeSelectedCells(cell => ({ ...cell, rawNumberFormat: pattern }))" />
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
.sheet-canvas-stack { position: relative; display: flex; flex: 1; min-width: 0; min-height: 0; overflow: hidden; }
.sheet-canvas-stack :deep(.spreadsheet-canvas) { flex: 1; }
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
