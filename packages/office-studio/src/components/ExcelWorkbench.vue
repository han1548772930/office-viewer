<script setup lang="ts">
import { computed, defineAsyncComponent, nextTick, onDeactivated, onMounted, onUnmounted, ref, watch } from 'vue'
import { Button, Input, Modal, Select, Tooltip } from 'antdv-next'
import { FUniver } from '../api/facade'
import type { OfficeStudio } from '../api/office-studio'
import type { ExcelCellView, ExcelWorkbookView as DomainExcelWorkbookView } from '../services/excelViewAdapter'
import { applyExcelWorkbookView } from '../services/excelViewAdapter'
import { createBlankWorkbook } from '../domain/model'
import type { ChartModel } from '../domain/model'
import SpreadsheetCanvas from './spreadsheet/SpreadsheetCanvas.vue'
import ChartOverlay from './ChartOverlay.vue'
import { chartPositionForRange, type ExcelWorkbookChartView } from './chart-overlay-model'
import type { DrawCommand } from '../rendering/contracts'
import type { GridHit, GridRange, SpreadsheetRenderSheet } from '../rendering/spreadsheet/layout'
import { addRange, containsCell, selectionBounds, selectionCount, subtractRange } from '../rendering/spreadsheet/selection'
import { Download, FileSpreadsheet, FolderOpen, Save, Search, X } from 'lucide-vue-next'
import WorkbenchRibbon from '../platform/ribbon/WorkbenchRibbon.vue'
import SheetFormulaBar from './spreadsheet/SheetFormulaBar.vue'
import SheetFooter from './spreadsheet/SheetFooter.vue'
import { createSheetRibbon } from './spreadsheet/sheet-ribbon'
import { calculateSelectionStatistics, createUniqueSheetName, normalizeSheetViewState, parseSheetAddress, validateSheetName, type SheetViewState } from '../features/sheet-workbench/sheet-workbench'
import { recalculateSheetViewFormulas } from '../features/sheet-workbench/sheet-view-formulas'
import { adjustSheetDecimals, applySheetBorders, applySheetNumberFormat, clearSheetCellFormat, nextSheetFontSize, setSheetTextFit, type SheetBorderPreset } from '../features/sheet-workbench/sheet-toolbar-format'
import { classifySheetNumberFormat, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from '../features/sheet-workbench/sheet-number-format'
import { buildSheetFunctionFormula, SHEET_FUNCTION_CATALOG } from '../features/sheet-workbench/sheet-function-catalog'
import { findVisibleSheetRow, hiddenSheetRows } from '../features/sheet-workbench/sheet-row-visibility'
import { validateSheetDataValidationChanges } from '../features/sheet-validation/validation-engine.ts'

const PivotTablePanel = defineAsyncComponent(() => import('./spreadsheet/PivotTablePanel.vue'))
const SlicerPanel = defineAsyncComponent(() => import('./spreadsheet/SlicerPanel.vue'))
const SheetDataPanel = defineAsyncComponent(() => import('./spreadsheet/SheetDataPanel.vue'))
const SheetDataValidationPanel = defineAsyncComponent(() => import('./spreadsheet/SheetDataValidationPanel.vue'))

export type ExcelCell = ExcelCellView
export type ExcelSheet = SpreadsheetRenderSheet
export type ExcelWorkbook = Omit<DomainExcelWorkbookView, 'sheets'> & { sheets: ExcelSheet[]; charts?: ChartModel[] }

const props = withDefaults(defineProps<{ modelValue?: ExcelWorkbook; model?: ExcelWorkbook; canUndo?: boolean; canRedo?: boolean; showHeader?: boolean; studio?: OfficeStudio }>(), {
  modelValue: undefined,
  model: undefined,
  canUndo: undefined,
  canRedo: undefined,
  showHeader: false,
  studio: undefined,
})

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
const activeSheetId = ref(workbook.value.sheets.find(sheet => sheet.id === workbook.value.activeSheetId && sheet.visible !== false)?.id ?? workbook.value.sheets.find(sheet => sheet.visible !== false)?.id ?? '')
const selectedKey = ref('A2')
const selectionAnchor = ref('A2')
const selectedRanges = ref<GridRange[]>([{ startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }])
const emptyCellSelection: string[] = []
const editingKey = ref<string | null>(null)
const editValue = ref('')
const selecting = ref(false)
const formulaInput = ref<InstanceType<typeof SheetFormulaBar> | null>(null)
const spreadsheetCanvas = ref<InstanceType<typeof SpreadsheetCanvas> | null>(null)
const fileInput = ref<HTMLInputElement | null>(null)
const zoom = ref(100)
const searchText = ref('')
const activeTab = ref('home')
const nameError = ref('')
const searchOpen = ref(false)
const searchInput = ref<{ focus: () => void } | null>(null)
const renameSheetId = ref<string | null>(null)
const renameValue = ref('')
const renameError = ref('')
const viewSessions = new Map<string, SheetViewState>()
const copiedCells = ref<ExcelCell[][] | null>(null)
const chartCommands = ref<DrawCommand[]>([])
const chartModalOpen = ref(false)
const chartSource = ref('')
const chartTitle = ref('')
const chartType = ref<'bar' | 'line'>('bar')
const chartError = ref('')
const pivotPanelOpen = ref(false)
const pivotSource = ref('')
const slicerPanelOpen = ref(false)
const slicerKind = ref<'slicer' | 'timeline'>('slicer')
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
const functionEntry = computed(() => SHEET_FUNCTION_CATALOG.find(entry => entry.name === functionName.value))
const functionOptions = computed(() => SHEET_FUNCTION_CATALOG
  .filter(entry => functionCategory.value === null || entry.category === functionCategory.value)
  .map(entry => ({ label: `${entry.name} · ${entry.label}`, value: entry.name })))
const sortBusy = ref(false)
const sortProgress = ref('')
let sortController: AbortController | undefined
const chartWorkbook = computed(() => workbook.value as ExcelWorkbookChartView)

const activeSheet = computed(() => workbook.value.sheets.find((sheet) => sheet.id === activeSheetId.value && sheet.visible !== false) ?? workbook.value.sheets.find(sheet => sheet.visible !== false))
const activeCell = computed(() => activeSheet.value?.cells[selectedKey.value] ?? { value: '' })
const selectedAddress = computed(() => {
  const range = selectedRanges.value[0]
  return selectedRanges.value.length === 1 && range && (range.startRow !== range.endRow || range.startColumn !== range.endColumn) ? rangeNotation(range) : selectedKey.value
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
    color: common('color', '#1f2937'), background: common('background', '#ffffff'),
    rotation: common('rotation', 0) ?? null, rotationVertical: common('rotationVertical', false),
    shrinkToFit: common('shrinkToFit', false),
  }
})
const validationOptions = computed(() => {
  const workbookId = workbook.value.id
  const sheetId = activeSheet.value?.id
  if (!props.studio || !workbookId || !sheetId) return []
  const [column, row] = splitKey(editingKey.value ?? selectedKey.value)
  return FUniver.newAPI(props.studio).getWorkbook(workbookId)?.getSheetBySheetId(sheetId)?.getCellValidationOptions(row - 1, column) ?? []
})
const canUnmerge = computed(() => {
  const [column, row] = splitKey(selectedKey.value)
  return activeSheet.value?.merges?.some(merge => { const range = parseChartRange(merge); return range && containsCell([range], row - 1, column) }) ?? false
})
const ribbonTabs = computed(() => createSheetRibbon({
  ...selectionFormat.value, canUndo: canUndo.value, canRedo: canRedo.value, hasSheet: Boolean(activeSheet.value),
  canMerge: selectedCount.value > 1, canUnmerge: canUnmerge.value,
  canDeleteRow: totalRows.value > 1, canDeleteColumn: (activeSheet.value?.columnCount ?? 0) > 1,
  chartAvailable: Boolean(props.studio && workbook.value.id), gridlines: activeSheet.value?.showGridlines !== false,
  dataValidationAvailable: Boolean(props.studio && workbook.value.id),
  canSort: Boolean(!sortBusy.value && props.studio && workbook.value.id && selectedRanges.value.length === 1 && selectedRanges.value[0] && selectedRanges.value[0].endRow > selectedRanges.value[0].startRow),
  frozen: isFrozen.value, zoom: zoom.value,
}))
const formulaValue = computed(() => editingKey.value ? editValue.value : activeCell.value.formula ?? activeCell.value.value)
const searchMatches = computed(() => {
  if (!searchText.value) return []
  const needle = searchText.value.toLocaleLowerCase()
  const hidden = hiddenSheetRows(activeSheet.value ?? { rowCount: 0 })
  return Object.entries(sheetCells.value).filter(([key, cell]) => !hidden.has(splitKey(key)[1] - 1) && String(cell.formula ? cell.cachedValue ?? '' : cell.value).toLocaleLowerCase().includes(needle)).map(([key]) => key)
})

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
  const requested = value.sheets.find(sheet => sheet.id === value.activeSheetId && sheet.visible !== false)
  const nextSheetId = requested?.id ?? value.sheets.find(sheet => sheet.id === previousSheetId && sheet.visible !== false)?.id ?? value.sheets.find(sheet => sheet.visible !== false)?.id ?? ''
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
  zoom.value = next.zoom
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
  chartError.value = ''
  chartModalOpen.value = true
}

function confirmChartDialog() {
  const sheet = activeSheet.value
  const range = parseChartRange(chartSource.value)
  if (!sheet || !range) { chartError.value = '请输入有效的数据范围，例如 A1:C4'; return }
  try {
    const target = getChartWorksheet(sheet.id)
    if (!target) throw new Error('图表命令上下文尚未绑定到活动工作簿')
    const position = chartPositionForRange(range, sheet)
    target.insertChart(target.newChart(chartType.value)
      .setSource(chartSource.value, 'columns')
      .setAbsolutePosition(position.x, position.y)
      .setSize(480, 300)
      .setTitle(chartTitle.value.trim())
      .build())
    chartModalOpen.value = false
  } catch (error) {
    chartError.value = error instanceof Error ? error.message : '图表数据范围不可用'
  }
}

function getChartWorksheet(sheetId: string) {
  if (!props.studio || !workbook.value.id) return null
  return FUniver.newAPI(props.studio).getWorkbook(workbook.value.id)?.getSheetBySheetId(sheetId) ?? null
}

async function updateChartLayout(payload: { chartId: string; patch: Partial<Pick<ChartModel, 'position' | 'size'>> }) {
  const chart = getChartWorksheet(activeSheet.value?.id ?? '')?.getChartById(payload.chartId)
  if (chart) await chart.update(payload.patch)
}

async function removeChart(chartId: string) {
  const chart = getChartWorksheet(activeSheet.value?.id ?? '')?.getChartById(chartId)
  if (chart) await chart.remove()
}

function setCellValue(key: string, value: string): boolean {
  const sheet = activeSheet.value
  if (!sheet) return false
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
  const previous = workbook.value.id ? props.studio?.getWorkbookModel(workbook.value.id) : undefined
  const result = recalculateSheetViewFormulas(workbook.value, previous)
  workbook.value = result.workbook
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

function selectCell(key: string, extend = false) {
  const [column, requestedRow] = splitKey(key)
  const visibleRow = activeSheet.value ? findVisibleSheetRow(activeSheet.value, requestedRow - 1, 1) : undefined
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
  if (target?.matches('input, textarea, select') && !target.classList.contains('canvas-cell-input') && !target.classList.contains('formula-input')) return
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
  if (!checkDirectSelection()) return false
  for (const range of selectedRanges.value) for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) action(`${columnName(column)}${row + 1}`)
  return true
}
function selectedExistingKeys(): string[] {
  return Object.keys(sheetCells.value).filter(key => { const [column, row] = splitKey(key); return containsCell(selectedRanges.value, row - 1, column) })
}
function copySelection(): boolean {
  const bounds = selectionBounds(selectedRanges.value)
  if (!activeSheet.value || !bounds || !checkDirectSelection(selectionCount([bounds]))) return false
  copiedCells.value = Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, rowOffset) =>
    Array.from({ length: bounds.endColumn - bounds.startColumn + 1 }, (_, columnOffset) => ({ ...getCell(`${columnName(bounds.startColumn + columnOffset)}${bounds.startRow + rowOffset + 1}`) })),
  )
  const text = copiedCells.value.map((row) => row.map((cell) => cell.formula ?? cell.value).join('\t')).join('\n')
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
  if (!text && copiedCells.value) text = copiedCells.value.map((row) => row.map((cell) => cell.formula ?? cell.value).join('\t')).join('\n')
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
  applySheetBorders(activeSheet.value.cells, selectedRanges.value, preset)
  commitHistory()
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

function mergeSelection() {
  const bounds = selectionBounds(selectedRanges.value)
  if (!activeSheet.value || selectedCount.value < 2 || !bounds) return
  const range = `${columnName(bounds.startColumn)}${bounds.startRow + 1}:${columnName(bounds.endColumn)}${bounds.endRow + 1}`
  activeSheet.value.merges ??= []
  activeSheet.value.merges.push(range)
  const master = `${columnName(bounds.startColumn)}${bounds.startRow + 1}`
  selectedExistingKeys().filter((key) => key !== master).forEach((key) => delete activeSheet.value?.cells[key])
  selectCell(master)
  commitHistory()
}

function transformMerges(axis: 'row' | 'column', index: number, delta: 1 | -1) {
  if (!activeSheet.value?.merges) return
  activeSheet.value.merges = activeSheet.value.merges.flatMap((range) => {
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

function insertRow() {
  if (!activeSheet.value) return
  const [, selectedRow] = splitKey(selectedKey.value)
  const updated: Record<string, ExcelCell> = {}
  Object.entries(activeSheet.value.cells).forEach(([key, cell]) => {
    const [column, row] = splitKey(key)
    updated[`${columnName(column)}${row >= selectedRow ? row + 1 : row}`] = cell
  })
  activeSheet.value.cells = updated
  transformMerges('row', selectedRow, 1)
  activeSheet.value.rowCount += 1
  moveSelection(1, 0)
  commitHistory()
}

function deleteRow() {
  if (!activeSheet.value || totalRows.value <= 1) return
  const [, selectedRow] = splitKey(selectedKey.value)
  const updated: Record<string, ExcelCell> = {}
  Object.entries(activeSheet.value.cells).forEach(([key, cell]) => {
    const [column, row] = splitKey(key)
    if (row < selectedRow) updated[key] = cell
    else if (row > selectedRow) updated[`${columnName(column)}${row - 1}`] = cell
  })
  activeSheet.value.cells = updated
  transformMerges('row', selectedRow, -1)
  activeSheet.value.rowCount -= 1
  selectedKey.value = `${columnName(splitKey(selectedKey.value)[0])}${Math.min(selectedRow, totalRows.value)}`
  selectCell(selectedKey.value)
  commitHistory()
}

function insertColumn() {
  if (!activeSheet.value) return
  const [selectedColumn] = splitKey(selectedKey.value)
  const updated: Record<string, ExcelCell> = {}
  Object.entries(activeSheet.value.cells).forEach(([key, cell]) => {
    const [column, row] = splitKey(key)
    updated[`${columnName(column >= selectedColumn ? column + 1 : column)}${row}`] = cell
  })
  activeSheet.value.cells = updated
  transformMerges('column', selectedColumn, 1)
  activeSheet.value.columnCount += 1
  moveSelection(0, 1)
  commitHistory()
}

function deleteColumn() {
  if (!activeSheet.value || activeSheet.value.columnCount <= 1) return
  const [selectedColumn] = splitKey(selectedKey.value)
  const updated: Record<string, ExcelCell> = {}
  Object.entries(activeSheet.value.cells).forEach(([key, cell]) => {
    const [column, row] = splitKey(key)
    if (column < selectedColumn) updated[key] = cell
    else if (column > selectedColumn) updated[`${columnName(column - 1)}${row}`] = cell
  })
  activeSheet.value.cells = updated
  transformMerges('column', selectedColumn, -1)
  activeSheet.value.columnCount -= 1
  selectedKey.value = `${columnName(Math.min(selectedColumn, activeSheet.value.columnCount - 1))}${splitKey(selectedKey.value)[1]}`
  selectCell(selectedKey.value)
  commitHistory()
}

function saveViewSession() {
  if (activeSheetId.value) viewSessions.set(activeSheetId.value, { key: selectedKey.value, anchor: selectionAnchor.value, ranges: selectedRanges.value.map(range => ({ ...range })), zoom: zoom.value })
}

function activateSheet(id: string) {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id && candidate.visible !== false)
  if (!sheet || id === activeSheetId.value) return
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
  zoom.value = next.zoom
  if (props.studio && workbook.value.id && props.studio.getWorkbookModel(workbook.value.id)?.sheets.some(candidate => candidate.id === id)) {
    FUniver.newAPI(props.studio).getWorkbook(workbook.value.id)?.setActiveSheet(id)
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

function switchNextSheet() {
  const sheets = workbook.value.sheets.filter(sheet => sheet.visible !== false)
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
  sheet.name = renameValue.value.trim()
  renameSheetId.value = null
  commitHistory()
}

function deleteSheet(id: string) {
  const sheets = workbook.value.sheets
  const sheet = sheets.find(candidate => candidate.id === id)
  if (!sheet || (sheet.visible !== false && sheets.filter(candidate => candidate.visible !== false).length <= 1)) return
  finishEdit()
  const remaining = sheets.filter(candidate => candidate.id !== id)
  if (activeSheetId.value === id) {
    const next = remaining.find(candidate => candidate.visible !== false)
    if (next) activateSheet(next.id)
  }
  workbook.value.sheets = remaining
  viewSessions.delete(id)
  commitHistory()
}

function hideSheet(id: string) {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id)
  const next = workbook.value.sheets.find(candidate => candidate.id !== id && candidate.visible !== false)
  if (!sheet || !next) return
  if (id === activeSheetId.value) activateSheet(next.id)
  sheet.visible = false
  commitHistory()
}

function showSheet(id: string) {
  const sheet = workbook.value.sheets.find(candidate => candidate.id === id)
  if (!sheet) return
  sheet.visible = true
  activateSheet(id)
  commitHistory()
}

function openFileDialog() {
  fileInput.value?.click()
}

function handleFile(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (file) emit('import', file)
  else emit('notify', { type: 'warning', message: '未选择文件' })
  if (fileInput.value) fileInput.value.value = ''
}

function saveWorkbook() {
  if (!finishEdit()) return
  emit('save', cloneWorkbook(workbook.value))
}

function exportWorkbook() {
  if (!finishEdit()) return
  emit('export', cloneWorkbook(workbook.value))
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
  selecting.value = false
}

function freeze(rows: number, columns: number) {
  if (!activeSheet.value) return
  activeSheet.value.freeze = { rows, columns }
  commitHistory()
}

function openSearch() {
  searchOpen.value = true
  void nextTick(() => searchInput.value?.focus())
}
function findNext() {
  if (!searchMatches.value.length) return
  const next = searchMatches.value[(searchMatches.value.indexOf(selectedKey.value) + 1) % searchMatches.value.length]
  if (next) { selectCell(next); spreadsheetCanvas.value?.ensureVisible(next) }
}
function closeSearch() { searchOpen.value = false; searchText.value = ''; spreadsheetCanvas.value?.focus() }

function openPivotPanel() {
  const bounds = selectionBounds(selectedRanges.value)
  pivotSource.value = bounds ? rangeNotation(bounds) : 'A1:C4'
  pivotPanelOpen.value = true
}

function openSlicerPanel(kind: 'slicer' | 'timeline') { slicerKind.value = kind; slicerPanelOpen.value = true }
function openDataPanel(mode: 'filter' | 'sort'): void {
  const bounds = selectionBounds(selectedRanges.value)
  dataPanelMode.value = mode
  dataPanelRange.value = bounds ? rangeNotation(bounds) : selectedKey.value
  dataPanelOpen.value = true
}
function openDataValidationPanel(): void {
  const bounds = selectionBounds(selectedRanges.value)
  validationPanelRange.value = bounds ? rangeNotation(bounds) : selectedKey.value
  validationPanelOpen.value = true
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

async function sortSelection(ascending: boolean) {
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

async function executeRibbon(commandId: string, params?: unknown) {
  if (commandId !== 'sheet.command.undo' && commandId !== 'sheet.command.redo' && !finishEdit()) return
  const commands: Record<string, () => void | Promise<void>> = {
    'sheet.command.undo': undo, 'sheet.command.redo': redo,
    'sheet.command.copy': () => { copySelection() }, 'sheet.command.cut': cutSelection,
    'sheet.command.paste': () => { void pasteSelection() }, 'sheet.command.clear-content': clearSelection,
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
    'sheet.command.set-borders': () => { if (typeof params === 'string' && ['all', 'outside', 'inside', 'top', 'bottom', 'left', 'right', 'none'].includes(params)) setBorders(params as SheetBorderPreset) },
    'sheet.command.set-color': () => { if (typeof params === 'string') setCellFormat('color', params) },
    'sheet.command.set-background': () => { if (typeof params === 'string') setBackground(params) },
    'sheet.command.set-align': () => { if (params === 'left' || params === 'center' || params === 'right') setAlign(params) },
    'sheet.command.toggle-wrap': () => changeSelectedCells(cell => setSheetTextFit(cell, 'wrap', selectionFormat.value.wrap !== true)),
    'sheet.command.set-shrink-to-fit': () => {
      if (typeof params === 'boolean') changeSelectedCells(cell => setSheetTextFit(cell, 'shrinkToFit', params))
    },
    'sheet.command.set-text-rotation': () => {
      if (params === 'v') setTextRotation(params)
      else if (typeof params === 'number' && [-90, -45, 0, 45, 90].includes(params)) setTextRotation(params)
    },
    'sheet.command.merge': mergeSelection, 'sheet.command.unmerge': unmergeSelection,
    'sheet.command.insert-row': insertRow, 'sheet.command.insert-column': insertColumn,
    'sheet.command.delete-row': deleteRow, 'sheet.command.delete-column': deleteColumn,
    'sheet.operation.insert-chart': openChartDialog,
    'sheet.operation.open-pivot': openPivotPanel,
    'sheet.operation.open-slicer': () => openSlicerPanel('slicer'),
    'sheet.operation.open-timeline': () => openSlicerPanel('timeline'),
    'sheet.operation.open-data-panel': () => openDataPanel(params === 'sort' ? 'sort' : 'filter'),
    'sheet.operation.open-data-validation': openDataValidationPanel,
    'sheet.operation.open-function': () => openFunctionPanel(typeof params === 'string' ? params : undefined),
    'sheet.command.sort-ascending': () => sortSelection(true),
    'sheet.command.sort-descending': () => sortSelection(false),
    'sheet.command.recalculate': () => { recalculateFormulas(); commitHistory() },
    'sheet.operation.start-formula': () => { beginEdit(selectedKey.value, activeCell.value.formula ?? '=', 'formula') },
    'sheet.command.toggle-gridlines': () => { if (activeSheet.value) { activeSheet.value.showGridlines = activeSheet.value.showGridlines === false; commitHistory() } },
    'sheet.command.freeze-selection': () => { const [column, row] = splitKey(selectedKey.value); freeze(row - 1, column) },
    'sheet.command.freeze-first-row': () => freeze(1, 0), 'sheet.command.freeze-first-column': () => freeze(0, 1),
    'sheet.command.unfreeze': () => freeze(0, 0),
    'sheet.operation.zoom-in': () => { zoom.value = Math.min(400, zoom.value + 10) },
    'sheet.operation.zoom-out': () => { zoom.value = Math.max(10, zoom.value - 10) },
    'sheet.operation.reset-zoom': () => { zoom.value = 100 },
    'sheet.operation.find': openSearch,
  }
  const command = commands[commandId]
  if (!command) return
  try { await command() } catch (error) { emit('notify', { type: 'error', message: error instanceof Error ? error.message : '操作失败' }); return }
  if (!commandId.startsWith('sheet.operation.') || commandId.startsWith('sheet.operation.zoom')) void nextTick(() => spreadsheetCanvas.value?.focus())
}

function onWindowKeydown(event: KeyboardEvent) {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
    // The application shell owns global shortcuts when it supplies the header.
    if (!props.showHeader) return
    event.preventDefault()
    saveWorkbook()
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f' && document.activeElement?.closest('.excel-workbench')) {
    event.preventDefault()
    openSearch()
  }
}

watch(selectedKey, () => { nameError.value = ''; void nextTick(scrollToSelected) })
defineExpose({ save: saveWorkbook, focus: () => spreadsheetCanvas.value?.focus() })

onMounted(() => {
  window.addEventListener('mouseup', stopSelecting)
  window.addEventListener('keydown', onWindowKeydown)
})

onUnmounted(() => {
  sortController?.abort()
  window.removeEventListener('mouseup', stopSelecting)
  window.removeEventListener('keydown', onWindowKeydown)
})
onDeactivated(() => sortController?.abort())

</script>

<template>
  <section class="excel-workbench" @keydown="handleGridKeydown">
    <input ref="fileInput" class="hidden-file-input" type="file" accept=".xlsx,.xls,.csv" @change="handleFile" />
    <header v-if="showHeader" class="excel-header">
      <div class="document-brand">
        <div class="excel-logo"><FileSpreadsheet :size="18" /></div>
        <div class="document-title-wrap">
          <div class="document-title">{{ workbook.name }}</div>
          <div class="document-status">本地工作簿</div>
        </div>
      </div>
      <div class="header-actions">
        <Tooltip title="打开文件"><Button type="text" size="small" aria-label="打开文件" @click="openFileDialog"><FolderOpen :size="16" /></Button></Tooltip>
        <Tooltip title="保存"><Button type="text" size="small" aria-label="保存工作簿" @click="saveWorkbook"><Save :size="16" /></Button></Tooltip>
        <Tooltip title="导出"><Button type="text" size="small" aria-label="导出工作簿" @click="exportWorkbook"><Download :size="16" /></Button></Tooltip>
      </div>
    </header>

    <WorkbenchRibbon v-model="activeTab" :tabs="ribbonTabs" @execute="executeRibbon" />
    <SheetFormulaBar ref="formulaInput" :address="selectedAddress" :value="formulaValue" :editing="Boolean(editingKey)" :error="nameError" @update:value="editValue = $event" @begin="beginFormulaEdit" @confirm="finishFormulaEdit" @cancel="cancelFormulaEdit" @navigate="navigateToAddress" @focus-grid="nameError = ''; spreadsheetCanvas?.focus()" />
    <div v-if="sortBusy" class="sheet-find-panel" role="status"><span>{{ sortProgress }}</span><Button size="small" aria-label="取消排序" @click="sortController?.abort()">取消</Button></div>
    <div v-if="searchOpen" class="sheet-find-panel" role="search" @keydown.stop @keydown.esc.prevent="closeSearch">
      <Search :size="15" />
      <Input ref="searchInput" v-model:value="searchText" size="small" placeholder="查找工作表内容" aria-label="查找工作表内容" @press-enter="findNext" />
      <span class="search-count">{{ searchText ? `${searchMatches.length} 个匹配` : '输入查找内容' }}</span>
      <Button size="small" :disabled="!searchMatches.length" @click="findNext">下一个</Button>
      <Button type="text" size="small" aria-label="关闭查找" @click="closeSearch"><X :size="15" /></Button>
    </div>

    <div class="sheet-canvas-stack">
      <SpreadsheetCanvas v-if="activeSheet" ref="spreadsheetCanvas" v-model:edit-value="editValue" :sheet="activeSheet" :selected-key="selectedKey" :selected-range="emptyCellSelection" :selected-ranges="selectedRanges" :editing-key="editingKey" :zoom="zoom" :search-text="searchText" :validation-options="validationOptions" :overlay-commands="chartCommands" @select-cell="startCellSelection" @extend-selection="extendCellSelection" @select-header="selectHeader" @edit-cell="beginEdit" @finish-edit="finishEdit" @editor-keydown="handleGridKeydown" @paste="handlePaste" @resize-axis="resizeAxis" />
      <ChartOverlay v-if="activeSheet" :workbook="chartWorkbook" :active-sheet-id="activeSheet.id" :zoom="zoom" @overlay-commands="chartCommands = $event" @update-chart="updateChartLayout" @remove-chart="removeChart" />
    </div>

    <Modal v-model:open="chartModalOpen" title="插入图表" ok-text="插入" cancel-text="取消" :width="440" @ok="confirmChartDialog">
      <div class="chart-insert-form">
        <label>数据区域<Input v-model:value="chartSource" placeholder="A1:C4" /></label>
        <label>图表类型<Select v-model:value="chartType" :options="[{ label: '簇状柱形图', value: 'bar' }, { label: '折线图', value: 'line' }]" /></label>
        <label>图表标题<Input v-model:value="chartTitle" placeholder="可选" /></label>
        <p v-if="chartError" role="alert" class="chart-insert-error">{{ chartError }}</p>
        <p class="chart-insert-hint">数据区域首行作为系列名，首列作为分类标签。支持柱形图和折线图。</p>
      </div>
    </Modal>

    <Modal :open="Boolean(renameSheetId)" title="重命名工作表" ok-text="确定" cancel-text="取消" :width="360" @ok="confirmRenameSheet" @cancel="renameSheetId = null">
      <Input v-model:value="renameValue" aria-label="工作表名称" :status="renameError ? 'error' : undefined" :maxlength="31" @press-enter="confirmRenameSheet" />
      <p v-if="renameError" class="chart-insert-error" role="alert">{{ renameError }}</p>
    </Modal>
    <Modal v-model:open="functionPanelOpen" title="插入函数" ok-text="插入到公式栏" cancel-text="取消" :width="480" @ok="insertFunction">
      <div class="chart-insert-form">
        <label>函数<Select v-model:value="functionName" :options="functionOptions" aria-label="函数名称" show-search option-filter-prop="label" /></label>
        <p>{{ functionEntry?.category }} · {{ functionEntry?.label }}</p>
        <label>参数或区域<Input v-model:value="functionArguments" aria-label="函数参数" :placeholder="functionEntry?.example" @press-enter="insertFunction" /></label>
        <p class="chart-insert-hint">例如 {{ functionEntry?.example }}。插入后可在公式栏修改并确认。目录列出当前引擎支持的函数。</p>
        <p v-if="functionError" role="alert" class="chart-insert-error">{{ functionError }}</p>
      </div>
    </Modal>
    <SheetFooter :sheets="workbook.sheets" :active-sheet-id="activeSheetId" :zoom="zoom" :statistics="selectedStatistics" :number-format="activeCell.rawNumberFormat ?? activeCell.numberFormat" @activate="activateSheet" @add="addSheet" @rename="renameSheet" @delete="deleteSheet" @hide="hideSheet" @show="showSheet" @update:zoom="zoom = $event" />
    <PivotTablePanel v-if="pivotPanelOpen && activeSheet" v-model:open="pivotPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="pivotSource" />
    <SlicerPanel v-if="slicerPanelOpen && activeSheet && studio && workbook.id" v-model:open="slicerPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :initial-kind="slicerKind" />
    <SheetDataPanel v-if="dataPanelOpen && activeSheet && studio && workbook.id" v-model:open="dataPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="dataPanelRange" :initial-mode="dataPanelMode" />
    <SheetDataValidationPanel v-if="validationPanelOpen && activeSheet && studio && workbook.id" v-model:open="validationPanelOpen" :studio="studio" :workbook-id="workbook.id" :sheet-id="activeSheet.id" :source-range="validationPanelRange" />
  </section>
</template>

<style scoped>
.excel-workbench { --ink: #172033; --muted: #718096; --line: #dfe5ee; --blue: #246bfd; display: flex; flex: 1; min-width: 0; min-height: 0; flex-direction: column; overflow: hidden; color: var(--ink); background: #fff; font-family: Inter, "Segoe UI", "Microsoft YaHei", sans-serif; font-size: 13px; }
.hidden-file-input { display: none; }
.excel-header { display: flex; flex: 0 0 53px; align-items: center; justify-content: space-between; border-bottom: 1px solid var(--line); padding: 0 18px; background: #fff; }
.document-brand, .header-actions { display: flex; align-items: center; gap: 10px; }
.excel-logo { display: grid; width: 30px; height: 30px; place-items: center; color: #fff; border-radius: 6px; background: #1f9d68; }
.document-title-wrap { display: grid; gap: 1px; }
.document-title { color: #1c2535; font-size: 14px; font-weight: 600; line-height: 18px; }
.document-status { color: #8994a6; font-size: 11px; }
.header-actions { gap: 3px; }
.sheet-canvas-stack { position: relative; display: flex; flex: 1; min-width: 0; min-height: 0; overflow: hidden; }
.sheet-canvas-stack :deep(.spreadsheet-canvas) { flex: 1; }
.chart-insert-form { display: flex; flex-direction: column; gap: 14px; padding: 12px 0 4px; }
.chart-insert-form label { display: flex; flex-direction: column; gap: 6px; color: #4b5563; font-size: 12px; }
.chart-insert-error { margin: 8px 0 0; color: #c2413a; }
.chart-insert-hint { margin: 0; color: #8490a0; font-size: 11px; }
.sheet-find-panel { display: flex; flex: 0 0 40px; box-sizing: border-box; min-width: 0; align-items: center; gap: 8px; border-bottom: 1px solid var(--line); padding: 4px 12px; color: #718096; background: #f8faff; }
.sheet-find-panel :deep(.ant-input) { width: 240px; min-width: 40px; }
.search-count { flex-shrink: 0; font-size: 12px; }
@media (max-width: 600px) { .sheet-find-panel { gap: 4px; padding: 4px; } .sheet-find-panel :deep(.ant-input) { flex: 1; } }
</style>
