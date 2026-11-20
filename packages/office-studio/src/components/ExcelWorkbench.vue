<script setup lang="ts">
import { computed, nextTick, onMounted, onUnmounted, ref, watch } from 'vue'
import { Button, Input, Modal, Select, Tooltip } from 'antdv-next'
import { FUniver } from '../api/facade'
import type { OfficeStudio } from '../api/office-studio'
import type { ExcelCellView, ExcelWorkbookView as DomainExcelWorkbookView } from '../services/excelViewAdapter'
import type { ChartModel } from '../domain/model'
import SpreadsheetCanvas from './spreadsheet/SpreadsheetCanvas.vue'
import ChartOverlay from './ChartOverlay.vue'
import { chartPositionForRange, type ExcelWorkbookChartView } from './chart-overlay-model'
import type { DrawCommand } from '../rendering/contracts'
import type { GridHit, GridRange, SpreadsheetRenderSheet } from '../rendering/spreadsheet/layout'
import { addRange, containsCell, selectionBounds, selectionCount, subtractRange } from '../rendering/spreadsheet/selection'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  Check,
  ChevronDown,
  Copy,
  Scissors,
  Download,
  FileSpreadsheet,
  FolderOpen,
  Grid2X2,
  Italic,
  Merge,
  MoreHorizontal,
  PaintBucket,
  Plus,
  Printer,
  Redo2,
  Save,
  Search,
  Trash2,
  Underline,
  Undo2,
  ZoomIn,
  ZoomOut,
} from 'lucide-vue-next'

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
const activeSheetId = ref(workbook.value.sheets[0]?.id ?? 'sheet-1')
const selectedKey = ref('A2')
const selectionAnchor = ref('A2')
const selectedRanges = ref<GridRange[]>([{ startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }])
const emptyCellSelection: string[] = []
const editingKey = ref<string | null>(null)
const editValue = ref('')
const selecting = ref(false)
const formulaInput = ref<HTMLInputElement | null>(null)
const spreadsheetCanvas = ref<InstanceType<typeof SpreadsheetCanvas> | null>(null)
const fileInput = ref<HTMLInputElement | null>(null)
const zoom = ref(100)
const searchText = ref('')
const activeTab = ref<'home' | 'insert' | 'view'>('home')
const fontFamily = ref('等线')
const fontSize = ref(11)
const numberFormat = ref<'general' | 'number' | 'percent'>('general')
const copiedCells = ref<ExcelCell[][] | null>(null)
const chartCommands = ref<DrawCommand[]>([])
const chartModalOpen = ref(false)
const chartSource = ref('')
const chartTitle = ref('')
const chartType = ref<'bar' | 'line'>('bar')
const chartError = ref('')
const chartWorkbook = computed(() => workbook.value as ExcelWorkbookChartView)

const activeSheet = computed(() => workbook.value.sheets.find((sheet) => sheet.id === activeSheetId.value) ?? workbook.value.sheets[0])
const activeCell = computed(() => activeSheet.value?.cells[selectedKey.value] ?? { value: '' })
const selectedAddress = computed(() => selectedKey.value)
const selectedCount = computed(() => selectionCount(selectedRanges.value))
const sheetCells = computed(() => activeSheet.value?.cells ?? {})
const totalRows = computed(() => activeSheet.value?.rowCount ?? 100)
const fontChoices = ['等线', '微软雅黑', '宋体', 'Arial', 'Times New Roman']
const freezeRows = computed(() => activeSheet.value?.freeze?.rows ?? 0)
const freezeColumnsCount = computed(() => activeSheet.value?.freeze?.columns ?? 0)
const isFrozen = computed({
  get: () => freezeRows.value > 0 || freezeColumnsCount.value > 0,
  set: (value: boolean) => {
    if (!activeSheet.value) return
    activeSheet.value.freeze = value ? { rows: 2, columns: 1 } : undefined
    commitHistory()
  },
})
const canUndo = computed(() => props.canUndo ?? false)
const canRedo = computed(() => props.canRedo ?? false)
const selectedStatistics = computed(() => {
  let count = 0, sum = 0, numericCount = 0
  for (const [key, cell] of Object.entries(sheetCells.value)) {
    const [column, row] = splitKey(key)
    if (!containsCell(selectedRanges.value, row - 1, column)) continue
    const raw = cell.formula ? cell.cachedValue : cell.value
    if (raw !== undefined && raw !== '') { count += 1; const number = Number(raw); if (Number.isFinite(number)) { sum += number; numericCount += 1 } }
  }
  return { count, sum, numericCount }
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
  const sameDocument = value.sheets.some((sheet) => sheet.id === previousSheetId)
  workbook.value = cloneWorkbook(value)
  activeSheetId.value = sameDocument ? previousSheetId : workbook.value.sheets[0]?.id ?? ''
  const sheet = activeSheet.value
  const isValidAddress = (key: string) => {
    const [column, row] = splitKey(key)
    return Boolean(sheet && row >= 1 && row <= sheet.rowCount && column >= 0 && column < sheet.columnCount)
  }
  if (!sameDocument || !isValidAddress(selectedKey.value)) {
    selectedKey.value = 'A1'
    selectedRanges.value = [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }]
    selectionAnchor.value = 'A1'
    editingKey.value = null
  } else {
    selectedRanges.value = selectedRanges.value.flatMap(range => sheet && range.startRow < sheet.rowCount && range.startColumn < sheet.columnCount ? [{ ...range, endRow: Math.min(sheet.rowCount - 1, range.endRow), endColumn: Math.min(sheet.columnCount - 1, range.endColumn) }] : [])
    if (!selectedRanges.value.length) selectCell(selectedKey.value)
    if (!isValidAddress(selectionAnchor.value)) selectionAnchor.value = selectedKey.value
  }
}

function commitHistory() {
  const snapshot = cloneWorkbook(workbook.value)
  emit('update:modelValue', snapshot)
  emit('update:model', snapshot)
  emit('change', snapshot)
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

function setCellValue(key: string, value: string) {
  if (!activeSheet.value) return
  const current = activeSheet.value.cells[key] ?? { value: '' }
  const formula = value.startsWith('=') ? value : undefined
  activeSheet.value.cells[key] = { ...current, value: formula ? '' : value, formula, cachedValue: formula ? '' : undefined }
  recalculateFormulas()
  emit('edit-cell', { sheetId: activeSheet.value.id, key, cell: activeSheet.value.cells[key] })
  commitHistory()
}

function evaluateFormula(formula: string, cells: Record<string, ExcelCell>): string {
  const expression = formula.slice(1).trim()
  const sumMatch = expression.match(/^SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i)
  if (sumMatch) {
    const [startColumn, startRow] = splitKey(sumMatch[1])
    const [endColumn, endRow] = splitKey(sumMatch[2])
    const values: number[] = []
    for (let row = startRow; row <= endRow; row += 1) {
      for (let column = startColumn; column <= endColumn; column += 1) {
        const cell = cells[`${columnName(column)}${row}`]
        const rawValue = cell?.formula ? cell.cachedValue : cell?.value
        const number = rawValue === undefined || rawValue === '' ? 0 : Number(rawValue)
        if (!Number.isNaN(number)) values.push(number)
      }
    }
    return String(values.reduce((sum, value) => sum + value, 0))
  }
  const tokens: Array<{ kind: 'number' | 'cell' | 'operator'; text: string }> = []
  const tokenPattern = /\s*(?:(\d+(?:\.\d*)?|\.\d+)|(\$?[A-Z]{1,3}\$?[1-9]\d*)|([()+\-*/]))/iy
  let cursor = 0
  while (cursor < expression.length) {
    tokenPattern.lastIndex = cursor
    const match = tokenPattern.exec(expression)
    if (!match) return '#UNSUPPORTED!'
    const text = match[1] ?? match[2] ?? match[3]
    tokens.push({ kind: match[1] ? 'number' : match[2] ? 'cell' : 'operator', text })
    cursor = tokenPattern.lastIndex
  }
  let tokenIndex = 0
  const readCell = (address: string) => {
    const cell = cells[address.replace(/\$/g, '').toUpperCase()]
    const value = cell?.formula ? cell.cachedValue : cell?.value
    if (value === undefined || value === '') return 0
    const number = Number(value)
    return Number.isFinite(number) ? number : Number.NaN
  }
  const parsePrimary = (): number => {
    const token = tokens[tokenIndex++]
    if (!token) throw new Error('expected-value')
    if (token.kind === 'number') return Number(token.text)
    if (token.kind === 'cell') return readCell(token.text)
    if (token.text === '(') {
      const value = parseAddition()
      if (tokens[tokenIndex++]?.text !== ')') throw new Error('expected-close')
      return value
    }
    throw new Error('expected-primary')
  }
  const parseUnary = (): number => {
    const operator = tokens[tokenIndex]?.text
    if (operator === '+' || operator === '-') {
      tokenIndex += 1
      const value = parseUnary()
      return operator === '-' ? -value : value
    }
    return parsePrimary()
  }
  const parseMultiplication = (): number => {
    let value = parseUnary()
    while (tokens[tokenIndex]?.text === '*' || tokens[tokenIndex]?.text === '/') {
      const operator = tokens[tokenIndex++].text
      const next = parseUnary()
      if (operator === '/' && next === 0) throw new Error('divide-by-zero')
      value = operator === '*' ? value * next : value / next
    }
    return value
  }
  const parseAddition = (): number => {
    let value = parseMultiplication()
    while (tokens[tokenIndex]?.text === '+' || tokens[tokenIndex]?.text === '-') {
      const operator = tokens[tokenIndex++].text
      const next = parseMultiplication()
      value = operator === '+' ? value + next : value - next
    }
    return value
  }
  try {
    const value = parseAddition()
    if (tokenIndex !== tokens.length || Number.isNaN(value)) return '#VALUE!'
    if (!Number.isFinite(value)) return '#NUM!'
    return String(value)
  } catch (error) {
    return error instanceof Error && error.message === 'divide-by-zero' ? '#DIV/0!' : '#VALUE!'
  }
}

function recalculateFormulas() {
  const cells = activeSheet.value?.cells
  if (!cells) return
  for (let pass = 0; pass < 5; pass += 1) {
    Object.entries(cells).forEach(([key, cell]) => {
      if (cell.formula) cells[key] = { ...cell, cachedValue: evaluateFormula(cell.formula, cells) }
    })
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

function selectCell(key: string, extend = false) {
  const [column, row] = splitKey(key)
  if (extend) {
    const [anchorColumn, anchorRow] = splitKey(selectionAnchor.value)
    selectedRanges.value = [{ startRow: Math.min(anchorRow, row) - 1, endRow: Math.max(anchorRow, row) - 1, startColumn: Math.min(anchorColumn, column), endColumn: Math.max(anchorColumn, column) }]
  } else {
    selectionAnchor.value = key
    selectedRanges.value = [{ startRow: row - 1, endRow: row - 1, startColumn: column, endColumn: column }]
  }
  selectedKey.value = key
}

function beginEdit(key = selectedKey.value, initialValue?: string, target: 'cell' | 'formula' = 'cell') {
  editingKey.value = key
  editValue.value = initialValue ?? getCell(key).formula ?? getCell(key).value
  nextTick(() => target === 'cell' ? spreadsheetCanvas.value?.focusEditor() : formulaInput.value?.focus())
}

function beginFormulaEdit() {
  if (!editingKey.value) beginEdit(selectedKey.value, undefined, 'formula')
}

function finishFormulaEdit(event: KeyboardEvent) {
  event.preventDefault()
  event.stopPropagation()
  finishEdit()
  moveSelection(1, 0)
}

function finishEdit() {
  if (!editingKey.value) return
  setCellValue(editingKey.value, editValue.value)
  editingKey.value = null
}

function cancelEdit() {
  editingKey.value = null
}

function moveSelection(rowOffset: number, columnOffset: number, extend = false) {
  const [column, row] = splitKey(selectedKey.value)
  const nextColumn = Math.max(0, Math.min((activeSheet.value?.columnCount ?? columnCount) - 1, column + columnOffset))
  const nextRow = Math.max(1, Math.min(activeSheet.value?.rowCount ?? 100, row + rowOffset))
  selectCell(`${columnName(nextColumn)}${nextRow}`, extend)
}

function handleGridKeydown(event: KeyboardEvent) {
  if (event.isComposing) return
  const target = event.target as HTMLElement | null
  if (target?.matches('input, textarea, select') && !target.classList.contains('canvas-cell-input') && target !== formulaInput.value) return
  if (event.ctrlKey && event.key === 'PageDown') {
    event.preventDefault()
    switchNextSheet()
    return
  }
  if (editingKey.value) {
    if (event.key === 'Enter' || event.key === 'Tab') {
      event.preventDefault()
      finishEdit()
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
  selectedExistingKeys().forEach((key) => { const cell = activeSheet.value?.cells[key]; if (cell) activeSheet.value!.cells[key] = { ...cell, value: '', formula: undefined, cachedValue: undefined, rawValue: '', displayValue: '' } })
  recalculateFormulas()
  commitHistory()
}

function applyPasteText(text: string) {
  if (!text || !activeSheet.value) return
  const [startColumn, startRow] = splitKey(selectedKey.value)
  text.replace(/\r/g, '').split('\n').forEach((rowText, rowOffset) => {
    rowText.split('\t').forEach((value, columnOffset) => {
      const column = startColumn + columnOffset
      const row = startRow + rowOffset
      if (column >= (activeSheet.value?.columnCount ?? columnCount) || row > totalRows.value) return
      const key = `${columnName(column)}${row}`
      const current = activeSheet.value?.cells[key] ?? { value: '' }
      const formula = value.startsWith('=') ? value : undefined
      activeSheet.value!.cells[key] = { ...current, value: formula ? '' : value, formula, cachedValue: formula ? '' : undefined }
    })
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

function toggleStyle(style: 'bold' | 'italic' | 'underline') {
  if (!activeSheet.value) return
  const value = !activeCell.value[style]
  if (!applyToSelection((key) => {
    const cell = activeSheet.value?.cells[key] ?? { value: '' }
    activeSheet.value!.cells[key] = { ...cell, [style]: value }
  })) return
  commitHistory()
}

function applyFont() {
  if (!activeSheet.value) return
  if (!applyToSelection((key) => {
    const cell = activeSheet.value?.cells[key] ?? { value: '' }
    activeSheet.value!.cells[key] = { ...cell, fontSize: fontSize.value, fontFamily: fontFamily.value }
  })) return
  commitHistory()
}

function applyNumberFormat() {
  if (!activeSheet.value) return
  if (!applyToSelection((key) => {
    const cell = activeSheet.value?.cells[key] ?? { value: '' }
    activeSheet.value!.cells[key] = { ...cell, numberFormat: numberFormat.value }
  })) return
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

function addSheet() {
  const id = `sheet-${Date.now()}`
  workbook.value.sheets.push({ id, name: `工作表${workbook.value.sheets.length + 1}`, cells: {}, rowCount: 100, columnCount })
  activeSheetId.value = id
  commitHistory()
}

function switchNextSheet() {
  const index = workbook.value.sheets.findIndex((sheet) => sheet.id === activeSheetId.value)
  const next = workbook.value.sheets[(index + 1) % workbook.value.sheets.length]
  if (next) activeSheetId.value = next.id
}

function renameSheet(sheet: ExcelSheet) {
  const name = window.prompt('重命名工作表', sheet.name)
  if (name?.trim()) {
    sheet.name = name.trim()
    commitHistory()
  }
}

function deleteSheet(sheet: ExcelSheet) {
  if (workbook.value.sheets.length <= 1) return
  const index = workbook.value.sheets.findIndex((item) => item.id === sheet.id)
  workbook.value.sheets = workbook.value.sheets.filter((item) => item.id !== sheet.id)
  if (activeSheetId.value === sheet.id) activeSheetId.value = workbook.value.sheets[Math.max(0, index - 1)].id
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
  emit('save', cloneWorkbook(workbook.value))
}

function exportWorkbook() {
  emit('export', cloneWorkbook(workbook.value))
}

function getDisplayValue(key: string) {
  const cell = getCell(key)
  const value = cell.formula ? cell.cachedValue ?? cell.formula : cell.value
  if (cell.numberFormat === 'percent' && value !== '' && !Number.isNaN(Number(value))) return `${(Number(value) * 100).toLocaleString('zh-CN')}%`
  if (cell.numberFormat === 'number' && value !== '' && !Number.isNaN(Number(value))) return Number(value).toLocaleString('zh-CN')
  return value
}

function scrollToSelected() {
  spreadsheetCanvas.value?.ensureVisible(selectedKey.value)
}

function startCellSelection(key: string, event: MouseEvent) {
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

function onWindowKeydown(event: KeyboardEvent) {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
    event.preventDefault()
    saveWorkbook()
  }
}

function onPrint() {
  window.print()
}

function updateFormulaInput(event: Event) {
  editValue.value = (event.target as HTMLInputElement).value
}

function updateFontFamily(event: Event) {
  fontFamily.value = (event.target as HTMLSelectElement).value
  applyFont()
}

function updateFontSize(event: Event) {
  fontSize.value = Number((event.target as HTMLSelectElement).value)
  applyFont()
}

function updateNumberFormat(event: Event) {
  numberFormat.value = (event.target as HTMLSelectElement).value as 'general' | 'number' | 'percent'
  applyNumberFormat()
}

watch(selectedKey, () => nextTick(scrollToSelected))
watch([selectedKey, activeSheetId], () => {
  fontFamily.value = activeCell.value.fontFamily ?? '等线'
  fontSize.value = activeCell.value.fontSize ?? 11
  numberFormat.value = activeCell.value.numberFormat ?? 'general'
})

onMounted(() => {
  window.addEventListener('mouseup', stopSelecting)
  window.addEventListener('keydown', onWindowKeydown)
})

onUnmounted(() => {
  window.removeEventListener('mouseup', stopSelecting)
  window.removeEventListener('keydown', onWindowKeydown)
})

</script>

<template>
  <section class="excel-workbench" @keydown="handleGridKeydown">
    <input ref="fileInput" class="hidden-file-input" type="file" accept=".xlsx,.xls,.csv" @change="handleFile" />
    <header v-if="showHeader" class="excel-header">
      <div class="document-brand">
        <div class="excel-logo"><FileSpreadsheet :size="18" /></div>
        <div class="document-title-wrap">
          <div class="document-title">{{ workbook.name }}</div>
          <div class="document-status"><span class="status-dot" />已保存到本地</div>
        </div>
      </div>
      <div class="header-actions">
        <Tooltip title="打开文件"><Button type="text" size="small" @click="openFileDialog"><FolderOpen :size="16" /></Button></Tooltip>
        <Tooltip title="保存"><Button type="text" size="small" @click="saveWorkbook"><Save :size="16" /></Button></Tooltip>
        <Tooltip title="导出"><Button type="text" size="small" @click="exportWorkbook"><Download :size="16" /></Button></Tooltip>
        <Tooltip title="打印"><Button type="text" size="small" @click="onPrint"><Printer :size="16" /></Button></Tooltip>
        <span class="header-divider" />
        <div class="avatar">州</div>
      </div>
    </header>

    <nav class="workbook-tabs" aria-label="Excel tools">
      <button :class="['ribbon-tab', { active: activeTab === 'home' }]" @click="activeTab = 'home'">开始</button>
      <button :class="['ribbon-tab', { active: activeTab === 'insert' }]" @click="activeTab = 'insert'">插入</button>
      <button :class="['ribbon-tab', { active: activeTab === 'view' }]" @click="activeTab = 'view'">视图</button>
      <div class="ribbon-spacer" />
      <div class="search-box"><Search :size="15" /><input v-model="searchText" placeholder="搜索工作簿" aria-label="搜索工作簿" /></div>
    </nav>

    <div class="ribbon-toolbar">
      <div class="toolbar-group history-group">
        <Tooltip title="撤销"><Button type="text" size="small" :disabled="!canUndo" @click="undo"><Undo2 :size="17" /></Button></Tooltip>
        <Tooltip title="重做"><Button type="text" size="small" :disabled="!canRedo" @click="redo"><Redo2 :size="17" /></Button></Tooltip>
      </div>
      <div class="toolbar-group">
        <label class="select-wrap font-select" title="字体"><select :value="fontFamily" aria-label="字体" @change="updateFontFamily"><option v-for="font in fontChoices" :key="font" :value="font">{{ font }}</option></select><ChevronDown :size="13" /></label>
        <label class="select-wrap size-select" title="字号"><select :value="fontSize" aria-label="字号" @change="updateFontSize"><option v-for="size in [10, 11, 12, 14, 16, 18, 20]" :key="size" :value="size">{{ size }}</option></select><ChevronDown :size="13" /></label>
        <Tooltip title="加粗"><Button type="text" size="small" :class="{ active: activeCell.bold }" @click="toggleStyle('bold')"><Bold :size="16" /></Button></Tooltip>
        <Tooltip title="斜体"><Button type="text" size="small" :class="{ active: activeCell.italic }" @click="toggleStyle('italic')"><Italic :size="16" /></Button></Tooltip>
        <Tooltip title="下划线"><Button type="text" size="small" :class="{ active: activeCell.underline }" @click="toggleStyle('underline')"><Underline :size="16" /></Button></Tooltip>
        <span class="toolbar-divider" />
        <Tooltip title="左对齐"><Button type="text" size="small" @click="setAlign('left')"><AlignLeft :size="16" /></Button></Tooltip>
        <Tooltip title="居中"><Button type="text" size="small" @click="setAlign('center')"><AlignCenter :size="16" /></Button></Tooltip>
        <Tooltip title="右对齐"><Button type="text" size="small" @click="setAlign('right')"><AlignRight :size="16" /></Button></Tooltip>
        <Tooltip title="填充颜色"><Button type="text" size="small" @click="setBackground('#fff4cc')"><PaintBucket :size="16" /></Button></Tooltip>
      </div>
      <div class="toolbar-group">
        <Tooltip title="合并单元格"><Button type="text" size="small" @click="mergeSelection"><Merge :size="16" /></Button></Tooltip>
        <Tooltip title="取消合并"><Button type="text" size="small" @click="unmergeSelection"><Merge :size="16" /></Button></Tooltip>
        <Tooltip title="剪切"><Button type="text" size="small" @click="cutSelection"><Scissors :size="16" /></Button></Tooltip>
        <Tooltip title="复制"><Button type="text" size="small" @click="copySelection"><Copy :size="16" /></Button></Tooltip>
        <Tooltip title="清除内容"><Button type="text" size="small" @click="clearSelection"><Trash2 :size="16" /></Button></Tooltip>
        <label class="select-wrap number-format-select" title="数字格式"><select :value="numberFormat" aria-label="数字格式" @change="updateNumberFormat"><option value="general">常规</option><option value="number">数字</option><option value="percent">百分比</option></select><ChevronDown :size="13" /></label>
      </div>
      <div v-if="activeTab === 'insert'" class="toolbar-group">
        <Tooltip title="插入行"><Button type="text" size="small" @click="insertRow"><Plus :size="16" /> 行</Button></Tooltip>
        <Tooltip title="删除行"><Button type="text" size="small" @click="deleteRow"><Trash2 :size="16" /> 行</Button></Tooltip>
        <Tooltip title="插入列"><Button type="text" size="small" @click="insertColumn"><Plus :size="16" /> 列</Button></Tooltip>
        <Tooltip title="删除列"><Button type="text" size="small" @click="deleteColumn"><Trash2 :size="16" /> 列</Button></Tooltip>
        <Tooltip title="从选区插入图表"><Button type="primary" size="small" :disabled="!studio || !workbook.id" @click="openChartDialog"><FileSpreadsheet :size="15" /> 图表</Button></Tooltip>
      </div>
      <div class="toolbar-group ribbon-summary">
        <span v-if="activeTab === 'insert'">从带表头的数据区域创建图表，拖动图表可调整位置</span>
        <span v-else-if="activeTab === 'view'">{{ isFrozen ? '已冻结前两行和首列' : '未冻结窗格' }}</span>
        <span v-else>编辑模式</span>
      </div>
    </div>

    <div class="formula-toolbar">
      <div class="name-box">{{ selectedAddress }} <ChevronDown :size="13" /></div>
      <div class="formula-symbol">fx</div>
      <input ref="formulaInput" v-model="editValue" class="formula-input" :placeholder="getDisplayValue(selectedKey) || '输入数据或公式'" @focus="beginFormulaEdit" @input="updateFormulaInput" @blur="finishEdit" @keydown.enter="finishFormulaEdit" />
      <Button type="text" size="small" class="formula-confirm" :class="{ visible: editingKey }" @click="finishEdit"><Check :size="16" /></Button>
    </div>

    <div class="sheet-canvas-stack">
      <SpreadsheetCanvas v-if="activeSheet" ref="spreadsheetCanvas" v-model:edit-value="editValue" :sheet="activeSheet" :selected-key="selectedKey" :selected-range="emptyCellSelection" :selected-ranges="selectedRanges" :editing-key="editingKey" :zoom="zoom" :search-text="searchText" :overlay-commands="chartCommands" @select-cell="startCellSelection" @extend-selection="extendCellSelection" @select-header="selectHeader" @edit-cell="beginEdit" @finish-edit="finishEdit" @editor-keydown="handleGridKeydown" @paste="handlePaste" @resize-axis="resizeAxis" />
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

    <div class="sheet-bar">
      <div class="sheet-controls">
        <Tooltip title="添加工作表"><Button type="text" size="small" @click="addSheet"><Plus :size="17" /></Button></Tooltip>
        <Tooltip title="所有工作表"><Button type="text" size="small"><MoreHorizontal :size="17" /></Button></Tooltip>
      </div>
      <div class="sheet-tabs">
        <button
          v-for="sheet in workbook.sheets"
          :key="sheet.id"
          :class="['sheet-tab', { active: activeSheetId === sheet.id }]"
          @click="activeSheetId = sheet.id"
          @dblclick="renameSheet(sheet)"
        >
          {{ sheet.name }}
          <span v-if="activeSheetId === sheet.id" class="sheet-tab-close" @click.stop="deleteSheet(sheet)">×</span>
        </button>
      </div>
      <div class="sheet-spacer" />
      <div class="selection-info" :title="`已选择 ${selectedCount.toLocaleString()} 个单元格`">
        <span v-if="selectedStatistics.numericCount > 0 && selectedStatistics.count > 1">求和：{{ selectedStatistics.sum.toLocaleString() }}</span>
        <span v-if="selectedStatistics.count > 1">计数：{{ selectedStatistics.count.toLocaleString() }}</span>
      </div>
      <div class="sheet-actions">
        <Tooltip title="冻结窗格"><Button type="text" size="small" :class="{ active: isFrozen }" @click="isFrozen = !isFrozen"><Grid2X2 :size="16" /></Button></Tooltip>
        <Tooltip title="缩小"><Button type="text" size="small" @click="zoom = Math.max(60, zoom - 10)"><ZoomOut :size="16" /></Button></Tooltip>
        <span class="zoom-value">{{ zoom }}%</span>
        <Tooltip title="放大"><Button type="text" size="small" @click="zoom = Math.min(150, zoom + 10)"><ZoomIn :size="16" /></Button></Tooltip>
      </div>
    </div>
  </section>
</template>

<style scoped>
.excel-workbench {
  --ink: #172033;
  --muted: #718096;
  --line: #dfe5ee;
  --soft: #f6f8fb;
  --blue: #246bfd;
  display: flex;
  flex: 1;
  min-width: 0;
  min-height: 0;
  flex-direction: column;
  color: var(--ink);
  background: #fff;
  font-family: Inter, "Segoe UI", "Microsoft YaHei", sans-serif;
  font-size: 13px;
  overflow: hidden;
}
.hidden-file-input { display: none; }
.excel-header { display: flex; flex: 0 0 53px; align-items: center; justify-content: space-between; border-bottom: 1px solid var(--line); padding: 0 18px; background: #fff; }
.document-brand, .header-actions, .toolbar-group, .sheet-actions, .status-bar > span, .sheet-controls { display: flex; align-items: center; }
.document-brand { gap: 10px; }
.excel-logo { display: grid; width: 30px; height: 30px; place-items: center; color: #fff; border-radius: 6px; background: #1f9d68; }
.document-title-wrap { display: grid; gap: 1px; }
.document-title { color: #1c2535; font-size: 14px; font-weight: 600; line-height: 18px; }
.document-status { display: flex; align-items: center; gap: 5px; color: #8994a6; font-size: 11px; }
.status-dot { width: 6px; height: 6px; border-radius: 50%; background: #28a970; }
.header-actions { gap: 3px; }
.header-actions :deep(.ant-btn), .toolbar-group :deep(.ant-btn), .sheet-actions :deep(.ant-btn), .sheet-controls :deep(.ant-btn), .formula-confirm { color: #5c6a7f; }
.header-actions :deep(.ant-btn:hover), .toolbar-group :deep(.ant-btn:hover), .sheet-actions :deep(.ant-btn:hover), .sheet-controls :deep(.ant-btn:hover) { color: var(--blue); background: #edf3ff; }
.header-divider, .toolbar-divider { width: 1px; height: 20px; margin: 0 8px; background: var(--line); }
.avatar { display: grid; width: 28px; height: 28px; margin-left: 10px; place-items: center; color: #fff; border-radius: 50%; background: #5568ce; font-size: 12px; }
.workbook-tabs { display: flex; flex: 0 0 38px; align-items: stretch; gap: 2px; border-bottom: 1px solid var(--line); padding: 0 16px; background: #fff; }
.ribbon-tab { position: relative; border: 0; padding: 0 16px; color: #6f7c8f; background: transparent; cursor: pointer; font-size: 13px; }
.ribbon-tab.active { color: var(--blue); font-weight: 600; }
.ribbon-tab.active::after { position: absolute; right: 13px; bottom: -1px; left: 13px; height: 2px; background: var(--blue); content: ""; }
.ribbon-spacer { flex: 1; }
.search-box { display: flex; width: 175px; align-items: center; gap: 6px; margin: 6px 0; border: 1px solid var(--line); border-radius: 4px; padding: 0 8px; color: #9ba5b4; background: #fafbfd; }
.search-box input { min-width: 0; flex: 1; border: 0; outline: 0; color: var(--ink); background: transparent; font-size: 12px; }
.ribbon-toolbar { display: flex; flex: 0 0 44px; align-items: center; gap: 14px; border-bottom: 1px solid var(--line); padding: 0 17px; background: #fbfcfe; }
.toolbar-group { gap: 1px; }
.toolbar-group :deep(.ant-btn.active), .sheet-actions :deep(.ant-btn.active) { color: var(--blue); background: #eaf1ff; }
.font-select, .size-select { display: flex; height: 26px; align-items: center; justify-content: space-between; gap: 12px; border: 1px solid var(--line); border-radius: 3px; padding: 0 7px; color: #445168; background: #fff; cursor: pointer; font-size: 12px; }
.font-select { width: 72px; }
.size-select { width: 42px; }
.select-wrap select { min-width: 0; flex: 1; border: 0; outline: 0; color: inherit; background: transparent; cursor: pointer; font: inherit; }
.number-format-select { width: 74px; gap: 4px; }
.ribbon-summary { margin-left: auto; color: #8692a4; font-size: 11px; }
.formula-toolbar { display: flex; flex: 0 0 38px; align-items: center; gap: 7px; border-bottom: 1px solid var(--line); padding: 0 17px; background: #fff; }
.sheet-canvas-stack { position: relative; display: flex; min-width: 0; min-height: 0; flex: 1; overflow: hidden; }
.sheet-canvas-stack :deep(.spreadsheet-canvas) { flex: 1; }
.chart-insert-form { display: flex; flex-direction: column; gap: 14px; padding: 12px 0 4px; }
.chart-insert-form label { display: flex; flex-direction: column; gap: 6px; color: #4b5563; font-size: 12px; }
.chart-insert-error { margin: 0; color: #c2413a; }
.chart-insert-hint { margin: 0; color: #8490a0; font-size: 11px; }
.name-box { display: flex; width: 65px; height: 25px; align-items: center; justify-content: space-between; border: 1px solid var(--line); border-radius: 3px; padding: 0 6px 0 8px; color: #37445a; background: #fbfcfe; font-size: 12px; }
.formula-symbol { width: 20px; color: #8792a4; font-family: Georgia, serif; font-size: 16px; font-style: italic; text-align: center; }
.formula-input { height: 25px; min-width: 0; flex: 1; border: 1px solid var(--line); border-radius: 3px; outline: 0; padding: 0 8px; color: #364257; background: #fff; font-family: Consolas, monospace; font-size: 12px; }
.formula-input:focus { border-color: #8aaefb; box-shadow: 0 0 0 2px rgb(36 107 253 / 10%); }
.formula-confirm { opacity: 0; transition: opacity .15s ease; }
.formula-confirm.visible { opacity: 1; }
.grid-viewport { position: relative; min-height: 0; flex: 1; overflow: auto; background: #fff; }
.grid { display: grid; width: max-content; transform-origin: top left; transform: scale(var(--zoom)); }
.row-spacer { grid-column: 1 / -1; pointer-events: none; }
.corner-cell, .column-header, .row-header, .grid-cell { box-sizing: border-box; height: 26px; border-right: 1px solid var(--line); border-bottom: 1px solid var(--line); }
.corner-cell { position: sticky; z-index: 4; left: 0; display: grid; width: 44px; place-items: center; color: #95a0ae; background: #f3f5f8; }
.column-header { position: sticky; z-index: 3; top: 0; display: grid; place-items: center; color: #617086; background: #f3f5f8; font-size: 11px; }
.column-header.selected, .row-header.selected { color: var(--blue); background: #e8f0ff; font-weight: 600; }
.row-header { position: sticky; z-index: 2; left: 0; display: grid; width: 44px; place-items: center; color: #7b8799; background: #f3f5f8; font-size: 11px; }
.grid-cell { position: relative; border-top: 0; border-left: 0; padding: 0 8px; overflow: hidden; color: #29364a; background: #fff; cursor: cell; font: inherit; line-height: 25px; outline: 0; text-overflow: ellipsis; white-space: nowrap; }
.grid-cell:hover { background: #f7faff; }
.grid-cell.selected { background-color: #f1f6ff; }
.grid-cell.current::after { position: absolute; z-index: 1; inset: -1px; border: 2px solid var(--blue); pointer-events: none; content: ""; }
.grid-cell.match { background-color: #fff2bc; }
.cell-editor { position: absolute; z-index: 3; inset: 0; width: 100%; box-sizing: border-box; border: 2px solid var(--blue); outline: 0; padding: 0 7px; color: #29364a; background: #fff; font: inherit; }
.freeze-line { position: absolute; z-index: 5; pointer-events: none; box-shadow: 2px 0 5px rgb(35 62 103 / 12%); }
.freeze-horizontal { right: 0; left: 0; height: 1px; border-top: 1px solid #9cb8ec; box-shadow: 0 2px 5px rgb(35 62 103 / 14%); }
.freeze-vertical { top: 0; bottom: 0; width: 1px; border-left: 1px solid #9cb8ec; }
.sheet-bar { display: flex; flex: 0 0 36px; align-items: center; gap: 6px; border-top: 1px solid var(--line); background: #f8fafc; }
.sheet-controls { gap: 1px; padding-left: 8px; }
.sheet-tabs { display: flex; height: 100%; align-items: stretch; gap: 2px; overflow-x: auto; }
.sheet-tab { display: flex; align-items: center; gap: 7px; border: 0; border-right: 1px solid transparent; border-left: 1px solid transparent; padding: 0 16px; color: #6c788a; background: transparent; cursor: pointer; font-size: 12px; white-space: nowrap; }
.sheet-tab.active { border-right-color: var(--line); border-left-color: var(--line); color: #27364c; background: #fff; box-shadow: inset 0 2px 0 var(--blue); font-weight: 600; }
.sheet-tab-close { color: #a0aaba; font-size: 15px; font-weight: 400; }
.sheet-spacer { flex: 1; }
.sheet-actions { gap: 2px; padding-right: 12px; }
.zoom-value { min-width: 38px; color: #6f7c8f; font-size: 11px; text-align: center; }
.status-bar { display: flex; flex: 0 0 25px; align-items: center; gap: 16px; border-top: 1px solid #d7e1ef; padding: 0 13px; color: #8090a5; background: #f2f7ff; font-size: 11px; }
.status-bar > span { gap: 4px; }
.status-spacer { flex: 1; }
@media (max-width: 800px) {
  .ribbon-summary, .search-box { display: none; }
  .ribbon-toolbar { gap: 6px; overflow-x: auto; }
  .header-actions .ant-btn:nth-child(3), .header-actions .ant-btn:nth-child(4) { display: none; }
  .status-bar { gap: 8px; }
  .status-bar > span:last-child { display: none; }
}
</style>
