import type { ExcelCellView, ExcelSheetView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { addRange, containsCell } from '../../rendering/spreadsheet/selection.ts'
import { findVisibleSheetRow, hiddenSheetRows } from './sheet-row-visibility.ts'

export interface SheetAddressBounds { rowCount: number; columnCount: number; hiddenRows?: readonly number[]; filteredRows?: readonly number[]; rowHeights?: Readonly<Record<number, number>> }
export interface ParsedSheetAddress { range: GridRange; address: string; anchor: string }

function columnIndex(label: string): number {
  let index = 0
  for (const character of label.toUpperCase()) {
    index = index * 26 + character.charCodeAt(0) - 64
    if (!Number.isSafeInteger(index)) return -1
  }
  return index - 1
}

function columnLabel(index: number): string {
  let value = index + 1, label = ''
  while (value > 0) {
    value -= 1
    label = String.fromCharCode(value % 26 + 65) + label
    value = Math.floor(value / 26)
  }
  return label
}

function hasValidBounds(bounds: SheetAddressBounds): boolean {
  return Number.isSafeInteger(bounds.rowCount) && bounds.rowCount > 0
    && Number.isSafeInteger(bounds.columnCount) && bounds.columnCount > 0
}

/** Parse a current-sheet reference. Named/cross-sheet references are resolved by the command layer. */
export function parseSheetAddress(input: string, bounds: SheetAddressBounds): ParsedSheetAddress | undefined {
  if (!hasValidBounds(bounds)) return undefined
  const text = input.trim().toUpperCase()
  let firstRow: number, lastRow: number, firstColumn: number, lastColumn: number
  let kind: 'cell' | 'row' | 'column' = 'cell'
  const cells = /^\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/.exec(text)
  const columns = /^\$?([A-Z]+):\$?([A-Z]+)$/.exec(text)
  const rows = /^\$?([1-9]\d*):\$?([1-9]\d*)$/.exec(text)
  if (cells) {
    firstRow = Number(cells[2]) - 1
    lastRow = Number(cells[4] ?? cells[2]) - 1
    firstColumn = columnIndex(cells[1]!)
    lastColumn = columnIndex(cells[3] ?? cells[1]!)
  } else if (columns) {
    kind = 'column'
    firstRow = 0; lastRow = bounds.rowCount - 1
    firstColumn = columnIndex(columns[1]!)
    lastColumn = columnIndex(columns[2]!)
  } else if (rows) {
    kind = 'row'
    firstRow = Number(rows[1]) - 1
    lastRow = Number(rows[2]) - 1
    firstColumn = 0; lastColumn = bounds.columnCount - 1
  } else return undefined
  if (![firstRow, lastRow, firstColumn, lastColumn].every(value => Number.isSafeInteger(value) && value >= 0)
    || Math.max(firstRow, lastRow) >= bounds.rowCount || Math.max(firstColumn, lastColumn) >= bounds.columnCount) return undefined
  const range = {
    startRow: Math.min(firstRow, lastRow), endRow: Math.max(firstRow, lastRow),
    startColumn: Math.min(firstColumn, lastColumn), endColumn: Math.max(firstColumn, lastColumn),
  }
  const anchor = `${columnLabel(range.startColumn)}${range.startRow + 1}`
  const end = `${columnLabel(range.endColumn)}${range.endRow + 1}`
  const address = kind === 'row' ? `${range.startRow + 1}:${range.endRow + 1}`
    : kind === 'column' ? `${columnLabel(range.startColumn)}:${columnLabel(range.endColumn)}`
      : anchor === end ? anchor : `${anchor}:${end}`
  return { range, address, anchor }
}

export type SelectionStatisticType = 'SUM' | 'COUNT' | 'COUNTA' | 'AVERAGE' | 'MIN' | 'MAX'
export interface SelectionStatistic { id: SelectionStatisticType; label: string; value: number }
export interface SelectionStatistics {
  sum: number
  count: number
  countA: number
  min: number | null
  max: number | null
  average: number | null
  items: SelectionStatistic[]
}
export interface SelectionStatisticsOptions { filteredRows?: readonly number[] }

const NUMERIC_TEXT = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i

function statisticCellValue(cell: ExcelCellView): { defined: boolean; number?: number } {
  // Explicit raw values preserve imported types: text "123" must not become a number.
  if (cell.rawValue !== undefined) {
    if (cell.rawValue === null) return { defined: false }
    return { defined: true, number: typeof cell.rawValue === 'number' && Number.isFinite(cell.rawValue) ? cell.rawValue : undefined }
  }
  // Legacy view cells carry editable text; never parse formatted displayValue or formula source.
  const value = cell.formula ? cell.cachedValue : cell.value
  if (value === undefined || (!cell.formula && value === '')) return { defined: false }
  const number = NUMERIC_TEXT.test(value.trim()) ? Number(value) : NaN
  return { defined: true, number: Number.isFinite(number) ? number : undefined }
}

function normalizeSelection(range: GridRange, bounds: SheetAddressBounds): GridRange | undefined {
  const values = [range.startRow, range.endRow, range.startColumn, range.endColumn]
  if (!values.every(Number.isSafeInteger)) return undefined
  const startRow = Math.max(0, Math.min(range.startRow, range.endRow))
  const endRow = Math.min(bounds.rowCount - 1, Math.max(range.startRow, range.endRow))
  const startColumn = Math.max(0, Math.min(range.startColumn, range.endColumn))
  const endColumn = Math.min(bounds.columnCount - 1, Math.max(range.startColumn, range.endColumn))
  return startRow <= endRow && startColumn <= endColumn ? { startRow, endRow, startColumn, endColumn } : undefined
}

export interface SheetViewState { key: string; anchor: string; ranges: GridRange[]; zoom: number }

/** Restore a view after model/axis changes without putting view operations into content history. */
export function normalizeSheetViewState(sheet: SheetAddressBounds, state: Partial<SheetViewState> = {}): SheetViewState {
  const requestedZoom = state.zoom ?? 100
  const zoom = Number.isFinite(requestedZoom) ? Math.max(10, Math.min(400, requestedZoom)) : 100
  if (!hasValidBounds(sheet)) return { key: 'A1', anchor: 'A1', ranges: [], zoom }
  const clampPoint = (key: string | undefined, fallback: { row: number; column: number }) => {
    const parsed = key && /^\$?[A-Z]+\$?[1-9]\d*$/i.test(key)
      ? parseSheetAddress(key, { rowCount: Number.MAX_SAFE_INTEGER, columnCount: Number.MAX_SAFE_INTEGER }) : undefined
    return parsed ? {
      row: Math.min(sheet.rowCount - 1, parsed.range.startRow),
      column: Math.min(sheet.columnCount - 1, parsed.range.startColumn),
    } : fallback
  }
  let point = clampPoint(state.key, { row: 0, column: 0 })
  let ranges: GridRange[] = []
  for (const range of state.ranges ?? []) {
    const normalized = normalizeSelection(range, sheet)
    if (normalized) ranges = addRange(ranges, normalized)
  }
  if (!ranges.length) ranges = [{ startRow: point.row, endRow: point.row, startColumn: point.column, endColumn: point.column }]
  if (!containsCell(ranges, point.row, point.column)) point = { row: ranges[0]!.startRow, column: ranges[0]!.startColumn }
  const visibleRow = findVisibleSheetRow(sheet, point.row, 1)
  if (visibleRow === undefined) return { key: 'A1', anchor: 'A1', ranges: [], zoom }
  if (visibleRow !== point.row) {
    point.row = visibleRow
    if (!containsCell(ranges, point.row, point.column)) ranges = [{ startRow: point.row, endRow: point.row, startColumn: point.column, endColumn: point.column }]
  }
  const candidate = clampPoint(state.anchor, point)
  const anchor = containsCell(ranges, candidate.row, candidate.column) && !hiddenSheetRows(sheet).has(candidate.row) ? candidate : point
  return {
    key: `${columnLabel(point.column)}${point.row + 1}`,
    anchor: `${columnLabel(anchor.column)}${anchor.row + 1}`,
    ranges, zoom,
  }
}

/** Match Univer's visible-row and countA/count display rules, visiting populated cells once. */
export function calculateSelectionStatistics(
  sheet: ExcelSheetView, ranges: readonly GridRange[], options: SelectionStatisticsOptions = {},
): SelectionStatistics {
  const result: SelectionStatistics = { sum: 0, count: 0, countA: 0, min: null, max: null, average: null, items: [] }
  if (!hasValidBounds(sheet)) return result
  const selections = ranges.map(range => normalizeSelection(range, sheet)).filter((range): range is GridRange => range !== undefined)
  if (!selections.length) return result
  const excludedRows = hiddenSheetRows({ ...sheet, filteredRows: [...(sheet.filteredRows ?? []), ...(options.filteredRows ?? [])] })
  const visited = new Set<string>()
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (!/^\$?[A-Z]+\$?[1-9]\d*$/i.test(key)) continue
    const parsed = parseSheetAddress(key, sheet)
    if (!parsed) continue
    const { startRow: row, startColumn: column, endRow, endColumn } = parsed.range
    // The sparse cell record is addressed by a single coordinate, never a range key.
    if (row !== endRow || column !== endColumn || excludedRows.has(row) || visited.has(parsed.anchor)
      || !selections.some(range => row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn)) continue
    visited.add(parsed.anchor)
    const value = statisticCellValue(cell)
    if (!value.defined) continue
    result.countA += 1
    if (value.number === undefined) continue
    result.count += 1
    result.sum += value.number
    result.min = result.min === null ? value.number : Math.min(result.min, value.number)
    result.max = result.max === null ? value.number : Math.max(result.max, value.number)
  }
  if (result.count > 0) result.average = result.sum / result.count
  if (result.countA <= 1) return result
  if (result.count > 0) {
    result.items.push(
      { id: 'MAX', label: '最大值', value: result.max! },
      { id: 'MIN', label: '最小值', value: result.min! },
      { id: 'SUM', label: '求和', value: result.sum },
    )
  }
  result.items.push({ id: 'COUNTA', label: '计数', value: result.countA })
  if (result.count > 0) result.items.push(
    { id: 'COUNT', label: '数值计数', value: result.count },
    { id: 'AVERAGE', label: '平均值', value: result.average! },
  )
  return result
}

export function selectSelectionStatistic(statistics: SelectionStatistics, preferred: SelectionStatisticType = 'SUM'): SelectionStatistic | undefined {
  return statistics.items.find(item => item.id === preferred)
    ?? statistics.items.find(item => item.id === 'SUM') ?? statistics.items[0]
}

export interface StatisticFormatOptions { locale?: string; numberFormat?: string }

/** Use the active numeric format for value statistics, never for COUNT/COUNTA. */
export function formatSelectionStatistic(item: SelectionStatistic, options: StatisticFormatOptions = {}): string {
  const count = item.id === 'COUNT' || item.id === 'COUNTA'
  const pattern = options.numberFormat ?? 'general'
  const percent = !count && (pattern === 'percent' || /^[#,0.]+%$/.test(pattern))
  const decimal = !count ? /\.([0#]+)/.exec(pattern) : null
  const minimumFractionDigits = !count && pattern === 'number' ? 2 : decimal?.[1]?.split('').filter(character => character === '0').length ?? 0
  const maximumFractionDigits = !count && pattern === 'number' ? 2 : decimal?.[1]?.length ?? (percent ? 0 : count ? 0 : 3)
  const formatted = Number.isFinite(item.value) ? new Intl.NumberFormat(options.locale ?? 'zh-CN', {
    style: percent ? 'percent' : 'decimal', minimumFractionDigits: Math.min(20, minimumFractionDigits),
    maximumFractionDigits: Math.min(20, Math.max(minimumFractionDigits, maximumFractionDigits)),
  }).format(Object.is(item.value, -0) ? 0 : item.value) : '—'
  return `${item.label}: ${formatted}`
}

export interface SheetMenuSource {
  id: string
  name: string
  visible?: boolean
  hiddenState?: 'visible' | 'hidden' | 'veryHidden' | 0 | 1 | 2
}
export type SheetNameValidation = { valid: true } | {
  valid: false
  code: 'empty' | 'too-long' | 'apostrophe' | 'invalid-character' | 'duplicate'
  message: string
}

export function validateSheetName(name: string, sheets: readonly Pick<SheetMenuSource, 'id' | 'name'>[], ignoreId?: string): SheetNameValidation {
  if (!name.trim()) return { valid: false, code: 'empty', message: '工作表名称不能为空。' }
  if (name.length > 31) return { valid: false, code: 'too-long', message: '工作表名称不能超过31个字符。' }
  if (name.startsWith("'") || name.endsWith("'")) return { valid: false, code: 'apostrophe', message: '工作表名称不能以单引号开头或结尾。' }
  if (/[:\\/?*\[\]\u0000-\u001f]/.test(name)) return { valid: false, code: 'invalid-character', message: '工作表名称不能包含 : \\ / ? * [ ] 或控制字符。' }
  if (sheets.some(sheet => sheet.id !== ignoreId && sheet.name.toLowerCase() === name.toLowerCase())) {
    return { valid: false, code: 'duplicate', message: '已存在同名工作表，请使用其他名称。' }
  }
  return { valid: true }
}

function truncateSheetName(name: string, maximum: number): string {
  let result = name.slice(0, maximum)
  const last = result.charCodeAt(result.length - 1)
  if (last >= 0xd800 && last <= 0xdbff) result = result.slice(0, -1)
  return result
}

export function createUniqueSheetName(sheets: readonly Pick<SheetMenuSource, 'id' | 'name'>[], baseName = '工作表'): string {
  const base = validateSheetName(baseName, []).valid ? baseName : '工作表'
  const names = new Set(sheets.map(sheet => sheet.name.toLowerCase()))
  for (let index = 1; ; index += 1) {
    const suffix = String(index)
    const candidate = truncateSheetName(base, 31 - suffix.length) + suffix
    if (!names.has(candidate.toLowerCase())) return candidate
  }
}

export interface SheetMenuEntry { id: string; name: string; hidden: boolean; selected: boolean }

/** VERY_HIDDEN is intentionally omitted; normal hidden sheets remain discoverable for Unhide. */
export function visibleSheetMenuEntries(sheets: readonly SheetMenuSource[], activeId?: string): SheetMenuEntry[] {
  return sheets.filter(sheet => sheet.hiddenState !== 'veryHidden' && sheet.hiddenState !== 2).map(sheet => ({
    id: sheet.id, name: sheet.name, selected: sheet.id === activeId,
    hidden: sheet.hiddenState !== undefined ? sheet.hiddenState === 'hidden' || sheet.hiddenState === 1 : sheet.visible === false,
  }))
}
