import type { CellRange, CellScalar, SpreadsheetSheet } from '../../domain/model.ts'
import { createCell } from '../../domain/model.ts'
import type { ExcelCellView, ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { setSheetViewCellStyleProperties, sheetViewCellAt } from '../../services/excelViewAdapter.ts'
import { resolveSheetCellStyle, setSheetCellStyleProperties, type SheetStyleContext } from './sheet-style-inheritance.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { validateCustomNumberFormat } from './custom-number-format.ts'
import { classifySheetNumberFormat, formatSheetValue, isSheetTextNumberFormat, SHEET_NUMBER_FORMATS, sheetNumberFormatPattern, type SheetNumberFormat } from './sheet-number-format.ts'
import { MAX_SHEET_DECIMALS, numberFormatDecimalPlaces, rawNumberDecimalPlaces, scientificPatternForValue, setNumberFormatDecimalPlaces } from './number-format-pattern.ts'
import { sheetCurrencyPattern } from './sheet-currency.ts'
import { sheetCellNumberValue } from './sheet-number-presentation.ts'

export const SHEET_NUMBER_FORMAT_COMMANDS = Object.freeze({
  set: 'sheet.command.set-number-format', currency: 'sheet.command.set-currency-format', percent: 'sheet.command.set-percent-format',
  increase: 'sheet.command.increase-decimals', decrease: 'sheet.command.decrease-decimals',
  setCells: 'sheet.command.numfmt.set.numfmt',
})
export const SHEET_NUMBER_FORMAT_ALIASES = Object.freeze({
  currency: 'sheet.command.numfmt.set.currency', percent: 'sheet.command.numfmt.set.percent',
  increase: 'sheet.command.numfmt.add.decimal.command', decrease: 'sheet.command.numfmt.subtract.decimal.command',
})

export interface SheetNumberFormatCellInput { row: number; col: number; pattern?: string | null }
export interface SheetNumberFormatCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  ranges?: CellRange[]
  pattern?: string | null
  format?: SheetNumberFormat
  values?: SheetNumberFormatCellInput[]
}
export type SheetNumberFormatAction = { kind: 'set'; pattern: string | null } | { kind: 'currency' } | { kind: 'percent' }
  | { kind: 'decimals'; delta: -1 | 1 } | { kind: 'cells'; values: SheetNumberFormatCellInput[] }
interface FormatTarget { row: number; col: number; value: CellScalar; pattern: string }
export interface SheetNumberFormatPatch { row: number; col: number; pattern: string }

function normalizedPattern(input: string | null | undefined): string {
  if (input === null || input === undefined || input === '' || /^general$/i.test(input)) return 'General'
  return validateCustomNumberFormat(input).pattern
}

export function numberFormatActionForCommand(id: string, params: unknown, region = 'zh-CN'): SheetNumberFormatAction | undefined {
  if (id === SHEET_NUMBER_FORMAT_COMMANDS.currency || id === SHEET_NUMBER_FORMAT_ALIASES.currency) return { kind: 'currency' }
  if (id === SHEET_NUMBER_FORMAT_COMMANDS.percent || id === SHEET_NUMBER_FORMAT_ALIASES.percent) return { kind: 'percent' }
  if (id === SHEET_NUMBER_FORMAT_COMMANDS.increase || id === SHEET_NUMBER_FORMAT_ALIASES.increase) return { kind: 'decimals', delta: 1 }
  if (id === SHEET_NUMBER_FORMAT_COMMANDS.decrease || id === SHEET_NUMBER_FORMAT_ALIASES.decrease) return { kind: 'decimals', delta: -1 }
  if (id !== SHEET_NUMBER_FORMAT_COMMANDS.set && id !== SHEET_NUMBER_FORMAT_COMMANDS.setCells) return undefined
  if (typeof params === 'string') return { kind: 'set', pattern: Object.hasOwn(SHEET_NUMBER_FORMATS, params) ? sheetNumberFormatPattern(params as SheetNumberFormat, region) : normalizedPattern(params) }
  if (!params || typeof params !== 'object') return undefined
  if ('values' in params && Array.isArray(params.values)) {
    const values: SheetNumberFormatCellInput[] = []
    for (const value of params.values) {
      if (!value || typeof value !== 'object' || typeof value.row !== 'number' || typeof value.col !== 'number'
        || value.pattern !== undefined && value.pattern !== null && typeof value.pattern !== 'string') throw new TypeError('数字格式单元格参数无效。')
      values.push({ row: value.row, col: value.col, pattern: value.pattern })
    }
    return { kind: 'cells', values }
  }
  if ('pattern' in params && (params.pattern === null || typeof params.pattern === 'string')) return { kind: 'set', pattern: params.pattern }
  if ('format' in params && typeof params.format === 'string' && Object.hasOwn(SHEET_NUMBER_FORMATS, params.format)) return { kind: 'set', pattern: sheetNumberFormatPattern(params.format as SheetNumberFormat, region) }
  return undefined
}

export function assertNumberFormatRanges(ranges: readonly CellRange[]): void {
  for (const range of ranges) {
    if (![range.startRow, range.startCol, range.endRow, range.endCol].every(Number.isSafeInteger) || range.startRow < 0 || range.startCol < 0
      || range.endRow < range.startRow || range.endCol < range.startCol || range.endRow >= 1_048_576 || range.endCol >= 16_384) throw new RangeError('数字格式范围无效。')
  }
}

function actionRanges(ranges: readonly CellRange[], action: SheetNumberFormatAction): CellRange[] {
  const targets = action.kind === 'cells' ? action.values.map(value => ({ startRow: value.row, endRow: value.row, startCol: value.col, endCol: value.col })) : [...ranges]
  assertNumberFormatRanges(targets)
  return targets
}

function* positions(ranges: readonly CellRange[]): Generator<{ row: number; col: number }> {
  const visited = new Set<number>()
  for (const range of ranges) for (let row = range.startRow; row <= range.endRow; row += 1) for (let col = range.startCol; col <= range.endCol; col += 1) {
    const key = row * 16_384 + col
    if (visited.has(key)) continue
    visited.add(key)
    yield { row, col }
  }
}

/** One plan is shared by domain commands and standalone Vue hosts; it never edits raw data. */
export function createSheetNumberFormatPlan(targets: readonly FormatTarget[], action: SheetNumberFormatAction, region = 'zh-CN'): SheetNumberFormatPatch[] {
  const requested = action.kind === 'cells' ? new Map(action.values.map(value => [`${value.row}:${value.col}`, normalizedPattern(value.pattern)])) : undefined
  const pattern = action.kind === 'set' ? normalizedPattern(action.pattern) : action.kind === 'currency' ? sheetCurrencyPattern(region) : action.kind === 'percent' ? '0%' : undefined
  const maximum = action.kind === 'decimals' ? targets.reduce((maximum, cell) => Math.max(maximum,
    !cell.pattern || /^general$/i.test(cell.pattern) ? typeof cell.value === 'number' ? rawNumberDecimalPlaces(cell.value) : 0 : numberFormatDecimalPlaces(cell.pattern)), 0) : 0
  const generalDecimals = action.kind === 'decimals' ? Math.max(0, Math.min(MAX_SHEET_DECIMALS, maximum + action.delta)) : 0
  return targets.flatMap(cell => {
    let next = requested?.get(`${cell.row}:${cell.col}`) ?? pattern
    if (action.kind === 'decimals') {
      const general = !cell.pattern || /^general$/i.test(cell.pattern)
      const scientific = typeof cell.value === 'number' ? scientificPatternForValue(cell.value) : undefined
      next = setNumberFormatDecimalPlaces(general ? scientific ?? '0' : cell.pattern,
        general ? generalDecimals : numberFormatDecimalPlaces(cell.pattern) + action.delta)
    }
    if (next === undefined || next === (cell.pattern || 'General')) return []
    // Fixed-version SetNumfmt preserves General scientific values when a lossy number/date format is requested.
    if (typeof cell.value === 'number' && scientificPatternForValue(cell.value)
      && next !== 'General' && !isSheetTextNumberFormat(next) && classifySheetNumberFormat(next) !== 'scientific') return []
    return [{ row: cell.row, col: cell.col, pattern: next }]
  })
}

function assertEditable(sheet: SpreadsheetSheet | ExcelSheetView, ranges: CellRange[], userId: string): void {
  if (!canEditWorksheetProtection(sheet, userId) || !canEditProtectedRanges(sheet, ranges, userId)) throw new Error('当前选区受保护，不能修改数字格式。')
}

export function applySheetNumberFormatCommand(sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetNumberFormatAction, region = 'zh-CN', date1904 = false, userId = 'local-user', styleContext: SheetStyleContext = {}): boolean {
  const scope = actionRanges(ranges, action)
  if (!scope.length) return false
  assertEditable(sheet, scope, userId)
  const targets = [...positions(scope)].map(({ row, col }) => ({ row, col, value: sheet.cells[`${row}:${col}`]?.value ?? null,
    pattern: resolveSheetCellStyle(sheet, row, col, sheet.cells[`${row}:${col}`], styleContext).numberFormat }))
  const patches = createSheetNumberFormatPlan(targets, action, region)
  for (const { row, col, pattern } of patches) {
    const cell = sheet.cells[`${row}:${col}`] ?? createCell(row, col)
    cell.style = resolveSheetCellStyle(sheet, row, col, sheet.cells[`${row}:${col}`], styleContext)
    setSheetCellStyleProperties(cell, { numberFormat: pattern })
    if (!cell.formula || cell.value !== null) cell.display = formatSheetValue(cell.value, pattern, date1904, region) ?? cell.display
    sheet.cells[`${row}:${col}`] = cell
  }
  return patches.length > 0
}

const address = (row: number, col: number) => {
  let letters = ''
  for (let index = col; index >= 0; index = Math.floor(index / 26) - 1) letters = String.fromCharCode(index % 26 + 65) + letters
  return `${letters}${row + 1}`
}
export function applySheetViewNumberFormatCommand(sheet: ExcelSheetView, ranges: readonly CellRange[], action: SheetNumberFormatAction, region = 'zh-CN', userId = 'local-user'): boolean {
  const scope = actionRanges(ranges, action)
  if (!scope.length) return false
  assertEditable(sheet, scope, userId)
  const targets = [...positions(scope)].map(({ row, col }) => {
    const cell = sheetViewCellAt(sheet, row, col)
    return { row, col, value: sheetCellNumberValue(cell) ?? null, pattern: cell?.rawNumberFormat ?? SHEET_NUMBER_FORMATS[cell?.numberFormat ?? 'general'] }
  })
  const patches = createSheetNumberFormatPlan(targets, action, region)
  for (const { row, col, pattern } of patches) {
    const key = address(row, col), cell = sheet.cells[key] ?? { value: '' }
    const value = sheetCellNumberValue(cell)
    setSheetViewCellStyleProperties(sheet, row, col, { numberFormat: pattern })
    sheet.cells[key] = { ...sheet.cells[key]!,
      displayValue: value === undefined ? cell.displayValue : formatSheetValue(value, pattern, sheet.date1904, region) ?? cell.displayValue }
  }
  return patches.length > 0
}
