import { createCell, defaultCellStyle, type CellRange, type SpreadsheetSheet } from '../../domain/model.ts'
import { cellKey, parseCellKey, type ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { RESET_BACKGROUND_COLOR, RESET_TEXT_COLOR } from './sheet-color-reset.ts'
import { sheetColorHex, type SheetColorField } from './sheet-colors.ts'

export type SheetFontMark = 'bold' | 'italic' | 'underline' | 'strike'
export type SheetFontFormatAction = { kind: 'mark'; field: SheetFontMark; value?: boolean }
  | { kind: 'color'; field: SheetColorField; value: string | null }
  | { kind: 'family'; value: string | null }
  | { kind: 'size'; value: number | null }
  | { kind: 'size-step'; direction: -1 | 1 }

export const SHEET_FONT_FORMAT_COMMANDS = Object.freeze({
  bold: 'sheet.command.toggle-bold', italic: 'sheet.command.toggle-italic',
  underline: 'sheet.command.toggle-underline', strike: 'sheet.command.toggle-strike',
  fontWeight: 'sheet.command.set-font-weight', fontStyle: 'sheet.command.set-font-style',
  fontFamily: 'sheet.command.set-font-family', fontSize: 'sheet.command.set-font-size',
  increaseFontSize: 'sheet.command.increase-font-size', decreaseFontSize: 'sheet.command.decrease-font-size',
  textColor: 'sheet.command.set-color', fontColor: 'sheet.command.set-font-color',
  background: 'sheet.command.set-background', resetTextColor: RESET_TEXT_COLOR, resetBackground: RESET_BACKGROUND_COLOR,
})

export const SHEET_FONT_FORMAT_ALIASES = Object.freeze({
  rangeBold: 'sheet.command.set-range-bold', bold: 'sheet.command.set-bold',
  rangeItalic: 'sheet.command.set-range-italic', italic: 'sheet.command.set-italic',
  rangeUnderline: 'sheet.command.set-range-underline', underline: 'sheet.command.set-underline',
  rangeStrike: 'sheet.command.set-range-stroke', strike: 'sheet.command.set-stroke',
  rangeTextColor: 'sheet.command.set-range-text-color', textColor: 'sheet.command.set-text-color',
  background: 'sheet.command.set-background-color', resetRangeTextColor: 'sheet.command.reset-range-text-color',
  rangeFontFamily: 'sheet.command.set-range-font-family', rangeFontSize: 'sheet.command.set-range-fontsize',
  rangeFontIncrease: 'sheet.command.set-range-font-increase', rangeFontDecrease: 'sheet.command.set-range-font-decrease',
})
export const SHEET_FONT_FORMAT_COMMAND_IDS: readonly string[] = [...Object.values(SHEET_FONT_FORMAT_COMMANDS), ...Object.values(SHEET_FONT_FORMAT_ALIASES)]
const textColorIds: readonly string[] = [SHEET_FONT_FORMAT_COMMANDS.textColor, SHEET_FONT_FORMAT_COMMANDS.fontColor, SHEET_FONT_FORMAT_ALIASES.textColor, SHEET_FONT_FORMAT_ALIASES.rangeTextColor]
const backgroundIds: readonly string[] = [SHEET_FONT_FORMAT_COMMANDS.background, SHEET_FONT_FORMAT_ALIASES.background]

export interface SheetFontPrimaryPosition { row: number; col: number }
export type SheetFontRangeInput = CellRange | { startRow: number; startColumn: number; endRow: number; endColumn: number }
export interface SheetFontFormatCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  ranges?: SheetFontRangeInput[]
  range?: SheetFontRangeInput
  /** Size steppers use the active cell; explicit API ranges default to their last start cell. */
  primary?: SheetFontPrimaryPosition
  value?: unknown
}
export type SheetFontFormatCommandParams = SheetFontFormatCommandInput | string | number | boolean | null | undefined

export interface SheetFontFormatShortcutInput {
  key: string; ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; altKey?: boolean;
  isComposing?: boolean; keyCode?: number; defaultPrevented?: boolean;
}

export function sheetFontFormatCommandForShortcut(event: SheetFontFormatShortcutInput): string | undefined {
  if ((!event.ctrlKey && !event.metaKey) || event.altKey || event.isComposing || event.keyCode === 229 || event.defaultPrevented) return undefined
  const key = event.key.toLowerCase()
  if (event.shiftKey) return key === 'x' ? SHEET_FONT_FORMAT_COMMANDS.strike : undefined
  return key === 'b' ? SHEET_FONT_FORMAT_COMMANDS.bold : key === 'i' ? SHEET_FONT_FORMAT_COMMANDS.italic : key === 'u' ? SHEET_FONT_FORMAT_COMMANDS.underline : undefined
}

function commandValue(params: unknown): unknown {
  return params && typeof params === 'object' ? 'value' in params ? params.value : undefined : params
}

export function sheetFontFormatCommandInput(params: unknown): SheetFontFormatCommandInput {
  return params && typeof params === 'object' && !Array.isArray(params) ? params as SheetFontFormatCommandInput
    : params === undefined ? {} : { value: params }
}
/** The fixed toolbar restricts input to 6–400; the public API retains finite positive sizes. */
export function normalizeSheetFontSize(value: unknown): number | null {
  if (value === null) return null
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) throw new RangeError('字号必须是有限正数或 null。')
  return value
}
export function nextSheetFontFormatSize(size: number, direction: -1 | 1): number {
  if (normalizeSheetFontSize(size) === null) throw new RangeError('字号步进需要有限正数。')
  if (direction !== -1 && direction !== 1) throw new RangeError('字号步进方向无效。')
  return direction === 1 ? size >= 400 ? 400 : size + 1 : size <= 6 ? 6 : size - 1
}

export function fontFormatActionForCommand(id: string, params?: unknown): SheetFontFormatAction | undefined {
  const marks: ReadonlyArray<[SheetFontMark, readonly string[]]> = [
    ['bold', [SHEET_FONT_FORMAT_COMMANDS.bold, SHEET_FONT_FORMAT_ALIASES.bold, SHEET_FONT_FORMAT_ALIASES.rangeBold]],
    ['italic', [SHEET_FONT_FORMAT_COMMANDS.italic, SHEET_FONT_FORMAT_ALIASES.italic, SHEET_FONT_FORMAT_ALIASES.rangeItalic]],
    ['underline', [SHEET_FONT_FORMAT_COMMANDS.underline, SHEET_FONT_FORMAT_ALIASES.underline, SHEET_FONT_FORMAT_ALIASES.rangeUnderline]],
    ['strike', [SHEET_FONT_FORMAT_COMMANDS.strike, SHEET_FONT_FORMAT_ALIASES.strike, SHEET_FONT_FORMAT_ALIASES.rangeStrike]],
  ]
  const value = commandValue(params)
  if (id === SHEET_FONT_FORMAT_COMMANDS.fontFamily || id === SHEET_FONT_FORMAT_ALIASES.rangeFontFamily) {
    if (value === null || typeof value === 'string' && !!value.trim()) return { kind: 'family', value }
    throw new TypeError('字体名称必须是非空字符串或 null。')
  }
  if (id === SHEET_FONT_FORMAT_COMMANDS.fontSize || id === SHEET_FONT_FORMAT_ALIASES.rangeFontSize) return { kind: 'size', value: normalizeSheetFontSize(value) }
  if (id === SHEET_FONT_FORMAT_COMMANDS.increaseFontSize || id === SHEET_FONT_FORMAT_ALIASES.rangeFontIncrease) return { kind: 'size-step', direction: 1 }
  if (id === SHEET_FONT_FORMAT_COMMANDS.decreaseFontSize || id === SHEET_FONT_FORMAT_ALIASES.rangeFontDecrease) return { kind: 'size-step', direction: -1 }
  for (const [field, ids] of marks) {
    if (!ids.includes(id)) continue
    if (value === undefined) return { kind: 'mark', field }
    if (typeof value === 'boolean' || value === 0 || value === 1) return { kind: 'mark', field, value: value === true || value === 1 }
    throw new TypeError('字体开关参数必须是布尔值。')
  }
  if (id === SHEET_FONT_FORMAT_COMMANDS.fontWeight) {
    if (value === 'bold' || value === 'normal' || value === null) return { kind: 'mark', field: 'bold', value: value === 'bold' }
    throw new TypeError('字体粗细必须是 bold、normal 或 null。')
  }
  if (id === SHEET_FONT_FORMAT_COMMANDS.fontStyle) {
    if (value === 'italic' || value === 'normal' || value === null) return { kind: 'mark', field: 'italic', value: value === 'italic' }
    throw new TypeError('字体样式必须是 italic、normal 或 null。')
  }
  if (id === RESET_TEXT_COLOR || id === SHEET_FONT_FORMAT_ALIASES.resetRangeTextColor) return { kind: 'color', field: 'color', value: null }
  if (id === RESET_BACKGROUND_COLOR) return { kind: 'color', field: 'background', value: null }
  const field = textColorIds.includes(id) ? 'color' : backgroundIds.includes(id) ? 'background' : undefined
  if (!field) return undefined
  if (value === null) return { kind: 'color', field, value }
  if (typeof value !== 'string') throw new TypeError('单元格颜色必须是字符串或 null。')
  if (!value) return undefined
  // Preserve the existing API's string color contract and original imported values.
  return { kind: 'color', field, value: sheetColorHex(value) ?? value }
}

export function sheetFontFormatInputRanges(input?: SheetFontFormatCommandInput): CellRange[] | undefined {
  const ranges = input?.ranges !== undefined ? input.ranges : input?.range !== undefined ? [input.range] : undefined
  if (ranges === undefined) return undefined
  if (!Array.isArray(ranges)) throw new TypeError('字体 ranges 必须是范围数组。')
  const normalized = ranges.map(range => {
    if (!range || typeof range !== 'object') throw new TypeError('字体范围必须包含行列坐标。')
    return 'startCol' in range ? { ...range } : {
      startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
    }
  })
  assertSheetFontFormatRanges(normalized)
  return ranges.every((range): range is CellRange => 'startCol' in range) ? ranges : normalized
}

export function assertSheetFontFormatRanges(ranges: readonly CellRange[]): void {
  for (const range of ranges) {
    if (!range || ![range.startRow, range.startCol, range.endRow, range.endCol].every(Number.isSafeInteger)
      || range.startRow < 0 || range.startCol < 0 || range.endRow < range.startRow || range.endCol < range.startCol
      || range.endRow >= 1_048_576 || range.endCol >= 16_384) throw new RangeError('单元格格式范围无效。')
  }
}

function assertEditable(sheet: SpreadsheetSheet | ExcelSheetView, ranges: readonly CellRange[], userId: string): void {
  assertSheetFontFormatRanges(ranges)
  if (!canEditWorksheetProtection(sheet, userId) || !canEditProtectedRanges(sheet, [...ranges], userId)) throw new Error('当前选区受保护，不能修改单元格格式。')
}

function inScope(ranges: readonly CellRange[], row: number, col: number): boolean {
  return ranges.some(range => row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol)
}

function* positions(ranges: readonly CellRange[], filteredRows: ReadonlySet<number>): Generator<{ row: number; col: number }> {
  const visited = new Set<number>()
  for (const range of ranges) for (let row = range.startRow; row <= range.endRow; row += 1) {
    if (filteredRows.has(row)) continue
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const index = row * 16_384 + col
      if (visited.has(index)) continue
      visited.add(index)
      yield { row, col }
    }
  }
}

type FontPropertyAction = Extract<SheetFontFormatAction, { kind: 'family' | 'size' | 'size-step' }>
type FontProperty = { family: string } | { size: number }
const fontDefaults = defaultCellStyle().font
function fontProperty(action: FontPropertyAction, primarySize: number): FontProperty {
  if (action.kind === 'family') {
    if (action.value !== null && (typeof action.value !== 'string' || !action.value.trim())) throw new TypeError('字体名称必须是非空字符串或 null。')
    return { family: action.value ?? fontDefaults.family }
  }
  return { size: action.kind === 'size' ? normalizeSheetFontSize(action.value) ?? fontDefaults.size : nextSheetFontFormatSize(primarySize, action.direction) }
}
function fontPrimary(ranges: readonly CellRange[], supplied?: SheetFontPrimaryPosition): SheetFontPrimaryPosition {
  if (supplied !== undefined && (supplied === null || typeof supplied !== 'object')) throw new TypeError('字号步进主单元格必须包含行列坐标。')
  const last = ranges.at(-1)!, primary = supplied ?? { row: last.startRow, col: last.startCol }
  assertSheetFontFormatRanges([{ startRow: primary.row, endRow: primary.row, startCol: primary.col, endCol: primary.col }])
  if (!inScope(ranges, primary.row, primary.col)) throw new RangeError('字号步进主单元格必须位于选区内。')
  return primary
}
function* resetPositions(existing: Iterable<SheetFontPrimaryPosition>, ranges: readonly CellRange[], filteredRows: ReadonlySet<number>): Generator<SheetFontPrimaryPosition> {
  for (const position of existing) if (!filteredRows.has(position.row) && inScope(ranges, position.row, position.col)) yield position
}

/** Model and standalone view paths share mixed-selection, filtered-row and sparse reset semantics. */
export function applySheetFontFormatCommand(sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetFontFormatAction, userId = 'local-user', filteredRows = sheet.autoFilter?.filteredRows ?? [], primary?: SheetFontPrimaryPosition): boolean {
  if (!ranges.length) return false
  assertEditable(sheet, ranges, userId)
  const filtered = new Set(filteredRows)
  if (action.kind === 'family' || action.kind === 'size' || action.kind === 'size-step') {
    const main = action.kind === 'size-step' ? fontPrimary(ranges, primary) : undefined
    const property = fontProperty(action, main ? sheet.cells[`${main.row}:${main.col}`]?.style.font.size ?? fontDefaults.size : fontDefaults.size)
    const reset = action.kind !== 'size-step' && action.value === null
    const targets = reset ? resetPositions(Object.values(sheet.cells), ranges, filtered) : positions(ranges, filtered)
    let changed = false
    for (const { row, col } of targets) {
      const key = `${row}:${col}`, current = sheet.cells[key]
      if ('family' in property ? (current?.style.font.family ?? fontDefaults.family) === property.family : (current?.style.font.size ?? fontDefaults.size) === property.size) continue
      const cell = current ?? createCell(row, col)
      cell.style.font = { ...cell.style.font, ...property }
      sheet.cells[key] = cell
      changed = true
    }
    return changed
  }
  if (action.kind === 'color' && action.value === null) {
    let changed = false
    for (const cell of Object.values(sheet.cells)) {
      if (filtered.has(cell.row) || !inScope(ranges, cell.row, cell.col)) continue
      if (action.field === 'color' ? cell.style.font.color == null : cell.style.fill == null) continue
      if (action.field === 'color') cell.style.font.color = null
      else cell.style.fill = null
      changed = true
    }
    return changed
  }
  const targets = [...positions(ranges, filtered)]
  const value = action.kind === 'mark' ? action.value ?? !targets.every(({ row, col }) => sheet.cells[`${row}:${col}`]?.style.font[action.field] === true) : action.value
  let changed = false
  for (const { row, col } of targets) {
    const key = `${row}:${col}`, current = sheet.cells[key]
    const previous = action.kind === 'mark' ? current?.style.font[action.field] ?? false : action.field === 'color' ? current?.style.font.color ?? null : current?.style.fill ?? null
    if (previous === value) continue
    const cell = current ?? createCell(row, col)
    if (action.kind === 'mark') cell.style.font[action.field] = value === true
    else if (action.field === 'color') cell.style.font.color = action.value
    else cell.style.fill = action.value
    sheet.cells[key] = cell
    changed = true
  }
  return changed
}

export function applySheetViewFontFormatCommand(sheet: ExcelSheetView, ranges: readonly CellRange[], action: SheetFontFormatAction, userId = 'local-user', primary?: SheetFontPrimaryPosition): boolean {
  if (!ranges.length) return false
  assertEditable(sheet, ranges, userId)
  const filtered = new Set(sheet.filteredRows ?? sheet.autoFilter?.filteredRows ?? [])
  if (action.kind === 'family' || action.kind === 'size' || action.kind === 'size-step') {
    const main = action.kind === 'size-step' ? fontPrimary(ranges, primary) : undefined
    const property = fontProperty(action, main ? sheet.cells[cellKey(main.row, main.col)]?.fontSize ?? fontDefaults.size : fontDefaults.size)
    const existing = () => Object.keys(sheet.cells).flatMap(key => { const position = parseCellKey(key); return position ? [position] : [] })
    const reset = action.kind !== 'size-step' && action.value === null
    const targets = reset ? resetPositions(existing(), ranges, filtered) : positions(ranges, filtered)
    let changed = false
    for (const { row, col } of targets) {
      const key = cellKey(row, col), current = sheet.cells[key]
      if ('family' in property ? (current?.fontFamily ?? fontDefaults.family) === property.family : (current?.fontSize ?? fontDefaults.size) === property.size) continue
      sheet.cells[key] = { ...(current ?? { value: '' }), ...('family' in property ? { fontFamily: property.family } : { fontSize: property.size }) }
      changed = true
    }
    return changed
  }
  if (action.kind === 'color' && action.value === null) {
    let changed = false
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const position = parseCellKey(key)
      if (!position || filtered.has(position.row) || !inScope(ranges, position.row, position.col) || cell[action.field] == null) continue
      sheet.cells[key] = { ...cell, [action.field]: null }
      changed = true
    }
    return changed
  }
  const targets = [...positions(ranges, filtered)]
  const value = action.kind === 'mark' ? action.value ?? !targets.every(({ row, col }) => sheet.cells[cellKey(row, col)]?.[action.field] === true) : action.value
  let changed = false
  for (const { row, col } of targets) {
    const key = cellKey(row, col), cell = sheet.cells[key]
    if ((cell?.[action.field] ?? (action.kind === 'mark' ? false : null)) === value) continue
    sheet.cells[key] = { ...(cell ?? { value: '' }), [action.field]: value }
    changed = true
  }
  return changed
}
