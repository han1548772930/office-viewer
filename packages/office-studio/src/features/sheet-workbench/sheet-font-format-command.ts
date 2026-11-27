import { createCell, defaultCellStyle, type CellRange, type SpreadsheetCell, type SpreadsheetSheet } from '../../domain/model.ts'
import { cellKey, parseCellKey, sheetViewCellStylePatch, toExcelCellView, type ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { RESET_BACKGROUND_COLOR, RESET_TEXT_COLOR } from './sheet-color-reset.ts'
import { sheetColorHex, type SheetColorField } from './sheet-colors.ts'
import { setSheetCellRichTextFont, sheetCellFontValues } from './sheet-cell-rich-text-style.ts'
import { adoptSheetCellStyleEdits, clearSheetCellStyleProperties, resolveSheetCellStyle, setSheetCellStyleProperties,
  sheetInheritedCellStyle, composeSheetCellStyles, hasInheritedSheetStyle, type SheetStyleContext } from './sheet-style-inheritance.ts'

export type SheetFontMark = 'bold' | 'italic' | 'underline' | 'strike'
export type SheetFontFormatAction = { kind: 'mark'; field: SheetFontMark; value?: boolean | null }
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
  if (event.shiftKey) {
    if (event.keyCode === 190) return SHEET_FONT_FORMAT_COMMANDS.increaseFontSize
    if (event.keyCode === 188) return SHEET_FONT_FORMAT_COMMANDS.decreaseFontSize
    if (key === '.' || key === '>') return SHEET_FONT_FORMAT_COMMANDS.increaseFontSize
    if (key === ',' || key === '<') return SHEET_FONT_FORMAT_COMMANDS.decreaseFontSize
    return key === 'x' ? SHEET_FONT_FORMAT_COMMANDS.strike : undefined
  }
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
    if (value === null) return { kind: 'mark', field, value: null }
    if (typeof value === 'boolean' || value === 0 || value === 1) return { kind: 'mark', field, value: value === true || value === 1 }
    throw new TypeError('字体开关参数必须是布尔值。')
  }
  if (id === SHEET_FONT_FORMAT_COMMANDS.fontWeight) {
    if (value === 'bold' || value === 'normal' || value === null) return { kind: 'mark', field: 'bold', value: value === null ? null : value === 'bold' }
    throw new TypeError('字体粗细必须是 bold、normal 或 null。')
  }
  if (id === SHEET_FONT_FORMAT_COMMANDS.fontStyle) {
    if (value === 'italic' || value === 'normal' || value === null) return { kind: 'mark', field: 'italic', value: value === null ? null : value === 'italic' }
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
export function applySheetFontFormatCommand(sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetFontFormatAction, userId = 'local-user', filteredRows = sheet.autoFilter?.filteredRows ?? [], primary?: SheetFontPrimaryPosition, styleContext: SheetStyleContext = {}): boolean {
  if (!ranges.length) return false
  assertEditable(sheet, ranges, userId)
  const filtered = new Set(filteredRows)
  if (action.kind === 'family' || action.kind === 'size' || action.kind === 'size-step') {
    const main = action.kind === 'size-step' ? fontPrimary(ranges, primary) : undefined
    const primaryCell = main ? cloneModel(sheet.cells[`${main.row}:${main.col}`] ?? createCell(main.row, main.col)) : undefined
    if (primaryCell && !filtered.has(primaryCell.row)) adoptSheetCellStyleEdits(sheet, primaryCell, styleContext)
    const property = fontProperty(action, main ? resolveSheetCellStyle(sheet, main.row, main.col, primaryCell, styleContext).font.size : fontDefaults.size)
    const reset = action.kind !== 'size-step' && action.value === null
    const targets = reset ? resetPositions(Object.values(sheet.cells), ranges, filtered) : positions(ranges, filtered)
    let changed = false
    for (const { row, col } of targets) {
      const key = `${row}:${col}`, current = sheet.cells[key]
      if (current) {
        adoptSheetCellStyleEdits(sheet, current, styleContext)
        current.style = resolveSheetCellStyle(sheet, row, col, current, styleContext)
      }
      const rich = setSheetCellRichTextFont(current?.richText, 'family' in property ? 'family' : 'size', reset ? null : 'family' in property ? property.family : property.size)
      const cell = current ?? createCell(row, col)
      if (!current) cell.style = sheetInheritedCellStyle(sheet, row, col, styleContext)
      const field = 'family' in property ? 'family' : 'size'
      const value = 'family' in property ? property.family : property.size
      if (reset) {
        const cleared = clearSheetCellStyleProperties(cell, [`font.${field}`])
        if (!cleared && !rich.changed) continue
        cell.style = resolveSheetCellStyle(sheet, row, col, cell, styleContext)
      } else {
        const same = cell.style.font[field] === value && (!cell.authoredStyleReference && !hasInheritedSheetStyle(sheet, row, col, styleContext, cell.themeStyle)
          || !cell.authoredStyleReference && cell.authoredStyle === undefined || cell.authoredStyle?.font?.[field] === value)
        if (same && !rich.changed) continue
        setSheetCellStyleProperties(cell, { font: property })
      }
      if (rich.changed) cell.richText = rich.runs ?? undefined
      sheet.cells[key] = cell
      changed = true
    }
    return changed
  }
  for (const cell of Object.values(sheet.cells)) if (!filtered.has(cell.row) && inScope(ranges, cell.row, cell.col)) {
    adoptSheetCellStyleEdits(sheet, cell, styleContext)
    cell.style = resolveSheetCellStyle(sheet, cell.row, cell.col, cell, styleContext)
  }
  if (action.kind === 'mark' && action.value === null) {
    let changed = false
    for (const cell of Object.values(sheet.cells)) {
      if (filtered.has(cell.row) || !inScope(ranges, cell.row, cell.col)) continue
      const rich = setSheetCellRichTextFont(cell.richText, action.field, null)
      const cleared = clearSheetCellStyleProperties(cell, [`font.${action.field}`])
      if (!cleared && !rich.changed) continue
      cell.style = resolveSheetCellStyle(sheet, cell.row, cell.col, cell, styleContext)
      if (rich.changed) cell.richText = rich.runs ?? undefined
      changed = true
    }
    return changed
  }
  const targets = action.kind === 'color' && action.value === null
    ? [...resetPositions(Object.values(sheet.cells), ranges, filtered)] : [...positions(ranges, filtered)]
  const value = action.kind === 'mark' ? action.value ?? !targets.every(({ row, col }) => {
    const cell = sheet.cells[`${row}:${col}`]
    return sheetCellFontValues(toExcelCellView(cell ?? { ...createCell(row, col), style: sheetInheritedCellStyle(sheet, row, col, styleContext) }), action.field).every(value => value === true)
  }) : action.value
  let changed = false
  for (const { row, col } of targets) {
    const key = `${row}:${col}`, current = sheet.cells[key]
    const effective = current?.style ?? sheetInheritedCellStyle(sheet, row, col, styleContext)
    const previous = action.kind === 'mark' ? effective.font[action.field] : action.field === 'color' ? effective.font.color : effective.fill
    const rich = action.kind === 'mark' ? setSheetCellRichTextFont(current?.richText, action.field, value === true)
      : action.field === 'color' ? setSheetCellRichTextFont(current?.richText, 'color', action.value) : undefined
    const direct = current?.authoredStyle
    const sameDirect = !current?.authoredStyleReference && !hasInheritedSheetStyle(sheet, row, col, styleContext, current?.themeStyle) || current && (!current.authoredStyleReference && direct === undefined
      || (action.kind === 'mark' ? direct?.font?.[action.field] : action.field === 'color' ? direct?.font?.color : direct?.fill) === value)
    if (previous === value && sameDirect && !rich?.changed) continue
    const cell = current ?? createCell(row, col)
    if (!current) cell.style = sheetInheritedCellStyle(sheet, row, col, styleContext)
    if (action.kind === 'mark') setSheetCellStyleProperties(cell, { font: { [action.field]: value === true } })
    else if (action.field === 'color') setSheetCellStyleProperties(cell, { font: { color: action.value } })
    else setSheetCellStyleProperties(cell, { fill: action.value })
    if (rich?.changed) cell.richText = rich.runs ?? undefined
    sheet.cells[key] = cell
    changed = true
  }
  return changed
}

export function applySheetViewFontFormatCommand(sheet: ExcelSheetView, ranges: readonly CellRange[], action: SheetFontFormatAction, userId = 'local-user', primary?: SheetFontPrimaryPosition): boolean {
  if (!ranges.length) return false
  assertEditable(sheet, ranges, userId)
  const model: SpreadsheetSheet = { id: sheet.id, name: sheet.name, visible: sheet.visible !== false, cells: {},
    mergedRanges: [], columnWidths: {}, rowHeights: {}, frozenRows: 0, frozenCols: 0, showGridlines: sheet.showGridlines !== false,
    defaultStyle: sheet.defaultStyle, rowStyles: sheet.rowStyles, columnStyles: sheet.columnStyles,
    rowStyleSpans: sheet.rowStyleSpans, columnStyleSpans: sheet.columnStyleSpans, rowStylePriority: sheet.rowStylePriority,
    rangeProtectionRules: sheet.rangeProtectionRules, worksheetProtection: sheet.worksheetProtection }
  const context: SheetStyleContext = { defaultStyle: sheet.workbookDefaultStyle, styles: sheet.styles }
  const signature = (cell: SpreadsheetCell) => JSON.stringify({ font: cell.style.font, fill: cell.style.fill, authored: cell.authoredStyle,
    reference: cell.authoredStyleReference, cleared: cell.clearedStyleProperties, rich: cell.richText })
  const before = new Map<string, string>()
  for (const [key, source] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position || !inScope(ranges, position.row, position.col)) continue
    const patch = sheetViewCellStylePatch(source), cell: SpreadsheetCell = { ...createCell(position.row, position.col, source.rawValue ?? source.value),
      formula: source.formula ?? null, display: source.displayValue ?? source.value, dataType: source.dataType,
      style: composeSheetCellStyles(patch), authoredStyle: source.authoredStyle === null ? {} : cloneModel(source.authoredStyle ?? (source.authoredStyleReference ? {} : patch)),
      authoredStyleReference: source.authoredStyleReference ?? undefined, clearedStyleProperties: source.clearedStyleProperties ? [...source.clearedStyleProperties] : undefined,
      themeStyle: source.themeStyle, richText: source.richText == null ? undefined : cloneModel(source.richText) as SpreadsheetCell['richText'] }
    cell.style = resolveSheetCellStyle(model, cell.row, cell.col, cell, context)
    model.cells[cell.id] = cell; before.set(cell.id, signature(cell))
  }
  const changed = applySheetFontFormatCommand(model, ranges, action, userId, sheet.filteredRows ?? sheet.autoFilter?.filteredRows ?? [], primary, context)
  if (!changed) return false
  for (const cell of Object.values(model.cells)) {
    if (signature(cell) === before.get(cell.id)) continue
    const key = cellKey(cell.row, cell.col), source = sheet.cells[key]
    sheet.cells[key] = { ...(source ?? { value: '' }), fontFamily: cell.style.font.family, fontSize: cell.style.font.size,
      bold: cell.style.font.bold, italic: cell.style.font.italic, underline: cell.style.font.underline, strike: cell.style.font.strike,
      color: cell.style.font.color, background: cell.style.fill, authoredStyle: cloneModel(cell.authoredStyle ?? cell.style),
      authoredStyleReference: cell.authoredStyleReference ?? null, clearedStyleProperties: cell.clearedStyleProperties ?? null,
      richText: cell.richText === undefined && source?.richText != null ? null : cell.richText ?? source?.richText }
  }
  return true
}
