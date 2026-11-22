import { createCell, defaultCellStyle, type CellRange, type SpreadsheetSheet } from '../../domain/model.ts'
import { cellKey, parseCellKey, type ExcelCellView, type ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { sheetCellShrinkToFit, sheetCellTextRotation, sheetCellWrapStrategy, type SheetCellLayoutStyle, type SheetHorizontalAlignment,
  type SheetTextRotation, type SheetVerticalAlignment, type SheetWrapStrategy } from './sheet-cell-layout.ts'
export { sheetCellShrinkToFit, sheetCellTextRotation, sheetCellWrapStrategy } from './sheet-cell-layout.ts'
export type { SheetHorizontalAlignment, SheetVerticalAlignment, SheetWrapStrategy, SheetTextRotation } from './sheet-cell-layout.ts'

type LayoutStyle = SheetCellLayoutStyle
type LayoutPatch = LayoutStyle

export const SHEET_LAYOUT_FORMAT_COMMANDS = Object.freeze({
  horizontal: 'sheet.command.set-align', vertical: 'sheet.command.set-vertical-align',
  wrapStrategy: 'sheet.command.set-wrap-strategy', wrap: 'sheet.command.set-wrap', toggleWrap: 'sheet.command.toggle-wrap',
  shrink: 'sheet.command.set-shrink-to-fit', rotation: 'sheet.command.set-text-rotation',
})
export const SHEET_LAYOUT_FORMAT_ALIASES = Object.freeze({
  horizontal: 'sheet.command.set-horizontal-text-align', vertical: 'sheet.command.set-vertical-text-align',
  wrap: 'sheet.command.set-text-wrap', horizontalAlignment: 'sheet.command.set-horizontal-alignment',
  verticalAlignment: 'sheet.command.set-vertical-alignment',
})
export const SHEET_LAYOUT_FORMAT_COMMAND_IDS: readonly string[] = [...Object.values(SHEET_LAYOUT_FORMAT_COMMANDS), ...Object.values(SHEET_LAYOUT_FORMAT_ALIASES)]
const horizontalIds: readonly string[] = [SHEET_LAYOUT_FORMAT_COMMANDS.horizontal, SHEET_LAYOUT_FORMAT_ALIASES.horizontal, SHEET_LAYOUT_FORMAT_ALIASES.horizontalAlignment]
const verticalIds: readonly string[] = [SHEET_LAYOUT_FORMAT_COMMANDS.vertical, SHEET_LAYOUT_FORMAT_ALIASES.vertical, SHEET_LAYOUT_FORMAT_ALIASES.verticalAlignment]
const defaults = defaultCellStyle()

export interface SheetLayoutFormatCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  ranges?: CellRange[]
  range?: CellRange | { startRow: number; startColumn: number; endRow: number; endColumn: number }
  value?: unknown
}
export type SheetLayoutFormatAction = { kind: 'horizontal'; value: SheetHorizontalAlignment }
  | { kind: 'vertical'; value: SheetVerticalAlignment }
  | { kind: 'wrap'; value: SheetWrapStrategy }
  | { kind: 'toggle-wrap' }
  | { kind: 'shrink'; value?: boolean }
  | { kind: 'rotation'; value: SheetTextRotation }

function commandValue(params: unknown): unknown {
  return params && typeof params === 'object' ? 'value' in params ? params.value : undefined : params
}
function booleanValue(value: unknown): boolean {
  if (typeof value === 'boolean' || value === 0 || value === 1) return value === true || value === 1
  throw new TypeError('单元格布局开关必须是布尔值。')
}

export function layoutFormatActionForCommand(id: string, params?: unknown): SheetLayoutFormatAction | undefined {
  const value = commandValue(params)
  if (horizontalIds.includes(id)) {
    if (value === undefined) return undefined
    const alignment = typeof value === 'number' ? (['general', 'left', 'center', 'right', 'justify', 'justify'] as const)[value] : value
    if (alignment === 'general' || alignment === 'left' || alignment === 'center' || alignment === 'right' || alignment === 'justify') return { kind: 'horizontal', value: alignment }
    throw new TypeError('水平对齐参数无效。')
  }
  if (verticalIds.includes(id)) {
    if (value === undefined) return undefined
    const alignment = typeof value === 'number' ? ([undefined, 'top', 'middle', 'bottom'] as const)[value] : value
    if (alignment === 'top' || alignment === 'middle' || alignment === 'bottom') return { kind: 'vertical', value: alignment }
    throw new TypeError('垂直对齐参数无效。')
  }
  if (id === SHEET_LAYOUT_FORMAT_COMMANDS.wrap) {
    if (value === undefined) return undefined
    return { kind: 'wrap', value: booleanValue(value) ? 'wrap' : 'overflow' }
  }
  if (id === SHEET_LAYOUT_FORMAT_COMMANDS.toggleWrap) return { kind: 'toggle-wrap' }
  if (id === SHEET_LAYOUT_FORMAT_COMMANDS.wrapStrategy || id === SHEET_LAYOUT_FORMAT_ALIASES.wrap) {
    if (value === undefined) return undefined
    const strategy = typeof value === 'number' ? ([undefined, 'overflow', 'clip', 'wrap'] as const)[value] : value
    if (strategy === 'overflow' || strategy === 'clip' || strategy === 'wrap') return { kind: 'wrap', value: strategy }
    throw new TypeError('文本换行策略无效。')
  }
  if (id === SHEET_LAYOUT_FORMAT_COMMANDS.shrink) return value === undefined ? { kind: 'shrink' } : { kind: 'shrink', value: booleanValue(value) }
  if (id === SHEET_LAYOUT_FORMAT_COMMANDS.rotation) {
    if (value === undefined) return undefined
    if (value === 'v' || typeof value === 'number' && Number.isFinite(value) && value >= -90 && value <= 90) return { kind: 'rotation', value }
    throw new RangeError('文字方向必须是 -90 到 90 度或竖排。')
  }
  return undefined
}

export function sheetLayoutFormatInputRanges(input?: SheetLayoutFormatCommandInput): CellRange[] | undefined {
  if (input?.ranges !== undefined) return input.ranges
  const range = input?.range
  return range ? ['startCol' in range ? { ...range } : { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }] : undefined
}
export function assertSheetLayoutFormatRanges(ranges: readonly CellRange[]): void {
  for (const range of ranges) {
    if (![range.startRow, range.startCol, range.endRow, range.endCol].every(Number.isSafeInteger)
      || range.startRow < 0 || range.startCol < 0 || range.endRow < range.startRow || range.endCol < range.startCol
      || range.endRow >= 1_048_576 || range.endCol >= 16_384) throw new RangeError('单元格布局范围无效。')
  }
}
function assertEditable(sheet: SpreadsheetSheet | ExcelSheetView, ranges: readonly CellRange[], userId: string): void {
  assertSheetLayoutFormatRanges(ranges)
  if (!canEditWorksheetProtection(sheet, userId) || !canEditProtectedRanges(sheet, [...ranges], userId)) throw new Error('当前选区受保护，不能修改单元格布局。')
}
function* positions(ranges: readonly CellRange[], filteredRows: ReadonlySet<number>): Generator<{ row: number; col: number }> {
  for (let index = 0; index < ranges.length; index++) {
    const range = ranges[index]!
    for (let row = range.startRow; row <= range.endRow; row++) {
      if (filteredRows.has(row)) continue
      for (let col = range.startCol; col <= range.endCol; col++) {
        // Overlapping selections are one formatting transaction and each position is visited once.
        let visited = false
        for (let earlier = 0; earlier < index; earlier++) {
          const previous = ranges[earlier]!
          if (row >= previous.startRow && row <= previous.endRow && col >= previous.startCol && col <= previous.endCol) { visited = true; break }
        }
        if (visited) continue
        yield { row, col }
      }
    }
  }
}
function resolvedAction(action: SheetLayoutFormatAction, ranges: readonly CellRange[], filteredRows: ReadonlySet<number>, read: (row: number, col: number) => LayoutStyle | undefined): Exclude<SheetLayoutFormatAction, { kind: 'toggle-wrap' }> {
  if (action.kind === 'toggle-wrap') {
    let allWrapped = true
    for (const { row, col } of positions(ranges, filteredRows)) if (sheetCellWrapStrategy(read(row, col)) !== 'wrap') { allWrapped = false; break }
    return { kind: 'wrap', value: allWrapped ? 'overflow' : 'wrap' }
  }
  if (action.kind === 'shrink' && action.value === undefined) {
    const primary = ranges[0]!
    return { kind: 'shrink', value: read(primary.startRow, primary.startCol)?.shrinkToFit !== true }
  }
  return action
}

/** Preserve the other layout fields while maintaining the fixed-version wrap/shrink invariant. */
function layoutPatch(style: LayoutStyle | undefined, action: Exclude<SheetLayoutFormatAction, { kind: 'toggle-wrap' }>): LayoutPatch | undefined {
  let patch: LayoutPatch
  switch (action.kind) {
    case 'horizontal': patch = { horizontal: action.value }; break
    case 'vertical': patch = { vertical: action.value }; break
    case 'wrap': patch = { wrap: action.value === 'wrap', wrapStrategy: action.value, ...(action.value === 'wrap' ? { shrinkToFit: false } : {}) }; break
    case 'shrink': patch = { shrinkToFit: action.value === true, ...(action.value && sheetCellWrapStrategy(style) === 'wrap' ? { wrap: false, wrapStrategy: 'overflow' } : {}) }; break
    case 'rotation': patch = { rotation: action.value === 'v' ? 0 : action.value, rotationVertical: action.value === 'v' }; break
  }
  const current: Required<LayoutStyle> = { horizontal: style?.horizontal ?? defaults.horizontal, vertical: style?.vertical ?? defaults.vertical,
    wrap: style?.wrap ?? sheetCellWrapStrategy(style) === 'wrap', wrapStrategy: sheetCellWrapStrategy(style),
    shrinkToFit: style?.shrinkToFit ?? false, rotation: style?.rotation ?? 0, rotationVertical: style?.rotationVertical ?? false }
  return (Object.keys(patch) as Array<keyof LayoutPatch>).some(key => patch[key] !== current[key]) ? patch : undefined
}
function viewLayout(cell: ExcelCellView | undefined): LayoutStyle | undefined {
  return cell ? { horizontal: cell.align, vertical: cell.vertical, wrap: cell.wrap, wrapStrategy: cell.wrapStrategy,
    shrinkToFit: cell.shrinkToFit, rotation: cell.rotation, rotationVertical: cell.rotationVertical } : undefined
}

export interface SheetLayoutFormatState {
  align?: SheetHorizontalAlignment
  vertical?: SheetVerticalAlignment
  wrap?: boolean
  wrapStrategy?: SheetWrapStrategy
  shrinkToFit?: boolean
  rotation: number | null
  rotationVertical?: boolean
}

/** Count the union by rectangular bands, without materializing empty cells in large selections. */
function scopeSize(ranges: readonly CellRange[], filteredRows: ReadonlySet<number>): number {
  const boundaries = [...new Set(ranges.flatMap(range => [range.startRow, range.endRow + 1]))].sort((a, b) => a - b)
  let count = 0
  for (let index = 0; index < boundaries.length - 1; index++) {
    const start = boundaries[index]!, end = boundaries[index + 1]!
    const intervals = ranges.filter(range => range.startRow <= start && range.endRow >= start)
      .map(range => [range.startCol, range.endCol + 1] as const).sort((a, b) => a[0] - b[0])
    let width = 0, right = -1
    for (const interval of intervals) {
      width += Math.max(0, interval[1] - Math.max(right, interval[0]))
      right = Math.max(right, interval[1])
    }
    let rows = end - start
    for (const row of filteredRows) if (row >= start && row < end) rows--
    count += width * rows
  }
  return count
}

export function getSheetViewLayoutFormatState(sheet: ExcelSheetView | undefined, ranges: readonly CellRange[]): SheetLayoutFormatState {
  const filtered = new Set(sheet?.filteredRows ?? sheet?.autoFilter?.filteredRows ?? [])
  const cells: Array<ExcelCellView | undefined> = []
  for (const [key, cell] of Object.entries(sheet?.cells ?? {})) {
    const position = parseCellKey(key)
    if (!position || filtered.has(position.row) || !ranges.some(range => position.row >= range.startRow && position.row <= range.endRow && position.col >= range.startCol && position.col <= range.endCol)) continue
    cells.push(cell)
  }
  if (!cells.length || scopeSize(ranges, filtered) > cells.length) cells.push(undefined)
  function common<T>(read: (cell: ExcelCellView | undefined) => T): T | undefined {
    const first = read(cells[0])
    return cells.every(cell => Object.is(read(cell), first)) ? first : undefined
  }
  const rotation = common(sheetCellTextRotation)
  return { align: common(cell => cell?.align ?? defaults.horizontal), vertical: common(cell => cell?.vertical ?? defaults.vertical),
    wrapStrategy: common(sheetCellWrapStrategy), wrap: common(cell => sheetCellWrapStrategy(cell) === 'wrap'),
    shrinkToFit: common(sheetCellShrinkToFit), rotation: rotation === 'v' ? 0 : rotation ?? null,
    rotationVertical: rotation === undefined ? undefined : rotation === 'v' }
}

export function applySheetLayoutFormatCommand(sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetLayoutFormatAction, userId = 'local-user', filteredRows = sheet.autoFilter?.filteredRows ?? []): boolean {
  if (!ranges.length) return false
  assertEditable(sheet, ranges, userId)
  const filtered = new Set(filteredRows)
  const resolved = resolvedAction(action, ranges, filtered, (row, col) => sheet.cells[`${row}:${col}`]?.style)
  let changed = false
  for (const { row, col } of positions(ranges, filtered)) {
    const key = `${row}:${col}`, previous = sheet.cells[key], patch = layoutPatch(previous?.style, resolved)
    if (!patch) continue
    const cell = previous ?? createCell(row, col)
    sheet.cells[key] = { ...cell, style: { ...cell.style, ...patch } }
    changed = true
  }
  return changed
}
export function applySheetViewLayoutFormatCommand(sheet: ExcelSheetView, ranges: readonly CellRange[], action: SheetLayoutFormatAction, userId = 'local-user'): boolean {
  if (!ranges.length) return false
  assertEditable(sheet, ranges, userId)
  const filtered = new Set(sheet.filteredRows ?? sheet.autoFilter?.filteredRows ?? [])
  const resolved = resolvedAction(action, ranges, filtered, (row, col) => viewLayout(sheet.cells[cellKey(row, col)]))
  let changed = false
  for (const { row, col } of positions(ranges, filtered)) {
    const key = cellKey(row, col), cell = sheet.cells[key], patch = layoutPatch(viewLayout(cell), resolved)
    if (!patch) continue
    const { horizontal, ...rest } = patch
    sheet.cells[key] = { ...(cell ?? { value: '' }), ...rest, ...(horizontal === undefined ? {} : { align: horizontal }) }
    changed = true
  }
  return changed
}
