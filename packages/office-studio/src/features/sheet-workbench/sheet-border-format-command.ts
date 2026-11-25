import { createCell, type CellRange, type SpreadsheetSheet } from '../../domain/model.ts'
import { parseCellKey, setSheetViewCellStyleProperties, sheetViewCellAt, type ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { assertSheetBorderRanges, createSheetBorderPlan, normalizeSheetBorderColor, normalizeSheetBorderPreset, normalizeSheetBorderStyle,
  type SheetBorderLineStyle, type SheetBorderOptions, type SheetBorderPlan, type SheetBorderPosition, type SheetBorderPreset, type SheetBorderSource } from './sheet-cell-border.ts'
import { resolveSheetCellStyle, setSheetCellStyleProperties, type SheetStyleContext, type SheetStyleSheet } from './sheet-style-inheritance.ts'
import { captureSheetStyleRegion } from './sheet-style-region.ts'

/** A temporary synchronous budget. Larger inherited-border clears still need the Worker protocol. */
export const SHEET_INHERITED_BORDER_SYNC_CELL_BUDGET = 100_000

export const SHEET_BORDER_FORMAT_COMMANDS = Object.freeze({
  set: 'sheet.command.set-border', basic: 'sheet.command.set-border-basic', position: 'sheet.command.set-border-position',
  color: 'sheet.command.set-border-color', style: 'sheet.command.set-border-style', legacy: 'sheet.command.set-borders',
})
export const SHEET_BORDER_FORMAT_COMMAND_IDS: readonly string[] = Object.values(SHEET_BORDER_FORMAT_COMMANDS)
export interface SheetBorderFormatState extends SheetBorderOptions { type: SheetBorderPreset; active: boolean }
export interface SheetBorderFormatAction { state: SheetBorderFormatState; apply: boolean }
export type SheetBorderRangeInput = CellRange | { startRow: number; startColumn: number; endRow: number; endColumn: number }
export interface SheetBorderFormatCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  ranges?: SheetBorderRangeInput[]
  range?: SheetBorderRangeInput
  value?: unknown
  /** UI color/style choices explicitly select the current preset before applying its pen. */
  activate?: boolean
}

export function createSheetBorderFormatState(): SheetBorderFormatState { return { type: 'all', style: 'thin', color: '#000000', active: false } }
function record(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : undefined
}
function commandValue(params: unknown): unknown { const input = record(params); return input ? input.value : params }
export function sheetBorderFormatCommandInput(params: unknown): SheetBorderFormatCommandInput {
  return record(params) ? params as SheetBorderFormatCommandInput : params === undefined ? {} : { value: params }
}

/** State is staged separately so canceled, invalid or protected requests never remember a new pen. */
export function borderFormatActionForCommand(
  id: string, params?: unknown, previous: SheetBorderFormatState = createSheetBorderFormatState(),
): SheetBorderFormatAction | undefined {
  if (!SHEET_BORDER_FORMAT_COMMAND_IDS.includes(id)) return undefined
  const state = { ...previous }, value = commandValue(params)
  const inputContext = record(params)
  if (inputContext?.activate !== undefined && typeof inputContext.activate !== 'boolean') throw new TypeError('边框 activate 必须是布尔值。')
  if (inputContext?.activate === true) state.active = true
  if (id === SHEET_BORDER_FORMAT_COMMANDS.basic) {
    const input = record(value)
    if (!input || input.type === undefined) throw new TypeError('边框设置必须包含有效的 type。')
    state.type = normalizeSheetBorderPreset(input.type)
    if (input.style !== undefined) state.style = normalizeSheetBorderStyle(input.style)
    if (input.color !== undefined) state.color = normalizeSheetBorderColor(input.color)
    state.active = true
  } else if (id === SHEET_BORDER_FORMAT_COMMANDS.legacy || id === SHEET_BORDER_FORMAT_COMMANDS.position) {
    state.type = normalizeSheetBorderPreset(value === undefined && id === SHEET_BORDER_FORMAT_COMMANDS.legacy ? state.type : value)
    state.active = true
  } else if (id === SHEET_BORDER_FORMAT_COMMANDS.style) state.style = normalizeSheetBorderStyle(value)
  else if (id === SHEET_BORDER_FORMAT_COMMANDS.color) state.color = normalizeSheetBorderColor(value)
  return { state, apply: state.active }
}

export function sheetBorderFormatInputRanges(input?: SheetBorderFormatCommandInput): CellRange[] | undefined {
  const ranges = input?.ranges !== undefined ? input.ranges : input?.range !== undefined ? [input.range] : undefined
  if (ranges === undefined) return undefined
  if (!Array.isArray(ranges)) throw new TypeError('边框 ranges 必须是范围数组。')
  const normalized = ranges.map(range => {
    if (!range || typeof range !== 'object') throw new TypeError('边框范围必须包含行列坐标。')
    return 'startCol' in range ? { ...range } : { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }
  })
  assertSheetBorderRanges(normalized)
  return normalized
}

function assertEditable(sheet: SpreadsheetSheet | ExcelSheetView, ranges: readonly CellRange[], userId: string): void {
  assertSheetBorderRanges(ranges)
  if (!canEditWorksheetProtection(sheet, userId) || !canEditProtectedRanges(sheet, ranges, userId)) throw new Error('边框选区或相邻边段受保护，不能修改边框。')
}
function assertPlanEditable(sheet: SpreadsheetSheet | ExcelSheetView, plan: SheetBorderPlan, userId: string): void {
  assertEditable(sheet, plan.ranges, userId)
  if (!canEditProtectedRanges(sheet, plan.patches.map(({ row, col }) => ({ startRow: row, endRow: row, startCol: col, endCol: col })), userId)) {
    throw new Error('相邻单元格的边框受保护，操作未应用。')
  }
}
function* inheritedBorderPositions(sheet: SheetStyleSheet, context: SheetStyleContext, ranges: readonly CellRange[],
  existing: Iterable<SheetBorderPosition>, readBorder: SheetBorderSource['readBorder']): Generator<SheetBorderPosition> {
  const visited = new Set<number>()
  for (const position of existing) { visited.add(position.row * 16_384 + position.col); yield position }
  let projected = 0
  for (const range of ranges) {
    const area = { startRow: Math.max(0, range.startRow - 1), endRow: Math.min(1_048_575, range.endRow + 1),
      startCol: Math.max(0, range.startCol - 1), endCol: Math.min(16_383, range.endCol + 1) }
    const source = captureSheetStyleRegion(sheet, { start: area.startRow, end: area.endRow }, { start: area.startCol, end: area.endCol }, context)
    const layers = [source.defaultStyle, ...Object.values(source.rowStyles ?? {}), ...Object.values(source.columnStyles ?? {}),
      ...(source.rowStyleSpans ?? []).map(span => span.value), ...(source.columnStyleSpans ?? []).map(span => span.value)]
    if (!layers.some(style => style?.border && Object.keys(style.border).length)) continue
    projected += (area.endRow - area.startRow + 1) * (area.endCol - area.startCol + 1)
    if (projected > SHEET_INHERITED_BORDER_SYNC_CELL_BUDGET) throw new RangeError('继承边框清除超出当前同步处理预算，需使用后续 Worker 路径。')
    for (let row = area.startRow; row <= area.endRow; row++) for (let col = area.startCol; col <= area.endCol; col++) {
      const key = row * 16_384 + col
      if (visited.has(key) || !Object.keys(readBorder(row, col) ?? {}).length) continue
      visited.add(key); yield { row, col }
    }
  }
}
function modelSource(sheet: SpreadsheetSheet, context: SheetStyleContext): SheetBorderSource {
  const readBorder: SheetBorderSource['readBorder'] = (row, col) => resolveSheetCellStyle(sheet, row, col, sheet.cells[`${row}:${col}`], context).border
  return { readBorder,
    existingCells: ranges => inheritedBorderPositions(sheet, context, ranges, Object.values(sheet.cells), readBorder), mergedRanges: sheet.mergedRanges }
}
function viewSource(sheet: ExcelSheetView): SheetBorderSource {
  const mergedRanges = (sheet.merges ?? []).map(merge => {
    const [start, end] = merge.split(':'), first = parseCellKey(start ?? ''), last = parseCellKey(end ?? start ?? '')
    if (!first || !last) throw new RangeError('合并单元格范围无效。')
    return { startRow: first.row, startCol: first.col, endRow: last.row, endCol: last.col }
  })
  const readBorder: SheetBorderSource['readBorder'] = (row, col) => sheetViewCellAt(sheet, row, col).border
  return { readBorder, existingCells: ranges => inheritedBorderPositions(sheet, { defaultStyle: sheet.workbookDefaultStyle, styles: sheet.styles }, ranges,
    Object.keys(sheet.cells).flatMap(key => { const position = parseCellKey(key); return position ? [position] : [] }), readBorder), mergedRanges }
}

export function applySheetBorderFormatCommand(
  sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetBorderFormatAction, userId = 'local-user', context: SheetStyleContext = {},
): boolean {
  if (!ranges.length) return false
  assertEditable(sheet, ranges, userId)
  if (!action.apply) return false
  const plan = createSheetBorderPlan(modelSource(sheet, context), ranges, action.state.type, action.state)
  assertPlanEditable(sheet, plan, userId)
  for (const patch of plan.patches) {
    const key = `${patch.row}:${patch.col}`, cell = sheet.cells[key] ?? createCell(patch.row, patch.col)
    cell.style = resolveSheetCellStyle(sheet, patch.row, patch.col, sheet.cells[key], context)
    setSheetCellStyleProperties(cell, { border: patch.border })
    sheet.cells[key] = cell
  }
  return plan.patches.length > 0
}
export function applySheetViewBorderFormatCommand(
  sheet: ExcelSheetView, ranges: readonly CellRange[], action: SheetBorderFormatAction, userId = 'local-user',
): boolean {
  if (!ranges.length) return false
  assertEditable(sheet, ranges, userId)
  if (!action.apply) return false
  const plan = createSheetBorderPlan(viewSource(sheet), ranges, action.state.type, action.state)
  assertPlanEditable(sheet, plan, userId)
  for (const patch of plan.patches) {
    setSheetViewCellStyleProperties(sheet, patch.row, patch.col, { border: patch.border })
  }
  return plan.patches.length > 0
}

export function sheetBorderFormatOptions(state: SheetBorderFormatState): { style: SheetBorderLineStyle; color: string } {
  return { style: state.style, color: state.color }
}
