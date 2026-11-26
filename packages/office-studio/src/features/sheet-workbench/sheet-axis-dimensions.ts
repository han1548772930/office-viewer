import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { assignSheetAxisSpan, canonicalSheetAxisSpans, sheetAxisSpanAt, sheetColumnWidth, sheetManualRowHeight, sheetRowHeight, sheetRowIsAutoHeight,
  type SheetAxisDimensions, type SheetAxisSpan } from '../../domain/sheet-axis.ts'
import { cellKey, parseCellKey, toExcelCellView, type ExcelCellView, type ExcelSheetView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { assertSheetCommandRanges, normalizeSheetCommandRange, type SheetRangeCommandInput } from './sheet-range-command.ts'
import { calculateSheetAxisAutoFit, type SheetAutoFitOptions } from './sheet-auto-fit.ts'
import type { SheetTextMeasure } from './sheet-cell-text-layout.ts'
import { hiddenSheetRows, hiddenSheetColumns, isSheetColumnVisible } from './sheet-row-visibility.ts'

export const SHEET_AXIS_DIMENSION_COMMANDS = {
  rowHeight: 'sheet.command.set-row-height', columnWidth: 'sheet.command.set-worksheet-col-width',
  rowAutoHeight: 'sheet.command.set-row-is-auto-height', columnAutoWidth: 'sheet.command.set-col-auto-width',
  deltaRowHeight: 'sheet.command.delta-row-height', deltaColumnWidth: 'sheet.command.delta-column-width',
  growingRowHeight: 'sheet.command.set-growing-row-height',
} as const
export const SHEET_AXIS_DIMENSION_COMMAND_IDS = [...Object.values(SHEET_AXIS_DIMENSION_COMMANDS),
  'sheet.command.set-col-width', 'sheet.command.auto-fit-row', 'sheet.command.auto-fit-column'] as const
export interface SheetAxisDimensionCommandInput extends SheetRangeCommandInput {
  value?: number
  index?: number
  anchorRow?: number
  anchorCol?: number
  deltaY?: number
  deltaX?: number
  expectedRevision?: number
}
export type SheetAxisDimensionAction = { axis: 'row' | 'column'; kind: 'set' | 'delta' | 'auto' | 'grow' }
export interface SheetAxisDimensionOptions { measure: SheetTextMeasure; region?: string }

function isRecord(value: unknown): value is Record<string, unknown> { return value !== null && typeof value === 'object' && !Array.isArray(value) }
export function sheetAxisDimensionCommandInput(params?: unknown): SheetAxisDimensionCommandInput {
  if (params === undefined) return {}
  if (typeof params === 'number') return { value: params }
  if (!isRecord(params)) throw new TypeError('行列尺寸命令参数无效。')
  const input: SheetAxisDimensionCommandInput = {}
  for (const key of ['unitId', 'sheetId', 'subUnitId'] as const) {
    if (params[key] === undefined) continue
    const value = params[key]
    if (typeof value !== 'string' || !value.trim()) throw new TypeError('工作簿或工作表标识无效。')
    input[key] = value
  }
  for (const key of ['value', 'index', 'anchorRow', 'anchorCol', 'deltaY', 'deltaX', 'expectedRevision'] as const) {
    const value = params[key]
    if (value === undefined) continue
    if (typeof value !== 'number' || !Number.isFinite(value)) throw new TypeError('行列尺寸参数必须是有限数值。')
    input[key] = value
  }
  for (const key of ['ranges', 'selections'] as const) {
    const ranges = params[key]
    if (ranges === undefined) continue
    if (!Array.isArray(ranges)) throw new TypeError('行列目标范围必须是数组。')
    input[key] = ranges.map(range => {
      if (!isRecord(range)) throw new TypeError('行列目标范围无效。')
      const startRow = range.startRow, endRow = range.endRow, startCol = range.startCol ?? range.startColumn, endCol = range.endCol ?? range.endColumn
      if (typeof startRow !== 'number' || typeof endRow !== 'number' || typeof startCol !== 'number' || typeof endCol !== 'number') throw new TypeError('行列目标坐标无效。')
      return normalizeSheetCommandRange({ startRow, endRow, startCol, endCol })
    })
  }
  if (params.range !== undefined) {
    const source = sheetAxisDimensionCommandInput({ ranges: [params.range] })
    input.range = source.ranges![0]!
  }
  return input
}
export function sheetAxisDimensionActionForCommand(id: string): SheetAxisDimensionAction | undefined {
  if (id === SHEET_AXIS_DIMENSION_COMMANDS.rowHeight) return { axis: 'row', kind: 'set' }
  if (id === SHEET_AXIS_DIMENSION_COMMANDS.columnWidth || id === 'sheet.command.set-col-width') return { axis: 'column', kind: 'set' }
  if (id === SHEET_AXIS_DIMENSION_COMMANDS.rowAutoHeight || id === 'sheet.command.auto-fit-row') return { axis: 'row', kind: 'auto' }
  if (id === SHEET_AXIS_DIMENSION_COMMANDS.columnAutoWidth || id === 'sheet.command.auto-fit-column') return { axis: 'column', kind: 'auto' }
  if (id === SHEET_AXIS_DIMENSION_COMMANDS.deltaRowHeight) return { axis: 'row', kind: 'delta' }
  if (id === SHEET_AXIS_DIMENSION_COMMANDS.deltaColumnWidth) return { axis: 'column', kind: 'delta' }
  if (id === SHEET_AXIS_DIMENSION_COMMANDS.growingRowHeight) return { axis: 'row', kind: 'grow' }
  return undefined
}

/** Delta/axis-header gestures use every selected whole row/column only when the anchor belongs to that selection. */
export function sheetAxisDimensionTargetRanges(sheet: { rowCount?: number; columnCount?: number }, axis: 'row' | 'column', ranges: readonly CellRange[], anchor?: number): CellRange[] {
  assertSheetCommandRanges(ranges)
  const rows = sheet.rowCount ?? 1_048_576, columns = sheet.columnCount ?? 16_384, count = axis === 'row' ? rows : columns
  if (anchor !== undefined && (!Number.isSafeInteger(anchor) || anchor < 0 || anchor >= count)) throw new RangeError('行列尺寸目标超出工作表。')
  const selected = ranges.map(range => ({ startRow: range.startRow, endRow: Math.min(rows - 1, range.endRow),
    startCol: range.startCol, endCol: Math.min(columns - 1, range.endCol) }))
    .filter(range => range.startRow <= range.endRow && range.startCol <= range.endCol)
  if (anchor === undefined) return selected
  const all = selected.length === 1 && selected[0]!.startRow === 0 && selected[0]!.endRow === rows - 1 && selected[0]!.startCol === 0 && selected[0]!.endCol === columns - 1
  const whole = selected.filter(range => axis === 'row' ? range.startCol === 0 && range.endCol === columns - 1 : range.startRow === 0 && range.endRow === rows - 1)
  const anchored = whole.some(range => axis === 'row' ? range.startRow <= anchor && anchor <= range.endRow : range.startCol <= anchor && anchor <= range.endCol)
  if (all || anchored) return all ? selected : whole
  return axis === 'row' ? [{ startRow: anchor, endRow: anchor, startCol: 0, endCol: columns - 1 }]
    : [{ startRow: 0, endRow: rows - 1, startCol: anchor, endCol: anchor }]
}
function axisSpans(ranges: readonly CellRange[], axis: 'row' | 'column'): SheetAxisSpan<boolean>[] {
  return canonicalSheetAxisSpans(ranges.map(range => ({ start: axis === 'row' ? range.startRow : range.startCol,
    end: axis === 'row' ? range.endRow : range.endCol, value: true })))
}
function removeSparseSizes(values: Record<number, number> | undefined, spans: readonly SheetAxisSpan<boolean>[]): void {
  for (const key of Object.keys(values ?? {})) if (sheetAxisSpanAt(spans, Number(key))) delete values![Number(key)]
}
function dimensionSignature(sheet: SheetAxisDimensions): string {
  return JSON.stringify([sheet.rowHeights ?? {}, sheet.columnWidths ?? {}, sheet.rowHeightSpans ?? [], sheet.columnWidthSpans ?? [], sheet.rowAutoHeightSpans ?? [], sheet.rowAutoHeights ?? {}])
}

interface MergeNode { center: number; crossing: Array<{ range: CellRange; order: number }>; disjoint: boolean; left?: MergeNode; right?: MergeNode }
function mergeIndex(entries: MergeNode['crossing']): MergeNode | undefined {
  if (!entries.length) return undefined
  const center = [...entries].sort((a, b) => a.range.startRow - b.range.startRow)[Math.floor(entries.length / 2)]!.range.startRow
  const left: typeof entries = [], right: typeof entries = [], crossing: typeof entries = []
  for (const entry of entries) (entry.range.endRow < center ? left : entry.range.startRow > center ? right : crossing).push(entry)
  crossing.sort((a, b) => a.range.startCol - b.range.startCol || a.order - b.order)
  return { center, crossing, disjoint: crossing.every((entry, i) => !i || crossing[i - 1]!.range.endCol < entry.range.startCol), left: mergeIndex(left), right: mergeIndex(right) }
}
function indexedMergeAt(node: MergeNode | undefined, row: number, column: number): { range: CellRange; order: number } | undefined {
  if (!node) return undefined
  let match: { range: CellRange; order: number } | undefined
  if (node.disjoint) {
    let low = 0, high = node.crossing.length
    while (low < high) { const middle = (low + high) >>> 1; if (node.crossing[middle]!.range.startCol <= column) low = middle + 1; else high = middle }
    const candidate = node.crossing[low - 1]
    if (candidate && candidate.range.endCol >= column && row >= candidate.range.startRow && row <= candidate.range.endRow) match = candidate
  } else for (const entry of node.crossing) {
    if (entry.range.startCol > column) break
    if (column <= entry.range.endCol && row >= entry.range.startRow && row <= entry.range.endRow && (!match || entry.order < match.order)) match = entry
  }
  const nested = indexedMergeAt(row < node.center ? node.left : row > node.center ? node.right : undefined, row, column)
  return nested && (!match || nested.order < match.order) ? nested : match
}
function measurementContext(sheet: ExcelSheetView, options: SheetAxisDimensionOptions): SheetAutoFitOptions {
  const merges = (sheet.merges ?? []).flatMap((value, order) => {
    const [start, end] = value.split(':'), first = parseCellKey(start ?? ''), last = parseCellKey(end ?? start ?? '')
    return first && last ? [{ range: { startRow: first.row, endRow: last.row, startCol: first.col, endCol: last.col }, order }] : []
  })
  const index = mergeIndex(merges), hiddenColumns = hiddenSheetColumns(sheet)
  return { ...options, hiddenRows: hiddenSheetRows(sheet), hiddenColumns,
    columnWidth: column => isSheetColumnVisible(sheet, column, hiddenColumns) ? sheetColumnWidth(sheet, column) : 0,
    mergeAt: (row, column) => {
      const range = indexedMergeAt(index, row, column)?.range
      return range ? { startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol } : undefined
    } }
}
function groupedCells(sheet: ExcelSheetView, axis: 'row' | 'column', targets?: readonly SheetAxisSpan<boolean>[]): Map<number, Array<[string, ExcelCellView]>> {
  const groups = new Map<number, Array<[string, ExcelCellView]>>()
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position || position.row < 0 || position.col < 0 || position.row >= sheet.rowCount || position.col >= sheet.columnCount) continue
    const index = axis === 'row' ? position.row : position.col
    if (targets && !sheetAxisSpanAt(targets, index)) continue
    const group = groups.get(index) ?? []; group.push([key, cell]); groups.set(index, group)
  }
  return groups
}
function setRowMode(sheet: ExcelSheetView, spans: readonly SheetAxisSpan<boolean>[], automatic: boolean): void {
  for (const span of spans) sheet.rowAutoHeightSpans = assignSheetAxisSpan(sheet.rowAutoHeightSpans, span.start, span.end, automatic)
  if (automatic) {
    removeSparseSizes(sheet.rowHeights, spans)
    for (const span of spans) sheet.rowHeightSpans = assignSheetAxisSpan(sheet.rowHeightSpans, span.start, span.end)
  } else removeSparseSizes(sheet.rowAutoHeights, spans)
}
function hasAxisSize(sheet: ExcelSheetView, axis: 'row' | 'column', target: SheetAxisSpan<boolean>, value: number): boolean {
  const boundaries = new Set([target.start])
  const add = (index: number) => { if (index >= target.start && index <= target.end) boundaries.add(index) }
  for (const span of axis === 'row' ? [...(sheet.rowHeightSpans ?? []), ...(sheet.rowAutoHeightSpans ?? [])] : sheet.columnWidthSpans ?? []) {
    add(span.start); add(span.end + 1)
  }
  for (const key of Object.keys(axis === 'row' ? sheet.rowHeights ?? {} : sheet.columnWidths ?? {})) { add(Number(key)); add(Number(key) + 1) }
  if (axis === 'column') add(1) // The standalone first-column fallback differs from other default columns.
  for (const index of boundaries) {
    if (axis === 'row') {
      if (sheetRowIsAutoHeight(sheet, index) || (sheetManualRowHeight(sheet, index) ?? sheet.defaultRowHeight ?? 26) !== value) return false
    } else if (sheetColumnWidth(sheet, index) !== value) return false
  }
  return true
}
function setSize(sheet: ExcelSheetView, axis: 'row' | 'column', spans: readonly SheetAxisSpan<boolean>[], value: number): void {
  const bounded = axis === 'row' ? Math.min(2000, value) : value
  const changed = spans.filter(span => !hasAxisSize(sheet, axis, span, bounded))
  if (!changed.length) return
  removeSparseSizes(axis === 'row' ? sheet.rowHeights : sheet.columnWidths, changed)
  for (const span of changed) {
    if (axis === 'row') sheet.rowHeightSpans = assignSheetAxisSpan(sheet.rowHeightSpans, span.start, span.end, bounded)
    else sheet.columnWidthSpans = assignSheetAxisSpan(sheet.columnWidthSpans, span.start, span.end, bounded)
  }
  if (axis === 'row') setRowMode(sheet, changed, false)
}

export function applySheetViewAxisDimensionCommand(sheet: ExcelSheetView, ranges: readonly CellRange[], action: SheetAxisDimensionAction,
  input: SheetAxisDimensionCommandInput, userId: string, options: SheetAxisDimensionOptions): boolean {
  if (!canEditWorksheetProtection(sheet, userId)) throw new Error('此工作表受保护，当前用户不能调整行列尺寸。')
  const anchor = input.index ?? (action.axis === 'row' ? input.anchorRow : input.anchorCol)
  const targets = sheetAxisDimensionTargetRanges(sheet, action.axis, ranges, anchor), spans = axisSpans(targets, action.axis)
  if (!spans.length) return false
  const before = dimensionSignature(sheet)
  if (action.kind === 'set' || action.kind === 'delta') {
    const delta = action.axis === 'row' ? input.deltaY : input.deltaX
    const value = action.kind === 'delta' ? anchor !== undefined && delta !== undefined
      ? (action.axis === 'row' ? sheetRowHeight(sheet, anchor) : sheetColumnWidth(sheet, anchor)) + delta : undefined : input.value
    if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) throw new RangeError('行列尺寸必须是有限的非负数。')
    setSize(sheet, action.axis, spans, value)
  } else {
    const context = measurementContext(sheet, options), groups = groupedCells(sheet, action.axis, spans)
    if (action.kind === 'grow') {
      const value = input.value
      if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) throw new RangeError('行高必须是有限的非负数。')
      const automatic = value <= (sheet.defaultRowHeight ?? 26)
      if (automatic) setRowMode(sheet, spans, true)
      else setSize(sheet, 'row', spans, value)
      for (const [row, entries] of groups) {
        const natural = calculateSheetAxisAutoFit(sheet, 'row', row, { ...context, entries })
        const singleton = [{ start: row, end: row, value: true }]
        if (value <= natural) { setRowMode(sheet, singleton, true); (sheet.rowAutoHeights ??= {})[row] = natural }
        else setSize(sheet, 'row', singleton, value)
      }
    } else if (action.axis === 'row') {
      setRowMode(sheet, spans, true)
      removeSparseSizes(sheet.rowAutoHeights, spans)
      for (const [row, entries] of groups) {
        const height = calculateSheetAxisAutoFit(sheet, 'row', row, { ...context, entries })
        if (height !== (sheet.defaultRowHeight ?? 26)) (sheet.rowAutoHeights ??= {})[row] = height
      }
    } else for (const [column, entries] of groups) {
      const width = calculateSheetAxisAutoFit(sheet, 'column', column, { ...context, entries })
      if (width !== sheetColumnWidth(sheet, column)) setSize(sheet, 'column', [{ start: column, end: column, value: true }], width)
    }
  }
  return before !== dimensionSignature(sheet)
}

function modelDimensionView(sheet: SpreadsheetSheet, date1904: boolean, withCells = true): ExcelSheetView {
  const cells: Record<string, ExcelCellView> = {}, view: ExcelSheetView = { ...sheet, cells, date1904,
    rowHeights: { ...sheet.rowHeights }, columnWidths: { ...sheet.columnWidths }, rowAutoHeights: { ...sheet.rowAutoHeights },
    rowHeightSpans: sheet.rowHeightSpans?.map(span => ({ ...span })), columnWidthSpans: sheet.columnWidthSpans?.map(span => ({ ...span })),
    rowAutoHeightSpans: sheet.rowAutoHeightSpans?.map(span => ({ ...span })),
    rowCount: sheet.rowCount ?? 1_048_576, columnCount: sheet.columnCount ?? 16_384,
    merges: sheet.mergedRanges.map(range => `${cellKey(range.startRow, range.startCol)}:${cellKey(range.endRow, range.endCol)}`) }
  if (withCells) for (const cell of Object.values(sheet.cells)) cells[cellKey(cell.row, cell.col)] = toExcelCellView(cell)
  return view
}
function copyDimensions(target: SpreadsheetSheet, source: ExcelSheetView): void {
  target.rowHeights = source.rowHeights ?? {}; target.columnWidths = source.columnWidths ?? {}
  target.rowHeightSpans = source.rowHeightSpans; target.columnWidthSpans = source.columnWidthSpans
  target.rowAutoHeightSpans = source.rowAutoHeightSpans; target.rowAutoHeights = source.rowAutoHeights
}
export function applySheetAxisDimensionCommand(sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetAxisDimensionAction,
  input: SheetAxisDimensionCommandInput, userId: string, options: SheetAxisDimensionOptions, date1904 = false): boolean {
  const view = modelDimensionView(sheet, date1904, action.kind === 'auto' || action.kind === 'grow')
  const changed = applySheetViewAxisDimensionCommand(view, ranges, action, input, userId, options)
  if (changed) copyDimensions(sheet, view)
  return changed
}

function heightCellSignature(cell: ExcelCellView | undefined): string {
  return cell ? JSON.stringify([cell.value, cell.formula, cell.rawValue, cell.cachedValue, cell.displayValue, cell.rawNumberFormat, cell.numberFormat,
    cell.fontFamily, cell.fontSize, cell.bold, cell.italic, cell.wrap, cell.wrapStrategy, cell.rotation, cell.rotationVertical, cell.indent, cell.shrinkToFit, cell.richText]) : ''
}
/** Refresh in the same content transaction, after formula values settle. Only sparse populated/cache rows are measured. */
export function refreshSheetViewAutoRowHeights(sheet: ExcelSheetView, previous: ExcelSheetView | undefined, options: SheetAxisDimensionOptions): boolean {
  const all = !previous || JSON.stringify([sheet.columnWidths, sheet.columnWidthSpans, sheet.hiddenColumns, sheet.columnHiddenSpans, sheet.outlineHiddenColumns, sheet.outlineHiddenColumnSpans,
    sheet.merges, sheet.defaultColumnWidth, sheet.defaultRowHeight, sheet.defaultRowAutoHeight, sheet.rowAutoHeightSpans, sheet.date1904]) !== JSON.stringify([previous.columnWidths, previous.columnWidthSpans, previous.hiddenColumns, previous.columnHiddenSpans, previous.outlineHiddenColumns, previous.outlineHiddenColumnSpans,
    previous.merges, previous.defaultColumnWidth, previous.defaultRowHeight, previous.defaultRowAutoHeight, previous.rowAutoHeightSpans, previous.date1904])
  const affected = new Set<number>()
  if (!all) for (const key of new Set([...Object.keys(sheet.cells), ...Object.keys(previous!.cells)])) {
    if (heightCellSignature(sheet.cells[key]) === heightCellSignature(previous!.cells[key])) continue
    const position = parseCellKey(key); if (position) affected.add(position.row)
  }
  if (!all && !affected.size) return false
  const groups = groupedCells(sheet, 'row'), context = measurementContext(sheet, options), heights = { ...(sheet.rowAutoHeights ?? {}) }
  let changed = false
  const candidates = all ? new Set([...groups.keys(), ...Object.keys(heights).map(Number)]) : affected
  for (const row of candidates) {
    if (!sheetRowIsAutoHeight(sheet, row) || row < 0 || row >= sheet.rowCount || !groups.has(row)) {
      if (heights[row] !== undefined) { delete heights[row]; changed = true }
      continue
    }
    const height = calculateSheetAxisAutoFit(sheet, 'row', row, { ...context, entries: groups.get(row)! })
    const next = height === (sheet.defaultRowHeight ?? 26) ? undefined : height
    if (heights[row] === next) continue
    if (next === undefined) delete heights[row]
    else heights[row] = next
    changed = true
  }
  if (changed) sheet.rowAutoHeights = heights
  return changed
}
export function refreshWorkbookAutoRowHeights(workbook: WorkbookModel, previous: WorkbookModel | undefined, options: SheetAxisDimensionOptions): void {
  for (const sheet of workbook.sheets) {
    const old = previous?.sheets.find(item => item.id === sheet.id)
    if (sheet === old) continue
    const view = modelDimensionView(sheet, workbook.date1904), before = old && modelDimensionView(old, previous!.date1904)
    if (refreshSheetViewAutoRowHeights(view, before, options)) copyDimensions(sheet, view)
  }
}
export function refreshWorkbookViewAutoRowHeights(workbook: ExcelWorkbookView, previous: ExcelWorkbookView | undefined, options: SheetAxisDimensionOptions): void {
  for (const sheet of workbook.sheets) refreshSheetViewAutoRowHeights(sheet, previous?.sheets.find(item => item.id === sheet.id), options)
}
