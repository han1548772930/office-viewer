import { createCell, defaultCellStyle, type CellRange, type CellStyle, type SpreadsheetSheet } from '../../domain/model.ts'
import { parseSheetCellText } from '../../domain/sheet-cell-value.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { cellKey, parseCellKey, type ExcelCellView, type ExcelSheetView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { applySheetMergeCommand, applySheetViewMergeCommand, createSheetMergePlan, previewSheetMergeCommand,
  type SheetMergePlan } from './sheet-merge-command.ts'
import { assertSheetCommandRanges, assertSheetRangesEditable, sheetRangeContainsCell, sheetRangesIntersect } from './sheet-range-command.ts'
import { classifySheetNumberFormat, formatSheetValue, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from './sheet-number-format.ts'
import { assertConditionalFormatPainterSource, captureConditionalFormatPainterSource, createConditionalFormatPainterPlan,
  type SheetConditionalFormatPainterSource } from './sheet-format-painter-conditional.ts'

/** The existing synchronous editor budget includes expanded brush and merge targets. */
export const MAXIMUM_FORMAT_PAINTER_CELLS = 100_000
/** Only formatting is sampled; sparse relative coordinates represent default-style cells implicitly. */
export interface SheetFormatPainterSource {
  readonly range: CellRange
  readonly styles: Readonly<Record<string, CellStyle>>
  readonly merges: readonly CellRange[]
  readonly conditional?: SheetConditionalFormatPainterSource
}
export interface SheetFormatPainterCanvasState {
  mode: 'once' | 'continuous'
  sourceSheetId: string
  sourceRange: GridRange
}
export interface SheetFormatPainterPlan {
  ranges: CellRange[]
  mergeRanges: CellRange[]
  mergePlan: SheetMergePlan
  permissionRanges: CellRange[]
}
export interface SheetFormatPainterResult {
  applied: boolean
  changed: boolean
  contentChanged: boolean
  requiresConfirmation: boolean
  discardedContentCount: number
  ranges: CellRange[]
  added: CellRange[]
  removed: CellRange[]
}
export interface SheetFormatPainterOptions { mergeConfirmed?: boolean; userId?: string; date1904?: boolean; region?: string }

function normalizedStyle(style?: CellStyle): CellStyle {
  const defaults = defaultCellStyle()
  return { ...style, font: { ...defaults.font, ...style?.font }, fill: style?.fill ?? defaults.fill,
    horizontal: style?.horizontal ?? defaults.horizontal, vertical: style?.vertical ?? defaults.vertical,
    wrap: style?.wrap ?? false, wrapStrategy: style?.wrapStrategy ?? (style?.wrap ? 'wrap' : 'overflow'),
    numberFormat: style?.numberFormat ?? defaults.numberFormat, border: style?.border ?? {},
    rotation: style?.rotation === 255 ? 0 : style?.rotation ?? 0,
    rotationVertical: style?.rotationVertical === true || style?.rotation === 255,
    indent: style?.indent ?? 0, shrinkToFit: style?.shrinkToFit ?? false,
    metadata: style?.metadata && Object.keys(style.metadata).length ? style.metadata : undefined }
}
function viewStyle(cell?: ExcelCellView): CellStyle {
  const defaults = defaultCellStyle()
  return normalizedStyle({ font: { ...defaults.font, family: cell?.fontFamily ?? defaults.font.family,
    size: cell?.fontSize ?? defaults.font.size, bold: cell?.bold ?? false, italic: cell?.italic ?? false,
    underline: cell?.underline ?? false, strike: cell?.strike ?? false, color: cell?.color ?? null },
    fill: cell?.background ?? null, horizontal: cell?.align ?? defaults.horizontal, vertical: cell?.vertical ?? defaults.vertical,
    wrap: cell?.wrap ?? false, wrapStrategy: cell?.wrapStrategy, numberFormat: cell?.rawNumberFormat
      ?? (cell?.numberFormat ? SHEET_NUMBER_FORMATS[cell.numberFormat] : 'General'),
    border: cell?.border, rotation: cell?.rotation, rotationVertical: cell?.rotationVertical,
    indent: cell?.indent, shrinkToFit: cell?.shrinkToFit, metadata: cell?.styleMetadata ?? undefined })
}
function stableStyle(style: CellStyle): string {
  return JSON.stringify(normalizedStyle(style), (_key, value: unknown) => value && typeof value === 'object' && !Array.isArray(value)
    ? Object.fromEntries(Object.entries(value).sort(([left], [right]) => left.localeCompare(right))) : value)
}
const DEFAULT_STYLE = normalizedStyle()
const DEFAULT_STYLE_KEY = stableStyle(DEFAULT_STYLE)

function freezeSource(source: SheetFormatPainterSource): SheetFormatPainterSource {
  const freeze = (value: unknown): void => {
    if (!value || typeof value !== 'object' || Object.isFrozen(value)) return
    for (const nested of Object.values(value)) freeze(nested)
    Object.freeze(value)
  }
  freeze(source)
  return source
}
function expandedSourceRange(range: CellRange, merges: readonly CellRange[]): CellRange {
  assertSheetCommandRanges([range, ...merges])
  let result = { ...range }, changed = true
  while (changed) {
    changed = false
    for (const merge of merges) {
      if (!sheetRangesIntersect(result, merge)) continue
      const expanded = { startRow: Math.min(result.startRow, merge.startRow), startCol: Math.min(result.startCol, merge.startCol),
        endRow: Math.max(result.endRow, merge.endRow), endCol: Math.max(result.endCol, merge.endCol) }
      if (expanded.startRow !== result.startRow || expanded.startCol !== result.startCol || expanded.endRow !== result.endRow || expanded.endCol !== result.endCol) {
        result = expanded
        changed = true
      }
    }
  }
  return result
}
function collectSource(range: CellRange, merges: readonly CellRange[], cells: Iterable<{ row: number; col: number; style: CellStyle }>, conditional?: SheetConditionalFormatPainterSource): SheetFormatPainterSource {
  const expanded = expandedSourceRange(range, merges), styles: Record<string, CellStyle> = {}
  for (const cell of cells) {
    if (!sheetRangeContainsCell(expanded, cell.row, cell.col)) continue
    const style = normalizedStyle(cell.style)
    if (stableStyle(style) !== DEFAULT_STYLE_KEY) styles[`${cell.row - expanded.startRow}:${cell.col - expanded.startCol}`] = cloneModel(style)
  }
  return freezeSource({ range: expanded, styles,
    merges: merges.filter(merge => sheetRangesIntersect(expanded, merge)).map(merge => ({ ...merge })), ...(conditional ? { conditional } : {}) })
}
export function captureSheetFormatPainterSource(sheet: SpreadsheetSheet, range: CellRange, userId = 'local-user'): SheetFormatPainterSource {
  const expanded = expandedSourceRange(range, sheet.mergedRanges)
  assertSheetRangesEditable(sheet, [expanded], userId)
  return collectSource(expanded, sheet.mergedRanges, Object.values(sheet.cells).map(cell => ({ row: cell.row, col: cell.col, style: cell.style })),
    captureConditionalFormatPainterSource(sheet.id, sheet.conditionalFormats, expanded))
}
function viewMerges(sheet: ExcelSheetView): CellRange[] {
  return (sheet.merges ?? []).map(value => {
    const [start, end] = value.split(':'), first = parseCellKey(start ?? ''), last = parseCellKey(end ?? start ?? '')
    if (!first || !last) throw new RangeError('合并区域坐标无效。')
    return { startRow: first.row, startCol: first.col, endRow: last.row, endCol: last.col }
  })
}
export function captureSheetViewFormatPainterSource(sheet: ExcelSheetView, range: CellRange, userId = 'local-user'): SheetFormatPainterSource {
  const merges = viewMerges(sheet), expanded = expandedSourceRange(range, merges)
  assertSheetRangesEditable(sheet, [expanded], userId)
  const cells = Object.entries(sheet.cells).flatMap(([key, cell]) => {
    const position = parseCellKey(key)
    return position ? [{ ...position, style: viewStyle(cell) }] : []
  })
  return collectSource(expanded, merges, cells, captureConditionalFormatPainterSource(sheet.id, sheet.conditionalFormats, expanded))
}
function assertJsonStyle(value: unknown, ancestors = new Set<object>()): void {
  if (value === undefined || value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value)) return
  if (typeof value !== 'object' || ancestors.has(value)) throw new TypeError('格式刷源样式必须是无循环的 JSON 数据。')
  const prototype = Object.getPrototypeOf(value)
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) throw new TypeError('格式刷源样式必须是 JSON 数据。')
  ancestors.add(value)
  for (const nested of Object.values(value)) assertJsonStyle(nested, ancestors)
  ancestors.delete(value)
}
function validateSource(source: SheetFormatPainterSource): void {
  if (!source || !source.styles || typeof source.styles !== 'object' || Array.isArray(source.styles) || !Array.isArray(source.merges)) throw new TypeError('格式刷源快照无效。')
  assertSheetCommandRanges([source.range, ...source.merges])
  if (source.conditional) { assertConditionalFormatPainterSource(source.conditional); assertJsonStyle(source.conditional) }
  const height = source.range.endRow - source.range.startRow + 1, width = source.range.endCol - source.range.startCol + 1
  for (const [key, style] of Object.entries(source.styles)) {
    const coordinate = /^(0|[1-9]\d*):(0|[1-9]\d*)$/.exec(key)
    if (!coordinate || !Number.isSafeInteger(Number(coordinate[1])) || !Number.isSafeInteger(Number(coordinate[2]))
      || Number(coordinate[1]) >= height || Number(coordinate[2]) >= width) throw new RangeError('格式刷样式坐标必须位于源采样范围内。')
    const record = (value: unknown) => !!value && typeof value === 'object' && !Array.isArray(value)
    const optionalBoolean = (value: unknown) => value === undefined || typeof value === 'boolean'
    const color = (value: unknown) => value === null || typeof value === 'string'
    if (!record(style) || !record(style.font) || typeof style.font.family !== 'string' || !style.font.family.trim()
      || !Number.isFinite(style.font.size) || style.font.size <= 0
      || [style.font.bold, style.font.italic, style.font.underline, style.font.strike, style.wrap].some(value => typeof value !== 'boolean')
      || !color(style.font.color) || !color(style.fill) || typeof style.numberFormat !== 'string'
      || !['general', 'left', 'center', 'right', 'justify'].includes(style.horizontal)
      || !['top', 'middle', 'bottom'].includes(style.vertical)
      || style.wrapStrategy !== undefined && !['overflow', 'wrap', 'clip'].includes(style.wrapStrategy)
      || style.rotation !== undefined && !Number.isFinite(style.rotation)
      || style.indent !== undefined && (!Number.isFinite(style.indent) || style.indent < 0)
      || !optionalBoolean(style.rotationVertical) || !optionalBoolean(style.shrinkToFit)
      || style.border !== undefined && !record(style.border) || style.metadata !== undefined && !record(style.metadata)) {
      throw new TypeError('格式刷源样式无效。')
    }
    // JSON snapshots can be Vue proxies; validate their data before any merge changes content.
    assertJsonStyle(style)
  }
  for (const merge of source.merges) if (!sheetRangeContainsCell(source.range, merge.startRow, merge.startCol) || !sheetRangeContainsCell(source.range, merge.endRow, merge.endCol)) {
    throw new RangeError('源合并区域必须位于格式刷采样范围内。')
  }
}
export function cloneSheetFormatPainterSource(source: SheetFormatPainterSource): SheetFormatPainterSource {
  validateSource(source)
  return freezeSource(cloneModel(source))
}
function* uniquePositions(ranges: readonly CellRange[]): Generator<{ row: number; col: number }> {
  const visited = new Set<number>()
  for (const range of ranges) {
    if ((range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1) > MAXIMUM_FORMAT_PAINTER_CELLS) throw new RangeError(`格式刷当前最多同步处理 ${MAXIMUM_FORMAT_PAINTER_CELLS.toLocaleString()} 个单元格。`)
    for (let row = range.startRow; row <= range.endRow; row++) for (let col = range.startCol; col <= range.endCol; col++) {
      const key = row * 16_384 + col
      if (visited.has(key)) continue
      visited.add(key)
      if (visited.size > MAXIMUM_FORMAT_PAINTER_CELLS) throw new RangeError(`格式刷当前最多同步处理 ${MAXIMUM_FORMAT_PAINTER_CELLS.toLocaleString()} 个单元格。`)
      yield { row, col }
    }
  }
}
function checkCost(ranges: readonly CellRange[]): void {
  const positions = uniquePositions(ranges)
  while (!positions.next().done) { /* Complete the bounded union check before producing mutations. */ }
}
const MERGE_ACTION = { kind: 'merge', mode: 'all', defaultMerge: true, isForceMerge: true } as const
/** Single cells expand to source size; other ranges retain their size and tile complete source merge blocks. */
export function createSheetFormatPainterPlan(merges: readonly CellRange[], source: SheetFormatPainterSource, targets: readonly CellRange[]): SheetFormatPainterPlan {
  validateSource(source)
  assertSheetCommandRanges(targets)
  const height = source.range.endRow - source.range.startRow + 1, width = source.range.endCol - source.range.startCol + 1
  const ranges = targets.map(target => target.startRow === target.endRow && target.startCol === target.endCol
    ? { ...target, endRow: target.startRow + height - 1, endCol: target.startCol + width - 1 } : { ...target })
  assertSheetCommandRanges(ranges)
  checkCost(ranges)
  const mergeRanges: CellRange[] = []
  for (const target of ranges) {
    const rowRepeats = Math.max(1, Math.floor((target.endRow - target.startRow + 1) / height))
    const colRepeats = Math.max(1, Math.floor((target.endCol - target.startCol + 1) / width))
    for (const merge of source.merges) for (let row = 0; row < rowRepeats; row++) for (let col = 0; col < colRepeats; col++) mergeRanges.push({
      startRow: target.startRow + row * height + merge.startRow - source.range.startRow,
      endRow: target.startRow + row * height + merge.endRow - source.range.startRow,
      startCol: target.startCol + col * width + merge.startCol - source.range.startCol,
      endCol: target.startCol + col * width + merge.endCol - source.range.startCol,
    })
  }
  assertSheetCommandRanges(mergeRanges)
  checkCost([...ranges, ...mergeRanges])
  const mergePlan = createSheetMergePlan(merges, mergeRanges, MERGE_ACTION)
  // Replacing an existing merge can restore cells well outside the painted rectangle.
  checkCost([...ranges, ...mergePlan.permissionRanges])
  return { ranges, mergeRanges, mergePlan, permissionRanges: [...ranges, ...mergePlan.permissionRanges] }
}
function sourceStyle(source: SheetFormatPainterSource, target: CellRange, row: number, col: number): CellStyle {
  const height = source.range.endRow - source.range.startRow + 1, width = source.range.endCol - source.range.startCol + 1
  return source.styles[`${(row - target.startRow) % height}:${(col - target.startCol) % width}`] ?? DEFAULT_STYLE
}
function styleAssignments(source: SheetFormatPainterSource, ranges: readonly CellRange[]): Map<number, { row: number; col: number; style: CellStyle }> {
  const assignments = new Map<number, { row: number; col: number; style: CellStyle }>()
  // A later explicit target owns overlapping cells; each physical cell is mutated once.
  for (const target of ranges) for (let row = target.startRow; row <= target.endRow; row++) for (let col = target.startCol; col <= target.endCol; col++) {
    assignments.set(row * 16_384 + col, { row, col, style: sourceStyle(source, target, row, col) })
  }
  return assignments
}
export function previewSheetFormatPainterCommand(sheet: SpreadsheetSheet, source: SheetFormatPainterSource, targets: readonly CellRange[], options: SheetFormatPainterOptions = {}): SheetFormatPainterResult {
  const plan = createSheetFormatPainterPlan(sheet.mergedRanges, source, targets)
  assertSheetRangesEditable(sheet, plan.permissionRanges, options.userId ?? 'local-user')
  const merge = previewSheetMergeCommand(sheet, plan.mergeRanges, { ...MERGE_ACTION, defaultMerge: options.mergeConfirmed ?? false }, options.userId ?? 'local-user')
  return { applied: false, changed: false, contentChanged: false, requiresConfirmation: merge.requiresConfirmation,
    discardedContentCount: merge.discardedContentCount, ranges: plan.ranges, added: plan.mergePlan.added, removed: plan.mergePlan.removed }
}
export function applySheetFormatPainterCommand(sheet: SpreadsheetSheet, source: SheetFormatPainterSource, targets: readonly CellRange[], options: SheetFormatPainterOptions = {}): SheetFormatPainterResult {
  const preview = previewSheetFormatPainterCommand(sheet, source, targets, options)
  if (preview.requiresConfirmation || !preview.ranges.length) return preview
  const assignments = styleAssignments(source, preview.ranges), plan = createSheetFormatPainterPlan(sheet.mergedRanges, source, targets)
  const conditional = createConditionalFormatPainterPlan(sheet.id, sheet.conditionalFormats, source.conditional, source.range, preview.ranges)
  const merge = applySheetMergeCommand(sheet, plan.mergeRanges, MERGE_ACTION, options.userId ?? 'local-user')
  let changed = merge.changed || conditional.changed
  for (const { row, col, style } of assignments.values()) {
    const key = `${row}:${col}`, current = sheet.cells[key]
    if (stableStyle(current?.style ?? DEFAULT_STYLE) === stableStyle(style)) continue
    const cell = current ?? createCell(row, col)
    cell.style = cloneModel(style)
    // Reformat destination raw/cache values; source values and formula caches are never sampled.
    if (cell.formula === null || cell.value !== null) cell.display = formatSheetValue(cell.value, style.numberFormat, options.date1904, options.region) ?? cell.display
    sheet.cells[key] = cell
    sheet.rowCount = Math.max(sheet.rowCount ?? 0, row + 1)
    sheet.columnCount = Math.max(sheet.columnCount ?? 0, col + 1)
    changed = true
  }
  if (conditional.changed) sheet.conditionalFormats = conditional.rules
  return { ...preview, applied: true, changed, contentChanged: merge.changed }
}
function withViewStyle(current: ExcelCellView | undefined, style: CellStyle, options: SheetFormatPainterOptions): ExcelCellView {
  const cell = current ?? { value: '', rawValue: null }, classification = classifySheetNumberFormat(style.numberFormat)
  const next: ExcelCellView = { ...cell, fontFamily: style.font.family, fontSize: style.font.size,
    bold: style.font.bold, italic: style.font.italic, underline: style.font.underline, strike: style.font.strike,
    color: style.font.color, background: style.fill, align: style.horizontal, vertical: style.vertical,
    wrap: style.wrap, wrapStrategy: style.wrapStrategy ?? (style.wrap ? 'wrap' : 'overflow'),
    rawNumberFormat: style.numberFormat, numberFormat: classification === 'custom' ? undefined : classification as SheetNumberFormat,
    border: cloneModel(style.border ?? {}), rotation: style.rotation ?? 0, rotationVertical: style.rotationVertical ?? false,
    indent: style.indent ?? 0, shrinkToFit: style.shrinkToFit ?? false,
    styleMetadata: style.metadata ? cloneModel(style.metadata) : null }
  if (viewStyle(cell).numberFormat !== style.numberFormat) {
    const raw = cell.rawValue !== undefined ? cell.rawValue : parseSheetCellText(cell.formula ? cell.cachedValue ?? cell.displayValue ?? '' : cell.value,
      cell.formula ? cell.cachedDataType ?? undefined : cell.dataType)
    if (!cell.formula || raw !== null) {
      const display = formatSheetValue(raw, style.numberFormat, options.date1904, options.region)
      if (display !== undefined) { next.displayValue = display; if (cell.formula) next.cachedValue = display }
    }
  }
  return next
}
export function applySheetViewFormatPainterCommand(sheet: ExcelSheetView, source: SheetFormatPainterSource, targets: readonly CellRange[], options: SheetFormatPainterOptions = {}): SheetFormatPainterResult {
  const plan = createSheetFormatPainterPlan(viewMerges(sheet), source, targets)
  assertSheetRangesEditable(sheet, plan.permissionRanges, options.userId ?? 'local-user')
  const candidate: ExcelSheetView = { ...sheet, cells: { ...sheet.cells }, merges: [...(sheet.merges ?? [])] }
  const conditional = createConditionalFormatPainterPlan(sheet.id, sheet.conditionalFormats, source.conditional, source.range, plan.ranges)
  const merge = applySheetViewMergeCommand(candidate, plan.mergeRanges, { ...MERGE_ACTION, defaultMerge: options.mergeConfirmed ?? false }, options.userId ?? 'local-user')
  const result: SheetFormatPainterResult = { applied: false, changed: false, contentChanged: false,
    requiresConfirmation: merge.requiresConfirmation, discardedContentCount: merge.discardedContentCount,
    ranges: plan.ranges, added: merge.added, removed: merge.removed }
  if (merge.requiresConfirmation || !plan.ranges.length) return result
  let changed = merge.changed || conditional.changed
  for (const { row, col, style } of styleAssignments(source, plan.ranges).values()) {
    const key = cellKey(row, col), current = candidate.cells[key]
    if (stableStyle(viewStyle(current)) === stableStyle(style)) continue
    candidate.cells[key] = withViewStyle(current, style, { ...options, date1904: options.date1904 ?? sheet.date1904 })
    candidate.rowCount = Math.max(candidate.rowCount, row + 1)
    candidate.columnCount = Math.max(candidate.columnCount, col + 1)
    changed = true
  }
  if (changed) { sheet.cells = candidate.cells; sheet.merges = candidate.merges; sheet.rowCount = candidate.rowCount; sheet.columnCount = candidate.columnCount }
  if (conditional.changed) sheet.conditionalFormats = conditional.rules
  return { ...result, applied: true, changed, contentChanged: merge.changed }
}

/** Compatibility for the earlier view helper; new callers use the shared model/view commands. */
export interface FormatPainterSource { range: GridRange; cells: readonly (readonly (ExcelCellView | undefined)[])[] }
export interface FormatPainterPatch { row: number; column: number; cell: ExcelCellView }
export function createFormatPainterPatches(source: FormatPainterSource, targets: readonly GridRange[], targetCells: Readonly<Record<string, ExcelCellView | undefined>>, address: (row: number, column: number) => string): FormatPainterPatch[] {
  const range = { startRow: source.range.startRow, endRow: source.range.endRow, startCol: source.range.startColumn, endCol: source.range.endColumn }
  const cells = source.cells.flatMap((row, rowIndex) => row.flatMap((cell, colIndex) => cell ? [{ row: range.startRow + rowIndex, col: range.startCol + colIndex, style: viewStyle(cell) }] : []))
  const snapshot = collectSource(range, [], cells), plan = createSheetFormatPainterPlan([], snapshot, targets.map(target => ({ startRow: target.startRow, endRow: target.endRow, startCol: target.startColumn, endCol: target.endColumn })))
  return [...styleAssignments(snapshot, plan.ranges).values()].map(({ row, col, style }) => ({ row, column: col, cell: withViewStyle(targetCells[address(row, col)], style, {}) }))
}
