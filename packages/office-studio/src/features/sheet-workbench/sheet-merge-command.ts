import { createCell, defaultCellStyle, type CellRange, type CellStyle, type SpreadsheetCell, type SpreadsheetSheet } from '../../domain/model.ts'
import { cellKey, parseCellKey, type ExcelCellView, type ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { inferSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { SHEET_NUMBER_FORMATS } from './sheet-number-format.ts'
import { clearSheetViewCellContent, clearSpreadsheetCellContent } from './sheet-clear-command.ts'
import { assertSheetCommandRanges, assertSheetRangesEditable, partitionSheetCommandCells, sheetRangesEqual, sheetRangesIntersect,
  type SheetRangeCommandInput } from './sheet-range-command.ts'

export const SHEET_MERGE_COMMANDS = Object.freeze({
  merge: 'sheet.command.merge', all: 'sheet.command.merge-all', horizontal: 'sheet.command.merge-horizontal',
  vertical: 'sheet.command.merge-vertical', unmerge: 'sheet.command.unmerge',
})
export const SHEET_MERGE_COMMAND_ALIASES = Object.freeze({
  merge: 'sheet.command.add-worksheet-merge', all: 'sheet.command.add-worksheet-merge-all',
  horizontal: 'sheet.command.add-worksheet-merge-horizontal', vertical: 'sheet.command.add-worksheet-merge-vertical',
  unmerge: 'sheet.command.remove-worksheet-merge', cells: 'sheet.command.merge-cells',
})
export const SHEET_MERGE_COMMAND_IDS: readonly string[] = [...Object.values(SHEET_MERGE_COMMANDS), ...Object.values(SHEET_MERGE_COMMAND_ALIASES)]
export type SheetMergeMode = 'all' | 'horizontal' | 'vertical'
export interface SheetMergeOptions { defaultMerge?: boolean; isForceMerge?: boolean }
export interface SheetMergeCommandInput extends SheetRangeCommandInput, SheetMergeOptions {
  mode?: SheetMergeMode
  value?: SheetMergeMode | number
  expectedRevision?: number
}
export type SheetMergeAction = { kind: 'merge'; mode: SheetMergeMode; defaultMerge: boolean; isForceMerge: boolean } | { kind: 'unmerge' }
export interface SheetMergePlan { added: CellRange[]; removed: CellRange[]; next: CellRange[]; permissionRanges: CellRange[] }
export interface SheetMergeResult {
  changed: boolean
  requiresConfirmation: boolean
  discardedContentCount: number
  added: CellRange[]
  removed: CellRange[]
}
export interface SheetMergeConfirmationRequest {
  requestId: string
  workbookId: string
  sheetId: string
  commandId: string
  baseRevision: number
  discardedContentCount: number
  input: SheetMergeCommandInput
}

export function mergeActionForCommand(id: string, input?: SheetMergeCommandInput | SheetMergeMode): SheetMergeAction | undefined {
  if (id === SHEET_MERGE_COMMANDS.unmerge || id === SHEET_MERGE_COMMAND_ALIASES.unmerge) return { kind: 'unmerge' }
  if (!SHEET_MERGE_COMMAND_IDS.includes(id)) return undefined
  if (input !== undefined && typeof input !== 'string' && (!input || typeof input !== 'object' || Array.isArray(input))) throw new TypeError('合并参数无效。')
  const params = typeof input === 'string' ? { mode: input } : input
  for (const flag of [params?.defaultMerge, params?.isForceMerge]) {
    if (flag !== undefined && typeof flag !== 'boolean') throw new TypeError('合并选项必须是布尔值。')
  }
  const value = params?.mode ?? params?.value
  const mode = id === SHEET_MERGE_COMMANDS.horizontal || id === SHEET_MERGE_COMMAND_ALIASES.horizontal ? 'horizontal'
    : id === SHEET_MERGE_COMMANDS.vertical || id === SHEET_MERGE_COMMAND_ALIASES.vertical ? 'vertical'
      : id === SHEET_MERGE_COMMANDS.all || id === SHEET_MERGE_COMMAND_ALIASES.all ? 'all'
        : value === 0 ? 'vertical' : value === 1 ? 'horizontal' : value ?? 'all'
  if (mode !== 'all' && mode !== 'vertical' && mode !== 'horizontal') throw new TypeError('合并方向无效。')
  return { kind: 'merge', mode, defaultMerge: params?.defaultMerge ?? false, isForceMerge: params?.isForceMerge ?? false }
}
function mergeRangesForMode(ranges: readonly CellRange[], mode: SheetMergeMode): CellRange[] {
  const result: CellRange[] = []
  for (const range of ranges) {
    if (mode === 'horizontal') {
      if (range.startCol === range.endCol) continue
      for (let row = range.startRow; row <= range.endRow; row++) result.push({ ...range, startRow: row, endRow: row })
    } else if (mode === 'vertical') {
      if (range.startRow === range.endRow) continue
      for (let col = range.startCol; col <= range.endCol; col++) result.push({ ...range, startCol: col, endCol: col })
    } else if (range.startRow !== range.endRow || range.startCol !== range.endCol) result.push({ ...range })
  }
  const seen = new Set<string>()
  return result.filter(range => {
    const key = mergeRangeKey(range)
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}
function mergeRangeKey(range: CellRange): string {
  return [range.startRow, range.startCol, range.endRow, range.endCol].join(':')
}
export function createSheetMergePlan(merges: readonly CellRange[], ranges: readonly CellRange[], action: SheetMergeAction): SheetMergePlan {
  assertSheetCommandRanges(ranges)
  assertSheetCommandRanges(merges)
  if (action.kind === 'unmerge') {
    const removed = merges.filter(merge => ranges.some(range => sheetRangesIntersect(range, merge)))
    const removedSet = new Set(removed)
    return { added: [], removed, next: merges.filter(merge => !removedSet.has(merge)).map(merge => ({ ...merge })), permissionRanges: [...ranges, ...removed] }
  }
  // Check input rectangles before splitting into potentially many independent row/column merges.
  for (let index = 0; index < ranges.length; index++) {
    if (ranges.slice(0, index).some(previous => !sheetRangesEqual(previous, ranges[index]!) && sheetRangesIntersect(previous, ranges[index]!))) {
      throw new RangeError('多个合并区域不能交叉。')
    }
  }
  const scopes = ranges.filter(range => action.mode === 'horizontal' ? range.startCol !== range.endCol
    : action.mode === 'vertical' ? range.startRow !== range.endRow : range.startRow !== range.endRow || range.startCol !== range.endCol)
  const candidates = mergeRangesForMode(scopes, action.mode)
  const removed = merges.filter(merge => scopes.some(range => sheetRangesIntersect(range, merge)))
  if (removed.length && !action.isForceMerge) throw new RangeError('所选区域与现有合并单元格重叠。')
  // A repeated toolbar merge is a no-op, while the Facade's strict overlap check remains above.
  const existingKeys = new Set(merges.map(mergeRangeKey)), candidateKeys = new Set(candidates.map(mergeRangeKey))
  const added = candidates.filter(range => !existingKeys.has(mergeRangeKey(range)))
  const replacements = removed.filter(merge => !candidateKeys.has(mergeRangeKey(merge)))
  const replacementSet = new Set(replacements)
  const next = merges.filter(merge => !replacementSet.has(merge)).map(merge => ({ ...merge }))
  next.push(...added)
  return { added, removed: replacements, next, permissionRanges: [...ranges, ...removed] }
}

function modelHasContent(cell: SpreadsheetCell): boolean {
  return Boolean(cell.formula || cell.richText?.length || cell.image) || cell.value !== null && cell.value !== ''
}
function viewHasContent(cell: ExcelCellView): boolean {
  return Boolean(cell.formula || cell.richText?.length || cell.image) || cell.rawValue !== undefined && cell.rawValue !== null && cell.rawValue !== '' || cell.value !== ''
}
function resultForPlan<C>(
  plan: SheetMergePlan, action: SheetMergeAction, groups: readonly (readonly { cell: C }[])[], hasContent: (cell: C) => boolean,
): SheetMergeResult {
  const discardedContentCount = groups.reduce((count, cells) => count + Math.max(0, cells.reduce((size, position) => size + Number(hasContent(position.cell)), 0) - 1), 0)
  return { changed: false, discardedContentCount, requiresConfirmation: action.kind === 'merge' && !action.defaultMerge && discardedContentCount > 0,
    added: plan.added, removed: plan.removed }
}
function firstNonEmpty<C>(members: readonly { row: number; col: number; cell: C }[], hasContent: (cell: C) => boolean) {
  let first: typeof members[number] | undefined
  for (const member of members) if (hasContent(member.cell) && (!first || member.row < first.row || member.row === first.row && member.col < first.col)) first = member
  return first
}
function restoredBorder(master: CellStyle['border'], actual: CellStyle['border'], range: CellRange, row: number, col: number): CellStyle['border'] {
  const isAnchor = row === range.startRow && col === range.startCol
  const border = { ...(actual ?? {}) }
  const sides = ['top', 'bottom', 'left', 'right'] as const
  for (const side of sides) {
    if (!isAnchor && actual && Object.prototype.hasOwnProperty.call(actual, side)) continue
    const outer = side === 'top' ? row === range.startRow : side === 'bottom' ? row === range.endRow : side === 'left' ? col === range.startCol : col === range.endCol
    if (outer && master && Object.prototype.hasOwnProperty.call(master, side)) border[side] = structuredClone(master[side])
    else if (isAnchor) delete border[side]
  }
  return Object.keys(border).length ? border : undefined
}
function hasNonDefaultStyle(style: CellStyle): boolean {
  const normalized = { ...style }
  if (!normalized.border || !Object.keys(normalized.border).length) delete normalized.border
  if (normalized.wrapStrategy === 'overflow' && !normalized.wrap) delete normalized.wrapStrategy
  const stable = (_key: string, value: unknown): unknown => value && typeof value === 'object' && !Array.isArray(value)
    ? Object.fromEntries(Object.entries(value).sort(([left], [right]) => left.localeCompare(right))) : value
  return JSON.stringify(normalized, stable) !== JSON.stringify(defaultCellStyle(), stable)
}
function restoreUnmergedModelStyles(sheet: SpreadsheetSheet, ranges: readonly CellRange[]): void {
  const groups = partitionSheetCommandCells(ranges, Object.values(sheet.cells))
  for (const [index, range] of ranges.entries()) {
    const master = sheet.cells[range.startRow + ':' + range.startCol]
    if (!master) continue
    const style = structuredClone(master.style), expand = hasNonDefaultStyle(style)
    const positions = expand ? undefined : groups[index]
    const restore = (row: number, col: number) => {
      const key = row + ':' + col, previous = sheet.cells[key]
      if (!previous && !expand) return
      const border = restoredBorder(style.border, previous?.style.border, range, row, col)
      const nextStyle = { ...structuredClone(style), border }
      sheet.cells[key] = { ...(previous ?? createCell(row, col, null)), style: nextStyle }
    }
    if (positions) positions.forEach(cell => restore(cell.row, cell.col))
    else for (let row = range.startRow; row <= range.endRow; row++) for (let col = range.startCol; col <= range.endCol; col++) restore(row, col)
  }
}
function viewStyle(cell: ExcelCellView): CellStyle {
  const defaults = defaultCellStyle()
  return { font: { ...defaults.font, family: cell.fontFamily ?? defaults.font.family, size: cell.fontSize ?? defaults.font.size,
    bold: cell.bold ?? false, italic: cell.italic ?? false, underline: cell.underline ?? false, strike: cell.strike ?? false, color: cell.color ?? null },
    fill: cell.background ?? null, horizontal: cell.align ?? defaults.horizontal, vertical: cell.vertical ?? defaults.vertical,
    wrap: cell.wrap ?? false, wrapStrategy: cell.wrapStrategy, border: cell.border, rotation: cell.rotation === 255 ? 0 : cell.rotation ?? 0,
    rotationVertical: cell.rotationVertical === true || cell.rotation === 255, indent: cell.indent ?? 0, shrinkToFit: cell.shrinkToFit ?? false,
    numberFormat: cell.rawNumberFormat ?? (cell.numberFormat ? SHEET_NUMBER_FORMATS[cell.numberFormat] : 'General'), metadata: cell.styleMetadata ?? undefined }
}
function withViewStyle(cell: ExcelCellView | undefined, style: CellStyle): ExcelCellView {
  return { ...(cell ?? { value: '', rawValue: null }), fontFamily: style.font.family, fontSize: style.font.size,
    bold: style.font.bold, italic: style.font.italic, underline: style.font.underline, strike: style.font.strike, color: style.font.color,
    background: style.fill, align: style.horizontal, vertical: style.vertical, wrap: style.wrap,
    wrapStrategy: style.wrapStrategy ?? (style.wrap ? 'wrap' : 'overflow'), border: style.border ?? {}, rotation: style.rotation ?? 0,
    rotationVertical: style.rotationVertical ?? false, indent: style.indent ?? 0, shrinkToFit: style.shrinkToFit ?? false,
    rawNumberFormat: style.numberFormat, numberFormat: undefined, styleMetadata: style.metadata ?? null }
}
/** A relocated donor replaces the whole cell; undefined view fields otherwise inherit the old master during model adaptation. */
function relocatedViewCell(source: ExcelCellView): ExcelCellView {
  const cell = withViewStyle(structuredClone(source), viewStyle(source))
  return { ...cell, hyperlink: cell.hyperlink ?? null, metadata: cell.metadata ?? null, richText: cell.richText ?? null, image: cell.image ?? null,
    dataType: cell.dataType ?? (cell.formula ? 'formula' : cell.rawValue !== undefined ? inferSheetCellValueType(cell.rawValue) : undefined),
    cachedDataType: cell.formula ? cell.cachedDataType === undefined
      ? cell.rawValue === undefined || cell.rawValue === null ? null : inferSheetCellValueType(cell.rawValue)
      : cell.cachedDataType : null }
}
function viewMergeRanges(sheet: ExcelSheetView): CellRange[] {
  return (sheet.merges ?? []).map(value => {
    const [start, end] = value.split(':'), first = parseCellKey(start ?? ''), last = parseCellKey(end ?? start ?? '')
    if (!first || !last) throw new RangeError('合并区域坐标无效。')
    return { startRow: first.row, startCol: first.col, endRow: last.row, endCol: last.col }
  })
}

export function applySheetMergeCommand(
  sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetMergeAction, userId = 'local-user',
): SheetMergeResult {
  const plan = createSheetMergePlan(sheet.mergedRanges, ranges, action)
  assertSheetRangesEditable(sheet, plan.permissionRanges, userId)
  const groups = partitionSheetCommandCells(plan.added, Object.values(sheet.cells).map(cell => ({ row: cell.row, col: cell.col, cell })))
  const result = resultForPlan(plan, action, groups, modelHasContent)
  if (result.requiresConfirmation || !plan.added.length && !plan.removed.length) return result
  if (action.kind === 'unmerge') restoreUnmergedModelStyles(sheet, plan.removed)
  else for (const [index, range] of plan.added.entries()) {
    const members = groups[index]!
    const first = firstNonEmpty(members, modelHasContent)
    if (first && (first.row !== range.startRow || first.col !== range.startCol)) {
      const key = range.startRow + ':' + range.startCol
      sheet.cells[key] = { ...structuredClone(first.cell), id: key, row: range.startRow, col: range.startCol }
    }
    for (const position of members) {
      if (position.row === range.startRow && position.col === range.startCol) continue
      sheet.cells[position.row + ':' + position.col] = clearSpreadsheetCellContent(position.cell)
    }
  }
  sheet.mergedRanges = plan.next
  for (const range of plan.added) { sheet.rowCount = Math.max(sheet.rowCount ?? 0, range.endRow + 1); sheet.columnCount = Math.max(sheet.columnCount ?? 0, range.endCol + 1) }
  result.changed = true
  return result
}
export function applySheetViewMergeCommand(
  sheet: ExcelSheetView, ranges: readonly CellRange[], action: SheetMergeAction, userId = 'local-user',
): SheetMergeResult {
  const plan = createSheetMergePlan(viewMergeRanges(sheet), ranges, action)
  assertSheetRangesEditable(sheet, plan.permissionRanges, userId)
  const cells = Object.entries(sheet.cells).flatMap(([key, cell]) => {
    const position = parseCellKey(key)
    return position ? [{ ...position, cell }] : []
  })
  const groups = partitionSheetCommandCells(plan.added, cells)
  const result = resultForPlan(plan, action, groups, viewHasContent)
  if (result.requiresConfirmation || !plan.added.length && !plan.removed.length) return result
  if (action.kind === 'unmerge') {
    const removedGroups = partitionSheetCommandCells(plan.removed, cells)
    for (const [index, range] of plan.removed.entries()) {
      const master = sheet.cells[cellKey(range.startRow, range.startCol)]
      if (!master) continue
      const style = viewStyle(master), expand = hasNonDefaultStyle(style)
      const restore = (row: number, col: number) => {
        const key = cellKey(row, col), previous = sheet.cells[key]
        if (!previous && !expand) return
        sheet.cells[key] = withViewStyle(previous, { ...structuredClone(style), border: restoredBorder(style.border, previous?.border, range, row, col) })
      }
      if (!expand) removedGroups[index]!.forEach(position => restore(position.row, position.col))
      else for (let row = range.startRow; row <= range.endRow; row++) for (let col = range.startCol; col <= range.endCol; col++) restore(row, col)
    }
  } else for (const [index, range] of plan.added.entries()) {
    const members = groups[index]!
    const first = firstNonEmpty(members, viewHasContent)
    if (first && (first.row !== range.startRow || first.col !== range.startCol)) sheet.cells[cellKey(range.startRow, range.startCol)] = relocatedViewCell(first.cell)
    for (const position of members) {
      if (position.row !== range.startRow || position.col !== range.startCol) sheet.cells[cellKey(position.row, position.col)] = clearSheetViewCellContent(position.cell)
    }
  }
  sheet.merges = plan.next.map(range => cellKey(range.startRow, range.startCol) + ':' + cellKey(range.endRow, range.endCol))
  for (const range of plan.added) { sheet.rowCount = Math.max(sheet.rowCount, range.endRow + 1); sheet.columnCount = Math.max(sheet.columnCount, range.endCol + 1) }
  result.changed = true
  return result
}

export function previewSheetMergeCommand(
  sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetMergeAction, userId = 'local-user',
): SheetMergeResult {
  const plan = createSheetMergePlan(sheet.mergedRanges, ranges, action)
  assertSheetRangesEditable(sheet, plan.permissionRanges, userId)
  const groups = partitionSheetCommandCells(plan.added, Object.values(sheet.cells).map(cell => ({ row: cell.row, col: cell.col, cell })))
  return resultForPlan(plan, action, groups, modelHasContent)
}
