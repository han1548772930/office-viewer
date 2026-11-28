import { createCell, type CellRange, type SpreadsheetCell, type SpreadsheetSheet, type WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { mapSheetCellShiftPosition, sheetCellShiftAffectedRange, sheetCellShiftContainsBand, transformSheetCellShiftRange,
  type SheetCellShiftEdit } from '../../domain/sheet-cell-shift.ts'
import { sheetRangeAxisInterval } from '../../domain/sheet-axis-structure.ts'
import { assertSheetCommandRanges, assertSheetRangesEditable, normalizeSheetCommandRange, sheetRangeCommandInputRanges, sheetRangeContainsCell, sheetRangesIntersect,
  type SheetRangeCommandInput } from './sheet-range-command.ts'
import { planSheetRangeThemeCellShift } from './sheet-range-theme-structure.ts'
import { refreshWorkbookCellStyles, sheetCellAuthoredStyle } from './sheet-style-inheritance.ts'
import { transformWorkbookCellShiftResources } from './sheet-cell-shift-resources.ts'
import { assertSheetClipboardWorkbookChangesEditable } from './sheet-clipboard-permissions.ts'
import { compactSheetClipboardRanges } from './sheet-clipboard-changes.ts'
import { assertSheetRangeChangedValuesValid } from './sheet-range-editor-validation.ts'

export const SHEET_CELL_SHIFT_COMMANDS = Object.freeze({
  right: 'sheet.command.insert-range-move-right', down: 'sheet.command.insert-range-move-down',
  left: 'sheet.command.delete-range-move-left', up: 'sheet.command.delete-range-move-up',
})
export const SHEET_CELL_SHIFT_CONFIRM_COMMANDS = Object.freeze({
  right: 'sheet.command.insert-range-move-right-confirm', down: 'sheet.command.insert-range-move-down-confirm',
  left: 'sheet.command.delete-range-move-left-confirm', up: 'sheet.command.delete-range-move-up-confirm',
})
export const SHEET_CELL_SHIFT_COMMAND_IDS: readonly string[] = [...Object.values(SHEET_CELL_SHIFT_COMMANDS), ...Object.values(SHEET_CELL_SHIFT_CONFIRM_COMMANDS)]
export type SheetCellShiftDirection = keyof typeof SHEET_CELL_SHIFT_COMMANDS
export const SheetCellShiftDimensions = Object.freeze({ COLUMNS: 0, ROWS: 1 } as const)
export type SheetCellShiftDimension = typeof SheetCellShiftDimensions[keyof typeof SheetCellShiftDimensions]
export interface SheetCellShiftCommandInput extends SheetRangeCommandInput {
  dismantlePartialMerges?: boolean
  expectedRevision?: number
  expectedHistoryGeneration?: number
  host?: { unitId: string; sheetId: string }
}
export interface SheetCellShiftConfirmationRequest {
  requestId: string
  workbookId: string
  sheetId: string
  commandId: string
  baseRevision: number
  mergeRanges: CellRange[]
  input: SheetCellShiftCommandInput
}
export interface SheetCellShiftPlan {
  workbook: WorkbookModel
  range: CellRange
  edit: SheetCellShiftEdit
  affected: Record<string, CellRange[]>
  mergeRanges: CellRange[]
  requiresConfirmation: boolean
}
export function sheetCellShiftDirectionForCommand(id: string): SheetCellShiftDirection | undefined {
  return (Object.keys(SHEET_CELL_SHIFT_COMMANDS) as SheetCellShiftDirection[]).find(direction =>
    SHEET_CELL_SHIFT_COMMANDS[direction] === id || SHEET_CELL_SHIFT_CONFIRM_COMMANDS[direction] === id)
}
export function sheetCellShiftCommandInput(value: unknown): SheetCellShiftCommandInput {
  if (value === undefined) return {}
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('单元格移位命令参数无效。')
  const input = value as SheetCellShiftCommandInput, ranges = sheetRangeCommandInputRanges(input)
  const range = input.range === undefined ? undefined : normalizeSheetCommandRange(input.range)
  if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('工作表标识不一致。')
  if (input.dismantlePartialMerges !== undefined && typeof input.dismantlePartialMerges !== 'boolean') throw new TypeError('拆分合并选项必须是布尔值。')
  for (const [key, number] of Object.entries({ expectedRevision: input.expectedRevision, expectedHistoryGeneration: input.expectedHistoryGeneration })) {
    if (number !== undefined && (!Number.isSafeInteger(number) || number < (key === 'expectedRevision' ? 0 : 1))) throw new TypeError('工作簿内容或历史版本无效。')
  }
  if (input.host !== undefined && (!input.host || typeof input.host !== 'object' || Array.isArray(input.host)
    || typeof input.host.unitId !== 'string' || !input.host.unitId.trim() || typeof input.host.sheetId !== 'string' || !input.host.sheetId.trim())) throw new TypeError('单元格移位宿主无效。')
  return { ...input, ...(range === undefined ? {} : { range }), ...(ranges === undefined ? {} : { ranges }), ...(input.host ? { host: { ...input.host } } : {}) }
}
/** An explicit primary range wins over the selection list; otherwise preserve the current active index. */
export function sheetCellShiftCommandRange(input: SheetCellShiftCommandInput, ranges: readonly CellRange[], activeRangeIndex?: number): CellRange | undefined {
  if (input.range !== undefined) return normalizeSheetCommandRange(input.range)
  const explicit = input.ranges !== undefined || input.selections !== undefined
  const range = ranges[explicit ? ranges.length - 1 : activeRangeIndex ?? ranges.length - 1] ?? ranges.at(-1)
  return range ? { ...range } : undefined
}
export function sheetCellShiftEdit(range: CellRange, direction: SheetCellShiftDirection): SheetCellShiftEdit {
  assertSheetCommandRanges([range])
  if (!(direction in SHEET_CELL_SHIFT_COMMANDS)) throw new TypeError('单元格移位方向无效。')
  const vertical = direction === 'down' || direction === 'up'
  return { axis: vertical ? 'row' : 'column', kind: direction === 'down' || direction === 'right' ? 'insert' : 'delete',
    start: vertical ? range.startRow : range.startCol, count: vertical ? range.endRow - range.startRow + 1 : range.endCol - range.startCol + 1,
    bandStart: vertical ? range.startCol : range.startRow, bandEnd: vertical ? range.endCol : range.endRow }
}
/** Fixed UI confirm commands reject vertical shifts through filtered rows, including the affected tail. */
export function assertSheetCellShiftConfirmationAllowed(model: WorkbookModel, sheetId: string, edit: SheetCellShiftEdit): void {
  if (edit.axis !== 'row') return
  const sheet = model.sheets.find(item => item.id === sheetId)
  if (!sheet) throw new Error('找不到移位目标工作表。')
  const end = sheet.rowCount ?? 1_048_576
  const filtered = [...(sheet.autoFilter?.filteredRows ?? []), ...(model.tables ?? [])
    .filter(table => table.sheetId === sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])]
  if (filtered.some(row => row >= edit.start && row < end)) throw new RangeError('受影响的尾部含筛选隐藏行，不能上下移动单元格。请先清除筛选。')
}
function styleOnlyCell(source: SpreadsheetCell): SpreadsheetCell | undefined {
  const style = sheetCellAuthoredStyle(source)
  if (!Object.keys(style).length && !source.authoredStyleReference && !source.clearedStyleProperties?.length && source.themeStyle == null) return undefined
  const cell = createCell(source.row, source.col, null)
  cell.authoredStyle = style
  if (source.authoredStyleReference !== undefined) cell.authoredStyleReference = source.authoredStyleReference
  if (source.clearedStyleProperties) cell.clearedStyleProperties = [...source.clearedStyleProperties]
  if (source.themeStyle !== undefined) cell.themeStyle = cloneModel(source.themeStyle)
  return cell
}
function shiftedMerges(sheet: SpreadsheetSheet, edit: SheetCellShiftEdit): { ranges: CellRange[]; dismantled: CellRange[] } {
  const ranges: CellRange[] = [], dismantled: CellRange[] = [], affected = sheetCellShiftAffectedRange(edit)
  for (const merge of sheet.mergedRanges) {
    if (!sheetRangesIntersect(merge, affected)) { ranges.push({ ...merge }); continue }
    const axis = sheetRangeAxisInterval(merge, edit.axis)
    if (sheetCellShiftContainsBand(merge, edit) && (axis.start >= edit.start + (edit.kind === 'delete' ? edit.count : 0)
      || edit.kind === 'delete' && axis.start >= edit.start && axis.end < edit.start + edit.count)) {
      ranges.push(...transformSheetCellShiftRange(merge, edit, false))
    } else dismantled.push({ ...merge })
  }
  return { ranges, dismantled }
}

/** Build and validate the complete candidate before asking about partial merges or publishing any content. */
export function planWorkbookCellShiftCommand(model: WorkbookModel, sheetId: string, range: CellRange, direction: SheetCellShiftDirection,
  userId = 'local-user', dismantlePartialMerges = false, selections: readonly CellRange[] = [range], region?: string): SheetCellShiftPlan {
  assertSheetCommandRanges(selections)
  const edit = sheetCellShiftEdit(range, direction), workbook = cloneModel(model), sheet = workbook.sheets.find(item => item.id === sheetId)
  if (!sheet) throw new Error('找不到移位目标工作表。')
  if (!Number.isSafeInteger(sheet.rowCount ?? 1_048_576) || (sheet.rowCount ?? 1_048_576) < 1 || (sheet.rowCount ?? 1_048_576) > 1_048_576
    || !Number.isSafeInteger(sheet.columnCount ?? 16_384) || (sheet.columnCount ?? 16_384) < 1 || (sheet.columnCount ?? 16_384) > 16_384) throw new RangeError('移位工作表行列容量无效。')
  if (range.endRow >= (sheet.rowCount ?? 1_048_576) || range.endCol >= (sheet.columnCount ?? 16_384)) throw new RangeError('移位选区超出工作表。')
  if (edit.kind === 'insert' && selections.some((first, index) => selections.slice(index + 1).some(second => sheetRangesIntersect(first, second)))) {
    throw new RangeError('选区相互重叠，无法插入并移动单元格。')
  }
  const scope = sheetCellShiftAffectedRange(edit)
  assertSheetRangesEditable(sheet, [scope], userId)
  const merges = shiftedMerges(sheet, edit)
  // A dismantled merge may extend outside the band; its complete original extent needs permission.
  assertSheetRangesEditable(sheet, merges.dismantled, userId)
  const before = model.sheets.find(item => item.id === sheetId)!
  const cells: SpreadsheetSheet['cells'] = {}
  for (const [sourceKey, source] of Object.entries(before.cells)) {
    if (!Number.isSafeInteger(source.row) || !Number.isSafeInteger(source.col) || source.row < 0 || source.row >= 1_048_576
      || source.col < 0 || source.col >= 16_384 || sourceKey !== `${source.row}:${source.col}`) throw new RangeError('移位来源单元格坐标无效。')
    const point = mapSheetCellShiftPosition(source.row, source.col, edit)
    if (point) { const key = `${point.row}:${point.col}`; cells[key] = { ...cloneModel(source), id: key, ...point } }
    if (edit.kind === 'insert' && sheetRangeContainsCell(range, source.row, source.col)) {
      const blank = styleOnlyCell(source)
      if (blank) cells[blank.id] = blank
    }
  }
  sheet.cells = cells; sheet.mergedRanges = merges.ranges
  if (before.rangeThemeRules) sheet.rangeThemeRules = planSheetRangeThemeCellShift(before.rangeThemeRules, edit)
  transformWorkbookCellShiftResources(workbook, model, sheetId, edit)
  // Axis geometry, manual hides, outline and freeze remain axis properties; only the necessary canvas capacity grows.
  let rows = sheet.rowCount ?? 1_048_576, columns = sheet.columnCount ?? 16_384
  const includeRange = (item: CellRange): void => {
    assertSheetCommandRanges([item])
    rows = Math.max(rows, item.endRow + 1); columns = Math.max(columns, item.endCol + 1)
  }
  const includePoint = (row: number, col: number): void => includeRange({ startRow: row, endRow: row, startCol: col, endCol: col })
  for (const image of sheet.images ?? []) includePoint(image.anchor.row, image.anchor.column)
  for (const cell of Object.values(sheet.cells)) { rows = Math.max(rows, cell.row + 1); columns = Math.max(columns, cell.col + 1) }
  for (const item of [...sheet.mergedRanges, ...(sheet.rangeThemeRules ?? []).map(rule => rule.range), ...(sheet.autoFilter ? [sheet.autoFilter.range] : []),
    ...(sheet.conditionalFormats ?? []).flatMap(rule => rule.ranges), ...(sheet.dataValidations ?? []).flatMap(rule => rule.ranges),
    ...(workbook.tables ?? []).filter(table => table.sheetId === sheetId).map(table => table.range)]) {
    includeRange(item)
  }
  for (const chart of workbook.charts ?? []) if ((chart.source.sheetId ?? chart.sheetId) === sheetId) includeRange(chart.source.range)
  for (const pivot of workbook.pivotTables ?? []) {
    if (pivot.source.subUnitId === sheetId && (!workbook.id || pivot.source.unitId === workbook.id)) includeRange(pivot.source.range)
    if (pivot.anchor.subUnitId !== sheetId || workbook.id && pivot.anchor.unitId !== workbook.id) continue
    includePoint(pivot.anchor.row, pivot.anchor.col)
    if (pivot.outputRange) includeRange(pivot.outputRange)
    for (const key of pivot.generatedCellKeys ?? []) {
      const matched = /^(\d+):(\d+)$/.exec(key)
      if (!matched) throw new TypeError('透视输出坐标无效。')
      includePoint(Number(matched[1]), Number(matched[2]))
    }
  }
  for (const slicer of workbook.slicers ?? []) {
    if (slicer.hostSheetId !== sheetId || workbook.id && slicer.unitId !== workbook.id || slicer.drawing?.anchorType === 'none') continue
    includePoint(slicer.position.row, slicer.position.column)
    if (slicer.drawing) {
      includePoint(slicer.drawing.from.row, slicer.drawing.from.column)
      if (slicer.drawing.anchorType === 'both') includePoint(slicer.drawing.to.row, slicer.drawing.to.column)
    } else if (!slicer.layout) includeRange({ startRow: slicer.position.row, startCol: slicer.position.column,
      endRow: slicer.position.row + slicer.size.rows - 1, endCol: slicer.position.column + slicer.size.columns - 1 })
  }
  sheet.rowCount = rows; sheet.columnCount = columns
  refreshWorkbookCellStyles(workbook, undefined, false)
  for (const original of model.sheets) assertSheetRangeChangedValuesValid(original, workbook, original.id, { userId, region })
  const affected = assertSheetClipboardWorkbookChangesEditable(model, workbook, userId)
  affected[sheetId] = compactSheetClipboardRanges([scope, ...merges.dismantled, ...(affected[sheetId] ?? [])])
  return { workbook, range: { ...range }, edit, affected, mergeRanges: merges.dismantled, requiresConfirmation: Boolean(merges.dismantled.length && !dismantlePartialMerges) }
}
