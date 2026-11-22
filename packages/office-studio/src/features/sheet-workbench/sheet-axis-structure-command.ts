import { createCell, type CellRange, type SpreadsheetCell, type SpreadsheetSheet, type WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { sheetColumnWidth, sheetRowHeight } from '../../domain/sheet-axis.ts'
import { transformSheetAxisDimensions } from '../../domain/sheet-axis-transform.ts'
import { assertSheetAxisEdit, mergeSheetAxisIntervals, sheetAxisEditSegments, sheetAxisMaximum, sheetRangeAxisInterval,
  sheetRangeWithAxisInterval, transformSheetAxisRange, type SheetAxis, type SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'
import { translateFormulaReferenceAreas } from '../formula/formula-reference-areas.ts'
import { assertSheetCommandRanges, assertSheetRangesEditable, normalizeSheetCommandRange, sheetCommandRangesWithoutFilteredRows, sheetRangeCommandInputRanges,
  type SheetCommandRange, type SheetRangeCommandInput } from './sheet-range-command.ts'
import { transformWorkbookAxisResources } from './sheet-axis-structure-resources.ts'

export const SHEET_AXIS_STRUCTURE_COMMAND_IDS = [
  'sheet.command.insert-row', 'sheet.command.insert-row-by-range', 'sheet.command.insert-row-before', 'sheet.command.insert-row-after',
  'sheet.command.insert-col', 'sheet.command.insert-column', 'sheet.command.insert-col-by-range', 'sheet.command.insert-col-before', 'sheet.command.insert-col-after',
  'sheet.command.insert-multi-rows-above', 'sheet.command.insert-multi-rows-after',
  'sheet.command.insert-multi-cols-before', 'sheet.command.insert-multi-cols-right',
  'sheet.command.remove-row', 'sheet.command.remove-row-by-range', 'sheet.command.delete-row',
  'sheet.command.remove-col', 'sheet.command.remove-col-by-range', 'sheet.command.delete-column',
  'sheet.command.move-rows', 'sheet.command.move-cols', 'sheet.command.move-columns',
  'sheet.command.copy-rows', 'sheet.command.copy-cols', 'sheet.command.copy-columns',
] as const
export interface SheetAxisStructureCommandInput extends SheetRangeCommandInput {
  index?: number
  start?: number
  count?: number
  value?: number
  direction?: 'before' | 'after' | 'up' | 'down' | 'left' | 'right' | 0 | 1 | 2 | 3
  fromRange?: SheetCommandRange
  toRange?: SheetCommandRange
  destinationIndex?: number
  destination?: number
  expectedRevision?: number
  includeFiltered?: boolean
}
export interface SheetAxisStructureAction { axis: SheetAxis; kind: SheetAxisEdit['kind']; position?: 'before' | 'after' }
export interface SheetAxisStructurePlan { workbook: WorkbookModel; edits: SheetAxisEdit[]; selection: CellRange; affected: CellRange[] }
const MAX_NEW_STYLED_CELLS = 250_000

export function sheetAxisStructureActionForCommand(id: string): SheetAxisStructureAction | undefined {
  if (!(SHEET_AXIS_STRUCTURE_COMMAND_IDS as readonly string[]).includes(id)) return undefined
  const axis: SheetAxis = /(?:-row|-rows)(?:-|$)/.test(id) ? 'row' : 'column'
  const kind = id.includes('.insert-') ? 'insert' : id.includes('.move-') ? 'move' : id.includes('.copy-') ? 'copy' : 'delete'
  return { axis, kind, ...(kind === 'insert' && /-(before|after|above|right)$/.test(id)
    ? { position: /-(after|right)$/.test(id) ? 'after' : 'before' } : {}) }
}
export function sheetAxisStructureCommandInput(value: unknown): SheetAxisStructureCommandInput {
  if (value === undefined) return {}
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('行列结构命令参数无效。')
  const input = value as SheetAxisStructureCommandInput
  const ranges = sheetRangeCommandInputRanges(input)
  for (const [name, number] of Object.entries({ index: input.index, start: input.start, count: input.count, value: input.value,
    destinationIndex: input.destinationIndex, destination: input.destination, expectedRevision: input.expectedRevision })) {
    if (number !== undefined && (!Number.isSafeInteger(number) || number < 0 || (name === 'count' || name === 'value') && number === 0)) {
      throw new RangeError(`行列结构参数 ${name} 无效。`)
    }
  }
  if (input.direction !== undefined && !(['before', 'after', 'up', 'down', 'left', 'right', 0, 1, 2, 3] as readonly unknown[]).includes(input.direction)) throw new TypeError('插入方向无效。')
  if (input.includeFiltered !== undefined && typeof input.includeFiltered !== 'boolean') throw new TypeError('筛选行选项无效。')
  if (input.fromRange !== undefined) normalizeSheetCommandRange(input.fromRange)
  if (input.toRange !== undefined) normalizeSheetCommandRange(input.toRange)
  return { ...input, ...(ranges === undefined ? {} : { ranges }) }
}
export function sheetAxisStructureEdits(sheet: SpreadsheetSheet, ranges: readonly CellRange[], action: SheetAxisStructureAction, input: SheetAxisStructureCommandInput,
  filteredRows: readonly number[] = sheet.autoFilter?.filteredRows ?? []): SheetAxisEdit[] {
  assertSheetCommandRanges(ranges)
  const explicitStart = input.start ?? input.index, sourceRanges = input.fromRange ? [normalizeSheetCommandRange(input.fromRange)] : [...ranges]
  const currentCount = (action.axis === 'row' ? sheet.rowCount : sheet.columnCount) ?? sheetAxisMaximum(action.axis)
  if (!Number.isSafeInteger(currentCount) || currentCount < 1 || currentCount > sheetAxisMaximum(action.axis)) throw new RangeError('工作表行列总数无效。')
  let intervals = mergeSheetAxisIntervals(sourceRanges.map(range => sheetRangeAxisInterval(range, action.axis)))
  if (explicitStart !== undefined) intervals = [{ start: explicitStart, end: explicitStart + (input.count ?? input.value ?? 1) - 1 }]
  if (!intervals.length) return []
  if (action.kind !== 'delete' && intervals.length !== 1) throw new Error('插入、移动和复制需要连续的行列选区。')
  const first = intervals[0]!, count = input.count ?? input.value ?? first.end - first.start + 1
  if (action.kind === 'insert') {
    const after = input.direction === 'after' || input.direction === 'down' || input.direction === 'right' || input.direction === 1 || input.direction === 2
      || input.direction === undefined && action.position === 'after'
    // By-range commands already describe the new interval; their direction
    // chooses the style anchor rather than moving the insertion boundary.
    const byRange = input.range !== undefined && (input.start === undefined && input.index === undefined)
      && action.position === undefined
    const start = after && !byRange ? (explicitStart ?? first.end) + 1 : first.start
    if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(count) || count <= 0) throw new RangeError('插入位置或数量无效。')
    if (start > currentCount) throw new RangeError('插入位置超出工作表。')
    if (start + count > sheetAxisMaximum(action.axis)) throw new RangeError('插入结果超出 Office 行列上限。')
    return [{ axis: action.axis, kind: 'insert', start, count }]
  }
  if (intervals.some(interval => interval.start < 0 || interval.end >= currentCount)) throw new RangeError('源行列选区超出工作表。')
  if (action.kind === 'move' || action.kind === 'copy') {
    const destination = input.destinationIndex ?? input.destination ?? (input.toRange ? sheetRangeAxisInterval(normalizeSheetCommandRange(input.toRange), action.axis).start : undefined)
    if (destination === undefined || destination > currentCount) throw new RangeError('请指定有效的目标行列位置。')
    const edit: SheetAxisEdit = { axis: action.axis, kind: action.kind, start: first.start, count, destination }
    assertSheetAxisEdit(edit)
    if (edit.start + edit.count > currentCount) throw new RangeError('源行列选区超出工作表。')
    return action.kind === 'move' && (destination === first.start || destination === first.start + count) ? [] : [edit]
  }
  if (action.axis === 'row' && input.includeFiltered !== true) {
    intervals = mergeSheetAxisIntervals(sheetCommandRangesWithoutFilteredRows(intervals.map(interval => ({ startRow: interval.start, endRow: interval.end,
      startCol: 0, endCol: (sheet.columnCount ?? 16_384) - 1 })), filteredRows).map(range => ({ start: range.startRow, end: range.endRow })))
  }
  const removed = intervals.reduce((sum, interval) => sum + interval.end - interval.start + 1, 0)
  if (removed >= currentCount) throw new Error(action.axis === 'row' ? '工作表至少需要保留一行。' : '工作表至少需要保留一列。')
  return intervals.reverse().map(interval => ({ axis: action.axis, kind: 'delete', start: interval.start, count: interval.end - interval.start + 1 }))
}

export function sheetAxisStructureAffectedRange(sheet: SpreadsheetSheet, edit: SheetAxisEdit): CellRange {
  const source = edit.kind === 'move' ? Math.min(edit.start, edit.destination) : edit.kind === 'copy' ? Math.min(edit.start, edit.destination) : edit.start
  const end = edit.kind === 'move' ? Math.max(edit.start + edit.count - 1, edit.destination - 1) : sheetAxisMaximum(edit.axis) - 1
  return edit.axis === 'row' ? { startRow: source, endRow: end, startCol: 0, endCol: (sheet.columnCount ?? 16_384) - 1 }
    : { startRow: 0, endRow: (sheet.rowCount ?? 1_048_576) - 1, startCol: source, endCol: end }
}
function prepareMerges(sheet: SpreadsheetSheet, edit: SheetAxisEdit): CellRange[] {
  const result: CellRange[] = []
  for (const merge of sheet.mergedRanges) {
    const axis = sheetRangeAxisInterval(merge, edit.axis)
    if (edit.kind === 'move' || edit.kind === 'copy') {
      const sourceIntersects = axis.start < edit.start + edit.count && axis.end >= edit.start
      if (sourceIntersects && (axis.start < edit.start || axis.end >= edit.start + edit.count)) throw new Error('源行列只包含部分合并单元格，请一起选择整个合并区域。')
      if (axis.start < edit.destination && axis.end >= edit.destination) throw new Error('目标位置跨过合并单元格，请选择合并区域边界。')
    }
    const ranges = transformSheetAxisRange(merge, edit)
    if (ranges.length > 1) throw new Error('移动会拆开合并单元格，请调整选区。')
    for (const range of ranges) if (range.startRow !== range.endRow || range.startCol !== range.endCol) result.push(range)
    if (edit.kind === 'copy' && axis.start >= edit.start && axis.end < edit.start + edit.count) {
      const range = sheetRangeWithAxisInterval(merge, edit.axis, { start: edit.destination + axis.start - edit.start, end: edit.destination + axis.end - edit.start })
      result.push(range)
    }
  }
  return result
}
function inheritedStyleCells(sheet: SpreadsheetSheet, edit: SheetAxisEdit): SpreadsheetCell[] {
  if (edit.kind !== 'insert') return []
  // Fixed-version copyRangeStyles uses the preceding row/column for both
  // before and after insertion; the dimension anchor is a separate contract.
  const index = edit.start - 1
  if (index < 0) return []
  return Object.values(sheet.cells).filter(cell => (edit.axis === 'row' ? cell.row : cell.col) === index)
}
function transformCells(sheet: SpreadsheetSheet, edit: SheetAxisEdit, styles: readonly SpreadsheetCell[]) {
  const segments = sheetAxisEditSegments(edit), maximum = sheetAxisMaximum(edit.axis), cells: SpreadsheetSheet['cells'] = {}, copies: Array<{ key: string; sourceRow: number; sourceColumn: number }> = []
  const position = (row: number, column: number, index: number) => edit.axis === 'row' ? { row: index, column } : { row, column: index }
  const map = (index: number) => { const segment = segments.find(item => index >= item.start && index <= item.end); return segment ? index + segment.offset : undefined }
  if (styles.length * edit.count > MAX_NEW_STYLED_CELLS) throw new Error('插入会生成过多格式单元格，请分批插入行列。')
  if (edit.kind === 'copy' && edit.destination + edit.count > maximum) throw new RangeError('复制结果超出 Office 行列上限。')
  for (const cell of Object.values(sheet.cells)) {
    const index = edit.axis === 'row' ? cell.row : cell.col, mapped = map(index)
    if (mapped !== undefined && mapped >= maximum) throw new Error('插入会把已有单元格移出 Office 行列上限。')
    if (mapped !== undefined) {
      const next = position(cell.row, cell.col, mapped), key = `${next.row}:${next.column}`
      cells[key] = { ...cell, id: key, row: next.row, col: next.column }
    }
    if (edit.kind === 'copy' && index >= edit.start && index < edit.start + edit.count) {
      const next = position(cell.row, cell.col, edit.destination + index - edit.start), key = `${next.row}:${next.column}`
      cells[key] = { ...cloneModel(cell), id: key, row: next.row, col: next.column }
      const source = position(cell.row, cell.col, mapped!)
      copies.push({ key, sourceRow: source.row, sourceColumn: source.column })
    }
  }
  for (const source of styles) for (let offset = 0; offset < edit.count; offset++) {
    const next = position(source.row, source.col, edit.start + offset), cell = createCell(next.row, next.column, null)
    cell.style = cloneModel(source.style)
    cells[cell.id] = cell
  }
  return { cells, copies }
}

/** Detach the full plan before any document/history mutation; all resources share each edit's original coordinates. */
export function planWorkbookAxisStructureCommand(model: WorkbookModel, sheetId: string, edits: readonly SheetAxisEdit[], userId = 'local-user',
  insertionStyle: 'before' | 'after' = 'before'): SheetAxisStructurePlan {
  const workbook = cloneModel(model), affected: CellRange[] = []
  let selection: CellRange = { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }
  for (const edit of edits) {
    assertSheetAxisEdit(edit)
    const sheet = workbook.sheets.find(item => item.id === sheetId)
    if (!sheet) throw new Error('找不到目标工作表。')
    const currentCount = (edit.axis === 'row' ? sheet.rowCount : sheet.columnCount) ?? sheetAxisMaximum(edit.axis)
    if (edit.kind === 'insert' ? edit.start > currentCount : edit.start + edit.count > currentCount) throw new RangeError('源行列选区超出工作表。')
    if ((edit.kind === 'move' || edit.kind === 'copy') && edit.destination > currentCount) throw new RangeError('行列目标位置超出工作表。')
    if (edit.kind === 'insert' && edit.start + edit.count > sheetAxisMaximum(edit.axis)) throw new RangeError('插入结果超出 Office 行列上限。')
    const original = cloneModel(workbook), before = original.sheets.find(item => item.id === sheetId)!
    const scope = sheetAxisStructureAffectedRange(before, edit)
    assertSheetRangesEditable(before, [scope], userId)
    affected.push(scope)
    const merges = prepareMerges(before, edit), styles = inheritedStyleCells(before, edit)
    const transformed = transformCells(before, edit, styles)
    // Materialize the legacy wider first column before permuting it. Its
    // implicit position must not grant the next column that same width.
    if (sheet.defaultColumnWidth === undefined) { sheet.columnWidths[0] = sheetColumnWidth(sheet, 0); sheet.defaultColumnWidth = 112 }
    const size = edit.kind === 'insert' ? edit.axis === 'row' ? sheetRowHeight(before, Math.max(0, Math.min((before.rowCount ?? 1_048_576) - 1,
      edit.start - (insertionStyle === 'after' ? 1 : 0)))) : sheetColumnWidth(before, Math.max(0, Math.min((before.columnCount ?? 16_384) - 1,
      edit.start - (insertionStyle === 'after' ? 1 : 0)))) : undefined
    Object.assign(sheet, transformSheetAxisDimensions(sheet, { ...edit, inheritedSize: size }))
    sheet.cells = transformed.cells
    sheet.mergedRanges = merges
    const count = edit.kind === 'insert' || edit.kind === 'copy' ? Math.min(sheetAxisMaximum(edit.axis), currentCount + edit.count)
      : edit.kind === 'delete' ? currentCount - edit.count : currentCount
    if (count < 1) throw new Error('工作表至少需要保留一行和一列。')
    if (edit.axis === 'row') sheet.rowCount = count; else sheet.columnCount = count
    const hidden = edit.axis === 'row' ? before.hiddenRows : before.hiddenColumns, segments = sheetAxisEditSegments(edit)
    const mapHidden = (index: number) => { const segment = segments.find(item => index >= item.start && index <= item.end); return segment ? index + segment.offset : undefined }
    const mappedHidden = hidden?.flatMap(index => {
      const mapped = mapHidden(index), values = mapped === undefined || mapped >= count ? [] : [mapped]
      if (edit.kind === 'copy' && index >= edit.start && index < edit.start + edit.count) values.push(edit.destination + index - edit.start)
      return values
    })
    const hiddenResult = mappedHidden ? [...new Set(mappedHidden)].sort((left, right) => left - right) : undefined
    if (edit.axis === 'row') sheet.hiddenRows = hiddenResult; else sheet.hiddenColumns = hiddenResult
    transformWorkbookAxisResources(workbook, original, sheet, edit)
    for (const copy of transformed.copies) {
      const cell = sheet.cells[copy.key]!, rowDelta = cell.row - copy.sourceRow, columnDelta = cell.col - copy.sourceColumn
      if (cell.formula) cell.formula = translateFormulaReferenceAreas(cell.formula, rowDelta, columnDelta)
      if (cell.sparkline) cell.sparkline.sourceRange = translateFormulaReferenceAreas(cell.sparkline.sourceRange, rowDelta, columnDelta)
    }
    const start = edit.kind === 'move' ? edit.destination - (edit.destination > edit.start ? edit.count : 0)
      : edit.kind === 'copy' ? edit.destination : Math.min(edit.start, count - 1)
    const end = edit.kind === 'delete' ? start : start + edit.count - 1
    selection = edit.axis === 'row' ? { startRow: start, endRow: end, startCol: 0, endCol: (sheet.columnCount ?? 16_384) - 1 }
      : { startRow: 0, endRow: (sheet.rowCount ?? 1_048_576) - 1, startCol: start, endCol: end }
  }
  return { workbook, edits: edits.map(edit => ({ ...edit })), selection, affected }
}
