import type { CellRange, SpreadsheetCell, SpreadsheetSheet } from '../../domain/model.ts'
import {
  assertSheetCommandRanges, partitionSheetCommandCells, sheetRangeContainsCell, sheetRangesIntersect,
} from '../sheet-workbench/sheet-range-command.ts'

export interface SheetDataRangeDirections { up?: boolean; down?: boolean; left?: boolean; right?: boolean }
export const SHEET_DATA_RANGE_ALL_DIRECTIONS: Readonly<SheetDataRangeDirections> = Object.freeze({ up: true, down: true, left: true, right: true })

/** Continuous data regions use stored values/body presence, never display text or formula identity. */
export function sheetCellHasData(cell: SpreadsheetCell | undefined): boolean {
  return !!cell && (cell.value !== undefined && cell.value !== null && cell.value !== ''
    || !!cell.richText?.length || cell.metadata?.publicBody !== undefined)
}

function unionRanges(left: CellRange, right: CellRange): CellRange {
  return { startRow: Math.min(left.startRow, right.startRow), endRow: Math.max(left.endRow, right.endRow),
    startCol: Math.min(left.startCol, right.startCol), endCol: Math.max(left.endCol, right.endCol) }
}
function hasIndexedValue(values: readonly number[] | undefined, start: number, end: number): boolean {
  if (!values) return false
  let low = 0, high = values.length
  while (low < high) {
    const middle = Math.floor((low + high) / 2)
    if (values[middle]! < start) low = middle + 1
    else high = middle
  }
  return low < values.length && values[low]! <= end
}
function addIndex(index: Map<number, number[]>, key: number, value: number): void {
  const values = index.get(key)
  if (values) values.push(value)
  else index.set(key, [value])
}

/**
 * Matches the fixed Sheets corner -> up/down/left/right expansion order without a dense cell matrix.
 * Merge interiors borrow their anchor's content; only spans encountered by an eligible edge join
 * the final union. Blank adjacent merges do not independently extend a data region.
 */
export function expandSheetDataRange(
  sheet: SpreadsheetSheet, source: CellRange, directions: SheetDataRangeDirections = SHEET_DATA_RANGE_ALL_DIRECTIONS,
): CellRange {
  assertSheetCommandRanges([source])
  if (!directions || typeof directions !== 'object' || Array.isArray(directions)
    || Object.entries(directions).some(([key, value]) => !['up', 'down', 'left', 'right'].includes(key)
      || value !== undefined && typeof value !== 'boolean')) throw new TypeError('连续数据区域的扩展方向无效。')
  const maxRows = sheet.rowCount ?? 1_048_576, maxColumns = sheet.columnCount ?? 16_384
  if (!Number.isSafeInteger(maxRows) || !Number.isSafeInteger(maxColumns) || maxRows <= source.endRow || maxColumns <= source.endCol
    || maxRows > 1_048_576 || maxColumns > 16_384) throw new RangeError('数据区域超出当前工作表范围。')
  const merges = sheet.mergedRanges
  const cells = Object.values(sheet.cells)
  const interior = new Set<SpreadsheetCell>()
  const groups = partitionSheetCommandCells(merges, cells)
  groups.forEach((group, index) => {
    const merge = merges[index]!
    for (const cell of group) if (cell.row !== merge.startRow || cell.col !== merge.startCol) interior.add(cell)
  })
  const rows = new Map<number, number[]>(), columns = new Map<number, number[]>()
  for (const cell of cells) {
    if (interior.has(cell) || !sheetCellHasData(cell)) continue
    addIndex(rows, cell.row, cell.col); addIndex(columns, cell.col, cell.row)
  }
  for (const values of [...rows.values(), ...columns.values()]) values.sort((left, right) => left - right)
  const contentSpans = new Set(merges.filter(merge => sheetCellHasData(sheet.cells[`${merge.startRow}:${merge.startCol}`])))
  const at = (row: number, col: number) => {
    const span = merges.find(merge => sheetRangeContainsCell(merge, row, col))
    return { span, hasValue: span ? contentSpans.has(span) : sheetCellHasData(sheet.cells[`${row}:${col}`]) }
  }
  const probe = (edge: CellRange, current: CellRange) => {
    let span: CellRange | undefined
    let hasValue = edge.startRow === edge.endRow
      ? hasIndexedValue(rows.get(edge.startRow), edge.startCol, edge.endCol)
      : hasIndexedValue(columns.get(edge.startCol), edge.startRow, edge.endRow)
    for (const merge of merges) {
      if (!sheetRangesIntersect(edge, merge)) continue
      span = span ? unionRanges(span, merge) : merge
      hasValue ||= contentSpans.has(merge)
    }
    return { hasValue, span: span && (hasValue || sheetRangesIntersect(current, span)) ? span : undefined }
  }
  let range = { ...source }, collected: CellRange | undefined
  for (;;) {
    const { startRow, endRow, startCol, endCol } = range
    let changed = false
    const corner = (row: number, col: number, vertical: 'startRow' | 'endRow', horizontal: 'startCol' | 'endCol') => {
      const cell = at(row, col)
      if (!cell.hasValue) return
      range[vertical] = cell.span?.[vertical] ?? row
      range[horizontal] = cell.span?.[horizontal] ?? col
      changed = true
    }
    if (directions.up && directions.left && startRow > 0 && startCol > 0) corner(startRow - 1, startCol - 1, 'startRow', 'startCol')
    if (directions.up && directions.right && startRow > 0 && endCol < maxColumns - 1) corner(startRow - 1, endCol + 1, 'startRow', 'endCol')
    if (directions.down && directions.left && endRow < maxRows - 1 && startCol > 0) corner(endRow + 1, startCol - 1, 'endRow', 'startCol')
    if (directions.down && directions.right && endRow < maxRows - 1 && endCol < maxColumns - 1) corner(endRow + 1, endCol + 1, 'endRow', 'endCol')
    if (changed) continue
    const edges: Array<{ enabled: boolean | undefined; edge: CellRange; coordinate: keyof CellRange; value: number }> = [
      { enabled: directions.up && range.startRow > 0, edge: { ...range, startRow: range.startRow - 1, endRow: range.startRow - 1 }, coordinate: 'startRow', value: range.startRow - 1 },
      { enabled: directions.down && range.endRow < maxRows - 1, edge: { ...range, startRow: range.endRow + 1, endRow: range.endRow + 1 }, coordinate: 'endRow', value: range.endRow + 1 },
      { enabled: directions.left && range.startCol > 0, edge: { ...range, startCol: range.startCol - 1, endCol: range.startCol - 1 }, coordinate: 'startCol', value: range.startCol - 1 },
      { enabled: directions.right && range.endCol < maxColumns - 1, edge: { ...range, startCol: range.endCol + 1, endCol: range.endCol + 1 }, coordinate: 'endCol', value: range.endCol + 1 },
    ]
    for (const edge of edges) {
      if (!edge.enabled) continue
      const result = probe(edge.edge, range)
      if (result.span) collected = collected ? unionRanges(collected, result.span) : { ...result.span }
      if (!result.hasValue) continue
      range[edge.coordinate] = edge.value
      changed = true
      break
    }
    if (!changed) return collected ? unionRanges(range, collected) : range
  }
}

/** Smart filter: one cell expands in four directions, a single row only down, otherwise keep it. */
export function prepareSmartSheetFilterRange(sheet: SpreadsheetSheet, source: CellRange): CellRange {
  const range = source.startRow === source.endRow
    ? expandSheetDataRange(sheet, source, source.startCol === source.endCol ? SHEET_DATA_RANGE_ALL_DIRECTIONS : { down: true })
    : { ...source }
  assertSheetCommandRanges([range])
  if (range.endRow <= range.startRow) throw new RangeError('请选择包含标题行和数据行的区域后再启用筛选。')
  return range
}
