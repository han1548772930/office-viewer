import type { CellRange } from '../../domain/model.ts'
import { canEditProtectedRanges, canEditWorksheetProtection } from '../sheet-protection/range-protection.ts'
import type { SpreadsheetSheet } from '../../domain/model.ts'

export type SheetCommandRange = CellRange | { startRow: number; startColumn: number; endRow: number; endColumn: number }
export interface SheetRangeCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  range?: SheetCommandRange
  ranges?: readonly SheetCommandRange[]
  selections?: readonly SheetCommandRange[]
}

export function normalizeSheetCommandRange(range: SheetCommandRange): CellRange {
  if (!range || typeof range !== 'object' || Array.isArray(range)) throw new RangeError('单元格范围无效。')
  const normalized = { startRow: range.startRow, endRow: range.endRow,
    startCol: 'startCol' in range ? range.startCol : range.startColumn,
    endCol: 'endCol' in range ? range.endCol : range.endColumn }
  assertSheetCommandRanges([normalized])
  return normalized
}
export function sheetRangeCommandInputRanges(input: SheetRangeCommandInput | undefined): CellRange[] | undefined {
  if (input === undefined) return undefined
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('范围命令参数无效。')
  for (const id of [input.unitId, input.sheetId, input.subUnitId]) {
    if (id !== undefined && (typeof id !== 'string' || !id.trim())) throw new TypeError('工作簿或工作表标识无效。')
  }
  const ranges = input.ranges !== undefined ? input.ranges : input.selections !== undefined ? input.selections
    : input.range !== undefined ? [input.range] : undefined
  if (ranges !== undefined && !Array.isArray(ranges)) throw new TypeError('单元格范围必须是数组。')
  return ranges?.map(normalizeSheetCommandRange)
}
export function assertSheetCommandRanges(ranges: readonly CellRange[]): void {
  for (const range of ranges) {
    if (!range || ![range.startRow, range.startCol, range.endRow, range.endCol].every(Number.isSafeInteger)
      || range.startRow < 0 || range.startCol < 0 || range.endRow < range.startRow || range.endCol < range.startCol
      || range.endRow >= 1_048_576 || range.endCol >= 16_384) throw new RangeError('单元格范围无效。')
  }
}
export function assertSheetRangesEditable(
  sheet: Pick<SpreadsheetSheet, 'worksheetProtection' | 'rangeProtectionRules'>,
  ranges: readonly CellRange[], userId: string,
): void {
  assertSheetCommandRanges(ranges)
  if (!canEditWorksheetProtection(sheet, userId) || !canEditProtectedRanges(sheet, ranges, userId)) {
    throw new Error('目标区域受保护，当前用户没有编辑权限。')
  }
}
export function sheetRangeContainsCell(range: CellRange, row: number, col: number): boolean {
  return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
}
export function sheetRangesIntersect(left: CellRange, right: CellRange): boolean {
  return left.startRow <= right.endRow && left.endRow >= right.startRow && left.startCol <= right.endCol && left.endCol >= right.startCol
}
export function sheetRangesEqual(left: CellRange, right: CellRange): boolean {
  return left.startRow === right.startRow && left.endRow === right.endRow && left.startCol === right.startCol && left.endCol === right.endCol
}

/** Four disjoint rectangles preserve the part of a resource outside the cleared range. */
export function subtractSheetCommandRange(source: CellRange, removed: CellRange): CellRange[] {
  if (!sheetRangesIntersect(source, removed)) return [{ ...source }]
  const top = Math.max(source.startRow, removed.startRow), bottom = Math.min(source.endRow, removed.endRow)
  const left = Math.max(source.startCol, removed.startCol), right = Math.min(source.endCol, removed.endCol)
  const ranges: CellRange[] = []
  if (source.startRow < top) ranges.push({ ...source, endRow: top - 1 })
  if (bottom < source.endRow) ranges.push({ ...source, startRow: bottom + 1 })
  if (source.startCol < left) ranges.push({ startRow: top, endRow: bottom, startCol: source.startCol, endCol: left - 1 })
  if (right < source.endCol) ranges.push({ startRow: top, endRow: bottom, startCol: right + 1, endCol: source.endCol })
  return ranges
}

/** Filtered rows are excluded; manual and outline hiding do not change the content scope. */
export function sheetCommandRangesWithoutFilteredRows(ranges: readonly CellRange[], filteredRows: readonly number[]): CellRange[] {
  const filtered = [...new Set(filteredRows)].sort((a, b) => a - b), result: CellRange[] = []
  for (const range of ranges) {
    let start = range.startRow
    for (const row of filtered) {
      if (row < start) continue
      if (row > range.endRow) break
      if (start < row) result.push({ ...range, startRow: start, endRow: row - 1 })
      start = row + 1
    }
    if (start <= range.endRow) result.push({ ...range, startRow: start })
  }
  return result
}

interface SheetRangeIndexNode {
  center: number
  crossing: Array<{ range: CellRange; index: number }>
  left?: SheetRangeIndexNode
  right?: SheetRangeIndexNode
}
function rangeIndex(entries: Array<{ range: CellRange; index: number }>): SheetRangeIndexNode | undefined {
  if (!entries.length) return undefined
  const ordered = [...entries].sort((left, right) => left.range.startRow - right.range.startRow)
  const center = ordered[Math.floor(ordered.length / 2)]!.range.startRow
  const crossing: typeof entries = [], left: typeof entries = [], right: typeof entries = []
  for (const entry of ordered) {
    if (entry.range.endRow < center) left.push(entry)
    else if (entry.range.startRow > center) right.push(entry)
    else crossing.push(entry)
  }
  crossing.sort((first, second) => first.range.startCol - second.range.startCol)
  for (let index = 1; index < crossing.length; index++) {
    if (crossing[index - 1]!.range.endCol >= crossing[index]!.range.startCol) throw new RangeError('合并区域不能交叉。')
  }
  return { center, crossing, left: rangeIndex(left), right: rangeIndex(right) }
}
function indexedRangeAt(node: SheetRangeIndexNode | undefined, row: number, col: number): number | undefined {
  if (!node) return undefined
  let low = 0, high = node.crossing.length
  while (low < high) {
    const middle = Math.floor((low + high) / 2)
    if (node.crossing[middle]!.range.startCol <= col) low = middle + 1
    else high = middle
  }
  const candidate = node.crossing[low - 1]
  if (candidate && sheetRangeContainsCell(candidate.range, row, col)) return candidate.index
  return indexedRangeAt(row < node.center ? node.left : node.right, row, col)
}
/** Partition sparse cells once into disjoint merge rectangles, without scanning the whole sheet for each merge. */
export function partitionSheetCommandCells<C extends { row: number; col: number }>(ranges: readonly CellRange[], cells: readonly C[]): C[][] {
  const root = rangeIndex(ranges.map((range, index) => ({ range, index })))
  const groups: C[][] = ranges.map(() => [])
  for (const cell of cells) {
    const index = indexedRangeAt(root, cell.row, cell.col)
    if (index !== undefined) groups[index]!.push(cell)
  }
  return groups
}
