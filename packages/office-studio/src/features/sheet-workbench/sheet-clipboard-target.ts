import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { assertSheetClipboardArea, expandSheetClipboardRange } from './sheet-clipboard-data.ts'
import { assertSheetCommandRanges, sheetRangeContainsCell, sheetRangesIntersect } from './sheet-range-command.ts'
import type { SheetClipboardSnapshot } from './sheet-clipboard-types.ts'

export interface SheetClipboardTargetGeometry {
  rows: number[]
  columns: number[]
  ranges: CellRange[]
  merges: CellRange[]
  /** A scalar pasted into a single existing merge changes its anchor content. */
  contentOnlyMerge?: CellRange
}

export function sheetClipboardRowsToRanges(rows: readonly number[], startCol: number, endCol: number): CellRange[] {
  const ranges: CellRange[] = []
  for (const row of rows) {
    const previous = ranges.at(-1)
    if (previous?.endRow === row - 1) previous.endRow = row
    else ranges.push({ startRow: row, endRow: row, startCol, endCol })
  }
  return ranges
}

function expandedRows(initial: readonly number[], count: number, filtered: ReadonlySet<number>): number[] {
  const rows = initial.slice(0, count)
  let row = (initial.at(-1) ?? -1) + 1
  while (rows.length < count && row < 1_048_576) { if (!filtered.has(row)) rows.push(row); row++ }
  if (rows.length !== count) throw new RangeError('粘贴后会超出工作表的最大行数。')
  return rows
}

function projectedMerges(snapshot: SheetClipboardSnapshot, rows: readonly number[], columns: readonly number[]): CellRange[] {
  const merges: CellRange[] = []
  for (let row = 0; row < rows.length; row += snapshot.rows.length) for (let col = 0; col < columns.length; col += snapshot.columns.length) {
    for (const merge of snapshot.merges) {
      const startRow = rows[row + merge.startRow], endRow = rows[row + merge.endRow]
      const startCol = columns[col + merge.startCol], endCol = columns[col + merge.endCol]
      if (startRow === undefined || endRow === undefined || startCol === undefined || endCol === undefined) throw new RangeError('粘贴范围不能截断源合并单元格。')
      if (endRow - startRow !== merge.endRow - merge.startRow) throw new RangeError('合并单元格不能跨过目标筛选行粘贴。')
      merges.push({ startRow, endRow, startCol, endCol })
    }
  }
  return merges
}

/** Selection repetition and filtering produce actual ranges, never the bounds of separate selections. */
export function createSheetClipboardTargetGeometry(
  workbook: WorkbookModel, sheet: SpreadsheetSheet, selected: CellRange, snapshot: SheetClipboardSnapshot,
): SheetClipboardTargetGeometry {
  assertSheetCommandRanges([selected])
  const expanded = expandSheetClipboardRange(sheet, selected)
  const filtered = new Set([...(sheet.autoFilter?.filteredRows ?? []), ...(workbook.tables ?? [])
    .filter(table => table.sheetId === sheet.id).flatMap(table => table.autoFilter?.filteredRows ?? [])])
  if (snapshot.copyType === 'cut') {
    const rows = Array.from({ length: snapshot.rows.length }, (_, index) => selected.startRow + index)
    const columns = Array.from({ length: snapshot.columns.length }, (_, index) => selected.startCol + index)
    const ranges = sheetClipboardRowsToRanges(rows, columns[0]!, columns.at(-1)!)
    assertSheetCommandRanges(ranges)
    return { rows, columns, ranges, merges: projectedMerges(snapshot, rows, columns) }
  }
  assertSheetClipboardArea(expanded.endRow - expanded.startRow + 1, expanded.endCol - expanded.startCol + 1)
  const initial: number[] = []
  for (let row = expanded.startRow; row <= expanded.endRow; row++) if (!filtered.has(row)) initial.push(row)
  if (!initial.length) throw new RangeError('粘贴目标没有可见行。')
  const width = expanded.endCol - expanded.startCol + 1
  const singleMerge = sheet.mergedRanges.find(merge => merge.startRow === expanded.startRow && merge.endRow === expanded.endRow
    && merge.startCol === expanded.startCol && merge.endCol === expanded.endCol)
  if (singleMerge && snapshot.rows.length === 1 && snapshot.columns.length === 1 && !snapshot.merges.length) {
    if (initial.length !== singleMerge.endRow - singleMerge.startRow + 1) throw new RangeError('合并单元格不能跨过目标筛选行粘贴。')
    return { rows: [singleMerge.startRow], columns: [singleMerge.startCol], ranges: [{ ...singleMerge }], merges: [], contentOnlyMerge: { ...singleMerge } }
  }
  const selectedRanges = sheetClipboardRowsToRanges(initial, expanded.startCol, expanded.endCol)
  const intersectsMerge = sheet.mergedRanges.some(merge => selectedRanges.some(range => sheetRangesIntersect(merge, range)))
  const repeat = !intersectsMerge && initial.length % snapshot.rows.length === 0 && width % snapshot.columns.length === 0
  const rows = expandedRows(initial, repeat ? initial.length : snapshot.rows.length, filtered)
  const columns = Array.from({ length: repeat ? width : snapshot.columns.length }, (_, index) => expanded.startCol + index)
  const ranges = sheetClipboardRowsToRanges(rows, columns[0]!, columns.at(-1)!)
  assertSheetClipboardArea(rows.length, columns.length); assertSheetCommandRanges(ranges)
  for (const merge of sheet.mergedRanges) {
    if (!ranges.some(range => sheetRangesIntersect(merge, range))) continue
    if (!ranges.some(range => sheetRangeContainsCell(range, merge.startRow, merge.startCol)
      && sheetRangeContainsCell(range, merge.endRow, merge.endCol))) throw new RangeError('粘贴范围会覆盖部分合并单元格，请重新选择完整区域。')
  }
  return { rows, columns, ranges, merges: projectedMerges(snapshot, rows, columns) }
}
