import type { CellRange, SpreadsheetCell, SpreadsheetSheet } from '../../domain/model.ts'

export interface SheetFilterColumnEntry {
  row: number
  column: number
  rowSpan: number
  cell: SpreadsheetCell | undefined
}

/** Match Studio range capabilities: preserve disjoint merge scopes instead of widening to a bounding rectangle. */
export function sheetFilterPermissionRanges(sheet: Pick<SpreadsheetSheet, 'mergedRanges'>, range: CellRange | null | undefined): CellRange[] {
  if (!range) return []
  return [{ ...range }, ...sheet.mergedRanges.filter(merge => merge.startRow <= range.endRow && merge.endRow >= range.startRow
    && merge.startCol <= range.endCol && merge.endCol >= range.startCol).map(merge => ({ ...merge }))]
}

/** Match Worksheet.iterateByColumn(false, skipNonTopLeft), including a merge's complete rowSpan. */
export function* sheetFilterColumnEntries(sheet: SpreadsheetSheet, range: CellRange, column: number, skipNonTopLeft = false): Generator<SheetFilterColumnEntry> {
  if (!Number.isSafeInteger(column) || column < range.startCol || column > range.endCol) throw new RangeError('筛选列必须位于筛选区域内。')
  const merges = sheet.mergedRanges.filter(merge => column >= merge.startCol && column <= merge.endCol
    && merge.endRow > range.startRow && merge.startRow <= range.endRow).sort((left, right) => left.startRow - right.startRow)
  let mergeIndex = 0, row = range.startRow + 1
  while (row <= range.endRow) {
    while (mergeIndex < merges.length && merges[mergeIndex]!.endRow < row) mergeIndex++
    const merge = merges[mergeIndex]
    if (merge && merge.startRow <= row && row <= merge.endRow) {
      if (row === merge.startRow && (!skipNonTopLeft || column === merge.startCol)) yield { row,
        column: merge.startCol, rowSpan: merge.endRow - merge.startRow + 1, cell: sheet.cells[`${row}:${merge.startCol}`] }
      row = merge.endRow + 1
    } else {
      yield { row, column, rowSpan: 1, cell: sheet.cells[`${row}:${column}`] }
      row++
    }
  }
}
