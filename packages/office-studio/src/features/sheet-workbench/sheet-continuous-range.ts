import type { CellRange, SpreadsheetSheet } from '../../domain/model.ts'
import { sheetRangeContainsCell, sheetRangesIntersect } from './sheet-range-command.ts'

export function expandSheetContinuousRange(sheet: Pick<SpreadsheetSheet, 'cells' | 'mergedRanges' | 'rowCount' | 'columnCount'>,
  range: CellRange): CellRange {
  const cells = new Map(Object.values(sheet.cells).map(cell => [`${cell.row}:${cell.col}`, cell]))
  const rowCount = sheet.rowCount ?? 1_048_576, columnCount = sheet.columnCount ?? 16_384
  const result = { ...range }, spans = new Set<CellRange>()
  function cellAt(row: number, col: number) {
    const span = sheet.mergedRanges.find(merge => sheetRangeContainsCell(merge, row, col))
    const cell = cells.get(span ? `${span.startRow}:${span.startCol}` : `${row}:${col}`)
    return { span, hasValue: cell !== undefined && (cell.value !== null && cell.value !== '' || cell.richText !== undefined) }
  }
  function sideHasValue(axis: 'row' | 'column', position: number, start: number, end: number): boolean {
    const merges = new Set<CellRange>()
    let hasValue = false
    for (let index = start; index <= end; index += 1) {
      const cell = cellAt(axis === 'row' ? position : index, axis === 'column' ? position : index)
      hasValue ||= cell.hasValue
      if (cell.span) merges.add(cell.span)
      if (hasValue && !sheet.mergedRanges.length) break
    }
    for (const merge of merges) if (hasValue || sheetRangesIntersect(result, merge)) spans.add(merge)
    return hasValue
  }
  for (;;) {
    const before = { ...result }
    let cornerExpanded = false
    for (const [row, col, top, left] of [
      [before.startRow - 1, before.startCol - 1, true, true],
      [before.startRow - 1, before.endCol + 1, true, false],
      [before.endRow + 1, before.startCol - 1, false, true],
      [before.endRow + 1, before.endCol + 1, false, false],
    ] as const) {
      if (row < 0 || col < 0 || row >= rowCount || col >= columnCount) continue
      const cell = cellAt(row, col)
      if (!cell.hasValue) continue
      if (top) result.startRow = cell.span?.startRow ?? row
      if (!top) result.endRow = cell.span?.endRow ?? row
      if (left) result.startCol = cell.span?.startCol ?? col
      if (!left) result.endCol = cell.span?.endCol ?? col
      cornerExpanded = true
    }
    if (cornerExpanded) continue
    if (result.startRow > 0 && sideHasValue('row', result.startRow - 1, result.startCol, result.endCol)) {
      result.startRow -= 1
      continue
    }
    if (result.endRow < rowCount - 1 && sideHasValue('row', result.endRow + 1, result.startCol, result.endCol)) {
      result.endRow += 1
      continue
    }
    if (result.startCol > 0 && sideHasValue('column', result.startCol - 1, result.startRow, result.endRow)) {
      result.startCol -= 1
      continue
    }
    if (result.endCol < columnCount - 1 && sideHasValue('column', result.endCol + 1, result.startRow, result.endRow)) {
      result.endCol += 1
      continue
    }
    break
  }
  for (const span of spans) {
    result.startRow = Math.min(result.startRow, span.startRow)
    result.endRow = Math.max(result.endRow, span.endRow)
    result.startCol = Math.min(result.startCol, span.startCol)
    result.endCol = Math.max(result.endCol, span.endCol)
  }
  return result
}
