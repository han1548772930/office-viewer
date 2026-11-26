import type { CellRange } from '../../domain/model.ts'
import { parseCellKey, sheetViewCellAt, type ExcelCellView, type ExcelSheetView } from '../../services/excelViewAdapter.ts'

export interface SheetStyleSelectionProbes { cells: ExcelCellView[]; complete: boolean }

/** Sample style/occupancy boundaries without expanding a whole row, column or worksheet selection. */
export function sheetStyleSelectionProbes(sheet: ExcelSheetView, ranges: readonly CellRange[], budget = 100_000): SheetStyleSelectionProbes {
  const filtered = new Set(sheet.filteredRows ?? sheet.autoFilter?.filteredRows ?? [])
  const occupied = Object.keys(sheet.cells).flatMap(key => { const position = parseCellKey(key); return position ? [position] : [] })
  const cells: ExcelCellView[] = [], visited = new Set<string>()
  for (const range of ranges) {
    const rows = new Set<number>(), columns = new Set<number>()
    const addRow = (value: number) => { if (value >= range.startRow && value <= range.endRow && !filtered.has(value)) rows.add(value) }
    const addColumn = (value: number) => { if (value >= range.startCol && value <= range.endCol) columns.add(value) }
    const around = (value: number, add: (value: number) => void) => { add(value - 1); add(value); add(value + 1) }
    addRow(range.startRow); addRow(range.endRow); addColumn(range.startCol); addColumn(range.endCol)
    for (const key of Object.keys(sheet.rowStyles ?? {})) around(Number(key), addRow)
    for (const key of Object.keys(sheet.columnStyles ?? {})) around(Number(key), addColumn)
    for (const span of sheet.rowStyleSpans ?? []) { around(span.start, addRow); around(span.end, addRow) }
    for (const span of sheet.columnStyleSpans ?? []) { around(span.start, addColumn); around(span.end, addColumn) }
    for (const rule of sheet.rangeThemeRules ?? []) {
      for (const row of [rule.range.startRow, rule.range.startRow + 1, rule.range.startRow + 2,
        rule.range.endRow - 2, rule.range.endRow - 1, rule.range.endRow]) around(row, addRow)
      for (const column of [rule.range.startCol, rule.range.startCol + 1, rule.range.startCol + 2,
        rule.range.endCol - 1, rule.range.endCol]) around(column, addColumn)
    }
    if (sheet.rangeThemeRules?.length) {
      for (const row of [...(sheet.hiddenRows ?? []), ...(sheet.outlineHiddenRows ?? [])]) around(row, addRow)
      for (const span of [...(sheet.rowHiddenSpans ?? []), ...(sheet.outlineHiddenRowSpans ?? []), ...(sheet.rowHeightSpans ?? []), ...(sheet.rowAutoHeightSpans ?? [])]) {
        around(span.start, addRow); around(span.end, addRow)
      }
    }
    for (const row of filtered) around(row, addRow)
    for (const position of occupied) if (position.row >= range.startRow && position.row <= range.endRow
      && position.col >= range.startCol && position.col <= range.endCol && !filtered.has(position.row)) {
      around(position.row, addRow); around(position.col, addColumn)
    }
    // Complex sparse cross products need the batch/Worker path; the toolbar reports an undetermined mixed state.
    if (rows.size * columns.size + cells.length > budget) return { cells, complete: false }
    for (const row of rows) for (const col of columns) {
      const key = `${row}:${col}`
      if (!visited.has(key)) { visited.add(key); cells.push(sheetViewCellAt(sheet, row, col)) }
    }
  }
  return { cells, complete: true }
}
