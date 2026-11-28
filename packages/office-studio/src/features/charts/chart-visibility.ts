import type { CellRange, SheetAutoFilter, SpreadsheetSheet, WorkbookModel } from '../../domain/model'
import { refreshAutoFilter } from '../sheet-sort-filter/sort-filter-engine'
import { hiddenSheetColumns, hiddenSheetRows, isSheetColumnVisible, isSheetRowVisible } from '../sheet-workbench/sheet-row-visibility'

/** Charts query the same size/manual/filter/outline reasons as the grid, without persisting derived hides. */
export function chartSourceVisibility(workbook: WorkbookModel, sheet: SpreadsheetSheet, range: CellRange): {
  rowVisible: (row: number) => boolean
  columnVisible: (column: number) => boolean
} {
  const filteredRows = new Set(sheet.autoFilter?.filteredRows ?? [])
  for (const table of workbook.tables ?? []) {
    if (table.sheetId !== sheet.id || !table.autoFilter) continue
    // refreshAutoFilter updates only this temporary filter. Reading a chart must
    // not refresh or otherwise mutate the workbook's authored table/filter state.
    const filter: SheetAutoFilter = { ...table.autoFilter, range: { ...table.autoFilter.range }, filteredRows: [] }
    for (const row of refreshAutoFilter({ ...sheet, autoFilter: filter }, workbook)) filteredRows.add(row)
  }
  const visibility = {
    ...sheet,
    rowCount: sheet.rowCount ?? range.endRow + 1,
    columnCount: sheet.columnCount ?? range.endCol + 1,
    filteredRows: [...filteredRows],
  }
  const rows = hiddenSheetRows(visibility), columns = hiddenSheetColumns(visibility)
  // Keep large collapsed groups as intervals instead of expanding all their
  // indexes. This is the same inclusive range contract used by sheet-outline.
  const collapsed = sheet.outline?.groups.filter(group => group.collapsed) ?? []
  return {
    rowVisible: row => isSheetRowVisible(visibility, row, rows) && !collapsed.some(group => group.axis === 'row' && row >= group.start && row <= group.end),
    columnVisible: column => isSheetColumnVisible(visibility, column, columns) && !collapsed.some(group => group.axis === 'column' && column >= group.start && column <= group.end),
  }
}
