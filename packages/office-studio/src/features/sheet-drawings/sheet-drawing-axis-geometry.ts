import type { SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'

/** The inputs shared by floating-object projections and the visible Sheet axes. */
export function sheetDrawingAxisGeometryInput(workbook: WorkbookModel, sheet: SpreadsheetSheet) {
  return {
    rowHeights: sheet.rowHeights, columnWidths: sheet.columnWidths, rowHeightSpans: sheet.rowHeightSpans,
    columnWidthSpans: sheet.columnWidthSpans, rowAutoHeightSpans: sheet.rowAutoHeightSpans, rowAutoHeights: sheet.rowAutoHeights,
    defaultRowHeight: sheet.defaultRowHeight, defaultRowAutoHeight: sheet.defaultRowAutoHeight, defaultColumnWidth: sheet.defaultColumnWidth,
    hiddenRows: sheet.hiddenRows, hiddenColumns: sheet.hiddenColumns, rowHiddenSpans: sheet.rowHiddenSpans,
    columnHiddenSpans: sheet.columnHiddenSpans, outline: sheet.outline, filteredRows: sheet.autoFilter?.filteredRows,
    tableFilters: (workbook.tables ?? []).filter(table => table.sheetId === sheet.id).map(table => table.autoFilter?.filteredRows),
  }
}
