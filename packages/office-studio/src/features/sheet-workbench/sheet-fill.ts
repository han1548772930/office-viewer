import { applyExcelCellView, toExcelCellView, type ExcelCellView, type ExcelSheetView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { createBlankWorkbook, createCell } from '../../domain/model.ts'
import { createSheetAutoFillPlan, SHEET_FILL_SYNC_CELL_BUDGET } from './sheet-auto-fill.ts'
import { resolveSheetFillTarget, sheetFillPointerTarget } from './sheet-fill-target.ts'

export interface SheetFillCellPatch {
  row: number
  column: number
  cell: ExcelCellView
}

export interface SheetFillPlan {
  range: GridRange
  cellPatches: SheetFillCellPatch[]
}

export type SheetFillApplyType = 'COPY' | 'SERIES' | 'ONLY_FORMAT' | 'NO_FORMAT'

export interface SheetFillOptions {
  applyType?: SheetFillApplyType
  targetCells?: readonly (readonly (ExcelCellView | undefined)[])[]
}

export function detectSheetFillDownEndRow(source: GridRange, sheet: Pick<ExcelSheetView, 'cells' | 'rowCount' | 'columnCount'>): number | undefined {
  const firstAdjacentRow = source.endRow + 1
  if (firstAdjacentRow >= sheet.rowCount) return undefined
  const leftColumn = source.startColumn > 0 ? source.startColumn - 1 : undefined
  const rightColumn = source.endColumn + 1 < sheet.columnCount ? source.endColumn + 1 : undefined
  const referenceColumn = leftColumn !== undefined && hasCellContent(sheet, leftColumn, firstAdjacentRow)
    ? leftColumn
    : rightColumn !== undefined && hasCellContent(sheet, rightColumn, firstAdjacentRow)
      ? rightColumn
      : undefined
  if (referenceColumn === undefined) return undefined

  let detectedEndRow = firstAdjacentRow
  while (detectedEndRow + 1 < sheet.rowCount && hasCellContent(sheet, referenceColumn, detectedEndRow + 1)) {
    detectedEndRow += 1
  }
  for (let row = firstAdjacentRow; row <= detectedEndRow; row += 1) {
    for (let column = source.startColumn; column <= source.endColumn; column += 1) {
      if (hasCellContent(sheet, column, row)) return row - 1 > source.endRow ? row - 1 : undefined
    }
  }
  return detectedEndRow > source.endRow ? detectedEndRow : undefined
}

function hasCellContent(sheet: Pick<ExcelSheetView, 'cells'>, column: number, row: number): boolean {
  const address = `${columnName(column)}${row + 1}`
  const cell = sheet.cells[address]
  return Boolean(cell?.formula || cell && (cell.rawValue !== undefined && cell.rawValue !== null && cell.rawValue !== '' || cell.value !== ''))
}

function columnName(index: number): string {
  let result = ''
  for (let value = index; value >= 0; value = Math.floor(value / 26) - 1) result = String.fromCharCode(value % 26 + 65) + result
  return result
}

/** View-only compatibility adapter delegates to the same domain planner. It
 * cannot carry workbook resources; Canvas authoring uses the complete model. */
export function createSheetFillPlan(
  source: GridRange,
  target: { row: number; column: number },
  sourceCells: readonly (readonly (ExcelCellView | undefined)[])[],
  options: SheetFillOptions = {},
): SheetFillPlan {
  const sourceRange = { startRow: source.startRow, endRow: source.endRow, startCol: source.startColumn, endCol: source.endColumn }
  const selection = sheetFillPointerTarget(sourceRange, target.row, target.column)
  const resolved = resolveSheetFillTarget(sourceRange, selection)
  const range = { startRow: selection.startRow, endRow: selection.endRow, startColumn: selection.startCol, endColumn: selection.endCol }
  if (!resolved) return { range, cellPatches: [] }
  const work = (selection.endRow - selection.startRow + 1) * (selection.endCol - selection.startCol + 1)
  if (work > SHEET_FILL_SYNC_CELL_BUDGET) throw new RangeError('填充范围超出当前同步预算，需要分块 Worker 填充。')
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]!
  sheet.rowCount = Math.max(sheet.rowCount ?? 1, source.endRow + 1, selection.endRow + 1)
  sheet.columnCount = Math.max(sheet.columnCount ?? 1, source.endColumn + 1, selection.endCol + 1)
  for (let r = 0; r < (options.targetCells?.length ?? 0); r++) {
    const row = options.targetCells![r]!
    for (let c = 0; c < row.length; c++) if (row[c]) {
      const position = { row: selection.startRow + r, col: selection.startCol + c }
      sheet.cells[`${position.row}:${position.col}`] = applyExcelCellView(position, row[c]!)
    }
  }
  for (let r = 0; r <= source.endRow - source.startRow; r++) for (let c = 0; c <= source.endColumn - source.startColumn; c++) {
    const cell = sourceCells[r]?.[c], position = { row: source.startRow + r, col: source.startColumn + c }, key = `${position.row}:${position.col}`
    if (cell) sheet.cells[key] = applyExcelCellView(position, cell)
    else delete sheet.cells[key]
  }
  const plan = createSheetAutoFillPlan(workbook, sheet.id, sourceRange, selection, { applyType: options.applyType })
  if (!plan.sheet || !plan.result.accepted) return { range, cellPatches: [] }
  const cellPatches: SheetFillCellPatch[] = []
  for (const destination of plan.result.ranges) for (let row = destination.startRow; row <= destination.endRow; row++) {
    for (let column = destination.startCol; column <= destination.endCol; column++) {
      const cell = plan.sheet.cells[`${row}:${column}`] ?? createCell(row, column, null)
      cellPatches.push({ row, column, cell: toExcelCellView(cell) })
    }
  }
  return { range, cellPatches }
}
