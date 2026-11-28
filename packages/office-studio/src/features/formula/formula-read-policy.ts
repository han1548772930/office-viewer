import type { CellRange, SpreadsheetSheet } from '../../domain/model'
import type { FormulaEvaluationOptions } from './formula-types'

const MAX_CELL_POLICY_CHECKS = 250_000
const overlaps = (left: CellRange, right: CellRange) => left.startRow <= right.endRow && left.endRow >= right.startRow
  && left.startCol <= right.endCol && left.endCol >= right.startCol

/** A range grant covers blank cells and merged subcells as well as sparse data.
 * Cell-only callers can prove bounded ranges; large callers supply the range
 * policy so a whole-column permission check never expands a million cells. */
export function formulaReadRangeAllowed(sheet: SpreadsheetSheet, range: CellRange, options: FormulaEvaluationOptions): boolean {
  if (![range.startRow, range.endRow, range.startCol, range.endCol].every(Number.isSafeInteger)
    || range.startRow < 0 || range.startCol < 0 || range.startRow > range.endRow || range.startCol > range.endCol
    || range.endRow > 1_048_575 || range.endCol > 16_383) return false
  if (!options.canReadCell && !options.canReadRange) return true
  const scopes = [range, ...sheet.mergedRanges.filter(merge => overlaps(range, merge))]
  try {
    if (options.canReadRange) return scopes.every(scope => options.canReadRange!(sheet, scope))
    let remaining = MAX_CELL_POLICY_CHECKS
    for (const scope of scopes) {
      const count = (scope.endRow - scope.startRow + 1) * (scope.endCol - scope.startCol + 1)
      if (!Number.isSafeInteger(count) || count > remaining) return false
      remaining -= count
      for (let row = scope.startRow; row <= scope.endRow; row++) for (let column = scope.startCol; column <= scope.endCol; column++) {
        if (!options.canReadCell!(sheet, row, column)) return false
      }
    }
    return true
  } catch { return false }
}

export function formulaReadCellAllowed(sheet: SpreadsheetSheet, row: number, column: number, options: FormulaEvaluationOptions): boolean {
  if (!formulaReadRangeAllowed(sheet, { startRow: row, endRow: row, startCol: column, endCol: column }, options)) return false
  try { return options.canReadCell?.(sheet, row, column) !== false } catch { return false }
}
