import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { evaluateFormula, evaluateFormulaReference } from '../formula/formula-engine.ts'
import { formulaError, type FormulaErrorValue, type FormulaEvaluationOptions, type FormulaResolvedReference, type FormulaValue } from '../formula/formula-types.ts'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'

export type SheetValidationSourceReader = (sheet: SpreadsheetSheet, range: CellRange) => boolean

export function sheetValidationSourceReader(userId?: string): SheetValidationSourceReader | undefined {
  return userId === undefined ? undefined : (sheet, range) => canViewWorksheetProtection(sheet, userId)
    && canViewProtectedRanges(sheet, [range, ...sheet.mergedRanges.filter(merge => merge.startRow <= range.endRow
      && merge.endRow >= range.startRow && merge.startCol <= range.endCol && merge.endCol >= range.startCol)], userId)
}

/** Permission policy is passed into the shared formula engine, including nested formula dependencies. */
export function sheetValidationFormulaValue(formula: string, workbook: WorkbookModel, sheetId: string,
  position?: { row: number; column: number }, canReadSource?: SheetValidationSourceReader): FormulaValue {
  return evaluateFormula(formula, workbook, sheetId, formulaReadOptions(position, canReadSource))
}

/** Names and dynamic candidates preserve the shared engine's reference result instead of reparsing formula text. */
export function sheetValidationFormulaReference(formula: string, workbook: WorkbookModel, sheetId: string,
  position?: { row: number; column: number }, canReadSource?: SheetValidationSourceReader): FormulaResolvedReference | FormulaErrorValue {
  return evaluateFormulaReference(formula, workbook, sheetId, formulaReadOptions(position, canReadSource))
}

function formulaReadOptions(position?: { row: number; column: number }, canReadSource?: SheetValidationSourceReader): FormulaEvaluationOptions {
  return { formulaCell: position, canReadRange: canReadSource,
    canReadCell: canReadSource && ((sheet, row, column) => canReadSource(sheet,
      { startRow: row, endRow: row, startCol: column, endCol: column })) }
}

export function sheetValidationCellValue(workbook: WorkbookModel, sheet: SpreadsheetSheet, row: number, column: number,
  canReadSource?: SheetValidationSourceReader): FormulaValue {
  if (canReadSource && !canReadSource(sheet, { startRow: row, endRow: row, startCol: column, endCol: column })) return formulaError('#REF!')
  const cell = sheet.cells[`${row}:${column}`]
  return cell?.formula ? sheetValidationFormulaValue(cell.formula, workbook, sheet.id, { row, column }, canReadSource) : cell?.value ?? null
}
