import type { SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import { conditionalFormatStyleAt } from '../conditional-formatting/conditional-formatting.ts'
import { conditionalFormatFormulaAt } from '../conditional-formatting/conditional-format-formulas.ts'
import { evaluateFormula } from '../formula/formula-engine.ts'
import { resolveSheetCellStyle } from '../sheet-workbench/sheet-style-inheritance.ts'
import { sheetOpaqueColorHex } from '../sheet-workbench/sheet-colors.ts'

export function canonicalSheetFilterColor(value: string): string { return sheetOpaqueColorHex(value) ?? value.trim().toLowerCase() }
export interface SheetFilterCellColors { fill: string | null; text: string }

/** Reuses the renderer's conditional-style evaluator; colors are derived, never authored into cells. */
export function createSheetFilterColorReader(sheet: SpreadsheetSheet, workbook?: WorkbookModel) {
  const rules = sheet.conditionalFormats ?? []
  const scope = rules.length ? Object.values(sheet.cells).map(cell => ({ row: cell.row, column: cell.col, value: cell.value })) : []
  const formulaWorkbook: WorkbookModel = workbook ?? { kind: 'workbook', name: sheet.name, sheets: [sheet],
    activeSheetId: sheet.id, date1904: false, modified: false }
  return (row: number, column: number, cell: SpreadsheetCell | undefined = sheet.cells[`${row}:${column}`]): SheetFilterCellColors => {
    const base = resolveSheetCellStyle(sheet, row, column, cell, workbook)
    const formulaValue = (formula: string, rule: typeof rules[number]) => {
      const translated = conditionalFormatFormulaAt(formula, rule, row, column)
      return translated === undefined ? undefined : evaluateFormula(translated.startsWith('=') ? translated : `=${translated}`,
        formulaWorkbook, sheet.id, { formulaCell: { row, column } })
    }
    const conditional = rules.length ? conditionalFormatStyleAt(rules, row, column, cell?.value, [], scope,
      (formula, rule) => { const result = formulaValue(formula, rule); return result === true || typeof result === 'number' && result !== 0 }, formulaValue) : undefined
    const fill = conditional?.fill ?? base.fill, text = conditional?.fontColor ?? base.font.color ?? '#000000'
    return { fill: fill === null ? null : canonicalSheetFilterColor(fill), text: canonicalSheetFilterColor(text) }
  }
}
