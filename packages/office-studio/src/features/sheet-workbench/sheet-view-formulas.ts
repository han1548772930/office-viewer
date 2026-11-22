import { createBlankWorkbook } from '../../domain/model.ts'
import type { WorkbookModel } from '../../domain/model.ts'
import { applyExcelWorkbookView, parseCellKey } from '../../services/excelViewAdapter.ts'
import type { ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { recalculateWorkbookFormulas } from '../formula/formula-engine.ts'
import type { FormulaEngineDiagnostic } from '../formula/formula-types.ts'

export interface SheetViewRecalculationResult {
  workbook: ExcelWorkbookView
  diagnostics: FormulaEngineDiagnostic[]
}

/** Preserve explicitly typed values when a standalone view has no previous domain model. */
function seedRawValues(workbook: WorkbookModel, view: ExcelWorkbookView): WorkbookModel {
  return {
    ...workbook,
    sheets: workbook.sheets.map(sheet => {
      const source = view.sheets.find(candidate => candidate.id === sheet.id)
      if (!source) return sheet
      const cells = { ...sheet.cells }
      for (const [address, cell] of Object.entries(source.cells)) {
        if (cell.rawValue === undefined || (!cell.formula && String(cell.rawValue ?? '') !== cell.value)) continue
        const position = parseCellKey(address)
        const key = position ? `${position.row}:${position.col}` : ''
        const converted = cells[key]
        if (converted) cells[key] = { ...converted, value: cell.rawValue }
      }
      return { ...sheet, cells }
    }),
  }
}

/**
 * The workbench uses the domain formula engine, not a second UI evaluator.
 * Only formula result fields are projected back; sheet dimensions, resources,
 * source formulas, formats and non-formula cells remain those of the view.
 */
export function recalculateSheetViewFormulas(view: ExcelWorkbookView, original?: WorkbookModel): SheetViewRecalculationResult {
  const baseline = original ?? seedRawValues(applyExcelWorkbookView(createBlankWorkbook(), view), view)
  const converted = original ? applyExcelWorkbookView(original, view) : baseline
  const result = recalculateWorkbookFormulas(converted, baseline)
  const preservedCaches = new Set(result.diagnostics
    .filter(diagnostic => diagnostic.code === 'FORMULA_UNSUPPORTED_CACHE_PRESERVED')
    .map(diagnostic => `${diagnostic.cell.sheetId}!${diagnostic.cell.row}:${diagnostic.cell.column}`))
  let workbookChanged = false
  const sheets = view.sheets.map(sheet => {
    const computed = result.workbook.sheets.find(candidate => candidate.id === sheet.id)
    if (!computed) return sheet
    const cells = { ...sheet.cells }
    let changed = false
    for (const [address, cell] of Object.entries(sheet.cells)) {
      if (!cell.formula) continue
      const position = parseCellKey(address)
      if (!position) continue
      const key = `${position.row}:${position.col}`
      const calculated = computed.cells[key]
      if (!calculated) continue
      const preserved = preservedCaches.has(`${sheet.id}!${key}`)
      const unchangedValue = cell.rawValue !== undefined && cell.rawValue === calculated.value
      const cachedValue = (preserved || unchangedValue) && cell.cachedValue !== undefined ? cell.cachedValue : calculated.display
      if (cell.cachedValue === cachedValue && cell.rawValue === calculated.value && cell.displayValue === cachedValue
        && cell.dataType === calculated.dataType && cell.cachedDataType === calculated.cachedDataType) continue
      cells[address] = { ...cell, cachedValue, rawValue: calculated.value, displayValue: cachedValue,
        dataType: calculated.dataType, cachedDataType: calculated.cachedDataType }
      changed = true
    }
    if (!changed) return sheet
    workbookChanged = true
    return { ...sheet, cells }
  })
  return { workbook: workbookChanged ? { ...view, sheets } : view, diagnostics: result.diagnostics }
}
