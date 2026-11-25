import type { ExcelSheetView, ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { applyExcelWorkbookView, toExcelWorkbookView, cellKey, parseCellKey } from '../../services/excelViewAdapter.ts'
import { createBlankWorkbook } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { mapSheetCellShiftPosition, sheetCellShiftAffectedRange } from '../../domain/sheet-cell-shift.ts'
import { transformFormulaReferencesForStructureEdit, type FormulaStructureEdit } from '../formula/formula-references.ts'
import { planWorkbookCellShiftCommand, sheetCellShiftEdit } from './sheet-cell-shift-command.ts'
import { formatProtectionRange } from '../sheet-protection/range-protection.ts'

export type SheetCellMoveDirection = 'right' | 'down' | 'left' | 'up'
export type SheetCellMoveResult =
  | { status: 'applied'; sheet: ExcelSheetView }
  | { status: 'confirm-merge'; mergeRanges: string[] }
  | { status: 'invalid'; reason: string }

/** Compatibility view adapter; its historical reference-only follow-up remains explicit. */
export function moveSheetCellRange(sheet: ExcelSheetView, range: GridRange, direction: SheetCellMoveDirection,
  dismantlePartialMerges = false, userId = 'local-user'): SheetCellMoveResult {
  try {
    const model = applyExcelWorkbookView(createBlankWorkbook(), {
      name: 'Cell shift', date1904: false, activeSheetId: sheet.id, sheets: [cloneModel(sheet)], calculationMode: 'manual',
    }, 'en-US')
    const target = { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }
    const plan = planWorkbookCellShiftCommand(model, sheet.id, target, direction, userId, dismantlePartialMerges)
    if (plan.requiresConfirmation) return { status: 'confirm-merge', mergeRanges: plan.mergeRanges.map(formatProtectionRange) }
    const next = toExcelWorkbookView(plan.workbook).sheets.find(item => item.id === sheet.id)!
    // Legacy callers publish formulas through transformWorkbookReferencesForSheetEdit.
    // Restore only that projection; UI and Facade use the complete planner directly.
    for (const [key, source] of Object.entries(sheet.cells)) {
      if (!source.formula) continue
      const original = parseCellKey(key), point = original && mapSheetCellShiftPosition(original.row, original.col, plan.edit)
      if (point) next.cells[cellKey(point.row, point.col)]!.formula = source.formula
    }
    return { status: 'applied', sheet: next }
  } catch (error: unknown) {
    return { status: 'invalid', reason: error instanceof Error ? error.message : '单元格移位未能准备。' }
  }
}

export function affectedCellMoveRange(_sheet: ExcelSheetView, range: GridRange, direction: SheetCellMoveDirection): GridRange {
  const affected = sheetCellShiftAffectedRange(structureEdit(range, direction))
  return { startRow: affected.startRow, endRow: affected.endRow, startColumn: affected.startCol, endColumn: affected.endCol }
}

/** Legacy reference-only helper; callers that edit content use planWorkbookCellShiftCommand. */
export function transformWorkbookReferencesForSheetEdit(workbook: ExcelWorkbookView, sheetId: string,
  edit: FormulaStructureEdit): ExcelWorkbookView {
  const target = workbook.sheets.find(sheet => sheet.id === sheetId)
  if (!target) return workbook
  const sheets = workbook.sheets.map(sheet => ({
    ...sheet,
    cells: Object.fromEntries(Object.entries(sheet.cells).map(([key, cell]) => [key, cell.formula
      ? { ...cell, formula: transformFormulaReferencesForStructureEdit(cell.formula, edit, sheet.name, target.name) }
      : cell])),
  }))
  const metadata = workbook.metadata ? cloneModel(workbook.metadata) : undefined
  if (metadata) {
    metadata.definedNames = transformDefinedNames(metadata.definedNames, edit, target.name, target.name)
    metadata.localDefinedNames = transformDefinedNames(metadata.localDefinedNames, edit, target.name, target.name, true)
  }
  return { ...workbook, sheets, metadata }
}

export function structureEdit(range: GridRange, direction: SheetCellMoveDirection): FormulaStructureEdit {
  return sheetCellShiftEdit({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }, direction)
}

function transformDefinedNames(value: unknown, edit: FormulaStructureEdit, formulaSheetName: string,
  targetSheetName: string, local = false): unknown {
  if (!Array.isArray(value)) return value
  return value.map(entry => {
    if (!entry || typeof entry !== 'object') return entry
    const record = entry as Record<string, unknown>
    if (typeof record.refersTo !== 'string' || local && typeof record.sheet !== 'string') return entry
    const ownerSheet = local && typeof record.sheet === 'string' ? record.sheet : formulaSheetName
    return { ...record, refersTo: transformFormulaReferencesForStructureEdit(record.refersTo, edit, ownerSheet, targetSheetName) }
  })
}
