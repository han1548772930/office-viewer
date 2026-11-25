import { formulaReferenceAreaSignature, translateFormulaReferenceAreas, transformFormulaReferenceAreasForCellShift } from './formula-reference-areas.ts'
import type { SheetCellShiftEdit } from '../../domain/sheet-cell-shift.ts'
export { mapFormulaReferenceAreas, translateFormulaReferenceAreas, transformFormulaReferenceAreasForAxisEdit,
  transformFormulaReferenceAreasForCellShift, type FormulaReferenceArea } from './formula-reference-areas.ts'

/** Compare relative references without converting a recoverable offset into #REF! or changing literal text. */
export function formulaRelativeReferenceSignature(formula: string, row: number, column: number): string {
  return formulaReferenceAreaSignature(formula, row, column)
}
export type FormulaStructureEdit = SheetCellShiftEdit

/** Copy and structural edits use the same lexical reference-area boundary, including whole axes and qualifiers. */
export function translateFormulaReferences(formula: string, rowDelta: number, columnDelta: number): string {
  if (typeof formula !== 'string' || !formula.length) throw new TypeError('Formula must be a non-empty string')
  if (!Number.isSafeInteger(rowDelta) || !Number.isSafeInteger(columnDelta)) throw new TypeError('Formula reference offsets must be safe integers')
  return translateFormulaReferenceAreas(formula, rowDelta, columnDelta)
}
export function transformFormulaReferencesForStructureEdit(formula: string, edit: FormulaStructureEdit,
  formulaSheetName: string | undefined, targetSheetName: string): string {
  if (typeof formula !== 'string' || !formula.length) throw new TypeError('Formula must be a non-empty string')
  return transformFormulaReferenceAreasForCellShift(formula, edit, formulaSheetName, targetSheetName)
}

