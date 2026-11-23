import type { CellRange, SheetDataValidationRule } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { formulaRelativeReferenceSignature } from '../formula/formula-references.ts'
import { translateFormulaReferenceAreas } from '../formula/formula-reference-areas.ts'

export function sheetValidationUsesRelativeAnchor(rule: Pick<SheetDataValidationRule, 'type'>): boolean {
  return rule.type !== 'any' && rule.type !== 'list' && rule.type !== 'listMultiple' && rule.type !== 'checkbox'
}

export function sheetValidationFormulaAnchor(rule: Pick<SheetDataValidationRule, 'ranges' | 'formulaAnchor'>): { row: number; col: number } | undefined {
  const first = rule.ranges[0]
  return rule.formulaAnchor ? { ...rule.formulaAnchor } : first ? { row: first.startRow, col: first.startCol } : undefined
}

export function sheetValidationFormulaAt(formula: string, rule: Pick<SheetDataValidationRule, 'ranges' | 'formulaAnchor'>, row: number, col: number): string {
  const anchor = sheetValidationFormulaAnchor(rule)
  return anchor ? translateFormulaReferenceAreas(formula, row - anchor.row, col - anchor.col) : formula
}

/** Range replacement preserves the rule's per-cell predicate, including partly invalid relative references. */
export function reanchorSheetDataValidationRule(rule: SheetDataValidationRule, ranges: readonly CellRange[]): SheetDataValidationRule {
  const result = cloneModel(rule), before = sheetValidationFormulaAnchor(rule), after = ranges[0]
  result.ranges = ranges.map(range => ({ ...range }))
  if (!before || !after || !sheetValidationUsesRelativeAnchor(rule)) return result
  const transform = (formula: string | undefined): string | undefined => formula === undefined || formula === '' ? formula
    : translateFormulaReferenceAreas(formula, after.startRow - before.row, after.startCol - before.col)
  const formulas = [rule.formula1, rule.formula2], translated = formulas.map(transform)
  const lossless = formulas.every((formula, index) => formula === undefined || formula === ''
    || formulaRelativeReferenceSignature(formula, before.row, before.col)
      === formulaRelativeReferenceSignature(translated[index]!, after.startRow, after.startCol))
  if (!lossless) { result.formulaAnchor = before; return result }
  result.formula1 = translated[0]; result.formula2 = translated[1]
  delete result.formulaAnchor
  return result
}
