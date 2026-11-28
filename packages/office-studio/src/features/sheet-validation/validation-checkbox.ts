import type { CellScalar, SheetDataValidationRule } from '../../domain/model.ts'
import type { FormulaValue } from '../formula/formula-types.ts'
import { isFormulaError } from '../formula/formula-types.ts'

export interface SheetCheckboxRuleValues { checked: CellScalar; unchecked: CellScalar }
export interface SheetCheckboxState { checked: boolean; renderable: boolean; valid: boolean }
export type SheetCheckboxFormulaReader = (formula: string, anchor: { row: number; column: number }) => FormulaValue | undefined

/** Only complete boolean words have numeric aliases. Custom text keeps its spelling and whitespace. */
export function normalizeSheetCheckboxValue(value: CellScalar | undefined): string {
  const text = String(value ?? '')
  return /^true$/i.test(text) ? '1' : /^false$/i.test(text) ? '0' : text
}

/** Checkbox operands belong to the rule, rather than shifting with each cell in its ranges. */
export function sheetCheckboxRuleValues(rule: SheetDataValidationRule,
  readFormula?: SheetCheckboxFormulaReader): SheetCheckboxRuleValues | undefined {
  if (rule.type !== 'checkbox') return undefined
  const anchor = { row: rule.formulaAnchor?.row ?? rule.ranges[0]?.startRow ?? 0,
    column: rule.formulaAnchor?.col ?? rule.ranges[0]?.startCol ?? 0 }
  function operand(formula: string | undefined, fallback: number): CellScalar | undefined {
    if (formula === undefined) return fallback
    const value = formula.trim().startsWith('=') ? readFormula?.(formula, anchor) : formula
    return value === undefined || value === null || value === '' || isFormulaError(value)
      || typeof value === 'number' && !Number.isFinite(value) ? undefined : value
  }
  const checked = operand(rule.formula1, 1), unchecked = operand(rule.formula2, 0)
  if (checked === undefined || unchecked === undefined
    || normalizeSheetCheckboxValue(checked) === normalizeSheetCheckboxValue(unchecked)) return undefined
  return { checked, unchecked }
}

/** Invalid existing text remains Canvas text; a missing scalar shows an unchecked control. */
export function sheetCheckboxState(values: SheetCheckboxRuleValues | undefined, actual: CellScalar | undefined,
  hasRichText = false, hasImage = false): SheetCheckboxState {
  if (!values || hasRichText || hasImage) return { checked: false, renderable: false, valid: false }
  const text = normalizeSheetCheckboxValue(actual), blank = actual === undefined || actual === null || actual === ''
  const checked = text === normalizeSheetCheckboxValue(values.checked)
  const valid = checked || text === normalizeSheetCheckboxValue(values.unchecked)
  return { checked: !blank && checked, renderable: blank || valid, valid }
}
