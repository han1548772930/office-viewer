import type { CellRange, SheetConditionalFormatFormulaAnchor, SheetConditionalFormatRule, SheetConditionalFormatValueConfig } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { formulaRelativeReferenceSignature, translateFormulaReferenceAreas } from '../formula/formula-references.ts'

export function conditionalFormatFormulaAnchor(rule: Pick<SheetConditionalFormatRule, 'ranges' | 'formulaAnchor'>): SheetConditionalFormatFormulaAnchor | undefined {
  if (rule.formulaAnchor) return { ...rule.formulaAnchor }
  const first = rule.ranges.reduce<CellRange | undefined>((origin, range) => !origin || range.startRow < origin.startRow
    || range.startRow === origin.startRow && range.startCol < origin.startCol ? range : origin, undefined)
  return first ? { row: first.startRow, col: first.startCol } : undefined
}

function hasStandardPredicate(rule: SheetConditionalFormatRule): boolean {
  if (rule.type === 'office' || !rule.formula1.trim()) return false
  const text = rule.subType === 'text' || ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(rule.type)
  if (!text) return true
  if (['containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'].includes(rule.operator)) return true
  return rule.value !== undefined && String(rule.value) !== rule.formula1
}

/** Standard predicates and formula operands move; matching literals and vendor data remain data. */
export function mapConditionalFormatFormulas(rule: SheetConditionalFormatRule, transform: (formula: string) => string): SheetConditionalFormatRule {
  const result = cloneModel(rule)
  if (rule.type === 'office') return result
  const expression = rule.type === 'expression' || rule.subType === 'formula'
  const numeric = rule.type === 'cellIs' || rule.type === 'highlightCell' && (rule.subType === undefined || rule.subType === 'number')
  if (expression || numeric || hasStandardPredicate(rule)) {
    result.formula1 = rule.formula1 ? transform(rule.formula1) : rule.formula1
    if (rule.formula2 !== undefined) result.formula2 = rule.formula2 ? transform(rule.formula2) : rule.formula2
  }
  if (expression && typeof rule.value === 'string' && rule.value) result.value = transform(rule.value)
  const threshold = (value: SheetConditionalFormatValueConfig): SheetConditionalFormatValueConfig => value.type === 'formula' && typeof value.value === 'string' && value.value
    ? { ...value, value: transform(value.value) } : { ...value }
  const config = result.config
  if (config?.min) config.min = threshold(config.min)
  if (config?.max) config.max = threshold(config.max)
  if (config?.thresholds) config.thresholds = config.thresholds.map(entry => ({ ...entry, value: threshold(entry.value) }))
  if (config?.icons) config.icons = config.icons.map(entry => ({ ...entry, value: threshold(entry.value) }))
  return result
}

function formulaSignature(rule: SheetConditionalFormatRule, origin: SheetConditionalFormatFormulaAnchor): string {
  const normalized = mapConditionalFormatFormulas(rule, formula => formulaRelativeReferenceSignature(formula, origin.row, origin.col))
  return JSON.stringify([normalized.formula1, normalized.formula2, normalized.value, normalized.config])
}

/**
 * Keep one rule and its complete statistical scope. If its earliest cell cannot represent every relative
 * reference, retain the original formula and origin; each evaluated cell can still have a valid reference.
 */
export function reanchorConditionalFormatRule(rule: SheetConditionalFormatRule, ranges: readonly CellRange[]): SheetConditionalFormatRule {
  const before = conditionalFormatFormulaAnchor(rule), after = conditionalFormatFormulaAnchor({ ranges: [...ranges] })
  const copiedRanges = ranges.map(range => ({ ...range }))
  if (rule.type === 'office') return {
    ...cloneModel(rule), ranges: copiedRanges,
    ...(before && after && (before.row !== after.row || before.col !== after.col) ? { formulaAnchor: before } : {}),
  }
  if (!before || !after) return { ...cloneModel(rule), ranges: copiedRanges }
  const moved = before.row === after.row && before.col === after.col ? cloneModel(rule)
    : mapConditionalFormatFormulas(rule, formula => translateFormulaReferenceAreas(formula, after.row - before.row, after.col - before.col))
  if (formulaSignature(rule, before) !== formulaSignature(moved, after)) {
    return { ...cloneModel(rule), ranges: copiedRanges, formulaAnchor: { ...before } }
  }
  delete moved.formulaAnchor
  return { ...moved, ranges: copiedRanges }
}

export function conditionalFormatFormulaAt(formula: string, rule: Pick<SheetConditionalFormatRule, 'ranges' | 'formulaAnchor'>, row: number, col: number): string | undefined {
  const origin = conditionalFormatFormulaAnchor(rule)
  return origin ? translateFormulaReferenceAreas(formula, row - origin.row, col - origin.col) : undefined
}
