import type {
  CellRange,
  SheetConditionalFormatOperator,
  SheetConditionalFormatRule,
  SheetConditionalFormatStyle,
  SpreadsheetSheet,
} from '../../domain/model.ts'

export interface ConditionalFormatDraft {
  ranges: CellRange[]
  operator: SheetConditionalFormatOperator
  formula1: string
  formula2?: string
  style: SheetConditionalFormatStyle
  stopIfTrue?: boolean
}

export function addConditionalFormatRule(
  sheet: SpreadsheetSheet,
  draft: ConditionalFormatDraft,
  id: string,
): SpreadsheetSheet {
  validateDraft(draft)
  const rules = (sheet.conditionalFormats ?? []).map(cloneRule)
  const next: SheetConditionalFormatRule = {
    id,
    type: 'cellIs',
    ranges: draft.ranges.map(range => ({ ...range })),
    operator: draft.operator,
    formula1: normalizeNumberFormula(draft.formula1),
    formula2: draft.formula2 === undefined ? undefined : normalizeNumberFormula(draft.formula2),
    style: { ...draft.style },
    priority: rules.length + 1,
    stopIfTrue: draft.stopIfTrue ?? false,
  }
  return { ...sheet, conditionalFormats: [...rules, next] }
}

export function updateConditionalFormatRule(
  sheet: SpreadsheetSheet,
  id: string,
  draft: ConditionalFormatDraft,
): SpreadsheetSheet {
  validateDraft(draft)
  const rules = sheet.conditionalFormats ?? []
  if (!rules.some(rule => rule.id === id)) throw new RangeError(`Unknown conditional format rule: ${id}`)
  return {
    ...sheet,
    conditionalFormats: rules.map(rule => rule.id === id ? {
      ...rule,
      ranges: draft.ranges.map(range => ({ ...range })),
      operator: draft.operator,
      formula1: normalizeNumberFormula(draft.formula1),
      formula2: draft.formula2 === undefined ? undefined : normalizeNumberFormula(draft.formula2),
      style: { ...draft.style },
      stopIfTrue: draft.stopIfTrue ?? false,
    } : cloneRule(rule)),
  }
}

export function deleteConditionalFormatRule(sheet: SpreadsheetSheet, id: string): SpreadsheetSheet {
  const remaining = (sheet.conditionalFormats ?? []).filter(rule => rule.id !== id).map(cloneRule)
  if (remaining.length === (sheet.conditionalFormats ?? []).length) return sheet
  return { ...sheet, conditionalFormats: withPriorities(remaining) }
}

export function moveConditionalFormatRule(sheet: SpreadsheetSheet, id: string, targetIndex: number): SpreadsheetSheet {
  const rules = (sheet.conditionalFormats ?? []).map(cloneRule)
  const sourceIndex = rules.findIndex(rule => rule.id === id)
  if (sourceIndex < 0) throw new RangeError(`Unknown conditional format rule: ${id}`)
  if (!Number.isInteger(targetIndex) || targetIndex < 0 || targetIndex >= rules.length) {
    throw new RangeError('Conditional format priority is out of range')
  }
  const [rule] = rules.splice(sourceIndex, 1)
  rules.splice(targetIndex, 0, rule!)
  return { ...sheet, conditionalFormats: withPriorities(rules) }
}

export function evaluateConditionalFormatRule(rule: SheetConditionalFormatRule, value: unknown): boolean {
  if (typeof value !== 'number' || !Number.isFinite(value)) return false
  const first = Number(rule.formula1)
  const second = rule.formula2 === undefined ? undefined : Number(rule.formula2)
  if (!Number.isFinite(first) || (rule.operator === 'between' || rule.operator === 'notBetween') && !Number.isFinite(second)) return false
  switch (rule.operator) {
    case 'equal': return value === first
    case 'notEqual': return value !== first
    case 'greaterThan': return value > first
    case 'greaterThanOrEqual': return value >= first
    case 'lessThan': return value < first
    case 'lessThanOrEqual': return value <= first
    case 'between': return value >= first && value <= second!
    case 'notBetween': return value < first || value > second!
  }
}

export function conditionalFormatStyleAt(
  rules: readonly SheetConditionalFormatRule[],
  row: number,
  column: number,
  value: unknown,
): SheetConditionalFormatStyle {
  const output: SheetConditionalFormatStyle = {}
  const assigned = new Set<keyof SheetConditionalFormatStyle>()
  for (const rule of [...rules].sort((a, b) => a.priority - b.priority)) {
    if (!rule.ranges.some(range => contains(range, row, column)) || !evaluateConditionalFormatRule(rule, value)) continue
    for (const key of ['fill', 'fontColor', 'bold'] as const) {
      if (rule.style[key] !== undefined && !assigned.has(key)) {
        if (key === 'fill' && rule.style.fill !== undefined) output.fill = rule.style.fill
        else if (key === 'fontColor' && rule.style.fontColor !== undefined) output.fontColor = rule.style.fontColor
        else if (key === 'bold' && rule.style.bold !== undefined) output.bold = rule.style.bold
        assigned.add(key)
      }
    }
    if (rule.stopIfTrue) break
  }
  return output
}

export function validateConditionalFormatDraft(draft: ConditionalFormatDraft): void {
  validateDraft(draft)
}

function validateDraft(draft: ConditionalFormatDraft): void {
  if (!draft.ranges.length || draft.ranges.some(range => !validRange(range))) {
    throw new RangeError('Conditional format requires valid cell ranges')
  }
  if (!['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'between', 'notBetween'].includes(draft.operator)) {
    throw new TypeError('Unsupported conditional format operator')
  }
  if (!isFiniteFormula(draft.formula1)) throw new TypeError('Conditional format value must be a finite number')
  if ((draft.operator === 'between' || draft.operator === 'notBetween') && (draft.formula2 === undefined || !isFiniteFormula(draft.formula2))) {
    throw new TypeError('Between conditions require a second finite number')
  }
  for (const color of [draft.style.fill, draft.style.fontColor]) {
    if (color !== undefined && !/^#[0-9a-f]{6}$/i.test(color)) throw new TypeError('Conditional format colors must use #RRGGBB')
  }
  if (draft.style.bold !== undefined && typeof draft.style.bold !== 'boolean') throw new TypeError('Conditional bold style must be boolean')
  if (!draft.style.fill && !draft.style.fontColor && draft.style.bold === undefined) throw new TypeError('Conditional format requires at least one style')
}

function validRange(range: CellRange): boolean {
  return Number.isInteger(range.startRow) && Number.isInteger(range.startCol) && Number.isInteger(range.endRow) && Number.isInteger(range.endCol) &&
    range.startRow >= 0 && range.startCol >= 0 && range.endRow >= range.startRow && range.endCol >= range.startCol
}

function isFiniteFormula(value: string): boolean {
  return Number.isFinite(Number(value.trim().replace(/^=/, '')))
}

function normalizeNumberFormula(value: string): string {
  return value.trim().replace(/^=/, '')
}

function contains(range: CellRange, row: number, column: number): boolean {
  return row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol
}

function withPriorities(rules: SheetConditionalFormatRule[]): SheetConditionalFormatRule[] {
  return rules.map((rule, index) => ({ ...rule, priority: index + 1 }))
}

function cloneRule(rule: SheetConditionalFormatRule): SheetConditionalFormatRule {
  return { ...rule, ranges: rule.ranges.map(range => ({ ...range })), style: { ...rule.style } }
}
