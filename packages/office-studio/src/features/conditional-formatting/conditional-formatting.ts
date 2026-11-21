import type {
  CellRange,
  SheetConditionalFormatOperator,
  SheetConditionalFormatRule,
  SheetConditionalFormatStyle,
  SpreadsheetSheet,
} from '../../domain/model.ts'
import { conditionalIconSequence } from './icon-set-config.ts'

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

export function evaluateConditionalFormatRule(rule: SheetConditionalFormatRule, value: unknown, scopeValues: readonly unknown[] = [], expressionMatches?: (formula: string) => boolean): boolean {
  if (rule.type === 'containsText' || rule.type === 'notContainsText' || rule.type === 'beginsWith' || rule.type === 'endsWith' || rule.subType === 'text') {
    const text = String(value ?? '').toLocaleLowerCase()
    const target = String(rule.value ?? rule.formula1).toLocaleLowerCase()
    const operator = String(rule.operator ?? 'containsText')
    if (operator === 'containsBlanks') return text.trim().length === 0
    if (operator === 'notContainsBlanks') return text.trim().length > 0
    const isError = typeof value === 'string' && ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!'].includes(value) || typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'error'
    if (operator === 'containsErrors') return isError
    if (operator === 'notContainsErrors') return !isError
    if (operator === 'equal') return text === target
    if (operator === 'notEqual') return text !== target
    if (!target) return false
    if (rule.type === 'notContainsText' || operator === 'notContainsText') return !text.includes(target)
    if (rule.type === 'beginsWith' || operator === 'beginsWith') return text.startsWith(target)
    if (rule.type === 'endsWith' || operator === 'endsWith') return text.endsWith(target)
    return text.includes(target)
  }
  if (rule.type === 'duplicateValues' || rule.type === 'uniqueValues') return false
  if (rule.type === 'top10' || rule.subType === 'rank') {
    if (typeof value !== 'number' || !Number.isFinite(value) || !scopeValues.length) return false
    const numeric = scopeValues.filter((item): item is number => typeof item === 'number' && Number.isFinite(item)).sort((a, b) => rule.isBottom ? a - b : b - a)
    const count = Math.max(1, Math.ceil(rule.isPercent ? numeric.length * (Number(rule.value) || 10) / 100 : Number(rule.value) || 10))
    const boundary = numeric[Math.min(count - 1, numeric.length - 1)]
    return boundary !== undefined && (rule.isBottom ? value <= boundary : value >= boundary)
  }
  if (rule.type === 'average' || rule.subType === 'average') {
    if (typeof value !== 'number' || !scopeValues.length) return false
    const numeric = scopeValues.filter((item): item is number => typeof item === 'number' && Number.isFinite(item))
    if (!numeric.length) return false
    const average = numeric.reduce((sum, item) => sum + item, 0) / numeric.length
    if (rule.operator === 'lessThanOrEqual') return value <= average
    if (rule.operator === 'greaterThanOrEqual') return value >= average
    if (rule.operator === 'equal') return value === average
    return rule.operator === 'lessThan' ? value < average : value > average
  }
  if (rule.type === 'expression' || rule.subType === 'formula') return expressionMatches?.(rule.formula1 || String(rule.value ?? '')) ?? false
  if (rule.type === 'timePeriod' || rule.subType === 'timePeriod') return evaluateTimePeriod(rule.operator, value)
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

export interface ConditionalFormatVisualStyle extends SheetConditionalFormatStyle {
  dataBar?: { fraction: number; start: number; color: string; showValue: boolean; gradient: boolean; negative: boolean }
  icon?: { iconType: string; iconId: string; showValue: boolean }
}

export function conditionalFormatStyleAt(
  rules: readonly SheetConditionalFormatRule[],
  row: number,
  column: number,
  value: unknown,
  values: readonly number[] = [],
  scopeValues: readonly { row: number; column: number; value: unknown }[] = [],
  expressionMatches?: (formula: string, rule: SheetConditionalFormatRule) => boolean,
  formulaValue?: (formula: string, rule: SheetConditionalFormatRule) => unknown,
): ConditionalFormatVisualStyle {
  const output: ConditionalFormatVisualStyle = {}
  const assigned = new Set<keyof SheetConditionalFormatStyle>()
  for (const rule of [...rules].sort((a, b) => a.priority - b.priority)) {
    if (!rule.ranges.some(range => contains(range, row, column))) continue
    const scopedValues = scopeValues.filter(item => rule.ranges.some(range => contains(range, item.row, item.column))).map(item => item.value)
    const numericValues = scopeValues.length ? scopedValues.filter((item): item is number => typeof item === 'number' && Number.isFinite(item)) : values
    const threshold = (config: { type: string; value?: number | string } | undefined) => resolveThreshold(config, numericValues, typeof value === 'number' ? value : 0, formula => formulaValue?.(formula, rule))
    if (rule.type === 'duplicateValues' || rule.subType === 'duplicateValues') {
      if (valueKey(value) === undefined) continue
      const count = scopedValues.filter(item => valueKey(item) === valueKey(value)).length
      if (count < 2) continue
      applyConditionalStyle(output, assigned, rule.style)
      if (rule.stopIfTrue) break
      continue
    }
    if (rule.type === 'uniqueValues' || rule.subType === 'uniqueValues') {
      if (valueKey(value) === undefined) continue
      const count = scopedValues.filter(item => valueKey(item) === valueKey(value)).length
      if (count !== 1) continue
      applyConditionalStyle(output, assigned, rule.style)
      if (rule.stopIfTrue) break
      continue
    }
    if (rule.type === 'dataBar' && typeof value === 'number' && Number.isFinite(value)) {
      if (output.dataBar) continue
      const min = threshold(rule.config?.min ?? { type: 'min' })
      const max = threshold(rule.config?.max ?? { type: 'max' })
      if (!Number.isFinite(min) || !Number.isFinite(max) || max <= min || value < min) continue
      const start = min < 0 ? max <= 0 ? 1 : -min / (max - min) : 0
      const negative = value < 0
      const fraction = negative ? max <= 0 ? Math.min(1, (max - value) / (max - min)) : Math.min(1, value / min) * start
        : min < 0 ? Math.min(1, value / max) * (1 - start) : Math.min(1, (value - min) / (max - min))
      output.dataBar = { fraction: Math.max(0, fraction), start, color: negative ? rule.config?.nativeColor ?? '#f8696b' : rule.config?.positiveColor ?? '#638ec6', showValue: rule.config?.isShowValue !== false, gradient: rule.config?.isGradient ?? false, negative }
      continue
    }
    if (rule.type === 'colorScale' && typeof value === 'number' && rule.config?.thresholds?.length) {
      if (!assigned.has('fill')) {
        const color = colorScaleColor(value, numericValues, rule.config.thresholds, formula => formulaValue?.(formula, rule))
        if (color) { output.fill = color; assigned.add('fill') }
      }
      continue
    }
    if (rule.type === 'iconSet' && typeof value === 'number' && Number.isFinite(value) && rule.config?.icons?.length) {
      if (output.icon) continue
      const icons = conditionalIconSequence(rule.config)
      const boundaries = icons.slice(0, -1).map(icon => ({ icon, value: threshold(icon.value) }))
      if (boundaries.some(boundary => !Number.isFinite(boundary.value))) continue
      const selected = boundaries.find(boundary => compareThreshold(boundary.icon.operator, boundary.value, value))?.icon ?? icons.at(-1)
      if (selected?.iconId !== undefined && selected.iconType) {
        output.icon = { iconType: selected.iconType, iconId: selected.iconId, showValue: rule.config.isShowValue !== false }
      }
      continue
    }
    if (!evaluateConditionalFormatRule(rule, value, scopedValues, formula => expressionMatches?.(formula, rule))) continue
    applyConditionalStyle(output, assigned, rule.style)
    if (rule.stopIfTrue) break
  }
  return output
}

function applyConditionalStyle(output: SheetConditionalFormatStyle, assigned: Set<keyof SheetConditionalFormatStyle>, style: SheetConditionalFormatStyle): void {
  for (const key of ['fill', 'fontColor', 'bold', 'italic', 'underline', 'strike'] as const) {
    if (style[key] === undefined || assigned.has(key)) continue
    Object.assign(output, { [key]: style[key] })
    assigned.add(key)
  }
}

function valueKey(value: unknown): string | undefined {
  if (value === null || value === undefined || value === '') return undefined
  return `${typeof value}:${String(value).toLocaleLowerCase()}`
}

function compareThreshold(operator: string | undefined, threshold: number, value: number): boolean {
  switch (operator) {
    case 'greaterThan': return value > threshold
    case 'greaterThanOrEqual': return value >= threshold
    case 'lessThan': return value < threshold
    case 'lessThanOrEqual': return value <= threshold
    case 'equal': return value === threshold
    case 'notEqual': return value !== threshold
    default: return value >= threshold
  }
}

function extreme(values: readonly number[], fallback: number, select: typeof Math.min | typeof Math.max): number {
  return values.reduce((current, value) => select(current, value), fallback)
}

function resolveThreshold(config: { type: string; value?: number | string } | undefined, values: readonly number[], fallback: number, formulaValue?: (formula: string) => unknown): number {
  const minimum = values.length ? extreme(values, Number.POSITIVE_INFINITY, Math.min) : fallback
  const maximum = values.length ? extreme(values, Number.NEGATIVE_INFINITY, Math.max) : fallback
  if (!config || config.type === 'min') return minimum
  if (config.type === 'max') return maximum
  if (config.type === 'num' || config.type === 'formula') {
    const result = config.type === 'formula' ? formulaValue?.(String(config.value ?? '')) : config.value
    const number = typeof result === 'number' || typeof result === 'string' && result.trim() ? Number(result) : Number.NaN
    return Number.isFinite(number) ? number : Number.NaN
  }
  const amount = Number(config.value)
  if (config.type === 'percent' && Number.isFinite(amount)) return minimum + (maximum - minimum) * amount / 100
  if (config.type === 'percentile' && Number.isFinite(amount)) {
    const sorted = [...values].sort((a, b) => a - b)
    if (!sorted.length) return fallback
    const position = (sorted.length - 1) * Math.max(0, Math.min(100, amount)) / 100
    const lower = Math.floor(position), fraction = position - lower
    return sorted[lower]! + ((sorted[lower + 1] ?? sorted[lower]!) - sorted[lower]!) * fraction
  }
  return fallback
}

function colorScaleColor(value: number, values: readonly number[], thresholds: Array<{ value: { type: string; value?: number | string }; color: string }>, formulaValue?: (formula: string) => unknown): string | undefined {
  const threshold = (config: { type: string; value?: number | string }) => resolveThreshold(config, values, value, formulaValue)
  if (thresholds.some(item => !Number.isFinite(threshold(item.value)))) return undefined
  const sorted = [...thresholds].sort((a, b) => threshold(a.value) - threshold(b.value))
  if (value <= threshold(sorted[0]!.value)) return sorted[0]!.color
  for (let index = 1; index < sorted.length; index++) {
    const lower = sorted[index - 1]!, upper = sorted[index]!
    const low = threshold(lower.value), high = threshold(upper.value)
    if (value <= high) return interpolateColor(lower.color, upper.color, high === low ? 1 : (value - low) / (high - low))
  }
  return sorted.at(-1)!.color
}

function interpolateColor(first: string, second: string, fraction: number): string {
  const a = first.match(/[\da-f]{2}/gi), b = second.match(/[\da-f]{2}/gi)
  if (a?.length !== 3 || b?.length !== 3) return first
  return `#${a.map((part, index) => Math.round(parseInt(part, 16) + (parseInt(b[index]!, 16) - parseInt(part, 16)) * fraction).toString(16).padStart(2, '0')).join('')}`
}

function evaluateTimePeriod(operator: string, value: unknown): boolean {
  if (typeof value !== 'number' || !Number.isFinite(value)) return false
  const date = new Date(Date.UTC(1899, 11, 30) + value * 86_400_000), now = new Date()
  const today = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate())
  const serial = (candidate: Date) => Math.floor((Date.UTC(candidate.getUTCFullYear(), candidate.getUTCMonth(), candidate.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86_400_000)
  const current = serial(date), todaySerial = Math.floor((today - Date.UTC(1899, 11, 30)) / 86_400_000)
  if (operator === 'today') return current === todaySerial
  if (operator === 'yesterday') return current === todaySerial - 1
  if (operator === 'tomorrow') return current === todaySerial + 1
  if (operator === 'last7Days') return current >= todaySerial - 6 && current <= todaySerial
  const todayDay = now.getUTCDay()
  if (operator === 'thisWeek') return current >= todaySerial - todayDay && current <= todaySerial + (6 - todayDay)
  if (operator === 'lastWeek') return current >= todaySerial - todayDay - 7 && current < todaySerial - todayDay
  if (operator === 'nextWeek') return current > todaySerial + (6 - todayDay) && current <= todaySerial + (13 - todayDay)
  if (operator === 'thisMonth') return date.getUTCFullYear() === now.getUTCFullYear() && date.getUTCMonth() === now.getUTCMonth()
  if (operator === 'lastMonth') {
    const priorMonth = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1))
    return date.getUTCFullYear() === priorMonth.getUTCFullYear() && date.getUTCMonth() === priorMonth.getUTCMonth()
  }
  if (operator === 'nextMonth') {
    const nextMonth = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1))
    return date.getUTCFullYear() === nextMonth.getUTCFullYear() && date.getUTCMonth() === nextMonth.getUTCMonth()
  }
  return false
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
  if (!draft.style.fill && !draft.style.fontColor && [draft.style.bold, draft.style.italic, draft.style.underline, draft.style.strike].every(value => value === undefined)) throw new TypeError('Conditional format requires at least one style')
}

function validRange(range: CellRange): boolean {
  return Number.isInteger(range.startRow) && Number.isInteger(range.startCol) && Number.isInteger(range.endRow) && Number.isInteger(range.endCol) &&
    range.startRow >= 0 && range.startCol >= 0 && range.endRow >= range.startRow && range.endCol >= range.startCol
}

function isFiniteFormula(value: string): boolean {
  const normalized = value.trim().replace(/^=/, '')
  return normalized !== '' && Number.isFinite(Number(normalized))
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
  return { ...rule, ranges: rule.ranges.map(range => ({ ...range })), style: { ...rule.style }, config: rule.config ? { ...rule.config, min: rule.config.min ? { ...rule.config.min } : undefined, max: rule.config.max ? { ...rule.config.max } : undefined, thresholds: rule.config.thresholds?.map(item => ({ ...item, value: { ...item.value } })), icons: rule.config.icons?.map(item => ({ ...item, value: { ...item.value } })) } : undefined, metadata: rule.metadata ? { ...rule.metadata } : undefined }
}
