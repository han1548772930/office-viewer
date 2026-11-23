import type { CellRange, SheetConditionalFormatRule, SheetConditionalFormatValueConfig, SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { parseFormula } from '../formula/formula-parser.ts'
import { canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { assertSheetCommandRanges, assertSheetRangesEditable } from '../sheet-workbench/sheet-range-command.ts'
import { validateConditionalIconConfig, validateConditionalThreshold } from './icon-set-config.ts'
import { assertConditionalFormatKnownFields, conditionalFormatOfficeBaseline, conditionalFormatRuleTypes, normalizeConditionalFormatOfficeSource } from './conditional-format-office-source.ts'
import { assertConditionalFormatModelBudget, conditionalFormatSerializedSize } from './conditional-format-source-budget.ts'

export const SET_SHEET_CONDITIONAL_FORMATS_COMMAND = 'sheet.command.set-conditional-formats'

export interface SetSheetConditionalFormatsInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  rules: readonly SheetConditionalFormatRule[]
  expectedRevision?: number
}

export type SheetConditionalFormatSubmitResult = { accepted: true; changed: boolean } | { accepted: false; error: string }

export interface SheetConditionalFormatsPlan {
  rules: SheetConditionalFormatRule[]
  ranges: CellRange[]
  changed: boolean
}

type ConditionalFormatSheet = Pick<SpreadsheetSheet, 'conditionalFormats' | 'worksheetProtection' | 'rangeProtectionRules'>
const numberOperators = ['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'between', 'notBetween']
const subTypes = ['number', 'text', 'timePeriod', 'formula', 'rank', 'average', 'duplicateValues', 'uniqueValues']

export function sheetConditionalFormatsCommandInput(input: SetSheetConditionalFormatsInput | undefined): SetSheetConditionalFormatsInput {
  if (!input || typeof input !== 'object' || Array.isArray(input) || !Array.isArray(input.rules)) throw new TypeError('条件格式命令需要规则列表。')
  for (const id of [input.unitId, input.sheetId, input.subUnitId]) {
    if (id !== undefined && (typeof id !== 'string' || !id.trim())) throw new TypeError('工作簿或工作表标识无效。')
  }
  if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw new TypeError('条件格式目标工作表不一致。')
  if (input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0)) throw new TypeError('工作簿版本无效。')
  return input
}

/** Object property order is not a document edit; arrays retain their observable order. */
function signature(value: unknown): string {
  return JSON.stringify(value, (_key, item: unknown) => item && typeof item === 'object' && !Array.isArray(item)
    ? Object.fromEntries(Object.entries(item).sort(([left], [right]) => left.localeCompare(right))) : item) ?? ''
}

function formulaRole(rule: SheetConditionalFormatRule): 'expression' | 'number' | undefined {
  if (rule.type === 'office') return undefined
  if (rule.type === 'expression' || rule.subType === 'formula') return 'expression'
  if (rule.type === 'cellIs' || rule.type === 'highlightCell' && (rule.subType === undefined || rule.subType === 'number')) return 'number'
  return undefined
}

function assertFormula(formula: string, previous: string | undefined): void {
  if (!formula.trim().replace(/^=/, '').trim()) throw new TypeError('条件格式公式不能为空。')
  // Imported formulas outside the current parser's grammar remain intact when they have not been edited.
  if (formula === previous) return
  try { parseFormula(formula) } catch (reason) {
    throw new TypeError(`条件格式公式无效：${reason instanceof Error ? reason.message : '无法解析公式。'}`)
  }
}

function assertThreshold(value: SheetConditionalFormatValueConfig, previous?: SheetConditionalFormatValueConfig, automatic?: 'min' | 'max'): void {
  if (!value || typeof value !== 'object' || Array.isArray(value) || typeof value.type !== 'string') throw new TypeError('条件格式阈值无效。')
  const isAutomatic = value.type === 'autoMin' || value.type === 'autoMax'
  if (isAutomatic && value.type !== (automatic === 'min' ? 'autoMin' : automatic === 'max' ? 'autoMax' : undefined)) throw new TypeError('自动阈值只适用于相应的数据条边界。')
  if (isAutomatic || signature(value) !== signature(previous)) validateConditionalThreshold(value, automatic !== undefined)
  if (value.type === 'formula') assertFormula(String(value.value ?? ''), previous?.type === 'formula' ? String(previous.value ?? '') : undefined)
}

/** Validate edited fields without stripping unedited parser/vendor configuration. */
export function assertSheetConditionalFormatRuleSnapshot(rule: SheetConditionalFormatRule, previous?: SheetConditionalFormatRule): void {
  if (!rule || typeof rule !== 'object' || Array.isArray(rule) || typeof rule.id !== 'string' || !rule.id.trim()) throw new TypeError('条件格式规则标识无效。')
  conditionalFormatSerializedSize(rule)
  if (!conditionalFormatRuleTypes.includes(rule.type)) throw new TypeError('条件格式规则类型无效。')
  if (rule.subType !== undefined && !subTypes.includes(rule.subType)) throw new TypeError('条件格式条件类型无效。')
  if (!Array.isArray(rule.ranges) || !rule.ranges.length) throw new RangeError('条件格式需要有效的应用范围。')
  assertSheetCommandRanges(rule.ranges)
  const anchor = rule.formulaAnchor
  if (anchor !== undefined && (!anchor || !Number.isSafeInteger(anchor.row) || !Number.isSafeInteger(anchor.col)
    || anchor.row < 0 || anchor.row >= 1_048_576 || anchor.col < 0 || anchor.col >= 16_384)) throw new RangeError('条件格式公式基准无效。')
  if (!Number.isSafeInteger(rule.priority) || rule.priority < 1) throw new RangeError('条件格式优先级无效。')
  if (typeof rule.stopIfTrue !== 'boolean' || typeof rule.operator !== 'string' || typeof rule.formula1 !== 'string'
    || rule.formula2 !== undefined && typeof rule.formula2 !== 'string') throw new TypeError('条件格式规则参数无效。')
  if (!rule.style || typeof rule.style !== 'object' || Array.isArray(rule.style)) throw new TypeError('条件格式样式无效。')
  assertConditionalFormatKnownFields(rule)
  for (const color of [rule.style.fill, rule.style.fontColor, rule.config?.positiveColor, rule.config?.nativeColor]) {
    if (color !== undefined && (typeof color !== 'string' || !/^#[0-9a-f]{6}$/i.test(color))) throw new TypeError('条件格式颜色必须使用 #RRGGBB。')
  }
  for (const flag of [rule.style.bold, rule.style.italic, rule.style.underline, rule.style.strike, rule.isBottom, rule.isPercent,
    rule.config?.isGradient, rule.config?.isShowValue, rule.config?.reverse]) {
    if (flag !== undefined && typeof flag !== 'boolean') throw new TypeError('条件格式样式和显示开关必须为布尔值。')
  }
  const officeSource = normalizeConditionalFormatOfficeSource(rule.officeSource)
  if (rule.type === 'office') {
    if (!officeSource) throw new TypeError('源文件条件格式缺少往返来源。')
    const baseline = conditionalFormatOfficeBaseline(previous ?? officeSource.baseline)
    const candidate = conditionalFormatOfficeBaseline(rule)
    for (const field of ['type', 'subType', 'operator', 'formula1', 'formula2', 'value', 'isBottom', 'isPercent', 'config'] as const) {
      if (signature(candidate[field]) !== signature(baseline[field])) throw new TypeError('原有条件规则的条件需要保留；可调整应用范围、样式和顺序。')
    }
    return
  }
  const role = formulaRole(rule), previousRole = previous && formulaRole(previous)
  if (role === 'number') {
    if (!numberOperators.includes(rule.operator)) throw new TypeError('条件格式数值比较条件无效。')
    assertFormula(rule.formula1, previousRole === role ? previous?.formula1 : undefined)
    if (rule.operator === 'between' || rule.operator === 'notBetween') {
      if (rule.formula2 === undefined) throw new TypeError('介于条件需要第二个数值或公式。')
    }
    if (rule.formula2 !== undefined) assertFormula(rule.formula2, previousRole === role ? previous?.formula2 : undefined)
  } else if (role === 'expression') {
    assertFormula(rule.formula1 || String(rule.value ?? ''), previousRole === role ? previous?.formula1 || String(previous?.value ?? '') : undefined)
  }
  if (rule.type === 'top10' || rule.subType === 'rank') {
    const count = rule.value ?? 10
    const numeric = typeof count === 'number' || typeof count === 'string' && /^\d+$/.test(count.trim()) ? Number(count) : Number.NaN
    if (!Number.isSafeInteger(numeric) || numeric < 1 || numeric > (rule.isPercent ? 100 : 1000)) throw new RangeError('条件格式排名项数无效。')
  }
  if (rule.config !== undefined && (!rule.config || typeof rule.config !== 'object' || Array.isArray(rule.config))) throw new TypeError('条件格式配置无效。')
  const config = rule.config, previousConfig = previous?.config
  if (config?.min !== undefined) assertThreshold(config.min, previousConfig?.min, rule.type === 'dataBar' ? 'min' : undefined)
  if (config?.max !== undefined) assertThreshold(config.max, previousConfig?.max, rule.type === 'dataBar' ? 'max' : undefined)
  if (rule.type === 'colorScale' && (!Array.isArray(config?.thresholds) || config.thresholds.length < 2 || config.thresholds.length > 3)) throw new RangeError('色阶需要两个或三个停止点。')
  if (config?.thresholds !== undefined) {
    if (!Array.isArray(config.thresholds)) throw new TypeError('色阶停止点必须为数组。')
    for (const [index, entry] of config.thresholds.entries()) {
      if (!entry || typeof entry !== 'object' || typeof entry.color !== 'string' || !/^#[0-9a-f]{6}$/i.test(entry.color)) throw new TypeError('色阶停止点颜色无效。')
      assertThreshold(entry.value, previousConfig?.thresholds?.[index]?.value)
    }
  }
  if (rule.type === 'iconSet' && signature(config) !== signature(previousConfig)) validateConditionalIconConfig(config ?? {})
  if (config?.icons !== undefined) {
    if (!Array.isArray(config.icons)) throw new TypeError('条件格式图标必须为数组。')
    for (const [index, entry] of config.icons.entries()) {
      if (!entry || typeof entry !== 'object') throw new TypeError('条件格式图标配置无效。')
      assertThreshold(entry.value, previousConfig?.icons?.[index]?.value)
    }
  }
}

function orderedRules(rules: readonly SheetConditionalFormatRule[]): SheetConditionalFormatRule[] {
  return rules.map(rule => cloneModel(rule)).sort((left, right) => left.priority - right.priority)
}

/** Changed rules touch both their old and new scope; tied priorities also retain stable stack order. */
export function planSheetConditionalFormatsCommand(sheet: Pick<SpreadsheetSheet, 'conditionalFormats'>, rules: readonly SheetConditionalFormatRule[]): SheetConditionalFormatsPlan {
  if (!Array.isArray(rules)) throw new TypeError('条件格式规则必须为数组。')
  assertConditionalFormatModelBudget(rules)
  const previous = orderedRules(sheet.conditionalFormats ?? []), previousById = new Map(previous.map(rule => [rule.id, rule]))
  const ids = new Set<string>()
  for (const rule of rules) {
    assertSheetConditionalFormatRuleSnapshot(rule, previousById.get(rule?.id))
    if (ids.has(rule.id)) throw new TypeError('条件格式规则标识不能重复。')
    ids.add(rule.id)
  }
  const next = orderedRules(rules), nextById = new Map(next.map(rule => [rule.id, rule]))
  const changed = signature(previous) !== signature(next), affected = new Set<string>()
  for (const rule of previous) if (!nextById.has(rule.id) || signature(rule) !== signature(nextById.get(rule.id))) affected.add(rule.id)
  for (const rule of next) if (!previousById.has(rule.id)) affected.add(rule.id)
  const beforeGroups = new Map<number, string[]>(), afterGroups = new Map<number, string[]>()
  for (const [source, other, groups] of [[previous, nextById, beforeGroups], [next, previousById, afterGroups]] as const) {
    for (const rule of source) {
      if (other.get(rule.id)?.priority !== rule.priority) continue
      const group = groups.get(rule.priority) ?? []
      group.push(rule.id)
      groups.set(rule.priority, group)
    }
  }
  for (const [priority, before] of beforeGroups) {
    const after = afterGroups.get(priority) ?? []
    before.forEach((id, index) => { if (id !== after[index]) { affected.add(id); if (after[index]) affected.add(after[index]!) } })
  }
  const uniqueRanges = new Map<string, CellRange>()
  for (const id of affected) {
    for (const range of [...(previousById.get(id)?.ranges ?? []), ...(nextById.get(id)?.ranges ?? [])]) {
      assertSheetCommandRanges([range])
      uniqueRanges.set(`${range.startRow}:${range.startCol}:${range.endRow}:${range.endCol}`, { ...range })
    }
  }
  return { rules: next, ranges: [...uniqueRanges.values()], changed }
}

/** Standalone and registered commands share the same atomic rule and permission contract. */
export function applySheetConditionalFormatsCommand(sheet: ConditionalFormatSheet, rules: readonly SheetConditionalFormatRule[], userId: string): SheetConditionalFormatsPlan {
  if (!canViewWorksheetProtection(sheet, userId)) throw new Error('该工作表受保护，当前用户没有查看权限。')
  const plan = planSheetConditionalFormatsCommand(sheet, rules)
  assertSheetRangesEditable(sheet, plan.ranges, userId)
  if (plan.changed) sheet.conditionalFormats = cloneModel(plan.rules)
  return plan
}
