import type { CellRange, SheetDataValidationRule, SheetDataValidationType, SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { assertSheetCommandRanges, normalizeSheetCommandRange, sheetRangesIntersect, type SheetCommandRange } from '../sheet-workbench/sheet-range-command.ts'
import { applySheetDataValidationRule, clearSheetDataValidationRanges, clearSheetDataValidationTypes,
  createSheetDataValidationRule, removeSheetDataValidationRule, replaceSheetDataValidationRule } from './validation-engine.ts'

export const SHEET_DATA_VALIDATION_COMMANDS = {
  set: 'sheet.command.set-data-validation', replace: 'sheet.command.replace-data-validation', remove: 'sheet.command.remove-data-validation',
  clear: 'sheet.command.clear-data-validation', clearTypes: 'sheet.command.clear-quick-data-validation',
} as const
export const SHEET_DATA_VALIDATION_ASYNC_COMMANDS = {
  set: 'sheet.command.set-data-validation-async', replace: 'sheet.command.replace-data-validation-async',
} as const
export type SheetDataValidationAction = keyof typeof SHEET_DATA_VALIDATION_COMMANDS
export interface SheetDataValidationCommandTarget {
  unitId: string; sheetId: string
  host?: 'api' | 'workbench'; expectedRevision?: number; expectedHistoryGeneration?: number; expectedUserId?: string
  expectedRulesSignature?: string
}
export interface SheetSetDataValidationInput extends SheetDataValidationCommandTarget { rule: SheetDataValidationRule }
export interface SheetRemoveDataValidationInput extends SheetDataValidationCommandTarget { ruleId: string }
export interface SheetClearDataValidationInput extends SheetDataValidationCommandTarget { ranges: readonly SheetCommandRange[] }
export interface SheetClearDataValidationTypesInput extends SheetClearDataValidationInput { types: readonly SheetDataValidationType[] }
export type SheetDataValidationCommandInput = SheetDataValidationCommandTarget & (
  { action: 'set' | 'replace'; rule: SheetDataValidationRule } | { action: 'remove'; ruleId: string }
  | { action: 'clear'; ranges: CellRange[] } | { action: 'clearTypes'; ranges: CellRange[]; types: SheetDataValidationType[] })
export interface SheetDataValidationRulesPlan { rules: SheetDataValidationRule[]; ranges: CellRange[]; changed: boolean }
export type SheetDataValidationSubmitResult = { accepted: true; changed: boolean } | { accepted: false; error: string }
const TYPES: readonly SheetDataValidationType[] = ['any', 'whole', 'decimal', 'textLength', 'date', 'time', 'checkbox', 'list', 'listMultiple', 'custom']

export function sheetDataValidationCommandRecord(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('数据验证命令参数必须是对象。')
  return value as Record<string, unknown>
}
function requiredText(value: unknown, label: string): string {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${label}无效。`)
  return value
}
export function sheetDataValidationCommandTarget(source: Record<string, unknown>): SheetDataValidationCommandTarget {
  const input: SheetDataValidationCommandTarget = { unitId: requiredText(source.unitId, '工作簿标识'), sheetId: requiredText(source.sheetId, '工作表标识') }
  const host = source.host, userId = source.expectedUserId, signature = source.expectedRulesSignature
  if (host !== undefined) { if (host !== 'api' && host !== 'workbench') throw new TypeError('数据验证宿主无效。'); input.host = host }
  for (const field of ['expectedRevision', 'expectedHistoryGeneration'] as const) {
    const value = source[field]
    if (value === undefined) continue
    if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < (field === 'expectedRevision' ? 0 : 1)) throw new TypeError('数据验证内容或历史版本无效。')
    input[field] = value
  }
  if (userId !== undefined) input.expectedUserId = requiredText(userId, '数据验证用户')
  if (signature !== undefined) input.expectedRulesSignature = requiredText(signature, '数据验证规则版本')
  return input
}

/** Payloads are detached once. API adapters capture authority before invoking a payload getter or JSON hook. */
export function sheetDataValidationCommandPayload(action: SheetDataValidationAction, source: Record<string, unknown>,
  target: SheetDataValidationCommandTarget): SheetDataValidationCommandInput {
  if (!Object.keys(SHEET_DATA_VALIDATION_COMMANDS).includes(action)) throw new TypeError('数据验证命令动作无效。')
  if (action === 'set' || action === 'replace') {
    const value = source.rule
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('数据验证命令需要规则。')
    const rule = cloneModel(value) as SheetDataValidationRule
    if (!Array.isArray(rule.ranges) || !rule.ranges.length) throw new RangeError('数据验证需要有效的应用区域。')
    rule.ranges = rule.ranges.map(normalizeSheetCommandRange)
    return { ...target, action, rule }
  }
  if (action === 'remove') return { ...target, action, ruleId: requiredText(source.ruleId, '数据验证规则标识') }
  const value = source.ranges
  if (!Array.isArray(value)) throw new TypeError('数据验证范围必须是数组。')
  const ranges = value.map(range => normalizeSheetCommandRange(range))
  if (!ranges.length) throw new RangeError('请选择至少一个数据验证区域。')
  if (action === 'clear') return { ...target, action, ranges }
  const rawTypes = source.types
  if (!Array.isArray(rawTypes) || !rawTypes.length || rawTypes.some(type => !TYPES.includes(type))) throw new TypeError('要清除的数据验证类型无效。')
  return { ...target, action, ranges, types: [...new Set(rawTypes)] }
}
export function sheetDataValidationCommandInput(action: SheetDataValidationAction, value: unknown): SheetDataValidationCommandInput {
  const source = sheetDataValidationCommandRecord(value)
  return sheetDataValidationCommandPayload(action, source, sheetDataValidationCommandTarget(source))
}
function signature(value: unknown): string {
  return JSON.stringify(value, (_key, item) => item && typeof item === 'object' && !Array.isArray(item)
    ? Object.fromEntries(Object.entries(item).sort(([left], [right]) => left.localeCompare(right))) : item) ?? ''
}

function completeRuleRanges(sheet: SpreadsheetSheet, ranges: readonly CellRange[]): CellRange[] {
  const complete = [...ranges, ...sheet.mergedRanges.filter(merge => ranges.some(range => sheetRangesIntersect(range, merge)))]
  return [...new Map(complete.map(range => [JSON.stringify(range), { ...range }])).values()]
}

/** Rule planning is cheap enough for capability checks; scalar initialization is prepared only after authorization. */
export function planSheetDataValidationCommand(sheet: SpreadsheetSheet, input: SheetDataValidationCommandInput): SheetDataValidationRulesPlan {
  if (!Object.keys(SHEET_DATA_VALIDATION_COMMANDS).includes(input.action)) throw new TypeError('数据验证命令动作无效。')
  const previous = sheet.dataValidations ?? [], candidate: SpreadsheetSheet = { ...sheet, dataValidations: cloneModel(previous) }
  let ranges: CellRange[]
  if (input.action === 'set' || input.action === 'replace') {
    const canonicalId = typeof input.rule.id === 'string' ? input.rule.id.trim() : undefined
    const current = previous.find(rule => rule.id === canonicalId)
    const rule = createSheetDataValidationRule(input.rule, input.rule.ranges, current)
    ranges = [...(current?.ranges ?? []), ...rule.ranges].map(range => ({ ...range }))
    if (current && signature(current) === signature(rule)) return { rules: cloneModel(previous), ranges: completeRuleRanges(sheet, ranges), changed: false }
    if (input.action === 'set') applySheetDataValidationRule(candidate, rule)
    else replaceSheetDataValidationRule(candidate, rule)
  } else if (input.action === 'remove') {
    ranges = previous.find(rule => rule.id === input.ruleId)?.ranges.map(range => ({ ...range })) ?? []
    removeSheetDataValidationRule(candidate, input.ruleId)
  } else {
    ranges = input.ranges.map(range => ({ ...range }))
    if (input.action === 'clear') clearSheetDataValidationRanges(candidate, ranges)
    else clearSheetDataValidationTypes(candidate, ranges, input.types)
  }
  assertSheetCommandRanges(ranges)
  const rules = candidate.dataValidations ?? []
  return { rules, ranges: completeRuleRanges(sheet, ranges), changed: signature(previous) !== signature(rules) }
}
