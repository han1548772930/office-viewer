import type { CellScalar, SheetDataValidationRule, WorkbookModel } from '../../domain/model.ts'
import { sheetOpaqueColorHex } from '../sheet-workbench/sheet-colors.ts'
import type { SheetCellInputDefaults, SheetCellInputPlan } from '../sheet-workbench/sheet-cell-input-command.ts'
import { parseValidationListItems, serializeValidationListItems, sheetDataValidationOptions, validateSheetDataValidationChanges } from './validation-engine.ts'
import { sheetValidationSourceReader } from './validation-value-read.ts'
import { planSheetValidationWidgetValue, resolveSheetValidationWidgetTarget, sheetValidationWidgetTargetInput,
  validationWidgetRecord, type SheetValidationWidgetTargetInput } from './validation-widget-input.ts'

export const COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND = 'sheet.command.commit-validation-list-input'
export interface SheetValidationListOption { value: string; label: string; color?: string }
export interface SheetValidationListInput extends SheetValidationWidgetTargetInput { values: string[] }
function key(value: string): string { return value.trim().toLocaleLowerCase() }

export function sheetValidationListOptions(rule: SheetDataValidationRule, workbook: WorkbookModel, sheetId: string, userId?: string, region = 'zh-CN'): SheetValidationListOption[] {
  const colors = rule.formula2?.split(',') ?? []
  const seen = new Set<string>()
  return sheetDataValidationOptions(rule, workbook, sheetId, sheetValidationSourceReader(userId), region).flatMap((label, index) => {
    const normalized = key(label)
    if (seen.has(normalized)) return []
    seen.add(normalized)
    const color = sheetOpaqueColorHex(colors[index] ?? '')
    return [{ value: label, label, ...(color ? { color } : {}) }]
  })
}
/** Stored CSV is read with the formula wrapper; candidate spelling, order and uniqueness remain authoritative. */
export function sheetValidationListSelected(options: readonly SheetValidationListOption[], value: CellScalar | undefined, multiple: boolean): string[] {
  const raw = String(value ?? ''), items = multiple ? parseValidationListItems(`"${raw}"`) : raw === '' ? [] : [raw]
  const selected = new Set(items.map(key))
  return options.filter(option => selected.has(key(option.value))).map(option => option.value).slice(0, multiple ? undefined : 1)
}
export function sheetValidationListValue(values: readonly string[], multiple: boolean): string {
  return multiple ? serializeValidationListItems(values).slice(1, -1) : values[0] ?? ''
}
export function sheetValidationListInput(value: unknown): SheetValidationListInput {
  const source = validationWidgetRecord(value, '列表输入命令参数'), values = source.values
  if (!Array.isArray(values) || values.some(item => typeof item !== 'string')) throw new TypeError('列表选择必须是文本数组。')
  return { ...sheetValidationWidgetTargetInput(source), values: [...values] }
}
export function planSheetValidationListInput(workbook: WorkbookModel, input: SheetValidationListInput,
  defaults: SheetCellInputDefaults): SheetCellInputPlan {
  const target = resolveSheetValidationWidgetTarget(workbook, input, defaults)
  if ('result' in target) return target
  if (target.rule.type !== 'list' && target.rule.type !== 'listMultiple') throw new TypeError('此规则没有列表选择控件。')
  const multiple = target.rule.type === 'listMultiple', options = sheetValidationListOptions(target.rule, workbook, target.sheet.id, defaults.userId, target.region)
  const selected = new Set(input.values.map(key)), allowed = new Set(options.map(option => key(option.value)))
  if ([...selected].some(item => !allowed.has(item))) throw new TypeError('列表候选已更改，所选项不在当前列表中。')
  if (!multiple && selected.size > 1) throw new TypeError('单选列表只能选择一个值。')
  const values = options.filter(option => selected.has(key(option.value))).map(option => option.value)
  const value = sheetValidationListValue(values, multiple), previous = target.sheet.cells[`${target.position.row}:${target.position.column}`]
  if (previous?.value === value && !previous.formula && !previous.richText?.length && !previous.image) {
    const validation = validateSheetDataValidationChanges(workbook, target.sheet.id, [target.position], sheetValidationSourceReader(defaults.userId), target.region)
    const issue = validation.issues.find(item => item.severity === 'stop')
    return { position: target.position, result: { accepted: !validation.blocked, changed: false, issues: validation.issues,
      ...(issue ? { reason: 'validation' as const, message: `${issue.title}：${issue.message}` } : {}) } }
  }
  return planSheetValidationWidgetValue(target, value, 'text')
}
