import type { CellRange, SheetDataValidationRule } from '../../domain/model.ts'
import { createSheetDataValidationRule, serializeValidationListItems } from './validation-engine.ts'

export type QuickSheetDataValidationPreset = 'checkbox' | 'dropdown' | 'date' | 'date-time'

export interface QuickSheetDataValidationOptions {
  dropdownValues?: readonly string[]
}

export function quickDropdownPresets(t: (key: string, fallback: string) => string) {
  const label = (key: string, fallback: string) => t(`sheet.validation.preset.${key}`, fallback)
  return {
    'dropdown-yes-no': [label('yes', '是'), label('no', '否')],
    'dropdown-status': [label('notStarted', '未开始'), label('inProgress', '进行中'), label('completed', '已完成')],
    'dropdown-abc': ['A', 'B', 'C'],
    'edit-dropdown': [label('option1', '选项 1'), label('option2', '选项 2')],
  }
}

export function planQuickValidation(action: unknown, ranges: readonly CellRange[], t: (key: string, fallback: string) => string) {
  if (!ranges.length) return
  if (action === 'add-default') {
    return { kind: 'set' as const, rule: createSheetDataValidationRule({ type: 'decimal', operator: 'equal', formula1: '100' }, ranges), edit: true }
  }
  if (action === 'clear-checkbox' || action === 'clear-dropdown' || action === 'clear-date') {
    const types = action === 'clear-checkbox' ? ['checkbox'] as const : action === 'clear-date' ? ['date'] as const : ['list', 'listMultiple'] as const
    return { kind: 'clear' as const, ranges, types }
  }
  if (action === 'checkbox' || action === 'date' || action === 'date-time' || action === 'edit-date') {
    return { kind: 'set' as const, rule: createQuickSheetDataValidationRule(action === 'edit-date' ? 'date' : action, ranges), edit: action === 'edit-date' }
  }
  if (action === 'dropdown-yes-no' || action === 'dropdown-status' || action === 'dropdown-abc' || action === 'edit-dropdown') {
    return { kind: 'set' as const, rule: createQuickSheetDataValidationRule('dropdown', ranges, { dropdownValues: quickDropdownPresets(t)[action] }), edit: action === 'edit-dropdown' }
  }
}

export function createQuickSheetDataValidationRule(
  preset: QuickSheetDataValidationPreset,
  ranges: readonly CellRange[],
  options: QuickSheetDataValidationOptions = {},
): SheetDataValidationRule {
  if (preset === 'checkbox') return createSheetDataValidationRule({ type: 'checkbox' }, ranges)
  if (preset === 'date' || preset === 'date-time') {
    return createSheetDataValidationRule({
      type: 'date',
      ...(preset === 'date-time' ? { metadata: { showTime: true } } : {}),
    }, ranges)
  }

  const values = options.dropdownValues?.map(value => value.trim()).filter(Boolean) ?? []
  if (values.length === 0) throw new RangeError('下拉列表至少需要一个选项。')
  return createSheetDataValidationRule({ type: 'list', formula1: serializeValidationListItems(values), renderMode: 'custom' }, ranges)
}
