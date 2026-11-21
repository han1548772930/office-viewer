import type { CellRange, SheetDataValidationRule } from '../../domain/model.ts'
import { createSheetDataValidationRule } from './validation-engine.ts'

export type QuickSheetDataValidationPreset = 'checkbox' | 'dropdown' | 'date' | 'date-time'

export interface QuickSheetDataValidationOptions {
  dropdownValues?: readonly string[]
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
  const formula1 = `"${values.map(value => value.replaceAll('"', '""')).join(',')}"`
  return createSheetDataValidationRule({ type: 'list', formula1, renderMode: 'custom' }, ranges)
}
