import { validationWidgetRecord } from './validation-widget-input.ts'

export const SHEET_VALIDATION_DROPDOWN_OPERATIONS = { show: 'sheet.operation.show-data-validation-dropdown', hide: 'sheet.operation.hide-data-validation-dropdown' } as const
export interface SheetValidationDropdownLocation { unitId: string; sheetId: string; row: number; column: number }
export interface SheetValidationDropdownRequest extends SheetValidationDropdownLocation {
  ruleId: string; expectedRuleSignature: string; expectedRevision: number; expectedHistoryGeneration: number; expectedUserId: string
}
export interface SheetValidationDropdownShowInput extends SheetValidationDropdownLocation {
  expectedRevision?: number; expectedHistoryGeneration?: number; expectedUserId?: string
}
export interface SheetValidationDropdownHideInput { unitId?: string; sheetId?: string }

function identifier(value: unknown, label: string): string {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${label}无效。`)
  return value
}
export function sheetValidationDropdownUnit(value: unknown): Pick<SheetValidationDropdownLocation, 'unitId' | 'sheetId'> {
  const source = validationWidgetRecord(value, '下拉控件操作参数'), unitId = identifier(source.unitId, '工作簿标识')
  const sheetId = source.sheetId, subUnitId = source.subUnitId
  if (sheetId !== undefined && subUnitId !== undefined && sheetId !== subUnitId) throw new TypeError('下拉控件工作表标识不一致。')
  return { unitId, sheetId: identifier(sheetId ?? subUnitId, '工作表标识') }
}
export function sheetValidationDropdownShowInput(value: unknown,
  target = sheetValidationDropdownUnit(value)): SheetValidationDropdownShowInput {
  const source = validationWidgetRecord(value, '下拉控件操作参数'), row = source.row, column = source.column
  if (typeof row !== 'number' || !Number.isSafeInteger(row) || row < 0 || row >= 1_048_576
    || typeof column !== 'number' || !Number.isSafeInteger(column) || column < 0 || column >= 16_384) throw new RangeError('下拉控件坐标无效。')
  const result: SheetValidationDropdownShowInput = { ...target, row, column }
  for (const field of ['expectedRevision', 'expectedHistoryGeneration'] as const) {
    const version = source[field]
    if (version !== undefined) {
      if (typeof version !== 'number' || !Number.isSafeInteger(version) || version < (field === 'expectedRevision' ? 0 : 1)) throw new TypeError('下拉控件版本无效。')
      result[field] = version
    }
  }
  if (source.expectedUserId !== undefined) result.expectedUserId = identifier(source.expectedUserId, '下拉控件用户')
  return result
}
export function sheetValidationDropdownHideInput(value: unknown): SheetValidationDropdownHideInput {
  const source = validationWidgetRecord(value, '关闭下拉控件操作参数'), result: SheetValidationDropdownHideInput = {}
  const unitId = source.unitId, sheetId = source.sheetId, subUnitId = source.subUnitId
  if (sheetId !== undefined && subUnitId !== undefined && sheetId !== subUnitId) throw new TypeError('下拉控件工作表标识不一致。')
  if (unitId !== undefined) result.unitId = identifier(unitId, '工作簿标识')
  if (sheetId !== undefined || subUnitId !== undefined) result.sheetId = identifier(sheetId ?? subUnitId, '工作表标识')
  return result
}
