import type { CellScalar, SheetDataValidationRule, WorkbookModel } from '../../domain/model.ts'
import { classifySheetNumberFormat } from '../sheet-workbench/sheet-number-format.ts'
import { sheetCellNumberPattern, sheetCellNumberValue } from '../sheet-workbench/sheet-number-presentation.ts'
import { DateSystem } from '../sheet-workbench/number-format-engine/date-system.ts'
import { excelDateTimePartsToSerial, excelSerialToDateTimeParts, getDateSystemMaxSerial,
  type IExcelDateTimeParts } from '../sheet-workbench/number-format-engine/serial-date.ts'
import type { SheetCellInputDefaults, SheetCellInputPlan } from '../sheet-workbench/sheet-cell-input-command.ts'
import { planSheetValidationWidgetValue, resolveSheetValidationWidgetTarget, sheetValidationWidgetTargetInput,
  validationWidgetRecord, type SheetValidationWidgetTargetInput } from './validation-widget-input.ts'

export const COMMIT_SHEET_DATE_INPUT_COMMAND = 'sheet.command.commit-date-input'
export type SheetDateInputMode = 'date' | 'datetime' | 'time' | 'duration'
export type SheetDateInputChange = { kind: 'unchanged' }
  | { kind: 'date'; year: number; month: number; day: number }
  | { kind: 'time'; hours: number; minutes: number; seconds: number; fractionalSecond: number }
  | ({ kind: 'date-time' } & IExcelDateTimeParts)
  | { kind: 'serial'; serial: number }
export interface SheetDateInputCommandInput extends SheetValidationWidgetTargetInput { change: SheetDateInputChange }
export interface SheetDateInputState {
  mode: SheetDateInputMode
  originalSerial: number | null
  parts: IExcelDateTimeParts | null
  date1904: boolean
  numberFormat: string
}
const DEFAULT_PATTERNS: Readonly<Record<SheetDateInputMode, string>> = {
  date: 'yyyy-MM-dd', datetime: 'yyyy-MM-dd hh:mm:ss', time: 'HH:mm:ss', duration: '[h]:mm:ss',
}
function calendar(date1904: boolean): DateSystem { return date1904 ? DateSystem.Date1904 : DateSystem.Date1900 }
export function sheetDateInputSerial(value: CellScalar | undefined): number | null {
  const serial = typeof value === 'number' ? value : typeof value === 'string' && value.trim() ? Number(value) : null
  return serial !== null && Number.isFinite(serial) ? serial : null
}
export function sheetDateInputMode(rule: SheetDataValidationRule, pattern: string): SheetDateInputMode {
  if (rule.type === 'date') return rule.metadata?.showTime === true ? 'datetime' : 'date'
  if (rule.type === 'time') return classifySheetNumberFormat(pattern) === 'duration' ? 'duration' : 'time'
  throw new TypeError('此数据验证规则没有日期或时间选择控件。')
}
export function sheetDateInputState(rule: SheetDataValidationRule, value: CellScalar | undefined,
  numberFormat: string, date1904: boolean): SheetDateInputState {
  const originalSerial = sheetDateInputSerial(value), mode = sheetDateInputMode(rule, numberFormat)
  return { mode, originalSerial, parts: originalSerial === null ? null : excelSerialToDateTimeParts(originalSerial, { dateSystem: calendar(date1904) }),
    date1904, numberFormat }
}
export function sheetDateInputFormat(mode: SheetDateInputMode, original: string): string {
  return classifySheetNumberFormat(original) === mode ? original : DEFAULT_PATTERNS[mode]
}

export function sheetDateInputChange(value: unknown): SheetDateInputChange {
  const source = validationWidgetRecord(value, '日期修改'), kind = source.kind
  if (kind === 'unchanged') return { kind }
  if (kind === 'serial') {
    const serial = source.serial
    if (typeof serial !== 'number' || !Number.isFinite(serial)) throw new TypeError('日期 serial 必须是有限数值。')
    return { kind, serial }
  }
  if (kind !== 'date' && kind !== 'time' && kind !== 'date-time') throw new TypeError('日期修改类型无效。')
  const fields = kind === 'date' ? ['year', 'month', 'day'] as const : kind === 'time'
    ? ['hours', 'minutes', 'seconds', 'fractionalSecond'] as const
    : ['year', 'month', 'day', 'hours', 'minutes', 'seconds', 'fractionalSecond'] as const
  const result: Record<string, number | string> = { kind }
  for (const key of fields) {
    const supplied = source[key], field = supplied === undefined && key === 'fractionalSecond' ? 0 : supplied
    if (typeof field !== 'number' || !Number.isFinite(field) || key !== 'fractionalSecond' && !Number.isSafeInteger(field)) throw new TypeError('日期或时间字段无效。')
    result[key] = field
  }
  for (const [key, maximum] of [['hours', 23], ['minutes', 59], ['seconds', 59]] as const) {
    const field = result[key]
    if (field !== undefined && (typeof field !== 'number' || field < 0 || field > maximum)) throw new RangeError('时间字段超出范围。')
  }
  if (typeof result.fractionalSecond === 'number' && (result.fractionalSecond < 0 || result.fractionalSecond >= 1)) throw new RangeError('小数秒超出范围。')
  return result as unknown as SheetDateInputChange
}
export function sheetDateInputCommandInput(value: unknown): SheetDateInputCommandInput {
  const source = validationWidgetRecord(value, '日期输入命令参数')
  return { ...sheetValidationWidgetTargetInput(source), change: sheetDateInputChange(source.change) }
}

/** Gregorian conversion is owned by the existing serial engine. Native Date/Dayjs only adapt controls. */
export function resolveSheetDateInputChange(state: SheetDateInputState, change: SheetDateInputChange): { changed: boolean; serial: number | null } {
  const { originalSerial, parts, mode, date1904 } = state, dateSystem = calendar(date1904)
  if (change.kind === 'unchanged') return { changed: false, serial: originalSerial }
  let serial: number | null
  if (change.kind === 'serial') serial = change.serial
  else {
    if (mode === 'duration') throw new TypeError('持续时间需要通过数值或持续时间输入修改。')
    if (originalSerial !== null && !parts) throw new RangeError('原日期超出日历范围，请使用数值输入纠正。')
    if (change.kind === 'date') {
      if (mode === 'time') throw new TypeError('时间选择控件不能修改日期。')
      if (parts && parts.year === change.year && parts.month === change.month && parts.day === change.day) return { changed: false, serial: originalSerial }
      const day = excelDateTimePartsToSerial({ ...change, hours: 0, minutes: 0, seconds: 0, fractionalSecond: 0 }, { dateSystem })
      serial = day === null ? null : day + (originalSerial === null ? 0 : originalSerial - Math.floor(originalSerial))
    } else if (change.kind === 'time') {
      if (mode === 'date') throw new TypeError('此日期选择控件没有启用时间。')
      if (parts && parts.hours === change.hours && parts.minutes === change.minutes && parts.seconds === change.seconds
        && parts.fractionalSecond === change.fractionalSecond) return { changed: false, serial: originalSerial }
      const anchor = parts ?? { year: date1904 ? 1904 : 1900, month: 1, day: date1904 ? 1 : 0 }
      serial = excelDateTimePartsToSerial({ ...anchor, ...change }, { dateSystem })
    } else {
      if (mode === 'time') throw new TypeError('时间选择控件不能修改日期。')
      if (parts && (['year', 'month', 'day', 'hours', 'minutes', 'seconds', 'fractionalSecond'] as const).every(key => parts[key] === change[key])) return { changed: false, serial: originalSerial }
      serial = excelDateTimePartsToSerial(mode === 'date' ? { ...change, hours: 0, minutes: 0, seconds: 0, fractionalSecond: 0 } : change, { dateSystem })
      if (mode === 'date' && serial !== null) serial += originalSerial === null ? 0 : originalSerial - Math.floor(originalSerial)
    }
  }
  if (serial === null || !Number.isFinite(serial) || serial < 0 || mode !== 'duration' && serial >= getDateSystemMaxSerial(dateSystem) + 1) throw new RangeError('日期或时间超出当前工作簿支持的范围。')
  return { changed: !Object.is(serial, originalSerial), serial }
}

export function planSheetDateInput(workbook: WorkbookModel, input: SheetDateInputCommandInput,
  defaults: SheetCellInputDefaults): SheetCellInputPlan {
  const target = resolveSheetValidationWidgetTarget(workbook, input, defaults)
  if ('result' in target) return target
  const state = sheetDateInputState(target.rule, sheetCellNumberValue(target.view), sheetCellNumberPattern(target.view), target.date1904)
  const change = resolveSheetDateInputChange(state, input.change)
  if (!change.changed) return { position: target.position, result: { accepted: true, changed: false, issues: [] } }
  return planSheetValidationWidgetValue(target, change.serial, 'number', sheetDateInputFormat(state.mode, state.numberFormat))
}
