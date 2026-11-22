import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { inferSheetCellValueType, parseSheetCellText } from '../../domain/sheet-cell-value.ts'
import { formatSheetValue } from './sheet-number-format.ts'
import { sheetCellNumberPattern } from './sheet-number-presentation.ts'

/** A committed edit replaces content and cache types while retaining the cell's formatting. */
export function editSheetCellText(cell: ExcelCellView | undefined, input: string, date1904 = false, region = 'zh-CN'): ExcelCellView {
  const current = cell ?? { value: '' }
  const pattern = sheetCellNumberPattern(current)
  const forcedText = input.startsWith("'")
  const textFormat = pattern === '@'
  if (!forcedText && !textFormat && input.startsWith('=')) {
    return { ...current, value: '', formula: input, dataType: 'formula', cachedDataType: null,
      rawValue: null, cachedValue: undefined, displayValue: input }
  }
  const value = forcedText ? input.slice(1) : input === '' ? null : textFormat ? input
    : /^(?:true|false)$/i.test(input.trim()) ? parseSheetCellText(input, 'boolean') : parseSheetCellText(input)
  return { ...current, value: value === null ? '' : String(value), formula: undefined,
    dataType: forcedText ? 'force_string' : inferSheetCellValueType(value), cachedDataType: null,
    rawValue: value, cachedValue: undefined,
    displayValue: formatSheetValue(value, pattern, date1904, region) ?? (value === null ? '' : String(value)) }
}
