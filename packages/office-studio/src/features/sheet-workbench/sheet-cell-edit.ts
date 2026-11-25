import { cellKey, parseCellKey, type ExcelCellView } from '../../services/excelViewAdapter.ts'
import { inferSheetCellValueType, isForcedSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { classifySheetNumberFormat, formatSheetValue, isSheetTextNumberFormat } from './sheet-number-format.ts'
import { sheetCellNumberPattern, sheetCellNumberValue } from './sheet-number-presentation.ts'
import { parseSheetNumberEdit, sheetNumberEditText } from './sheet-number-format-edit.ts'
import { normalizeInputDigits } from './number-format-engine/number-input.ts'

/** Visible merged primaries can differ from the cell that owns their persisted content. */
export function sheetCellContentKey(sheet: { merges?: readonly string[] }, key: string): string {
  const point = parseCellKey(key)
  if (!point) return key
  for (const merge of sheet.merges ?? []) {
    const [first, last = first] = merge.split(':')
    const start = parseCellKey(first ?? ''), end = parseCellKey(last ?? '')
    if (!start || !end) continue
    const startRow = Math.min(start.row, end.row), endRow = Math.max(start.row, end.row)
    const startColumn = Math.min(start.col, end.col), endColumn = Math.max(start.col, end.col)
    if (point.row >= startRow && point.row <= endRow && point.col >= startColumn && point.col <= endColumn) return cellKey(startRow, startColumn)
  }
  return cellKey(point.row, point.col)
}

export function sheetCellInputText(cell: ExcelCellView | undefined, date1904 = false, region = 'zh-CN'): string {
  if (!cell) return ''
  if (cell.formula) return cell.formula
  const pattern = sheetCellNumberPattern(cell)
  if (isForcedSheetCellValueType(cell.dataType) && !isSheetTextNumberFormat(pattern)) return "'" + cell.value
  const value = sheetCellNumberValue(cell)
  return typeof value === 'number' ? sheetNumberEditText(value, pattern, { date1904, region }) ?? cell.value : cell.value
}

export interface SheetCellEditorStart {
  value: string
  selection: { start: number; end: number }
}

/** A percentage's retained suffix is outside the initial caret/selection and never applies to formulas or text. */
export function createSheetCellEditorStart(cell: ExcelCellView | undefined, initialValue?: string, date1904 = false, region = 'zh-CN'): SheetCellEditorStart {
  const text = sheetCellInputText(cell, date1904, region)
  const isPercent = !cell?.formula && typeof sheetCellNumberValue(cell) === 'number' && text.endsWith('%')
  const retainSuffix = isPercent && (initialValue === undefined || /^\d/.test(normalizeInputDigits(initialValue)))
  const value = initialValue === undefined ? text : retainSuffix ? initialValue + '%' : initialValue
  const end = value.length - (retainSuffix ? 1 : 0)
  return { value, selection: { start: end, end } }
}

/** A committed edit replaces content and cache types while retaining the cell's formatting. */
export function editSheetCellText(cell: ExcelCellView | undefined, input: string, date1904 = false, region = 'zh-CN', currentYear?: number): ExcelCellView {
  const current = cell ?? { value: '' }
  const pattern = sheetCellNumberPattern(current)
  const textFormat = isSheetTextNumberFormat(pattern)
  const forcedText = !textFormat && input.startsWith("'")
  if (!forcedText && !textFormat && input.startsWith('=')) {
    return { ...current, value: '', formula: input, dataType: 'formula', cachedDataType: null,
      rawValue: null, cachedValue: undefined, displayValue: input, image: null }
  }
  const parsed = forcedText || textFormat || input === '' ? undefined
    : parseSheetNumberEdit(input, sheetCellNumberValue(current), pattern, { date1904, region, currentYear })
  const value = forcedText ? input.slice(1) : input === '' ? null : textFormat ? input : parsed?.value ?? input
  const numberPattern = parsed?.numberFormat ?? pattern
  const detectedFormat = parsed?.numberFormat ? classifySheetNumberFormat(parsed.numberFormat) : undefined
  return { ...current, value: value === null ? '' : String(value), formula: undefined,
    image: null,
    dataType: forcedText ? 'force_string' : inferSheetCellValueType(value), cachedDataType: null,
    rawValue: value, cachedValue: undefined,
    ...(parsed?.numberFormat ? { rawNumberFormat: parsed.numberFormat, numberFormat: detectedFormat === 'custom' ? undefined : detectedFormat } : {}),
    displayValue: formatSheetValue(value, numberPattern, date1904, region) ?? (value === null ? '' : String(value)) }
}

/** Rich text remains text when an existing text cell happens to contain a number or date. */
export function editSheetCellInlineText(cell: ExcelCellView | undefined, input: string, date1904 = false, region = 'zh-CN', currentYear?: number): ExcelCellView {
  const edited = editSheetCellText(cell, input, date1904, region, currentYear)
  if (edited.formula || typeof cell?.rawValue !== 'string'
    || typeof edited.rawValue !== 'number' && typeof edited.rawValue !== 'boolean') return edited
  const pattern = sheetCellNumberPattern(cell)
  const text = !isSheetTextNumberFormat(pattern) && input.startsWith("'") ? input.slice(1) : input
  return { ...cell, value: text, rawValue: text, formula: undefined,
    image: null,
    dataType: isForcedSheetCellValueType(cell.dataType) ? 'force_string' : 'text', cachedDataType: null,
    cachedValue: undefined, displayValue: formatSheetValue(text, pattern, date1904, region) ?? text }
}
