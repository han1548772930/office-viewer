import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { classifySheetNumberFormat } from './sheet-number-format.ts'

const numericText = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i

export function convertTextToNumber(cell: ExcelCellView): ExcelCellView | undefined {
  if (cell.formula || typeof cell.rawValue === 'boolean' || cell.rawValue === null) return undefined
  const isTextFormat = cell.rawNumberFormat ? classifySheetNumberFormat(cell.rawNumberFormat) === 'text' : cell.numberFormat === 'text'
  if (typeof cell.rawValue === 'number' && !isTextFormat) return undefined
  const text = cell.rawValue ?? cell.value
  if (typeof text === 'string' && (!text.trim() || !numericText.test(text.trim()))) return undefined
  const value = Number(typeof text === 'string' ? text.trim() : text)
  if (!Number.isFinite(value)) return undefined
  return { ...cell, value: String(value), rawValue: value, displayValue: undefined, ...(isTextFormat ? { numberFormat: 'general', rawNumberFormat: 'General' } : {}) }
}
