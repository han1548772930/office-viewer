import type { CellScalar } from '../../domain/model.ts'
import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { renderSheetNumberFormat, type SheetNumberPresentation } from './number-format-render.ts'
import { SHEET_NUMBER_FORMATS } from './sheet-number-format.ts'

/** Legacy hosts can supply strings; typed raw values remain authoritative. */
export function sheetCellNumberValue(cell: ExcelCellView | undefined): CellScalar | undefined {
  if (!cell) return undefined
  if (cell.rawValue !== undefined) return cell.formula && cell.rawValue === null ? undefined : cell.rawValue
  const value = cell.formula ? cell.cachedValue : cell.value
  if (value === undefined) return undefined
  if (value !== '' && /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(value)) {
    const number = Number(value)
    if (Number.isFinite(number)) return number
  }
  return value
}

export function sheetCellNumberPattern(cell: ExcelCellView | undefined): string {
  return cell?.rawNumberFormat ?? SHEET_NUMBER_FORMATS[cell?.numberFormat ?? 'general']
}

/** A format preview only changes numeric presentation, never the cell or its formula cache. */
export function presentSheetCellNumber(cell: ExcelCellView | undefined, date1904 = false, region?: string, previewPattern?: string): SheetNumberPresentation {
  if (!cell) return { text: '' }
  const value = sheetCellNumberValue(cell)
  if (value !== undefined) {
    const pattern = previewPattern !== undefined && typeof value === 'number' ? previewPattern : sheetCellNumberPattern(cell)
    const presentation = renderSheetNumberFormat(value, pattern, date1904, region)
    if (presentation) return presentation
  }
  return { text: cell.displayValue ?? (cell.formula ? cell.cachedValue ?? cell.formula : cell.value) }
}
