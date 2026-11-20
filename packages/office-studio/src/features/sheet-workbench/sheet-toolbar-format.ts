import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { defaultCellStyle } from '../../domain/model.ts'
import { cellAddress } from '../../rendering/spreadsheet/layout.ts'
import { changeSheetDecimalPlaces, classifySheetNumberFormat, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from './sheet-number-format.ts'

export type SheetBorderPreset = 'all' | 'outside' | 'inside' | 'top' | 'bottom' | 'left' | 'right' | 'none'
const FONT_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
export function nextSheetFontSize(size: number, direction: -1 | 1): number {
  return direction === 1 ? FONT_STEPS.find(step => step > size) ?? Math.min(400, size + 1)
    : [...FONT_STEPS].reverse().find(step => step < size) ?? Math.max(1, size - 1)
}
export function applySheetNumberFormat(cell: ExcelCellView, format: SheetNumberFormat): ExcelCellView {
  return { ...cell, numberFormat: format, rawNumberFormat: SHEET_NUMBER_FORMATS[format] }
}
export function adjustSheetDecimals(cell: ExcelCellView, delta: -1 | 1): ExcelCellView {
  const pattern = changeSheetDecimalPlaces(cell.rawNumberFormat ?? SHEET_NUMBER_FORMATS[cell.numberFormat ?? 'general'], delta)
  const kind = classifySheetNumberFormat(pattern)
  return { ...cell, rawNumberFormat: pattern, numberFormat: kind === 'custom' ? cell.numberFormat : kind }
}
export function clearSheetCellFormat(cell: ExcelCellView): ExcelCellView {
  const style = defaultCellStyle()
  return { ...cell, fontFamily: style.font.family, fontSize: style.font.size, bold: false, italic: false,
    underline: false, strike: false, color: style.font.color, background: style.fill, align: 'left', vertical: style.vertical,
    wrap: false, border: {}, rotation: 0, indent: 0, shrinkToFit: false, numberFormat: 'general', rawNumberFormat: 'General' }
}

/** Scope perimeter borders to each selected rectangle; do not fake outside borders on every cell. */
export function applySheetBorders(cells: Record<string, ExcelCellView>, ranges: readonly GridRange[], preset: SheetBorderPreset): void {
  for (const range of ranges) for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    const key = cellAddress(row, column), cell = cells[key] ?? { value: '' }
    const border: Record<string, unknown> = preset === 'none' ? {} : { ...cell.border }
    const edges = { top: row === range.startRow, bottom: row === range.endRow, left: column === range.startColumn, right: column === range.endColumn }
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const selected = preset === 'all' || preset === 'outside' && edges[side] || preset === 'inside' && !edges[side] || preset === side && edges[side]
      if (selected) border[side] = { style: 'thin', color: '#1f2937' }
    }
    cells[key] = { ...cell, border }
  }
}
