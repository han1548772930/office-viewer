import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { defaultCellStyle } from '../../domain/model.ts'
import { cellAddress } from '../../rendering/spreadsheet/layout.ts'
import { changeSheetDecimalPlaces, classifySheetNumberFormat, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from './sheet-number-format.ts'

export const sheetBorderLabels = {
  top: '上框线', bottom: '下框线', left: '左框线', right: '右框线', none: '无框线',
  all: '所有边框', outside: '外侧框线', inside: '内部框线', horizontal: '内侧横线', vertical: '内侧竖线',
  'diagonal-down': '右下斜线', 'diagonal-down-pair': '右下双斜线', 'diagonal-down-triple': '右下三斜线',
  'diagonal-up': '右上斜线', 'diagonal-up-pair': '右上双斜线',
} as const
export type SheetBorderPreset = keyof typeof sheetBorderLabels
export type SheetBorderStyle = 'thin' | 'hair' | 'dotted' | 'dashed' | 'dash-dot' | 'dash-dot-dot' | 'medium' | 'medium-dashed' | 'medium-dash-dot' | 'medium-dash-dot-dot' | 'thick' | 'double'
export interface SheetBorderOptions { style: SheetBorderStyle; color: string }
export const sheetBorderStyleLabels: Record<SheetBorderStyle, string> = {
  thin: '细实线', hair: '极细线', dotted: '点线', dashed: '虚线', 'dash-dot': '点划线', 'dash-dot-dot': '双点划线',
  medium: '中等实线', 'medium-dashed': '中等虚线', 'medium-dash-dot': '中等点划线', 'medium-dash-dot-dot': '中等双点划线', thick: '粗实线', double: '双线',
}
export function isSheetBorderPreset(value: unknown): value is SheetBorderPreset {
  return typeof value === 'string' && Object.hasOwn(sheetBorderLabels, value)
}

const DEFAULT_BORDER_OPTIONS: SheetBorderOptions = { style: 'thin', color: '#1f2937' }
export function nextSheetFontSize(size: number, direction: -1 | 1): number {
  return Math.max(6, Math.min(400, size + direction))
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
    wrap: false, border: {}, rotation: 0, rotationVertical: false, indent: 0, shrinkToFit: false, numberFormat: 'general', rawNumberFormat: 'General' }
}

export function setSheetTextFit(cell: ExcelCellView, mode: 'wrap' | 'shrinkToFit', enabled: boolean): ExcelCellView {
  if (mode === 'wrap') return { ...cell, wrap: enabled, ...(enabled ? { shrinkToFit: false } : {}) }
  return { ...cell, shrinkToFit: enabled, ...(enabled ? { wrap: false } : {}) }
}

/** Scope perimeter borders to each selected rectangle; do not fake outside borders on every cell. */
export function applySheetBorders(
  cells: Record<string, ExcelCellView>, ranges: readonly GridRange[], preset: SheetBorderPreset,
  options: SheetBorderOptions = DEFAULT_BORDER_OPTIONS,
): void {
  for (const range of ranges) for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    const key = cellAddress(row, column), cell = cells[key] ?? { value: '' }
    const border: Record<string, unknown> = preset === 'none' ? {} : { ...cell.border }
    const edges = { top: row === range.startRow, bottom: row === range.endRow, left: column === range.startColumn, right: column === range.endColumn }
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const selected = preset === 'all' || preset === 'outside' && edges[side]
        || preset === 'inside' && !edges[side]
        || preset === side && edges[side]
        || preset === 'horizontal' && (side === 'top' || side === 'bottom') && !edges[side]
        || preset === 'vertical' && (side === 'left' || side === 'right') && !edges[side]
      if (selected) border[side] = { ...options }
    }
    if (preset === 'diagonal-down' || preset === 'diagonal-down-triple') border.diagonalDown = { ...options }
    if (preset === 'diagonal-up') border.diagonalUp = { ...options }
    if (preset === 'diagonal-down-pair' || preset === 'diagonal-down-triple') {
      border.topLeftBottomCenter = { ...options }; border.topLeftMiddleRight = { ...options }
    }
    if (preset === 'diagonal-up-pair') {
      border.middleLeftTopRight = { ...options }; border.bottomCenterTopRight = { ...options }
    }
    cells[key] = { ...cell, border }
  }
}
