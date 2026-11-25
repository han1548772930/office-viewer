import { parseCellKey, type ExcelCellView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'
import { defaultCellStyle } from '../../domain/model.ts'
import { cellAddress } from '../../rendering/spreadsheet/layout.ts'
import { changeSheetDecimalPlaces, classifySheetNumberFormat, SHEET_NUMBER_FORMATS, type SheetNumberFormat } from './sheet-number-format.ts'
import { createSheetBorderPlan, type SheetBorderOptions, type SheetBorderPreset } from './sheet-cell-border.ts'

export { sheetBorderLabels, sheetBorderStyleLabels, isSheetBorderPreset, type SheetBorderOptions, type SheetBorderPreset, type SheetBorderStyle } from './sheet-cell-border.ts'

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
  return { ...cell, authoredStyle: {}, authoredStyleReference: null, clearedStyleProperties: null,
    fontFamily: style.font.family, fontSize: style.font.size, bold: false, italic: false,
    underline: false, strike: false, color: style.font.color, background: style.fill, align: style.horizontal, vertical: style.vertical,
    wrap: false, wrapStrategy: 'overflow', border: {}, rotation: 0, rotationVertical: false, indent: 0, shrinkToFit: false, numberFormat: 'general', rawNumberFormat: 'General' }
}

export function setSheetTextFit(cell: ExcelCellView, mode: 'wrap' | 'shrinkToFit', enabled: boolean): ExcelCellView {
  if (mode === 'wrap') return { ...cell, wrap: enabled, wrapStrategy: enabled ? 'wrap' : 'overflow', ...(enabled ? { shrinkToFit: false } : {}) }
  const wasWrapped = cell.wrapStrategy === 'wrap' || cell.wrapStrategy === undefined && cell.wrap === true
  return { ...cell, shrinkToFit: enabled, ...(enabled && wasWrapped ? { wrap: false, wrapStrategy: 'overflow' as const } : {}) }
}

/** Keep the legacy four-argument view entry on the same union/neighbor plan as registered commands. */
export function applySheetBorders(
  cells: Record<string, ExcelCellView>, ranges: readonly GridRange[], preset: SheetBorderPreset,
  options: SheetBorderOptions = DEFAULT_BORDER_OPTIONS,
): void {
  const plan = createSheetBorderPlan({
    readBorder: (row, col) => cells[cellAddress(row, col)]?.border,
    existingCells: () => Object.keys(cells).flatMap(key => { const position = parseCellKey(key); return position ? [position] : [] }),
  }, ranges.map(range => ({ startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn })), preset, options)
  for (const patch of plan.patches) {
    const key = cellAddress(patch.row, patch.col)
    cells[key] = { ...(cells[key] ?? { value: '' }), border: patch.border }
  }
}
