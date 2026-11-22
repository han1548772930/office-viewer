import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetCellNumberValue, sheetCellStoredValueType } from './sheet-number-presentation.ts'

export type SheetCellTextAlignment = 'left' | 'center' | 'right'

/** General alignment follows the stored scalar type and preserves explicit formatting. */
export function sheetCellTextAlignment(cell: ExcelCellView | undefined): SheetCellTextAlignment {
  if (cell?.align && cell.align !== 'general') return cell.align === 'justify' ? 'left' : cell.align
  const dataType = sheetCellStoredValueType(cell)
  if (dataType === 'number') return 'right'
  if (dataType === 'boolean' || dataType === 'error') return 'center'
  if (dataType === 'text' || dataType === 'blank') return 'left'
  const value = sheetCellNumberValue(cell)
  return typeof value === 'number' ? 'right' : typeof value === 'boolean' ? 'center' : 'left'
}
