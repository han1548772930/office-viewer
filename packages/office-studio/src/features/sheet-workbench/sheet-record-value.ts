import type { CellScalar, SpreadsheetCell } from '../../domain/model.ts'
import { normalizeSheetCellValueType } from '../../domain/sheet-cell-value.ts'
import { recordValueFromScalar, type RecordValue } from '../../domain/record-value.ts'
import { classifySheetNumberFormat } from './sheet-number-format.ts'
import { parseTimelineDate } from '../slicer/timeline-date.ts'

/** Reads stored/cached values without interpreting a localized rendered display. */
export function sheetCellRecordValue(cell: SpreadsheetCell | undefined, date1904 = false): RecordValue {
  if (!cell) return { type: 'blank' }
  const value = cell.value, type = normalizeSheetCellValueType(cell.formula ? cell.cachedDataType ?? cell.dataType : cell.dataType)
  if (typeof cell.metadata?.error === 'string') return { type: 'error', value: cell.metadata.error }
  if (type === 'error') return { type: 'error', value: typeof value === 'string' ? value : cell.display }
  if (value === null) return { type: 'blank' }
  const storageType = (cell.formula ? cell.cachedDataType ?? cell.dataType : cell.dataType)?.toLowerCase()
  const format = classifySheetNumberFormat(cell.style.numberFormat)
  const dateHint = storageType === 'date' || storageType === 'd' || storageType === 'datetime'
    || typeof value === 'number' && ['date', 'datetime', 'time'].includes(format)
  if (dateHint) {
    const serial = parseTimelineDate(value, date1904)?.serial
    if (serial !== undefined) return { type: 'date', value: serial }
  }
  return recordValueFromScalar(value)
}

export function sheetRecordRawValue(cell: SpreadsheetCell | undefined): CellScalar { return cell?.value ?? null }
