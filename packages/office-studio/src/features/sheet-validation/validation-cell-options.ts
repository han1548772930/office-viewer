import type { SheetDataValidationRule, WorkbookModel } from '../../domain/model.ts'
import { cellKey } from '../../services/excelViewAdapter.ts'
import type { SheetValidationPosition } from './validation-engine.ts'
import { sheetValidationListOptions } from './validation-list-input.ts'
import { sheetValidationListRuleAt, type SheetValidationTagOption, type SheetValidationTagOptionsAt } from './validation-tag-layout.ts'
import { sheetValidationSourceReader } from './validation-value-read.ts'
import type { SheetValidationListOptionsRead } from './validation-list-cache.ts'

export type SheetValidationTagOptionsReader = (positions: readonly SheetValidationPosition[]) => ReadonlyMap<string, readonly SheetValidationTagOption[]>

/** A synchronous, snapshot-owned resolver shares each rule's formatted candidates within one layout/dimension operation. */
export function createSheetValidationTagOptionsResolver(workbook: WorkbookModel, sheetId: string, userId?: string,
  region = 'zh-CN', readOptions?: SheetValidationListOptionsRead): SheetValidationTagOptionsAt {
  const sheet = workbook.sheets.find(candidate => candidate.id === sheetId), canRead = sheetValidationSourceReader(userId)
  const cached = new Map<SheetDataValidationRule, readonly SheetValidationTagOption[]>()
  return (row, column) => {
    if (!sheet || !Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || column < 0
      || row >= (sheet.rowCount ?? 1_048_576) || column >= (sheet.columnCount ?? 16_384)) return undefined
    const merge = sheet.mergedRanges.find(range => row >= range.startRow && row <= range.endRow && column >= range.startCol && column <= range.endCol)
    const mainRow = merge?.startRow ?? row, mainColumn = merge?.startCol ?? column
    if (canRead && !canRead(sheet, merge ?? { startRow: row, endRow: row, startCol: column, endCol: column })) return undefined
    const rule = sheetValidationListRuleAt(sheet.dataValidations, mainRow, mainColumn)
    if (!rule) return undefined
    let options = cached.get(rule)
    if (!options) {
      options = (readOptions ? readOptions(rule) : sheetValidationListOptions(rule, workbook, sheet.id, userId, region))
        .map(({ label, color }) => ({ label, ...(color ? { color } : {}) }))
      cached.set(rule, options)
    }
    return options
  }
}

/** Only requested, viewable main cells reach the Canvas. The result contains labels/colors, never source DTOs or rules. */
export function sheetValidationCellTagOptions(workbook: WorkbookModel, sheetId: string, positions: readonly SheetValidationPosition[],
  userId?: string, region = 'zh-CN', readOptions?: SheetValidationListOptionsRead): Map<string, readonly SheetValidationTagOption[]> {
  const sheet = workbook.sheets.find(candidate => candidate.id === sheetId), result = new Map<string, readonly SheetValidationTagOption[]>()
  if (!sheet) return result
  const read = createSheetValidationTagOptionsResolver(workbook, sheetId, userId, region, readOptions)
  for (const position of positions) {
    const options = read(position.row, position.column)
    if (!options) continue
    const merge = sheet.mergedRanges.find(range => position.row >= range.startRow && position.row <= range.endRow
      && position.column >= range.startCol && position.column <= range.endCol)
    result.set(cellKey(merge?.startRow ?? position.row, merge?.startCol ?? position.column), options)
  }
  return result
}
