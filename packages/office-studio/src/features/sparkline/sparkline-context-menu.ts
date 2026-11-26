import type { CellRange } from '../../domain/model.ts'
import type { ExcelSheetView } from '../../services/excelViewAdapter.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { parseSparklineReference } from './sparkline-reference.ts'

export interface SparklineContextCapability { visible: boolean; canEdit: boolean; canDelete: boolean }
/** The fixed desktop submenu requires a single selection with a sparkline primary. */
export function sparklineContextCapability(sheet: ExcelSheetView, ranges: readonly CellRange[], primary: string, userId: string, sheets: readonly ExcelSheetView[] = [sheet]): SparklineContextCapability {
  const selected = sheet.cells[primary]?.sparkline
  if (!selected || ranges.length !== 1) return { visible: false, canEdit: false, canDelete: false }
  const members = Object.entries(sheet.cells).flatMap(([address, cell]) => {
    const reference = cell.sparkline && parseSparklineReference(address)
    return reference ? [{ spec: cell.sparkline!, address, range: reference.range }] : []
  })
  const picked = members.filter(member => ranges.some(range => member.range.startRow >= range.startRow && member.range.endRow <= range.endRow
    && member.range.startCol >= range.startCol && member.range.endCol <= range.endCol))
  if (!picked.some(member => member.address === primary)) return { visible: false, canEdit: false, canDelete: false }
  const groupKeys = new Set(picked.map(member => member.spec.groupId ?? `ungrouped:${member.address}`))
  const groups = members.filter(member => groupKeys.has(member.spec.groupId ?? `ungrouped:${member.address}`)), targets = groups.map(member => member.range)
  const canDelete = canViewWorksheetProtection(sheet, userId) && canViewProtectedRanges(sheet, targets, userId)
    && canEditWorksheetProtection(sheet, userId) && canEditProtectedRanges(sheet, targets, userId)
  const sourcesReadable = groups.every(member => {
    const reference = parseSparklineReference(member.spec.sourceRange)
    const source = reference && (reference.sheetName === undefined ? sheet : sheets.find(sheet => sheet.id === reference.sheetName || sheet.name.toLocaleLowerCase() === reference.sheetName!.toLocaleLowerCase()))
    return !!source && canViewWorksheetProtection(source, userId) && canViewProtectedRanges(source, [reference!.range], userId)
  })
  return { visible: canDelete, canDelete, canEdit: canDelete && sourcesReadable }
}
