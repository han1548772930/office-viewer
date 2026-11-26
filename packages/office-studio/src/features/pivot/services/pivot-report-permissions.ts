import type { PivotTableModel, WorkbookModel } from '../../../domain/model.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../../sheet-protection/range-protection.ts'

export function assertPivotReportReadPermission(workbook: WorkbookModel, report: PivotTableModel, userId: string): void {
  const source = workbook.sheets.find(sheet => sheet.id === report.source.subUnitId)
  const output = workbook.sheets.find(sheet => sheet.id === report.anchor.subUnitId)
  const outputRange = report.outputRange ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col }
  if (report.source.unitId !== workbook.id || report.anchor.unitId !== workbook.id || !source || !output
    || !canViewWorksheetProtection(source, userId) || !canViewProtectedRanges(source, [report.source.range], userId)
    || !canViewWorksheetProtection(output, userId) || !canViewProtectedRanges(output, [outputRange], userId)) {
    throw new Error('没有查看透视报告来源或输出范围的权限。')
  }
}

export function assertPivotReportEditPermission(workbook: WorkbookModel, report: PivotTableModel, userId: string): void {
  assertPivotReportReadPermission(workbook, report, userId)
  const output = workbook.sheets.find(sheet => sheet.id === report.anchor.subUnitId)!
  const range = report.outputRange ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col }
  if (!canEditWorksheetProtection(output, userId) || !canEditProtectedRanges(output, [range], userId)) {
    throw new Error('没有编辑透视报告输出范围的权限。')
  }
}
