import type { CellRange, SlicerModel, WorkbookModel } from '../../domain/model.ts'
import { sheetDrawingRefAffectedRanges } from '../sheet-drawings/sheet-drawing-host.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewProtectedRanges, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { getSlicerCache, getSlicerConnectionReports, getSlicerPivotSource, SlicerConnectionError, type SlicerConnectionReport } from './slicer-connections.ts'

function assertRanges(workbook: WorkbookModel, sheetId: string, ranges: CellRange[], userId: string, edit: boolean): void {
  const sheet = workbook.sheets.find(candidate => candidate.id === sheetId)
  if (!sheet || !canViewWorksheetProtection(sheet, userId) || !canViewProtectedRanges(sheet, ranges, userId)
    || edit && (!canEditWorksheetProtection(sheet, userId) || !canEditProtectedRanges(sheet, ranges, userId))) {
    throw new SlicerConnectionError('permission-denied', edit ? '没有编辑切片器、连接报告或其完整范围的权限。' : '没有查看切片器或其来源范围的权限。')
  }
}

export function assertSlicerViewPermission(workbook: WorkbookModel, view: SlicerModel, userId: string, edit: boolean): void {
  assertRanges(workbook, view.hostSheetId, sheetDrawingRefAffectedRanges(workbook, view.hostSheetId, [{ kind: view.type, id: view.id }]), userId, edit)
}

export function assertSlicerSourceReadPermission(workbook: WorkbookModel, view: SlicerModel, userId: string): void {
  const cache = getSlicerCache(workbook, view)
  assertSlicerViewPermission(workbook, view, userId, false)
  if (cache.source.kind === 'pivot') {
    const source = getSlicerPivotSource(workbook, view)
    assertRanges(workbook, source.subUnitId, [source.range], userId, false)
  } else {
    const source = cache.source, table = workbook.tables?.find(candidate => candidate.id === source.tableId)
    if (!table) throw new SlicerConnectionError('source-missing', '切片器表格数据源不存在。')
    assertRanges(workbook, source.subUnitId, [table.range], userId, false)
  }
}

export function assertSlicerFilterPermission(workbook: WorkbookModel, view: SlicerModel, userId: string, reports?: readonly SlicerConnectionReport[]): void {
  assertSlicerSourceReadPermission(workbook, view, userId)
  const cache = getSlicerCache(workbook, view)
  for (const control of workbook.slicers ?? []) {
    let shared: boolean
    try { shared = getSlicerCache(workbook, control).id === cache.id }
    catch (error) { if (error instanceof SlicerConnectionError) continue; throw error }
    if (shared) assertSlicerViewPermission(workbook, control, userId, true)
  }
  if (cache.source.kind === 'table') {
    // The field filter belongs to the source worksheet, independent of the drawing host.
    const source = cache.source, sourceTable = workbook.tables?.find(candidate => candidate.id === source.tableId)
    if (!sourceTable) throw new SlicerConnectionError('source-missing', '切片器表格来源不存在。')
    assertRanges(workbook, source.subUnitId, [sourceTable.range], userId, true)
    return
  }
  for (const report of reports ?? getSlicerConnectionReports(workbook, view)) {
    if (report.kind === 'ownedPivotChart') {
      assertRanges(workbook, report.subUnitId, sheetDrawingRefAffectedRanges(workbook, report.subUnitId, [{ kind: 'chart', id: report.id }]), userId, true)
    } else {
      const range = report.pivot.outputRange ?? { startRow: report.pivot.anchor.row, endRow: report.pivot.anchor.row,
        startCol: report.pivot.anchor.col, endCol: report.pivot.anchor.col }
      assertRanges(workbook, report.subUnitId, [range], userId, true)
    }
  }
}
