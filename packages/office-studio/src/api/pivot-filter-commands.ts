import type { OfficeStudio } from './office-studio.ts'
import type { PivotTableModel, WorkbookModel } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { clearPivotFieldFilters, setPivotDateFilter, setPivotLabelConditionFilter, setPivotLabelFilter } from '../features/pivot/services/pivot-engine.ts'
import { normalizePivotDateFilterInput, normalizePivotLabelConditionInput,
  type PivotDateFilterInput, type PivotLabelConditionInput } from '../features/pivot/services/pivot-filter-api.ts'
import type { PivotReportCommandTarget } from '../features/pivot/services/pivot-report-api-types.ts'

export const PIVOT_REPORT_FILTER_COMMANDS = {
  setFilter: 'sheet.command.set-pivot-filter',
  setDateFilter: 'sheet.command.set-pivot-date-filter',
  setLabelConditionFilter: 'sheet.command.set-pivot-label-condition-filter',
} as const

interface PivotFilterTarget extends PivotReportCommandTarget {
  subUnitId?: string
  tableFieldId: string
}
export interface PivotReportManualFilterInput extends PivotFilterTarget {
  items: string[] | null
  partialItems?: string[]
  isAll?: boolean
  clearAllFilters?: boolean
}
export interface PivotReportDateFilterInput extends PivotFilterTarget { filterInfo?: PivotDateFilterInput | null }
export interface PivotReportLabelConditionInput extends PivotFilterTarget { filterInfo?: PivotLabelConditionInput | null }

function requireReport(workbook: WorkbookModel, id: string): PivotTableModel {
  const report = workbook.pivotTables?.find(report => report.id === id)
  if (!report) throw new Error('透视报告已被删除。')
  return report
}

function resolveTarget<T extends PivotFilterTarget>(studio: OfficeStudio, raw: T): T {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw) || typeof raw.unitId !== 'string' || !raw.unitId
    || typeof raw.pivotTableId !== 'string' || !raw.pivotTableId || typeof raw.tableFieldId !== 'string' || !raw.tableFieldId) {
    throw new TypeError('透视筛选需要工作簿、报告和本地字段身份。')
  }
  const input = cloneModel(raw), workbook = studio.getWorkbookModel(input.unitId)
  if (!workbook) throw new Error('透视报告的工作簿已关闭。')
  const report = requireReport(workbook, input.pivotTableId)
  if (input.subUnitId !== undefined && input.subUnitId !== report.anchor.subUnitId) throw new Error('透视报告的宿主工作表身份不一致。')
  return input
}

export function registerPivotFilterCommands(studio: OfficeStudio): void {
  studio.registerCommand<PivotReportManualFilterInput, boolean>(PIVOT_REPORT_FILTER_COMMANDS.setFilter, raw => {
    const input = resolveTarget(studio, raw)
    if (input.isAll !== undefined && typeof input.isAll !== 'boolean'
      || input.clearAllFilters !== undefined && typeof input.clearAllFilters !== 'boolean') throw new TypeError('透视筛选标志无效。')
    return studio.commitPivotChange(input.unitId, input.pivotTableId, undefined, PIVOT_REPORT_FILTER_COMMANDS.setFilter, '设置透视成员筛选', workbook => {
      const report = requireReport(workbook, input.pivotTableId)
      if (input.clearAllFilters) clearPivotFieldFilters(workbook, report, input.tableFieldId)
      else setPivotLabelFilter(workbook, report, input.tableFieldId, input.isAll ? null : input.items, 'worksheet',
        { keyMode: 'grouped', partialItems: input.isAll ? undefined : input.partialItems })
    }, input, input.hostSheetId ?? input.subUnitId)
  })
  studio.registerCommand<PivotReportDateFilterInput, boolean>(PIVOT_REPORT_FILTER_COMMANDS.setDateFilter, raw => {
    const input = resolveTarget(studio, raw), filter = normalizePivotDateFilterInput(input.filterInfo)
    return studio.commitPivotChange(input.unitId, input.pivotTableId, undefined, PIVOT_REPORT_FILTER_COMMANDS.setDateFilter, '设置透视日期条件', workbook => {
      setPivotDateFilter(workbook, requireReport(workbook, input.pivotTableId), input.tableFieldId, filter)
    }, input, input.hostSheetId ?? input.subUnitId)
  })
  studio.registerCommand<PivotReportLabelConditionInput, boolean>(PIVOT_REPORT_FILTER_COMMANDS.setLabelConditionFilter, raw => {
    const input = resolveTarget(studio, raw), filter = normalizePivotLabelConditionInput(input.filterInfo)
    return studio.commitPivotChange(input.unitId, input.pivotTableId, undefined, PIVOT_REPORT_FILTER_COMMANDS.setLabelConditionFilter, '设置透视标签条件', workbook => {
      setPivotLabelConditionFilter(workbook, requireReport(workbook, input.pivotTableId), input.tableFieldId, filter)
    }, input, input.hostSheetId ?? input.subUnitId)
  })
}
