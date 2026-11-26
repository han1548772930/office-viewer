import type { OfficeStudio } from './office-studio.ts'
import { cloneModel } from '../domain/serialization.ts'
import type { PivotTableModel, WorkbookModel } from '../domain/model.ts'
import { setPivotCategoryLevel, setPivotCollapse, stepPivotCategoryLevel } from '../features/pivot/services/pivot-engine.ts'
import type { PivotChartReportEdit } from '../features/pivot/services/pivot-chart-engine.ts'
import { PIVOT_HIERARCHY_COMMANDS, type PivotChartHierarchyTarget, type PivotReportHierarchyTarget } from '../features/pivot/services/pivot-hierarchy-api-types.ts'
import { commitPivotChartReportEdit, resolvePivotChartTarget } from './pivot-chart-commands.ts'
import { getPivotChartAxisSelection } from './pivot-chart-axis-selection.ts'

function requireReport(workbook: WorkbookModel, id: string): PivotTableModel {
  const report = workbook.pivotTables?.find(report => report.id === id)
  if (!report) throw new Error('透视报告已被删除。')
  return report
}

function chartEdit(studio: OfficeStudio, raw: PivotChartHierarchyTarget, edit: PivotChartReportEdit, commandId: string): Promise<boolean> {
  return commitPivotChartReportEdit(studio, resolvePivotChartTarget(raw), edit, commandId)
}

export function registerPivotHierarchyCommands(studio: OfficeStudio): void {
  studio.registerCommand<PivotReportHierarchyTarget & { fieldId: string; collapse: boolean; item?: string }, boolean>(PIVOT_HIERARCHY_COMMANDS.setCollapse, raw => {
    const input = cloneModel(raw)
    return studio.commitPivotChange(input.unitId, input.pivotTableId, undefined, PIVOT_HIERARCHY_COMMANDS.setCollapse, '设置透视层级展开折叠',
      workbook => setPivotCollapse(workbook, requireReport(workbook, input.pivotTableId), input.fieldId, input.collapse, input.item), input, input.hostSheetId)
  })
  studio.registerCommand<PivotReportHierarchyTarget & { fieldId: string }, boolean>(PIVOT_HIERARCHY_COMMANDS.setCategoryLevel, raw => {
    const input = cloneModel(raw)
    return studio.commitPivotChange(input.unitId, input.pivotTableId, undefined, PIVOT_HIERARCHY_COMMANDS.setCategoryLevel, '设置透视类别层级',
      workbook => setPivotCategoryLevel(workbook, requireReport(workbook, input.pivotTableId), input.fieldId), input, input.hostSheetId)
  })
  studio.registerCommand<PivotReportHierarchyTarget & { direction: 'expand' | 'collapse' }, boolean>(PIVOT_HIERARCHY_COMMANDS.stepCategoryLevel, raw => {
    const input = cloneModel(raw)
    return studio.commitPivotChange(input.unitId, input.pivotTableId, undefined, PIVOT_HIERARCHY_COMMANDS.stepCategoryLevel, '切换透视类别层级',
      workbook => { stepPivotCategoryLevel(workbook, requireReport(workbook, input.pivotTableId), input.direction) }, input, input.hostSheetId)
  })
  studio.registerCommand<PivotChartHierarchyTarget & { fieldId: string; collapse: boolean; item?: string }, Promise<boolean>>(PIVOT_HIERARCHY_COMMANDS.chartSetCollapse, input =>
    chartEdit(studio, input, { type: 'setCollapse', fieldId: input.fieldId, collapse: input.collapse, ...(input.item !== undefined ? { item: input.item } : {}) }, PIVOT_HIERARCHY_COMMANDS.chartSetCollapse), { asynchronous: true })
  studio.registerCommand<PivotChartHierarchyTarget & { fieldId: string }, Promise<boolean>>(PIVOT_HIERARCHY_COMMANDS.chartSetCategoryLevel, input =>
    chartEdit(studio, input, { type: 'setCategoryLevel', fieldId: input.fieldId }, PIVOT_HIERARCHY_COMMANDS.chartSetCategoryLevel), { asynchronous: true })
  studio.registerCommand<PivotChartHierarchyTarget & { direction: 'expand' | 'collapse' }, Promise<boolean>>(PIVOT_HIERARCHY_COMMANDS.chartStepCategoryLevel, input =>
    chartEdit(studio, input, { type: 'stepCategoryLevel', direction: input.direction }, PIVOT_HIERARCHY_COMMANDS.chartStepCategoryLevel), { asynchronous: true })
  for (const [commandId, collapse] of [[PIVOT_HIERARCHY_COMMANDS.chartExpandItem, false], [PIVOT_HIERARCHY_COMMANDS.chartCollapseItem, true]] as const) {
    studio.registerCommand<(PivotChartHierarchyTarget & { fieldId: string; item: string }) | undefined, Promise<boolean>>(commandId, async raw => {
      const input = raw === undefined ? getPivotChartAxisSelection(studio) : cloneModel(raw)
      if (!input) return false
      return chartEdit(studio, input, { type: 'setCollapse', fieldId: input.fieldId, item: input.item, collapse }, commandId)
    }, { asynchronous: true })
  }
}
