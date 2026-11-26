import type { PivotChartCommandContext } from './pivot-chart-api-types.ts'
import type { PivotReportCommandTarget } from './pivot-report-api-types.ts'

export interface PivotReportHierarchyTarget extends PivotReportCommandTarget {}

export interface PivotChartHierarchyTarget extends PivotChartCommandContext {
  unitId: string
  pivotChartId: string
  /** Accept the fixed public subUnitId or the existing Office Studio sheetId. */
  subUnitId?: string
  sheetId?: string
}

export const PIVOT_HIERARCHY_COMMANDS = {
  setCollapse: 'sheet.command.set-pivot-collapse',
  setCategoryLevel: 'sheet.command.set-pivot-category-level',
  stepCategoryLevel: 'sheet.command.step-pivot-category-level',
  chartSetCollapse: 'sheet.command.set-pivot-chart-collapse',
  chartSetCategoryLevel: 'sheet.command.set-pivot-chart-category-level',
  chartStepCategoryLevel: 'sheet.command.step-pivot-chart-category-level',
  chartExpandItem: 'sheet.command.expand-pivot-chart-category-item',
  chartCollapseItem: 'sheet.command.collapse-pivot-chart-category-item',
} as const
