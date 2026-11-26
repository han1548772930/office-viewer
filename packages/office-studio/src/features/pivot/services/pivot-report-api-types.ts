import type { PivotChartCommandContext } from './pivot-chart-api-types.ts'

/** Report commands may originate on a dashboard separate from the output worksheet. */
export interface PivotReportCommandContext extends PivotChartCommandContext {
  hostSheetId?: string
}

export interface PivotReportCommandTarget extends PivotReportCommandContext {
  unitId: string
  pivotTableId: string
}
