import type { ChartConfig, ChartUpdatePatch } from '../../charts/chart-engine.ts'
import type { PivotTableAnchorInput, PivotTableSourceInput } from './pivot-engine.ts'
import type { PivotChartFieldsConfig } from './pivot-chart-engine.ts'

export interface PivotChartTableReference {
  unitId: string
  subUnitId: string
  pivotTableId: string
}

/** Ownership is explicit; drawing source ranges are derived from this report. */
export type PivotChartSourceSpec =
  | { kind: 'range'; sourceRangeInfo: PivotTableSourceInput }
  | { kind: 'pivotTable'; pivotTableRef: PivotChartTableReference }

export type PivotChartInfo = Omit<ChartConfig, 'source' | 'pivotTableId' | 'ownedPivot'> & {
  source: PivotChartSourceSpec
  initialFieldsConfig?: PivotChartFieldsConfig
}

/** Keep existing linked ChartConfig callers on the same content command. */
export type LinkedPivotChartInfo = ChartConfig & { pivotTableId: string }
export type PivotChartInsertInfo = PivotChartInfo | LinkedPivotChartInfo

/** Local counterpart of the menu's withPivotTable destination, outside chart authoring data. */
export interface PivotChartInsertOptions {
  pivotTable?: {
    anchor: PivotTableAnchorInput
    pivotTableId?: string
  }
}

export interface PivotChartCommandContext {
  expectedRevision?: number
  expectedHistoryGeneration?: number
  expectedUserId?: string
  host?: 'workbench'
}

export interface PivotChartConfigureInput {
  patch: ChartUpdatePatch
  fieldsConfig?: PivotChartFieldsConfig
}

export const PIVOT_CHART_COMMANDS = {
  insert: 'sheet.command.add-pivot-chart',
  edit: 'sheet.command.edit-pivot-chart-report',
  configure: 'sheet.command.configure-pivot-chart',
  setControls: 'sheet.command.update-pivot-chart-controls',
  refresh: 'sheet.command.refresh-pivot-chart',
  remove: 'sheet.command.remove-pivot-chart',
} as const
