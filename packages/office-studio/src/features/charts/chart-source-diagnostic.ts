import type { ChartModel, WorkbookModel } from '../../domain/model'
import { ComputeError } from '../../compute/errors'

export const CHART_SOURCE_NOT_FOUND = 'chart.source-not-found' as const
export type ChartSourceIdentity = Pick<ChartModel, 'sheetId' | 'source' | 'pivotTableId'> & Partial<Pick<ChartModel, 'id' | 'ownedPivot'>>
export interface ChartSourceNotFoundDiagnostic {
  code: typeof CHART_SOURCE_NOT_FOUND; sourceSheetId: string; hostSheetId: string; chartId?: string
}
export class ChartSourceNotFoundError extends ComputeError {
  readonly diagnostic: ChartSourceNotFoundDiagnostic
  constructor(diagnostic: ChartSourceNotFoundDiagnostic) {
    super('COMPUTE_INVALID_INPUT', '图表源工作表不存在，请恢复工作表或重新选择数据源。', JSON.stringify(diagnostic))
    this.name = 'ChartSourceNotFoundError'; this.diagnostic = diagnostic
  }
}

/** An explicit source identity is authoritative. The chart's host sheet and
 * a previously prepared snapshot can never replace a missing source sheet. */
export function chartSourceNotFoundDiagnostic(workbook: WorkbookModel, chart: ChartSourceIdentity): ChartSourceNotFoundDiagnostic | undefined {
  if (chart.pivotTableId) return undefined
  const sourceSheetId = chart.ownedPivot?.source.subUnitId ?? chart.source.sheetId ?? chart.sheetId
  return workbook.sheets.some(sheet => sheet.id === sourceSheetId) ? undefined
    : { code: CHART_SOURCE_NOT_FOUND, sourceSheetId, hostSheetId: chart.sheetId, ...(chart.id ? { chartId: chart.id } : {}) }
}
export function assertChartSourcePresent(workbook: WorkbookModel, chart: ChartSourceIdentity): void {
  const diagnostic = chartSourceNotFoundDiagnostic(workbook, chart)
  if (diagnostic) throw new ChartSourceNotFoundError(diagnostic)
}
/** The stable feature code survives generic compute-error serialization. */
export function readChartSourceNotFoundDiagnostic(error: unknown): ChartSourceNotFoundDiagnostic | undefined {
  if (error instanceof ChartSourceNotFoundError) return error.diagnostic
  if (!(error instanceof ComputeError) || !error.details) return undefined
  try {
    const value: unknown = JSON.parse(error.details)
    if (!value || typeof value !== 'object') return undefined
    const item = value as Partial<ChartSourceNotFoundDiagnostic>
    return item.code === CHART_SOURCE_NOT_FOUND && typeof item.sourceSheetId === 'string' && typeof item.hostSheetId === 'string'
      && (item.chartId === undefined || typeof item.chartId === 'string') ? item as ChartSourceNotFoundDiagnostic : undefined
  } catch { return undefined }
}
