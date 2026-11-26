import type { PivotTableModel, WorkbookModel } from '../../../domain/model.ts'
import { createDetachedPivotReport } from './pivot-engine.ts'
import { assertPivotReportEditPermission } from './pivot-report-permissions.ts'
import type { PivotChartInsertInfo, PivotChartInsertOptions } from './pivot-chart-api-types.ts'

const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)

/** Read-only preparation: the paired report and its output are registered only in the content candidate. */
export function preparePivotChartPairedTable(workbook: WorkbookModel, chartId: string, info: PivotChartInsertInfo,
  options: PivotChartInsertOptions | undefined, userId: string): PivotTableModel | undefined {
  if (options === undefined) return undefined
  if (!record(options) || Object.keys(options).some(key => key !== 'pivotTable')) throw new TypeError('透视图创建选项无效。')
  if (options.pivotTable === undefined) return undefined
  const paired = options.pivotTable
  if (!record(paired) || Object.keys(paired).some(key => !['anchor', 'pivotTableId'].includes(key)) || !record(paired.anchor)
    || Object.keys(paired.anchor).some(key => !['unitId', 'subUnitId', 'row', 'col'].includes(key))) throw new TypeError('同时创建的透视表位置无效。')
  if (!record(info.source) || !('kind' in info.source) || info.source.kind !== 'range') {
    throw new TypeError('同时创建透视表与图必须使用区域来源。')
  }
  const id = paired.pivotTableId ?? `${chartId}:pivot-table`
  if (typeof id !== 'string' || !id.trim()) throw new TypeError('同时创建的透视表 ID 无效。')
  if (workbook.pivotTables?.some(report => report.id === id) || workbook.charts?.some(chart => chart.ownedPivot?.id === id)) {
    throw new Error('同时创建的透视表 ID 已存在。')
  }
  const report = createDetachedPivotReport(workbook, info.source.sourceRangeInfo, paired.anchor, id)
  assertPivotReportEditPermission(workbook, report, userId)
  return report
}
