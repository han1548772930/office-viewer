import type { WorkbookModel } from '../../../domain/model.ts'
import { cloneModel } from '../../../domain/serialization.ts'
import { equivalentModelData } from '../../../platform/commands/model-equivalence.ts'

/** Direct report edits, source structure edits and Slicer fan-out share one ownership projection. */
export function synchronizeLinkedPivotChartReports(before: WorkbookModel, after: WorkbookModel): void {
  const previousReports = new Map((before.pivotTables ?? []).map(report => [report.id, report]))
  const reports = new Map((after.pivotTables ?? []).map(report => [report.id, report]))
  const previousCharts = new Map((before.charts ?? []).map(chart => [chart.id, chart]))
  for (const chart of after.charts ?? []) {
    if (chart.pivotTableId && chart.ownedPivot || !chart.pivotTableId && !chart.ownedPivot) continue
    const previous = previousCharts.get(chart.id)
    const report = chart.ownedPivot ?? reports.get(chart.pivotTableId!)
    // Missing imported or disconnected resources retain their original identity.
    if (!report) continue
    const source = { ...chart.source,
      sheetId: chart.ownedPivot ? report.source.subUnitId : report.anchor.subUnitId,
      range: cloneModel(chart.ownedPivot ? report.source.range : report.outputRange
        ?? { startRow: report.anchor.row, endRow: report.anchor.row, startCol: report.anchor.col, endCol: report.anchor.col }) }
    const previousReport = chart.ownedPivot ? previous?.ownedPivot : previousReports.get(report.id)
    const changed = !equivalentModelData(previousReport, report) || !equivalentModelData(chart.source, source)
    chart.source = source
    if (changed && previous && chart.revision === previous.revision) chart.revision += 1
  }
}
