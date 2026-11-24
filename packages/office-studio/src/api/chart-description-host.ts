import type { ChartModel, WorkbookModel } from '../domain/model'
import { chartDescriptionFieldIndexes, type ChartDescriptionContext } from '../features/charts/chart-description'

/** Fetch only requested header cells. Data/role inference belongs to the source
 * engine; a description must not trigger a statistics task or scan all records. */
export function chartDescriptionSourceContext(workbook: WorkbookModel, chart: ChartModel): ChartDescriptionContext {
  const context: ChartDescriptionContext = { id: chart.id }
  // Synthetic pivot headers depend on the retained projection. Until that
  // service supplies metadata, use authored field overrides without recomputing.
  if (chart.pivotTableId) return context
  const sheet = workbook.sheets.find(candidate => candidate.id === (chart.source.sheetId ?? chart.sheetId))
  if (!sheet) throw new Error('图表源工作表不存在。')
  const { range, orientation } = chart.source
  const fieldCount = orientation === 'columns' ? range.endCol - range.startCol + 1 : range.endRow - range.startRow + 1
  const fields: Record<number, string> = {}
  for (const field of chartDescriptionFieldIndexes(chart)) {
    if (field >= fieldCount) continue
    const row = range.startRow + (orientation === 'rows' ? field : 0)
    const column = range.startCol + (orientation === 'columns' ? field : 0)
    const value = sheet.cells[`${row}:${column}`]?.value
    fields[field] = value == null ? '' : String(value)
  }
  context.fields = fields
  return context
}
