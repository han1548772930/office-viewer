import type { CellScalar } from '../../domain/model'
import type { ChartConfig, ChartData } from './chart-engine'
import { chartNumber, type ChartRecords } from './chart-data'
import { prepareChartDateSource, supportsChartDateAxis, type ChartDateAxisConfig } from './chart-date-axis'
import { projectChartRecords } from './chart-record-projection'
import { chartColumnName, parseChartSourceRange, type ChartSettingsDraft, type ChartSettingsSourceSheet } from './chart-settings-draft'
import { hiddenSheetColumns, hiddenSheetRows, isSheetColumnVisible, isSheetRowVisible } from '../sheet-workbench/sheet-row-visibility'

export const CHART_SETTINGS_PREVIEW_MAX_CELLS = 65_536
export interface ChartSettingsProjection { source?: ChartRecords; data?: ChartData; error?: string }

/** The settings preview shares the renderer's binding projection, with a bounded UI input. */
export function projectChartSettings(draft: ChartSettingsDraft, sheet?: ChartSettingsSourceSheet, provided?: ChartData, date1904 = provided?.date1904): ChartSettingsProjection {
  let source: ChartRecords
  try {
    if (draft.config.pivotTableId || !sheet) {
      if (!provided?.fields || !provided.records) return provided ? { data: date1904 === undefined ? provided : { ...provided, date1904 } } : {}
      previewBudget(provided.fields.length, provided.records.length + 1)
      source = { fields: provided.fields, records: provided.records,
        sourceCategoryFields: provided.sourceCategoryFields, sourceValueFields: provided.sourceValueFields,
        visibleFields: provided.visibleFields, syntheticValues: provided.syntheticValues }
    } else source = chartSettingsRecords(draft, sheet)
    if (date1904 !== undefined) source = { ...source, date1904 }
  } catch (reason) { return { error: reason instanceof Error ? reason.message : String(reason) } }
  try {
    return { source, data: projectChartRecords(source, { type: draft.config.type, categoryField: draft.config.categoryField ?? 0,
      valueFields: draft.config.valueFields ?? [], seriesToCategory: draft.config.seriesToCategory, options: draft.config.options }) }
  } catch (reason) { return { source, error: reason instanceof Error ? reason.message : String(reason) } }
}

/** Report real date/bounds errors in the bounded settings projection. Histogram dates come from final bins. */
export function chartSettingsDateAxisError(config: ChartConfig, data?: ChartData, date1904 = data?.date1904): string | undefined {
  if (!config.options?.mapping?.useDateAxis || !supportsChartDateAxis(config)) return undefined
  const chart: ChartDateAxisConfig = { type: config.type, options: config.options, xAxis: { scaleType: 'linear', ...config.xAxis } }
  const current: ChartData = { ...(data ?? { categories: [], series: [] }), date1904 }
  try { prepareChartDateSource(current, chart); return undefined }
  catch (reason) { return reason instanceof Error ? reason.message : String(reason) }
}

export function chartSettingsRecords(draft: ChartSettingsDraft, sheet: ChartSettingsSourceSheet): ChartRecords {
  const range = parseChartSourceRange(draft.sourceText)
  if (!range) throw new RangeError('请输入有效的数据范围，例如 A1:C4。')
  if (sheet.rowCount !== undefined && range.endRow >= sheet.rowCount || sheet.columnCount !== undefined && range.endCol >= sheet.columnCount) {
    throw new RangeError('图表数据区域超出了所选工作表的边界。')
  }
  const columns = draft.config.source.orientation === 'columns'
  const fieldCount = columns ? range.endCol - range.startCol + 1 : range.endRow - range.startRow + 1
  const recordCount = columns ? range.endRow - range.startRow : range.endCol - range.startCol
  if (!recordCount || !fieldCount) throw new RangeError('图表数据区域须包含字段名和至少一条数据。')
  previewBudget(fieldCount, recordCount + 1)
  const view = { ...sheet, rowCount: sheet.rowCount ?? range.endRow + 1, columnCount: sheet.columnCount ?? range.endCol + 1 }
  const rows = hiddenSheetRows(view), hiddenColumns = hiddenSheetColumns(view)
  const collapsed = sheet.outline?.groups.filter(group => group.collapsed) ?? []
  const rowVisible = (row: number): boolean => isSheetRowVisible(view, row, rows) && !collapsed.some(group => group.axis === 'row' && row >= group.start && row <= group.end)
  const columnVisible = (column: number): boolean => isSheetColumnVisible(view, column, hiddenColumns) && !collapsed.some(group => group.axis === 'column' && column >= group.start && column <= group.end)
  const read = (field: number, record: number): CellScalar => {
    const row = range.startRow + (columns ? record : field), column = range.startCol + (columns ? field : record)
    const cell = sheet.cells[`${chartColumnName(column)}${row + 1}`]
    if (!cell) return null
    if (cell.rawValue !== undefined) return cell.rawValue
    if (cell.formula && cell.cachedValue !== undefined && cell.cachedValue !== '') {
      const cached = chartNumber(cell.cachedValue)
      if (cached !== null) return cached
    }
    return cell.value ?? null
  }
  const visible = draft.config.options?.plotVisibleOnly !== false
  const fields = Array.from({ length: fieldCount }, (_unused, field) => String(read(field, 0) ?? ''))
  const visibleFields = visible ? fields.map((_name, field) => field).filter(field => columns
    ? columnVisible(range.startCol + field) : rowVisible(range.startRow + field)) : undefined
  const records: CellScalar[][] = []
  for (let record = 1; record <= recordCount; record++) {
    if (visible && !(columns ? rowVisible(range.startRow + record) : columnVisible(range.startCol + record))) continue
    records.push(Array.from({ length: fieldCount }, (_unused, field) => read(field, record)))
  }
  return { fields, records, ...(visibleFields ? { visibleFields } : {}) }
}

function previewBudget(fields: number, rows: number): void {
  if (fields * rows > CHART_SETTINGS_PREVIEW_MAX_CELLS) {
    throw new RangeError(`数据源包含 ${fields * rows} 个数据位置，超过设置面板 ${CHART_SETTINGS_PREVIEW_MAX_CELLS} 个位置的预览预算；字段配置仍可编辑，完整图表计算须通过计算服务处理。`)
  }
}
