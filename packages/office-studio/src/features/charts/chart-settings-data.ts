import type { CellScalar } from '../../domain/model'
import type { ChartConfig, ChartData } from './chart-engine'
import type { ChartComputedStatistics } from './chart-derived'
import { chartNumber, type ChartRecords } from './chart-data'
import { prepareChartDateSource, supportsChartDateAxis, type ChartDateAxisConfig } from './chart-date-axis'
import { projectChartRecords } from './chart-record-projection'
import { chartColumnName, parseChartSourceRange, type ChartSettingsDraft, type ChartSettingsSourceSheet } from './chart-settings-draft'
import { hiddenSheetColumns, hiddenSheetRows, isSheetColumnVisible, isSheetRowVisible } from '../sheet-workbench/sheet-row-visibility'

export const CHART_SETTINGS_PREVIEW_MAX_CELLS = 65_536
export type ChartSettingsSourceRevision = number | string
/** Only the draft-source service produces prepared-source data. Statistical
 * bins, five-number rows and other derived render data cannot supply samples. */
export interface ChartSettingsProvidedData {
  kind: 'prepared-source' | 'derived'
  data: ChartData
  bindingKey: string
  sourceRevision?: ChartSettingsSourceRevision
  /** Statistics and preparedData are captured by the same draft read. */
  statistics?: ChartComputedStatistics
}
export interface ChartSettingsSourceIdentity { hostSheetId?: string; sourceRevision?: ChartSettingsSourceRevision }
export interface ChartSettingsProjection {
  source?: ChartRecords
  data?: ChartData
  statistics?: ChartComputedStatistics
  error?: string
  needsPreparation?: boolean
  prepared?: boolean
}
export type ChartSettingsSourceLoader = (config: ChartConfig, options: { signal: AbortSignal }) => Promise<ChartData>
export type ChartSettingsStatisticsLoader = (config: ChartConfig, options: { signal: AbortSignal }) => Promise<ChartComputedStatistics>

/** Configuration affecting the settings read, excluding authored paint and
 * Boxplot summaries. The workbook/session revision is separate. Pareto and
 * Histogram classifications include their statistical author options. */
export function chartSettingsBindingKey(draft: ChartSettingsDraft, hostSheetId?: string, date1904?: boolean): string {
  const range = parseChartSourceRange(draft.sourceText)
  if (!range) throw new RangeError('请输入有效的数据范围，例如 A1:C4。')
  const config = draft.config, options = config.options
  const dateAxis = options?.mapping?.useDateAxis ? {
    scaleType: config.xAxis?.scaleType ?? 'linear', min: config.xAxis?.min, max: config.xAxis?.max,
    baseTimeUnit: config.xAxis?.baseTimeUnit, majorTimeUnit: config.xAxis?.majorTimeUnit, majorUnit: config.xAxis?.majorUnit,
  } : undefined
  return JSON.stringify([config.type, hostSheetId, config.source.sheetId ?? hostSheetId, range, config.source.orientation, config.pivotTableId,
    config.categoryField ?? 0, config.valueFields ?? [], config.seriesToCategory ?? {}, date1904,
    options?.mapping, options?.plotVisibleOnly !== false, options?.invalidValueStrategy, options?.aggregation,
    options?.trendlines?.map(item => item.seriesIndex),
    options?.errorBars?.map(item => [item.seriesIndex, item.direction, item.type, item.valueType, item.plusIndex, item.minusIndex]),
    options?.allSeriesStyle?.type, Object.entries(options?.series ?? {}).filter(([_field, style]) => style.type !== undefined).map(([field, style]) => [field, style.type]),
    (options?.namedSeries ?? []).filter(item => item.style.type !== undefined).map(item => [item.name, item.style.type]), dateAxis,
    config.type === 'pareto' ? options?.pareto?.includeZeroValues === true : undefined,
    config.type === 'histogram' ? [options?.histogram?.binning, options?.histogram?.underflowThreshold, options?.histogram?.overflowThreshold] : undefined])
}

export function chartSettingsProvidedValue(provided?: ChartData | ChartSettingsProvidedData): ChartData | undefined {
  return provided && 'kind' in provided ? provided.data : provided
}

export function chartSettingsPreparedData(draft: ChartSettingsDraft, data: ChartData, identity: ChartSettingsSourceIdentity = {}, date1904 = data.date1904): ChartSettingsProvidedData {
  return { kind: 'prepared-source', data, bindingKey: chartSettingsBindingKey(draft, identity.hostSheetId, date1904), sourceRevision: identity.sourceRevision }
}

/** Preserve one read's prepared source and materialized statistical result.
 * No normalization, aggregation or Pareto sorting runs in the settings UI. */
export function chartSettingsPreparedStatistics(draft: ChartSettingsDraft, statistics: ChartComputedStatistics, identity: ChartSettingsSourceIdentity = {}, date1904 = statistics.preparedData?.date1904): ChartSettingsProvidedData {
  if (!statistics.preparedData) throw new RangeError('图表设置统计读取未返回准备后的原始数据源。')
  if (draft.config.type === 'pareto' && !statistics.pareto) throw new RangeError('图表设置统计读取未返回当前帕累托分类。')
  if (draft.config.type === 'histogram' && !statistics.histogram) throw new RangeError('图表设置统计读取未返回当前直方图分箱。')
  return { ...chartSettingsPreparedData(draft, statistics.preparedData, identity, date1904), statistics }
}

/** Bound service data bypasses the UI preview budget and is never reprojected.
 * Unbound records still use the bounded projection for standalone consumers. */
export function projectChartSettings(draft: ChartSettingsDraft, sheet?: ChartSettingsSourceSheet, provided?: ChartData | ChartSettingsProvidedData, date1904 = chartSettingsProvidedValue(provided)?.date1904, identity: ChartSettingsSourceIdentity = {}): ChartSettingsProjection {
  let source: ChartRecords
  try {
    if (provided && 'kind' in provided) {
      if (provided.kind !== 'prepared-source') return { needsPreparation: true, error: '图表统计结果不能作为原始数据源，请重新读取数据源。' }
      const expectedKey = chartSettingsBindingKey(draft, identity.hostSheetId ?? sheet?.id, date1904)
      if (provided.bindingKey !== expectedKey || provided.sourceRevision !== identity.sourceRevision) {
        return { needsPreparation: true, error: '图表数据源或字段绑定已更新，请重新读取数据源。' }
      }
      const data = date1904 === undefined || date1904 === provided.data.date1904 ? provided.data : { ...provided.data, date1904 }
      return { data, statistics: provided.statistics, prepared: true, ...(data.fields && data.records ? { source: settingsSourceRecords(data) } : {}) }
    }
    if (draft.config.pivotTableId || !sheet) {
      if (!provided?.fields || !provided.records) return { needsPreparation: true,
        ...(provided ? { error: '尚未取得与当前字段绑定一致的原始数据源，请重新读取数据源。' } : {}) }
      previewBudget(provided.fields.length, provided.records.length + 1)
      source = settingsSourceRecords(provided)
    } else source = chartSettingsRecords(draft, sheet)
    if (date1904 !== undefined) source = { ...source, date1904 }
  } catch (reason) { return { error: reason instanceof Error ? reason.message : String(reason),
    ...(reason instanceof ChartSettingsPreparationRequired ? { needsPreparation: true } : {}) } }
  try {
    return { source, data: projectChartRecords(source, { type: draft.config.type, categoryField: draft.config.categoryField ?? 0,
      valueFields: draft.config.valueFields ?? [], seriesToCategory: draft.config.seriesToCategory, options: draft.config.options }) }
  } catch (reason) { return { source, error: reason instanceof Error ? reason.message : String(reason) } }
}

/** Large prepared sources use constant-size author-option preflight. Full
 * source/derived date and final tick validation belongs to candidate apply. */
export function chartSettingsDateAxisError(config: ChartConfig, data?: ChartData, date1904 = data?.date1904, validation: 'source' | 'options' = 'source'): string | undefined {
  if (!config.options?.mapping?.useDateAxis || !supportsChartDateAxis(config)) return undefined
  const chart: ChartDateAxisConfig = { type: config.type, options: config.options, xAxis: { scaleType: 'linear', ...config.xAxis } }
  const current: ChartData = { ...(validation === 'options' ? { categories: [], series: [] } : data ?? { categories: [], series: [] }), date1904 }
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
    throw new ChartSettingsPreparationRequired('图表数据源较大，请读取完整数据源后编辑原始样本。')
  }
}

class ChartSettingsPreparationRequired extends RangeError {}

function settingsSourceRecords(data: ChartData): ChartRecords {
  return { fields: data.fields!, records: data.records!, date1904: data.date1904,
    formattedRecords: data.formattedRecords, sourceCategoryFields: data.sourceCategoryFields, sourceValueFields: data.sourceValueFields, sourceNumericFields: data.sourceNumericFields,
    visibleFields: data.visibleFields, syntheticValues: data.syntheticValues }
}
