import type {
  CellRange,
  CellScalar,
  ChartAxisOptions,
  ChartLegendOptions,
  ChartModel,
  ChartSemanticOptions,
  ChartSource,
  ChartType,
  SpreadsheetSheet,
  WorkbookModel,
} from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { randomId } from '../../api/model-utils'
import { readPivotChartSource } from './chart-pivot-data'
import { resolveChartType } from './chart-catalog'
import { chartField, readSpecialChartData, type ChartCandle, type ChartGraphData, type ChartHierarchyNode, type ChartRecords } from './chart-data'
import { cloneChartOptions, mergeChartOptions, validateChartPalette, validateChartAxis, validateChartFont } from './chart-options'
import { chartCompoundPieData, type ChartComputedStatistics } from './chart-derived'
import { layoutChartData } from './chart-layout'
import { drawChartData } from './chart-draw'
import { chartSourceVisibility } from './chart-visibility'

export interface ChartConfig {
  type: ChartType
  source: ChartSource
  position: { x: number; y: number }
  size: { width: number; height: number }
  title?: string
  subtitle?: string
  palette?: string[]
  options?: ChartSemanticOptions
  legend?: Partial<ChartLegendOptions>
  categoryField?: number
  valueFields?: number[]
  seriesToCategory?: Record<number, number[]>
  pivotTableId?: string
  xAxis?: Partial<ChartAxisOptions>
  yAxis?: Partial<ChartAxisOptions>
}

export interface ChartSeriesData {
  field: number
  name: string
  values: Array<number | null>
  xValues?: Array<number | null>
  sizeValues?: Array<number | null>
}

export interface ChartData {
  categories: string[]
  series: ChartSeriesData[]
  fields?: string[]
  records?: CellScalar[][]
  graph?: ChartGraphData
  hierarchy?: ChartHierarchyNode[]
  candles?: Array<ChartCandle | null>
}

export interface ChartPoint { x: number; y: number; value: number; category: string; categoryIndex: number; baseY?: number; baseX?: number; baseValue?: number; xValue?: number; size?: number }
export interface ChartSeriesLayout {
  name: string
  color: string
  field?: number
  type?: 'column' | 'bar' | 'line' | 'area'
  stacked?: boolean
  rightAxis?: boolean
  points: ChartPoint[]
}
export interface ChartLayout {
  frame: RenderRect
  plot: RenderRect
  min: number
  max: number
  xMin?: number
  xMax?: number
  rightMin?: number
  rightMax?: number
  horizontal?: boolean
  categories: string[]
  series: ChartSeriesLayout[]
}

const DEFAULT_LEGEND: ChartLegendOptions = { visible: true, position: 'bottom' }
const DEFAULT_AXIS: ChartAxisOptions = { scaleType: 'linear' }

export function createChart(workbook: WorkbookModel, sheetId: string, config: ChartConfig, id = randomId('chart')): ChartModel {
  const sheet = requireSheet(workbook, sheetId)
  const source = normalizeSource(config.source)
  if (!config.pivotTableId) validateSourceInSheet(source.range, source.sheetId ? requireSheet(workbook, source.sheetId) : sheet)
  const type = config.type
  validateChartType(type)
  const position = normalizeSize(config.position, 'chart position', true)
  const size = normalizeSize(config.size, 'chart size', false)
  const categoryField = normalizeField(config.categoryField ?? 0, 'categoryField')
  const valueFields = config.valueFields?.map((field) => normalizeField(field, 'valueField'))
  if (valueFields && new Set(valueFields).size !== valueFields.length) throw new RangeError('Chart value fields must be unique')
  validateChartPalette(config.palette)
  validateChartText(config.title, '标题'); validateChartText(config.subtitle, '副标题')
  const model: ChartModel = {
    id,
    sheetId,
    type,
    source,
    position,
    size,
    title: config.title,
    subtitle: config.subtitle,
    palette: config.palette ? [...config.palette] : undefined,
    options: cloneChartOptions(config.options),
    legend: cloneChartLegend({ ...DEFAULT_LEGEND, ...(config.legend ?? {}) }),
    categoryField,
    valueFields: valueFields ?? [],
    ...(config.seriesToCategory ? { seriesToCategory: cloneSeriesCategories(config.seriesToCategory) } : {}),
    pivotTableId: config.pivotTableId,
    xAxis: cloneChartAxis({ ...DEFAULT_AXIS, ...(config.xAxis ?? {}) }),
    yAxis: cloneChartAxis({ ...DEFAULT_AXIS, ...(config.yAxis ?? {}) }),
    revision: 0,
  }
  validateAxis(model.xAxis, 'xAxis')
  validateAxis(model.yAxis, 'yAxis')
  validateChartData(model, readChartData(workbook, model), model.yAxis)
  if (model.pivotTableId && !workbook.pivotTables?.some((pivot) => pivot.id === model.pivotTableId)) throw new Error(`Pivot table not found: ${model.pivotTableId}`)
  return model
}

export function updateChart(workbook: WorkbookModel, chart: ChartModel, patch: Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>): void {
  const next: ChartModel = {
    ...chart,
    ...patch,
    source: patch.source ? normalizeSource(patch.source) : chart.source,
    position: patch.position ? normalizeSize(patch.position, 'chart position', true) : chart.position,
    size: patch.size ? normalizeSize(patch.size, 'chart size', false) : chart.size,
    legend: cloneChartLegend({ ...chart.legend, ...patch.legend }),
    xAxis: cloneChartAxis({ ...chart.xAxis, ...patch.xAxis }),
    yAxis: cloneChartAxis({ ...chart.yAxis, ...patch.yAxis }),
    valueFields: patch.valueFields ? [...patch.valueFields] : chart.valueFields,
    seriesToCategory: cloneSeriesCategories(patch.seriesToCategory ?? chart.seriesToCategory ?? {}),
    options: mergeChartOptions(chart.options, patch.options),
    palette: Object.prototype.hasOwnProperty.call(patch, 'palette') ? patch.palette ? [...patch.palette] : undefined : chart.palette,
    revision: chart.revision + 1,
  }
  const sheet = requireSheet(workbook, chart.sheetId)
  if (!next.pivotTableId) validateSourceInSheet(next.source.range, next.source.sheetId ? requireSheet(workbook, next.source.sheetId) : sheet)
  validateChartType(next.type)
  validateChartPalette(next.palette)
  validateChartText(next.title, '标题'); validateChartText(next.subtitle, '副标题')
  next.categoryField = normalizeField(next.categoryField, 'categoryField')
  next.valueFields = next.valueFields.map((field) => normalizeField(field, 'valueField'))
  if (new Set(next.valueFields).size !== next.valueFields.length) throw new RangeError('Chart value fields must be unique')
  validateAxis(next.xAxis, 'xAxis')
  validateAxis(next.yAxis, 'yAxis')
  validateChartData(next, readChartData(workbook, next), next.yAxis)
  Object.assign(chart, next)
}

export function readChartData(workbook: WorkbookModel, chart: ChartModel): ChartData {
  if (chart.pivotTableId) {
    const data = readPivotChartData(readPivotChartSource(workbook, chart.pivotTableId), chart)
    if (data.fields && data.records) {
      const special = readSpecialChartData({ fields: data.fields, records: data.records }, resolveChartType(chart.type).dataShape === 'edge' ? chart : { ...chart, valueFields: chart.valueFields.length ? chart.valueFields : data.series.map(series => series.field) })
      if (special) return special
    }
    if (chart.type === 'scatter') {
      if (data.fields && data.records) return readScatterChartData({ fields: data.fields, records: data.records }, chart, data.series.map(series => series.field))
      return { ...data, series: data.series.map(series => ({ ...series, xValues: data.categories.map(numericValue) })) }
    }
    return data
  }
  const sheet = requireSheet(workbook, chart.source.sheetId ?? chart.sheetId)
  const range = chart.source.range
  const matrix = readMatrix(sheet, range)
  const vectors = chart.source.orientation === 'rows' ? matrix[0]?.map((_cell, index) => matrix.map(row => row[index] ?? null)) ?? [] : matrix
  if (vectors.length < 2 || !vectors[0]?.length) throw new RangeError('图表数据区域须包含字段名和至少一条数据。')
  const visibleOnly = chart.options?.plotVisibleOnly === true
  const visibility = visibleOnly ? chartSourceVisibility(workbook, sheet, range) : undefined
  const records: ChartRecords = { fields: vectors[0]!.map(scalarLabel), records: vectors.slice(1).filter((_record, index) => {
    if (!visibleOnly) return true
    const position = chart.source.orientation === 'columns' ? range.startRow + index + 1 : range.startCol + index + 1
    return chart.source.orientation === 'columns' ? visibility!.rowVisible(position) : visibility!.columnVisible(position)
  }) }
  const categoryFields = chart.options?.mapping?.categoryIndexes ?? [chart.categoryField]
  const availableFields = records.fields.map((_value, index) => index).filter(index => {
    if (categoryFields.includes(index) && resolveChartType(chart.type).dataShape !== 'distribution' || !records.records.some(record => numericValue(record[index]) !== null)) return false
    if (!visibleOnly) return true
    const position = chart.source.orientation === 'columns' ? range.startCol + index : range.startRow + index
    return chart.source.orientation === 'columns' ? visibility!.columnVisible(position) : visibility!.rowVisible(position)
  })
  const shape = resolveChartType(chart.type).dataShape
  const specialChart = shape === 'edge' || shape === 'distribution' ? chart : { ...chart, valueFields: chart.valueFields.length ? chart.valueFields : availableFields }
  const special = readSpecialChartData(records, specialChart)
  if (special) return special
  if (vectors[0]!.length < 2) throw new RangeError('Chart source must contain a category and at least one data value')
  if (chart.type === 'scatter') return readScatterChartData(records, chart, availableFields)
  for (const field of categoryFields) chartField(records, field, '分类字段')
  const valueFields = chart.valueFields.length ? chart.valueFields : availableFields
  for (const field of valueFields) chartField(records, field, '数值字段')
  const categories = records.records.map(row => categoryFields.map(field => scalarLabel(row[field])).join(' / '))
  const series = valueFields.map((field) => ({
    field,
    name: records.fields[field] || `Series ${field + 1}`,
    values: records.records.map(row => numericValue(row[field])),
  }))
  return { ...records, categories, series }
}

function readPivotChartData(source: ReturnType<typeof readPivotChartSource>, chart: ChartModel): ChartData {
  const categoryFields = chart.options?.mapping?.categoryIndexes ?? (chart.categoryField === 0 ? source.categoryFields : [chart.categoryField])
  for (const field of categoryFields) chartField(source, field, '分类字段')
  const valueFields = chart.valueFields.length && !(source.syntheticValues && chart.valueFields.includes(0)) ? chart.valueFields : source.valueFields
  if (source.syntheticValues && chart.valueFields.some(field => field > source.valueFields.length)) throw new RangeError('Pivot chart value field is outside pivot output')
  for (const field of valueFields) if (!source.valueFields.includes(field)) throw new RangeError('Pivot chart value field is outside pivot output')
  return { fields: source.fields, records: source.records,
    categories: source.records.map(row => categoryFields.map(field => scalarLabel(row[field])).join(' / ')),
    series: valueFields.map(field => ({ field, name: source.fields[field] || `Series ${field + 1}`, values: source.records.map(row => numericValue(row[field])) })) }
}

function cloneSeriesCategories(mapping: Record<number, number[]>): Record<number, number[]> {
  return Object.fromEntries(Object.entries(mapping).map(([field, categories]) => {
    normalizeField(Number(field), 'seriesToCategory value field')
    if (!Array.isArray(categories) || categories.length !== 1) throw new RangeError('Scatter series require exactly one X field')
    return [field, categories.map((category) => normalizeField(category, 'seriesToCategory X field'))]
  }))
}

function readScatterChartData(source: ChartRecords, chart: ChartModel, availableFields: number[]): ChartData {
  const categoryField = normalizeField(chart.categoryField, 'categoryField')
  chartField(source, categoryField, 'X 字段')
  const mapping = cloneSeriesCategories(chart.seriesToCategory ?? {})
  const xFields = new Set([categoryField, ...Object.values(mapping).flat()])
  const fields = chart.valueFields.length ? chart.valueFields : availableFields.filter(field => !xFields.has(field))
  for (const [valueField, xField] of Object.entries(mapping)) {
    if (!fields.includes(Number(valueField))) throw new RangeError('Scatter series mapping is outside selected source fields')
    chartField(source, xField[0]!, 'X 字段')
  }
  return {
    ...source,
    categories: source.records.map(record => scalarLabel(record[categoryField])),
    series: fields.map((field) => {
      chartField(source, field, 'Y 字段')
      const xField = mapping[field]?.[0] ?? categoryField
      return {
        field,
        name: source.fields[field] || `Series ${field + 1}`,
        values: source.records.map(record => numericValue(record[field])),
        xValues: source.records.map(record => numericValue(record[xField])),
      }
    }),
  }
}

export function chartLayout(workbook: WorkbookModel, chart: ChartModel, computed?: ChartComputedStatistics): ChartLayout {
  return layoutChartData(readChartData(workbook, chart), chart, computed)
}

export function chartRenderCommands(workbook: WorkbookModel, chart: ChartModel, computed?: ChartComputedStatistics): DrawCommand[] {
  return drawChartData(readChartData(workbook, chart), chart, {}, computed)
}

export function radarValueExtent(data: ChartData): { min: number; max: number } {
  let min = 0, max = 0
  for (const series of data.series) for (const value of series.values) if (value !== null && Number.isFinite(value)) {
    min = Math.min(min, value); max = Math.max(max, value)
  }
  return { min, max: min === max ? min + 1 : max }
}
export function normalizeSource(source: ChartSource): ChartSource {
  if (!source || !source.range) throw new TypeError('Chart source is required')
  const range = normalizeRange(source.range)
  if (source.orientation !== 'columns' && source.orientation !== 'rows') throw new TypeError(`Invalid chart source orientation: ${source.orientation}`)
  if (source.sheetId !== undefined && (typeof source.sheetId !== 'string' || !source.sheetId.trim())) throw new TypeError('Invalid chart source worksheet')
  return { range, orientation: source.orientation, ...(source.sheetId === undefined ? {} : { sheetId: source.sheetId }) }
}

function validateChartType(type: ChartType): void {
  resolveChartType(type)
}

function validateChartData(chart: ChartModel, data: ChartData, axis: ChartAxisOptions): void {
  const type = chart.type, descriptor = resolveChartType(type)
  if (!data.categories.length || !data.series.length) throw new RangeError('Chart source must contain at least one category and value series')
  if (type === 'radar' && data.categories.length < 3) throw new RangeError('Radar charts require at least three categories')
  if (data.series.some((series) => !series.values.length || series.values.length !== data.categories.length)) throw new RangeError('Chart series must have one value per category')
  if ((type === 'scatter' || type === 'bubble') && data.series.some((series) => !series.xValues || series.xValues.length !== series.values.length)) throw new RangeError('Scatter series must have one X value per Y value')
  if ((type === 'scatter' || type === 'bubble') && !data.series.some((series) => series.values.some((value, index) => value !== null && series.xValues?.[index] != null))) throw new RangeError('Scatter chart requires at least one numeric X/Y pair')
  if (descriptor.stacked && axis.scaleType !== 'linear') throw new TypeError('堆积图须使用线性数值轴。')
  if (type === 'radar' && axis.scaleType !== 'linear') throw new TypeError('Radar charts require a linear value axis')
  if (descriptor.family === 'pie') {
    if (data.series.length !== 1) throw new RangeError(`${type} charts require exactly one value series`)
    const values = data.series[0]!.values.filter((value): value is number => value !== null && Number.isFinite(value))
    if (!values.length || values.some((value) => value < 0) || values.every((value) => value === 0)) throw new RangeError(`${type} charts require non-negative values with a positive total`)
  }
  if (type === 'bubble' && data.series.some(series => series.sizeValues?.some(value => value !== null && value < 0))) throw new RangeError('气泡大小不能为负数。')
  if (type === 'candlestick' && !data.candles?.some(Boolean)) throw new RangeError('K 线须至少包含一条完整有效的 HLC 或 OHLC 数据。')
  if (type === 'histogram' && data.series.length !== 1) throw new RangeError('直方图须选择一个数值样本字段。')
  if (type === 'pareto' && data.series.length !== 1) throw new RangeError('帕累托图须选择一个数值权重字段。')
  if ((type === 'funnel' || type === 'wordCloud') && data.series.length !== 1) throw new RangeError('漏斗和词云图须选择一个权重字段。')
  if ((type === 'pareto' || type === 'wordCloud') && data.series.some(series => series.values.some(value => value !== null && value < 0))) throw new RangeError('此图表的权重不能为负数。')
}

function readMatrix(sheet: SpreadsheetSheet, range: CellRange): CellScalar[][] {
  return Array.from({ length: range.endRow - range.startRow + 1 }, (_unused, rowOffset) => Array.from({ length: range.endCol - range.startCol + 1 }, (_value, colOffset) => {
    const cell = sheet.cells[`${range.startRow + rowOffset}:${range.startCol + colOffset}`]
    if (!cell) return null
    if (cell.formula && cell.display !== '') {
      const cached = numericValue(cell.display)
      if (cached !== null) return cached
    }
    return cell.value
  }))
}

function validateSourceInSheet(range: CellRange, sheet: SpreadsheetSheet): void {
  const maxRow = (sheet.rowCount ?? Number.MAX_SAFE_INTEGER) - 1, maxCol = (sheet.columnCount ?? Number.MAX_SAFE_INTEGER) - 1
  if (range.endRow > maxRow || range.endCol > maxCol) throw new RangeError('Chart source is outside worksheet bounds')
}

function normalizeRange(range: CellRange): CellRange {
  if (!range || typeof range !== 'object') throw new TypeError('Chart source range is required')
  for (const name of ['startRow', 'endRow', 'startCol', 'endCol'] as const) {
    const value = range[name], max = name.endsWith('Row') ? 1_048_575 : 16_383
    if (!Number.isSafeInteger(value) || value < 0 || value > max) throw new RangeError(`Invalid chart range ${name}`)
  }
  if (range.endRow < range.startRow || range.endCol < range.startCol) throw new RangeError('Invalid chart source range')
  return { startRow: range.startRow, startCol: range.startCol, endRow: range.endRow, endCol: range.endCol }
}

function normalizeSize(value: { x: number; y: number } | { width: number; height: number }, name: string, position: true): { x: number; y: number }
function normalizeSize(value: { x: number; y: number } | { width: number; height: number }, name: string, position: false): { width: number; height: number }
function normalizeSize(value: { x?: number; y?: number; width?: number; height?: number }, name: string, position: boolean): { x: number; y: number } | { width: number; height: number } {
  const first = position ? value.x : value.width, second = position ? value.y : value.height
  if (typeof first !== 'number' || !Number.isFinite(first) || typeof second !== 'number' || !Number.isFinite(second) || (position ? first < 0 || second < 0 : first <= 0 || second <= 0)) throw new RangeError(`Invalid ${name}`)
  return position ? { x: first, y: second } : { width: first, height: second }
}

function normalizeField(value: number, name: string): number {
  if (!Number.isSafeInteger(value) || value < 0 || value > 1_048_575) throw new RangeError(`Invalid chart ${name}`)
  return value
}

function validateAxis(axis: ChartAxisOptions, name: string): void {
  validateChartAxis(axis, name)
  if (axis.scaleType !== 'linear' && axis.scaleType !== 'log') throw new TypeError(`Invalid ${name} scale type`)
  if (axis.logBase !== undefined && (!Number.isFinite(axis.logBase) || axis.logBase < 2 || axis.logBase > 1000)) throw new RangeError(`${name} log base must be between 2 and 1000`)
  for (const bound of [axis.min, axis.max]) if (bound !== undefined && (!Number.isFinite(bound) || axis.scaleType === 'log' && bound <= 0)) throw new RangeError(`${name} bounds must be finite and positive for a log axis`)
  if (axis.min !== undefined && axis.max !== undefined && axis.min >= axis.max) throw new RangeError(`${name} minimum must be smaller than maximum`)
}

function cloneChartAxis(axis: ChartAxisOptions): ChartAxisOptions {
  validateAxis(axis, 'Chart axis')
  return cloneChartOptions({ rightYAxis: axis })!.rightYAxis!
}

function cloneChartLegend(legend: ChartLegendOptions): ChartLegendOptions {
  if (typeof legend.visible !== 'boolean' || !['top', 'right', 'bottom', 'left'].includes(legend.position)) throw new TypeError('Invalid chart legend')
  validateChartFont(legend.font, '图例字体')
  return { ...legend, ...(legend.font ? { font: { ...legend.font } } : {}) }
}

function validateChartText(value: string | undefined, name: string): void {
  if (value !== undefined && (typeof value !== 'string' || value.length > 100_000)) throw new TypeError(`图表${name}须为不超过 100000 字符的文本。`)
}

function numericValue(value: CellScalar | undefined): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value))) return Number(value)
  return null
}

function scalarLabel(value: CellScalar | undefined): string { return value == null ? '' : String(value) }

function requireSheet(workbook: WorkbookModel, sheetId: string): SpreadsheetSheet {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
  return sheet
}
