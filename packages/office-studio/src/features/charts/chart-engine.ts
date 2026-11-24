import type {
  CellRange,
  CellScalar,
  ChartAxisOptions,
  ChartLegendOptions,
  ChartModel,
  ChartSemanticOptions,
  ChartSource,
  ChartType,
  ChartWaterfallStyleTarget,
  SpreadsheetSheet,
  WorkbookModel,
} from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { randomId } from '../../api/model-utils'
import { resolveChartType } from './chart-catalog'
import { type ChartCandle, type ChartGraphData, type ChartHierarchyNode, type ChartRecords } from './chart-data'
import { cloneChartSeriesCategories, projectChartRecords } from './chart-record-projection'
import { cloneChartOptions, mergeChartOptions, chartSeriesStyle, validateChartPalette, validateChartAxis, validateChartFont } from './chart-options'
import { validateChartNamedSeriesBindings } from './chart-series-selector'
import { prepareChartDateSource } from './chart-date-axis'
import { chartSeriesKind, type ChartComputedStatistics } from './chart-derived'
import { convertChartType } from './chart-type-conversion'
import { layoutChartData } from './chart-layout'
import { drawChartData } from './chart-draw'
import { readChartRecordSteps } from './chart-source-compute'
import { chartCellPosition, chartPlacementGeometry } from './chart-placement'
import { validateChartGradientAuthoring } from './chart-gradient'

export interface ChartConfig {
  type: ChartType
  source: ChartSource
  position: { x: number; y: number }
  /** Builder placement shortcut; resolved against host axes before creating the model. */
  positionCell?: string
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
  /** Derived labels may change while named selectors keep the original source identity. */
  sourceName?: string
  values: Array<number | null>
  xValues?: Array<number | null>
  sizeValues?: Array<number | null>
}

export interface ChartData {
  categories: string[]
  /** Raw base-axis values retained through aggregation and category transforms. */
  categoryValues?: CellScalar[]
  /** Derived Excel serial days; never persisted in the chart model. */
  categoryDateValues?: Array<number | null>
  series: ChartSeriesData[]
  /** Runtime date-system owner inherited from the source workbook. */
  date1904?: boolean
  fields?: string[]
  records?: CellScalar[][]
  /** Source defaults stay independent of the chart's current authored selection. */
  sourceCategoryFields?: number[]
  sourceValueFields?: number[]
  visibleFields?: number[]
  syntheticValues?: boolean
  graph?: ChartGraphData
  hierarchy?: ChartHierarchyNode[]
  candles?: Array<ChartCandle | null>
}

export interface ChartPoint { x: number; y: number; value: number; category: string; categoryIndex: number; baseY?: number; baseX?: number; baseValue?: number; xValue?: number; size?: number; rightAxis?: boolean; waterfallRole?: ChartWaterfallStyleTarget }
export interface ChartWaterfallPointLayout extends ChartPoint {
  seriesPosition: number
  dataPointIndex: number
  synthetic: boolean
  baseY: number
  rightAxis: boolean
  waterfallRole: ChartWaterfallStyleTarget
}
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
  dateAxis?: boolean
  date1904?: boolean
  categoryPositions?: Array<number | null>
  categoryBandwidth?: number
  categories: string[]
  series: ChartSeriesLayout[]
  waterfall?: {
    points: ChartWaterfallPointLayout[]
    connectors: Array<{ from: { x: number; y: number }; to: { x: number; y: number } }>
  }
}

const DEFAULT_LEGEND: ChartLegendOptions = { visible: true, position: 'bottom' }
const DEFAULT_AXIS: ChartAxisOptions = { scaleType: 'linear' }

export function createChart(workbook: WorkbookModel, sheetId: string, config: ChartConfig, id = randomId('chart'), preloaded?: ChartRecords): ChartModel {
  const sheet = requireSheet(workbook, sheetId)
  const source = normalizeSource(config.source)
  if (!config.pivotTableId) validateSourceInSheet(source.range, source.sheetId ? requireSheet(workbook, source.sheetId) : sheet)
  const type = config.type
  validateChartType(type)
  const position = config.positionCell === undefined ? normalizeSize(config.position, 'chart position', true)
    : chartCellPosition(config.positionCell, chartPlacementGeometry(workbook, sheetId))
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
    ...(config.seriesToCategory ? { seriesToCategory: cloneChartSeriesCategories(config.seriesToCategory) } : {}),
    pivotTableId: config.pivotTableId,
    xAxis: cloneChartAxis({ ...DEFAULT_AXIS, ...(config.xAxis ?? {}) }),
    yAxis: cloneChartAxis({ ...DEFAULT_AXIS, ...(config.yAxis ?? {}) }),
    revision: 0,
  }
  validateChartGradientAuthoring(model.type, model.options)
  validateAxis(model.xAxis, 'xAxis')
  validateAxis(model.yAxis, 'yAxis')
  validateChartData(model, readChartData(workbook, model, preloaded), model.yAxis)
  if (model.pivotTableId && !workbook.pivotTables?.some((pivot) => pivot.id === model.pivotTableId)) throw new Error(`Pivot table not found: ${model.pivotTableId}`)
  return model
}

export function updateChart(workbook: WorkbookModel, chart: ChartModel, patch: Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>, preloaded?: ChartRecords): void {
  const source = patch.source ? normalizeSource(patch.source) : chart.source
  const pivotTableId = Object.prototype.hasOwnProperty.call(patch, 'pivotTableId') ? patch.pivotTableId : chart.pivotTableId
  const sourceChart = { ...chart, source, pivotTableId, options: mergeChartOptions(chart.options, patch.options) }
  validateChartGradientAuthoring(patch.type ?? chart.type, sourceChart.options, chart.options)
  const sheet = requireSheet(workbook, chart.sheetId)
  if (!pivotTableId) validateSourceInSheet(source.range, source.sheetId ? requireSheet(workbook, source.sheetId) : sheet)
  validateChartType(patch.type ?? chart.type)
  const records = preloaded ?? readChartRecords(workbook, sourceChart)
  const defaults = patch.type !== undefined && patch.type !== chart.type ? convertChartType(sourceChart, patch.type, records) : chart
  const next: ChartModel = {
    ...chart,
    ...defaults,
    ...patch,
    source,
    pivotTableId,
    position: patch.position ? normalizeSize(patch.position, 'chart position', true) : chart.position,
    size: patch.size ? normalizeSize(patch.size, 'chart size', false) : chart.size,
    legend: cloneChartLegend({ ...DEFAULT_LEGEND, ...defaults.legend, ...patch.legend }),
    xAxis: cloneChartAxis({ ...DEFAULT_AXIS, ...defaults.xAxis, ...patch.xAxis }),
    yAxis: cloneChartAxis({ ...DEFAULT_AXIS, ...defaults.yAxis, ...patch.yAxis }),
    valueFields: patch.valueFields ? [...patch.valueFields] : defaults.valueFields ? [...defaults.valueFields] : [],
    categoryField: patch.categoryField ?? defaults.categoryField ?? 0,
    seriesToCategory: cloneChartSeriesCategories(patch.seriesToCategory ?? defaults.seriesToCategory ?? {}),
    options: mergeChartOptions(defaults.options, patch.options),
    palette: Object.prototype.hasOwnProperty.call(patch, 'palette') ? patch.palette ? [...patch.palette] : undefined : chart.palette,
    revision: chart.revision + 1,
  }
  validateChartPalette(next.palette)
  validateChartText(next.title, '标题'); validateChartText(next.subtitle, '副标题')
  next.categoryField = normalizeField(next.categoryField, 'categoryField')
  next.valueFields = next.valueFields.map((field) => normalizeField(field, 'valueField'))
  if (new Set(next.valueFields).size !== next.valueFields.length) throw new RangeError('Chart value fields must be unique')
  validateAxis(next.xAxis, 'xAxis')
  validateAxis(next.yAxis, 'yAxis')
  validateChartData(next, readChartData(workbook, next, records), next.yAxis)
  Object.assign(chart, next)
}

/** Read the current source without interpreting the chart's field roles. */
export function readChartRecords(workbook: WorkbookModel, chart: Pick<ChartModel, 'sheetId' | 'source' | 'pivotTableId' | 'options'>): ChartRecords {
  const steps = readChartRecordSteps(workbook, chart)
  while (true) { const step = steps.next(); if (step.done) return step.value }
}

export function readChartData(workbook: WorkbookModel, chart: ChartModel, preloaded?: ChartRecords): ChartData {
  const data = projectChartRecords(preloaded ?? readChartRecords(workbook, chart), chart)
  validateChartNamedSeriesBindings(chart.options, data.series)
  return data
}

export function chartLayout(workbook: WorkbookModel, chart: ChartModel, computed?: ChartComputedStatistics): ChartLayout {
  return layoutChartData(computed?.preparedData ?? readChartData(workbook, chart), chart, computed)
}

export function chartRenderCommands(workbook: WorkbookModel, chart: ChartModel, computed?: ChartComputedStatistics): DrawCommand[] {
  return drawChartData(computed?.preparedData ?? readChartData(workbook, chart), chart, {}, computed)
}

export { chartRadarExtent as radarValueExtent } from './chart-radar'
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
  data = prepareChartDateSource(data, chart)
  const type = chart.type, descriptor = resolveChartType(type)
  if (descriptor.stacked && axis.scaleType !== 'linear') throw new TypeError('堆积图须使用线性数值轴。')
  data.series.forEach((series, index) => {
    const style = chartSeriesStyle(chart, series.field, index, undefined, series.sourceName ?? series.name)
    const valueAxis = style.rightAxis ? chart.options?.rightYAxis ?? DEFAULT_AXIS : axis
    if (chartSeriesKind(chart, series.field, index, series.sourceName ?? series.name).stacked && valueAxis.scaleType !== 'linear') throw new TypeError('堆积系列须使用线性数值轴。')
  })
  if (type === 'radar' && axis.scaleType !== 'linear') throw new TypeError('Radar charts require a linear value axis')
  if (data.visibleFields !== undefined && (!data.categories.length || !data.series.length)) return
  if (!data.categories.length || !data.series.length) throw new RangeError('Chart source must contain at least one category and value series')
  if (type === 'radar' && data.categories.length < 3) throw new RangeError('Radar charts require at least three categories')
  if (data.series.some((series) => !series.values.length || series.values.length !== data.categories.length)) throw new RangeError('Chart series must have one value per category')
  if ((type === 'scatter' || type === 'bubble') && data.series.some((series) => !series.xValues || series.xValues.length !== series.values.length)) throw new RangeError('Scatter series must have one X value per Y value')
  if ((type === 'scatter' || type === 'bubble') && !data.series.some((series) => series.values.some((value, index) => value !== null && series.xValues?.[index] != null))) throw new RangeError('Scatter chart requires at least one numeric X/Y pair')
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
  if (type === 'wordCloud' && data.series.some(series => series.values.some(value => value !== null && value < 0))) throw new RangeError('此图表的权重不能为负数。')
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

function requireSheet(workbook: WorkbookModel, sheetId: string): SpreadsheetSheet {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
  return sheet
}
