import type {
  CellRange,
  CellScalar,
  ChartAxisOptions,
  ChartCandlestickMappingOptions,
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
import { cloneChartOptions, definedChartStyleFields, mergeChartOptions, validateChartPalette, validateChartAxis } from './chart-options'
import { validateChartNamedSeriesBindings } from './chart-series-selector'
import type { ChartComputedStatistics } from './chart-derived'
import { convertChartType } from './chart-type-conversion'
import { layoutChartData } from './chart-layout'
import { drawChartData } from './chart-draw'
import { readChartRecordSteps } from './chart-source-compute'
import { chartCellPosition, chartPlacementGeometry } from './chart-placement'
import { validateChartGradientAuthoring } from './chart-gradient'
import { initializeChartWaterfallOptions } from './chart-waterfall'
import type { ChartSankeyLayout } from './chart-sankey-layout'
import type { ChartTreeRect } from './chart-special-layout'
import type { ChartSunburstSector } from './chart-sunburst-layout'
import type { ChartChordData } from './chart-chord-data'
import type { ChartChordLayout } from './chart-chord-layout'
import type { ChartCloudWord, ChartWordCloudLayoutResult } from './chart-word-cloud-layout'
import type { ChartRelationGraphLayout, ChartRelationLayoutResult } from './chart-relation-layout'
import { validateChartSourceData } from './chart-source-validation'
import { assertChartSourcePresent } from './chart-source-diagnostic'
import { chartLegendPatch } from './chart-legend'
import { drawingAffineFrame, drawingAffineWithFrame, normalizeDrawingAffineGeometry, type DrawingAffineGeometry } from '../../domain/drawing-affine.ts'
import { GRID_HEADER_WIDTH, GRID_HEADER_HEIGHT } from '../../rendering/spreadsheet/layout.ts'

export interface ChartConfig {
  type: ChartType
  source: ChartSource
  position: { x: number; y: number }
  /** Builder placement shortcut; resolved against host axes before creating the model. */
  positionCell?: string
  size: { width: number; height: number }
  drawingGeometry?: DrawingAffineGeometry
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

export type ChartUpdatePatch = Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>

export interface ChartSeriesData {
  field: number
  name: string
  /** Derived labels may change while named selectors keep the original source identity. */
  sourceName?: string
  values: Array<number | null>
  /** Source number formatting aligned to values; statistics still consume values only. */
  formattedValues?: Array<string | null>
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
  formattedRecords?: Array<Array<string | null>>
  /** Source defaults stay independent of the chart's current authored selection. */
  sourceCategoryFields?: number[]
  sourceValueFields?: number[]
  sourceNumericFields?: number[]
  visibleFields?: number[]
  syntheticValues?: boolean
  graph?: ChartGraphData
  chord?: ChartChordData
  /** Browser-measured glyphs are placed once by the chart Worker/Rust kernel. */
  wordCloud?: ChartWordCloudLayoutResult
  /** Plot-relative node positions from the chart Worker/Rust kernel. */
  relation?: ChartRelationLayoutResult
  hierarchy?: ChartHierarchyNode[]
  hierarchyPaletteOrder?: string[]
  hierarchyValueFieldName?: string
  hierarchyDiagnostic?: string
  candles?: Array<ChartCandle | null>
  /** Explicit source roles survive all-empty HLC/OHLC data and runtime filtering. */
  candleFields?: ChartCandlestickMappingOptions
}

export interface ChartPoint { x: number; y: number; value: number; category: string; categoryIndex: number; sequenceIndex?: number; baseY?: number; baseX?: number; baseValue?: number; xValue?: number; size?: number; rightAxis?: boolean; waterfallRole?: ChartWaterfallStyleTarget }
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
  /** Calendar-axis coordinates and labels share one derived tick sequence. */
  dateTicks?: Array<{ value: number; serial: number }>
  /** Ordinary Cartesian sources use stable date order with separate duplicate
   * observations; Pareto preserves its weight order. Indices retain source identity. */
  dateOrder?: number[]
  /** Source bounds stay in Excel serials even when xMin/xMax are calendar buckets. */
  dateSourceMin?: number
  dateSourceMax?: number
  categoryPositions?: Array<number | null>
  categoryBandwidth?: number
  categories: string[]
  series: ChartSeriesLayout[]
  sankey?: ChartSankeyLayout
  chord?: ChartChordLayout
  relation?: ChartRelationGraphLayout
  wordCloud?: ChartCloudWord[]
  treemap?: ChartTreeRect[]
  sunburst?: ChartSunburstSector[]
  waterfall?: {
    points: ChartWaterfallPointLayout[]
    connectors: Array<{ from: { x: number; y: number }; to: { x: number; y: number } }>
  }
}

const DEFAULT_LEGEND: ChartLegendOptions = { visible: true, position: 'bottom' }
export const defaultChartLegend = (type: ChartType): ChartLegendOptions => ({ ...DEFAULT_LEGEND, visible: !['histogram', 'candlestick', 'treemap', 'sunburst', 'chord', 'wordCloud'].includes(type) })
const DEFAULT_AXIS: ChartAxisOptions = { scaleType: 'linear' }

/** Normalize authored configuration without reading its source. Read-only
 * settings and asynchronous creation share this exact model contract. */
export function createChartDefinition(workbook: WorkbookModel, sheetId: string, config: ChartConfig, id = randomId('chart')): ChartModel {
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
    drawingGeometry: config.drawingGeometry === undefined ? undefined : normalizeDrawingAffineGeometry(config.drawingGeometry),
    title: config.title,
    subtitle: config.subtitle,
    palette: config.palette ? [...config.palette] : undefined,
    options: cloneChartOptions(config.options),
    legend: cloneChartLegend({ ...defaultChartLegend(type), ...(config.legend ?? {}) }),
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
  if (model.pivotTableId && !workbook.pivotTables?.some((pivot) => pivot.id === model.pivotTableId)) throw new Error(`Pivot table not found: ${model.pivotTableId}`)
  return model
}

export function createChart(workbook: WorkbookModel, sheetId: string, config: ChartConfig, id = randomId('chart'), preloaded?: ChartRecords): ChartModel {
  const model = createChartDefinition(workbook, sheetId, config, id)
  const data = readChartData(workbook, model, preloaded)
  model.options = initializeChartWaterfallOptions(model, data.series)
  validateChartSourceData(model, data)
  return model
}

/** Normalize an edit independently of source I/O. A source-aware second pass
 * supplies conversion defaults from the very same records used for projection. */
export function updateChartDefinition(workbook: WorkbookModel, chart: ChartModel, patch: ChartUpdatePatch, records?: ChartRecords): ChartModel {
  const source = patch.source ? normalizeSource(patch.source) : chart.source
  const pivotTableId = Object.prototype.hasOwnProperty.call(patch, 'pivotTableId') ? patch.pivotTableId : chart.pivotTableId
  const sourceChart = { ...chart, source, pivotTableId, options: mergeChartOptions(chart.options, patch.options) }
  validateChartGradientAuthoring(patch.type ?? chart.type, sourceChart.options, chart.options)
  const sheet = requireSheet(workbook, chart.sheetId)
  if (!pivotTableId) validateSourceInSheet(source.range, source.sheetId ? requireSheet(workbook, source.sheetId) : sheet)
  validateChartType(patch.type ?? chart.type)
  const defaults = patch.type !== undefined && patch.type !== chart.type ? convertChartType(sourceChart, patch.type, records) : chart
  const next: ChartModel = {
    ...chart,
    ...defaults,
    ...patch,
    source,
    pivotTableId,
    position: patch.position ? normalizeSize(patch.position, 'chart position', true) : chart.position,
    size: patch.size ? normalizeSize(patch.size, 'chart size', false) : chart.size,
    legend: cloneChartLegend({ ...defaultChartLegend(patch.type ?? chart.type), ...defaults.legend, ...patch.legend }),
    xAxis: cloneChartAxis({ ...DEFAULT_AXIS, ...defaults.xAxis, ...patch.xAxis }),
    yAxis: cloneChartAxis({ ...DEFAULT_AXIS, ...defaults.yAxis, ...patch.yAxis }),
    valueFields: patch.valueFields ? [...patch.valueFields] : defaults.valueFields ? [...defaults.valueFields] : [],
    categoryField: patch.categoryField ?? defaults.categoryField ?? 0,
    seriesToCategory: cloneChartSeriesCategories(patch.seriesToCategory ?? defaults.seriesToCategory ?? {}),
    options: mergeChartOptions(defaults.options, patch.options),
    palette: Object.prototype.hasOwnProperty.call(patch, 'palette') ? patch.palette ? [...patch.palette] : undefined : chart.palette,
    revision: chart.revision + 1,
  }
  if (Object.hasOwn(patch, 'drawingGeometry')) next.drawingGeometry = patch.drawingGeometry === undefined ? undefined : normalizeDrawingAffineGeometry(patch.drawingGeometry)
  else if (chart.drawingGeometry && (patch.position || patch.size)) {
    const before = drawingAffineFrame(chart.drawingGeometry)
    next.drawingGeometry = drawingAffineWithFrame(chart.drawingGeometry, {
      left: patch.position ? next.position.x - GRID_HEADER_WIDTH : before.left,
      top: patch.position ? next.position.y - GRID_HEADER_HEIGHT : before.top,
      width: patch.size?.width ?? before.width, height: patch.size?.height ?? before.height,
    })
  } else if (chart.drawingGeometry) next.drawingGeometry = normalizeDrawingAffineGeometry(chart.drawingGeometry)
  validateChartPalette(next.palette)
  validateChartText(next.title, '标题'); validateChartText(next.subtitle, '副标题')
  next.categoryField = normalizeField(next.categoryField, 'categoryField')
  next.valueFields = next.valueFields.map((field) => normalizeField(field, 'valueField'))
  if (new Set(next.valueFields).size !== next.valueFields.length) throw new RangeError('Chart value fields must be unique')
  validateAxis(next.xAxis, 'xAxis')
  validateAxis(next.yAxis, 'yAxis')
  if (pivotTableId && !workbook.pivotTables?.some(pivot => pivot.id === pivotTableId)) throw new Error(`Pivot table not found: ${pivotTableId}`)
  return next
}

export function updateChart(workbook: WorkbookModel, chart: ChartModel, patch: ChartUpdatePatch, preloaded?: ChartRecords): void {
  const seed = updateChartDefinition(workbook, chart, patch)
  const records = preloaded ?? readChartRecords(workbook, seed)
  const next = updateChartDefinition(workbook, chart, patch, records)
  validateChartSourceData(next, readChartData(workbook, next, records))
  Object.assign(chart, next)
}

/** Read the current source without interpreting the chart's field roles. */
export function readChartRecords(workbook: WorkbookModel, chart: Pick<ChartModel, 'sheetId' | 'source' | 'pivotTableId' | 'options'> & Partial<Pick<ChartModel, 'type'>>): ChartRecords {
  const steps = readChartRecordSteps(workbook, chart)
  while (true) { const step = steps.next(); if (step.done) return step.value }
}

export function readChartData(workbook: WorkbookModel, chart: ChartModel, preloaded?: ChartRecords): ChartData {
  assertChartSourcePresent(workbook, chart)
  const data = projectChartRecords(preloaded ?? readChartRecords(workbook, chart), chart)
  validateChartNamedSeriesBindings(chart.options, data.series)
  return data
}

export function chartLayout(workbook: WorkbookModel, chart: ChartModel, computed?: ChartComputedStatistics): ChartLayout {
  assertChartSourcePresent(workbook, chart)
  return layoutChartData(computed?.preparedData ?? readChartData(workbook, chart), chart, computed)
}

export function chartRenderCommands(workbook: WorkbookModel, chart: ChartModel, computed?: ChartComputedStatistics): DrawCommand[] {
  assertChartSourcePresent(workbook, chart)
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
  return { visible: legend.visible, position: legend.position, ...definedChartStyleFields(chartLegendPatch(undefined, legend)) }
}

function validateChartText(value: string | undefined, name: string): void {
  if (value !== undefined && (typeof value !== 'string' || value.length > 100_000)) throw new TypeError(`图表${name}须为不超过 100000 字符的文本。`)
}

function requireSheet(workbook: WorkbookModel, sheetId: string): SpreadsheetSheet {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
  return sheet
}
