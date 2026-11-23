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
import { pivotOutputValues } from '../pivot/services/pivot-engine.ts'
import { resolveChartType } from './chart-catalog'
import { readSpecialChartData, type ChartCandle, type ChartGraphData, type ChartHierarchyNode } from './chart-data'
import { cloneChartOptions, mergeChartOptions } from './chart-options'
import { chartCompoundPieData, type ChartComputedStatistics } from './chart-derived'
import { layoutChartData } from './chart-layout'
import { drawChartData } from './chart-draw'

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
    legend: { ...DEFAULT_LEGEND, ...(config.legend ?? {}) },
    categoryField,
    valueFields: valueFields ?? [],
    ...(config.seriesToCategory ? { seriesToCategory: cloneSeriesCategories(config.seriesToCategory) } : {}),
    pivotTableId: config.pivotTableId,
    xAxis: { ...DEFAULT_AXIS, ...(config.xAxis ?? {}) },
    yAxis: { ...DEFAULT_AXIS, ...(config.yAxis ?? {}) },
    revision: 0,
  }
  validateAxis(model.xAxis, 'xAxis')
  validateAxis(model.yAxis, 'yAxis')
  validateChartData(model, readChartData(workbook, model), model.yAxis)
  if (model.pivotTableId && !workbook.pivotTables?.some((pivot) => pivot.id === model.pivotTableId)) throw new Error(`Pivot table not found: ${model.pivotTableId}`)
  // Resolve once during creation so invalid field indexes fail atomically.
  readChartData(workbook, model)
  return model
}

export function updateChart(workbook: WorkbookModel, chart: ChartModel, patch: Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>>): void {
  const next: ChartModel = {
    ...chart,
    ...patch,
    source: patch.source ? normalizeSource(patch.source) : chart.source,
    position: patch.position ? normalizeSize(patch.position, 'chart position', true) : chart.position,
    size: patch.size ? normalizeSize(patch.size, 'chart size', false) : chart.size,
    legend: patch.legend ? { ...chart.legend, ...patch.legend } : chart.legend,
    xAxis: patch.xAxis ? { ...chart.xAxis, ...patch.xAxis } : chart.xAxis,
    yAxis: patch.yAxis ? { ...chart.yAxis, ...patch.yAxis } : chart.yAxis,
    valueFields: patch.valueFields ? [...patch.valueFields] : chart.valueFields,
    seriesToCategory: cloneSeriesCategories(patch.seriesToCategory ?? chart.seriesToCategory ?? {}),
    options: mergeChartOptions(chart.options, patch.options),
    palette: patch.palette ? [...patch.palette] : chart.palette,
    revision: chart.revision + 1,
  }
  const sheet = requireSheet(workbook, chart.sheetId)
  if (!next.pivotTableId) validateSourceInSheet(next.source.range, next.source.sheetId ? requireSheet(workbook, next.source.sheetId) : sheet)
  validateChartType(next.type)
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
    const pivot = workbook.pivotTables?.find((candidate) => candidate.id === chart.pivotTableId)
    if (!pivot) throw new Error(`Pivot table not found: ${chart.pivotTableId}`)
    const data = readPivotChartData(pivotOutputValues(workbook, pivot), pivot, chart)
    if (chart.type === 'scatter') {
      if (Object.keys(chart.seriesToCategory ?? {}).length) throw new RangeError('Pivot scatter charts use the shared pivot row field for X values')
      return { ...data, series: data.series.map((series) => ({ ...series, xValues: data.categories.map(numericValue) })) }
    }
    return data
  }
  const sheet = requireSheet(workbook, chart.source.sheetId ?? chart.sheetId)
  const range = chart.source.range
  const matrix = readMatrix(sheet, range)
  const vectors = chart.source.orientation === 'rows' ? matrix[0]?.map((_cell, index) => matrix.map(row => row[index] ?? null)) ?? [] : matrix
  if (vectors.length < 2 || !vectors[0]?.length) throw new RangeError('图表数据区域须包含字段名和至少一条数据。')
  const records = { fields: vectors[0]!.map(scalarLabel), records: vectors.slice(1) }
  const special = readSpecialChartData(records, chart)
  if (special) return special
  if (vectors[0]!.length < 2) throw new RangeError('Chart source must contain a category and at least one data value')
  if (chart.type === 'scatter') return { ...records, ...readScatterChartData(matrix, chart) }
  if (chart.source.orientation === 'rows') return { ...records, ...readRowsChartData(matrix, chart) }
  const categoryOffset = chart.categoryField
  if (categoryOffset >= matrix[0]!.length) throw new RangeError('Chart category field is outside source range')
  const valueFields = chart.valueFields.length ? chart.valueFields : matrix[0]!.map((_value, index) => index).filter((index) => index !== categoryOffset)
  for (const field of valueFields) if (field >= matrix[0]!.length) throw new RangeError('Chart value field is outside source range')
  const categories = matrix.slice(1).map((row) => scalarLabel(row[categoryOffset]))
  const series = valueFields.map((field) => ({
    field,
    name: scalarLabel(matrix[0]![field]) || `Series ${field + 1}`,
    values: matrix.slice(1).map((row) => numericValue(row[field])),
  }))
  return { ...records, categories, series }
}

function readPivotChartData(matrix: CellScalar[][], pivot: NonNullable<WorkbookModel['pivotTables']>[number], chart: ChartModel): ChartData {
  if (matrix.length < 2 || matrix[0]!.length < 2) throw new RangeError('Pivot chart requires a pivot table with row and value fields')
  const rowFields = pivot.fields.filter((field) => field.area === 'row').sort((a, b) => a.position - b.position)
  const columnFields = pivot.fields.filter((field) => field.area === 'column').sort((a, b) => a.position - b.position)
  if (!rowFields.length) throw new RangeError('Pivot chart requires at least one row field to define categories')
  const valueRows = pivot.options.valuePosition === 'row'
  const headerRows = Math.max(1, columnFields.length + (valueRows ? 0 : 1))
  const rowHeaderCount = rowFields.length + (valueRows ? 1 : 0)
  const rows = matrix.slice(headerRows).filter((row) => row.some((value) => typeof value === 'number') && !isPivotTotalRow(row, rowFields.length))
  if (!rows.length) throw new RangeError('Pivot chart has no numeric categories to display')
  const previousLabels = Array.from({ length: rowFields.length }, () => '')
  const categories = rows.map((row) => row.slice(0, rowFields.length).map((value, index) => {
    if (value != null && value !== '') previousLabels[index] = scalarLabel(value)
    return previousLabels[index]!
  }).filter(Boolean).join(' / '))
  if (!valueRows) {
    const valueColumns = matrix[0]!.map((_value, index) => index)
      .filter((index) => index >= rowHeaderCount && rows.some((row) => typeof row[index] === 'number') && !isPivotTotalColumn(matrix, index, headerRows))
    const valueFields = chart.valueFields.length ? chart.valueFields : valueColumns
    for (const field of valueFields) if (!valueColumns.includes(field)) throw new RangeError('Pivot chart value field is outside pivot output')
    return {
      categories,
      series: valueFields.map((field) => ({
        field,
        name: matrix.slice(0, headerRows).map((row) => row[field]).filter((value) => value != null && value !== '').map(scalarLabel).join(' · ') || `Series ${field}`,
        values: rows.map((row) => numericValue(row[field])),
      })),
    }
  }

  const seriesByName = new Map<string, { name: string; values: Map<string, number> }>()
  for (const row of rows) {
    const category = row.slice(0, rowFields.length).map(scalarLabel).filter(Boolean).join(' / ')
    const valueName = scalarLabel(row[rowFields.length])
    for (let column = rowHeaderCount; column < matrix[0]!.length; column += 1) {
      if (isPivotTotalColumn(matrix, column, headerRows) || (columnFields.length > 0 && pivot.options.showColGrandTotal && column === matrix[0]!.length - 1)) continue
      const value = numericValue(row[column])
      if (value === null) continue
      const columnName = matrix.slice(0, headerRows).map((header) => header[column]).filter((label) => label != null && label !== '').map(scalarLabel).join(' · ')
      const name = [columnName, valueName].filter(Boolean).join(' · ')
      const series = seriesByName.get(name) ?? { name, values: new Map<string, number>() }
      series.values.set(category, value)
      seriesByName.set(name, series)
    }
  }
  const uniqueCategories = [...new Set(rows.map((row) => row.slice(0, rowFields.length).map(scalarLabel).filter(Boolean).join(' / ')))]
  return { categories: uniqueCategories, series: [...seriesByName.values()].map((series, field) => ({ field, name: series.name, values: uniqueCategories.map((category) => series.values.get(category) ?? null) })) }
}

function isPivotTotalColumn(matrix: CellScalar[][], column: number, headerRows: number): boolean {
  return matrix.slice(0, headerRows).some((row) => typeof row[column] === 'string' && /^grand\s+total(?:\s+.+)?$/i.test(row[column]!.trim()))
}

function isPivotTotalRow(row: CellScalar[], rowFieldCount: number): boolean {
  return row.slice(0, rowFieldCount).some((value) => typeof value === 'string' && /^grand\s+total$/i.test(value.trim()))
}

function cloneSeriesCategories(mapping: Record<number, number[]>): Record<number, number[]> {
  return Object.fromEntries(Object.entries(mapping).map(([field, categories]) => {
    normalizeField(Number(field), 'seriesToCategory value field')
    if (!Array.isArray(categories) || categories.length !== 1) throw new RangeError('Scatter series require exactly one X field')
    return [field, categories.map((category) => normalizeField(category, 'seriesToCategory X field'))]
  }))
}

function readScatterChartData(matrix: CellScalar[][], chart: ChartModel): ChartData {
  const vectors = chart.source.orientation === 'rows'
    ? matrix
    : matrix[0]!.map((_value, column) => matrix.map((row) => row[column] ?? null))
  const categoryField = normalizeField(chart.categoryField, 'categoryField')
  const categoryVector = vectors[categoryField]
  if (!categoryVector) throw new RangeError('Chart X field is outside source range')
  const mapping = cloneSeriesCategories(chart.seriesToCategory ?? {})
  const xFields = new Set([categoryField, ...Object.values(mapping).flat()])
  const fields = chart.valueFields.length ? chart.valueFields : vectors.map((_vector, field) => field).filter((field) => !xFields.has(field))
  for (const [valueField, xField] of Object.entries(mapping)) {
    if (!fields.includes(Number(valueField)) || !vectors[xField[0]!]) throw new RangeError('Scatter series mapping is outside selected source fields')
  }
  return {
    categories: categoryVector.slice(1).map(scalarLabel),
    series: fields.map((field) => {
      const vector = vectors[field]
      if (!vector) throw new RangeError('Chart Y field is outside source range')
      const xField = mapping[field]?.[0] ?? categoryField
      return {
        field,
        name: scalarLabel(vector[0]) || `Series ${field + 1}`,
        values: vector.slice(1).map(numericValue),
        xValues: vectors[xField]!.slice(1).map(numericValue),
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

function readRowsChartData(matrix: CellScalar[][], chart: ChartModel): ChartData {
  const categoryOffset = chart.categoryField
  if (categoryOffset >= matrix.length) throw new RangeError('Chart category field is outside source rows')
  const categoryRow = matrix[categoryOffset]!
  const categories = categoryRow.slice(1).map(scalarLabel)
  const fields = chart.valueFields.length ? chart.valueFields : matrix.map((_row, index) => index).filter((index) => index !== categoryOffset)
  for (const field of fields) if (field <= 0 || field >= matrix.length) throw new RangeError('Chart value field is outside source rows')
  return { categories, series: fields.map((field) => ({ field, name: scalarLabel(matrix[field]![0]) || `Series ${field + 1}`, values: categories.map((_category, index) => numericValue(matrix[field]![index + 1])) })) }
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
  for (const [name, value] of Object.entries(range)) if (!Number.isInteger(value) || value < 0) throw new RangeError(`Invalid chart range ${name}`)
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
  if (!Number.isInteger(value) || value < 0) throw new RangeError(`Invalid chart ${name}`)
  return value
}

function validateAxis(axis: ChartAxisOptions, name: string): void {
  if (axis.scaleType !== 'linear' && axis.scaleType !== 'log') throw new TypeError(`Invalid ${name} scale type`)
  if (axis.logBase !== undefined && (!Number.isFinite(axis.logBase) || axis.logBase < 2 || axis.logBase > 1000)) throw new RangeError(`${name} log base must be between 2 and 1000`)
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
