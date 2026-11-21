import type {
  CellRange,
  CellScalar,
  ChartAxisOptions,
  ChartLegendOptions,
  ChartModel,
  ChartSource,
  ChartType,
  SpreadsheetSheet,
  WorkbookModel,
} from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { randomId } from '../../api/model-utils'
import { pivotOutputValues } from '../pivot/services/pivot-engine.ts'

export interface ChartConfig {
  type: ChartType
  source: ChartSource
  position: { x: number; y: number }
  size: { width: number; height: number }
  title?: string
  legend?: Partial<ChartLegendOptions>
  categoryField?: number
  valueFields?: number[]
  pivotTableId?: string
  xAxis?: Partial<ChartAxisOptions>
  yAxis?: Partial<ChartAxisOptions>
}

export interface ChartSeriesData {
  field: number
  name: string
  values: Array<number | null>
}

export interface ChartData {
  categories: string[]
  series: ChartSeriesData[]
}

export interface ChartPoint { x: number; y: number; value: number; category: string; categoryIndex: number }
export interface ChartSeriesLayout {
  name: string
  color: string
  points: ChartPoint[]
}
export interface ChartLayout {
  frame: RenderRect
  plot: RenderRect
  min: number
  max: number
  categories: string[]
  series: ChartSeriesLayout[]
}

const CHART_COLORS = ['#2563eb', '#dc2626', '#059669', '#d97706', '#7c3aed', '#0891b2']
const DEFAULT_LEGEND: ChartLegendOptions = { visible: true, position: 'bottom' }
const DEFAULT_AXIS: ChartAxisOptions = { scaleType: 'linear' }

export function createChart(workbook: WorkbookModel, sheetId: string, config: ChartConfig, id = randomId('chart')): ChartModel {
  const sheet = requireSheet(workbook, sheetId)
  const source = normalizeSource(config.source)
  if (!config.pivotTableId) validateSourceInSheet(source.range, sheet)
  const type = config.type
  if (type !== 'bar' && type !== 'line') throw new TypeError(`Unsupported chart type: ${type}`)
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
    legend: { ...DEFAULT_LEGEND, ...(config.legend ?? {}) },
    categoryField,
    valueFields: valueFields ?? [],
    pivotTableId: config.pivotTableId,
    xAxis: { ...DEFAULT_AXIS, ...(config.xAxis ?? {}) },
    yAxis: { ...DEFAULT_AXIS, ...(config.yAxis ?? {}) },
    revision: 0,
  }
  validateAxis(model.xAxis, 'xAxis')
  validateAxis(model.yAxis, 'yAxis')
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
    revision: chart.revision + 1,
  }
  const sheet = requireSheet(workbook, chart.sheetId)
  if (!next.pivotTableId) validateSourceInSheet(next.source.range, sheet)
  if (next.type !== 'bar' && next.type !== 'line') throw new TypeError(`Unsupported chart type: ${next.type}`)
  next.categoryField = normalizeField(next.categoryField, 'categoryField')
  next.valueFields = next.valueFields.map((field) => normalizeField(field, 'valueField'))
  if (new Set(next.valueFields).size !== next.valueFields.length) throw new RangeError('Chart value fields must be unique')
  validateAxis(next.xAxis, 'xAxis')
  validateAxis(next.yAxis, 'yAxis')
  readChartData(workbook, next)
  Object.assign(chart, next)
}

export function readChartData(workbook: WorkbookModel, chart: ChartModel): ChartData {
  if (chart.pivotTableId) {
    const pivot = workbook.pivotTables?.find((candidate) => candidate.id === chart.pivotTableId)
    if (!pivot) throw new Error(`Pivot table not found: ${chart.pivotTableId}`)
    return readPivotChartData(pivotOutputValues(workbook, pivot), chart)
  }
  const sheet = requireSheet(workbook, chart.sheetId)
  const range = chart.source.range
  const matrix = readMatrix(sheet, range)
  if (matrix.length < 2 || matrix[0]!.length < 2) throw new RangeError('Chart source must contain a header and at least one data value')
  if (chart.source.orientation === 'rows') return readRowsChartData(matrix, chart)
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
  return { categories, series }
}

function readPivotChartData(matrix: CellScalar[][], chart: ChartModel): ChartData {
  if (matrix.length < 2 || matrix[0]!.length < 2) throw new RangeError('Pivot chart requires a pivot table with row and value fields')
  const rows = matrix.slice(2).filter((row) => row.some((value) => typeof value === 'number') && !isPivotTotalRow(row))
  if (!rows.length) throw new RangeError('Pivot chart has no numeric categories to display')
  const firstValueField = matrix[0]!.findIndex((_value, index) => index > 0 && rows.some((row) => typeof row[index] === 'number'))
  const valueFields = chart.valueFields.length
    ? chart.valueFields
    : matrix[0]!.map((_value, index) => index).filter((index) => index >= firstValueField && rows.some((row) => typeof row[index] === 'number') && !isPivotTotalColumn(matrix, index))
  for (const field of valueFields) if (field <= 0 || field >= matrix[0]!.length || isPivotTotalColumn(matrix, field)) throw new RangeError('Pivot chart value field is outside pivot output')
  const previousLabels = Array.from({ length: firstValueField }, () => '')
  const categories = rows.map((row) => row.slice(0, firstValueField).map((value, index) => {
    if (value != null && value !== '') previousLabels[index] = scalarLabel(value)
    return previousLabels[index]!
  }).filter(Boolean).join(' / '))
  return {
    categories,
    series: valueFields.map((field) => ({
      field,
      name: [matrix[0]![field], matrix[1]![field]].filter((value) => value != null && value !== '').map(scalarLabel).join(' · ') || `Series ${field}`,
      values: rows.map((row) => numericValue(row[field])),
    })),
  }
}

function isPivotTotalColumn(matrix: CellScalar[][], column: number): boolean {
  return matrix.slice(0, 2).some((row) => typeof row[column] === 'string' && /(?:^|\s)(?:grand\s+)?total$/i.test(row[column]!.trim()))
}

function isPivotTotalRow(row: CellScalar[]): boolean {
  return row.some((value) => typeof value === 'string' && /(?:^|\s)(?:grand\s+)?total$/i.test(value.trim()))
}

export function chartLayout(workbook: WorkbookModel, chart: ChartModel): ChartLayout {
  const data = readChartData(workbook, chart)
  const frame = { x: chart.position.x, y: chart.position.y, width: chart.size.width, height: chart.size.height }
  const top = chart.title ? 32 : 12
  const legendSize = chart.legend.visible ? 24 : 0
  const left = 46, right = chart.yAxis.title ? 18 : 10, bottom = 34 + legendSize
  const plot: RenderRect = {
    x: frame.x + left,
    y: frame.y + top,
    width: Math.max(1, frame.width - left - right),
    height: Math.max(1, frame.height - top - bottom),
  }
  const values = data.series.flatMap((series) => series.values.filter((value): value is number => value !== null && Number.isFinite(value) && (chart.yAxis.scaleType !== 'log' || value > 0)))
  if (!values.length) return { frame, plot, min: 0, max: 1, categories: data.categories, series: data.series.map((series, index) => ({ name: series.name, color: CHART_COLORS[index % CHART_COLORS.length]!, points: [] })) }
  const rawMin = Math.min(...values), rawMax = Math.max(...values)
  const min = chart.yAxis.scaleType === 'log' ? Math.max(Number.MIN_VALUE, rawMin) : Math.min(0, rawMin)
  const max = chart.yAxis.scaleType === 'log' ? rawMax : rawMax === min ? min + 1 : rawMax
  const toY = (value: number): number => {
    if (chart.yAxis.scaleType === 'log') {
      const base = chart.yAxis.logBase ?? 10
      return plot.y + plot.height - ((Math.log(value) / Math.log(base) - Math.log(min) / Math.log(base)) / (Math.log(max) / Math.log(base) - Math.log(min) / Math.log(base))) * plot.height
    }
    return plot.y + plot.height - ((value - min) / (max - min)) * plot.height
  }
  const series = data.series.map((sourceSeries, seriesIndex) => {
    const points = sourceSeries.values.map((value, index) => value == null || !Number.isFinite(value) || (chart.yAxis.scaleType === 'log' && value <= 0)
      ? undefined
      : { x: plot.x + (data.categories.length <= 1 ? plot.width / 2 : index * plot.width / (data.categories.length - 1)), y: toY(value), value, category: data.categories[index] ?? '', categoryIndex: index }).filter((point): point is ChartPoint => point !== undefined)
    return { name: sourceSeries.name, color: CHART_COLORS[seriesIndex % CHART_COLORS.length]!, points }
  })
  return { frame, plot, min, max, categories: data.categories, series }
}

export function chartRenderCommands(workbook: WorkbookModel, chart: ChartModel): DrawCommand[] {
  const layout = chartLayout(workbook, chart)
  const commands: DrawCommand[] = [
    { kind: 'rect', ...layout.frame, fill: '#ffffff', stroke: '#cbd5e1', strokeWidth: 1 },
    { kind: 'rect', ...layout.plot, fill: '#ffffff', stroke: '#e2e8f0', strokeWidth: 1 },
  ]
  if (chart.title) commands.push({ kind: 'text', x: layout.frame.x + layout.frame.width / 2, y: layout.frame.y + 20, text: chart.title, font: '600 14px Aptos, Arial, sans-serif', color: '#172033', align: 'center', baseline: 'middle' })
  const zero = chart.yAxis.scaleType === 'linear' && layout.min < 0 && layout.max > 0
    ? layout.plot.y + layout.plot.height - ((0 - layout.min) / (layout.max - layout.min)) * layout.plot.height
    : layout.plot.y + layout.plot.height
  commands.push({ kind: 'line', x1: layout.plot.x, y1: zero, x2: layout.plot.x + layout.plot.width, y2: zero, color: '#64748b', width: 1 })
  commands.push({ kind: 'line', x1: layout.plot.x, y1: layout.plot.y, x2: layout.plot.x, y2: layout.plot.y + layout.plot.height, color: '#64748b', width: 1 })
  const categoryStep = layout.categories.length ? layout.plot.width / Math.max(1, layout.categories.length) : layout.plot.width
  if (chart.type === 'bar') {
    const groupWidth = categoryStep * 0.72
    const barWidth = groupWidth / Math.max(1, layout.series.length)
    layout.series.forEach((series, seriesIndex) => series.points.forEach((point) => {
      const categoryIndex = point.categoryIndex
      const x = layout.plot.x + categoryIndex * categoryStep + (categoryStep - groupWidth) / 2 + seriesIndex * barWidth
      const y = Math.min(point.y, zero), height = Math.max(1, Math.abs(zero - point.y))
      commands.push({ kind: 'rect', x, y, width: Math.max(1, barWidth - 1), height, fill: series.color, stroke: '#ffffff', strokeWidth: 0.5 })
    }))
  } else {
    layout.series.forEach((series) => {
      for (let index = 1; index < series.points.length; index++) {
        const previous = series.points[index - 1]!, point = series.points[index]!
        commands.push({ kind: 'line', x1: previous.x, y1: previous.y, x2: point.x, y2: point.y, color: series.color, width: 2 })
      }
      series.points.forEach((point) => commands.push({ kind: 'rect', x: point.x - 3, y: point.y - 3, width: 6, height: 6, fill: series.color }))
    })
  }
  if (chart.legend.visible) {
    const y = chart.legend.position === 'top' ? layout.frame.y + 12 : layout.frame.y + layout.frame.height - 12
    layout.series.forEach((series, index) => {
      const x = layout.frame.x + 12 + index * Math.max(72, layout.frame.width / Math.max(1, layout.series.length + 1))
      commands.push({ kind: 'rect', x, y: y - 5, width: 10, height: 10, fill: series.color })
      commands.push({ kind: 'text', x: x + 15, y, text: series.name, font: '11px Aptos, Arial, sans-serif', color: '#475569', baseline: 'middle' })
    })
  }
  return commands
}

export function normalizeSource(source: ChartSource): ChartSource {
  if (!source || !source.range) throw new TypeError('Chart source is required')
  const range = normalizeRange(source.range)
  if (source.orientation !== 'columns' && source.orientation !== 'rows') throw new TypeError(`Invalid chart source orientation: ${source.orientation}`)
  return { range, orientation: source.orientation }
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
