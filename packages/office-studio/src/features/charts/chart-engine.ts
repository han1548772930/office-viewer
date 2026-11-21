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
}

export interface ChartData {
  categories: string[]
  series: ChartSeriesData[]
}

export interface ChartPoint { x: number; y: number; value: number; category: string; categoryIndex: number; baseY?: number }
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
  xMin?: number
  xMax?: number
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
  validateChartData(model.type, readChartData(workbook, model), model.yAxis)
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
    revision: chart.revision + 1,
  }
  const sheet = requireSheet(workbook, chart.sheetId)
  if (!next.pivotTableId) validateSourceInSheet(next.source.range, sheet)
  validateChartType(next.type)
  next.categoryField = normalizeField(next.categoryField, 'categoryField')
  next.valueFields = next.valueFields.map((field) => normalizeField(field, 'valueField'))
  if (new Set(next.valueFields).size !== next.valueFields.length) throw new RangeError('Chart value fields must be unique')
  validateAxis(next.xAxis, 'xAxis')
  validateAxis(next.yAxis, 'yAxis')
  validateChartData(next.type, readChartData(workbook, next), next.yAxis)
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
  const sheet = requireSheet(workbook, chart.sheetId)
  const range = chart.source.range
  const matrix = readMatrix(sheet, range)
  if (matrix.length < 2 || matrix[0]!.length < 2) throw new RangeError('Chart source must contain a header and at least one data value')
  if (chart.type === 'scatter') return readScatterChartData(matrix, chart)
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

function axisExtent(values: number[], axis: ChartAxisOptions): { min: number; max: number } {
  if (!values.length) return { min: axis.scaleType === 'log' ? 1 : 0, max: axis.scaleType === 'log' ? axis.logBase ?? 10 : 1 }
  let min = values[0]!, max = values[0]!
  for (const value of values) { min = Math.min(min, value); max = Math.max(max, value) }
  if (axis.scaleType !== 'log') min = Math.min(0, min)
  if (min === max) max = axis.scaleType === 'log' ? min * (axis.logBase ?? 10) : min + 1
  return { min, max }
}

function axisRatio(value: number, min: number, max: number, axis: ChartAxisOptions): number {
  return axis.scaleType === 'log'
    ? (Math.log(value) - Math.log(min)) / (Math.log(max) - Math.log(min))
    : (value - min) / (max - min)
}

function scatterLayout(data: ChartData, chart: ChartModel, frame: RenderRect, plot: RenderRect): ChartLayout {
  const pairs = data.series.map((series) => series.values.flatMap((value, categoryIndex) => {
    const xValue = series.xValues?.[categoryIndex]
    if (value == null || xValue == null || !Number.isFinite(value) || !Number.isFinite(xValue)) return []
    if (chart.xAxis.scaleType === 'log' && xValue <= 0 || chart.yAxis.scaleType === 'log' && value <= 0) return []
    return [{ xValue, value, categoryIndex }]
  }))
  const xExtent = axisExtent(pairs.flatMap((points) => points.map((point) => point.xValue)), chart.xAxis)
  const yExtent = axisExtent(pairs.flatMap((points) => points.map((point) => point.value)), chart.yAxis)
  return {
    frame, plot, ...yExtent, xMin: xExtent.min, xMax: xExtent.max, categories: data.categories,
    series: data.series.map((series, index) => ({
      name: series.name, color: CHART_COLORS[index % CHART_COLORS.length]!,
      points: pairs[index]!.map(({ xValue, value, categoryIndex }) => ({
        x: plot.x + axisRatio(xValue, xExtent.min, xExtent.max, chart.xAxis) * plot.width,
        y: plot.y + plot.height - axisRatio(value, yExtent.min, yExtent.max, chart.yAxis) * plot.height,
        value, category: scalarLabel(xValue), categoryIndex,
      })),
    })),
  }
}

export function chartLayout(workbook: WorkbookModel, chart: ChartModel): ChartLayout {
  const data = readChartData(workbook, chart)
  const frame = { x: chart.position.x, y: chart.position.y, width: chart.size.width, height: chart.size.height }
  const top = (chart.title ? 32 : 12) + (chart.legend.visible && chart.legend.position === 'top' ? 24 : 0)
  const bottomLegendSize = chart.legend.visible && chart.legend.position === 'bottom' ? 24 : 0
  const axisTitleSize = chart.xAxis.title ? 16 : 0
  const left = 54 + (chart.yAxis.title ? 18 : 0)
  const right = 10
  const bottom = 40 + bottomLegendSize + axisTitleSize
  const plot: RenderRect = {
    x: frame.x + left,
    y: frame.y + top,
    width: Math.max(1, frame.width - left - right),
    height: Math.max(1, frame.height - top - bottom),
  }
  if (chart.type === 'scatter') return scatterLayout(data, chart, frame, plot)
  const values = data.series.flatMap((series) => series.values.filter((value): value is number => value !== null && Number.isFinite(value) && (chart.yAxis.scaleType !== 'log' || value > 0)))
  const stackTotals = chart.type === 'stacked-column' ? stackedExtents(data) : undefined
  const rangeValues = stackTotals ? [...stackTotals] : values
  if (!rangeValues.length) {
    const min = chart.yAxis.scaleType === 'log' ? 1 : 0
    const max = chart.yAxis.scaleType === 'log' ? chart.yAxis.logBase ?? 10 : 1
    return { frame, plot, min, max, categories: data.categories, series: data.series.map((series, index) => ({ name: series.name, color: CHART_COLORS[index % CHART_COLORS.length]!, points: [] })) }
  }
  const rawMin = Math.min(...rangeValues), rawMax = Math.max(...rangeValues)
  const min = chart.yAxis.scaleType === 'log' ? Math.max(Number.MIN_VALUE, rawMin) : Math.min(0, rawMin)
  const max = chart.yAxis.scaleType === 'log'
    ? rawMax === min ? min * (chart.yAxis.logBase ?? 10) : rawMax
    : rawMax === min ? min + 1 : rawMax
  const toY = (value: number): number => {
    if (chart.yAxis.scaleType === 'log') {
      const base = chart.yAxis.logBase ?? 10
      return plot.y + plot.height - ((Math.log(value) / Math.log(base) - Math.log(min) / Math.log(base)) / (Math.log(max) / Math.log(base) - Math.log(min) / Math.log(base))) * plot.height
    }
    return plot.y + plot.height - ((value - min) / (max - min)) * plot.height
  }
  const positiveStacks = Array.from({ length: data.categories.length }, () => 0)
  const negativeStacks = Array.from({ length: data.categories.length }, () => 0)
  const series = data.series.map((sourceSeries, seriesIndex) => {
    const points = sourceSeries.values.map((value, index) => {
      if (value == null || !Number.isFinite(value) || (chart.yAxis.scaleType === 'log' && value <= 0)) return undefined
      const baseValue = chart.type === 'stacked-column'
        ? value >= 0 ? positiveStacks[index]! : negativeStacks[index]!
        : 0
      if (chart.type === 'stacked-column') {
        if (value >= 0) positiveStacks[index] = baseValue + value
        else negativeStacks[index] = baseValue + value
      }
      return {
        x: plot.x + (index + 0.5) * plot.width / Math.max(1, data.categories.length),
        y: toY(baseValue + value), value, category: data.categories[index] ?? '', categoryIndex: index,
        ...(chart.type === 'stacked-column' ? { baseY: toY(baseValue) } : {}),
      }
    }).filter((point): point is ChartPoint => point !== undefined)
    return { name: sourceSeries.name, color: CHART_COLORS[seriesIndex % CHART_COLORS.length]!, points }
  })
  return { frame, plot, min, max, categories: data.categories, series }
}

export function chartRenderCommands(workbook: WorkbookModel, chart: ChartModel): DrawCommand[] {
  const layout = chartLayout(workbook, chart)
  const commands: DrawCommand[] = [
    { kind: 'rect', ...layout.frame, fill: '#ffffff', stroke: '#cbd5e1', strokeWidth: 1 },
  ]
  if (chart.title) commands.push({ kind: 'text', x: layout.frame.x + layout.frame.width / 2, y: layout.frame.y + 20, text: chart.title, font: '600 14px Aptos, Arial, sans-serif', color: '#172033', align: 'center', baseline: 'middle' })
  if (chart.type === 'pie' || chart.type === 'doughnut') {
    renderPieChart(commands, layout, readChartData(workbook, chart), chart)
    return commands
  }
  commands.push({ kind: 'rect', ...layout.plot, fill: '#ffffff', stroke: '#e2e8f0', strokeWidth: 1 })
  const tickValues = axisTickValues(layout.min, layout.max, chart.yAxis.scaleType)
  tickValues.forEach((value) => {
    const ratio = chart.yAxis.scaleType === 'log'
      ? (Math.log(value) - Math.log(layout.min)) / (Math.log(layout.max) - Math.log(layout.min))
      : (value - layout.min) / (layout.max - layout.min)
    const y = layout.plot.y + layout.plot.height - ratio * layout.plot.height
    commands.push({ kind: 'line', x1: layout.plot.x, y1: y, x2: layout.plot.x + layout.plot.width, y2: y, color: '#e5e7eb', width: 1 })
    commands.push({ kind: 'text', x: layout.plot.x - 8, y, text: formatAxisValue(value), font: '10px Aptos, Arial, sans-serif', color: '#64748b', align: 'right', baseline: 'middle' })
  })
  const zero = chart.yAxis.scaleType === 'linear' && layout.min < 0 && layout.max > 0
    ? layout.plot.y + layout.plot.height - ((0 - layout.min) / (layout.max - layout.min)) * layout.plot.height
    : layout.plot.y + layout.plot.height
  commands.push({ kind: 'line', x1: layout.plot.x, y1: zero, x2: layout.plot.x + layout.plot.width, y2: zero, color: '#64748b', width: 1.25 })
  commands.push({ kind: 'line', x1: layout.plot.x, y1: layout.plot.y, x2: layout.plot.x, y2: layout.plot.y + layout.plot.height, color: '#64748b', width: 1 })
  const categoryStep = layout.categories.length ? layout.plot.width / Math.max(1, layout.categories.length) : layout.plot.width
  if (chart.type === 'scatter') {
    const min = layout.xMin!, max = layout.xMax!
    axisTickValues(min, max, chart.xAxis.scaleType).forEach((value) => {
      const x = layout.plot.x + axisRatio(value, min, max, chart.xAxis) * layout.plot.width
      commands.push({ kind: 'line', x1: x, y1: layout.plot.y, x2: x, y2: layout.plot.y + layout.plot.height, color: '#e5e7eb', width: 1 })
      commands.push({ kind: 'text', x, y: layout.plot.y + layout.plot.height + 14, text: formatAxisValue(value), font: '10px Aptos, Arial, sans-serif', color: '#64748b', align: 'center', baseline: 'middle' })
    })
  } else layout.categories.forEach((category, index) => commands.push({
    kind: 'text',
    x: layout.plot.x + (index + 0.5) * categoryStep,
    y: layout.plot.y + layout.plot.height + 14,
    text: category,
    font: '10px Aptos, Arial, sans-serif',
    color: '#64748b',
    align: 'center',
    baseline: 'middle',
    maxWidth: Math.max(1, categoryStep - 4),
  }))
  if (chart.xAxis.title) commands.push({ kind: 'text', x: layout.plot.x + layout.plot.width / 2, y: layout.plot.y + layout.plot.height + 34, text: chart.xAxis.title, font: '11px Aptos, Arial, sans-serif', color: '#475569', align: 'center', baseline: 'middle' })
  if (chart.yAxis.title) commands.push({ kind: 'text', x: layout.frame.x + 13, y: layout.plot.y + layout.plot.height / 2, text: chart.yAxis.title, font: '11px Aptos, Arial, sans-serif', color: '#475569', align: 'center', baseline: 'middle', rotation: -90 })
  if (chart.type === 'bar' || chart.type === 'stacked-column') {
    const groupWidth = categoryStep * 0.72
    if (chart.type === 'stacked-column') {
      const barWidth = groupWidth * 0.8
      layout.series.forEach((series) => series.points.forEach((point) => {
        const x = layout.plot.x + point.categoryIndex * categoryStep + (categoryStep - barWidth) / 2
        const y = Math.min(point.y, point.baseY ?? zero), height = Math.max(1, Math.abs((point.baseY ?? zero) - point.y))
        commands.push({ kind: 'rect', x, y, width: barWidth, height, fill: series.color, stroke: '#ffffff', strokeWidth: 0.5 })
      }))
    } else {
      const barWidth = groupWidth / Math.max(1, layout.series.length)
      layout.series.forEach((series, seriesIndex) => series.points.forEach((point) => {
        const categoryIndex = point.categoryIndex
        const x = layout.plot.x + categoryIndex * categoryStep + (categoryStep - groupWidth) / 2 + seriesIndex * barWidth
        const y = Math.min(point.y, zero), height = Math.max(1, Math.abs(zero - point.y))
        commands.push({ kind: 'rect', x, y, width: Math.max(1, barWidth - 1), height, fill: series.color, stroke: '#ffffff', strokeWidth: 0.5 })
      }))
    }
  } else if (chart.type === 'area') {
    layout.series.forEach((series) => {
      for (let index = 1; index < series.points.length; index += 1) {
        const previous = series.points[index - 1]!, point = series.points[index]!
        if (point.categoryIndex !== previous.categoryIndex + 1) continue
        for (const segment of areaSegments(previous, point, zero)) commands.push({ kind: 'polygon', points: segment, fill: `rgba(${hexRgb(series.color)}, 0.22)` })
        commands.push({ kind: 'line', x1: previous.x, y1: previous.y, x2: point.x, y2: point.y, color: series.color, width: 2 })
      }
      series.points.forEach((point) => commands.push({ kind: 'rect', x: point.x - 2, y: point.y - 2, width: 4, height: 4, fill: series.color }))
    })
  } else if (chart.type === 'scatter') {
    layout.series.forEach((series) => series.points.forEach((point) => {
      commands.push({ kind: 'rect', x: point.x - 3, y: point.y - 3, width: 6, height: 6, fill: series.color })
    }))
  } else {
    layout.series.forEach((series) => {
      for (let index = 1; index < series.points.length; index++) {
        const previous = series.points[index - 1]!, point = series.points[index]!
        if (point.categoryIndex !== previous.categoryIndex + 1) continue
        commands.push({ kind: 'line', x1: previous.x, y1: previous.y, x2: point.x, y2: point.y, color: series.color, width: 2 })
      }
      series.points.forEach((point) => commands.push({ kind: 'rect', x: point.x - 3, y: point.y - 3, width: 6, height: 6, fill: series.color }))
    })
  }
  if (chart.legend.visible) {
    const y = chart.legend.position === 'top'
      ? layout.frame.y + (chart.title ? 44 : 12)
      : layout.frame.y + layout.frame.height - 12
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

function validateChartType(type: ChartType): void {
  if (!['bar', 'line', 'area', 'pie', 'doughnut', 'stacked-column', 'scatter'].includes(type)) throw new TypeError(`Unsupported chart type: ${type}`)
}

function validateChartData(type: ChartType, data: ChartData, axis: ChartAxisOptions): void {
  if (!data.categories.length || !data.series.length) throw new RangeError('Chart source must contain at least one category and value series')
  if (data.series.some((series) => !series.values.length || series.values.length !== data.categories.length)) throw new RangeError('Chart series must have one value per category')
  if (type === 'scatter' && data.series.some((series) => !series.xValues || series.xValues.length !== series.values.length)) throw new RangeError('Scatter series must have one X value per Y value')
  if (type === 'scatter' && !data.series.some((series) => series.values.some((value, index) => value !== null && series.xValues?.[index] != null))) throw new RangeError('Scatter chart requires at least one numeric X/Y pair')
  if (type === 'stacked-column' && axis.scaleType !== 'linear') throw new TypeError('Stacked column charts require a linear value axis')
  if (type === 'pie' || type === 'doughnut') {
    if (data.series.length !== 1) throw new RangeError(`${type} charts require exactly one value series`)
    const values = data.series[0]!.values.filter((value): value is number => value !== null && Number.isFinite(value))
    if (!values.length || values.some((value) => value < 0) || values.every((value) => value === 0)) throw new RangeError(`${type} charts require non-negative values with a positive total`)
  }
}

function axisTickValues(min: number, max: number, scaleType: ChartAxisOptions['scaleType']): number[] {
  if (min === max) return [min]
  return Array.from({ length: 5 }, (_unused, index) => {
    const ratio = index / 4
    return scaleType === 'log'
      ? min * (max / min) ** ratio
      : min + (max - min) * ratio
  })
}

function formatAxisValue(value: number): string {
  return String(Number(value.toPrecision(4)))
}

function stackedExtents(data: ChartData): number[] {
  const positive = Array.from({ length: data.categories.length }, () => 0)
  const negative = Array.from({ length: data.categories.length }, () => 0)
  for (const series of data.series) {
    series.values.forEach((value, index) => {
      if (value === null || !Number.isFinite(value)) return
      if (value >= 0) positive[index] += value
      else negative[index] += value
    })
  }
  return [...positive, ...negative]
}

function renderPieChart(commands: DrawCommand[], layout: ChartLayout, data: ChartData, chart: ChartModel): void {
  const series = data.series[0]
  if (!series) return
  const values = series.values.map((value) => Math.max(0, value ?? 0))
  const total = values.reduce((sum, value) => sum + value, 0)
  if (total <= 0) return
  const centerX = layout.plot.x + layout.plot.width / 2
  const centerY = layout.plot.y + layout.plot.height / 2
  const radius = Math.max(1, Math.min(layout.plot.width, layout.plot.height) * 0.43)
  const colors = data.categories.map((_category, index) => CHART_COLORS[index % CHART_COLORS.length]!)
  let start = -Math.PI / 2
  values.forEach((value, index) => {
    if (value <= 0) return
    const end = start + Math.PI * 2 * value / total
    const steps = Math.max(2, Math.ceil((end - start) / (Math.PI / 24)))
    let previous = { x: centerX + Math.cos(start) * radius, y: centerY + Math.sin(start) * radius }
    for (let step = 1; step <= steps; step += 1) {
      const angle = start + (end - start) * step / steps
      const point = { x: centerX + Math.cos(angle) * radius, y: centerY + Math.sin(angle) * radius }
      commands.push({ kind: 'polygon', points: [{ x: centerX, y: centerY }, previous, point], fill: colors[index]! })
      previous = point
    }
    start = end
  })
  if (chart.type === 'doughnut') {
    const innerRadius = radius * 0.55
    for (let step = 0; step < 96; step += 1) {
      const startAngle = step * Math.PI * 2 / 96
      const endAngle = (step + 1) * Math.PI * 2 / 96
      commands.push({ kind: 'polygon', points: [
        { x: centerX, y: centerY },
        { x: centerX + Math.cos(startAngle) * innerRadius, y: centerY + Math.sin(startAngle) * innerRadius },
        { x: centerX + Math.cos(endAngle) * innerRadius, y: centerY + Math.sin(endAngle) * innerRadius },
      ], fill: '#ffffff' })
    }
  }
  if (chart.legend.visible) {
    const y = chart.legend.position === 'top'
      ? layout.frame.y + (chart.title ? 44 : 12)
      : layout.frame.y + layout.frame.height - 12
    data.categories.forEach((category, index) => {
      const x = layout.frame.x + 12 + index * Math.max(72, layout.frame.width / Math.max(1, data.categories.length + 1))
      commands.push({ kind: 'rect', x, y: y - 5, width: 10, height: 10, fill: colors[index]! })
      commands.push({ kind: 'text', x: x + 15, y, text: category, font: '11px Aptos, Arial, sans-serif', color: '#475569', baseline: 'middle' })
    })
  }
}

function areaSegments(previous: ChartPoint, point: ChartPoint, zero: number): Array<Array<{ x: number; y: number }>> {
  if ((previous.y - zero) * (point.y - zero) >= 0) return [[
    { x: previous.x, y: zero }, { x: previous.x, y: previous.y }, { x: point.x, y: point.y }, { x: point.x, y: zero },
  ]]
  const ratio = (zero - previous.y) / (point.y - previous.y)
  const crossing = { x: previous.x + (point.x - previous.x) * ratio, y: zero }
  return [
    [{ x: previous.x, y: zero }, { x: previous.x, y: previous.y }, crossing],
    [crossing, { x: point.x, y: point.y }, { x: point.x, y: zero }],
  ]
}

function hexRgb(hex: string): string {
  const match = /^#([0-9a-f]{6})$/i.exec(hex)
  if (!match) return '37, 99, 235'
  const value = match[1]!
  return `${Number.parseInt(value.slice(0, 2), 16)}, ${Number.parseInt(value.slice(2, 4), 16)}, ${Number.parseInt(value.slice(4, 6), 16)}`
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
