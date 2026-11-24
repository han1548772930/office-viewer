import type { CellRange, ChartAxisOptions, ChartLegendOptions, ChartModel, ChartSemanticOptions, ChartSourceOrientation, PivotTableModel, SpreadsheetSheet } from '../domain/model.ts'
import { effectiveSheetRowSizes, sheetAxisSpanAt, SparseAxisIndex } from '../domain/sheet-axis.ts'
import { cloneModel } from '../domain/serialization.ts'
import { isChartType } from '../features/charts/chart-catalog.ts'
import { validateChartAxis, validateChartFont, validateChartOptions, validateChartPalette } from '../features/charts/chart-options.ts'
import { array, finite, record } from './validation.ts'

interface ChartRange extends CellRange { sheetId: string }
interface SeriesVectors {
  raw: Record<string, unknown>
  x: ChartRange
  y: ChartRange
  size?: ChartRange
  errors: Array<{ raw: Record<string, unknown>; plus?: ChartRange; minus?: ChartRange }>
}

export interface UnresolvedOfficeChart {
  sheetId: string
  index: number
  reason: string
}
export interface WorkbookChartImport {
  charts: ChartModel[]
  /** Original chart DTO/XML stays in sheet.metadata.source.charts. Export must
   * reject unresolved objects instead of silently writing a chart-free file. */
  unresolved: UnresolvedOfficeChart[]
}

export function normalizeWorkbookCharts(rawSheets: unknown, sheets: readonly SpreadsheetSheet[], pivots: readonly PivotTableModel[]): WorkbookChartImport {
  const charts: ChartModel[] = [], unresolved: UnresolvedOfficeChart[] = []
  for (const [sheetIndex, sheetValue] of array(rawSheets).entries()) {
    const host = sheets[sheetIndex]
    if (!host) continue
    // Reuse one sparse geometry per host; no whole-axis enumeration or cell index.
    const rowSizes = effectiveSheetRowSizes(host)
    const rows = new SparseAxisIndex(1_048_576, host.defaultRowHeight ?? 26, rowSizes.sizes,
      host.hiddenRows, rowSizes.spans, host.rowHiddenSpans)
    const columnSizes = host.defaultColumnWidth === undefined && !sheetAxisSpanAt(host.columnWidthSpans, 0) ? { 0: 152, ...host.columnWidths } : host.columnWidths
    const columns = new SparseAxisIndex(16_384, host.defaultColumnWidth ?? 112, columnSizes,
      host.hiddenColumns, host.columnWidthSpans, host.columnHiddenSpans)
    const differences: Record<string, { index: number; reasons: string[] }> = {}
    for (const [chartIndex, chartValue] of array(record(sheetValue).charts).entries()) {
      const raw = record(chartValue), index = Number.isSafeInteger(raw.index) ? Number(raw.index) : chartIndex
      try {
        const pivotId = typeof raw.pivotTableId === 'string' ? raw.pivotTableId : undefined
        const pivot = pivotId ? pivots.find(item => item.id === pivotId) : undefined
        if (pivotId && (!pivot || !pivot.outputRange)) throw new TypeError('CHART_PIVOT_SOURCE_UNRESOLVED')
        const authored = record(raw.authored)
        const chart = Object.keys(authored).length ? fromAuthored(authored, raw, host, sheets) : fromNative(raw, host, sheets, pivot)
        chart.id = `${pivot ? 'pivot-' : ''}chart-${sheetIndex + 1}-${index + 1}`
        const reasons = array(raw.unsupportedReasons)
        if (reasons.some(reason => typeof reason !== 'string' || reason.length === 0)) throw new TypeError('CHART_IMPORT_DIAGNOSTICS_INVALID')
        const codes = reasons.filter((reason): reason is string => typeof reason === 'string')
        if (pivot) {
          chart.pivotTableId = pivot.id
          chart.source = { range: { ...pivot.outputRange! }, orientation: 'columns', sheetId: pivot.anchor.subUnitId }
        }
        const geometry = drawingGeometry(raw, rows, columns)
        chart.position = geometry.position
        chart.size = geometry.size
        restoreHelperOwnership(authored.helperSheets, chart, sheets)
        const source = sheets.find(sheet => sheet.id === chart.source.sheetId)
        if (source) {
          source.rowCount = Math.max(source.rowCount, chart.source.range.endRow + 1)
          source.columnCount = Math.max(source.columnCount, chart.source.range.endCol + 1)
        }
        if (codes.length) differences[chart.id] = { index, reasons: [...new Set(codes)] }
        charts.push(chart)
      } catch (error) {
        unresolved.push({ sheetId: host.id, index, reason: error instanceof Error ? error.message : 'CHART_IMPORT_INVALID' })
      }
    }
    if (Object.keys(differences).length) host.metadata.officeChartDifferences = differences
    else delete host.metadata.officeChartDifferences
  }
  return { charts, unresolved }
}

function restoreHelperOwnership(value: unknown, chart: ChartModel, sheets: readonly SpreadsheetSheet[]): void {
  for (const item of array(value)) {
    const helper = record(item)
    if (typeof helper.name !== 'string' || typeof helper.key !== 'string' || !helper.name.startsWith('__ov_chart_')) continue
    const sheet = sheets.find(sheet => sheet.name === helper.name && !sheet.visible)
    if (!sheet || sheet.id === chart.sheetId || sheet.id === chart.source.sheetId
      || sheet.cells['0:0']?.value !== `office-studio:chart-helper:v1:${helper.key}`) continue
    const previous = record(sheet.metadata.officeChartHelper)
    if (previous.chartId !== undefined && previous.chartId !== chart.id) throw new TypeError('CHART_HELPER_OWNER_CONFLICT')
    sheet.metadata.officeChartHelper = { version: 1, chartId: chart.id, key: helper.key }
  }
}

function chartType(value: unknown): ChartModel['type'] {
  if (!isChartType(value)) throw new TypeError(`CHART_NATIVE_TYPE_UNRESOLVED: ${String(value)}`)
  return value
}

function range(value: unknown): CellRange {
  const raw = record(value)
  const result = { startRow: finite(raw.startRow, -1), startCol: finite(raw.startCol, -1), endRow: finite(raw.endRow, -1), endCol: finite(raw.endCol, -1) }
  if (![result.startRow, result.startCol, result.endRow, result.endCol].every(Number.isSafeInteger)
    || result.startRow < 0 || result.startCol < 0 || result.endRow < result.startRow || result.endCol < result.startCol
    || result.endRow > 1_048_575 || result.endCol > 16_383) throw new RangeError('CHART_SOURCE_RANGE_INVALID')
  return result
}

function orientation(value: unknown): ChartSourceOrientation {
  if (value !== 'columns' && value !== 'rows') throw new TypeError('CHART_SOURCE_ORIENTATION_INVALID')
  return value
}

function integer(value: unknown, maximum: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value >= maximum) throw new RangeError('CHART_SOURCE_FIELD_INVALID')
  return value
}

function fieldList(value: unknown, maximum: number): number[] {
  if (!Array.isArray(value)) throw new TypeError('CHART_SOURCE_FIELDS_INVALID')
  const fields = value.map(field => integer(field, maximum))
  if (new Set(fields).size !== fields.length) throw new TypeError('CHART_SOURCE_FIELDS_DUPLICATE')
  return fields
}

function options(value: unknown): ChartSemanticOptions {
  // Runtime validators cover known typed properties and the complete extension
  // tree before its values can enter the persistent model.
  const result: ChartSemanticOptions = { ...optionalObject(value, 'CHART_OPTIONS_INVALID') }
  validateChartOptions(result)
  return cloneModel(result)
}

function axis(value: unknown, title?: unknown): ChartAxisOptions {
  const raw = optionalObject(value, 'CHART_AXIS_INVALID')
  if (raw.scaleType !== undefined && raw.scaleType !== 'linear' && raw.scaleType !== 'log') throw new TypeError('CHART_AXIS_SCALE_INVALID')
  const result: ChartAxisOptions = { ...raw, scaleType: raw.scaleType === 'log' ? 'log' : 'linear',
    ...(typeof title === 'string' && raw.title === undefined ? { title } : {}) }
  validateChartAxis(result, '导入图表轴')
  return cloneModel(result)
}

function legend(raw: Record<string, unknown>, authored = false): ChartLegendOptions {
  const value = authored ? raw : { visible: raw.legend !== false, position: raw.legendPosition, font: raw.legendFont }
  const position = value.position === 'top' || value.position === 'right' || value.position === 'left' ? value.position : 'bottom'
  const result: ChartLegendOptions = { visible: value.visible !== false, position }
  if (value.font !== undefined && value.font !== null) {
    const font = { ...optionalObject(value.font, 'CHART_LEGEND_FONT_INVALID') }
    validateChartFont(font, '导入图例字体')
    result.font = cloneModel(font)
  }
  return result
}

function optionalObject(value: unknown, message: string): Record<string, unknown> {
  if (value === undefined || value === null) return {}
  if (typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype) throw new TypeError(message)
  return record(value)
}

function base(raw: Record<string, unknown>, host: SpreadsheetSheet, type: ChartModel['type']): ChartModel {
  return { id: '', sheetId: host.id, type, source: { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, orientation: 'columns' },
    position: { x: 0, y: 0 }, size: { width: 480, height: 300 }, categoryField: 0, valueFields: [],
    ...(typeof raw.title === 'string' ? { title: raw.title } : {}), ...(typeof raw.subtitle === 'string' ? { subtitle: raw.subtitle } : {}),
    legend: legend(raw), xAxis: axis(raw.xAxis, raw.xAxisTitle), yAxis: axis(raw.yAxis, raw.yAxisTitle), revision: 0 }
}

function fromAuthored(authored: Record<string, unknown>, raw: Record<string, unknown>, host: SpreadsheetSheet, sheets: readonly SpreadsheetSheet[]): ChartModel {
  const source = record(authored.source), sourceRange = range(source.range), direction = orientation(source.orientation)
  const sheet = sheets.find(item => typeof source.sheetName === 'string' && item.name.toLowerCase() === source.sheetName.toLowerCase())
  if (!sheet) throw new TypeError('CHART_AUTHORED_SOURCE_SHEET_UNRESOLVED')
  const fields = direction === 'columns' ? sourceRange.endCol - sourceRange.startCol + 1 : sourceRange.endRow - sourceRange.startRow + 1
  if ((direction === 'columns' ? sourceRange.endRow - sourceRange.startRow : sourceRange.endCol - sourceRange.startCol) < 1) throw new RangeError('CHART_SOURCE_DATA_MISSING')
  const result = base(authored, host, chartType(authored.type))
  result.source = { sheetId: sheet.id, range: sourceRange, orientation: direction }
  result.categoryField = integer(authored.categoryField, fields)
  result.valueFields = fieldList(authored.valueFields, fields)
  if (authored.seriesToCategory !== undefined) {
    result.seriesToCategory = {}
    for (const [key, value] of Object.entries(record(authored.seriesToCategory))) {
      if (!/^(0|[1-9]\d*)$/.test(key)) throw new TypeError('CHART_SCATTER_FIELD_INVALID')
      result.seriesToCategory[integer(Number(key), fields)] = fieldList(value, fields)
    }
  }
  result.options = options(authored.options)
  result.legend = authored.legend === undefined || authored.legend === null ? legend(raw) : legend(record(authored.legend), true)
  result.xAxis = axis(authored.xAxis ?? raw.xAxis, raw.xAxisTitle)
  result.yAxis = axis(authored.yAxis ?? raw.yAxis, raw.yAxisTitle)
  if (authored.palette !== undefined) {
    if (!Array.isArray(authored.palette) || !authored.palette.every((color): color is string => typeof color === 'string')) throw new TypeError('CHART_PALETTE_INVALID')
    validateChartPalette(authored.palette)
    result.palette = [...authored.palette]
  }
  validateFieldRoles(result, fields)
  return result
}

function fromNative(raw: Record<string, unknown>, host: SpreadsheetSheet, sheets: readonly SpreadsheetSheet[], pivot?: PivotTableModel): ChartModel {
  const result = base(raw, host, chartType(raw.kind)), series = array(raw.series)
  if (!series.length) throw new TypeError('CHART_NATIVE_SERIES_MISSING')
  const vectors: SeriesVectors[] = series.map(value => {
    const item = record(value), x = parseChartRange(item.categories, sheets, host), y = parseChartRange(item.values, sheets, host)
    if (!x || !y) throw new TypeError('CHART_LITERAL_OR_UNBOUND_SERIES')
    const size = item.bubbleSizes ? parseChartRange(item.bubbleSizes, sheets, host) : undefined
    if (item.bubbleSizes && !size) throw new TypeError('CHART_BUBBLE_SIZE_SOURCE_UNRESOLVED')
    const errors = array(item.errorBars).map(value => {
      const error = record(value), plus = error.plusFormula ? parseChartRange(error.plusFormula, sheets, host) : undefined,
        minus = error.minusFormula ? parseChartRange(error.minusFormula, sheets, host) : undefined
      if (error.plusFormula && !plus || error.minusFormula && !minus) throw new TypeError('CHART_ERROR_SOURCE_UNRESOLVED')
      return { raw: error, plus, minus }
    })
    return { raw: item, x, y, size, errors }
  })
  const first = vectors[0]!.y, direction: ChartSourceOrientation = first.startCol === first.endCol ? 'columns' : 'rows'
  const ranges = vectors.flatMap(vector => [vector.x, vector.y, ...(vector.size ? [vector.size] : []), ...vector.errors.flatMap(error => [error.plus, error.minus].filter((range): range is ChartRange => range !== undefined))])
  const aligned = ranges.every(vector => vector.sheetId === first.sheetId && (direction === 'columns'
    ? vector.startRow === first.startRow && vector.endRow === first.endRow
    : vector.startCol === first.startCol && vector.endCol === first.endCol))
  if (!aligned || (direction === 'columns' ? first.startRow === 0 : first.startCol === 0)) throw new TypeError('CHART_DISJOINT_OR_HEADERLESS_SOURCE_UNRESOLVED')
  if (vectors.some(vector => direction === 'columns' ? vector.y.startCol !== vector.y.endCol : vector.y.startRow !== vector.y.endRow)) throw new TypeError('CHART_VALUE_VECTOR_INVALID')
  for (const vector of vectors) {
    for (const auxiliary of [vector.size, ...vector.errors.flatMap(error => [error.plus, error.minus])]) {
      if (auxiliary && (direction === 'columns' ? auxiliary.startCol !== auxiliary.endCol : auxiliary.startRow !== auxiliary.endRow)) throw new TypeError('CHART_AUXILIARY_VECTOR_INVALID')
    }
    if (result.type === 'bubble' && (direction === 'columns' ? vector.x.startCol !== vector.x.endCol : vector.x.startRow !== vector.x.endRow)) throw new TypeError('CHART_BUBBLE_X_VECTOR_INVALID')
  }
  const sourceRange = { startRow: Math.min(...ranges.map(item => item.startRow)) - (direction === 'columns' ? 1 : 0),
    startCol: Math.min(...ranges.map(item => item.startCol)) - (direction === 'rows' ? 1 : 0),
    endRow: Math.max(...ranges.map(item => item.endRow)), endCol: Math.max(...ranges.map(item => item.endCol)) }
  const field = (vector: ChartRange) => direction === 'columns' ? vector.startCol - sourceRange.startCol : vector.startRow - sourceRange.startRow
  const categories = (vector: ChartRange) => Array.from({ length: direction === 'columns' ? vector.endCol - vector.startCol + 1 : vector.endRow - vector.startRow + 1 }, (_, index) => field(vector) + index)
  result.source = { range: sourceRange, orientation: direction, sheetId: first.sheetId }
  result.categoryField = field(vectors[0]!.x)
  result.valueFields = vectors.map(vector => field(vector.y))
  if (new Set(result.valueFields).size !== result.valueFields.length) throw new TypeError('CHART_DUPLICATE_NATIVE_SERIES_UNRESOLVED')
  result.options = options(raw.options)
  const opts = result.options
  opts.series = {}
  const categoryFields = categories(vectors[0]!.x)
  if (result.type === 'scatter') result.seriesToCategory = Object.fromEntries(vectors.map(vector => {
    const fields = categories(vector.x)
    if (fields.length !== 1) throw new TypeError('CHART_SCATTER_X_VECTOR_INVALID')
    return [field(vector.y), fields]
  }))
  else if (vectors.some(vector => JSON.stringify(categories(vector.x)) !== JSON.stringify(categoryFields))) throw new TypeError('CHART_SERIES_CATEGORY_MISMATCH')
  if (categoryFields.length > 1) opts.mapping = { ...opts.mapping, categoryIndexes: categoryFields }
  const trends: unknown[] = [], errors: unknown[] = []
  for (const vector of vectors) {
    const sourceField = field(vector.y), style = { ...optionalObject(vector.raw.options, 'CHART_SERIES_OPTIONS_INVALID') }
    if (typeof vector.raw.name === 'string' && style.name === undefined) style.name = vector.raw.name
    if (raw.dataLabels === true && style.label === undefined) style.label = { visible: true, contentType: 'value' }
    if (result.type === 'combination' && !['column', 'columnStacked', 'line', 'area', 'areaStacked'].includes(String(style.type))) throw new TypeError('CHART_COMBINATION_SERIES_TYPE_UNRESOLVED')
    // validateChartOptions checks the complete per-field style before committing.
    const checked = options({ series: { [sourceField]: style } })
    opts.series[sourceField] = checked.series![sourceField]!
    trends.push(...array(vector.raw.trendlines).map(value => ({ ...record(value), seriesIndex: sourceField })))
    for (const error of vector.errors) {
      const { plusFormula: _plus, minusFormula: _minus, ...config } = error.raw
      errors.push({ ...config, seriesIndex: sourceField, ...(error.plus ? { plusIndex: field(error.plus) } : {}), ...(error.minus ? { minusIndex: field(error.minus) } : {}) })
    }
  }
  Object.assign(opts, options({ trendlines: trends, errorBars: errors }))
  if (result.type === 'bubble') {
    if (vectors.length !== 1) throw new TypeError('CHART_MULTI_BUBBLE_SERIES_UNRESOLVED')
    const vector = vectors[0]!
    opts.mapping = { ...opts.mapping, bubble: { xIndex: field(vector.x), yIndex: field(vector.y), ...(vector.size ? { sizeIndex: field(vector.size) } : {}) } }
    result.valueFields = [field(vector.y), ...(vector.size ? [field(vector.size)] : [])]
  }
  if (result.type === 'candlestick') {
    const values = result.valueFields
    if (values.length !== 3 && values.length !== 4) throw new TypeError('CHART_OHLC_SERIES_INVALID')
    const offset = values.length === 4 ? 1 : 0
    opts.mapping = { ...opts.mapping, candlestick: { ...(offset ? { openIndex: values[0] } : {}), highIndex: values[offset]!, lowIndex: values[offset + 1]!, closeIndex: values[offset + 2]! } }
  }
  if (pivot && pivot.outputRange) {
    // Preserve field/style identities relative to the persisted pivot output.
    if (direction !== 'columns' || first.sheetId !== pivot.anchor.subUnitId || sourceRange.startRow !== pivot.outputRange.startRow
      || sourceRange.startCol !== pivot.outputRange.startCol) throw new TypeError('CHART_PIVOT_BINDING_UNRESOLVED')
  }
  validateChartOptions(opts)
  validateFieldRoles(result, direction === 'columns' ? sourceRange.endCol - sourceRange.startCol + 1 : sourceRange.endRow - sourceRange.startRow + 1)
  return result
}

function validateFieldRoles(chart: ChartModel, maximum: number): void {
  const opts = chart.options
  for (const role of [opts?.mapping?.bubble, opts?.mapping?.candlestick, opts?.mapping?.edge, opts?.mapping?.gauge]) {
    for (const value of Object.values(role ?? {})) if (value !== undefined) integer(value, maximum)
  }
  for (const fields of [opts?.mapping?.categoryIndexes, opts?.mapping?.hierarchyIndexes]) for (const value of fields ?? []) integer(value, maximum)
  for (const key of Object.keys(opts?.series ?? {})) integer(Number(key), maximum)
  for (const trend of opts?.trendlines ?? []) integer(trend.seriesIndex, maximum)
  for (const error of opts?.errorBars ?? []) for (const value of [error.seriesIndex, error.plusIndex, error.minusIndex]) if (value !== undefined) integer(value, maximum)
}

/** Native formulas use arbitrary worksheet names and absolute A1 coordinates. */
export function parseChartRange(value: unknown, sheets: readonly SpreadsheetSheet[], host: SpreadsheetSheet): ChartRange | undefined {
  if (typeof value !== 'string') return undefined
  const match = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/i.exec(value.trim().replace(/^=/, ''))
  if (!match) return undefined
  const name = match[1]?.replaceAll("''", "'") ?? match[2]
  const source = name === undefined ? host : sheets.find(sheet => sheet.name.toLowerCase() === name.toLowerCase())
  if (!source) return undefined
  const column = (label: string) => [...label.toUpperCase()].reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0) - 1
  try { return { sheetId: source.id, ...range({ startRow: Number(match[4]) - 1, startCol: column(match[3]!), endRow: Number(match[6] ?? match[4]) - 1, endCol: column(match[5] ?? match[3]!) }) } }
  catch { return undefined }
}

function drawingGeometry(raw: Record<string, unknown>, rows: SparseAxisIndex, columns: SparseAxisIndex): Pick<ChartModel, 'position' | 'size'> {
  const drawing = record(raw.drawing), from = record(drawing.fromCell ?? raw.from), to = record(drawing.toCell ?? raw.to)
  const offset = array(drawing.fromOffsetEmu), endOffset = array(drawing.toOffsetEmu), ext = array(drawing.absoluteSizeEmu)
  const point = (cell: Record<string, unknown>, offset: unknown[]) => ({
    x: columns.offset(integer(cell.col ?? 0, 16_384)) + finite(offset[0]) / 9525,
    y: rows.offset(integer(cell.row ?? 0, 1_048_576)) + finite(offset[1]) / 9525,
  })
  // editAs="absolute" on a two-cell anchor changes resize behavior; its
  // coordinates still come from the cells. Only absoluteAnchor has no fromCell.
  const absoluteAnchor = drawing.behavior === 'absolute' && drawing.fromCell == null
  const position = absoluteAnchor ? { x: finite(offset[0]) / 9525, y: finite(offset[1]) / 9525 } : point(from, offset)
  const end = point(to, endOffset)
  const width = ext[0] === undefined ? end.x > position.x ? end.x - position.x : 480 : finite(ext[0], -1) / 9525
  const height = ext[1] === undefined ? end.y > position.y ? end.y - position.y : 300 : finite(ext[1], -1) / 9525
  if (!Number.isFinite(position.x) || !Number.isFinite(position.y) || position.x < 0 || position.y < 0
    || !Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) throw new RangeError('CHART_DRAWING_GEOMETRY_INVALID')
  return { position, size: { width, height } }
}
