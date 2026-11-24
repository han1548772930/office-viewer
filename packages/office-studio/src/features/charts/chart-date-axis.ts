import type { CellScalar, ChartAxisOptions, ChartModel } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartAxisRatio } from './chart-geometry'
import { chartSeriesStyle } from './chart-options'
import { DateSystem } from '../sheet-workbench/number-format-engine/date-system'
import { excelDateTimePartsToSerial, excelSerialToDateTimeParts, getDateSystemMaxSerial } from '../sheet-workbench/number-format-engine/serial-date'
import { formatSheetValue } from '../sheet-workbench/sheet-number-format'
import { chartDateCoordinate, chartDateTicks, validateChartDateCalendarOptions, type ChartDateAxisTick } from './chart-date-calendar'
import { validateChartAxis } from './chart-option-validation'
import { CHART_PREPARATION_CHUNK, consumeChartPreparationSync } from './chart-compute-lifecycle'
export { chartDateAdvance, chartDateCoordinate, chartDateCoordinateSerial, chartDateTicks, CHART_DATE_AXIS_MAX_TICKS } from './chart-date-calendar'
export type { ChartDateAxisTick } from './chart-date-calendar'

export interface ChartDateAxis {
  /** Original Excel serials, retaining each source row and repeated date. */
  values: Array<number | null>
  min: number
  max: number
  sourceMin: number
  sourceMax: number
  coordinates: Array<number | null>
  coordinateMin: number
  coordinateMax: number
  coordinateStep: number
  /** Stable chronological order; indices continue to identify original rows. */
  order: number[]
  ticks: ChartDateAxisTick[]
  /** Smallest positive distance between observations, in Excel serial days. */
  step: number
  bandRatio: number
  date1904: boolean
}

export type ChartDateAxisConfig = Pick<ChartModel, 'type' | 'options' | 'xAxis'>

/** Numeric serials and strict ISO calendar text share Sheets' date-system semantics. */
export function chartDateSerial(value: CellScalar | undefined, date1904 = false): number | null {
  const options = { dateSystem: date1904 ? DateSystem.Date1904 : DateSystem.Date1900 }
  if (typeof value === 'number') return excelSerialToDateTimeParts(value, options) ? value : null
  if (typeof value !== 'string' || !value.trim()) return null
  const text = value.trim()
  if (/^[+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) {
    const serial = Number(text)
    return excelSerialToDateTimeParts(serial, options) ? serial : null
  }
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?(Z|[+-]\d{2}:\d{2})?)?$/.exec(text)
  if (!match) return null
  const serial = excelDateTimePartsToSerial({ year: Number(match[1]), month: Number(match[2]), day: Number(match[3]),
    hours: Number(match[4] ?? 0), minutes: Number(match[5] ?? 0), seconds: Number(match[6] ?? 0), fractionalSecond: Number(`0.${match[7] ?? '0'}`) }, options)
  if (serial === null) return null
  const zone = match[8]
  if (!zone || zone === 'Z') return serial
  const hours = Number(zone.slice(1, 3)), minutes = Number(zone.slice(4, 6))
  if (hours > 23 || minutes > 59) return null
  let adjusted = serial - (zone.startsWith('-') ? -1 : 1) * (hours * 60 + minutes) / 1440
  // UTC offsets on real calendar dates skip Excel's fictional 1900-02-29.
  // An explicitly authored fictional day keeps serial arithmetic instead.
  if (!date1904) {
    if (serial >= 61 && adjusted < 61) adjusted -= 1
    else if (serial < 60 && adjusted >= 60) adjusted += 1
  }
  return excelSerialToDateTimeParts(adjusted, options) ? adjusted : null
}

/** Matches the upstream X-axis control where a Cartesian axis is present. */
export function supportsChartDateAxis(chart: Pick<ChartModel, 'type'>): boolean {
  return ['line', 'column', 'bar', 'area', 'bubble', 'combination', 'pareto', 'boxplot', 'candlestick', 'histogram'].includes(resolveChartType(chart.type).family)
}

export function prepareChartDateData(data: ChartData, chart: ChartDateAxisConfig): ChartData {
  return consumeChartPreparationSync(chartDateDataSteps(data, chart))
}

export function* chartDateDataSteps(data: ChartData, chart: ChartDateAxisConfig): Generator<void, ChartData, void> {
  if (!chart.options?.mapping?.useDateAxis || !supportsChartDateAxis(chart)) return data
  validateDateAxisOptions(chart.xAxis, Boolean(data.date1904))
  const raw = data.categoryValues ?? data.categories
  if (raw.length !== data.categories.length) throw new RangeError('日期轴的分类值血缘与当前数据不匹配。')
  const values: Array<number | null> = []
  let work = 0
  for (let index = 0; index < raw.length; index++) {
    const value = raw[index]
    const serial = chartDateSerial(value, data.date1904)
    if (serial === null && value !== null && value !== undefined && !(typeof value === 'string' && !value.trim())) {
      throw new RangeError(`第 ${index + 1} 个分类无法作为日期轴值：${String(value)}。请使用有效的 Excel 日期序号或 ISO 日期。`)
    }
    values.push(serial)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  let hasValue = false, hasPair = false
  for (const series of data.series) for (let index = 0; index < series.values.length; index++) {
    const value = series.values[index]
    if (value != null && Number.isFinite(value)) { hasValue = true; if (values[index] != null) hasPair = true }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (data.categories.length && hasValue && !hasPair) {
    throw new RangeError('日期轴至少需要一个与数值对应的有效日期。')
  }
  const bubble = resolveChartType(chart.type).family === 'bubble'
  const series: ChartData['series'] = []
  for (const input of data.series) {
    const xValues: Array<number | null> = bubble ? [] : values
    if (bubble) for (let index = 0; index < values.length; index++) {
      xValues.push(input.xValues?.[index] != null || input.values[index] != null ? values[index]! : null)
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    series.push({ ...input, xValues })
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return { ...data, categoryDateValues: values, series }
}

/** Candidate/API and form validation share the same raw-source semantics, before expensive statistics. */
export function prepareChartDateSource(data: ChartData, chart: ChartDateAxisConfig): ChartData {
  if (!chart.options?.mapping?.useDateAxis || !supportsChartDateAxis(chart)) return data
  const family = resolveChartType(chart.type).family
  if (family === 'histogram') {
    // The actual time coordinates are final bin boundaries/centers, not source category labels.
    chartDateAxis({ categories: [], series: [], date1904: data.date1904 }, chart)
    return data
  }
  if (family === 'boxplot') {
    const categories = data.series.map((series, index) => chartSeriesStyle(chart, series.field, index, undefined, series.sourceName ?? series.name).name ?? series.name)
    const values = data.series.map(series => series.values.find(value => value !== null && Number.isFinite(value)) ?? null)
    // A finite representative only validates date/sample pairing; it is never a rendered or stored summary.
    chartDateAxis({ categories, categoryValues: categories, date1904: data.date1904,
      series: [{ field: data.series[0]?.field ?? 0, name: '样本分布', values }] }, chart)
    return data
  }
  const prepared = prepareChartDateData(data, chart)
  chartDateAxis(prepared, chart)
  if (family === 'pareto') chartParetoDateValues(prepared, prepared.categories, chart.options?.pareto?.includeZeroValues)
  return prepared
}

/** A grouped category has one date identity; elapsed-time coordinates survive Pareto's weight order. */
export function chartParetoDateValues(data: ChartData, categories: readonly string[], includeZeroValues = false): Array<number | null> {
  const dates = data.categoryDateValues
  if (!dates || dates.length !== data.categories.length) throw new RangeError('帕累托图缺少与源分类匹配的日期轴值。')
  const lineage = new Map<string, number | null>(), active = new Set(categories)
  const values = data.series[0]?.values ?? []
  for (let index = 0; index < values.length; index++) {
    const value = values[index], category = data.categories[index]
    if (value == null || !Number.isFinite(value) || !includeZeroValues && value === 0 || category === undefined || !active.has(category)) continue
    const date = dates[index] ?? null
    if (lineage.has(category) && lineage.get(category) !== date) throw new RangeError('帕累托图合并的同名分类不能包含不同的日期轴值。')
    lineage.set(category, date)
  }
  return categories.map(category => lineage.get(category) ?? null)
}

export function chartDateAxis(data: ChartData, chart: ChartDateAxisConfig): ChartDateAxis | undefined {
  if (!chart.options?.mapping?.useDateAxis || !supportsChartDateAxis(chart)) return undefined
  validateDateAxisOptions(chart.xAxis, Boolean(data.date1904))
  const prepared = data.categoryDateValues ? data : prepareChartDateData(data, chart)
  const values = prepared.categoryDateValues ?? [], date1904 = Boolean(data.date1904)
  if (values.length !== data.categories.length) throw new RangeError('日期轴位置与当前分类数据不匹配。')
  const order = values.flatMap((value, index) => value === null ? [] : [index]).sort((a, b) => values[a]! - values[b]! || a - b)
  const unique = [...new Set(order.map(index => values[index]!))]
  if (!unique.length) return undefined
  if (unique.some(value => chartDateSerial(value, date1904) === null)) throw new RangeError('日期轴值超出工作簿日期系统的范围。')
  let step = unique.length === 1 ? 1 : Infinity
  for (let index = 1; index < unique.length; index++) step = Math.min(step, unique[index]! - unique[index - 1]!)
  const limit = getDateSystemMaxSerial(date1904 ? DateSystem.Date1904 : DateSystem.Date1900) + 1
  const min = chart.xAxis.min ?? Math.max(0, unique[0]! - step / 2)
  const max = chart.xAxis.max ?? Math.min(limit, unique.at(-1)! + step / 2)
  if (min >= max) throw new RangeError('日期轴最小值必须小于最大值。')
  const calendar = chart.xAxis.baseTimeUnit === 'months' || chart.xAxis.baseTimeUnit === 'years'
  const coordinates = values.map(value => value === null ? null : chartDateCoordinate(value, chart.xAxis.baseTimeUnit, date1904) ?? null)
  // A valid authored edge can clip every observation within the automatic
  // padded range. The nearest visible edge still gives a valid tick anchor.
  const sourceMin = chart.xAxis.min ?? Math.min(unique[0]!, max), sourceMax = chart.xAxis.max ?? Math.max(unique.at(-1)!, min)
  const coordinateMin = calendar ? chartDateCoordinate(sourceMin, chart.xAxis.baseTimeUnit, date1904)! - 0.5 : min,
    coordinateMax = calendar ? chartDateCoordinate(sourceMax, chart.xAxis.baseTimeUnit, date1904)! + 0.5 : max,
    coordinateStep = calendar ? 1 : step
  const ticks = chartDateTicks({ min: coordinateMin, max: coordinateMax, sourceMin, sourceMax, date1904 }, chart.xAxis)
  return { values, min, max, sourceMin, sourceMax, step, coordinates, coordinateMin, coordinateMax, coordinateStep, order, ticks,
    bandRatio: Math.min(1, coordinateStep / (coordinateMax - coordinateMin)), date1904 }
}

function validateDateAxisOptions(axis: ChartAxisOptions, date1904: boolean): void {
  validateChartAxis(axis, '日期 X 轴')
  validateChartDateCalendarOptions(axis)
  if (axis.scaleType === 'log') throw new TypeError('日期轴须使用线性时间刻度。')
  for (const bound of [axis.min, axis.max]) if (bound !== undefined && chartDateSerial(bound, date1904) === null) throw new RangeError('日期轴边界须为当前日期系统中的有效序号。')
  if (axis.min !== undefined && axis.max !== undefined && axis.min >= axis.max) throw new RangeError('日期轴最小值必须小于最大值。')
}

/** The Pareto converter keeps its computed weight order; chronological ordering
 * applies to ordinary Cartesian source rows without changing their identities. */
export function chartDateSeriesOrder(axis: ChartDateAxis, chart: Pick<ChartModel, 'type'>): number[] {
  return chart.type === 'pareto' ? axis.values.flatMap((value, index) => value === null ? [] : [index]) : axis.order
}

export function chartDatePosition(axis: ChartDateAxis, options: ChartAxisOptions, index: number): number | undefined {
  const value = axis.coordinates[index]
  return value == null ? undefined : chartAxisRatio(value, axis.coordinateMin, axis.coordinateMax, options)
}

export function chartDateAxisLabel(value: number, axis: Pick<ChartAxisOptions, 'format'>, date1904 = false, span = 1): string {
  if (chartDateSerial(value, date1904) === null) return ''
  return formatSheetValue(value, axis.format ?? (span < 1 ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd'), date1904) ?? ''
}
