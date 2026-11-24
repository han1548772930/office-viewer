import type { CellScalar } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { finiteChartValue } from './chart-math'
import { chartDateSerial, chartDateDataSteps, type ChartDateAxisConfig } from './chart-date-axis'
import { CHART_PREPARATION_CHUNK, consumeChartPreparation, type ChartComputePreparationContext } from './chart-compute-lifecycle'

interface CategoryEntry { category: string; indexes: number[]; values: Array<number | null>; weight?: number; synthetic?: boolean }

/** Both the synchronous public contract and asynchronous preparation consume
 * this implementation. The yielded values are scheduling boundaries only. */
export function* chartComputeInputSteps(source: ChartData, chart: ChartDateAxisConfig): Generator<void, ChartData, void> {
  const data = resolveChartType(chart.type).dataShape === 'category' ? yield* aggregateCategorySteps(source, chart) : source
  const percent = resolveChartType(chart.type).percent
  const scales: number[] = [], totals: number[] = []
  let work = 0
  if (percent) {
    for (let row = 0; row < data.categories.length; row++) {
      let scale = 0
      for (const series of data.series) {
        scale = Math.max(scale, Math.abs(series.values[row] ?? 0))
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
      let total = 0
      for (const series of data.series) {
        if (scale) total += Math.abs(series.values[row] ?? 0) / scale
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
      scales.push(scale); totals.push(total)
    }
  }
  const series: ChartData['series'] = []
  for (const input of data.series) {
    const values: Array<number | null> = []
    const formattedValues: Array<string | null> | undefined = input.formattedValues ? [] : undefined
    for (let row = 0; row < input.values.length; row++) {
      const value = input.values[row]!
      values.push(value === null ? chart.type !== 'gauge' && chart.type !== 'histogram' && chart.options?.invalidValueStrategy === 'zero' ? 0 : null
        : percent ? totals[row] ? (value / scales[row]!) / totals[row]! * 100 : 0 : value)
      if (formattedValues) formattedValues.push(Object.is(value, values[row]) ? input.formattedValues![row] ?? null : null)
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    series.push({ ...input, values, ...(formattedValues ? { formattedValues } : {}) })
  }
  const prepared = { ...data, series }
  // Statistical plots create their dates from the final bin centers or field
  // labels. Ordinary Cartesian statistics use the actual source serials.
  return ['histogram', 'boxplot'].includes(chart.type) ? prepared : yield* chartDateDataSteps(prepared, chart)
}

export function prepareChartComputeInputs(source: ChartData, chart: ChartDateAxisConfig, context: ChartComputePreparationContext = {}): Promise<ChartData> {
  return consumeChartPreparation(chartComputeInputSteps(source, chart), context)
}

function* finiteSumSteps<T>(items: readonly T[], valueAt: (item: T) => number | null | undefined, label: string): Generator<void, number | null, void> {
  let scale = 0, count = 0, work = 0
  for (const item of items) {
    const value = valueAt(item)
    if (value != null) { scale = Math.max(scale, Math.abs(finiteChartValue(value, label))); count++ }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (!count) return null
  if (!scale) return 0
  let total = 0, compensation = 0
  for (const item of items) {
    const value = valueAt(item)
    if (value != null) {
      const addend = value / scale - compensation, next = total + addend
      compensation = next - total - addend; total = next
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return finiteChartValue(total * scale, label)
}

function* aggregateCategorySteps(data: ChartData, chart: ChartDateAxisConfig): Generator<void, ChartData, void> {
  const options = chart.options?.aggregation
  if (!options?.aggregate && !options?.topN) return data
  const groups = new Map<string, number[]>()
  let work = 0
  for (let row = 0; row < data.categories.length; row++) {
    const category = data.categories[row]!, key = options.aggregate ? category : `${row}\u0000${category}`
    const group = groups.get(key)
    if (group) group.push(row); else groups.set(key, [row])
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  let entries: CategoryEntry[] = []
  for (const indexes of groups.values()) {
    const values: Array<number | null> = []
    for (const series of data.series) {
      values.push(yield* finiteSumSteps(indexes, row => series.values[row], '图表分类合计'))
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    entries.push({ category: data.categories[indexes[0]!]!, indexes, values })
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  groups.clear()
  if (options.topN) {
    let scale = 1
    for (const entry of entries) for (const value of entry.values) {
      scale = Math.max(scale, Math.abs(value ?? 0))
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    for (const entry of entries) {
      let weight = 0
      for (const value of entry.values) {
        weight += Math.abs(value ?? 0) / scale
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
      entry.weight = weight
    }
    const sorted = yield* sortCategorySteps(entries)
    const retained: CategoryEntry[] = [], remaining: CategoryEntry[] = []
    for (let index = 0; index < sorted.length; index++) {
      ;(index < options.topN.count ? retained : remaining).push(sorted[index]!)
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    entries = retained
    if (options.topN.includeOthers && remaining.length) {
      const indexes: number[] = [], values: Array<number | null> = []
      for (const entry of remaining) for (const row of entry.indexes) {
        indexes.push(row)
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
      for (let series = 0; series < data.series.length; series++) {
        values.push(yield* finiteSumSteps(remaining, entry => entry.values[series], '图表其他分类合计'))
      }
      entries.push({ category: options.topN.othersLabel ?? '其他', indexes, values, synthetic: true })
    }
  }
  const categories: string[] = [], records: NonNullable<ChartData['records']> = []
  const categoryValues: CellScalar[] | undefined = data.categoryValues ? [] : undefined
  let individualRecords = Boolean(data.records)
  for (const entry of entries) {
    categories.push(entry.category)
    if (categoryValues) categoryValues.push(yield* categorySourceValueSteps(data, entry, chart))
    if (entry.synthetic || entry.indexes.length !== 1) individualRecords = false
    else if (data.records) records.push(data.records[entry.indexes[0]!]!)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const series: ChartData['series'] = []
  for (let index = 0; index < data.series.length; index++) {
    const values: Array<number | null> = []
    for (const entry of entries) {
      values.push(entry.values[index] ?? null)
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    const input = data.series[index]!
    const xValues = input.xValues ? yield* projectCategoryCoordinatesSteps(input.xValues, entries) : undefined
    const sizeValues = input.sizeValues ? yield* projectCategoryCoordinatesSteps(input.sizeValues, entries) : undefined
    series.push({ ...input, values, ...(xValues ? { xValues } : {}), ...(sizeValues ? { sizeValues } : {}) })
  }
  return { categories, fields: data.fields, date1904: data.date1904, sourceCategoryFields: data.sourceCategoryFields,
    sourceValueFields: data.sourceValueFields, sourceNumericFields: data.sourceNumericFields, visibleFields: data.visibleFields,
    ...(categoryValues ? { categoryValues } : {}), ...(individualRecords ? { records } : {}), series }
}

/** A merged category has one coordinate only when all contributing rows agree.
 * Synthetic Others is an aggregate label with no individual source position. */
function* projectCategoryCoordinatesSteps(source: Array<number | null>, entries: CategoryEntry[]): Generator<void, Array<number | null>, void> {
  const projected: Array<number | null> = []
  let work = 0
  for (const entry of entries) {
    let coordinate = entry.synthetic ? null : source[entry.indexes[0]!] ?? null
    for (const row of entry.indexes) {
      if (!Object.is(coordinate, source[row] ?? null)) coordinate = null
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    projected.push(coordinate === null ? null : finiteChartValue(coordinate, '图表聚合坐标'))
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return projected
}

function* categorySourceValueSteps(data: ChartData, entry: CategoryEntry, chart: ChartDateAxisConfig): Generator<void, CellScalar, void> {
  if (entry.synthetic) return entry.category
  const first = data.categoryValues?.[entry.indexes[0]!] ?? null
  const date = chart.options?.mapping?.useDateAxis === true
  const serial = date ? chartDateSerial(first, data.date1904) : null
  let work = 0
  for (const row of entry.indexes) {
    const value = data.categoryValues?.[row] ?? null
    if (!Object.is(first, value) && !(date && serial !== null && chartDateSerial(value, data.date1904) === serial)) return entry.category
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return first
}

/** Stable bottom-up merge sorting yields during writes, including a large
 * Top-N source. Weight ties keep the first source row's observable order. */
function* sortCategorySteps(input: CategoryEntry[]): Generator<void, CategoryEntry[], void> {
  let source = input, target = new Array<CategoryEntry>(input.length), work = 0
  const before = (a: CategoryEntry, b: CategoryEntry): boolean => a.weight! > b.weight!
    || a.weight === b.weight && a.indexes[0]! <= b.indexes[0]!
  for (let width = 1; width < input.length; width *= 2) {
    for (let start = 0; start < input.length; start += width * 2) {
      const middle = Math.min(start + width, input.length), end = Math.min(start + width * 2, input.length)
      let left = start, right = middle
      for (let index = start; index < end; index++) {
        target[index] = left < middle && (right >= end || before(source[left]!, source[right]!)) ? source[left++]! : source[right++]!
        if (++work % CHART_PREPARATION_CHUNK === 0) yield
      }
    }
    ;[source, target] = [target, source]
  }
  return source
}
