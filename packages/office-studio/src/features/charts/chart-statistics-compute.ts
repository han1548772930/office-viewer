import type { ChartModel } from '../../domain/model'
import type { ChartStatisticsJob, ChartStatisticsParams, ChartStatisticsTask } from '../../compute/chart-protocol.ts'
import { CHART_COMPUTE_LIMITS } from '../../compute/chart-protocol.ts'
import { estimateChartTaskMemory, isChartStatisticsParams, validateChartTask } from '../../compute/chart-tasks.ts'
import { readChartStatisticsOutput } from '../../compute/chart-results.ts'
import { CELL_STATE, COMPUTE_LIMITS } from '../../compute/protocol.ts'
import { ComputeError } from '../../compute/errors.ts'
import type { ChartData } from './chart-engine'
import type { ChartComputedStatistics } from './chart-derived'
import { supportsChartErrorDirection, supportsChartStatistics } from './chart-catalog'
import { chartNumber } from './chart-data'
import { validateChartOptions } from './chart-options'
import { validateChartNamedSeriesBindings } from './chart-series-selector'
import { formatChartTrendlineEquation, type HistogramBin, type ChartErrorInterval, type WaterfallPoint } from './chart-statistics'
import { assertChartPreparationCurrent, assertChartPreparationMemory, CHART_PREPARATION_CHUNK, consumeChartPreparation,
  type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { estimateChartDecodedMemory } from './chart-compute-memory'

export interface ChartStatisticsPlan {
  task: ChartStatisticsTask | null
  /** Already aggregated/normalized, never the histogram/boxplot output. */
  preparedData: ChartData
  categoryDictionary: string[]
  trendlineCount: number
  errorBarCount: number
}

interface ColumnSource { at(row: number): number | null }

/** Encode only requested numeric series/X/error/total columns. All category
 * strings stay in this owner; the Worker receives dictionary IDs only. */
export function prepareChartStatisticsTask(data: ChartData, chart: ChartModel, context: ChartComputePreparationContext = {}): Promise<ChartStatisticsPlan> {
  return consumeChartPreparation(chartStatisticsTaskSteps(data, chart, context), context)
}

function* chartStatisticsTaskSteps(data: ChartData, chart: ChartModel, context: ChartComputePreparationContext): Generator<void, ChartStatisticsPlan, void> {
  validateChartOptions(chart.options ?? {})
  validateChartNamedSeriesBindings(chart.options, data.series)
  const rowCount = data.categories.length, trendlineCount = chart.options?.trendlines?.length ?? 0, errorBarCount = chart.options?.errorBars?.length ?? 0
  if (!Number.isSafeInteger(rowCount) || rowCount > COMPUTE_LIMITS.maxRows || data.series.length > CHART_COMPUTE_LIMITS.maxSeries
    || trendlineCount > CHART_COMPUTE_LIMITS.maxJobs || errorBarCount > CHART_COMPUTE_LIMITS.maxJobs) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表样本、系列或统计配置数量超出预算')
  }
  if (data.series.some(series => series.values.length !== rowCount || series.xValues && series.xValues.length !== rowCount)) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '图表系列与分类样本数量不一致')
  }
  if (!data.series.length || !rowCount) return { task: null, preparedData: data, categoryDictionary: [], trendlineCount, errorBarCount }
  const columns: ColumnSource[] = [], numericColumns = new Map<readonly (number | null)[], number>(), customColumns = new Map<number, number>()
  const series: ChartStatisticsParams['series'] = [], seriesIndexes = new Map<number, number>(), jobs: ChartStatisticsJob[] = []
  let work = 0
  const addColumn = (source: ColumnSource): number => {
    if (columns.length >= CHART_COMPUTE_LIMITS.maxColumns) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表统计输入列数量超出预算')
    const column = columns.length; columns.push(source); return column
  }
  const addJob = (job: ChartStatisticsJob): void => {
    if (jobs.length >= CHART_COMPUTE_LIMITS.maxJobs) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表分组统计任务数量超出预算')
    jobs.push(job)
  }
  const addNumeric = (values: readonly (number | null)[]): number => {
    const existing = numericColumns.get(values)
    if (existing !== undefined) return existing
    const column = addColumn({ at: row => values[row]! }); numericColumns.set(values, column)
    return column
  }
  const addSeries = (position: number): number => {
    const input = data.series[position]!
    const existing = seriesIndexes.get(position)
    if (existing !== undefined) return existing
    const index = series.length
    series.push({ field: input.field, seriesPosition: position, column: addNumeric(input.values), ...(input.xValues ? { xColumn: addNumeric(input.xValues) } : {}) })
    seriesIndexes.set(position, index)
    return index
  }
  const addCustom = (field: number | undefined): number => {
    if (field === undefined || !Number.isSafeInteger(field) || field < 0 || field >= (data.fields?.length ?? data.records?.[0]?.length ?? 0)
      || !data.records || data.records.length !== rowCount) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '自定义误差字段须与当前图表数据点逐项对应')
    }
    const existing = customColumns.get(field)
    if (existing !== undefined) return existing
    const column = addColumn({ at: row => chartNumber(data.records![row]?.[field]) }); customColumns.set(field, column)
    return column
  }
  if (chart.type === 'histogram' || chart.type === 'pareto') {
    if (data.series.length !== 1) throw new ComputeError('COMPUTE_INVALID_INPUT', '直方图或帕累托图需要一个样本系列')
    const seriesIndex = addSeries(0)
    const histogram = chart.options?.histogram
    addJob(chart.type === 'histogram'
      ? { kind: 'histogram', seriesIndex, slot: 0, ...(histogram?.binning ? { binning: { ...histogram.binning } } : {}),
        ...(histogram?.underflowThreshold === undefined ? {} : { underflowThreshold: histogram.underflowThreshold }),
        ...(histogram?.overflowThreshold === undefined ? {} : { overflowThreshold: histogram.overflowThreshold }) }
      : { kind: 'pareto', seriesIndex, slot: 0, includeZeroValues: chart.options?.pareto?.includeZeroValues ?? false })
  } else if (chart.type === 'boxplot') {
    for (let position = 0; position < data.series.length; position++) addJob({ kind: 'boxplot', seriesIndex: addSeries(position), slot: position,
      quartileMethod: chart.options?.boxplot?.quartileMethod ?? 'inclusive', whiskerMultiplier: chart.options?.boxplot?.whiskerMultiplier ?? 1.5 })
  } else if (chart.type === 'waterfall') {
    for (let position = 0; position < data.series.length; position++) {
      const input = data.series[position]!, totals = new Set<number>()
      for (const role of chart.options?.waterfall?.pointRoles ?? []) {
        if (role.seriesIndex === input.field && role.role === 'total' && role.dataPointIndex < rowCount) totals.add(role.dataPointIndex)
      }
      const totalColumn = totals.size ? addColumn({ at: row => totals.has(row) ? 1 : 0 }) : undefined
      addJob({ kind: 'waterfall', seriesIndex: addSeries(position), slot: position, ...(totalColumn === undefined ? {} : { totalColumn }) })
    }
  }
  for (const [slot, config] of (chart.options?.trendlines ?? []).entries()) {
    for (let position = 0; position < data.series.length; position++) {
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
      if (data.series[position]!.field !== config.seriesIndex || !supportsChartStatistics(chart, config.seriesIndex, position, data.series[position]!.sourceName ?? data.series[position]!.name)) continue
      addJob({ kind: 'trendline', seriesIndex: addSeries(position), slot,
        type: config.type, ...(config.order === undefined ? {} : { order: config.order }), ...(config.period === undefined ? {} : { period: config.period }) })
    }
  }
  for (const [slot, config] of (chart.options?.errorBars ?? []).entries()) {
    if (!supportsChartErrorDirection(chart, config.direction)) continue
    for (let position = 0; position < data.series.length; position++) {
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
      if (data.series[position]!.field !== config.seriesIndex || !supportsChartStatistics(chart, config.seriesIndex, position, data.series[position]!.sourceName ?? data.series[position]!.name)) continue
      addJob({ kind: 'errorBars', seriesIndex: addSeries(position), slot, direction: config.direction, type: config.type, valueType: config.valueType,
        ...(config.value === undefined ? {} : { value: config.value }), ...(config.valueType !== 'cust' ? {} : {
          ...(config.type === 'minus' ? {} : { plusColumn: addCustom(config.plusIndex) }),
          ...(config.type === 'plus' ? {} : { minusColumn: addCustom(config.minusIndex) }),
        }) })
    }
  }
  const plan: ChartStatisticsPlan = { task: null, preparedData: data, categoryDictionary: [], trendlineCount, errorBarCount }
  if (!jobs.length) return plan
  const params: ChartStatisticsParams = { rowCount, columnCount: columns.length, categoryCount: rowCount ? 1 : 0, series, jobs }
  if (!isChartStatisticsParams(params)) throw new ComputeError('COMPUTE_INVALID_INPUT', '图表统计字段、配置或样本维度无效')
  assertChartPreparationMemory(estimateChartTaskMemory(params).estimatedPeakBytes, context)
  const values = new Float64Array(rowCount * columns.length), states = new Uint8Array(values.length), categoryIds = new Uint32Array(rowCount)
  const categoryIndexes = new Map<string, number>()
  for (let row = 0; row < rowCount; row++) {
    const category = data.categories[row]!
    if (typeof category !== 'string') throw new ComputeError('COMPUTE_INVALID_INPUT', '图表分类必须是字符串')
    let categoryId = categoryIndexes.get(category)
    if (categoryId === undefined) { categoryId = categoryIndexes.size; categoryIndexes.set(category, categoryId); plan.categoryDictionary.push(category) }
    categoryIds[row] = categoryId
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  params.categoryCount = categoryIndexes.size
  categoryIndexes.clear()
  for (let column = 0; column < columns.length; column++) {
    const offset = column * rowCount
    for (let row = 0; row < rowCount; row++) {
      const value = columns[column]!.at(row)
      if (value === null) states[offset + row] = CELL_STATE.empty
      else {
        if (!Number.isFinite(value)) throw new ComputeError('COMPUTE_INVALID_INPUT', '图表数值输入必须有限或为空')
        values[offset + row] = value; states[offset + row] = CELL_STATE.number
      }
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
  }
  plan.task = { kind: 'chartStatistics', values, states, categoryIds, params }
  validateChartTask(plan.task)
  assertChartPreparationCurrent(context)
  return plan
}

/** Validate each bulk value/index while decoding it, avoiding another full
 * scan in the Worker protocol guard. The result owns ordinary JS arrays. */
export function decodeChartStatistics(plan: ChartStatisticsPlan, output: unknown, context: ChartComputePreparationContext = {}): Promise<ChartComputedStatistics> {
  return consumeChartPreparation(decodeChartStatisticsSteps(plan, output, context), context)
}

function* decodeChartStatisticsSteps(plan: ChartStatisticsPlan, raw: unknown, context: ChartComputePreparationContext): Generator<void, ChartComputedStatistics, void> {
  if (!plan.task) throw protocolError('没有可解码的图表统计任务')
  const output = readChartStatisticsOutput(raw, plan.task.params)
  if (!output) throw protocolError('图表统计描述符或输出缓冲区无效')
  assertChartPreparationMemory(output.values.byteLength + output.indices.byteLength + estimateChartDecodedMemory(output), context)
  const result: ChartComputedStatistics = { preparedData: plan.preparedData }
  if (plan.trendlineCount) result.trendlines = []
  if (plan.errorBarCount) result.errorBars = []
  let work = 0
  const number = (position: number): number => {
    const value = output.values[position]
    if (value === undefined || !Number.isFinite(value)) throw protocolError('图表统计结果包含非有限值')
    return value
  }
  for (const segment of output.segments) {
    const job = plan.task.params.jobs[segment.jobIndex]!, start = segment.valueOffset, indexStart = segment.indexOffset
    switch (job.kind) {
      case 'histogram': {
        const bins: HistogramBin[] = []
        let count = 0, previousUpper: number | undefined
        for (let offset = 0; offset < segment.valueLength; offset += 2) {
          const lower = number(start + offset), upper = number(start + offset + 1), kind = output.indices[indexStart + offset]!, samples = output.indices[indexStart + offset + 1]!
          count += samples
          if (samples > segment.validSampleCount || count > segment.validSampleCount) throw protocolError('直方图分箱计数超出样本数量')
          if (kind === 0) {
            if (offset !== 0 || job.underflowThreshold === undefined || lower !== 0 || upper !== job.underflowThreshold) throw protocolError('直方图下溢边界无效')
            bins.push({ kind: 'underflow', upperBound: upper, count: samples }); previousUpper = upper
          } else if (kind === 1) {
            if (lower >= upper || previousUpper !== undefined && lower !== previousUpper) throw protocolError('直方图普通分箱边界不连续')
            bins.push({ kind: 'regular', lowerBound: lower, upperBound: upper, count: samples }); previousUpper = upper
          } else if (kind === 2) {
            if (offset + 2 !== segment.valueLength || job.overflowThreshold === undefined || lower !== job.overflowThreshold || upper !== 0
              || previousUpper !== undefined && previousUpper > lower) throw protocolError('直方图上溢边界无效')
            bins.push({ kind: 'overflow', lowerBound: lower, count: samples })
          } else throw protocolError('直方图分箱类型无效')
          if (++work % CHART_PREPARATION_CHUNK === 0) yield
        }
        if (count !== segment.validSampleCount || segment.available && (job.underflowThreshold !== undefined && bins[0]?.kind !== 'underflow'
          || job.overflowThreshold !== undefined && bins.at(-1)?.kind !== 'overflow')) throw protocolError('直方图分箱没有覆盖有效样本')
        result.histogram = { validSampleCount: count, bins }
        break
      }
      case 'boxplot': {
        let statistics: NonNullable<ChartComputedStatistics['boxplots']>[number]['statistics']
        if (segment.available) {
          const min = number(start), q1 = number(start + 1), median = number(start + 2), q3 = number(start + 3), max = number(start + 4), outliers: number[] = []
          if (q1 > median || median > q3 || min > max) throw protocolError('箱线图四分位或须线顺序无效')
          for (let offset = 5; offset < segment.valueLength; offset++) {
            const value = number(start + offset)
            if (outliers.length && outliers.at(-1)! > value || value >= min && value <= max) throw protocolError('箱线图异常点范围或顺序无效')
            outliers.push(value)
            if (++work % CHART_PREPARATION_CHUNK === 0) yield
          }
          if (outliers.length > segment.validSampleCount) throw protocolError('箱线图异常点数量无效')
          statistics = { min, q1, median, q3, max, outliers }
        }
        ;(result.boxplots ??= []).push({ field: segment.field, seriesPosition: segment.seriesPosition, statistics })
        break
      }
      case 'pareto': {
        const categories: string[] = [], values: number[] = [], cumulativePercent: number[] = [], seen = new Set<number>()
        let positive = false
        for (let offset = 0; offset < segment.indexLength; offset++) {
          const categoryId = output.indices[indexStart + offset]!, category = plan.categoryDictionary[categoryId]
          const value = number(start + offset * 2), percent = number(start + offset * 2 + 1)
          if (category === undefined || seen.has(categoryId) || value < 0 || !job.includeZeroValues && value === 0 || values.length && values.at(-1)! < value
            || percent < 0 || percent > 100 || cumulativePercent.length && cumulativePercent.at(-1)! > percent) throw protocolError('帕累托分类、权重或累计比例无效')
          seen.add(categoryId); categories.push(category); values.push(value); cumulativePercent.push(percent); positive ||= value > 0
          if (++work % CHART_PREPARATION_CHUNK === 0) yield
        }
        if (positive && cumulativePercent.at(-1) !== 100) throw protocolError('帕累托累计比例没有到达百分之百')
        result.pareto = { categories, values, cumulativePercent }
        break
      }
      case 'waterfall': {
        const points: WaterfallPoint[] = []
        let previous = 0
        for (let offset = 0; offset < segment.indexLength / 2; offset++) {
          const categoryIndex = output.indices[indexStart + offset * 2]!, total = output.indices[indexStart + offset * 2 + 1]!
          const startValue = number(start + offset * 3), end = number(start + offset * 3 + 1), value = number(start + offset * 3 + 2)
          if (categoryIndex !== offset || total > 1 || startValue !== (total ? 0 : previous) || end !== (total ? value : startValue + value)) throw protocolError('瀑布图数据点或累计基线无效')
          points.push({ categoryIndex, start: startValue, end, value, role: total ? 'total' : 'change' }); previous = end
          if (++work % CHART_PREPARATION_CHUNK === 0) yield
        }
        ;(result.waterfalls ??= []).push({ field: segment.field, seriesPosition: segment.seriesPosition, points })
        break
      }
      case 'trendline': {
        if (segment.slot >= plan.trendlineCount) throw protocolError('趋势线结果不属于当前配置槽位')
        const points: Array<[number, number]> = []
        let previousX: number | undefined
        for (let offset = 0; offset < segment.valueLength; offset += 2) {
          const x = number(start + offset), y = number(start + offset + 1)
          if (job.type !== 'movingAvg' && previousX !== undefined && x < previousX) throw protocolError('趋势线预测点顺序无效')
          points.push([x, y]); previousX = x
          if (++work % CHART_PREPARATION_CHUNK === 0) yield
        }
        result.trendlines!.push({ configIndex: segment.slot, seriesPosition: segment.seriesPosition,
          statistics: { points, ...(segment.fit ? { formula: formatChartTrendlineEquation(job.type, segment.fit) } : {}),
            ...(segment.r2 === undefined ? {} : { r2: segment.r2 }) } })
        break
      }
      case 'errorBars': {
        if (segment.slot >= plan.errorBarCount) throw protocolError('误差线结果不属于当前配置槽位')
        const intervals: ChartErrorInterval[] = []
        let previous = -1
        for (let offset = 0; offset < segment.indexLength; offset++) {
          const categoryIndex = output.indices[indexStart + offset]!, x = number(start + offset * 4), y = number(start + offset * 4 + 1)
          const minus = number(start + offset * 4 + 2), plus = number(start + offset * 4 + 3), target = job.direction === 'x' ? x : y
          if (categoryIndex <= previous || categoryIndex >= plan.preparedData.categories.length || minus < 0 || plus < 0
            || job.type === 'plus' && minus !== 0 || job.type === 'minus' && plus !== 0
            || job.type !== 'plus' && !Number.isFinite(target - minus) || job.type !== 'minus' && !Number.isFinite(target + plus)) throw protocolError('误差线数据点、方向或数值范围无效')
          intervals.push({ categoryIndex, x, y, minus, plus }); previous = categoryIndex
          if (++work % CHART_PREPARATION_CHUNK === 0) yield
        }
        result.errorBars!.push({ configIndex: segment.slot, seriesPosition: segment.seriesPosition, intervals })
        break
      }
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return result
}

function protocolError(message: string): ComputeError { return new ComputeError('COMPUTE_PROTOCOL_MISMATCH', message) }
