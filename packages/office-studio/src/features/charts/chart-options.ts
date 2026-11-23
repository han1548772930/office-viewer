import type { ChartModel, ChartSemanticOptions, ChartSeriesOptions } from '../../domain/model'

const DEFAULT_COLORS = ['#2563eb', '#dc2626', '#059669', '#d97706', '#7c3aed', '#0891b2']

export function cloneChartOptions(options: ChartSemanticOptions | undefined): ChartSemanticOptions | undefined {
  if (options === undefined) return undefined
  validateChartOptions(options)
  return JSON.parse(JSON.stringify(options)) as ChartSemanticOptions
}

/** Groups omitted from a patch are retained. An explicitly supplied collection replaces that collection. */
export function mergeChartOptions(current: ChartSemanticOptions | undefined, patch: ChartSemanticOptions | undefined): ChartSemanticOptions | undefined {
  if (patch === undefined) return cloneChartOptions(current)
  const merged: ChartSemanticOptions = { ...cloneChartOptions(current), ...cloneChartOptions(patch) }
  for (const key of ['appearance', 'mapping', 'threeD', 'bar', 'area', 'pie', 'radar', 'bubble', 'funnel', 'relation', 'waterfall', 'pareto', 'histogram', 'boxplot', 'candlestick', 'heatmap', 'treemap', 'sunburst', 'gauge', 'chord', 'wordCloud'] as const) {
    if (current?.[key] && patch[key]) Object.assign(merged, { [key]: { ...current[key], ...patch[key] } })
  }
  return cloneChartOptions(merged)
}

export function chartPalette(chart: ChartModel, fallback: readonly string[] = DEFAULT_COLORS): string[] {
  return [...(chart.palette?.length ? chart.palette : fallback.length ? fallback : DEFAULT_COLORS)]
}

export function chartSeriesStyle(chart: ChartModel, field: number, index: number, palette: readonly string[] = DEFAULT_COLORS): ChartSeriesOptions {
  const shared = chart.options?.allSeriesStyle, specific = chart.options?.series?.[field]
  return {
    color: palette[index % palette.length], fillOpacity: 1,
    ...shared, ...specific,
    border: { ...shared?.border, ...specific?.border },
    point: { ...shared?.point, ...specific?.point },
    label: { ...shared?.label, ...specific?.label },
  }
}

export function validateChartOptions(options: ChartSemanticOptions): void {
  let nodes = 0
  function visit(value: unknown, path: string, depth: number): void {
    if (++nodes > 200_000 || depth > 12) throw new RangeError('图表配置超出允许的大小或嵌套深度。')
    if (value === undefined || value === null || typeof value === 'string' || typeof value === 'boolean') return
    if (typeof value === 'number') { if (!Number.isFinite(value)) throw new RangeError(`${path}须为有限数值。`); return }
    if (Array.isArray(value)) { value.forEach((entry, index) => visit(entry, `${path}[${index}]`, depth + 1)); return }
    if (typeof value !== 'object' || Object.getPrototypeOf(value) !== Object.prototype) throw new TypeError(`${path}须为可序列化的普通配置对象。`)
    for (const [key, entry] of Object.entries(value)) {
      if (key === '__proto__' || key === 'constructor' || key === 'prototype') throw new TypeError('图表配置包含无效属性。')
      visit(entry, `${path}.${key}`, depth + 1)
      if (entry === undefined) continue
      if (['seriesIndex', 'dataPointIndex', 'xIndex', 'yIndex', 'sizeIndex', 'labelIndex', 'sourceIndex', 'targetIndex', 'valueIndex', 'openIndex', 'highIndex', 'lowIndex', 'closeIndex', 'plusIndex', 'minusIndex'].includes(key)) integer(entry, 0, 1_048_575, `${path}.${key}`)
      if (['opacity', 'fillOpacity', 'ribbonOpacity'].includes(key)) number(entry, 0, 1, `${path}.${key}`)
      if (['fontSize', 'minFontSize', 'maxFontSize'].includes(key)) number(entry, 1, 256, `${path}.${key}`)
      if (key === 'width' && path.endsWith('binning')) number(entry, Number.MIN_VALUE, Number.MAX_VALUE, `${path}.${key}`)
    }
  }
  visit(options, '图表', 0)
  if (options.invalidValueStrategy !== undefined && !['zero', 'break', 'link'].includes(options.invalidValueStrategy)) throw new TypeError('无效的空值策略。')
  for (const fields of [options.mapping?.categoryIndexes, options.mapping?.hierarchyIndexes]) {
    if (fields && (!fields.length || new Set(fields).size !== fields.length)) throw new RangeError('分类或层级字段不能为空或重复。')
    fields?.forEach(field => integer(field, 0, 1_048_575, '图表字段'))
  }
  if (options.histogram?.binning?.mode === 'count') integer(options.histogram.binning.count, 1, 10_000, '直方图箱数')
  if (options.histogram?.binning && !['count', 'width'].includes(options.histogram.binning.mode)) throw new TypeError('无效的直方图分箱方式。')
  if (options.histogram?.binGap !== undefined) number(options.histogram.binGap, 0, 0.9, '直方图箱间距')
  if (options.histogram?.underflowThreshold !== undefined && options.histogram.overflowThreshold !== undefined && options.histogram.underflowThreshold >= options.histogram.overflowThreshold) throw new RangeError('下溢阈值须小于上溢阈值。')
  if (options.gauge?.min !== undefined && options.gauge.max !== undefined && options.gauge.min >= options.gauge.max) throw new RangeError('仪表最大刻度须大于最小刻度。')
  if (options.bubble?.scale !== undefined) number(options.bubble.scale, Number.MIN_VALUE, 10_000, '气泡比例')
  if (options.boxplot?.whiskerMultiplier !== undefined) number(options.boxplot.whiskerMultiplier, 0, 100, '箱线图须线倍数')
  if (options.threeD?.view?.rotX !== undefined) number(options.threeD.view.rotX, -90, 90, '三维俯仰角')
  if (options.threeD?.view?.rotY !== undefined) number(options.threeD.view.rotY, 0, 360, '三维旋转角')
  if (options.pie?.doughnutHole !== undefined) number(options.pie.doughnutHole, 0, 0.95, '圆环内径比例')
  if (options.pie?.secondaryPlot?.split?.type === 'position') integer(options.pie.secondaryPlot.split.count, 1, 100_000, '副图项目数')
  options.trendlines?.forEach(line => {
    if (!['exp', 'linear', 'log', 'movingAvg', 'poly', 'pow'].includes(line.type)) throw new TypeError('无效的趋势线类型。')
    if (line.type === 'poly') integer(line.order ?? 2, 2, 6, '多项式阶数')
    if (line.type === 'movingAvg') integer(line.period ?? 2, 2, 100_000, '移动平均周期')
  })
  options.errorBars?.forEach(error => {
    if (!['x', 'y'].includes(error.direction) || !['both', 'plus', 'minus'].includes(error.type) || !['fixedVal', 'percentage', 'stdDev', 'stdErr', 'cust'].includes(error.valueType)) throw new TypeError('无效的误差线配置。')
    if (error.value !== undefined) number(error.value, 0, Number.MAX_VALUE, '误差值')
    if (error.valueType === 'cust' && error.plusIndex === undefined && error.minusIndex === undefined) throw new RangeError('自定义误差线须选择正或负误差字段。')
  })
}

function integer(value: unknown, min: number, max: number, label: string): void {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < min || value > max) throw new RangeError(`${label}须为 ${min} 到 ${max} 的整数。`)
}
function number(value: unknown, min: number, max: number, label: string): void {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) throw new RangeError(`${label}须位于 ${min} 到 ${max}。`)
}
