import type { ChartData } from './chart-engine'
import { resolveChartType, supportsChartStatistics } from './chart-catalog'
import { chartBubbleMapping, chartCandlestickMapping, chartEdgeMapping, chartField, chartHierarchyValueField, chartNumber, chartValueField,
  readSpecialChartData, type ChartDataBinding, type ChartRecords } from './chart-data'

/** Source defaults are explicit for pivot output; ordinary sheets infer numeric columns. */
export function projectChartRecords(source: ChartRecords, chart: ChartDataBinding, availableValueFields?: readonly number[]): ChartData {
  const descriptor = resolveChartType(chart.type), shape = descriptor.dataShape
  const categories = chart.options?.mapping?.categoryIndexes
    ?? (chart.categoryField === 0 && source.sourceCategoryFields?.length ? source.sourceCategoryFields : [chart.categoryField])
  const visible = source.visibleFields === undefined ? undefined : new Set(source.visibleFields)
  const numericFields = availableValueFields !== undefined ? [...availableValueFields] : source.sourceValueFields !== undefined ? [...source.sourceValueFields]
    : source.fields.map((_field, index) => index).filter(index => source.records.some(record => chartNumber(record[index]) !== null))
  const defaults = numericFields.filter(field => {
    if (visible && !visible.has(field)) return false
    if (shape === 'distribution') return true
    if (shape === 'xy' || shape === 'bubble') return field !== chart.categoryField
    return !categories.includes(field)
  })
  const authoredFields = source.syntheticValues && chart.valueFields.includes(0) ? [] : chart.valueFields
  const singleValue = ['pie', 'funnel', 'wordCloud', 'pareto', 'histogram'].includes(descriptor.family)
  const fields = authoredFields.length ? authoredFields : singleValue ? defaults.slice(0, 1) : defaults
  for (const field of chart.valueFields) chartField(source, field, '数值字段')
  const binding: ChartDataBinding = shape === 'edge' ? chart : { ...chart, valueFields: fields,
    ...(shape === 'hierarchy' && !chart.options?.mapping?.hierarchyIndexes && source.sourceCategoryFields?.length
      ? { options: { ...chart.options, mapping: { ...chart.options?.mapping, hierarchyIndexes: source.sourceCategoryFields } } } : {}) }

  const roles = primaryFields(source, binding, categories)
  for (const field of roles) chartField(source, field, '图表字段')
  if (shape !== 'edge') fields.forEach(field => chartValueField(source, field, '数值字段'))
  validateChartAuxiliaryReferences(source, binding, fields)
  const values = visible ? fields.filter(field => visible.has(field)) : fields
  // Keep raw records in the settings source. The plotted records hide auxiliary
  // fields as well, so custom errors cannot read a hidden column as zero.
  const plotted = visible && visible.size < source.fields.length ? { ...source,
    records: source.records.map(record => record.map((value, field) => visible.has(field) ? value : null)) } : source
  const empty: ChartData = { ...plotted, categories: [], series: [] }
  if (visible && roles.some(field => !visible.has(field))) return empty
  if ((shape === 'distribution' || shape === 'category' || shape === 'heatmap') && !values.length) return empty
  const special = readSpecialChartData(plotted, shape === 'edge' ? binding : { ...binding, valueFields: values })
  if (special) return special
  if (shape === 'xy') return scatterData(plotted, binding, fields, visible)
  for (const field of categories) chartField(source, field, '分类字段')
  for (const field of values) chartField(source, field, '数值字段')
  return { ...plotted,
    categories: plotted.records.map(record => categories.map(field => String(record[field] ?? '')).join(' / ')),
    categoryValues: plotted.records.map(record => categories.length === 1 ? record[categories[0]!] ?? null : categories.map(field => String(record[field] ?? '')).join(' / ')),
    series: values.map(field => ({ field, name: source.fields[field] || `Series ${field + 1}`, values: plotted.records.map(record => chartNumber(record[field])) })),
  }
}

function primaryFields(source: ChartRecords, chart: ChartDataBinding, categories: number[]): number[] {
  const shape = resolveChartType(chart.type).dataShape
  if (shape === 'bubble') return Object.values(chartBubbleMapping(source, chart)).filter((field): field is number => field !== undefined)
  if (shape === 'edge') return Object.values(chartEdgeMapping(source, chart)).filter((field): field is number => field !== undefined)
  if (shape === 'ohlc') return [chart.categoryField, ...Object.values(chartCandlestickMapping(source, chart)).filter((field): field is number => field !== undefined)]
  if (shape === 'hierarchy') return [...(chart.options?.mapping?.hierarchyIndexes ?? [chart.categoryField]), chartHierarchyValueField(source, chart)]
  if (shape === 'gauge') {
    const mapping = chart.options?.mapping?.gauge ?? { valueIndex: chart.valueFields[0] ?? (source.fields.length > 1 ? 1 : 0),
      labelIndex: source.fields.length > 1 ? chart.categoryField : undefined }
    return Object.values(mapping).filter((field): field is number => field !== undefined)
  }
  return shape === 'xy' || shape === 'distribution' ? [] : categories
}

function validateChartAuxiliaryReferences(source: ChartRecords, chart: ChartDataBinding, valueFields: number[]): void {
  for (const trend of chart.options?.trendlines ?? []) chartField(source, trend.seriesIndex, '趋势线系列字段')
  for (const error of chart.options?.errorBars ?? []) {
    chartField(source, error.seriesIndex, '误差线系列字段')
    if (error.valueType !== 'cust' || !supportsChartStatistics(chart) || !valueFields.includes(error.seriesIndex)) continue
    for (const field of [error.type === 'minus' ? undefined : error.plusIndex, error.type === 'plus' ? undefined : error.minusIndex]) {
      if (field !== undefined) chartValueField(source, field, '自定义误差字段')
    }
  }
}

export function cloneChartSeriesCategories(mapping: Record<number, number[]>): Record<number, number[]> {
  return Object.fromEntries(Object.entries(mapping).map(([field, categories]) => {
    const valueField = Number(field)
    if (!Number.isSafeInteger(valueField) || valueField < 0 || valueField > 1_048_575) throw new RangeError('散点图 Y 字段须为非负字段索引。')
    if (!Array.isArray(categories) || categories.length !== 1 || !Number.isSafeInteger(categories[0]) || categories[0]! < 0 || categories[0]! > 1_048_575) {
      throw new RangeError('散点图的每个 Y 系列须对应一个 X 字段。')
    }
    return [field, [...categories]]
  }))
}

function scatterData(source: ChartRecords, chart: ChartDataBinding, selectedFields: number[], visible?: ReadonlySet<number>): ChartData {
  chartValueField(source, chart.categoryField, 'X 字段')
  const mapping = cloneChartSeriesCategories(chart.seriesToCategory ?? {})
  const xFields = new Set([chart.categoryField, ...Object.values(mapping).flat()])
  const fields = chart.valueFields.length ? chart.valueFields : selectedFields.filter(field => !xFields.has(field))
  for (const [valueField, xField] of Object.entries(mapping)) {
    chartValueField(source, Number(valueField), 'Y 字段'); chartValueField(source, xField[0]!, 'X 字段')
    if (!fields.includes(Number(valueField))) throw new RangeError('散点图的独立 X 字段须对应已选择的 Y 系列。')
  }
  return { ...source,
    categories: source.records.map(record => String(record[chart.categoryField] ?? '')),
    series: fields.filter(field => !visible || visible.has(field) && visible.has(mapping[field]?.[0] ?? chart.categoryField)).map(field => {
      chartValueField(source, field, 'Y 字段')
      const xField = mapping[field]?.[0] ?? chart.categoryField
      return { field, name: source.fields[field] || `Series ${field + 1}`,
        values: source.records.map(record => chartNumber(record[field])), xValues: source.records.map(record => chartNumber(record[xField])) }
    }),
  }
}
