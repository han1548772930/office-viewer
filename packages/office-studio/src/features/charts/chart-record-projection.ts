import type { ChartData } from './chart-engine'
import { resolveChartType, supportsChartStatistics } from './chart-catalog'
import { chartBubbleMapping, chartCandlestickMapping, chartEdgeMapping, chartField, chartHierarchyValueField, chartNumber, chartValueField,
  readSpecialChartDataSteps, type ChartDataBinding, type ChartRecords } from './chart-data'
import { assertChartPreparationMemory, CHART_PREPARATION_CHUNK, consumeChartPreparation, consumeChartPreparationSync,
  type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { chartMapSteps, chartNumericFieldSteps, chartProjectionSourceBytes } from './chart-data-preparation'
import { chartHierarchyChartDataSteps, chartHierarchyFields, chartHierarchyNumber, chartHierarchyResultAsync } from './chart-hierarchy-data'
import { chartChordChartDataSteps, chartChordResultAsync } from './chart-chord-data'
import { chartRelationSourceFormat } from './chart-relation-data'

interface ChartRecordProjection {
  source: ChartRecords; binding: ChartDataBinding; categories: number[]; fields: number[]; values: number[]
  visible?: ReadonlySet<number>; empty?: ChartData
}

/** Source defaults are explicit for pivot output; ordinary sheets infer numeric columns. */
export function projectChartRecords(source: ChartRecords, chart: ChartDataBinding, availableValueFields?: readonly number[]): ChartData {
  return consumeChartPreparationSync(projectChartRecordSteps(source, chart, availableValueFields))
}

export async function projectChartRecordsAsync(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Promise<ChartData> {
  const plan = await consumeChartPreparation(prepareChartRecordProjectionSteps(source, chart, undefined, context), context)
  if (plan.empty) return plan.empty
  if (chart.type === 'chord') {
    const result = await chartChordResultAsync(plan.source, plan.binding, context)
    return consumeChartPreparation(chartChordChartDataSteps(plan.source, result), context)
  }
  if (resolveChartType(chart.type).dataShape === 'hierarchy') {
    const binding = { ...plan.binding, valueFields: plan.values }
    const valueField = chartHierarchyValueField(plan.source, binding)
    const result = await chartHierarchyResultAsync(plan.source, binding, context)
    return consumeChartPreparation(chartHierarchyChartDataSteps(plan.source, result, valueField), context)
  }
  return consumeChartPreparation(projectPreparedChartRecordSteps(plan, context), context)
}

export function* projectChartRecordSteps(source: ChartRecords, chart: ChartDataBinding, availableValueFields?: readonly number[],
  context: ChartComputePreparationContext = {}): Generator<void, ChartData, void> {
  const plan = yield* prepareChartRecordProjectionSteps(source, chart, availableValueFields, context)
  return yield* projectPreparedChartRecordSteps(plan, context)
}

function* prepareChartRecordProjectionSteps(source: ChartRecords, chart: ChartDataBinding, availableValueFields?: readonly number[],
  context: ChartComputePreparationContext = {}): Generator<void, ChartRecordProjection, void> {
  assertChartPreparationMemory(chartProjectionSourceBytes(source.records.length, source.fields.length), context)
  const descriptor = resolveChartType(chart.type), shape = descriptor.dataShape
  const sourceNumericFields = shape === 'hierarchy' || chart.type === 'chord' ? yield* chartNumericFieldSteps(source.records, source.fields.length, chartHierarchyNumber)
    : source.sourceNumericFields ?? (yield* chartNumericFieldSteps(source.records, source.fields.length, chartNumber))
  source = { ...source, sourceNumericFields }
  const categories = chart.options?.mapping?.categoryIndexes
    ?? (chart.categoryField === 0 && source.sourceCategoryFields?.length ? source.sourceCategoryFields : [chart.categoryField])
  const visible = source.visibleFields === undefined ? undefined : new Set(source.visibleFields)
  const numericFields = availableValueFields !== undefined ? [...availableValueFields] : source.sourceValueFields !== undefined ? [...source.sourceValueFields]
    : [...sourceNumericFields]
  const defaults = numericFields.filter(field => {
    if (visible && !visible.has(field)) return false
    if (shape === 'distribution') return true
    if (shape === 'xy' || shape === 'bubble') return field !== chart.categoryField
    return !(shape === 'hierarchy' ? chartHierarchyFields(source, chart) : categories).includes(field)
  })
  const authoredFields = source.syntheticValues && chart.valueFields.includes(0) ? [] : chart.valueFields
  const singleValue = shape === 'hierarchy' || ['pie', 'funnel', 'wordCloud', 'pareto', 'histogram'].includes(descriptor.family)
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
  let plotted = source
  if (visible && visible.size < source.fields.length) {
    const records: ChartRecords['records'] = []
    let work = 0
    for (const record of source.records) {
      records.push(yield* chartMapSteps(record, (value, field) => visible.has(field) ? value : null))
      work += record.length + 1
      if (work >= CHART_PREPARATION_CHUNK) { work = 0; yield }
    }
    const formattedRecords: ChartRecords['formattedRecords'] = source.formattedRecords ? [] : undefined
    if (formattedRecords) for (const record of source.formattedRecords!) {
      formattedRecords.push(yield* chartMapSteps(record, (value, field) => visible.has(field) ? value : null))
      work += record.length + 1
      if (work >= CHART_PREPARATION_CHUNK) { work = 0; yield }
    }
    plotted = { ...source, records, ...(formattedRecords ? { formattedRecords } : {}) }
  }
  const empty: ChartData = { ...plotted, categories: [], series: [] }
  const hidden = Boolean(visible && roles.some(field => !visible.has(field)))
    || (shape === 'distribution' || shape === 'category' || shape === 'heatmap') && !values.length
  return { source: plotted, binding, categories, fields, values, visible, ...(hidden ? { empty } : {}) }
}

function* projectPreparedChartRecordSteps(plan: ChartRecordProjection,
  context: ChartComputePreparationContext): Generator<void, ChartData, void> {
  if (plan.empty) return plan.empty
  const { source: plotted, binding, fields, values, categories, visible } = plan
  const source = plotted, shape = resolveChartType(binding.type).dataShape
  const special = yield* readSpecialChartDataSteps(plotted, shape === 'edge' ? binding : { ...binding, valueFields: values }, context)
  if (special) return special
  if (shape === 'xy') return yield* scatterDataSteps(plotted, binding, fields, visible)
  for (const field of categories) chartField(source, field, '分类字段')
  for (const field of values) chartField(source, field, '数值字段')
  const labels: string[] = [], categoryValues: NonNullable<ChartData['categoryValues']> = []
  let work = 0
  for (const record of plotted.records) {
    const parts = yield* chartMapSteps(categories, field => String(record[field] ?? ''))
    const label = parts.join(' / ')
    labels.push(label); categoryValues.push(categories.length === 1 ? record[categories[0]!] ?? null : label)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const series: ChartData['series'] = []
  for (const field of values) {
    const seriesValues = yield* chartMapSteps(plotted.records, record => chartNumber(record[field]))
    const formattedValues = shape === 'heatmap' && plotted.formattedRecords
      ? yield* chartMapSteps(plotted.formattedRecords, record => record[field] ?? null) : undefined
    series.push({ field, name: source.fields[field] || `Series ${field + 1}`, values: seriesValues, ...(formattedValues ? { formattedValues } : {}) })
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return { ...plotted, categories: labels, categoryValues, series }
}

function primaryFields(source: ChartRecords, chart: ChartDataBinding, categories: number[]): number[] {
  const shape = resolveChartType(chart.type).dataShape
  if (chart.type === 'relation' && chartRelationSourceFormat(chart, source) === 'coWordMatrix') return source.fields.map((_field, index) => index)
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

function* scatterDataSteps(source: ChartRecords, chart: ChartDataBinding, selectedFields: number[], visible?: ReadonlySet<number>): Generator<void, ChartData, void> {
  chartValueField(source, chart.categoryField, 'X 字段')
  const mapping = cloneChartSeriesCategories(chart.seriesToCategory ?? {})
  const xFields = new Set([chart.categoryField, ...Object.values(mapping).flat()])
  const fields = chart.valueFields.length ? chart.valueFields : selectedFields.filter(field => !xFields.has(field))
  for (const [valueField, xField] of Object.entries(mapping)) {
    chartValueField(source, Number(valueField), 'Y 字段'); chartValueField(source, xField[0]!, 'X 字段')
    if (!fields.includes(Number(valueField))) throw new RangeError('散点图的独立 X 字段须对应已选择的 Y 系列。')
  }
  const categories = yield* chartMapSteps(source.records, record => String(record[chart.categoryField] ?? ''))
  const series: ChartData['series'] = []
  for (const field of fields) {
    if (visible && (!visible.has(field) || !visible.has(mapping[field]?.[0] ?? chart.categoryField))) continue
    chartValueField(source, field, 'Y 字段')
    const xField = mapping[field]?.[0] ?? chart.categoryField
    const values = yield* chartMapSteps(source.records, record => chartNumber(record[field]))
    const xValues = yield* chartMapSteps(source.records, record => chartNumber(record[xField]))
    series.push({ field, name: source.fields[field] || `Series ${field + 1}`, values, xValues })
  }
  return { ...source, categories, series }
}
