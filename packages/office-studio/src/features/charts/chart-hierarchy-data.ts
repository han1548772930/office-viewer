import type { ChartData } from './chart-engine'
import type { ChartStatisticsTask } from '../../compute/chart-protocol'
import { CHART_COMPUTE_LIMITS } from '../../compute/chart-protocol'
import { estimateChartTaskMemory } from '../../compute/chart-tasks'
import { COMPUTE_LIMITS } from '../../compute/protocol'
import { ComputeError } from '../../compute/errors'
import { chartField, chartValueField, type ChartDataBinding, type ChartHierarchyNode, type ChartRecords } from './chart-data'
import { assertChartPreparationMemory, CHART_PREPARATION_CHUNK, consumeChartPreparation, consumeChartPreparationSync,
  type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { chartMapSteps, chartProjectionSourceBytes } from './chart-data-preparation'
import { chartHierarchySourceFormatterSteps } from './chart-hierarchy-format'
import { decodeChartHierarchySteps, executeChartHierarchyTaskSteps } from './chart-hierarchy-compute'

export interface ChartHierarchyData { valueFieldName: string; nodes: ChartHierarchyNode[]; paletteOrder: string[]; diagnostic?: string }
export interface ChartHierarchyPlan {
  task: ChartStatisticsTask; keys: string[]; labels: string[]; valueField: number; valueFieldName: string
  formatter: (value: number) => string; retainedBytes: number
}

export function chartHierarchyNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

/** Whitespace is an authored name. Only null/undefined and the empty string
 * participate in scoped carry, as in fixed Va. */
export const chartHierarchyBlank = (value: unknown): boolean => value == null || value === ''

export function chartHierarchyTokenKey(value: unknown): string {
  if (value instanceof Date) return `date:${value.getTime()}`
  if (typeof value === 'number') return `number:${Object.is(value, -0) ? '-0' : String(value)}`
  if (typeof value === 'string') return `string:${JSON.stringify(value)}`
  if (typeof value === 'boolean') return `boolean:${String(value)}`
  return `${typeof value}:${String(value)}`
}

export function chartHierarchyFields(source: ChartRecords, chart: ChartDataBinding): number[] {
  return chart.options?.mapping?.hierarchyIndexes ?? (source.sourceCategoryFields?.length ? source.sourceCategoryFields : [chart.categoryField])
}

export function chartHierarchyValueField(source: ChartRecords, chart: ChartDataBinding): number {
  if (chart.valueFields.length > 1) throw new RangeError('层级图须选择且仅选择一个数值权重字段。')
  if (chart.valueFields.length) return chart.valueFields[0]!
  const hierarchy = chartHierarchyFields(source, chart)
  const numeric = source.sourceValueFields ?? source.sourceNumericFields ?? source.fields.map((_field, index) => index)
    .filter(index => source.records.some(row => chartHierarchyNumber(row[index]) !== null))
  return numeric.find(index => !hierarchy.includes(index)) ?? source.fields.findIndex((_name, index) => !hierarchy.includes(index))
}

/** Strings stay in dictionaries on the owner thread. Only typed key/label IDs,
 * numeric values and state bytes cross the Worker/WASM boundary. */
export function* createChartHierarchyPlanSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartHierarchyPlan, void> {
  const fields = chartHierarchyFields(source, chart), valueField = chartHierarchyValueField(source, chart)
  if (!fields.length || new Set(fields).size !== fields.length) throw new RangeError('层级字段须按从根到叶顺序选择且不能重复。')
  for (const field of fields) chartField(source, field, '层级字段')
  chartValueField(source, valueField, '层级权重字段')
  if (fields.includes(valueField)) throw new RangeError('层级名称字段不能同时作为权重字段。')
  const rows = source.records.length, columns = fields.length * 2 + 1
  if (rows > COMPUTE_LIMITS.maxRows || columns > CHART_COMPUTE_LIMITS.maxColumns || rows * columns > COMPUTE_LIMITS.maxElements) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '层级图完整数据的列数或样本数超出计算预算。')
  }
  const params: ChartStatisticsTask['params'] = { rowCount: rows, columnCount: columns, categoryCount: rows ? 1 : 0,
    series: [{ field: valueField, seriesPosition: 0, column: 0 }], jobs: [{ kind: 'hierarchy', seriesIndex: 0, slot: 0,
      hierarchyColumns: fields.map((_field, depth) => depth + 1), labelColumns: fields.map((_field, depth) => fields.length + depth + 1), keyCount: 0, labelCount: 0 }] }
  const job = params.jobs[0]!
  if (job.kind !== 'hierarchy') throw new ComputeError('COMPUTE_INVALID_INPUT', '层级图任务类型无效。')
  const sourceBytes = chartProjectionSourceBytes(rows, source.fields.length)
  const peak = estimateChartTaskMemory(params).estimatedPeakBytes
  assertChartPreparationMemory(sourceBytes + peak, context)
  const values = new Float64Array(rows * columns), states = new Uint8Array(values.length), categoryIds = new Uint32Array(rows)
  const keys: string[] = [], labels: string[] = [], keyIndex = new Map<string, number>(), labelIndex = new Map<string, number>()
  let retainedBytes = sourceBytes, work = 0
  const dictionary = (text: string, entries: string[], index: Map<string, number>): number => {
    const previous = index.get(text)
    if (previous !== undefined) return previous
    retainedBytes += 96 + text.length * 2
    assertChartPreparationMemory(retainedBytes + peak, context)
    const next = entries.length; entries.push(text); index.set(text, next); return next
  }
  for (let row = 0; row < rows; row++) {
    const record = source.records[row]!, weight = chartHierarchyNumber(record[valueField])
    if (weight !== null) { values[row] = weight; states[row] = 1 }
    for (let depth = 0; depth < fields.length; depth++) {
      const field = fields[depth]!, raw = record[field]
      if (!chartHierarchyBlank(raw)) {
        const key = dictionary(chartHierarchyTokenKey(raw), keys, keyIndex)
        const label = dictionary(source.formattedRecords?.[row]?.[field] ?? String(raw), labels, labelIndex)
        const keyOffset = (depth + 1) * rows + row, labelOffset = (fields.length + depth + 1) * rows + row
        values[keyOffset] = key; states[keyOffset] = 1; values[labelOffset] = label; states[labelOffset] = 1
      }
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  job.keyCount = keys.length; job.labelCount = labels.length
  const formatter = yield* chartHierarchySourceFormatterSteps(source, valueField)
  return { task: { kind: 'chartStatistics', values, states, categoryIds, params }, keys, labels, valueField,
    valueFieldName: source.fields[valueField] ?? '', formatter, retainedBytes }
}

export function* chartHierarchyResultSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartHierarchyData, void> {
  const plan = yield* createChartHierarchyPlanSteps(source, chart, context)
  const output = yield* executeChartHierarchyTaskSteps(plan.task, context)
  return yield* decodeChartHierarchySteps(plan, output, context)
}

export async function chartHierarchyResultAsync(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Promise<ChartHierarchyData> {
  if (!context.runChartTask) return consumeChartPreparation(chartHierarchyResultSteps(source, chart, context), context)
  const plan = await consumeChartPreparation(createChartHierarchyPlanSteps(source, chart, context), context)
  const output = await context.runChartTask(plan.task)
  assertChartPreparationMemory(plan.retainedBytes + output.values.byteLength + output.indices.byteLength, context)
  return consumeChartPreparation(decodeChartHierarchySteps(plan, output, context), context)
}

export function chartHierarchyData(source: ChartRecords, chart: ChartDataBinding): ChartHierarchyNode[] {
  return consumeChartPreparationSync(chartHierarchyDataSteps(source, chart))
}

export function* chartHierarchyDataSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartHierarchyNode[], void> {
  return (yield* chartHierarchyResultSteps(source, chart, context)).nodes
}

export function* chartHierarchyChartDataSteps(source: ChartRecords, result: ChartHierarchyData,
  valueField: number): Generator<void, ChartData, void> {
  const categories = yield* chartMapSteps(result.nodes, node => node.name)
  const values = yield* chartMapSteps(result.nodes, node => node.value)
  const formattedValues = yield* chartMapSteps(result.nodes, node => node.formattedValue ?? null)
  return { ...source, categories, series: [{ field: valueField, name: result.valueFieldName, sourceName: result.valueFieldName, values, formattedValues }],
    hierarchy: result.nodes, hierarchyPaletteOrder: result.paletteOrder, hierarchyValueFieldName: result.valueFieldName,
    ...(result.diagnostic ? { hierarchyDiagnostic: result.diagnostic } : {}) }
}
