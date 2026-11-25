import type { ChartChordMapping } from '../../domain/model'
import type { ChartStatisticsTask } from '../../compute/chart-protocol'
import { estimateChartTaskMemory } from '../../compute/chart-tasks'
import { COMPUTE_LIMITS } from '../../compute/protocol'
import { ComputeError } from '../../compute/errors'
import type { ChartData } from './chart-engine'
import { chartField, chartValueField, type ChartDataBinding, type ChartEdge, type ChartRecords } from './chart-data'
import { chartChordMappingValues } from './chart-chord-api'
import { resolveChartChordLayout } from './chart-chord'
import { chartHierarchySourceFormatterSteps } from './chart-hierarchy-format'
import { assertChartPreparationMemory, CHART_PREPARATION_CHUNK, consumeChartPreparation, consumeChartPreparationSync,
  type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { chartMapSteps, chartProjectionSourceBytes } from './chart-data-preparation'
import { decodeChartChordSteps, executeChartChordTaskSteps } from './chart-chord-compute'

export interface ChartChordNode {
  id: string; index: number; name: string; value: number; formattedValue: string
  start: number; end: number
}
export interface ChartChordLink extends ChartEdge {
  id: string; index: number; sourceId: string; targetId: string; sourceIndex: number; targetIndex: number
  formattedValue: string; sourceStart: number; sourceEnd: number; targetStart: number; targetEnd: number
}
/** Only runtime facts/angles. No source node IDs or renderer geometry are saved. */
export interface ChartChordData {
  mapping: ChartChordMapping; weightMode: 'value' | 'count'; valueFieldName: string
  nodes: ChartChordNode[]; links: ChartChordLink[]; validSampleCount: number
  angleOptions: Pick<ReturnType<typeof resolveChartChordLayout>, 'startAngle' | 'clockwise' | 'padAngle' | 'minAngle'>
}
export interface ChartChordPlan {
  task: ChartStatisticsTask; labels: string[]; mapping: ChartChordMapping; valueFieldName: string
  formatter: (value: number) => string; retainedBytes: number
}

export function chartChordNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : null
}

export function chartChordMapping(source: ChartRecords, chart: ChartDataBinding): ChartChordMapping {
  const authored = chart.options?.mapping?.chord ?? chart.options?.mapping?.edge
    ?? { sourceIndex: chart.categoryField, targetIndex: chart.valueFields[0] ?? 1,
      valueIndex: chart.valueFields[1] ?? (source.fields.length > 2 ? 2 : undefined) }
  const mapping = chartChordMappingValues(authored)
  chartField(source, mapping.sourceIndex, '和弦来源字段'); chartField(source, mapping.targetIndex, '和弦目标字段')
  if (mapping.valueIndex !== undefined) chartValueField(source, mapping.valueIndex, '和弦权重字段')
  return mapping
}

/** Display strings stay with the owner. The Worker receives three retained
 * numeric/state columns (weight, source ID, target ID), without coercion. */
export function* createChartChordPlanSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartChordPlan, void> {
  const mapping = chartChordMapping(source, chart), rows = source.records.length, columns = 3
  if (rows > COMPUTE_LIMITS.maxRows || rows * columns > COMPUTE_LIMITS.maxElements) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '和弦图完整数据超出计算预算。')
  }
  const { startAngle, clockwise, padAngle, minAngle } = resolveChartChordLayout(chart.options?.chord)
  const params: ChartStatisticsTask['params'] = { rowCount: rows, columnCount: columns, categoryCount: rows ? 1 : 0,
    series: [{ field: mapping.valueIndex ?? mapping.targetIndex, seriesPosition: 0, column: 0 }],
    jobs: [{ kind: 'chord', seriesIndex: 0, slot: 0, sourceColumn: 1, targetColumn: 2, labelCount: 0, startAngle, clockwise, padAngle, minAngle }] }
  const job = params.jobs[0]!
  if (job.kind !== 'chord') throw new ComputeError('COMPUTE_INVALID_INPUT', '和弦图任务类型无效。')
  const sourceBytes = chartProjectionSourceBytes(rows, source.fields.length), peak = estimateChartTaskMemory(params).estimatedPeakBytes
  assertChartPreparationMemory(sourceBytes + peak, context)
  const values = new Float64Array(rows * columns), states = new Uint8Array(values.length), categoryIds = new Uint32Array(rows)
  const labels: string[] = [], dictionary = new Map<string, number>()
  let retainedBytes = sourceBytes, work = 0
  const display = (row: number, field: number): string => source.formattedRecords?.[row]?.[field]
    ?? (source.records[row]?.[field] == null ? '' : String(source.records[row]![field]))
  const label = (text: string): number => {
    const found = dictionary.get(text)
    if (found !== undefined) return found
    retainedBytes += 96 + text.length * 2
    assertChartPreparationMemory(retainedBytes + peak, context)
    const next = labels.length; labels.push(text); dictionary.set(text, next); return next
  }
  for (let row = 0; row < rows; row++) {
    const weight = mapping.valueIndex === undefined ? 1 : chartChordNumber(source.records[row]?.[mapping.valueIndex])
    const sourceName = display(row, mapping.sourceIndex), targetName = display(row, mapping.targetIndex)
    if (weight !== null && sourceName !== '' && targetName !== '' && sourceName !== targetName) {
      values[row] = weight; states[row] = 1
      values[rows + row] = label(sourceName); states[rows + row] = 1
      values[rows * 2 + row] = label(targetName); states[rows * 2 + row] = 1
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  job.labelCount = labels.length
  const formatter = mapping.valueIndex === undefined ? String : yield* chartHierarchySourceFormatterSteps(source, mapping.valueIndex)
  const valueFieldName = mapping.valueIndex === undefined ? '关系次数' : source.fields[mapping.valueIndex] ?? '关系权重'
  return { task: { kind: 'chartStatistics', values, states, categoryIds, params }, labels, mapping, valueFieldName, formatter, retainedBytes }
}

export function* chartChordResultSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartChordData, void> {
  const plan = yield* createChartChordPlanSteps(source, chart, context)
  const output = yield* executeChartChordTaskSteps(plan.task, context)
  return yield* decodeChartChordSteps(plan, output, context)
}
export function chartChordData(source: ChartRecords, chart: ChartDataBinding): ChartChordData {
  return consumeChartPreparationSync(chartChordResultSteps(source, chart))
}
export async function chartChordResultAsync(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Promise<ChartChordData> {
  if (!context.runChartTask) return consumeChartPreparation(chartChordResultSteps(source, chart, context), context)
  const plan = await consumeChartPreparation(createChartChordPlanSteps(source, chart, context), context)
  const output = await context.runChartTask(plan.task)
  assertChartPreparationMemory(plan.retainedBytes + output.values.byteLength + output.indices.byteLength, context)
  return consumeChartPreparation(decodeChartChordSteps(plan, output, context), context)
}
export function* chartChordChartDataSteps(source: ChartRecords, result: ChartChordData): Generator<void, ChartData, void> {
  const categories = yield* chartMapSteps(result.nodes, node => node.name)
  const values = yield* chartMapSteps(result.nodes, node => node.value)
  const formattedValues = yield* chartMapSteps(result.nodes, node => node.formattedValue)
  return { ...source, categories,
    series: [{ field: result.mapping.valueIndex ?? result.mapping.targetIndex, name: result.valueFieldName, sourceName: result.valueFieldName, values, formattedValues }],
    // Compatibility graph references the same arrays rather than retaining a
    // second graph or running a directed SCC/layout engine for Chord.
    graph: { nodes: result.nodes, links: result.links, cyclic: false }, chord: result }
}
