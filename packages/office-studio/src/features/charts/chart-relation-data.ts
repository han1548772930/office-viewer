import type { CellScalar } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { chartField, chartValueField, type ChartDataBinding, type ChartEdge, type ChartGraphData, type ChartRecords } from './chart-data'
import { chartMapSteps, chartProjectionSourceBytes } from './chart-data-preparation'
import { assertChartPreparationMemory, CHART_PREPARATION_CHUNK, type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { finiteChartValue } from './chart-math'

export type ChartRelationSourceFormat = 'edgeList' | 'coWordMatrix'
/** The local row-major adapter requires a complete square and matching node
 * labels. It preserves the upstream matrix-before-edge recognition contract
 * without applying its irregular converted column DTO to ordinary cells. */
export function isChartRelationCoWordMatrix(source: ChartRecords): boolean {
  const count = source.records.length
  if (!count || source.fields.length !== count + 2) return false
  const names = new Set<string>()
  for (let index = 0; index < count; index++) {
    const record = source.records[index], name = source.fields[index + 2]
    if (!record || record.length !== count + 2 || !endpoint(name) || record[0] !== name || names.has(name)) return false
    names.add(name)
  }
  return true
}
export function chartRelationSourceFormat(chart: Pick<ChartDataBinding, 'options'>, source?: ChartRecords): ChartRelationSourceFormat {
  const mapping = chart.options?.mapping
  if (mapping?.relationFormat) return mapping.relationFormat
  if (mapping?.edge) return 'edgeList'
  return source && isChartRelationCoWordMatrix(source) ? 'coWordMatrix' : 'edgeList'
}
export function chartRelationMapping(source: ChartRecords, chart: ChartDataBinding): { sourceIndex: number; targetIndex: number; valueIndex?: number } {
  const mapping = chart.options?.mapping?.edge ?? { sourceIndex: 0, targetIndex: 1, valueIndex: source.fields.length > 2 ? 2 : undefined }
  chartField(source, mapping.sourceIndex, '关系来源字段'); chartField(source, mapping.targetIndex, '关系目标字段')
  if (mapping.sourceIndex === mapping.targetIndex || mapping.valueIndex === mapping.sourceIndex || mapping.valueIndex === mapping.targetIndex) {
    throw new RangeError('关系来源、目标和权重须使用互不重复的字段。')
  }
  if (mapping.valueIndex !== undefined) chartValueField(source, mapping.valueIndex, '关系权重字段')
  return mapping
}
function empty(format: ChartRelationSourceFormat, diagnostic: string): ChartGraphData {
  return { nodes: [], links: [], categories: [], cyclic: false, valid: false, diagnostic, sourceKind: format }
}
const endpoint = (value: CellScalar | undefined): value is string => typeof value === 'string' && value.trim().length > 0
const weight = (value: CellScalar | undefined): value is number => typeof value === 'number' && Number.isFinite(value) && value >= 0

export function* chartRelationGraphSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartGraphData, void> {
  if (chartRelationSourceFormat(chart, source) === 'coWordMatrix') return yield* chartRelationMatrixSteps(source, context)
  const mapping = chartRelationMapping(source, chart), categories = mapping.valueIndex === undefined || !source.fields[mapping.valueIndex]
    ? [] : [source.fields[mapping.valueIndex]!]
  if (!source.records.length) return empty('edgeList', '关系来源需要至少一条有效连线。')
  const sourceBytes = chartProjectionSourceBytes(source.records.length, source.fields.length)
  let work = 0
  // Fixed Gae invalidates the complete source when any endpoint/weight is bad.
  // No trimming, numeric-string coercion, zero suppression or edge coalescing.
  for (const record of source.records) {
    if (!endpoint(record[mapping.sourceIndex]) || !endpoint(record[mapping.targetIndex])
      || mapping.valueIndex !== undefined && !weight(record[mapping.valueIndex])) {
      return empty('edgeList', '关系来源和目标必须为非空文本，权重必须为有限非负数；请修复全部无效行。')
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const nodes: ChartGraphData['nodes'] = [], links: ChartEdge[] = [], ids = new Map<string, number>()
  let names = 0
  const nodeAt = (name: string): number => {
    const previous = ids.get(name)
    if (previous !== undefined) return previous
    const index = nodes.length
    ids.set(name, index); names += name.length * 2
    nodes.push({ id: String(index), name, value: 0, ...(categories.length ? { category: categories[0] } : {}) })
    return index
  }
  for (const record of source.records) {
    const first = record[mapping.sourceIndex], second = record[mapping.targetIndex]
    if (!endpoint(first) || !endpoint(second)) throw new TypeError('关系来源在准备期间已更改。')
    const value = mapping.valueIndex === undefined ? 1 : record[mapping.valueIndex]
    if (!weight(value)) throw new TypeError('关系权重在准备期间已更改。')
    const from = nodeAt(first), to = nodeAt(second)
    nodes[from]!.value = finiteChartValue(nodes[from]!.value + value, '关系节点权重合计')
    if (from !== to) nodes[to]!.value = finiteChartValue(nodes[to]!.value + value, '关系节点权重合计')
    links.push({ source: first, target: second, sourceIndex: from, targetIndex: to, value })
    assertChartPreparationMemory(sourceBytes + names + nodes.length * 192 + links.length * 192, context)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return { nodes, links, categories, cyclic: false, valid: true, sourceKind: 'edgeList' }
}

/** Explicit row-major adapter: N rows and N+2 headers, with categories in field
 * 1 and node names in headers 2..N+1. The fixed upstream DTO has N+1 source
 * columns; this adapter retains both authored label/category columns instead
 * of silently applying that irregular DTO to an ordinary N*N table. */
export function* chartRelationMatrixSteps(source: ChartRecords,
  context: ChartComputePreparationContext = {}): Generator<void, ChartGraphData, void> {
  const count = source.records.length
  if (!count || source.fields.length !== count + 2 || source.records.some(record => record.length !== count + 2)) {
    return empty('coWordMatrix', '共词矩阵需要 N 行节点、名称和分类两列，以及 N 列矩阵权重。')
  }
  const nodes: ChartGraphData['nodes'] = [], links: ChartEdge[] = [], categories: string[] = [], categorySet = new Set<string>()
  const sourceBytes = chartProjectionSourceBytes(count, source.fields.length)
  let work = 0, names = 0
  for (let index = 0; index < count; index++) {
    const name = source.fields[index + 2]!, category = String(source.records[index]![1] ?? '')
    names += (name.length + category.length) * 2
    nodes.push({ id: String(index), name, category, value: 0 })
    if (!categorySet.has(category)) { categorySet.add(category); categories.push(category) }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  for (let from = 0; from < count; from++) for (let to = 0; to < count; to++) {
    const value = source.records[from]![to + 2]
    if (typeof value === 'number') {
      if (!weight(value)) return empty('coWordMatrix', '共词矩阵的数值权重必须有限且非负。')
      links.push({ source: nodes[from]!.name, target: nodes[to]!.name, sourceIndex: from, targetIndex: to, value })
      // Wae accumulates only outgoing matrix weights, including self-links.
      nodes[from]!.value = finiteChartValue(nodes[from]!.value + value, '共词矩阵节点出边权重合计')
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) {
      assertChartPreparationMemory(sourceBytes + names + nodes.length * 192 + links.length * 192, context); yield
    }
  }
  assertChartPreparationMemory(sourceBytes + names + nodes.length * 192 + links.length * 192, context)
  return { nodes, links, categories, cyclic: false, valid: true, sourceKind: 'coWordMatrix' }
}

export function* chartRelationChartDataSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartData, void> {
  const graph = yield* chartRelationGraphSteps(source, chart, context)
  const mapping = graph.sourceKind === 'coWordMatrix' ? undefined : chartRelationMapping(source, chart)
  const field = mapping?.valueIndex ?? (mapping?.targetIndex ?? 2)
  const name = graph.categories?.length === 1 ? graph.categories[0]! : mapping?.valueIndex === undefined ? '关系次数' : source.fields[field] ?? '关系权重'
  const categories = yield* chartMapSteps(graph.nodes, node => node.name), values = yield* chartMapSteps(graph.nodes, node => node.value)
  return { ...source, categories, series: [{ field, name, sourceName: source.fields[field] ?? name, values }], graph }
}
