import type { ChartModel } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartStatisticsParams, ChartStatisticsTask } from '../../compute/chart-protocol'
import { readChartStatisticsOutput } from '../../compute/chart-results'
import { estimateChartTaskMemory, validateChartTask } from '../../compute/chart-tasks'
import { ComputeError } from '../../compute/errors'
import type { ChartData } from './chart-engine'
import type { ChartGraphData } from './chart-data'
import type { ChartGraphLayout, ChartGraphLinkLayout, ChartGraphNodeLayout } from './chart-special-layout'
import { chartPlotFrame } from './chart-layout'
import { chartBezier, type ChartVector } from './chart-geometry'
import { CHART_RELATION_DEFAULTS, chartRelationNameSeed, resolvedChartRelationForce } from './chart-relation'
import { assertChartPreparationCurrent, assertChartPreparationMemory, yieldChartPreparation,
  CHART_PREPARATION_CHUNK, type ChartComputePreparationContext } from './chart-compute-lifecycle'

export interface ChartRelationNodePosition { index: number; x: number; y: number; symbolSize: number }
export interface ChartRelationLayoutResult { width: number; height: number; nodes: ChartRelationNodePosition[] }
export interface ChartRelationNodeLayout extends ChartGraphNodeLayout {
  id: string; category?: string; colorIndex: number; symbolSize: number
}
export interface ChartRelationLinkLayout extends ChartGraphLinkLayout {
  index: number; sourceIndex: number; targetIndex: number; selfLoop: boolean
}
export interface ChartRelationGraphLayout extends ChartGraphLayout {
  nodes: ChartRelationNodeLayout[]; links: ChartRelationLinkLayout[]; center: ChartVector
  width: number; height: number; categories: string[]; plot: RenderRect
}
export interface ChartRelationLayoutPlan { params: ChartStatisticsParams; graph: ChartGraphData }

/** Only numeric columns cross the Worker boundary; names/categories and edge
 * identity remain with the owner. A later translation uses this same result. */
export async function prepareChartRelationLayout(data: ChartData, chart: ChartModel,
  context: ChartComputePreparationContext): Promise<ChartData> {
  if (chart.type !== 'relation') return data
  assertChartPreparationCurrent(context)
  const graph = data.graph, plot = chartPlotFrame(chart, data.series).plot
  if (!graph) throw new ComputeError('COMPUTE_INVALID_INPUT', '关系图缺少同一版本的节点和连线数据')
  if (!graph.nodes.length) return { ...data, relation: { width: plot.width, height: plot.height, nodes: [] } }
  if (!context.runChartTask) throw new ComputeError('COMPUTE_WASM_API_MISSING', '关系图排布需要图表 Worker/Rust 计算服务')
  const nodeCount = graph.nodes.length, linkCount = graph.links.length, rowCount = Math.max(nodeCount, linkCount)
  const useValueAsSymbolSize = chart.options?.relation?.useValueAsSymbolSize ?? CHART_RELATION_DEFAULTS.useValueAsSymbolSize
  const force = resolvedChartRelationForce(chart.options?.relation, chart.size.width, chart.size.height)
  const params: ChartStatisticsParams = { rowCount, columnCount: 5, categoryCount: 1,
    series: [{ field: data.series[0]?.field ?? 0, seriesPosition: 0, column: 0 }],
    jobs: [{ kind: 'relationLayout', seriesIndex: 0, slot: 0, seedColumn: 1, sourceColumn: 2, targetColumn: 3, linkValueColumn: 4,
      nodeCount, linkCount, width: plot.width, height: plot.height,
      layout: chart.options?.relation?.layout ?? CHART_RELATION_DEFAULTS.layout,
      useValueAsSymbolSize, symbolSize: useValueAsSymbolSize ? Math.min(chart.size.width, chart.size.height) * 0.08 : 10,
      ...force }] }
  assertChartPreparationMemory(estimateChartTaskMemory(params).estimatedPeakBytes + nodeCount * 256, context)
  const values = new Float64Array(rowCount * 5), states = new Uint8Array(values.length), categoryIds = new Uint32Array(rowCount)
  let work = 0
  for (let index = 0; index < nodeCount; index++) {
    values[index] = graph.nodes[index]!.value; states[index] = 1
    values[rowCount + index] = chartRelationNameSeed(graph.nodes[index]!.name); states[rowCount + index] = 1
    if (++work % CHART_PREPARATION_CHUNK === 0) await yieldChartPreparation(context)
  }
  for (let index = 0; index < linkCount; index++) {
    const edge = graph.links[index]!
    if (!Number.isSafeInteger(edge.sourceIndex) || !Number.isSafeInteger(edge.targetIndex)
      || edge.sourceIndex! < 0 || edge.targetIndex! < 0 || edge.sourceIndex! >= nodeCount || edge.targetIndex! >= nodeCount) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '关系图连线缺少有效的节点身份')
    }
    values[rowCount * 2 + index] = edge.sourceIndex!; states[rowCount * 2 + index] = 1
    values[rowCount * 3 + index] = edge.targetIndex!; states[rowCount * 3 + index] = 1
    values[rowCount * 4 + index] = edge.value; states[rowCount * 4 + index] = 1
    if (++work % CHART_PREPARATION_CHUNK === 0) await yieldChartPreparation(context)
  }
  const task: ChartStatisticsTask = { kind: 'chartStatistics', values, states, categoryIds, params }
  validateChartTask(task)
  const output = await context.runChartTask(task)
  assertChartPreparationCurrent(context)
  return { ...data, relation: await decodeChartRelationLayout({ params, graph }, output, context) }
}

export async function decodeChartRelationLayout(plan: ChartRelationLayoutPlan, candidate: unknown,
  context: ChartComputePreparationContext = {}): Promise<ChartRelationLayoutResult> {
  const output = readChartStatisticsOutput(candidate, plan.params), job = plan.params.jobs[0]
  if (!output || output.segments.length !== 1 || job?.kind !== 'relationLayout' || job.nodeCount !== plan.graph.nodes.length) {
    throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '关系图排布结果与节点任务不匹配')
  }
  const segment = output.segments[0]!, nodes: ChartRelationNodePosition[] = Array(job.nodeCount), seen = new Set<number>()
  if (segment.indexLength !== job.nodeCount || segment.valueLength !== job.nodeCount * 3
    || segment.validSampleCount !== job.nodeCount || segment.available !== (job.nodeCount > 0)) {
    throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '关系图排布结果没有完整覆盖节点')
  }
  assertChartPreparationMemory(job.nodeCount * 256 + output.values.byteLength + output.indices.byteLength, context)
  let maximum = 0
  for (const node of plan.graph.nodes) maximum = Math.max(maximum, node.value)
  for (let index = 0; index < job.nodeCount; index++) {
    const id = output.indices[segment.indexOffset + index]!, source = plan.graph.nodes[id], at = segment.valueOffset + index * 3
    const x = output.values[at]!, y = output.values[at + 1]!, symbolSize = output.values[at + 2]!
    const expectedSize = !job.useValueAsSymbolSize ? job.symbolSize : maximum === 0 ? Math.min(job.symbolSize, 10) : (source?.value ?? 0) / maximum * job.symbolSize
    if (!source || seen.has(id) || ![x, y, symbolSize].every(Number.isFinite) || symbolSize < 0
      || Math.abs(symbolSize - expectedSize) > Math.max(1, job.symbolSize) * 1e-10
      || job.layout !== 'none' && (x < 0 || x > job.width || y < 0 || y > job.height)) {
      throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '关系图节点身份、坐标或线性大小无效')
    }
    seen.add(id); nodes[id] = { index: id, x, y, symbolSize }
    if (index % CHART_PREPARATION_CHUNK === 0) await yieldChartPreparation(context)
  }
  assertChartPreparationCurrent(context)
  return { width: job.width, height: job.height, nodes }
}

/** Scene geometry is derived from one prepared coordinate set, never from a
 * second solver. Explicit numeric identities also preserve matrix names which
 * happen to be equal, duplicate edges and self-links. */
export function chartRelationLayout(data: ChartData, chart: ChartModel, plot: RenderRect): ChartRelationGraphLayout {
  const graph = data.graph, retained = data.relation
  const center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }
  if (!graph?.nodes.length) return { nodes: [], links: [], categories: graph?.categories ?? [], center, width: plot.width, height: plot.height, plot: { ...plot } }
  if (!retained || retained.nodes.length !== graph.nodes.length) throw new ComputeError('COMPUTE_NOT_READY', '关系图尚未获得同一版本的节点排布结果')
  if (Math.abs(retained.width - plot.width) > 1e-9 || Math.abs(retained.height - plot.height) > 1e-9) {
    throw new ComputeError('COMPUTE_STALE_RESULT', '关系图画布尺寸已改变，需要重新排布')
  }
  const categories = graph.categories ?? [], categoryIndexes = new Map(categories.map((category, index) => [category, index]))
  const nodes: ChartRelationNodeLayout[] = graph.nodes.map((node, index) => {
    const point = retained.nodes[index]!
    if (point.index !== index) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '关系图排布节点顺序已失效')
    return { ...node, id: node.id ?? String(index), index, x: plot.x + point.x, y: plot.y + point.y,
      width: point.symbolSize, height: point.symbolSize, symbolSize: point.symbolSize,
      colorIndex: categories.length ? categoryIndexes.get(node.category ?? '') ?? 0 : index }
  })
  const links: ChartRelationLinkLayout[] = graph.links.map((link, index) => {
    const first = nodes[link.sourceIndex!], second = nodes[link.targetIndex!]
    if (!first || !second) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '关系图连线引用的节点已失效')
    const selfLoop = link.sourceIndex === link.targetIndex, radius = Math.max(4, first.symbolSize / 2)
    const points = selfLoop
      ? chartBezier({ x: first.x - radius * 0.65, y: first.y - radius * 0.75 },
        { x: first.x - radius - 22, y: first.y - radius - 36 },
        { x: first.x + radius + 22, y: first.y - radius - 36 },
        { x: first.x + radius * 0.65, y: first.y - radius * 0.75 })
      : [{ x: first.x, y: first.y }, { x: second.x, y: second.y }]
    return { ...link, sourceIndex: link.sourceIndex!, targetIndex: link.targetIndex!, index, selfLoop,
      points, width: 1, colorIndex: first.colorIndex }
  })
  return { nodes, links, categories, center, width: plot.width, height: plot.height, plot: { ...plot } }
}
