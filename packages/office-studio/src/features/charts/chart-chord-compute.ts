import type { ChartStatisticsJob, ChartStatisticsOutput, ChartStatisticsTask } from '../../compute/chart-protocol'
import { readChartStatisticsOutput } from '../../compute/chart-results'
import { estimateChartTaskMemory, validateChartTask } from '../../compute/chart-tasks'
import { ComputeError } from '../../compute/errors'
import type { ChartChordData, ChartChordPlan } from './chart-chord-data'
import { assertChartPreparationMemory, CHART_PREPARATION_CHUNK, type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { finiteChartValue } from './chart-math'

type ChordJob = Extract<ChartStatisticsJob, { kind: 'chord' }>
interface NumericNode { label: number; value: number; span: number; start: number; end: number }
interface NumericLink { source: number; target: number; value: number; sourceStart: number; sourceEnd: number; targetStart: number; targetEnd: number }
const TAU = Math.PI * 2
const invalid = (message: string): never => { throw new ComputeError('COMPUTE_INVALID_INPUT', message) }
const mismatch = (message: string): never => { throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', message) }

export function chartChordAngleParameters(job: Pick<ChordJob, 'startAngle' | 'clockwise' | 'padAngle' | 'minAngle'>, nodes: number): { start: number; direction: 1 | -1; padding: number; minimum: number; available: number } {
  const start = -(job.startAngle % 360) / 180 * Math.PI, direction = job.clockwise ? 1 : -1
  if (!nodes) return { start, direction, padding: 0, minimum: 0, available: 0 }
  const maximum = TAU / nodes
  let minimum = Math.min(maximum, job.minAngle / 180 * Math.PI), padding = job.padAngle / 180 * Math.PI
  if (padding >= maximum) padding = Math.max(0, maximum - minimum)
  if (padding + minimum >= maximum) minimum = Math.max(0, maximum - padding)
  const available = Math.max(0, TAU - padding * nodes)
  minimum = Math.min(minimum, available / nodes)
  return { start, direction, padding, minimum, available }
}

/** Small synchronous compatibility uses the same packed numeric contract as
 * Rust. Studio preparation supplies runChartTask and executes this job there. */
export function* executeChartChordTaskSteps(task: ChartStatisticsTask,
  context: ChartComputePreparationContext = {}): Generator<void, ChartStatisticsOutput, void> {
  validateChartTask(task)
  const job = task.params.jobs[0], series = task.params.series[0], rows = task.params.rowCount
  if (!job || job.kind !== 'chord' || job.seriesIndex !== 0 || job.slot !== 0 || task.params.jobs.length !== 1
    || !series || task.params.series.length !== 1) invalid('和弦计算须包含一个和弦任务和一个权重系列。')
  assertChartPreparationMemory(estimateChartTaskMemory(task.params).estimatedPeakBytes, context)
  const nodes: NumericNode[] = [], links: NumericLink[] = [], nodeIds = new Map<number, number>(), linkIds = new Map<string, number>()
  let work = 0, validSampleCount = 0
  const labelAt = (column: number, row: number): number | undefined => {
    const at = column * rows + row, state = task.states[at], label = task.values[at]!
    if (state === 0) return undefined
    if (state !== 1 || !Number.isSafeInteger(label) || label < 0 || label >= job.labelCount) invalid('和弦端点字典索引或状态无效。')
    return label
  }
  const nodeAt = (label: number): number => {
    const found = nodeIds.get(label)
    if (found !== undefined) return found
    const index = nodes.length; nodes.push({ label, value: 0, span: 0, start: 0, end: 0 }); nodeIds.set(label, index); return index
  }
  for (let row = 0; row < rows; row++) {
    const first = labelAt(job.sourceColumn, row), second = labelAt(job.targetColumn, row), at = series.column * rows + row
    const state = task.states[at], value = task.values[at]!
    if (state !== 0 && state !== 1 && state !== 2) invalid('和弦权重状态无效。')
    if (state === 1 && Number.isFinite(value) && value > 0 && first !== undefined && second !== undefined && first !== second) {
      const source = nodeAt(first), target = nodeAt(second), key = `${Math.min(source, target)}:${Math.max(source, target)}`
      const previous = linkIds.get(key)
      if (previous === undefined) {
        linkIds.set(key, links.length)
        links.push({ source, target, value, sourceStart: 0, sourceEnd: 0, targetStart: 0, targetEnd: 0 })
      } else links[previous]!.value = finiteChartValue(links[previous]!.value + value, '和弦连接权重合计')
      validSampleCount++
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  // Fixed wae aggregates links first, then adds them to each incident node in
  // first-link order. Row-order node summing can produce different roundoff.
  for (const link of links) {
    nodes[link.source]!.value = finiteChartValue(nodes[link.source]!.value + link.value, '和弦节点关联权重合计')
    nodes[link.target]!.value = finiteChartValue(nodes[link.target]!.value + link.value, '和弦节点关联权重合计')
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const angles = chartChordAngleParameters(job, nodes.length)
  let scale = 0, total = 0
  for (const node of nodes) { scale = Math.max(scale, node.value); if (++work % CHART_PREPARATION_CHUNK === 0) yield }
  for (const node of nodes) { total += node.value / scale; if (++work % CHART_PREPARATION_CHUNK === 0) yield }
  let deficit = 0, surplus = 0
  for (const node of nodes) {
    node.span = angles.available * ((node.value / scale) / total)
    deficit += Math.max(0, angles.minimum - node.span); surplus += Math.max(0, node.span - angles.minimum)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const fraction = surplus > 0 ? Math.min(1, deficit / surplus) : 0
  let allocated = 0, allocationCorrection = 0
  for (const node of nodes) {
    node.span = node.span < angles.minimum ? angles.minimum : node.span - (node.span - angles.minimum) * fraction
    const addend = node.span - allocationCorrection, next = allocated + addend
    allocationCorrection = next - allocated - addend; allocated = next
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  // A bounded floating residual is assigned to the final arc, without changing
  // value proportions or introducing an unowned gap beyond the authored pads.
  if (nodes.length) nodes.at(-1)!.span = Math.max(0, nodes.at(-1)!.span + angles.available - allocated)
  let consumed = 0, angleCorrection = 0
  const offsets = new Float64Array(nodes.length), consumedValues = new Float64Array(nodes.length)
  for (const [index, node] of nodes.entries()) {
    node.start = angles.start + consumed * angles.direction; node.end = node.start + node.span * angles.direction; offsets[index] = node.start
    const addend = node.span + angles.padding - angleCorrection, next = consumed + addend
    angleCorrection = next - consumed - addend; consumed = next
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  for (const link of links) {
    const source = nodes[link.source]!, target = nodes[link.target]!
    link.sourceStart = offsets[link.source]!; link.targetStart = offsets[link.target]!
    consumedValues[link.source] = finiteChartValue(consumedValues[link.source]! + link.value, '和弦端点累计权重')
    consumedValues[link.target] = finiteChartValue(consumedValues[link.target]! + link.value, '和弦端点累计权重')
    link.sourceEnd = source.start + source.span * (consumedValues[link.source]! / source.value) * angles.direction
    link.targetEnd = target.start + target.span * (consumedValues[link.target]! / target.value) * angles.direction
    offsets[link.source] = link.sourceEnd; offsets[link.target] = link.targetEnd
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const values = new Float64Array(nodes.length * 3 + links.length * 5), indices = new Uint32Array(2 + nodes.length + links.length * 2)
  indices[0] = nodes.length; indices[1] = links.length
  for (const [index, node] of nodes.entries()) {
    values[index * 3] = node.value; values[index * 3 + 1] = node.start; values[index * 3 + 2] = node.end; indices[index + 2] = node.label
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  for (const [index, link] of links.entries()) {
    const at = nodes.length * 3 + index * 5, id = 2 + nodes.length + index * 2
    values[at] = link.value; values[at + 1] = link.sourceStart; values[at + 2] = link.sourceEnd
    values[at + 3] = link.targetStart; values[at + 4] = link.targetEnd; indices[id] = link.source; indices[id + 1] = link.target
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return { kind: 'chartStatistics', values, indices, segments: [{ kind: 'chord', jobIndex: 0, field: series.field, seriesPosition: 0,
    slot: 0, valueOffset: 0, valueLength: values.length, indexOffset: 0, indexLength: indices.length, validSampleCount, available: links.length > 0 }] }
}

const near = (first: number, second: number): boolean => Math.abs(first - second) <= Math.max(1, Math.abs(first), Math.abs(second)) * 1e-10
const nearValue = (first: number, second: number): boolean => first === second
  || Math.abs(first - second) <= Math.max(Math.abs(first), Math.abs(second)) * 1e-10

/** Validate each consumed value/index and reconstruct semantics from the
 * retained dictionary. The decoder neither aggregates source rows nor reruns
 * a second angle layout after the Worker result has become authoritative. */
export function* decodeChartChordSteps(plan: ChartChordPlan, candidate: unknown,
  context: ChartComputePreparationContext = {}): Generator<void, ChartChordData, void> {
  const output = readChartStatisticsOutput(candidate, plan.task.params), job = plan.task.params.jobs[0]
  if (!output || output.segments.length !== 1 || !job || job.kind !== 'chord' || job.labelCount !== plan.labels.length) mismatch('和弦计算结果与任务或字典不匹配。')
  const segment = output.segments[0]!, nodeCount = output.indices[0]!, linkCount = output.indices[1]!
  const angles = chartChordAngleParameters(job, nodeCount), nodes: ChartChordData['nodes'] = [], links: ChartChordData['links'] = []
  const labels = new Set<number>(), pairs = new Set<string>(), sums = new Float64Array(nodeCount), offsets = new Float64Array(nodeCount)
  let work = 0, angle = angles.start, retainedBytes = plan.retainedBytes + output.values.byteLength + output.indices.byteLength + nodeCount * 16
  assertChartPreparationMemory(retainedBytes + nodeCount * 512 + linkCount * 384, context)
  for (let index = 0; index < nodeCount; index++) {
    const at = index * 3, label = output.indices[index + 2]!, name = plan.labels[label], value = output.values[at]!
    const start = output.values[at + 1]!, end = output.values[at + 2]!, span = (end - start) * angles.direction
    if (label >= plan.labels.length || name === undefined || name === '' || labels.has(label)
      || ![value, start, end].every(Number.isFinite) || value <= 0 || span < -1e-10 || span > TAU + 1e-10
      || span + 1e-10 < angles.minimum || !near(start, angle)) mismatch('和弦节点身份、数值或角度无效。')
    const formattedValue = plan.formatter(value), id = `chord-node:${index}`
    if (typeof formattedValue !== 'string') mismatch('和弦数值格式化结果无效。')
    retainedBytes += 512 + (name.length + id.length + formattedValue.length) * 2
    assertChartPreparationMemory(retainedBytes + (nodeCount - index - 1) * 512 + linkCount * 384, context)
    nodes.push({ id, index, name, value, formattedValue, start, end }); labels.add(label); offsets[index] = start
    angle = end + angles.padding * angles.direction
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (nodeCount && !near(angle, angles.start + TAU * angles.direction)) mismatch('和弦节点和间隔未组成完整圆周。')
  for (let index = 0; index < linkCount; index++) {
    const at = nodeCount * 3 + index * 5, idAt = 2 + nodeCount + index * 2
    const sourceIndex = output.indices[idAt]!, targetIndex = output.indices[idAt + 1]!, source = nodes[sourceIndex], target = nodes[targetIndex]
    const value = output.values[at]!, sourceStart = output.values[at + 1]!, sourceEnd = output.values[at + 2]!
    const targetStart = output.values[at + 3]!, targetEnd = output.values[at + 4]!, key = `${Math.min(sourceIndex, targetIndex)}:${Math.max(sourceIndex, targetIndex)}`
    if (!source || !target || sourceIndex === targetIndex || pairs.has(key) || value <= 0
      || ![value, sourceStart, sourceEnd, targetStart, targetEnd].every(Number.isFinite)
      || !near(sourceStart, offsets[sourceIndex]!) || !near(targetStart, offsets[targetIndex]!)) mismatch('和弦连接身份、合计或起始角度无效。')
    for (const [node, start, end] of [[source, sourceStart, sourceEnd], [target, targetStart, targetEnd]] as const) {
      const span = (end - start) * angles.direction, maximum = (node.end - node.start) * angles.direction
      if (value > node.value && !nearValue(value, node.value) || span < -1e-10
        || (start - node.start) * angles.direction < -1e-10 || (node.end - end) * angles.direction < -1e-10
        || !near(span, maximum * (value / node.value))) mismatch('和弦弦带角度未按端点关联数值分配。')
    }
    const formattedValue = plan.formatter(value), id = `chord-link:${index}`
    if (typeof formattedValue !== 'string') mismatch('和弦连接格式化结果无效。')
    retainedBytes += 384 + (id.length + formattedValue.length) * 2
    assertChartPreparationMemory(retainedBytes + (linkCount - index - 1) * 384, context)
    links.push({ id, index, source: source.name, target: target.name, sourceId: source.id, targetId: target.id, sourceIndex, targetIndex,
      value, formattedValue, sourceStart, sourceEnd, targetStart, targetEnd })
    pairs.add(key); sums[sourceIndex] = finiteChartValue(sums[sourceIndex]! + value, '和弦解码节点关联数值')
    sums[targetIndex] = finiteChartValue(sums[targetIndex]! + value, '和弦解码节点关联数值')
    offsets[sourceIndex] = sourceEnd; offsets[targetIndex] = targetEnd
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  for (const node of nodes) {
    if (!nearValue(node.value, sums[node.index]!) || !near(node.end, offsets[node.index]!)) mismatch('和弦节点合计或弦带闭合结果无效。')
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return { mapping: { ...plan.mapping }, valueFieldName: plan.valueFieldName, weightMode: plan.mapping.valueIndex === undefined ? 'count' : 'value',
    nodes, links, validSampleCount: segment.validSampleCount,
    angleOptions: { startAngle: job.startAngle, clockwise: job.clockwise, padAngle: job.padAngle, minAngle: job.minAngle } }
}
