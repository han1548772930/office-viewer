import type { ChartStatisticsOutput, ChartStatisticsTask } from '../../compute/chart-protocol'
import { readChartStatisticsOutput } from '../../compute/chart-results'
import { validateChartTask } from '../../compute/chart-tasks'
import { ComputeError } from '../../compute/errors'
import type { ChartHierarchyNode } from './chart-data'
import type { ChartHierarchyData, ChartHierarchyPlan } from './chart-hierarchy-data'
import { assertChartPreparationMemory, CHART_PREPARATION_CHUNK, type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { finiteChartValue } from './chart-math'

export const CHART_HIERARCHY_ROOT = 0xffff_ffff
interface NumericNode { parent: number; key: number; label: number; selfValue: number; value: number; children: number[] }
interface Token { key: number; label: number }
const invalid = (message: string): never => { throw new ComputeError('COMPUTE_INVALID_INPUT', message) }
const mismatch = (message: string): never => { throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', message) }

/** Synchronous small-source compatibility uses the same numeric contract as
 * Rust. The owner service routes asynchronous preparation through its Worker. */
export function* executeChartHierarchyTaskSteps(task: ChartStatisticsTask,
  _context: ChartComputePreparationContext = {}): Generator<void, ChartStatisticsOutput, void> {
  validateChartTask(task)
  const job = task.params.jobs[0], series = task.params.series[0], rows = task.params.rowCount
  if (!job || job.kind !== 'hierarchy' || job.seriesIndex !== 0 || job.slot !== 0 || task.params.jobs.length !== 1 || !series || task.params.series.length !== 1) {
    invalid('层级计算须包含一个层级任务和一个权重系列。')
  }
  const nodes: NumericNode[] = [{ parent: CHART_HIERARCHY_ROOT, key: 0, label: 0, selfValue: 0, value: 0, children: [] }]
  const lookup = new Map<string, number>(), cache: Array<Token | undefined> = Array(job.hierarchyColumns.length)
  let validSampleCount = 0, work = 0
  const at = (column: number, row: number): number | undefined => {
    const offset = column * rows + row
    return task.states[offset] === 1 && Number.isFinite(task.values[offset]) ? task.values[offset] : undefined
  }
  const tokenAt = (depth: number, row: number): Token | undefined => {
    const keyOffset = job.hierarchyColumns[depth]! * rows + row, labelOffset = job.labelColumns[depth]! * rows + row
    const keyState = task.states[keyOffset], labelState = task.states[labelOffset]
    if (keyState === 0 && labelState === 0) return undefined
    const key = task.values[keyOffset]!, label = task.values[labelOffset]!
    if (keyState !== 1 || labelState !== 1 || !Number.isSafeInteger(key) || key < 0 || key >= job.keyCount
      || !Number.isSafeInteger(label) || label < 0 || label >= job.labelCount) invalid('层级图 token 字典索引或状态无效。')
    return { key, label }
  }
  for (let row = 0; row < rows; row++) {
    const tokens: Array<Token | undefined> = []
    for (let depth = 0; depth < cache.length; depth++) {
      tokens.push(tokenAt(depth, row))
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    const path: Token[] = []
    let rejected = false
    for (let depth = 0; depth < cache.length; depth++) {
      const token = tokens[depth]
      if (!token) {
        const carried = cache[depth]
        if (carried) { path.push(carried); continue }
        rejected = depth === 0 || tokens.slice(depth + 1).some(Boolean)
        break
      }
      if (cache[depth]?.key !== token.key) for (let child = depth + 1; child < cache.length; child++) cache[child] = undefined
      cache[depth] = token; path.push(token)
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    // Even a textual/non-finite weight has already updated the scoped cache.
    const value = at(series.column, row)
    if (rejected || !path.length || value === undefined) continue
    validSampleCount++
    let parent = 0
    for (const token of path) {
      const identity = `${parent}:${token.key}`
      let child = lookup.get(identity)
      if (child === undefined) {
        child = nodes.length
        nodes.push({ parent, key: token.key, label: token.label, selfValue: 0, value: 0, children: [] })
        nodes[parent]!.children.push(child); lookup.set(identity, child)
      }
      parent = child
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    nodes[parent]!.selfValue = finiteChartValue(nodes[parent]!.selfValue + value, '层级路径自身权重合计')
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const compare = (left: number, right: number): number => nodes[right]!.value - nodes[left]!.value || left - right
  for (let index = nodes.length - 1; index > 0; index--) {
    const node = nodes[index]!
    node.children = node.children.filter(child => nodes[child]!.value > 0).sort(compare)
    let childValue = 0
    for (const child of node.children) {
      childValue = finiteChartValue(childValue + nodes[child]!.value, '层级子节点权重合计')
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    node.value = finiteChartValue(node.selfValue + childValue, '层级节点权重合计')
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const sourceRoots = nodes[0]!.children.filter(child => nodes[child]!.value > 0), roots = [...sourceRoots].sort(compare)
  const preorder: number[] = [], stack = [...roots].reverse(), outputIndex = new Map<number, number>()
  while (stack.length) {
    const index = stack.pop()!
    outputIndex.set(index, preorder.length); preorder.push(index)
    for (let child = nodes[index]!.children.length - 1; child >= 0; child--) stack.push(nodes[index]!.children[child]!)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const values = new Float64Array(preorder.length * 2), indices = new Uint32Array(2 + preorder.length * 4 + roots.length)
  indices[0] = preorder.length; indices[1] = roots.length
  for (let index = 0; index < preorder.length; index++) {
    const original = preorder[index]!, node = nodes[original]!, offset = 2 + index * 4
    values[index * 2] = node.value; values[index * 2 + 1] = node.selfValue
    indices[offset] = node.parent === 0 ? CHART_HIERARCHY_ROOT : outputIndex.get(node.parent)!
    indices[offset + 1] = node.key; indices[offset + 2] = node.label; indices[offset + 3] = original - 1
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  for (let index = 0; index < sourceRoots.length; index++) indices[2 + preorder.length * 4 + index] = outputIndex.get(sourceRoots[index]!)!
  return { kind: 'chartStatistics', values, indices, segments: [{ kind: 'hierarchy', jobIndex: 0, field: series.field, seriesPosition: 0, slot: 0,
    valueOffset: 0, valueLength: values.length, indexOffset: 0, indexLength: indices.length, validSampleCount, available: roots.length > 0 }] }
}

/** Decode a materialized forest without recomputing scoped carry or signed
 * aggregation. Structural totals, source identities and ordering are checked. */
export function* decodeChartHierarchySteps(plan: ChartHierarchyPlan, candidate: ChartStatisticsOutput,
  context: ChartComputePreparationContext = {}): Generator<void, ChartHierarchyData, void> {
  const output = readChartStatisticsOutput(candidate, plan.task.params)
  if (!output) mismatch('层级图计算结果的段或缓冲区协议无效。')
  const job = plan.task.params.jobs[0]
  if (!job || job.kind !== 'hierarchy' || plan.task.params.jobs.length !== 1 || job.keyCount !== plan.keys.length || job.labelCount !== plan.labels.length) {
    mismatch('层级图来源字典与专属任务不一致。')
  }
  const segment = output.segments[0]!, count = output.indices[0]!, rootCount = output.indices[1]!
  const roots: ChartHierarchyNode[] = [], nodes: ChartHierarchyNode[] = [], paths: string[][] = [], parents: number[] = [], sourceOrders: number[] = []
  const siblingKeys = new Map<number, Set<number>>(), previous = new Map<number, number>(), created = new Set<number>(), active: number[] = []
  const bulkBytes = output.values.byteLength + output.indices.byteLength
  let decodedBytes = 1024, work = 0
  for (let index = 0; index < count; index++) {
    const offset = 2 + index * 4, parent = output.indices[offset]!, key = output.indices[offset + 1]!, label = output.indices[offset + 2]!, sourceOrder = output.indices[offset + 3]!
    const value = output.values[index * 2]!, selfValue = output.values[index * 2 + 1]!
    if (!Number.isFinite(value) || value <= 0 || !Number.isFinite(selfValue) || key >= plan.keys.length || label >= plan.labels.length
      || parent !== CHART_HIERARCHY_ROOT && (parent >= index || sourceOrders[parent]! >= sourceOrder)
      || sourceOrder >= segment.validSampleCount * job.hierarchyColumns.length || created.has(sourceOrder)) mismatch('层级图节点的权重、字典、父级或来源顺序无效。')
    const keys = siblingKeys.get(parent) ?? new Set<number>()
    if (keys.has(key)) mismatch('层级图结果包含重复的完整路径。')
    keys.add(key); siblingKeys.set(parent, keys); created.add(sourceOrder)
    if (parent === CHART_HIERARCHY_ROOT) active.length = 0
    else {
      while (active.length && active.at(-1) !== parent) active.pop()
      if (!active.length) mismatch('层级图节点没有保持深度优先顺序。')
    }
    active.push(index)
    if (active.length > job.hierarchyColumns.length) mismatch('层级图节点深度超过所选层级字段。')
    const prior = previous.get(parent)
    if (prior !== undefined && (nodes[prior]!.value < value || nodes[prior]!.value === value && sourceOrders[prior]! > sourceOrder)) mismatch('层级图同级节点没有保持权重与首次来源的稳定顺序。')
    previous.set(parent, index)
    const path = parent === CHART_HIERARCHY_ROOT ? [plan.labels[label]!] : [...nodes[parent]!.path, plan.labels[label]!]
    const tokenPath = parent === CHART_HIERARCHY_ROOT ? [plan.keys[key]!] : [...paths[parent]!, plan.keys[key]!]
    const id = JSON.stringify(tokenPath), formattedValue = plan.formatter(value), formattedSelfValue = plan.formatter(selfValue)
    decodedBytes += 256 + (path.length + tokenPath.length) * 16 + (id.length + formattedValue.length + formattedSelfValue.length) * 2
    assertChartPreparationMemory(plan.retainedBytes + bulkBytes + decodedBytes, context)
    const node: ChartHierarchyNode = { id, name: plan.labels[label]!, path, value, selfValue, formattedValue, formattedSelfValue }
    nodes.push(node); paths.push(tokenPath); parents.push(parent); sourceOrders.push(sourceOrder)
    if (parent === CHART_HIERARCHY_ROOT) roots.push(node)
    else (nodes[parent]!.children ??= []).push(node)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  if (roots.length !== rootCount) mismatch('层级图根节点数量与协议不一致。')
  for (let index = nodes.length - 1; index >= 0; index--) {
    const node = nodes[index]!
    let childrenValue = 0
    for (const child of node.children ?? []) {
      childrenValue += child.value
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    if (!Number.isFinite(childrenValue) || node.selfValue + childrenValue !== node.value) mismatch('层级图父节点的自身权重与子节点合计不一致。')
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const paletteOrder: string[] = [], paletteRoots = new Set<number>()
  let lastOrder = -1
  for (let paletteIndex = 0; paletteIndex < rootCount; paletteIndex++) {
    const root = output.indices[2 + count * 4 + paletteIndex]!
    if (root >= count || parents[root] !== CHART_HIERARCHY_ROOT || paletteRoots.has(root) || sourceOrders[root]! <= lastOrder) mismatch('层级图根节点配色身份没有保持首次来源顺序。')
    paletteRoots.add(root); lastOrder = sourceOrders[root]!
    nodes[root]!.paletteIndex = paletteIndex; paletteOrder.push(nodes[root]!.id)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return { valueFieldName: plan.valueFieldName, nodes: roots, paletteOrder,
    ...(segment.available ? {} : { diagnostic: '层级图没有最终为正的数值路径。' }) }
}
