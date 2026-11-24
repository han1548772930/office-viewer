import type { RenderRect } from '../../rendering/contracts'
import type { ChartGraphData } from './chart-data'
import type { ChartGraphLayout, ChartGraphLinkLayout, ChartGraphNodeLayout } from './chart-special-layout'
import { chartBezier, type ChartVector } from './chart-geometry'
import { finiteChartValue } from './chart-math'

export interface ChartSankeyNodeLayout extends ChartGraphNodeLayout {
  layer: number
  inValue: number
  outValue: number
}
export interface ChartSankeyLinkLayout extends ChartGraphLinkLayout {
  index: number
  sourceIndex: number
  targetIndex: number
  feedback: boolean
  /** The filled band, including exact vertical source and target ports. */
  band: ChartVector[]
  feedbackY?: number
}
export interface ChartSankeyLayout extends ChartGraphLayout {
  nodes: ChartSankeyNodeLayout[]
  links: ChartSankeyLinkLayout[]
  plot: RenderRect
  nodeArea: RenderRect
  cyclic: boolean
  gap: number
}
export interface ChartSankeyLayoutBudget { maxNodes: number; maxLinks: number; maxPoints: number }
/** Explicit allocation limits, pending large-data Worker and memory measurement. */
export const CHART_SANKEY_LAYOUT_BUDGET: Readonly<ChartSankeyLayoutBudget> = {
  maxNodes: 100_000, maxLinks: 100_000, maxPoints: 2_000_000,
}

interface LinkInput { index: number; source: number; target: number; value: number; normalized: number; feedback: boolean }
interface NodeFlow { incoming: number; outgoing: number; inLinks: number[]; outLinks: number[] }

/** Fixed Gje/ECharts default: horizontal, justified terminals, 20px nodes,
 * 8px gaps and 0.5 curve. Cyclic SCCs use real feedback bands instead of the
 * fixed converter's reciprocal-edge empty state. Source identities survive
 * vertical ordering; neither layout coordinates nor allocations are authored. */
export function chartSankeyLayout(graph: ChartGraphData, plot: RenderRect,
  budget: Readonly<ChartSankeyLayoutBudget> = CHART_SANKEY_LAYOUT_BUDGET): ChartSankeyLayout {
  validatePlot(plot)
  for (const limit of Object.values(budget)) if (!Number.isSafeInteger(limit) || limit < 1) throw new RangeError('桑基图布局资源预算须为正整数。')
  if (graph.nodes.length > budget.maxNodes || graph.links.length > budget.maxLinks) throw new RangeError('桑基图布局超过当前节点或边资源预算。')
  const empty: ChartSankeyLayout = { plot: { ...plot }, nodeArea: { ...plot }, nodes: [], links: [], cyclic: false, gap: 0 }
  if (!graph.nodes.length || plot.width === 0 || plot.height === 0) return empty
  const names = new Map<string, number>(), flows: NodeFlow[] = []
  graph.nodes.forEach((node, index) => {
    if (typeof node.name !== 'string' || !node.name || names.has(node.name)) throw new TypeError('桑基图节点须有唯一非空名称。')
    names.set(node.name, index); flows.push({ incoming: 0, outgoing: 0, inLinks: [], outLinks: [] })
  })
  let weightScale = 0
  const links: LinkInput[] = []
  graph.links.forEach((link, index) => {
    const source = names.get(link.source), target = names.get(link.target)
    if (source === undefined || target === undefined) throw new TypeError('桑基图流带引用不存在的节点。')
    if (!Number.isFinite(link.value) || link.value < 0) throw new RangeError('桑基图流带权重须为有限非负数值。')
    if (link.value === 0) return
    const position = links.length
    flows[source]!.outgoing = finiteChartValue(flows[source]!.outgoing + link.value, '桑基节点出流合计')
    flows[target]!.incoming = finiteChartValue(flows[target]!.incoming + link.value, '桑基节点入流合计')
    flows[source]!.outLinks.push(position); flows[target]!.inLinks.push(position)
    links.push({ index, source, target, value: link.value, normalized: 0, feedback: false })
    weightScale = Math.max(weightScale, link.value)
  })
  weightScale ||= 1
  for (const link of links) link.normalized = link.value / weightScale
  const components = connectedComponents(flows, links), componentByNode = new Uint32Array(graph.nodes.length)
  components.forEach((nodes, index) => nodes.forEach(node => { componentByNode[node] = index }))
  const outgoing = components.map(() => new Set<number>()), degree = new Uint32Array(components.length), layers = new Uint32Array(components.length)
  let cyclic = false, feedbackWeight = 0, largestFeedback = 0, feedbackCount = 0
  for (const link of links) {
    const source = componentByNode[link.source]!, target = componentByNode[link.target]!
    if (source === target) {
      link.feedback = true; cyclic = true; feedbackCount++
      feedbackWeight += link.normalized; largestFeedback = Math.max(largestFeedback, link.normalized)
    } else if (!outgoing[source]!.has(target)) { outgoing[source]!.add(target); degree[target]!++ }
  }
  const queue: number[] = []
  degree.forEach((count, index) => { if (!count) queue.push(index) })
  let maxLayer = 0
  for (let cursor = 0; cursor < queue.length; cursor++) {
    const source = queue[cursor]!
    maxLayer = Math.max(maxLayer, layers[source]!)
    for (const target of outgoing[source]!) {
      layers[target] = Math.max(layers[target]!, layers[source]! + 1)
      degree[target]!--; if (!degree[target]) queue.push(target)
    }
  }
  // Terminals in disconnected short branches still occupy the final column.
  layers.forEach((_layer, index) => { if (!outgoing[index]!.size) layers[index] = maxLayer })
  const columns: number[][] = Array.from({ length: maxLayer + 1 }, () => [])
  graph.nodes.forEach((_node, index) => { columns[layers[componentByNode[index]!]!]!.push(index) })
  const estimatedPoints = links.reduce((total, link) => total + (link.feedback ? 114 : 51), 0)
  if (estimatedPoints > budget.maxPoints) throw new RangeError('桑基图流带几何超过当前资源预算。')
  const largestColumn = columns.reduce((maximum, column) => Math.max(maximum, column.length), 0)
  const gap = Math.min(8, plot.height / Math.max(2, 4 * (largestColumn + feedbackCount + 1)),
    cyclic ? plot.width / (16 * columns.length) : 8)
  const nodeWidth = Math.min(20, plot.width / Math.max(1, cyclic ? columns.length * 2 : maxLayer * 1.5 + 1))
  let scale = Infinity
  for (const column of columns) {
    const weight = column.reduce((sum, index) => sum + Math.max(flows[index]!.incoming, flows[index]!.outgoing) / weightScale, 0)
    const available = plot.height - Math.max(0, column.length - 1) * gap - (feedbackCount ? feedbackCount + 2 : 0) * gap
    if (weight + feedbackWeight > 0) scale = Math.min(scale, available / (weight + feedbackWeight))
  }
  if (largestFeedback > 0) scale = Math.min(scale, (plot.width / (4 * columns.length) - gap) / largestFeedback)
  if (!Number.isFinite(scale)) scale = 0
  scale = Math.max(0, scale)
  const margin = cyclic ? largestFeedback * scale + gap : 0
  // Two clear gaps below the lanes leave room for both quarter-circle turns.
  const top = feedbackCount ? feedbackWeight * scale + (feedbackCount + 2) * gap : 0
  const nodeArea = { x: plot.x + margin, y: plot.y + top, width: Math.max(0, plot.width - margin * 2), height: Math.max(0, plot.height - top) }
  const nodes: ChartSankeyNodeLayout[] = graph.nodes.map((node, index) => {
    const flow = flows[index]!, value = Math.max(flow.incoming, flow.outgoing), layer = layers[componentByNode[index]!]!
    return { name: node.name, value, inValue: flow.incoming, outValue: flow.outgoing, index, layer,
      x: nodeArea.x + (maxLayer ? layer / maxLayer * (nodeArea.width - nodeWidth) : (nodeArea.width - nodeWidth) / 2),
      y: nodeArea.y, width: nodeWidth, height: value / weightScale * scale }
  })
  for (const column of columns) {
    const height = column.reduce((sum, index) => sum + nodes[index]!.height, 0) + Math.max(0, column.length - 1) * gap
    let y = nodeArea.y + Math.max(0, (nodeArea.height - height) / 2)
    for (const index of column) { nodes[index]!.y = y; y += nodes[index]!.height + gap }
  }
  // Relax only the condensation DAG. Feedback loops do not pull their own
  // nodes into impossible ordering constraints or change source style indices.
  for (let iteration = 0; iteration < 32; iteration++) {
    const alpha = Math.pow(0.99, iteration + 1)
    relax(columns, nodes, flows, links, nodeArea, gap, alpha, false)
    relax(columns, nodes, flows, links, nodeArea, gap, alpha, true)
  }
  let feedbackOffset = plot.y + gap
  const outputLinks: ChartSankeyLinkLayout[] = links.map(link => {
    const width = link.normalized * scale, feedbackY = link.feedback ? feedbackOffset + width / 2 : undefined
    if (link.feedback) feedbackOffset += width + gap
    return { source: graph.nodes[link.source]!.name, target: graph.nodes[link.target]!.name, value: link.value,
      index: link.index, sourceIndex: link.source, targetIndex: link.target, colorIndex: link.source,
      feedback: link.feedback, width, points: [], band: [], ...(feedbackY === undefined ? {} : { feedbackY }) }
  })
  const result: ChartSankeyLayout = { plot: { ...plot }, nodeArea, nodes, links: outputLinks, cyclic, gap }
  routeLinks(result)
  return result
}

/** View-only drag keeps weights and source identities, and rebuilds ports and
 * affected bands without source reads, SCC discovery or layout relaxation. */
export function moveChartSankeyNode(layout: ChartSankeyLayout, index: number, position: ChartVector): ChartSankeyLayout {
  if (!Number.isSafeInteger(index) || !layout.nodes[index]) throw new RangeError('桑基图节点序号不存在。')
  if (!Number.isFinite(position.x) || !Number.isFinite(position.y)) throw new RangeError('桑基图节点位置须为有限数值。')
  const nodes = layout.nodes.map(node => ({ ...node })), node = nodes[index]!, area = layout.nodeArea
  node.x = Math.max(area.x, Math.min(area.x + area.width - node.width, position.x))
  node.y = Math.max(area.y, Math.min(area.y + area.height - node.height, position.y))
  const moved = { ...layout, nodes, links: layout.links.map(link => ({ ...link, points: [], band: [] })) }
  routeLinks(moved)
  return moved
}

function routeLinks(layout: ChartSankeyLayout): void {
  const inLinks: ChartSankeyLinkLayout[][] = layout.nodes.map(() => []), outLinks: ChartSankeyLinkLayout[][] = layout.nodes.map(() => [])
  for (const link of layout.links) { outLinks[link.sourceIndex]!.push(link); inLinks[link.targetIndex]!.push(link) }
  const startY = new Map<number, number>(), endY = new Map<number, number>()
  for (const node of layout.nodes) for (const [incoming, edges] of [[false, outLinks[node.index]!], [true, inLinks[node.index]!]] as const) {
    edges.sort((first, second) => {
      const a = layout.nodes[incoming ? first.sourceIndex : first.targetIndex]!, b = layout.nodes[incoming ? second.sourceIndex : second.targetIndex]!
      return a.y + a.height / 2 - b.y - b.height / 2 || first.index - second.index
    })
    const total = edges.reduce((sum, link) => sum + link.width, 0)
    let cursor = node.y + Math.max(0, (node.height - total) / 2)
    for (const link of edges) { (incoming ? endY : startY).set(link.index, cursor + link.width / 2); cursor += link.width }
  }
  for (const link of layout.links) {
    const source = layout.nodes[link.sourceIndex]!, target = layout.nodes[link.targetIndex]!
    const start = { x: source.x + source.width, y: startY.get(link.index)! }, end = { x: target.x, y: endY.get(link.index)! }
    if (link.feedback) {
      const radius = link.width / 2 + layout.gap
      const lane = Math.min(link.feedbackY!, Math.min(start.y, end.y) - radius * 2)
      link.points = feedbackPath(start, end, lane, radius)
      link.band = [...feedbackPath(start, end, lane, radius, link.width / 2), ...feedbackPath(start, end, lane, radius, -link.width / 2).reverse()]
    } else {
      const middle = (start.x + end.x) / 2, half = link.width / 2
      // Dragged ports may share an X coordinate. Keep an actual S curve rather
      // than reducing the filled flow to a line of zero area.
      const turn = end.x > start.x ? 0 : Math.max(source.width, target.width, layout.gap)
      const firstX = turn ? start.x + turn : middle, secondX = turn ? end.x - turn : middle
      link.points = chartBezier(start, { x: firstX, y: start.y }, { x: secondX, y: end.y }, end, 16)
      const boundary = (offset: number): ChartVector[] => chartBezier({ x: start.x, y: start.y + offset },
        { x: firstX, y: start.y + offset }, { x: secondX, y: end.y + offset }, { x: end.x, y: end.y + offset }, 16)
      link.band = [...boundary(-half), ...boundary(half).reverse()]
    }
  }
}

/** Four analytic quarter-circles keep the inner radius positive and ports
 * exactly vertical. Normal-offsetting a reversing cubic folds a feedback band. */
function feedbackPath(start: ChartVector, end: ChartVector, lane: number, radius: number, offset = 0): ChartVector[] {
  // These caps retain the clockwise top route after a horizontal node drag.
  // The initial/final straight sections still meet the exact vertical ports.
  const right = Math.max(start.x, end.x), left = Math.min(start.x, end.x)
  const points: ChartVector[] = [{ x: start.x, y: start.y + offset }], adjusted = radius + offset
  for (const [x, y, angle] of [[right, start.y - radius, Math.PI / 2], [right, lane + radius, 0],
    [left, lane + radius, -Math.PI / 2], [left, end.y - radius, -Math.PI]] as const) {
    for (let step = 0; step <= 8; step++) {
      const current = angle - step * Math.PI / 16
      points.push({ x: x + adjusted * Math.cos(current), y: y + adjusted * Math.sin(current) })
    }
  }
  points.push({ x: end.x, y: end.y + offset })
  return points
}

function relax(columns: number[][], nodes: ChartSankeyNodeLayout[], flows: NodeFlow[], links: LinkInput[],
  area: RenderRect, gap: number, alpha: number, incoming: boolean): void {
  const start = incoming ? 0 : columns.length - 1, end = incoming ? columns.length : -1, step = incoming ? 1 : -1
  for (let columnIndex = start; columnIndex !== end; columnIndex += step) {
    const column = columns[columnIndex]!
    for (const index of column) {
      let weighted = 0, weight = 0
      for (const position of incoming ? flows[index]!.inLinks : flows[index]!.outLinks) {
        const link = links[position]!
        if (link.feedback) continue
        const other = nodes[incoming ? link.source : link.target]!
        weighted += ((other.y - area.y + other.height / 2) / Math.max(Number.MIN_VALUE, area.height)) * link.normalized
        weight += link.normalized
      }
      if (weight > 0) {
        const node = nodes[index]!, center = area.y + weighted / weight * area.height
        node.y += (center - node.y - node.height / 2) * alpha
      }
    }
    column.sort((a, b) => nodes[a]!.y - nodes[b]!.y || a - b)
    let next = area.y
    for (const index of column) { const node = nodes[index]!; node.y = Math.max(node.y, next); next = node.y + node.height + gap }
    let bottom = area.y + area.height
    for (let cursor = column.length - 1; cursor >= 0; cursor--) {
      const node = nodes[column[cursor]!]!; node.y = Math.min(node.y, bottom - node.height); bottom = node.y - gap
    }
  }
}

function connectedComponents(flows: NodeFlow[], links: LinkInput[]): number[][] {
  const visited = new Uint8Array(flows.length), order: number[] = []
  for (let index = 0; index < flows.length; index++) {
    if (visited[index]) continue
    const stack = [{ index, next: 0 }]; visited[index] = 1
    while (stack.length) {
      const current = stack[stack.length - 1]!, edges = flows[current.index]!.outLinks
      if (current.next < edges.length) {
        const target = links[edges[current.next++]!]!.target
        if (!visited[target]) { visited[target] = 1; stack.push({ index: target, next: 0 }) }
      } else { order.push(current.index); stack.pop() }
    }
  }
  visited.fill(0)
  const components: number[][] = []
  for (let cursor = order.length - 1; cursor >= 0; cursor--) {
    const index = order[cursor]!
    if (visited[index]) continue
    const stack = [index], component: number[] = []; visited[index] = 1
    while (stack.length) {
      const current = stack.pop()!; component.push(current)
      for (const position of flows[current]!.inLinks) {
        const source = links[position]!.source
        if (!visited[source]) { visited[source] = 1; stack.push(source) }
      }
    }
    components.push(component)
  }
  return components
}

function validatePlot(plot: RenderRect): void {
  if (![plot.x, plot.y, plot.width, plot.height, plot.x + plot.width, plot.y + plot.height].every(Number.isFinite)
    || plot.width < 0 || plot.height < 0) throw new RangeError('桑基图绘图区须为有限非负矩形。')
}
