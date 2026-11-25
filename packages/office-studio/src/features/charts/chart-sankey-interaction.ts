import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { chartColorOpacity, type ChartVector } from './chart-geometry'
import { chartColorChannels } from './chart-color'
import type { ChartSankeyLayout } from './chart-sankey-layout'
import type { ChartSankeyScene } from './chart-sankey-draw'

export type ChartSankeyHit = { kind: 'node'; index: number } | { kind: 'link'; index: number }
export interface ChartSankeyHitIndex {
  plot: RenderRect
  nodes: Array<{ index: number; bounds: RenderRect }>
  links: Array<{ index: number; bounds: RenderRect; band: readonly ChartVector[] }>
}
export interface ChartSankeyTopology {
  incoming: number[][]
  outgoing: number[][]
  links: Map<number, { source: number; target: number }>
}
export interface ChartSankeyTrajectory { nodes: Set<number>; links: Set<number> }
/** Stripped from ECharts options and cache signatures. Its owner is the
 * renderer entry, which releases it on source/config/size changes or dispose. */
export interface ChartSankeyPointerRuntime {
  mode: 'sankey'
  groupId: string
  zOffset: number
  scene: ChartSankeyScene
  /** Unfiltered view positions survive legend filtering and subsequent drags. */
  baseScene?: ChartSankeyScene
  hitIndex: ChartSankeyHitIndex
  topology: ChartSankeyTopology
}

export function createChartSankeyHitIndex(layout: ChartSankeyLayout): ChartSankeyHitIndex {
  return {
    plot: { ...layout.plot },
    nodes: layout.nodes.filter(node => node.width > 0 && node.height > 0).map(node => ({ index: node.index,
      bounds: { x: node.x, y: node.y, width: node.width, height: node.height } })),
    links: layout.links.filter(link => link.width > 0 && link.band.length >= 3).map(link => {
      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity
      for (const point of link.band) { minX = Math.min(minX, point.x); minY = Math.min(minY, point.y); maxX = Math.max(maxX, point.x); maxY = Math.max(maxY, point.y) }
      return { index: link.index, bounds: { x: minX, y: minY, width: maxX - minX, height: maxY - minY }, band: link.band }
    }),
  }
}

/** Node rectangles have priority. A band's actual filled polygon determines
 * hits; neither its bounding rectangle nor its old center line is a hit. */
export function hitChartSankey(index: ChartSankeyHitIndex, point: ChartVector | undefined): ChartSankeyHit | undefined {
  if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y) || !insideRect(index.plot, point)) return undefined
  for (let cursor = index.nodes.length - 1; cursor >= 0; cursor--) {
    const node = index.nodes[cursor]!
    if (insideRect(node.bounds, point)) return { kind: 'node', index: node.index }
  }
  for (let cursor = index.links.length - 1; cursor >= 0; cursor--) {
    const link = index.links[cursor]!
    if (insideRect(link.bounds, point) && insidePolygon(link.band, point)) return { kind: 'link', index: link.index }
  }
  return undefined
}

export function createChartSankeyTopology(layout: ChartSankeyLayout): ChartSankeyTopology {
  const topology: ChartSankeyTopology = { incoming: layout.nodes.map(() => []), outgoing: layout.nodes.map(() => []), links: new Map() }
  for (const link of layout.links) {
    if (link.width <= 0) continue
    topology.links.set(link.index, { source: link.sourceIndex, target: link.targetIndex })
    topology.incoming[link.targetIndex]!.push(link.index); topology.outgoing[link.sourceIndex]!.push(link.index)
  }
  return topology
}

/** Fixed trajectory semantics: upstream and downstream directed walks.
 * Hovering A→B keeps ancestors of A and descendants of B; the sibling
 * branches of A do not become part of the selected link's trajectory. */
export function chartSankeyTrajectory(topology: ChartSankeyTopology, hit: ChartSankeyHit | undefined): ChartSankeyTrajectory | undefined {
  if (!hit) return undefined
  const result: ChartSankeyTrajectory = { nodes: new Set(), links: new Set() }
  let upstream: number, downstream: number
  if (hit.kind === 'link') {
    const link = topology.links.get(hit.index)
    if (!link) return undefined
    upstream = link.source; downstream = link.target; result.links.add(hit.index)
  } else {
    if (!topology.incoming[hit.index] || !topology.outgoing[hit.index]) return undefined
    upstream = downstream = hit.index
  }
  for (const [start, incoming] of [[upstream, true], [downstream, false]] as const) {
    const queue = [start], visited = new Set<number>(queue)
    for (let cursor = 0; cursor < queue.length; cursor++) {
      const node = queue[cursor]!; result.nodes.add(node)
      for (const linkIndex of incoming ? topology.incoming[node]! : topology.outgoing[node]!) {
        const link = topology.links.get(linkIndex)!
        const other = incoming ? link.source : link.target
        result.links.add(linkIndex); result.nodes.add(other)
        if (!visited.has(other)) { visited.add(other); queue.push(other) }
      }
    }
  }
  return result
}

/** The base scene is immutable. Fade multiplies existing author alpha,
 * including independent outlines, textures and labels, then restores exactly. */
export function chartSankeyFocusedCommands(scene: ChartSankeyScene, trajectory?: ChartSankeyTrajectory, hit?: ChartSankeyHit): DrawCommand[] {
  if (!trajectory) return scene.commands
  const commands: DrawCommand[] = []
  for (const mark of scene.marks) {
    const focused = (mark.kind === 'node' ? trajectory.nodes : trajectory.links).has(mark.index)
    for (const command of mark.commands) commands.push(focused
      ? mark.kind === 'link' && hit?.kind === 'link' && hit.index === mark.index ? scaleCommand(command, 2.5) : command
      : scaleCommand(command, 0.15))
  }
  return commands
}

function scaleCommand(command: DrawCommand, opacity: number): DrawCommand {
  if (command.kind === 'rect') return { ...command, fill: scaleColor(command.fill, opacity),
    ...(command.stroke === undefined ? {} : { stroke: scaleColor(command.stroke, opacity) }) }
  if (command.kind === 'polygon') return { ...command, fill: scaleColor(command.fill, opacity) }
  if (command.kind === 'line' || command.kind === 'text') return { ...command, color: scaleColor(command.color, opacity) }
  return { ...command, opacity: (command.opacity ?? 1) * opacity }
}

function scaleColor(color: string, scale: number): string {
  if (scale <= 1) return chartColorOpacity(color, scale)
  const [red, green, blue, alpha] = chartColorChannels(color)
  return `rgba(${Math.round(red * 255)}, ${Math.round(green * 255)}, ${Math.round(blue * 255)}, ${Math.min(1, alpha * scale)})`
}

function insideRect(rect: RenderRect, point: ChartVector): boolean {
  return point.x >= rect.x && point.y >= rect.y && point.x <= rect.x + rect.width && point.y <= rect.y + rect.height
}
function insidePolygon(polygon: readonly ChartVector[], point: ChartVector): boolean {
  let inside = false
  for (let current = 0, previous = polygon.length - 1; current < polygon.length; previous = current++) {
    const first = polygon[previous]!, second = polygon[current]!
    const dx = second.x - first.x, dy = second.y - first.y, px = point.x - first.x, py = point.y - first.y
    const length = Math.hypot(dx, dy)
    if (length > 0 && Math.abs(dx * py - dy * px) <= 1e-7 * length
      && px * dx + py * dy >= 0 && px * dx + py * dy <= length * length) return true
    if ((first.y > point.y) !== (second.y > point.y)
      && point.x < first.x + (point.y - first.y) / (second.y - first.y) * (second.x - first.x)) inside = !inside
  }
  return inside
}
