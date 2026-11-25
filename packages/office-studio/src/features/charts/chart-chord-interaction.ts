import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { chartColorOpacity, type ChartVector } from './chart-geometry'
import type { ChartChordLayout } from './chart-chord-layout'
import type { ChartChordScene } from './chart-chord-draw'

export type ChartChordHit = { kind: 'node'; index: number } | { kind: 'link'; index: number }
/** Text bounds come from the actual renderer, after its font/truncation and
 * transform are applied. Pure geometry never guesses the width of a label. */
export interface ChartChordLabelHit {
  index: number
  polygon: readonly ChartVector[]
  clip?: RenderRect
}
export interface ChartChordHitIndex {
  plot: RenderRect
  center: ChartVector
  innerRadius: number
  outerRadius: number
  nodes: Array<{ index: number; start: number; end: number }>
  links: Array<{ index: number; bounds: RenderRect; band: readonly ChartVector[] }>
  labels: ChartChordLabelHit[]
}
export interface ChartChordTopology {
  incident: number[][]
  links: Map<number, { source: number; target: number }>
}
export interface ChartChordFocusSelection { nodes: Set<number>; links: Set<number> }
export interface ChartChordContentTooltip {
  id: string
  kind: ChartChordHit['kind']
  index: number
  title: string
  lines: string[]
}
/** The renderer owns this transient scene/index. It is removed from setOption
 * and cache signatures, and released on source/config/size changes or dispose. */
export interface ChartChordPointerRuntime {
  mode: 'chord'
  groupId: string
  zOffset: number
  scene: ChartChordScene
  hitIndex: ChartChordHitIndex
  topology: ChartChordTopology
}

export function createChartChordHitIndex(layout: ChartChordLayout, labels: readonly ChartChordLabelHit[] = []): ChartChordHitIndex {
  return {
    plot: { ...layout.plot }, center: { ...layout.center }, innerRadius: layout.innerRadius, outerRadius: layout.outerRadius,
    nodes: layout.nodes.filter(node => Math.abs(node.end - node.start) > 0)
      .map(node => ({ index: node.index, start: node.start, end: node.end })),
    links: layout.innerRadius > 0 ? layout.links.filter(link => link.band.length >= 3
      && Math.abs(link.sourceEnd - link.sourceStart) > 0 && Math.abs(link.targetEnd - link.targetStart) > 0)
      .map(link => ({ index: link.index, bounds: { ...link.bounds }, band: link.band })) : [],
    labels: labels.map(label => ({ ...label, polygon: label.polygon.map(point => ({ ...point })), ...(label.clip ? { clip: { ...label.clip } } : {}) })),
  }
}

/** Label rectangles are measured by the renderer and may extend past the
 * plot. Arc hits take priority over overlapping ribbons, which use their
 * filled polygon rather than an enclosing rectangle or a center line. */
export function hitChartChord(index: ChartChordHitIndex, point?: ChartVector): ChartChordHit | undefined {
  if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y)) return undefined
  for (let cursor = index.labels.length - 1; cursor >= 0; cursor--) {
    const label = index.labels[cursor]!
    if ((!label.clip || insideRect(label.clip, point)) && insidePolygon(label.polygon, point)) return { kind: 'node', index: label.index }
  }
  if (!insideRect(index.plot, point)) return undefined
  const radius = Math.hypot(point.x - index.center.x, point.y - index.center.y)
  if (radius >= index.innerRadius && radius <= index.outerRadius && index.outerRadius > index.innerRadius) {
    const angle = Math.atan2(point.y - index.center.y, point.x - index.center.x)
    for (let cursor = index.nodes.length - 1; cursor >= 0; cursor--) {
      const node = index.nodes[cursor]!
      if (insideArc(node.start, node.end, angle)) return { kind: 'node', index: node.index }
    }
  }
  for (let cursor = index.links.length - 1; cursor >= 0; cursor--) {
    const link = index.links[cursor]!
    if (insideRect(link.bounds, point) && insidePolygon(link.band, point)) return { kind: 'link', index: link.index }
  }
  return undefined
}

export function createChartChordTopology(layout: ChartChordLayout): ChartChordTopology {
  const topology: ChartChordTopology = { incident: layout.nodes.map(() => []), links: new Map() }
  for (const link of layout.links) {
    topology.links.set(link.index, { source: link.sourceIndex, target: link.targetIndex })
    topology.incident[link.sourceIndex]!.push(link.index)
    topology.incident[link.targetIndex]!.push(link.index)
  }
  return topology
}

/** Chord is undirected. Adjacency is exactly one hop; it must not become the
 * directed transitive trajectory used by Sankey. Storage stays O(N + E). */
export function chartChordFocusSelection(topology: ChartChordTopology, focus: ChartChordScene['focus'], hit?: ChartChordHit): ChartChordFocusSelection | undefined {
  if (!hit || focus === 'none') return undefined
  const selected: ChartChordFocusSelection = { nodes: new Set(), links: new Set() }
  if (hit.kind === 'node') {
    const incident = topology.incident[hit.index]
    if (!incident) return undefined
    selected.nodes.add(hit.index)
    if (focus === 'adjacency') for (const linkIndex of incident) {
      const link = topology.links.get(linkIndex)!
      selected.links.add(linkIndex); selected.nodes.add(link.source); selected.nodes.add(link.target)
    }
  } else {
    const link = topology.links.get(hit.index)
    if (!link) return undefined
    selected.links.add(hit.index)
    if (focus === 'adjacency') { selected.nodes.add(link.source); selected.nodes.add(link.target) }
  }
  return selected
}

/** Base paint remains immutable. Focus keeps its independent author alpha;
 * unrelated fills, borders, patterns and labels multiply that alpha by .15. */
export function chartChordFocusedCommands(scene: ChartChordScene, selection?: ChartChordFocusSelection): DrawCommand[] {
  if (!selection) return scene.commands
  const commands: DrawCommand[] = []
  for (const mark of scene.marks) {
    const focused = (mark.kind === 'node' ? selection.nodes : selection.links).has(mark.index)
    for (const command of mark.commands) commands.push(focused ? command : fadeCommand(command))
  }
  return commands
}

/** Aggregate formatters have already run once after link/node totals. A
 * pointer move never reparses cells or formats each raw contributing row. */
export function chartChordContentTooltip(scene: ChartChordScene, hit?: ChartChordHit): ChartChordContentTooltip | undefined {
  if (!hit) return undefined
  const label = scene.weightMode === 'count' ? '计数' : scene.valueFieldName || '权重'
  if (hit.kind === 'node') {
    const node = scene.layout.nodes[hit.index]
    return node ? { id: node.id, ...hit, title: node.name, lines: [`${label}：${node.formattedValue}`] } : undefined
  }
  const link = scene.layout.links[hit.index]
  return link ? { id: link.id, ...hit, title: `${link.source} — ${link.target}`, lines: [`${label}：${link.formattedValue}`] } : undefined
}

function fadeCommand(command: DrawCommand): DrawCommand {
  if (command.kind === 'rect') return { ...command, fill: chartColorOpacity(command.fill, 0.15),
    ...(command.stroke === undefined ? {} : { stroke: chartColorOpacity(command.stroke, 0.15) }) }
  if (command.kind === 'polygon') return { ...command, fill: chartColorOpacity(command.fill, 0.15) }
  if (command.kind === 'line' || command.kind === 'text') return { ...command, color: chartColorOpacity(command.color, 0.15) }
  return { ...command, opacity: (command.opacity ?? 1) * 0.15 }
}
function insideArc(start: number, end: number, angle: number): boolean {
  const circle = Math.PI * 2, span = end - start, direction = span >= 0 ? 1 : -1
  const progress = ((angle - start) * direction % circle + circle) % circle
  return progress <= Math.abs(span) + 1e-10 || circle - progress < 1e-10
}
function insideRect(rect: RenderRect, point: ChartVector): boolean {
  return point.x >= rect.x && point.y >= rect.y && point.x <= rect.x + rect.width && point.y <= rect.y + rect.height
}
function insidePolygon(polygon: readonly ChartVector[], point: ChartVector): boolean {
  let inside = false
  for (let current = 0, previous = polygon.length - 1; current < polygon.length; previous = current++) {
    const first = polygon[previous]!, second = polygon[current]!
    const dx = second.x - first.x, dy = second.y - first.y, px = point.x - first.x, py = point.y - first.y, length = Math.hypot(dx, dy)
    if (length > 0 && Math.abs(dx * py - dy * px) <= 1e-7 * length
      && px * dx + py * dy >= 0 && px * dx + py * dy <= length * length) return true
    if ((first.y > point.y) !== (second.y > point.y)
      && point.x < first.x + (point.y - first.y) / (second.y - first.y) * (second.x - first.x)) inside = !inside
  }
  return inside
}
