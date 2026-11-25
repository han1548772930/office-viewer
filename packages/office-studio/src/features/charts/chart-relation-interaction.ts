import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { chartColorChannels } from './chart-color'
import { chartColorOpacity, type ChartVector } from './chart-geometry'
import type { ChartRelationGraphLayout } from './chart-relation-layout'
import type { ChartRelationScene, ChartRelationViewport } from './chart-relation-draw'

export type ChartRelationHit = { kind: 'node'; index: number } | { kind: 'link'; index: number }
export interface ChartRelationLabelHit { index: number; polygon: readonly ChartVector[]; clip?: RenderRect }
interface HitMark { index: number; bounds: RenderRect; commands: readonly DrawCommand[] }
export interface ChartRelationHitIndex { plot: RenderRect; nodes: HitMark[]; links: HitMark[]; labels: ChartRelationLabelHit[] }
export interface ChartRelationTopology { incident: number[][]; links: Map<number, { source: number; target: number }> }
export interface ChartRelationFocusSelection { nodes: Set<number>; links: Set<number> }
export interface ChartRelationContentTooltip { id: string; kind: ChartRelationHit['kind']; index: number; title: string; lines: string[] }
export interface ChartRelationPointerRuntime {
  mode: 'relation'; groupId: string; zOffset: number; scene: ChartRelationScene; baseScene: ChartRelationScene
  hitIndex: ChartRelationHitIndex; topology: ChartRelationTopology; viewport: ChartRelationViewport
}

/** Node and edge hits use the actual submitted triangles, rectangles and
 * stroke segments. Label polygons are supplied by the real text renderer. */
export function createChartRelationHitIndex(scene: ChartRelationScene, labels: readonly ChartRelationLabelHit[] = []): ChartRelationHitIndex {
  const index: ChartRelationHitIndex = { plot: { ...scene.layout.plot }, nodes: [], links: [],
    labels: labels.map(label => ({ ...label, polygon: label.polygon.map(point => ({ ...point })), ...(label.clip ? { clip: { ...label.clip } } : {}) })) }
  const nodeCommands = new Map<number, DrawCommand[]>()
  for (const mark of scene.marks) {
    const commands = mark.commands.filter(command => command.kind !== 'text' && command.kind !== 'image' && visibleCommand(command))
    if (!commands.length) continue
    if (mark.kind === 'node') {
      const retained = nodeCommands.get(mark.index) ?? []
      for (const command of commands) retained.push(command)
      nodeCommands.set(mark.index, retained)
    } else index.links.push({ index: mark.index, bounds: commandBounds(commands), commands })
  }
  for (const [id, commands] of nodeCommands) index.nodes.push({ index: id, bounds: commandBounds(commands), commands })
  return index
}

/** A two-CSS-pixel stroke tolerance is applied to the real path, including
 * .3 curves and self-loops; an enclosing box alone never counts as a hit. */
export function hitChartRelation(index: ChartRelationHitIndex, point?: ChartVector): ChartRelationHit | undefined {
  if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y)) return undefined
  for (let cursor = index.labels.length - 1; cursor >= 0; cursor--) {
    const label = index.labels[cursor]!
    if ((!label.clip || insideRect(label.clip, point)) && insidePolygon(label.polygon, point)) return { kind: 'node', index: label.index }
  }
  if (!insideRect(index.plot, point)) return undefined
  for (const kind of ['node', 'link'] as const) {
    const marks = kind === 'node' ? index.nodes : index.links
    for (let cursor = marks.length - 1; cursor >= 0; cursor--) {
      const mark = marks[cursor]!
      if (insideRect(mark.bounds, point) && mark.commands.some(command => hitCommand(command, point))) return { kind, index: mark.index }
    }
  }
  return undefined
}

export function createChartRelationTopology(layout: ChartRelationGraphLayout): ChartRelationTopology {
  const topology: ChartRelationTopology = { incident: layout.nodes.map(() => []), links: new Map() }
  for (const link of layout.links) {
    topology.links.set(link.index, { source: link.sourceIndex, target: link.targetIndex })
    topology.incident[link.sourceIndex]!.push(link.index)
    if (link.sourceIndex !== link.targetIndex) topology.incident[link.targetIndex]!.push(link.index)
  }
  return topology
}

/** Adjacency is one undirected hop. Duplicate edges and self-loops retain
 * their numeric identities; focus never walks a transitive graph trajectory. */
export function chartRelationFocusSelection(topology: ChartRelationTopology, enabled: boolean, hit?: ChartRelationHit): ChartRelationFocusSelection | undefined {
  if (!enabled || !hit) return undefined
  const selected: ChartRelationFocusSelection = { nodes: new Set(), links: new Set() }
  if (hit.kind === 'node') {
    const incident = topology.incident[hit.index]
    if (!incident) return undefined
    selected.nodes.add(hit.index)
    for (const id of incident) {
      const edge = topology.links.get(id)!
      selected.links.add(id); selected.nodes.add(edge.source); selected.nodes.add(edge.target)
    }
  } else {
    const edge = topology.links.get(hit.index)
    if (!edge) return undefined
    selected.links.add(hit.index); selected.nodes.add(edge.source); selected.nodes.add(edge.target)
  }
  return selected
}

/** Immutable base alpha is multiplied for unrelated fills, borders, patterns
 * and labels; leaving restores the exact author paint. */
export function chartRelationFocusedCommands(scene: ChartRelationScene, selected?: ChartRelationFocusSelection): DrawCommand[] {
  if (!selected) return scene.commands
  const commands: DrawCommand[] = []
  for (const mark of scene.marks) {
    const focused = (mark.kind === 'node' ? selected.nodes : selected.links).has(mark.index)
    for (const command of mark.commands) commands.push(focused ? command : fadeCommand(command))
  }
  return commands
}
export function chartRelationContentTooltip(scene: ChartRelationScene, hit?: ChartRelationHit): ChartRelationContentTooltip | undefined {
  if (!hit) return undefined
  if (hit.kind === 'node') {
    const node = scene.layout.nodes[hit.index], paint = scene.nodes[hit.index]
    return node && paint ? { ...hit, id: node.id, title: node.name,
      lines: [...(node.category ? [`分类：${node.category}`] : []), `${scene.valueFieldName}：${paint.formattedValue}`] } : undefined
  }
  const link = scene.layout.links[hit.index]
  return link ? { ...hit, id: `link:${link.index}`, title: `${link.source} → ${link.target}`,
    lines: [`${scene.valueFieldName}：${scene.linkFormattedValues[link.index] ?? link.value}`] } : undefined
}

function visibleCommand(command: DrawCommand): boolean {
  if (command.kind === 'rect') return chartColorChannels(command.fill)[3] > 0
    || command.stroke !== undefined && (command.strokeWidth ?? 1) > 0 && chartColorChannels(command.stroke)[3] > 0
  if (command.kind === 'polygon') return chartColorChannels(command.fill)[3] > 0
  if (command.kind === 'line') return (command.width ?? 1) > 0 && chartColorChannels(command.color)[3] > 0
  return false
}
function hitCommand(command: DrawCommand, point: ChartVector): boolean {
  if (command.clip && !insideRect(command.clip, point)) return false
  if (command.kind === 'polygon') return insidePolygon(command.points, point)
  if (command.kind === 'line') return segmentDistance(point, { x: command.x1, y: command.y1 }, { x: command.x2, y: command.y2 }) <= (command.width ?? 1) / 2 + 2
  if (command.kind === 'rect') {
    if (chartColorChannels(command.fill)[3] > 0 && insideRect(command, point)) return true
    if (command.stroke === undefined || chartColorChannels(command.stroke)[3] <= 0) return false
    const polygon = [{ x: command.x, y: command.y }, { x: command.x + command.width, y: command.y },
      { x: command.x + command.width, y: command.y + command.height }, { x: command.x, y: command.y + command.height }]
    return polygon.some((item, index) => segmentDistance(point, item, polygon[(index + 1) % 4]!) <= (command.strokeWidth ?? 1) / 2 + 2)
  }
  return false
}
function commandBounds(commands: readonly DrawCommand[]): RenderRect {
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity
  const add = (point: ChartVector, tolerance = 0): void => {
    left = Math.min(left, point.x - tolerance); right = Math.max(right, point.x + tolerance)
    top = Math.min(top, point.y - tolerance); bottom = Math.max(bottom, point.y + tolerance)
  }
  for (const command of commands) {
    if (command.kind === 'polygon') for (const point of command.points) add(point)
    else if (command.kind === 'line') { const tolerance = (command.width ?? 1) / 2 + 2;
      add({ x: command.x1, y: command.y1 }, tolerance); add({ x: command.x2, y: command.y2 }, tolerance) }
    else if (command.kind === 'rect') { const tolerance = command.stroke ? (command.strokeWidth ?? 1) / 2 + 2 : 0;
      add(command, tolerance); add({ x: command.x + command.width, y: command.y + command.height }, tolerance) }
  }
  return { x: left, y: top, width: right - left, height: bottom - top }
}
function segmentDistance(point: ChartVector, first: ChartVector, second: ChartVector): number {
  const dx = second.x - first.x, dy = second.y - first.y, squared = dx * dx + dy * dy
  const t = squared > 0 ? Math.max(0, Math.min(1, ((point.x - first.x) * dx + (point.y - first.y) * dy) / squared)) : 0
  return Math.hypot(point.x - first.x - dx * t, point.y - first.y - dy * t)
}
function insideRect(rect: RenderRect, point: ChartVector): boolean {
  return point.x >= rect.x && point.y >= rect.y && point.x <= rect.x + rect.width && point.y <= rect.y + rect.height
}
function insidePolygon(polygon: readonly ChartVector[], point: ChartVector): boolean {
  let inside = false
  for (let current = 0, previous = polygon.length - 1; current < polygon.length; previous = current++) {
    const first = polygon[previous]!, second = polygon[current]!
    if (segmentDistance(point, first, second) <= 1e-7) return true
    if ((first.y > point.y) !== (second.y > point.y)
      && point.x < first.x + (point.y - first.y) / (second.y - first.y) * (second.x - first.x)) inside = !inside
  }
  return inside
}
function fadeCommand(command: DrawCommand): DrawCommand {
  if (command.kind === 'rect') return { ...command, fill: chartColorOpacity(command.fill, 0.1),
    ...(command.stroke === undefined ? {} : { stroke: chartColorOpacity(command.stroke, 0.1) }) }
  if (command.kind === 'polygon') return { ...command, fill: chartColorOpacity(command.fill, 0.1) }
  if (command.kind === 'line' || command.kind === 'text') return { ...command, color: chartColorOpacity(command.color, 0.1) }
  return { ...command, opacity: (command.opacity ?? 1) * 0.1 }
}
