import type { ChartModel, ChartSemanticOptions } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartGraphData, ChartHierarchyNode } from './chart-data'
import type { ChartData } from './chart-engine'
import { chartBezier, chartPolar, type ChartVector } from './chart-geometry'
import type { ChartTreemapRenderNode } from './chart-treemap'
export { chartSankeyLayout } from './chart-sankey-layout'
export { chartTreemapLayout } from './chart-treemap-layout'
export { chartSunburstLayout } from './chart-sunburst-layout'
export { chartChordLayout } from './chart-chord-layout'
export type { ChartChordLayout, ChartChordNodeLayout, ChartChordLinkLayout } from './chart-chord-layout'
export type { ChartSunburstSector } from './chart-sunburst-layout'

export interface ChartGraphNodeLayout extends ChartVector { name: string; value: number; width: number; height: number; index: number }
export interface ChartGraphLinkLayout { source: string; target: string; value: number; width: number; points: ChartVector[]; colorIndex: number }
export interface ChartGraphLayout { nodes: ChartGraphNodeLayout[]; links: ChartGraphLinkLayout[] }
export interface ChartTreeRect extends RenderRect { node: ChartHierarchyNode; renderNode: ChartTreemapRenderNode; pointIndex: number; rootIndex: number; depth: number; colorIndex: number; path: string[]; labelHeight: number }
export interface ChartCloudWord extends RenderRect { text: string; value: number; categoryIndex: number; fontSize: number; rotation: number }

/** Fixed-order, deterministic graph coordinates, shared by both Canvas APIs. */
export function chartRelationLayout(graph: ChartGraphData, chart: ChartModel, plot: RenderRect): ChartGraphLayout {
  const options = chart.options?.relation, count = graph.nodes.length
  const center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }
  let maxWeight = 1
  for (const node of graph.nodes) maxWeight = Math.max(maxWeight, node.value)
  const nodes = graph.nodes.map((node, index) => {
    const point = count === 1 ? center : chartPolar(center, Math.min(plot.width, plot.height) * 0.36, -Math.PI / 2 + index * Math.PI * 2 / count)
    const size = options?.useValueAsSymbolSize ? 8 + Math.sqrt(node.value / maxWeight) * 24 : 16
    return { ...node, ...point, width: size, height: size, index }
  })
  if (options?.layout === 'force' && nodes.length > 1) {
    const indexes = new Map(nodes.map((node, index) => [node.name, index]))
    const repulsion = options.force?.repulsion ?? 80, gravity = options.force?.gravity ?? 0.12, length = options.force?.edgeLength ?? 72
    // Spatial bins bound the repulsion work while retaining a deterministic solver.
    const cellSize = Math.max(24, Math.sqrt(repulsion) * 5)
    for (let iteration = 0; iteration < 64; iteration++) {
      const bins = new Map<string, number[]>()
      nodes.forEach((node, index) => { const key = `${Math.floor(node.x / cellSize)}:${Math.floor(node.y / cellSize)}`, list = bins.get(key) ?? []; list.push(index); bins.set(key, list) })
      const movements = nodes.map(node => ({ x: (center.x - node.x) * gravity * 0.06, y: (center.y - node.y) * gravity * 0.06 }))
      nodes.forEach((node, index) => {
        const col = Math.floor(node.x / cellSize), row = Math.floor(node.y / cellSize)
        for (let dx = -2; dx <= 2; dx++) for (let dy = -2; dy <= 2; dy++) for (const otherIndex of bins.get(`${col + dx}:${row + dy}`) ?? []) {
          if (otherIndex <= index) continue
          const other = nodes[otherIndex]!, x = node.x - other.x || 0.01, y = node.y - other.y || 0.01, distance = Math.max(1, Math.hypot(x, y))
          const force = Math.min(8, repulsion / (distance * distance))
          movements[index]!.x += x / distance * force; movements[index]!.y += y / distance * force
          movements[otherIndex]!.x -= x / distance * force; movements[otherIndex]!.y -= y / distance * force
        }
      })
      for (const link of graph.links) {
        const a = indexes.get(link.source)!, b = indexes.get(link.target)!
        if (a === b) continue
        const first = nodes[a]!, second = nodes[b]!, x = second.x - first.x, y = second.y - first.y, distance = Math.max(1, Math.hypot(x, y))
        const force = (distance - length) * 0.025
        movements[a]!.x += x / distance * force; movements[a]!.y += y / distance * force
        movements[b]!.x -= x / distance * force; movements[b]!.y -= y / distance * force
      }
      const cooling = 1 - iteration / 80
      nodes.forEach((node, index) => {
        node.x = Math.max(plot.x + node.width, Math.min(plot.x + plot.width - node.width, node.x + Math.max(-10, Math.min(10, movements[index]!.x)) * cooling))
        node.y = Math.max(plot.y + node.height, Math.min(plot.y + plot.height - node.height, node.y + Math.max(-10, Math.min(10, movements[index]!.y)) * cooling))
      })
    }
  }
  const byName = new Map(nodes.map(node => [node.name, node]))
  const links = graph.links.map(link => {
    const start = byName.get(link.source)!, end = byName.get(link.target)!
    const self = start === end
    const points = self ? chartBezier({ x: start.x - 4, y: start.y - 4 }, { x: start.x - 36, y: start.y - 44 }, { x: start.x + 36, y: start.y - 44 }, { x: start.x + 4, y: start.y - 4 })
      : [start, end]
    return { ...link, points, width: 1 + Math.sqrt(link.value / maxWeight) * 3, colorIndex: start.index }
  })
  return { nodes, links }
}

export function chartRibbon(points: readonly ChartVector[], width: number): ChartVector[] {
  const first: ChartVector[] = [], second: ChartVector[] = []
  points.forEach((point, index) => {
    const previous = points[Math.max(0, index - 1)]!, next = points[Math.min(points.length - 1, index + 1)]!
    const dx = next.x - previous.x, dy = next.y - previous.y, distance = Math.max(1e-9, Math.hypot(dx, dy))
    const x = -dy / distance * width / 2, y = dx / distance * width / 2
    first.push({ x: point.x + x, y: point.y + y }); second.push({ x: point.x - x, y: point.y - y })
  })
  return [...first, ...second.reverse()]
}

export function chartWordCloudLayout(data: ChartData, options: ChartSemanticOptions['wordCloud'], plot: RenderRect): ChartCloudWord[] {
  const words = data.categories.map((text, categoryIndex) => ({ text, categoryIndex, value: data.series[0]?.values[categoryIndex] ?? 0 })).filter(word => word.text && word.value > 0)
    .sort((a, b) => b.value - a.value || a.categoryIndex - b.categoryIndex)
  if (!words.length) return []
  const largest = words[0]!.value, minSize = options?.minFontSize ?? 10, maxSize = options?.maxFontSize ?? 54, rotation = options?.rotation ?? 0
  const radians = rotation * Math.PI / 180, cos = Math.abs(Math.cos(radians)), sin = Math.abs(Math.sin(radians))
  // Font metrics are conservatively estimated here; the measured Canvas/Worker layout remains unaccepted.
  for (let expansion = 1; expansion <= 32; expansion *= 1.6) {
    const width = plot.width * expansion, height = plot.height * expansion, bins = new Map<string, ChartCloudWord[]>(), result: ChartCloudWord[] = [], cellSize = 24
    const collides = (rect: RenderRect): boolean => {
      for (let col = Math.floor(rect.x / cellSize); col <= Math.floor((rect.x + rect.width) / cellSize); col++) for (let row = Math.floor(rect.y / cellSize); row <= Math.floor((rect.y + rect.height) / cellSize); row++) {
        for (const word of bins.get(`${col}:${row}`) ?? []) if (rect.x < word.x + word.width + 2 && rect.x + rect.width + 2 > word.x && rect.y < word.y + word.height + 2 && rect.y + rect.height + 2 > word.y) return true
      }
      return false
    }
    const add = (word: ChartCloudWord): void => {
      result.push(word)
      for (let col = Math.floor(word.x / cellSize); col <= Math.floor((word.x + word.width) / cellSize); col++) for (let row = Math.floor(word.y / cellSize); row <= Math.floor((word.y + word.height) / cellSize); row++) {
        const key = `${col}:${row}`, entries = bins.get(key) ?? []; entries.push(word); bins.set(key, entries)
      }
    }
    let failed = false
    const candidates = options?.repeat ? [...words, ...words.slice(0, Math.min(words.length, 80))] : words
    for (const [index, word] of candidates.entries()) {
      const fontSize = minSize + (maxSize - minSize) * Math.sqrt(word.value / largest)
      const textWidth = [...word.text].reduce((sum, char) => sum + (char.codePointAt(0)! > 255 ? 1.05 : /[MW@#%]/.test(char) ? 0.95 : 0.65), 0) * fontSize
      const rectWidth = cos * textWidth + sin * fontSize * 1.2, rectHeight = sin * textWidth + cos * fontSize * 1.2
      let placed = false
      for (let step = 0; step < 2400; step++) {
        const angle = step * 0.45 + index * 0.7, radius = Math.sqrt(step / 2399) * Math.min(width, height) * 0.48
        const x = width / 2 + Math.cos(angle) * radius * width / Math.max(1, height) - rectWidth / 2, y = height / 2 + Math.sin(angle) * radius - rectHeight / 2
        const rect = { x, y, width: rectWidth, height: rectHeight }
        if (x < 0 || y < 0 || x + rectWidth > width || y + rectHeight > height || !cloudRectInside(rect, width, height, options?.shape ?? 'cloud') || collides(rect)) continue
        add({ ...word, ...rect, fontSize, rotation }); placed = true; break
      }
      if (!placed) { if (index >= words.length) break; failed = true; break }
    }
    if (!failed) return result.map(word => ({ ...word, x: plot.x + word.x / expansion, y: plot.y + word.y / expansion, width: word.width / expansion, height: word.height / expansion, fontSize: word.fontSize / expansion }))
  }
  throw new RangeError('词云数据无法在当前图表区域内完整排布，请增大图表或调整字号。')
}

function cloudRectInside(rect: RenderRect, width: number, height: number, shape: NonNullable<ChartSemanticOptions['wordCloud']>['shape']): boolean {
  return [[rect.x, rect.y], [rect.x + rect.width, rect.y], [rect.x, rect.y + rect.height], [rect.x + rect.width, rect.y + rect.height]].every(([px, py]) => {
    const x = px! / width * 2 - 1, y = py! / height * 2 - 1
    if (shape === 'rect') return true
    if (shape === 'circle') return x * x + y * y < 0.94
    if (shape === 'diamond') return Math.abs(x) + Math.abs(y) < 0.96
    if (shape === 'heart') return (x * x + y * y - 0.7) ** 3 + x * x * y ** 3 < 0
    if (shape === 'cardioid') { const angle = Math.atan2(y, x); return Math.hypot(x, y) < 0.5 * (1 - Math.sin(angle)) + 0.15 }
    if (shape === 'cloud') return x * x / 0.98 + y * y / 0.68 < 1 || (x + 0.42) ** 2 + (y + 0.12) ** 2 < 0.24 || (x - 0.38) ** 2 + (y - 0.1) ** 2 < 0.23
    const count = shape === 'pentagon' ? 5 : shape === 'star' ? 10 : 3, start = shape === 'triangle-forward' ? 0 : -Math.PI / 2
    const polygon = Array.from({ length: count }, (_unused, index) => chartPolar({ x: 0, y: 0 }, shape === 'star' && index % 2 ? 0.45 : 0.98, start + index * Math.PI * 2 / count))
    let inside = false
    for (let index = 0, previous = polygon.length - 1; index < polygon.length; previous = index++) {
      const a = polygon[index]!, b = polygon[previous]!
      if ((a.y > y) !== (b.y > y) && x < (b.x - a.x) * (y - a.y) / (b.y - a.y) + a.x) inside = !inside
    }
    return inside
  })
}
