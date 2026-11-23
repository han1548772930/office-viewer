import type { ChartModel, ChartSemanticOptions } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartGraphData, ChartHierarchyNode } from './chart-data'
import type { ChartData } from './chart-engine'
import { chartBezier, chartPolar, type ChartVector } from './chart-geometry'

export interface ChartGraphNodeLayout extends ChartVector { name: string; value: number; width: number; height: number; index: number }
export interface ChartGraphLinkLayout { source: string; target: string; value: number; width: number; points: ChartVector[]; colorIndex: number }
export interface ChartGraphLayout { nodes: ChartGraphNodeLayout[]; links: ChartGraphLinkLayout[] }
export interface ChartChordNodeLayout { name: string; value: number; start: number; end: number; index: number }
export interface ChartChordLinkLayout { sourceStart: number; sourceEnd: number; targetStart: number; targetEnd: number; value: number; sourceIndex: number; targetIndex: number }
export interface ChartChordLayout { center: ChartVector; innerRadius: number; outerRadius: number; nodes: ChartChordNodeLayout[]; links: ChartChordLinkLayout[] }
export interface ChartTreeRect extends RenderRect { node: ChartHierarchyNode; depth: number; colorIndex: number; path: string[] }
export interface ChartSunburstSector { node: ChartHierarchyNode; start: number; end: number; inner: number; outer: number; depth: number; colorIndex: number }
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

/** Cycles remain directed links; SCCs occupy a column and feedback links use an outside port. */
export function chartSankeyLayout(graph: ChartGraphData, plot: RenderRect): ChartGraphLayout {
  const components = stronglyConnectedComponents(graph)
  const componentByName = new Map<string, number>()
  components.forEach((names, index) => names.forEach(name => componentByName.set(name, index)))
  const outgoing = components.map(() => new Set<number>()), degrees = components.map(() => 0), layers = components.map(() => 0)
  for (const link of graph.links) {
    const source = componentByName.get(link.source)!, target = componentByName.get(link.target)!
    if (source === target || outgoing[source]!.has(target)) continue
    outgoing[source]!.add(target); degrees[target]!++
  }
  const queue = degrees.flatMap((degree, index) => degree ? [] : [index])
  for (let index = 0; index < queue.length; index++) for (const next of outgoing[queue[index]!]!) {
    layers[next] = Math.max(layers[next]!, layers[queue[index]!]! + 1)
    degrees[next]!--; if (!degrees[next]) queue.push(next)
  }
  const nodeWeights = new Map(graph.nodes.map(node => [node.name, { in: 0, out: 0 }]))
  for (const link of graph.links) { nodeWeights.get(link.source)!.out += link.value; nodeWeights.get(link.target)!.in += link.value }
  const columns = new Map<number, string[]>()
  for (const node of graph.nodes) { const layer = layers[componentByName.get(node.name)!]!, list = columns.get(layer) ?? []; list.push(node.name); columns.set(layer, list) }
  const maxLayer = Math.max(0, ...layers), nodeWidth = Math.max(4, Math.min(18, plot.width / Math.max(4, (maxLayer + 1) * 3)))
  const gap = Math.max(2, Math.min(12, plot.height / Math.max(4, graph.nodes.length * 2)))
  const reserved = graph.cyclic ? Math.min(36, plot.height * 0.14) : 4
  let scale = Infinity
  for (const names of columns.values()) {
    const sum = names.reduce((total, name) => total + Math.max(nodeWeights.get(name)!.in, nodeWeights.get(name)!.out), 0)
    scale = Math.min(scale, Math.max(1, plot.height - reserved * 2 - (names.length - 1) * gap) / Math.max(1, sum))
  }
  if (!Number.isFinite(scale)) scale = 1
  const byName = new Map<string, ChartGraphNodeLayout>()
  for (const [layer, names] of columns) {
    const heights = names.map(name => Math.max(nodeWeights.get(name)!.in, nodeWeights.get(name)!.out) * scale)
    let y = plot.y + (plot.height - heights.reduce((sum, height) => sum + height, 0) - (names.length - 1) * gap) / 2
    names.forEach((name, index) => {
      const source = graph.nodes.find(node => node.name === name)!
      const x = plot.x + (maxLayer ? layer / maxLayer * (plot.width - nodeWidth - reserved) : (plot.width - nodeWidth - reserved) / 2)
      byName.set(name, { ...source, x, y, width: nodeWidth, height: heights[index]!, index: graph.nodes.indexOf(source) }); y += heights[index]! + gap
    })
  }
  const usedIn = new Map<string, number>(), usedOut = new Map<string, number>()
  const links = graph.links.map((link, index) => {
    const source = byName.get(link.source)!, target = byName.get(link.target)!, width = link.value * scale
    const out = usedOut.get(source.name) ?? 0, input = usedIn.get(target.name) ?? 0
    usedOut.set(source.name, out + width); usedIn.set(target.name, input + width)
    const start = { x: source.x + source.width, y: source.y + (source.height - nodeWeights.get(source.name)!.out * scale) / 2 + out + width / 2 }
    const forward = target.x > source.x
    const end = { x: forward ? target.x : target.x + target.width, y: target.y + (target.height - nodeWeights.get(target.name)!.in * scale) / 2 + input + width / 2 }
    const middle = forward ? (start.x + end.x) / 2 : Math.min(plot.x + plot.width - width / 2, Math.max(start.x, end.x) + reserved * (0.5 + index % 3 / 6))
    const self = source === target
    const points = self ? chartBezier(start, { x: middle, y: Math.max(plot.y + width / 2, start.y - Math.max(16, source.height)) }, { x: middle, y: Math.max(plot.y + width / 2, end.y - Math.max(16, source.height)) }, end)
      : chartBezier(start, { x: middle, y: start.y }, { x: middle, y: end.y }, end)
    return { ...link, points, width, colorIndex: source.index }
  })
  return { nodes: graph.nodes.map(node => byName.get(node.name)!), links }
}

function stronglyConnectedComponents(graph: ChartGraphData): string[][] {
  const forward = new Map(graph.nodes.map(node => [node.name, [] as string[]])), reverse = new Map(graph.nodes.map(node => [node.name, [] as string[]]))
  for (const link of graph.links) { forward.get(link.source)!.push(link.target); reverse.get(link.target)!.push(link.source) }
  const visited = new Set<string>(), order: string[] = []
  for (const node of graph.nodes) {
    if (visited.has(node.name)) continue
    const stack: Array<{ name: string; next: number }> = [{ name: node.name, next: 0 }]; visited.add(node.name)
    while (stack.length) {
      const current = stack.at(-1)!, edges = forward.get(current.name)!
      if (current.next < edges.length) { const name = edges[current.next++]!; if (!visited.has(name)) { visited.add(name); stack.push({ name, next: 0 }) } }
      else { order.push(current.name); stack.pop() }
    }
  }
  visited.clear()
  const components: string[][] = []
  for (const name of order.reverse()) {
    if (visited.has(name)) continue
    const component: string[] = [], stack = [name]; visited.add(name)
    while (stack.length) { const current = stack.pop()!; component.push(current); for (const next of reverse.get(current)!) if (!visited.has(next)) { visited.add(next); stack.push(next) } }
    components.push(component)
  }
  return components
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

export function chartChordLayout(graph: ChartGraphData, chart: ChartModel, plot: RenderRect): ChartChordLayout {
  const options = chart.options?.chord, clockwise = options?.clockwise !== false ? 1 : -1
  const radius = Math.max(1, Math.min(plot.width, plot.height) / 2), innerRadius = radius * (options?.innerRadiusRatio ?? 0.66), outerRadius = radius * (options?.outerRadiusRatio ?? 0.73)
  const padding = Math.min(Math.PI / Math.max(1, graph.nodes.length), (options?.padAngle ?? 3) * Math.PI / 180)
  const total = graph.nodes.reduce((sum, node) => sum + node.value, 0)
  const anglePerWeight = Math.max(0, Math.PI * 2 - graph.nodes.length * padding) / Math.max(1e-9, total)
  let angle = (options?.startAngle ?? -90) * Math.PI / 180
  const nodes = graph.nodes.map((node, index) => { const start = angle, end = start + clockwise * node.value * anglePerWeight; angle = end + clockwise * padding; return { ...node, start, end, index } })
  const byName = new Map(nodes.map(node => [node.name, node])), offsets = new Map(nodes.map(node => [node.name, node.start]))
  const links = graph.links.map(link => {
    const source = byName.get(link.source)!, target = byName.get(link.target)!
    const sourceStart = offsets.get(source.name)!, sourceEnd = sourceStart + clockwise * link.value * anglePerWeight; offsets.set(source.name, sourceEnd)
    const targetStart = offsets.get(target.name)!, targetEnd = targetStart + clockwise * link.value * anglePerWeight; offsets.set(target.name, targetEnd)
    return { sourceStart, sourceEnd, targetStart, targetEnd, value: link.value, sourceIndex: source.index, targetIndex: target.index }
  })
  return { center: { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }, innerRadius, outerRadius, nodes, links }
}

/** Squarified rows preserve each sibling's weight and its nested parent rectangle. */
export function chartTreemapLayout(nodes: ChartHierarchyNode[], plot: RenderRect, banner = true): ChartTreeRect[] {
  const output: ChartTreeRect[] = []
  function layout(siblings: ChartHierarchyNode[], area: RenderRect, depth: number, colorIndex: number, path: string[]): void {
    const total = siblings.reduce((sum, node) => sum + node.value, 0)
    if (total <= 0 || area.width <= 0 || area.height <= 0) return
    const sorted = siblings.map((node, index) => ({ node, index, size: node.value / total * area.width * area.height })).sort((a, b) => b.size - a.size || a.index - b.index)
    let remaining = { ...area }, cursor = 0
    while (cursor < sorted.length) {
      const side = Math.min(remaining.width, remaining.height), row: typeof sorted = []
      const aspect = (entries: typeof sorted): number => {
        const sum = entries.reduce((value, entry) => value + entry.size, 0), smallest = entries.at(-1)!.size, largest = entries[0]!.size
        return Math.max(side * side * largest / (sum * sum), sum * sum / (side * side * smallest))
      }
      row.push(sorted[cursor++]!)
      while (cursor < sorted.length && aspect([...row, sorted[cursor]!]) <= aspect(row)) row.push(sorted[cursor++]!)
      const rowSize = row.reduce((sum, entry) => sum + entry.size, 0), vertical = remaining.width >= remaining.height, thickness = rowSize / Math.max(1e-9, side)
      let offset = 0
      for (const entry of row) {
        const length = entry.size / Math.max(1e-9, thickness), index = depth === 0 ? entry.index : colorIndex
        const rect = vertical ? { x: remaining.x, y: remaining.y + offset, width: thickness, height: length } : { x: remaining.x + offset, y: remaining.y, width: length, height: thickness }
        const nextPath = [...path, entry.node.name]
        output.push({ ...rect, node: entry.node, depth, colorIndex: index, path: nextPath })
        if (entry.node.children?.length) {
          const header = banner ? Math.min(18, rect.height * 0.2) : 0
          layout(entry.node.children, { x: rect.x + 2, y: rect.y + header + 2, width: Math.max(0, rect.width - 4), height: Math.max(0, rect.height - header - 4) }, depth + 1, index, nextPath)
        }
        offset += length
      }
      remaining = vertical ? { ...remaining, x: remaining.x + thickness, width: Math.max(0, remaining.width - thickness) } : { ...remaining, y: remaining.y + thickness, height: Math.max(0, remaining.height - thickness) }
    }
  }
  layout(nodes, plot, 0, 0, [])
  return output
}

export function chartSunburstLayout(nodes: ChartHierarchyNode[], plot: RenderRect): ChartSunburstSector[] {
  const output: ChartSunburstSector[] = []
  const depthOf = (siblings: ChartHierarchyNode[]): number => siblings.reduce((max, node) => Math.max(max, 1 + (node.children?.length ? depthOf(node.children) : 0)), 0)
  const depthCount = Math.max(1, depthOf(nodes)), radius = Math.max(1, Math.min(plot.width, plot.height) * 0.43), step = radius / (depthCount + 1)
  function layout(siblings: ChartHierarchyNode[], start: number, end: number, depth: number, colorIndex: number): void {
    const total = siblings.reduce((sum, node) => sum + node.value, 0); let angle = start
    siblings.forEach((node, index) => {
      if (node.value <= 0 || total <= 0) return
      const next = angle + (end - start) * node.value / total, color = depth === 0 ? index : colorIndex
      output.push({ node, start: angle, end: next, inner: step * (depth + 1), outer: step * (depth + 2), depth, colorIndex: color })
      if (node.children?.length) layout(node.children, angle, next, depth + 1, color)
      angle = next
    })
  }
  layout(nodes, -Math.PI / 2, Math.PI * 1.5, 0, 0)
  return output
}

export function chartWordCloudLayout(data: ChartData, options: ChartSemanticOptions['wordCloud'], plot: RenderRect): ChartCloudWord[] {
  const words = data.categories.map((text, categoryIndex) => ({ text, categoryIndex, value: data.series[0]?.values[categoryIndex] ?? 0 })).filter(word => word.text && word.value > 0)
    .sort((a, b) => b.value - a.value || a.categoryIndex - b.categoryIndex)
  if (!words.length) return []
  const largest = words[0]!.value, minSize = options?.minFontSize ?? 10, maxSize = options?.maxFontSize ?? Math.min(54, plot.height / 3), rotation = options?.rotation ?? 0
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
