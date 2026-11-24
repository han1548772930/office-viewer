import type { ChartTreemapOptions } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartHierarchyNode } from './chart-data'
import type { ChartTreeRect } from './chart-special-layout'
import { chartTreemapHeaderHeight, chartTreemapRenderNodes, type ChartTreemapRenderNode } from './chart-treemap'

export interface ChartTreemapLayoutOptions { parentLabelLayout?: ChartTreemapOptions['parentLabelLayout']; fontSize?: number }

/** Iterative squarified layout consumes rendered weights, including parent-owned
 * residuals. The scene retains semantic nodes and source style identities. */
export function chartTreemapLayout(nodes: ChartHierarchyNode[], plot: RenderRect,
  input: boolean | ChartTreemapLayoutOptions = true): ChartTreeRect[] {
  const options = typeof input === 'boolean' ? { parentLabelLayout: input ? 'banner' as const : 'none' as const } : input
  const mode = options.parentLabelLayout ?? 'banner', fontSize = options.fontSize ?? 12
  const output: ChartTreeRect[] = [], pending: Array<{ siblings: ChartTreemapRenderNode[]; area: RenderRect }> = [{ siblings: chartTreemapRenderNodes(nodes), area: plot }]
  while (pending.length) {
    const { siblings, area } = pending.pop()!
    const scale = siblings.reduce((maximum, node) => Math.max(maximum, node.value), 0) || 1
    const total = siblings.reduce((sum, node) => sum + node.value / scale, 0)
    if (total <= 0 || area.width <= 0 || area.height <= 0) continue
    const sorted = siblings.map((node, index) => ({ node, index, size: node.value / scale / total * area.width * area.height }))
      .filter(entry => entry.size > 0).sort((a, b) => b.size - a.size || a.index - b.index)
    let remaining = { ...area }, cursor = 0
    while (cursor < sorted.length && remaining.width > 0 && remaining.height > 0) {
      const side = Math.min(remaining.width, remaining.height), row: typeof sorted = []
      const aspect = (sum: number, largest: number, smallest: number): number => Math.max(side * side * largest / (sum * sum), sum * sum / (side * side * smallest))
      const first = sorted[cursor++]!
      row.push(first)
      let rowSize = first.size, smallest = first.size
      while (cursor < sorted.length) {
        const next = sorted[cursor]!, nextSize = rowSize + next.size
        if (aspect(nextSize, first.size, next.size) > aspect(rowSize, first.size, smallest)) break
        row.push(next); cursor++; rowSize = nextSize; smallest = next.size
      }
      const vertical = remaining.width >= remaining.height
      const thickness = Math.min(vertical ? remaining.width : remaining.height, rowSize / side)
      let offset = 0
      for (let entryIndex = 0; entryIndex < row.length; entryIndex++) {
        const entry = row[entryIndex]!, available = side - offset
        const length = entryIndex === row.length - 1 ? available : Math.min(available, entry.size / Math.max(Number.MIN_VALUE, thickness))
        const rect = vertical ? { x: remaining.x, y: remaining.y + offset, width: thickness, height: length }
          : { x: remaining.x + offset, y: remaining.y, width: length, height: thickness }
        const parent = Boolean(entry.node.children?.length), labelHeight = parent ? chartTreemapHeaderHeight(mode, fontSize) : 0
        output.push({ ...entry.node, ...rect, renderNode: entry.node, colorIndex: entry.node.rootIndex, labelHeight })
        if (parent) {
          const padding = entry.node.depth === 0 ? 2 : 1, reserved = mode === 'banner' ? labelHeight : 0
          pending.push({ siblings: entry.node.children!, area: { x: rect.x + padding, y: rect.y + reserved + padding,
            width: Math.max(0, rect.width - padding * 2), height: Math.max(0, rect.height - reserved - padding * 2) } })
        }
        offset += length
      }
      remaining = vertical ? { ...remaining, x: remaining.x + thickness, width: Math.max(0, remaining.width - thickness) }
        : { ...remaining, y: remaining.y + thickness, height: Math.max(0, remaining.height - thickness) }
    }
  }
  return output
}
