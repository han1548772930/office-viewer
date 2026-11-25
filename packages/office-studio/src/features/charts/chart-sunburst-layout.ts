import type { ChartHierarchyNode } from './chart-data'
import type { RenderRect } from '../../rendering/contracts'
import { finiteChartValue } from './chart-math'
import { CHART_SUNBURST_INNER_RADIUS_RATIO, CHART_SUNBURST_OUTER_RADIUS_RATIO, chartSunburstRenderNodes, chartSunburstTotal,
  type ChartSunburstRenderNode } from './chart-sunburst'

export interface ChartSunburstSector extends Omit<ChartSunburstRenderNode, 'children'> {
  renderNode: Omit<ChartSunburstRenderNode, 'children'>
  start: number; end: number; inner: number; outer: number; colorIndex: number
}

/** Locked SunburstLayout uses one global angle unit rather than normalizing
 * each parent's children. A negative self fact can leave positive child arcs
 * wider than their parent. Keep this observable geometry and signed metadata. */
export function chartSunburstLayout(nodes: ChartHierarchyNode[], plot: RenderRect): ChartSunburstSector[] {
  const roots = chartSunburstRenderNodes(nodes), total = chartSunburstTotal(nodes), output: ChartSunburstSector[] = []
  if (total <= 0 || plot.width <= 0 || plot.height <= 0) return output
  let depthCount = 1
  const inspect = [...roots]
  while (inspect.length) {
    const node = inspect.pop()!; depthCount = Math.max(depthCount, node.depth + 1)
    for (const child of node.children ?? []) inspect.push(child)
  }
  const radius = Math.min(plot.width, plot.height) / 2, inner = radius * CHART_SUNBURST_INNER_RADIUS_RATIO
  const outer = radius * CHART_SUNBURST_OUTER_RADIUS_RATIO, step = (outer - inner) / depthCount
  const pending: Array<{ node: ChartSunburstRenderNode; start: number; end: number }> = []
  const schedule = (siblings: ChartSunburstRenderNode[], start: number): void => {
    let angle = start
    const entries: typeof pending = []
    for (const node of siblings) {
      const span = finiteChartValue(node.value / total * Math.PI * 2, '旭日图扇区角度'), end = finiteChartValue(angle + span, '旭日图扇区位置')
      entries.push({ node, start: angle, end }); angle = end
    }
    for (let index = entries.length - 1; index >= 0; index--) pending.push(entries[index]!)
  }
  schedule(roots, -Math.PI / 2)
  while (pending.length) {
    const { node, start, end } = pending.pop()!
    // Layout packets retain one node's metadata rather than serializing its
    // entire subtree again at every ancestor/descendant position.
    const { children: renderChildren, ...item } = node, semanticNode = { ...item.node }
    delete semanticNode.children
    const metadata = { ...item, node: semanticNode }
    output.push({ ...metadata, renderNode: metadata, start, end, inner: inner + step * node.depth,
      outer: inner + step * (node.depth + 1), colorIndex: node.rootIndex })
    if (renderChildren?.length) schedule(renderChildren, start)
  }
  return output
}
