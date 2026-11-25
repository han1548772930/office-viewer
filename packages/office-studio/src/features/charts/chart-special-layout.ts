import type { RenderRect } from '../../rendering/contracts'
import type { ChartHierarchyNode } from './chart-data'
import type { ChartVector } from './chart-geometry'
import type { ChartTreemapRenderNode } from './chart-treemap'
export { chartSankeyLayout } from './chart-sankey-layout'
export { chartTreemapLayout } from './chart-treemap-layout'
export { chartSunburstLayout } from './chart-sunburst-layout'
export { chartChordLayout } from './chart-chord-layout'
export { chartRelationLayout } from './chart-relation-layout'
export type { ChartRelationGraphLayout, ChartRelationNodeLayout, ChartRelationLinkLayout } from './chart-relation-layout'
export type { ChartChordLayout, ChartChordNodeLayout, ChartChordLinkLayout } from './chart-chord-layout'
export type { ChartSunburstSector } from './chart-sunburst-layout'

export interface ChartGraphNodeLayout extends ChartVector { name: string; value: number; width: number; height: number; index: number }
export interface ChartGraphLinkLayout { source: string; target: string; value: number; width: number; points: ChartVector[]; colorIndex: number }
export interface ChartGraphLayout { nodes: ChartGraphNodeLayout[]; links: ChartGraphLinkLayout[] }
export interface ChartTreeRect extends RenderRect { node: ChartHierarchyNode; renderNode: ChartTreemapRenderNode; pointIndex: number; rootIndex: number; depth: number; colorIndex: number; path: string[]; labelHeight: number }
export { chartWordCloudLayout } from './chart-word-cloud-layout'
export type { ChartCloudWord } from './chart-word-cloud-layout'

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

