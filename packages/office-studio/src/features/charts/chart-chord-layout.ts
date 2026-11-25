import type { ChartModel } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartChordData, ChartChordLink, ChartChordNode } from './chart-chord-data'
import { resolveChartChordLayout } from './chart-chord'
import { chartArc, chartBezier, type ChartVector } from './chart-geometry'

export interface ChartChordNodeLayout extends ChartChordNode {}
export interface ChartChordLinkLayout extends ChartChordLink { band: ChartVector[]; bounds: RenderRect }
export interface ChartChordLayout {
  center: ChartVector; innerRadius: number; outerRadius: number; plot: RenderRect; clockwise: boolean
  nodes: ChartChordNodeLayout[]; links: ChartChordLinkLayout[]
}

/** Only map authoritative normalized angles into the viewport. A layout change
 * requires a new source/config task rather than another renderer-owned engine. */
export function chartChordLayout(data: ChartChordData, chart: ChartModel, plot: RenderRect): ChartChordLayout {
  if (!data) throw new TypeError('和弦绘制需要匹配的计算结果。')
  const options = resolveChartChordLayout(chart.options?.chord)
  for (const key of ['startAngle', 'clockwise', 'padAngle', 'minAngle'] as const) {
    if (data.angleOptions[key] !== options[key]) throw new TypeError('和弦角度结果已过期，需要更新图表计算任务。')
  }
  if (![plot.x, plot.y, plot.width, plot.height].every(Number.isFinite) || plot.width < 0 || plot.height < 0) throw new RangeError('和弦视口须为有效矩形。')
  const radius = Math.min(plot.width, plot.height) / 2, center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }
  const innerRadius = radius * options.innerRadiusRatio, outerRadius = radius * options.outerRadiusRatio
  const nodes = data.nodes.map(node => ({ ...node }))
  const links = data.links.map(link => {
    const source = chartArc(center, innerRadius, link.sourceStart, link.sourceEnd), target = chartArc(center, innerRadius, link.targetStart, link.targetEnd)
    const first = chartBezier(source.at(-1)!, center, center, target[0]!), second = chartBezier(target.at(-1)!, center, center, source[0]!)
    const band = [...source, ...first.slice(1), ...target.slice(1), ...second.slice(1)]
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity
    for (const point of band) { minX = Math.min(minX, point.x); minY = Math.min(minY, point.y); maxX = Math.max(maxX, point.x); maxY = Math.max(maxY, point.y) }
    return { ...link, band, bounds: { x: minX, y: minY, width: maxX - minX, height: maxY - minY } }
  })
  return { center, innerRadius, outerRadius, plot: { ...plot }, clockwise: options.clockwise, nodes, links }
}
