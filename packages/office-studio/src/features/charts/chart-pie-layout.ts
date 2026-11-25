import type { ChartModel } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartData } from './chart-engine'
import { chartCompoundPieData, type ChartPiePart } from './chart-derived'
import { chartPolar, type ChartVector } from './chart-geometry'
import { chartPieAggregateExplosion, chartPieMaximumExplosion, chartPieSliceExplosion } from './chart-pie'

export interface ChartPieSectorLayout {
  part: ChartPiePart
  center: ChartVector
  start: number
  end: number
  inner: number
  outer: number
  percentage: number
  secondary: boolean
}
export interface ChartPieLayout {
  sectors: ChartPieSectorLayout[]
  bars: Array<{ part: ChartPiePart; rect: RenderRect; percentage: number }>
  connectors: Array<[ChartVector, ChartVector]>
}

export interface ChartFunnelStep { category: string; value: number; categoryIndex: number; points: ChartVector[]; center: ChartVector }

export function chartPieOuterRadius(plot: RenderRect, explosion = 0): number {
  return Math.max(1, Math.min(plot.width, plot.height) * 0.43 / (1 + Math.max(0, explosion) / 100))
}

export function chartPieLayout(data: ChartData, chart: ChartModel, plot: RenderRect): ChartPieLayout {
  const options = chart.options?.pie, compound = chart.type === 'pieOfPie' || chart.type === 'barOfPie'
  const result: ChartPieLayout = { sectors: [], bars: [], connectors: [] }
  if (!data.series.length || !data.categories.length) return result
  const maximumExplosion = chartPieMaximumExplosion(data, chart)
  const radius = chartPieOuterRadius(plot, maximumExplosion)
  const startAngle = (options?.startAngle ?? -90) * Math.PI / 180, angleSize = options?.isHalfPie ? Math.PI : Math.PI * 2
  const scale = data.series[0]!.values.reduce<number>((max, value) => Math.max(max, value ?? 0), 0) || 1
  const total = data.series[0]!.values.reduce<number>((sum, value) => sum + Math.max(0, value ?? 0) / scale, 0)
  const percentage = (value: number): number => total > 0 ? (value / scale) / total * 100 : 0
  const secondaryBar = chart.type === 'barOfPie' || options?.secondaryPlot?.type === 'bar'
  const secondaryRatio = (options?.secondaryPlot?.size ?? 65) / 100, gapRatio = (options?.secondaryPlot?.gapWidth ?? 50) / 100
  const secondaryWidthRatio = secondaryRatio * (secondaryBar ? 0.45 : 1), explosionRatio = 1 + maximumExplosion / 100
  const primaryRadius = compound ? Math.min(radius, plot.width / (2 * explosionRatio * (1 + secondaryWidthRatio) + gapRatio)) : radius
  const secondaryRadius = primaryRadius * secondaryRatio, secondaryHalfWidth = primaryRadius * secondaryWidthRatio
  const usedWidth = 2 * explosionRatio * (primaryRadius + secondaryHalfWidth) + primaryRadius * gapRatio
  const primaryCenter = { x: compound ? plot.x + (plot.width - usedWidth) / 2 + primaryRadius * explosionRatio : plot.x + plot.width / 2, y: plot.y + plot.height / 2 }
  const secondaryCenter = { x: primaryCenter.x + explosionRatio * (primaryRadius + secondaryHalfWidth) + primaryRadius * gapRatio, y: primaryCenter.y }
  const inner = chart.type === 'doughnut' ? primaryRadius * (options?.doughnutHole ?? 0.55) : 0
  const partition = compound ? chartCompoundPieData(data, chart) : undefined
  const parts = partition?.primary ?? data.categories.map((category, index) => ({ category, index, value: Math.max(0, data.series[0]!.values[index] ?? 0) }))
  const sectorLayout = (items: ChartPiePart[], center: ChartVector, outer: number, hole: number, secondary: boolean): void => {
    const partScale = items.reduce((max, part) => Math.max(max, part.value), 0) || 1, sum = items.reduce((value, part) => value + part.value / partScale, 0); let angle = startAngle
    let largest = 0
    for (const part of items) largest = Math.max(largest, part.value)
    for (const part of items) {
      if (part.value <= 0 || sum <= 0) continue
      const end = angle + angleSize * (part.value / partScale) / sum, middle = (angle + end) / 2
      const explosion = part.index === partition?.aggregateIndex ? chartPieAggregateExplosion(chart) : chartPieSliceExplosion(chart, part.index)
      const exploded = chartPolar(center, Math.max(0, explosion) / 100 * outer, middle)
      result.sectors.push({ part, center: exploded, start: angle, end, inner: hole, outer: options?.rosePie ? outer * Math.sqrt(part.value / largest) : outer, percentage: percentage(part.value), secondary })
      angle = end
    }
  }
  sectorLayout(parts, primaryCenter, primaryRadius, inner, false)
  if (partition) {
    let ends: [ChartVector, ChartVector]
    if (secondaryBar) {
      const width = secondaryRadius * 0.9, height = secondaryRadius * 2, partScale = partition.secondary.reduce((max, part) => Math.max(max, part.value), 0) || 1, sum = partition.secondary.reduce((value, part) => value + part.value / partScale, 0)
      let y = secondaryCenter.y - height / 2
      for (const part of partition.secondary) {
        const slice = sum > 0 ? height * (part.value / partScale) / sum : 0
        result.bars.push({ part, rect: { x: secondaryCenter.x - width / 2, y, width, height: slice }, percentage: percentage(part.value) }); y += slice
      }
      ends = [{ x: secondaryCenter.x - width / 2, y: secondaryCenter.y - height / 2 }, { x: secondaryCenter.x - width / 2, y: secondaryCenter.y + height / 2 }]
    } else {
      sectorLayout(partition.secondary, secondaryCenter, secondaryRadius, 0, true)
      ends = [chartPolar(secondaryCenter, secondaryRadius, -Math.PI * 0.75), chartPolar(secondaryCenter, secondaryRadius, Math.PI * 0.75)]
    }
    const aggregate = result.sectors.find(sector => !sector.secondary && sector.part.index === partition.aggregateIndex)
    if (aggregate) {
      const starts = [chartPolar(aggregate.center, aggregate.outer, aggregate.start), chartPolar(aggregate.center, aggregate.outer, aggregate.end)].sort((a, b) => a.y - b.y)
      result.connectors.push([starts[0]!, ends[0]], [starts[1]!, ends[1]])
    }
  }
  return result
}

/** Sorting changes drawing order while preserving each original data-point index. */
export function chartFunnelLayout(data: ChartData, chart: ChartModel, plot: RenderRect): ChartFunnelStep[] {
  const source = data.series[0], options = chart.options?.funnel
  if (!source || !data.categories.length) return []
  const parts = data.categories.map((category, categoryIndex) => ({ category, categoryIndex, value: options?.useAbsoluteValue ? Math.abs(source.values[categoryIndex] ?? 0) : source.values[categoryIndex] ?? 0 }))
  if (options?.sort !== 'none') parts.sort((a, b) => options?.sort === 'ascending' ? a.value - b.value || a.categoryIndex - b.categoryIndex : b.value - a.value || a.categoryIndex - b.categoryIndex)
  const max = parts.reduce((value, part) => Math.max(value, part.value), 0) || 1, step = plot.height / parts.length, gap = Math.min(options?.gap ?? 2, step)
  return parts.map((part, index) => {
    const width = Math.max(0, part.value / max) * plot.width, nextWidth = Math.max(0, (parts[index + 1]?.value ?? 0) / max) * plot.width
    const x = plot.x + plot.width / 2, top = plot.y + index * step + gap / 2, bottom = plot.y + (index + 1) * step - gap / 2
    return { ...part, center: { x, y: (top + bottom) / 2 }, points: [{ x: x - width / 2, y: top }, { x: x + width / 2, y: top }, { x: x + nextWidth / 2, y: bottom }, { x: x - nextWidth / 2, y: bottom }] }
  })
}

export function chartGaugeTiles(data: ChartData, plot: RenderRect): Array<{ categoryIndex: number; center: ChartVector; radius: number }> {
  const indexes = data.categories.flatMap((_category, index) => Number.isFinite(data.series[0]?.values[index]) ? [index] : [])
  const columns = Math.max(1, Math.ceil(Math.sqrt(indexes.length * plot.width / Math.max(1, plot.height)))), rows = Math.max(1, Math.ceil(indexes.length / columns))
  return indexes.map((categoryIndex, tileIndex) => ({ categoryIndex,
    center: { x: plot.x + (tileIndex % columns + 0.5) * plot.width / columns, y: plot.y + (Math.floor(tileIndex / columns) + 0.5) * plot.height / rows },
    radius: Math.max(1, Math.min(plot.width / columns, plot.height / rows) * 0.38),
  }))
}
