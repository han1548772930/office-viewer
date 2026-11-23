import type { ChartModel } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartData } from './chart-engine'
import { chartCompoundPieData, type ChartPiePart } from './chart-derived'
import { chartPolar, type ChartVector } from './chart-geometry'

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

export function chartPieLayout(data: ChartData, chart: ChartModel, plot: RenderRect): ChartPieLayout {
  const options = chart.options?.pie, compound = chart.type === 'pieOfPie' || chart.type === 'barOfPie'
  const result: ChartPieLayout = { sectors: [], bars: [], connectors: [] }
  const radius = Math.max(1, Math.min(plot.width, plot.height) * 0.43)
  const startAngle = (options?.startAngle ?? -90) * Math.PI / 180, angleSize = options?.isHalfPie ? Math.PI : Math.PI * 2
  const primaryCenter = { x: plot.x + plot.width * (compound ? 0.3 : 0.5), y: plot.y + plot.height * 0.5 }
  const total = data.series[0]!.values.reduce<number>((sum, value) => sum + Math.max(0, value ?? 0), 0)
  const primaryRadius = compound ? Math.min(radius, plot.width * 0.25) : radius
  const inner = chart.type === 'doughnut' ? primaryRadius * (options?.doughnutHole ?? 0.55) : 0
  const partition = compound ? chartCompoundPieData(data, chart) : undefined
  const parts = partition?.primary ?? data.categories.map((category, index) => ({ category, index, value: Math.max(0, data.series[0]!.values[index] ?? 0) }))
  const sectorLayout = (items: ChartPiePart[], center: ChartVector, outer: number, hole: number, secondary: boolean): void => {
    const sum = items.reduce((value, part) => value + part.value, 0); let angle = startAngle
    let largest = 0
    for (const part of items) largest = Math.max(largest, part.value)
    for (const part of items) {
      if (part.value <= 0 || sum <= 0) continue
      const end = angle + angleSize * part.value / sum, middle = (angle + end) / 2
      const exploded = chartPolar(center, Math.max(0, options?.explosion ?? 0) / 100 * outer, middle)
      result.sectors.push({ part, center: exploded, start: angle, end, inner: hole, outer: options?.rosePie ? outer * Math.sqrt(part.value / Math.max(1, largest)) : outer, percentage: total ? part.value * 100 / total : 0, secondary })
      angle = end
    }
  }
  sectorLayout(parts, primaryCenter, primaryRadius, inner, false)
  if (partition) {
    const secondaryRatio = Math.max(0.15, Math.min(1, (options?.secondaryPlot?.size ?? 65) / 100)), secondaryRadius = primaryRadius * secondaryRatio
    const secondaryCenter = { x: plot.x + plot.width * 0.79, y: primaryCenter.y }
    if (chart.type === 'barOfPie' || options?.secondaryPlot?.type === 'bar') {
      const width = secondaryRadius * 0.9, height = secondaryRadius * 2, sum = partition.secondary.reduce((value, part) => value + part.value, 0)
      let y = secondaryCenter.y - height / 2
      for (const part of partition.secondary) {
        const slice = sum > 0 ? height * part.value / sum : 0
        result.bars.push({ part, rect: { x: secondaryCenter.x - width / 2, y, width, height: slice }, percentage: total ? part.value * 100 / total : 0 }); y += slice
      }
      result.connectors.push([{ x: primaryCenter.x + primaryRadius * 0.75, y: primaryCenter.y - primaryRadius * 0.45 }, { x: secondaryCenter.x - width / 2, y: secondaryCenter.y - height / 2 }],
        [{ x: primaryCenter.x + primaryRadius * 0.75, y: primaryCenter.y + primaryRadius * 0.45 }, { x: secondaryCenter.x - width / 2, y: secondaryCenter.y + height / 2 }])
    } else {
      sectorLayout(partition.secondary, secondaryCenter, secondaryRadius, 0, true)
      result.connectors.push([{ x: primaryCenter.x + primaryRadius * 0.75, y: primaryCenter.y - primaryRadius * 0.45 }, { x: secondaryCenter.x - secondaryRadius * 0.75, y: secondaryCenter.y - secondaryRadius * 0.45 }],
        [{ x: primaryCenter.x + primaryRadius * 0.75, y: primaryCenter.y + primaryRadius * 0.45 }, { x: secondaryCenter.x - secondaryRadius * 0.75, y: secondaryCenter.y + secondaryRadius * 0.45 }])
    }
  }
  return result
}

export function chartGaugeTiles(data: ChartData, plot: RenderRect): Array<{ categoryIndex: number; center: ChartVector; radius: number }> {
  const columns = Math.max(1, Math.ceil(Math.sqrt(data.categories.length * plot.width / Math.max(1, plot.height)))), rows = Math.max(1, Math.ceil(data.categories.length / columns))
  return data.categories.map((_category, categoryIndex) => ({ categoryIndex,
    center: { x: plot.x + (categoryIndex % columns + 0.5) * plot.width / columns, y: plot.y + (Math.floor(categoryIndex / columns) + 0.5) * plot.height / rows },
    radius: Math.max(1, Math.min(plot.width / columns, plot.height / rows) * 0.38),
  }))
}
