import type { ChartLineOptions, ChartModel } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartData, ChartLayout, ChartPoint } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartPalette, chartSeriesStyle } from './chart-options'
import { chartPointStyle } from './chart-derived'
import {
  chartAxisRatio, chartAxisTicks, chartFont, chartLine, chartPolygon, chartShade, formatChartNumber,
  type ChartVector3,
} from './chart-geometry'

interface Chart3DFace { points: ChartVector3[]; color: string; border?: ChartLineOptions }
export interface Chart3DScene { commands: DrawCommand[]; points: ChartPoint[][] }

/** Applies camera rotation and perspective to authored geometry; it does not change source values. */
export function chart3DScene(data: ChartData, chart: ChartModel, layout: ChartLayout, textColor = '#64748b', palette = chartPalette(chart)): Chart3DScene {
  const descriptor = resolveChartType(chart.type), { plot } = layout, project = chart3DProjector(plot, chart)
  const commands: DrawCommand[] = [], faces: Chart3DFace[] = [], points: ChartPoint[][] = data.series.map(() => [])
  const horizontal = descriptor.family === 'bar', count = Math.max(1, data.categories.length), seriesCount = Math.max(1, data.series.length)
  const ratio = (value: number): number => chartAxisRatio(value, layout.min, layout.max, chart.yAxis)
  const categoryRatio = (index: number): number => chart.xAxis.reverse ? 1 - (index + 0.5) / count : (index + 0.5) / count
  const world = (category: number, value: number, depth: number): ChartVector3 => horizontal ? { x: value, y: 1 - category, z: depth } : { x: category, y: value, z: depth }
  const addFace = (vertices: ChartVector3[], color: string, border?: ChartLineOptions): void => { faces.push({ points: vertices.map(project), color, border }) }
  const baseline = chart.yAxis.scaleType === 'log' ? layout.min : 0
  if (descriptor.family === 'pie') {
    const source = data.series[0]!, total = source.values.reduce<number>((sum, value) => sum + Math.max(0, value ?? 0), 0)
    let start = (chart.options?.pie?.startAngle ?? -90) * Math.PI / 180
    const height = Math.max(0.04, Math.min(0.38, (chart.options?.threeD?.view?.depthPercent ?? 100) / 500))
    source.values.forEach((value, categoryIndex) => {
      if (value === null || value <= 0 || total <= 0) return
      const end = start + value / total * Math.PI * 2, middle = (start + end) / 2
      const style = chartPointStyle(chartSeriesStyle(chart, source.field, 0, palette), categoryIndex), color = style.color ?? palette[categoryIndex % palette.length]!
      const explosion = Math.min(0.18, (chart.options?.pie?.explosion ?? 0) / 100 * 0.5)
      const cx = 0.5 + Math.cos(middle) * explosion, cz = 0.5 + Math.sin(middle) * explosion
      const radial = (angle: number, level: number): ChartVector3 => ({ x: cx + Math.cos(angle) * 0.43, y: level, z: cz + Math.sin(angle) * 0.43 })
      const steps = Math.max(2, Math.ceil((end - start) / (Math.PI / 48)))
      for (let index = 0; index < steps; index++) {
        const a = start + (end - start) * index / steps, b = start + (end - start) * (index + 1) / steps
        addFace([radial(a, 0), radial(b, 0), radial(b, height), radial(a, height)], chartShade(color, 0.67), style.border)
        addFace([{ x: cx, y: height, z: cz }, radial(a, height), radial(b, height)], color)
      }
      addFace([{ x: cx, y: 0, z: cz }, radial(start, 0), radial(start, height), { x: cx, y: height, z: cz }], chartShade(color, 0.8), style.border)
      addFace([{ x: cx, y: 0, z: cz }, radial(end, 0), radial(end, height), { x: cx, y: height, z: cz }], chartShade(color, 0.74), style.border)
      const anchor = project({ x: cx + Math.cos(middle) * 0.28, y: height, z: cz + Math.sin(middle) * 0.28 })
      points[0]!.push({ x: anchor.x, y: anchor.y, value, category: data.categories[categoryIndex] ?? '', categoryIndex })
      start = end
    })
  } else {
    const floor = [world(0, ratio(baseline), 0), world(1, ratio(baseline), 0), world(1, ratio(baseline), 1), world(0, ratio(baseline), 1)].map(project)
    chartPolygon(commands, floor, chart.options?.appearance?.background ?? 'rgba(148, 163, 184, 0.06)', { color: chart.yAxis.line?.color ?? textColor, width: 0.7 })
    for (const value of chartAxisTicks(layout.min, layout.max, chart.yAxis)) {
      const a = project(world(0, ratio(value), 0)), b = project(world(1, ratio(value), 0)), c = project(world(0, ratio(value), 1))
      chartLine(commands, a, b, chart.yAxis.gridLine, '#cbd5e1', plot)
      chartLine(commands, a, c, chart.yAxis.gridLine, '#cbd5e1', plot)
      if (chart.yAxis.label?.visible !== false) commands.push({ kind: 'text', x: a.x - 5, y: a.y, text: `${formatChartNumber(value, chart.yAxis.format)}${descriptor.percent && !chart.yAxis.format ? '%' : ''}`, font: chartFont(chart.yAxis.label, 10), color: chart.yAxis.label?.color ?? textColor, align: 'right', baseline: 'middle' })
    }
    data.categories.forEach((category, index) => {
      const point = project(world(categoryRatio(index), ratio(baseline), 0))
      if (chart.xAxis.label?.visible !== false) commands.push({ kind: 'text', x: point.x, y: point.y + 12, text: category, font: chartFont(chart.xAxis.label, 10), color: chart.xAxis.label?.color ?? textColor, align: 'center', baseline: 'middle', maxWidth: Math.max(12, plot.width / count) })
    })
    const positive = Array(count).fill(0) as number[], negative = Array(count).fill(0) as number[]
    const gapRatio = 1 / (1 + (chart.options?.bar?.gapWidth ?? 150) / 100), overlap = (chart.options?.bar?.overlap ?? 0) / 100
    for (const [seriesIndex, source] of data.series.entries()) {
      const style = chartSeriesStyle(chart, source.field, seriesIndex, palette)
      const grouped = chart.options?.threeD?.grouping === 'clustered' && !descriptor.stacked
      const depthGap = Math.max(0, (chart.options?.threeD?.gapDepth ?? 150) / 100)
      const depthWidth = Math.max(0.02, 0.75 / (descriptor.stacked || grouped ? 1 : seriesCount) / (1 + depthGap * 0.3))
      const depth = descriptor.stacked || grouped ? 0.35 : (seriesIndex + 0.5) / seriesCount
      const band = gapRatio / count, width = grouped ? band / Math.max(1, seriesCount - (seriesCount - 1) * overlap) : band
      let previous: { categoryIndex: number; end: number; base: number; category: number; value: number } | undefined
      source.values.forEach((value, categoryIndex) => {
        if (value === null || !Number.isFinite(value) || chart.yAxis.scaleType === 'log' && value <= 0) { if (chart.options?.invalidValueStrategy !== 'link') previous = undefined; return }
        const base = descriptor.stacked ? value >= 0 ? positive[categoryIndex]! : negative[categoryIndex]! : baseline
        if (descriptor.stacked) { if (value >= 0) positive[categoryIndex] = base + value; else negative[categoryIndex] = base + value }
        const end = descriptor.stacked ? base + value : value
        const center = categoryRatio(categoryIndex) + (grouped ? (seriesIndex - (seriesCount - 1) / 2) * width * (1 - overlap) : 0)
        const point = project(world(center, ratio(end), depth)), basePoint = project(world(center, ratio(base), depth))
        points[seriesIndex]!.push({ x: point.x, y: point.y, value, category: data.categories[categoryIndex] ?? '', categoryIndex, baseY: basePoint.y, baseX: basePoint.x, baseValue: base })
        const pointStyle = chartPointStyle(style, categoryIndex), color = pointStyle.color ?? palette[seriesIndex % palette.length]!
        if (descriptor.family === 'column' || descriptor.family === 'bar') {
          const a = world(center - width / 2, ratio(base), depth - depthWidth / 2), b = world(center + width / 2, ratio(base), depth - depthWidth / 2)
          const c = world(center + width / 2, ratio(end), depth - depthWidth / 2), d = world(center - width / 2, ratio(end), depth - depthWidth / 2)
          const e = { ...a, z: depth + depthWidth / 2 }, f = { ...b, z: depth + depthWidth / 2 }, g = { ...c, z: depth + depthWidth / 2 }, h = { ...d, z: depth + depthWidth / 2 }
          addFace([a, b, c, d], color, pointStyle.border); addFace([e, f, g, h], chartShade(color, 0.76), pointStyle.border)
          addFace([a, e, h, d], chartShade(color, 0.68), pointStyle.border); addFace([b, f, g, c], chartShade(color, 0.88), pointStyle.border)
          addFace([d, c, g, h], chartShade(color, 1.18), pointStyle.border); addFace([a, b, f, e], chartShade(color, 0.6), pointStyle.border)
        } else if (previous && (chart.options?.invalidValueStrategy === 'link' || categoryIndex === previous.categoryIndex + 1)) {
          const a = world(previous.category, ratio(previous.end), depth - depthWidth / 2), b = world(center, ratio(end), depth - depthWidth / 2)
          const c = { ...b, z: depth + depthWidth / 2 }, d = { ...a, z: depth + depthWidth / 2 }
          addFace([a, b, c, d], color, pointStyle.border)
          if (descriptor.family === 'area') {
            const aBase = world(previous.category, ratio(previous.base), depth - depthWidth / 2), bBase = world(center, ratio(base), depth - depthWidth / 2)
            if (previous.value * value < 0) {
              const t = Math.abs(previous.value) / (Math.abs(previous.value) + Math.abs(value)), crossCategory = previous.category + (center - previous.category) * t
              const crossBase = previous.base + (base - previous.base) * t, crossing = world(crossCategory, ratio(crossBase), depth - depthWidth / 2)
              addFace([aBase, a, crossing], chartShade(color, 0.84), pointStyle.border); addFace([crossing, b, bBase], chartShade(color, 0.84), pointStyle.border)
              addFace([{ ...aBase, z: depth + depthWidth / 2 }, d, { ...crossing, z: depth + depthWidth / 2 }], chartShade(color, 0.68), pointStyle.border)
              addFace([{ ...crossing, z: depth + depthWidth / 2 }, c, { ...bBase, z: depth + depthWidth / 2 }], chartShade(color, 0.68), pointStyle.border)
            } else {
              addFace([aBase, a, b, bBase], chartShade(color, 0.84), pointStyle.border)
              addFace([{ ...aBase, z: depth + depthWidth / 2 }, d, c, { ...bBase, z: depth + depthWidth / 2 }], chartShade(color, 0.68), pointStyle.border)
            }
          }
        }
        previous = { categoryIndex, end, base, category: center, value }
      })
    }
  }
  faces.sort((a, b) => a.points.reduce((sum, point) => sum + point.z, 0) / a.points.length - b.points.reduce((sum, point) => sum + point.z, 0) / b.points.length)
  for (const face of faces) chartPolygon(commands, face.points, face.color, face.border, plot)
  return { commands, points }
}

function chart3DProjector(plot: RenderRect, chart: ChartModel): (point: ChartVector3) => ChartVector3 {
  const view = chart.options?.threeD?.view, pitch = (view?.rotX ?? 28) * Math.PI / 180, yaw = (view?.rotY ?? 28) * Math.PI / 180
  const height = (view?.heightPercent ?? 100) / 100, depth = (view?.depthPercent ?? 100) / 100 * 0.65, perspective = view?.rightAngleAxes ? 0 : view?.perspective ?? 30
  const distance = 2 + 90 / Math.max(1, perspective)
  const rotate = (point: ChartVector3): ChartVector3 => {
    const x = point.x - 0.5, y = (point.y - 0.5) * height, z = (point.z - 0.5) * depth
    const rx = x * Math.cos(yaw) + z * Math.sin(yaw), rz = -x * Math.sin(yaw) + z * Math.cos(yaw)
    const ry = y * Math.cos(pitch) - rz * Math.sin(pitch), cameraZ = y * Math.sin(pitch) + rz * Math.cos(pitch)
    const scale = perspective ? distance / Math.max(0.1, distance - cameraZ) : 1
    return { x: rx * scale, y: ry * scale, z: cameraZ }
  }
  const corners = [0, 1].flatMap(x => [0, 1].flatMap(y => [0, 1].map(z => rotate({ x, y, z }))))
  const minX = Math.min(...corners.map(point => point.x)), maxX = Math.max(...corners.map(point => point.x)), minY = Math.min(...corners.map(point => point.y)), maxY = Math.max(...corners.map(point => point.y))
  return point => {
    const rotated = rotate(point)
    return { x: plot.x + (rotated.x - minX) / Math.max(1e-9, maxX - minX) * plot.width, y: plot.y + plot.height - (rotated.y - minY) / Math.max(1e-9, maxY - minY) * plot.height, z: rotated.z }
  }
}
