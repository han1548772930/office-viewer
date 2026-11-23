import type { ChartAxisOptions, ChartLineOptions, ChartModel, ChartSeriesFillPattern } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartData, ChartLayout, ChartPoint } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartPalette, chartSeriesStyle } from './chart-options'
import { chartPointStyle } from './chart-derived'
import {
  chartAxisExtent, chartAxisRatio, chartAxisTicks, chartColorOpacity, chartFont, chartLine, chartPolygon, chartShade, chartSymbol, formatChartNumber,
  type ChartVector3,
} from './chart-geometry'
import { chartInterpolate, finiteChartValue } from './chart-math'
import { chartAxisTitleCommand, chartAxisTitleFont } from './chart-axis-title'
import { chartDateAxisLabel } from './chart-date-axis'
import { chartFillPatternPaint } from './chart-fill-pattern'

interface Chart3DFace { points: ChartVector3[]; color: string; border?: ChartLineOptions; pattern?: ChartSeriesFillPattern; marker?: { size: number; shape: string } }
export interface Chart3DScene { commands: DrawCommand[]; points: ChartPoint[][] }

/** Applies camera rotation and perspective to authored geometry; it does not change source values. */
export function chart3DScene(data: ChartData, chart: ChartModel, layout: ChartLayout, textColor = '#64748b', palette = chartPalette(chart)): Chart3DScene {
  const descriptor = resolveChartType(chart.type), { plot } = layout, project = chart3DProjector(plot, chart)
  const commands: DrawCommand[] = [], faces: Chart3DFace[] = [], points: ChartPoint[][] = data.series.map(() => [])
  if (!data.categories.length || !data.series.length) return { commands, points }
  const horizontal = descriptor.family === 'bar', count = Math.max(1, data.categories.length), seriesCount = Math.max(1, data.series.length)
  const ratio = (value: number): number => chartAxisRatio(value, layout.min, layout.max, chart.yAxis)
  const categoryRatio = (index: number): number => layout.categoryPositions ? layout.categoryPositions[index] ?? NaN : chart.xAxis.reverse ? 1 - (index + 0.5) / count : (index + 0.5) / count
  const depthRatio = (value: number): number => chart.options?.depthAxis?.reverse ? 1 - value : value
  const world = (category: number, value: number, depth: number): ChartVector3 => horizontal ? { x: value, y: 1 - category, z: depth } : { x: category, y: value, z: depth }
  const addFace = (vertices: ChartVector3[], color: string, border?: ChartLineOptions, pattern?: ChartSeriesFillPattern): void => { faces.push({ points: vertices.map(project), color, border, pattern }) }
  const baseline = chart.yAxis.scaleType === 'log' ? layout.min : Math.max(layout.min, Math.min(layout.max, 0))
  if (descriptor.family === 'pie') {
    const source = data.series[0]!, scale = source.values.reduce<number>((max, value) => Math.max(max, value ?? 0), 0) || 1
    const total = source.values.reduce<number>((sum, value) => sum + Math.max(0, (value ?? 0) / scale), 0)
    let start = (chart.options?.pie?.startAngle ?? -90) * Math.PI / 180
    const height = Math.max(0.04, Math.min(0.38, (chart.options?.threeD?.view?.depthPercent ?? 100) / 500))
    source.values.forEach((value, categoryIndex) => {
      if (value === null || value <= 0 || total <= 0) return
      const end = start + (value / scale) / total * Math.PI * 2, middle = (start + end) / 2
      const style = chartPointStyle(chartSeriesStyle(chart, source.field, categoryIndex, palette, source.sourceName ?? source.name), categoryIndex)
      const color = chartColorOpacity(style.color ?? palette[categoryIndex % palette.length]!, style.fillOpacity ?? 1)
      const pattern = (shade = 1) => chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1, shade)
      const explosion = Math.min(0.18, (chart.options?.pie?.explosion ?? 0) / 100 * 0.5)
      const cx = 0.5 + Math.cos(middle) * explosion, cz = 0.5 + Math.sin(middle) * explosion
      const radial = (angle: number, level: number): ChartVector3 => ({ x: cx + Math.cos(angle) * 0.43, y: level, z: cz + Math.sin(angle) * 0.43 })
      const steps = Math.max(2, Math.ceil((end - start) / (Math.PI / 48)))
      for (let index = 0; index < steps; index++) {
        const a = start + (end - start) * index / steps, b = start + (end - start) * (index + 1) / steps
        addFace([radial(a, 0), radial(b, 0), radial(b, height), radial(a, height)], chartShade(color, 0.67), style.border, pattern(0.67))
        addFace([{ x: cx, y: height, z: cz }, radial(a, height), radial(b, height)], color, undefined, pattern())
      }
      addFace([{ x: cx, y: 0, z: cz }, radial(start, 0), radial(start, height), { x: cx, y: height, z: cz }], chartShade(color, 0.8), style.border, pattern(0.8))
      addFace([{ x: cx, y: 0, z: cz }, radial(end, 0), radial(end, height), { x: cx, y: height, z: cz }], chartShade(color, 0.74), style.border, pattern(0.74))
      const anchor = project({ x: cx + Math.cos(middle) * 0.28, y: height, z: cz + Math.sin(middle) * 0.28 })
      points[0]!.push({ x: anchor.x, y: anchor.y, value, category: data.categories[categoryIndex] ?? '', categoryIndex })
      start = end
    })
  } else {
    const front = depthRatio(0), back = depthRatio(1), floorRatio = ratio(baseline)
    const floor = [world(0, floorRatio, front), world(1, floorRatio, front), world(1, floorRatio, back), world(0, floorRatio, back)].map(project)
    chartPolygon(commands, floor, chart.options?.appearance?.background ?? 'rgba(148, 163, 184, 0.06)', undefined, plot)
    const rightAxis: ChartAxisOptions = chart.options?.rightYAxis ?? { scaleType: 'linear' }
    const rightExtent = layout.rightMin === undefined || layout.rightMax === undefined ? chartAxisExtent([], rightAxis) : { min: layout.rightMin, max: layout.rightMax }
    const axisScene = (axis: ChartAxisOptions, min: number, max: number, right = false): void => {
      const axisRatio = (value: number): number => chartAxisRatio(value, min, max, axis)
      const category = horizontal ? right ? 0 : 1 : right ? 1 : 0
      chartLine(commands, project(world(category, axisRatio(min), front)), project(world(category, axisRatio(max), front)), axis.line, textColor, plot)
      for (const value of chartAxisTicks(min, max, axis)) {
        const a = project(world(category, axisRatio(value), front)), b = project(world(1 - category, axisRatio(value), front)), c = project(world(category, axisRatio(value), back))
        chartLine(commands, a, b, axis.gridLine, '#cbd5e1', plot)
        chartLine(commands, a, c, axis.gridLine, '#cbd5e1', plot)
        if (axis.tick?.visible) {
          const length = (axis.tick.length ?? 4) * (axis.tick.position === 'inside' ? -1 : 1)
          chartLine(commands, a, { x: a.x + (horizontal ? 0 : length * (right ? 1 : -1)), y: a.y + (horizontal ? length * (right ? -1 : 1) : 0) }, axis.tick, textColor)
        }
        if (axis.label?.visible !== false) commands.push({ kind: 'text', x: a.x + (horizontal ? 0 : right ? 6 : -6), y: a.y + (horizontal ? right ? -6 : 6 : 0),
          text: `${formatChartNumber(value, axis.format)}${descriptor.percent && !axis.format ? '%' : ''}`, font: chartFont({ ...chart.options?.appearance?.font, ...axis.label }, 10), color: axis.label?.color ?? textColor,
          align: horizontal ? 'center' : right ? 'left' : 'right', baseline: horizontal ? right ? 'bottom' : 'top' : 'middle', rotation: axis.label?.rotate, clip: layout.frame })
      }
      const offset = Math.max(28, (chartAxisTitleFont(chart, axis).fontSize ?? 11) + 17)
      const titlePoint = (value: number): { x: number; y: number } => {
        const anchor = project(world(category, value, front))
        return { x: anchor.x + (horizontal ? 0 : right ? offset : -offset), y: anchor.y + (horizontal ? right ? -offset : offset : 0) }
      }
      const title = chartAxisTitleCommand(chart, axis, titlePoint(0), titlePoint(1), textColor, horizontal ? 0 : right ? 90 : -90, layout.frame)
      if (title) commands.push(title)
    }
    axisScene(chart.yAxis, layout.min, layout.max)
    if (data.series.some((source, index) => chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name).rightAxis)) axisScene(rightAxis, rightExtent.min, rightExtent.max, true)
    chartLine(commands, project(world(0, floorRatio, front)), project(world(1, floorRatio, front)), chart.xAxis.line, textColor, plot)
    const categoryTicks = layout.dateAxis ? chartAxisTicks(layout.xMin!, layout.xMax!, chart.xAxis).map(value => ({
      category: chartDateAxisLabel(value, chart.xAxis, layout.date1904, layout.xMax! - layout.xMin!), coordinate: chartAxisRatio(value, layout.xMin!, layout.xMax!, chart.xAxis) }))
      : data.categories.map((category, index) => ({ category, coordinate: categoryRatio(index) }))
    categoryTicks.forEach(({ category, coordinate }) => {
      if (!Number.isFinite(coordinate)) return
      const point = project(world(coordinate, floorRatio, front))
      chartLine(commands, point, project(world(coordinate, floorRatio, back)), chart.xAxis.gridLine, '#cbd5e1', plot)
      if (chart.xAxis.tick?.visible) {
        const length = (chart.xAxis.tick.length ?? 4) * (chart.xAxis.tick.position === 'inside' ? -1 : 1)
        chartLine(commands, point, { x: point.x + (horizontal ? -length : 0), y: point.y + (horizontal ? 0 : length) }, chart.xAxis.tick, textColor)
      }
      if (chart.xAxis.label?.visible !== false) commands.push({ kind: 'text', x: point.x + (horizontal ? -8 : 0), y: point.y + (horizontal ? 0 : 12), text: category,
        font: chartFont({ ...chart.options?.appearance?.font, ...chart.xAxis.label }, 10), color: chart.xAxis.label?.color ?? textColor, align: horizontal ? 'right' : 'center', baseline: 'middle', rotation: chart.xAxis.label?.rotate,
        maxWidth: horizontal ? Math.max(12, plot.x - layout.frame.x - 12) : Math.max(12, plot.width / count), clip: layout.frame })
    })
    const xTitleOffset = Math.max(32, (chartAxisTitleFont(chart, chart.xAxis).fontSize ?? 11) + 21)
    const xTitlePoint = (value: number): { x: number; y: number } => {
      const point = project(world(value, floorRatio, front))
      return { x: point.x + (horizontal ? -xTitleOffset : 0), y: point.y + (horizontal ? 0 : xTitleOffset) }
    }
    const xTitle = chartAxisTitleCommand(chart, chart.xAxis, xTitlePoint(0), xTitlePoint(1), textColor, horizontal ? -90 : 0, layout.frame)
    if (xTitle) commands.push(xTitle)
    const depthAxis = chart.options?.depthAxis ?? {}
    chartLine(commands, project(world(1, floorRatio, front)), project(world(1, floorRatio, back)), depthAxis.line, textColor, plot)
    data.series.forEach((source, index) => {
      const depth = depthRatio((index + 0.5) / seriesCount), point = project(world(1, floorRatio, depth))
      chartLine(commands, point, project(world(0, floorRatio, depth)), depthAxis.gridLine, '#cbd5e1', plot)
      if (depthAxis.tick?.visible) {
        const length = (depthAxis.tick.length ?? 4) * (depthAxis.tick.position === 'inside' ? -1 : 1)
        chartLine(commands, point, { x: point.x + length, y: point.y }, depthAxis.tick, textColor)
      }
      if (depthAxis.label?.visible !== false) commands.push({ kind: 'text', x: point.x + 8, y: point.y, text: chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name).name ?? source.name,
        font: chartFont({ ...chart.options?.appearance?.font, ...depthAxis.label }, 10), color: depthAxis.label?.color ?? textColor, baseline: 'middle', rotation: depthAxis.label?.rotate, maxWidth: Math.max(12, plot.width / seriesCount), clip: layout.frame })
    })
    const depthTitlePoint = (value: number): { x: number; y: number } => {
      const point = project(world(1, floorRatio, value))
      return { x: point.x + 24, y: point.y + 20 }
    }
    const depthTitle = chartAxisTitleCommand(chart, depthAxis, depthTitlePoint(0), depthTitlePoint(1), textColor, 0, layout.frame)
    if (depthTitle) commands.push(depthTitle)
    const positive = new Map<string, number[]>(), negative = new Map<string, number[]>()
    const gapRatio = 1 / (1 + (chart.options?.bar?.gapWidth ?? 150) / 100), overlap = (chart.options?.bar?.overlap ?? 0) / 100
    for (const [seriesIndex, source] of data.series.entries()) {
      const style = chartSeriesStyle(chart, source.field, seriesIndex, palette, source.sourceName ?? source.name)
      const axis = style.rightAxis ? rightAxis : chart.yAxis, extent = style.rightAxis ? rightExtent : { min: layout.min, max: layout.max }
      const valueRatio = (value: number): number => chartAxisRatio(value, extent.min, extent.max, axis)
      const seriesBaseline = axis.scaleType === 'log' ? extent.min : Math.max(extent.min, Math.min(extent.max, 0))
      const grouped = chart.options?.threeD?.grouping === 'clustered' && !descriptor.stacked
      const depthGap = Math.max(0, (chart.options?.threeD?.gapDepth ?? 150) / 100)
      const depthWidth = Math.max(0.02, 0.75 / (descriptor.stacked || grouped ? 1 : seriesCount) / (1 + depthGap * 0.3))
      const depth = depthRatio(descriptor.stacked || grouped ? 0.35 : (seriesIndex + 0.5) / seriesCount)
      const band = gapRatio * (layout.categoryBandwidth === undefined ? 1 / count : layout.categoryBandwidth / (horizontal ? plot.height : plot.width))
      const width = grouped ? band / Math.max(1, seriesCount - (seriesCount - 1) * overlap) : band
      let previous: { categoryIndex: number; end: number; base: number; category: number; value: number } | undefined
      source.values.forEach((value, categoryIndex) => {
        if (value === null || !Number.isFinite(value) || !Number.isFinite(categoryRatio(categoryIndex)) || axis.scaleType === 'log' && value <= 0) { if (chart.options?.invalidValueStrategy !== 'link') previous = undefined; return }
        const stackKey = style.rightAxis ? 'right' : 'left', stacks = value >= 0 ? positive : negative
        const stack = descriptor.stacked ? stacks.get(stackKey) ?? Array(count).fill(0) as number[] : undefined
        const base = stack?.[categoryIndex] ?? seriesBaseline
        const end = descriptor.stacked ? finiteChartValue(base + value, '三维堆积合计') : value
        if (stack) { stack[categoryIndex] = end; stacks.set(stackKey, stack) }
        const center = categoryRatio(categoryIndex) + (grouped ? (seriesIndex - (seriesCount - 1) / 2) * width * (1 - overlap) : 0)
        const point = project(world(center, valueRatio(end), depth)), basePoint = project(world(center, valueRatio(base), depth))
        points[seriesIndex]!.push({ x: point.x, y: point.y, value, category: data.categories[categoryIndex] ?? '', categoryIndex, baseY: basePoint.y, baseX: basePoint.x, baseValue: base,
          ...(layout.dateAxis ? { xValue: data.categoryDateValues?.[categoryIndex] ?? undefined } : {}) })
        const pointStyle = chartPointStyle(style, categoryIndex), color = chartColorOpacity(pointStyle.color ?? palette[seriesIndex % palette.length]!, pointStyle.fillOpacity ?? 1)
        const pattern = (shade = 1) => chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1, shade)
        if (descriptor.family === 'column' || descriptor.family === 'bar') {
          const a = world(center - width / 2, valueRatio(base), depth - depthWidth / 2), b = world(center + width / 2, valueRatio(base), depth - depthWidth / 2)
          const c = world(center + width / 2, valueRatio(end), depth - depthWidth / 2), d = world(center - width / 2, valueRatio(end), depth - depthWidth / 2)
          const e = { ...a, z: depth + depthWidth / 2 }, f = { ...b, z: depth + depthWidth / 2 }, g = { ...c, z: depth + depthWidth / 2 }, h = { ...d, z: depth + depthWidth / 2 }
          addFace([a, b, c, d], color, pointStyle.border, pattern()); addFace([e, f, g, h], chartShade(color, 0.76), pointStyle.border, pattern(0.76))
          addFace([a, e, h, d], chartShade(color, 0.68), pointStyle.border, pattern(0.68)); addFace([b, f, g, c], chartShade(color, 0.88), pointStyle.border, pattern(0.88))
          addFace([d, c, g, h], chartShade(color, 1.18), pointStyle.border, pattern(1.18)); addFace([a, b, f, e], chartShade(color, 0.6), pointStyle.border, pattern(0.6))
        } else if (previous && (chart.options?.invalidValueStrategy === 'link' || categoryIndex === previous.categoryIndex + 1)) {
          const a = world(previous.category, valueRatio(previous.end), depth - depthWidth / 2), b = world(center, valueRatio(end), depth - depthWidth / 2)
          const c = { ...b, z: depth + depthWidth / 2 }, d = { ...a, z: depth + depthWidth / 2 }
          addFace([a, b, c, d], color, pointStyle.border, pattern())
          if (descriptor.family === 'area') {
            const aBase = world(previous.category, valueRatio(previous.base), depth - depthWidth / 2), bBase = world(center, valueRatio(base), depth - depthWidth / 2)
            if (previous.value < 0 && value > 0 || previous.value > 0 && value < 0) {
              const scale = Math.max(Math.abs(previous.value), Math.abs(value)), aWeight = Math.abs(previous.value) / scale, bWeight = Math.abs(value) / scale
              const t = aWeight / (aWeight + bWeight), crossCategory = previous.category + (center - previous.category) * t
              const crossBase = chartInterpolate(previous.base, base, t), crossing = world(crossCategory, valueRatio(crossBase), depth - depthWidth / 2)
              addFace([aBase, a, crossing], chartShade(color, 0.84), pointStyle.border, pattern(0.84)); addFace([crossing, b, bBase], chartShade(color, 0.84), pointStyle.border, pattern(0.84))
              addFace([{ ...aBase, z: depth + depthWidth / 2 }, d, { ...crossing, z: depth + depthWidth / 2 }], chartShade(color, 0.68), pointStyle.border, pattern(0.68))
              addFace([{ ...crossing, z: depth + depthWidth / 2 }, c, { ...bBase, z: depth + depthWidth / 2 }], chartShade(color, 0.68), pointStyle.border, pattern(0.68))
            } else {
              addFace([aBase, a, b, bBase], chartShade(color, 0.84), pointStyle.border, pattern(0.84))
              addFace([{ ...aBase, z: depth + depthWidth / 2 }, d, c, { ...bBase, z: depth + depthWidth / 2 }], chartShade(color, 0.68), pointStyle.border, pattern(0.68))
            }
          }
        }
        if (descriptor.family === 'line') faces.push({ points: [point], color: chartColorOpacity(pointStyle.point?.color ?? pointStyle.color!, pointStyle.fillOpacity ?? 1), border: pointStyle.border, pattern: pattern(),
          marker: { size: pointStyle.point?.size ?? 5, shape: pointStyle.point?.shape ?? 'circle' } })
        previous = { categoryIndex, end, base, category: center, value }
      })
    }
  }
  faces.sort((a, b) => a.points.reduce((sum, point) => sum + point.z, 0) / a.points.length - b.points.reduce((sum, point) => sum + point.z, 0) / b.points.length)
  for (const face of faces) {
    if (face.marker) chartSymbol(commands, face.points[0]!, face.marker.size, face.color, face.marker.shape, face.border, plot, face.pattern)
    else chartPolygon(commands, face.points, face.color, face.border, plot, face.pattern)
  }
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
