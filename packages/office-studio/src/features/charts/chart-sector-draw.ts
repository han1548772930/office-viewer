import type { ChartSeriesFillPattern, ChartSeriesOptions } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartAutomaticGradient } from './chart-auto-gradient'
import { chartArc, chartLine, chartPolar, chartPolygon, type ChartVector } from './chart-geometry'

export function drawChartSector(commands: DrawCommand[], center: ChartVector, inner: number, outer: number, start: number, end: number, color: string,
  border?: ChartSeriesOptions['border'], clip?: RenderRect, pattern?: ChartSeriesFillPattern, gradient?: ChartAutomaticGradient): void {
  const count = Math.max(2, Math.ceil(Math.abs(end - start) / (Math.PI / 48)))
  const paint = gradient && { ...gradient, bounds: gradient.bounds ?? { x: center.x - outer, y: center.y - outer, width: outer * 2, height: outer * 2 } }
  for (let index = 0; index < count; index++) {
    const a = start + (end - start) * index / count, b = start + (end - start) * (index + 1) / count
    chartPolygon(commands, inner > 0 ? [chartPolar(center, inner, a), chartPolar(center, outer, a), chartPolar(center, outer, b), chartPolar(center, inner, b)]
      : [center, chartPolar(center, outer, a), chartPolar(center, outer, b)], color, undefined, clip, pattern, paint)
  }
  if (border && border.visible !== false) {
    const outerArc = chartArc(center, outer, start, end), innerArc = inner > 0 ? chartArc(center, inner, start, end) : [center]
    for (const arc of [outerArc, innerArc]) for (let index = 1; index < arc.length; index++) chartLine(commands, arc[index - 1]!, arc[index]!, border, color, clip)
    chartLine(commands, outerArc[0]!, innerArc[0]!, border, color, clip); chartLine(commands, outerArc.at(-1)!, innerArc.at(-1)!, border, color, clip)
  }
}
