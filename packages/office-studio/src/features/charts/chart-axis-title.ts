import type { ChartAxisOptions, ChartFontOptions, ChartModel } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { chartFont, type ChartVector } from './chart-geometry'
import { definedChartStyleFields } from './chart-options'

type TitleAxis = Pick<ChartAxisOptions, 'title' | 'titleStyle' | 'titleAlignment' | 'reverse'>

export function chartAxisTitleFont(chart: ChartModel, axis: Pick<ChartAxisOptions, 'titleStyle'>): ChartFontOptions {
  return { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(axis.titleStyle) }
}

export function chartAxisTitleSpace(chart: ChartModel, axis: TitleAxis): number {
  return axis.title ? (chartAxisTitleFont(chart, axis).fontSize ?? 11) + 7 : 0
}

/** Start/end follow the logical axis, including reversal. Text stays within its axis span. */
export function chartAxisTitleCommand(chart: ChartModel, axis: TitleAxis, start: ChartVector, end: ChartVector,
  color: string, rotation = 0, clip?: RenderRect): DrawCommand | undefined {
  if (!axis.title) return undefined
  const font = chartAxisTitleFont(chart, axis), alignment = axis.titleAlignment ?? 'center'
  const ratio = alignment === 'start' ? 0 : alignment === 'end' ? 1 : 0.5
  const position = axis.reverse ? 1 - ratio : ratio
  const angle = rotation * Math.PI / 180
  const direction = ((end.x - start.x) * Math.cos(angle) + (end.y - start.y) * Math.sin(angle)) * (axis.reverse ? -1 : 1)
  const align = alignment === 'center' ? 'center' : (alignment === 'start') === (direction >= 0) ? 'left' : 'right'
  return { kind: 'text', x: start.x + (end.x - start.x) * position, y: start.y + (end.y - start.y) * position,
    text: axis.title, font: chartFont(font, 11), color: font.color ?? color, align, baseline: 'middle', rotation,
    maxWidth: Math.max(1, Math.hypot(end.x - start.x, end.y - start.y)), ...(clip ? { clip } : {}) }
}
