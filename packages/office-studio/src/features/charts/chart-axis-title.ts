import type { ChartAxisOptions, ChartFontOptions, ChartModel } from '../../domain/model'
import type { RenderRect, TextCommand } from '../../rendering/contracts'
import type { ChartLayout } from './chart-engine'
import { chartFont, type ChartVector } from './chart-geometry'
import { definedChartStyleFields } from './chart-options'

type TitleAxis = Pick<ChartAxisOptions, 'title' | 'titleStyle' | 'titleAlignment' | 'reverse'>
export const CHART_AXIS_TITLE_DEFAULT_FONT_SIZE = 12

export function chartAxisTitleFont(chart: Pick<ChartModel, 'options'>, axis: Pick<ChartAxisOptions, 'titleStyle'>): ChartFontOptions {
  return { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(axis.titleStyle) }
}

export function chartAxisTitleSpace(chart: Pick<ChartModel, 'options'>, axis: TitleAxis): number {
  return axis.title ? (chartAxisTitleFont(chart, axis).fontSize ?? CHART_AXIS_TITLE_DEFAULT_FONT_SIZE) + 7 : 0
}

/** Fixed wke positions start/end against the physical title box, independently
 * of data-axis reversal. The Office-style rotated text stays inside that span. */
export function chartAxisTitleCommand(chart: ChartModel, axis: TitleAxis, start: ChartVector, end: ChartVector,
  color: string, rotation = 0, clip?: RenderRect): TextCommand | undefined {
  if (!axis.title) return undefined
  const font = chartAxisTitleFont(chart, axis), alignment = axis.titleAlignment ?? 'center'
  const horizontal = Math.abs(end.x - start.x) >= Math.abs(end.y - start.y)
  const forward = horizontal ? start.x <= end.x : start.y <= end.y
  const first = forward ? start : end, last = forward ? end : start
  const position = alignment === 'start' ? 0 : alignment === 'end' ? 1 : 0.5
  const angle = rotation * Math.PI / 180
  const direction = (last.x - first.x) * Math.cos(angle) + (last.y - first.y) * Math.sin(angle)
  const align = alignment === 'center' ? 'center' : (alignment === 'start') === (direction >= 0) ? 'left' : 'right'
  return { kind: 'text', x: first.x + (last.x - first.x) * position, y: first.y + (last.y - first.y) * position,
    text: axis.title, font: chartFont(font, CHART_AXIS_TITLE_DEFAULT_FONT_SIZE), color: font.color ?? color, align, baseline: 'middle', rotation,
    maxWidth: Math.max(1, Math.hypot(end.x - start.x, end.y - start.y)), ...(clip ? { clip } : {}) }
}

/** Shared chrome geometry for native content and common Canvas content. It
 * binds the category/value axes to their real sides, including horizontal Bar. */
export function chartRectangularAxisTitleCommands(chart: ChartModel,
  layout: Pick<ChartLayout, 'frame' | 'plot' | 'horizontal' | 'rightMin' | 'rightMax'>, color: string): TextCommand[] {
  const { plot, frame } = layout, horizontal = layout.horizontal
  const bottom = horizontal ? chart.yAxis : chart.xAxis, side = horizontal ? chart.xAxis : chart.yAxis
  const titleY = plot.y + plot.height + 30 + (chartAxisTitleFont(chart, bottom).fontSize ?? CHART_AXIS_TITLE_DEFAULT_FONT_SIZE) / 2
  const titleX = plot.x - 54 - chartAxisTitleSpace(chart, side) / 2
  const commands: TextCommand[] = []
  function append(axis: TitleAxis, start: ChartVector, end: ChartVector, rotation: number): void {
    const command = chartAxisTitleCommand(chart, axis, start, end, color, rotation, frame)
    if (command) commands.push(command)
  }
  append(bottom, { x: plot.x, y: titleY }, { x: plot.x + plot.width, y: titleY }, 0)
  append(side, { x: titleX, y: plot.y }, { x: titleX, y: plot.y + plot.height }, -90)
  if (layout.rightMin !== undefined && layout.rightMax !== undefined) {
    const axis = chart.options?.rightYAxis ?? { scaleType: 'linear' }, offset = chartAxisTitleSpace(chart, axis) / 2
    append(axis, horizontal ? { x: plot.x, y: plot.y - 40 - offset } : { x: plot.x + plot.width + 54 + offset, y: plot.y },
      horizontal ? { x: plot.x + plot.width, y: plot.y - 40 - offset } : { x: plot.x + plot.width + 54 + offset, y: plot.y + plot.height }, horizontal ? 0 : 90)
  }
  return commands
}
