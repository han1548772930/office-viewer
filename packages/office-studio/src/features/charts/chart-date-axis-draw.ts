import type { ChartModel } from '../../domain/model'
import type { DrawCommand } from '../../rendering/contracts'
import type { ChartLayout } from './chart-engine'
import type { ChartRenderTheme } from './chart-theme'
import { chartAxisRatio, chartFont, chartLine } from './chart-geometry'
import { chartDateAxisLabel } from './chart-date-axis'
import { definedChartStyleFields } from './chart-options'
import { normalizeChartColor } from './chart-color'
import { tagChartCanvasElementCommands, withChartCanvasElement } from './chart-element-interaction'

/** Shared date chrome leaves native content and the source date identities intact. */
export function chartDateAxisCommands(chart: ChartModel, layout: ChartLayout, color: string, theme: ChartRenderTheme = {}): DrawCommand[] {
  if (!layout.dateAxis || layout.xMin === undefined || layout.xMax === undefined) return []
  const { plot } = layout, axis = chart.xAxis, horizontal = layout.horizontal, commands: DrawCommand[] = []
  for (const tick of layout.dateTicks ?? []) {
    const first = commands.length
    const ratio = chartAxisRatio(tick.value, layout.xMin, layout.xMax, axis), coordinate = horizontal ? plot.y + ratio * plot.height : plot.x + ratio * plot.width
    chartLine(commands, horizontal ? { x: plot.x, y: coordinate } : { x: coordinate, y: plot.y },
      horizontal ? { x: plot.x + plot.width, y: coordinate } : { x: coordinate, y: plot.y + plot.height }, { visible: false, ...definedChartStyleFields(axis.gridLine) }, theme.splitLineColor ?? '#e5e7eb', plot)
    tagChartCanvasElementCommands(commands, first, { kind: 'gridLine', axisId: 'xAxis' })
    if (axis.label?.visible !== false) commands.push(withChartCanvasElement({ kind: 'text', x: horizontal ? plot.x - 8 : coordinate, y: horizontal ? coordinate : plot.y + plot.height + 14,
      text: chartDateAxisLabel(tick.serial, axis, layout.date1904, (layout.dateSourceMax ?? 1) - (layout.dateSourceMin ?? 0)),
      font: chartFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(axis.label) }, 10),
      color: normalizeChartColor(axis.label?.color ?? color), align: horizontal ? 'right' : 'center', baseline: 'middle', rotation: axis.label?.rotate, clip: layout.frame }, { kind: 'axis', axisId: 'xAxis', part: 'label' }))
    if (axis.tick?.visible) {
      const tickStart = commands.length
      const length = (axis.tick.length ?? 4) * (axis.tick.position === 'inside' ? -1 : 1)
      chartLine(commands, horizontal ? { x: plot.x, y: coordinate } : { x: coordinate, y: plot.y + plot.height },
        horizontal ? { x: plot.x - length, y: coordinate } : { x: coordinate, y: plot.y + plot.height + length }, axis.tick, theme.axisColor ?? '#64748b', layout.frame)
      tagChartCanvasElementCommands(commands, tickStart, { kind: 'tick', axisId: 'xAxis' })
    }
  }
  return commands
}
