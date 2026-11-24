import type { ChartAxisOptions, ChartFontOptions, ChartModel } from '../../domain/model'
import type { DrawCommand, RenderRect, TextCommand } from '../../rendering/contracts'
import type { ChartData, ChartLayout } from './chart-engine'
import type { ChartRenderTheme } from './chart-theme'
import { chartRectangularAxisTitleCommands } from './chart-axis-title'
import { chartColorOpacity, chartFont, chartLine, chartRect } from './chart-geometry'
import { chartHeatmapPointStyle } from './chart-derived'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { definedChartStyleFields } from './chart-options'
import { chartHeatmapBoundaryText, chartHeatmapCellRect, chartHeatmapFieldNames, chartHeatmapLabelAnchor,
  chartHeatmapLabelText, chartHeatmapLegendBand, chartHeatmapRatioColor, chartHeatmapValueVisible, resolveChartHeatmapLabel,
  type ChartHeatmapScale } from './chart-heatmap'

export interface ChartHeatmapLegendPiece { pieceIndex: number; rect: RenderRect; hit: RenderRect }
export interface ChartHeatmapLegend {
  visible: boolean
  horizontal: boolean
  position: 'top' | 'right' | 'bottom' | 'left'
  frame: RenderRect
  font: ChartFontOptions
  color: string
  scale: ChartHeatmapScale
  bar?: RenderRect
  pieces: ChartHeatmapLegendPiece[]
  commands: DrawCommand[]
}

export function drawChartHeatmapAxes(commands: DrawCommand[], data: ChartData, chart: ChartModel,
  layout: ChartLayout, color: string, theme: ChartRenderTheme): void {
  const { plot, frame } = layout, names = chartHeatmapFieldNames(data, chart)
  const axisColor = theme.axisColor ?? '#94a3b8', gridColor = theme.splitLineColor ?? '#e2e8f0'
  const xAxis = chart.xAxis, yAxis = chart.yAxis
  chartLine(commands, { x: plot.x, y: plot.y + plot.height }, { x: plot.x + plot.width, y: plot.y + plot.height }, xAxis.line, axisColor)
  chartLine(commands, { x: plot.x, y: plot.y }, { x: plot.x, y: plot.y + plot.height }, yAxis.line, axisColor)
  function drawCategories(axis: ChartAxisOptions, labels: string[], horizontal: boolean): void {
    const font = { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(axis.label) }
    labels.forEach((label, index) => {
      const ratio = horizontal ? (axis.reverse ? labels.length - index - 0.5 : index + 0.5) / labels.length
        : (axis.reverse ? index + 0.5 : labels.length - index - 0.5) / labels.length
      const coordinate = horizontal ? plot.x + ratio * plot.width : plot.y + ratio * plot.height
      if (axis.tick?.visible) {
        const length = (axis.tick.length ?? 4) * (axis.tick.position === 'inside' ? -1 : 1)
        const start = horizontal ? { x: coordinate, y: plot.y + plot.height } : { x: plot.x, y: coordinate }
        chartLine(commands, start, { x: start.x - (horizontal ? 0 : length), y: start.y + (horizontal ? length : 0) }, axis.tick, axisColor)
      }
      if (axis.label?.visible !== false) commands.push({ kind: 'text', x: horizontal ? coordinate : plot.x - 8, y: horizontal ? plot.y + plot.height + 14 : coordinate,
        text: label, font: chartFont(font, 10), color: axis.label?.color ?? color, align: horizontal ? 'center' : 'right', baseline: 'middle', rotation: axis.label?.rotate,
        maxWidth: horizontal ? Math.max(1, plot.width / labels.length - 4) : Math.max(1, plot.x - frame.x - 12), clip: frame })
    })
    if (axis.gridLine?.visible) for (let index = 0; index <= labels.length; index++) {
      const coordinate = horizontal ? plot.x + index / labels.length * plot.width : plot.y + index / labels.length * plot.height
      chartLine(commands, horizontal ? { x: coordinate, y: plot.y } : { x: plot.x, y: coordinate },
        horizontal ? { x: coordinate, y: plot.y + plot.height } : { x: plot.x + plot.width, y: coordinate }, axis.gridLine, gridColor, plot)
    }
  }
  drawCategories(xAxis, names, true); drawCategories(yAxis, data.categories, false)
  for (const command of chartRectangularAxisTitleCommands(chart, layout, color)) commands.push(command)
}

export function drawChartHeatmapCells(commands: DrawCommand[], data: ChartData, chart: ChartModel,
  layout: ChartLayout, color: string, palette: readonly string[], scale: ChartHeatmapScale): void {
  data.series.forEach((series, seriesPosition) => series.values.forEach((value, categoryIndex) => {
    if (value === null || !chartHeatmapValueVisible(scale, value)) return
    const rect = chartHeatmapCellRect(data, chart, layout.plot, seriesPosition, categoryIndex)
    if (!rect) return
    const style = chartHeatmapPointStyle(chart, series.field, seriesPosition, categoryIndex, value, scale.extent, palette, series.sourceName ?? series.name, scale)
    chartRect(commands, rect, chartColorOpacity(style.color!, style.fillOpacity ?? 1), style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    const label = resolveChartHeatmapLabel(chart, style.label, color)
    const text = chartHeatmapLabelText(value, series.formattedValues?.[categoryIndex], label)
    if (label.visible && text) {
      const { nativePosition: _nativePosition, ...anchor } = chartHeatmapLabelAnchor(rect, label)
      commands.push({ kind: 'text', ...anchor, text, font: chartFont(label, 12), color: label.color!, maxWidth: Math.max(1, rect.width - 4), clip: layout.frame })
    }
  }))
}

/** The actual legend is a shared Canvas scene in both adapters. The internal
 * native visualMap is noninteractive; surface pointer forwarding queries these
 * exact rectangles rather than depending on an unattached Canvas DOM event. */
export function createChartHeatmapLegend(chart: Pick<ChartModel, 'legend' | 'options'>, layout: Pick<ChartLayout, 'frame' | 'plot'>,
  scale: ChartHeatmapScale, color: string): ChartHeatmapLegend {
  const visible = chart.legend.visible, position = chart.legend.position, horizontal = position === 'top' || position === 'bottom'
  const { frame, plot } = layout, font = { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(chart.legend.font) }
  const fontSize = font.fontSize ?? 11, commands: DrawCommand[] = []
  const legend: ChartHeatmapLegend = { visible, position, horizontal, frame, font, color: font.color ?? color, scale, pieces: [], commands }
  if (!visible) return legend
  function text(value: string, x: number, y: number, maxWidth: number, align: TextCommand['align'] = 'center'): void {
    commands.push({ kind: 'text', x, y, text: value, font: chartFont(font, 11), color: legend.color, align, baseline: 'middle', maxWidth: Math.max(1, maxWidth), clip: frame })
  }
  if (scale.type === 'continuous') {
    const sideWidth = position === 'left' ? Math.max(1, plot.x - frame.x - 54) : Math.max(1, frame.x + frame.width - plot.x - plot.width - 12)
    const centerX = horizontal ? frame.x + frame.width / 2 : position === 'left' ? frame.x + sideWidth / 2 + 4 : plot.x + plot.width + 12 + sideWidth / 2
    const barWidth = horizontal ? Math.max(1, Math.min(160, (frame.width - 48) * 0.42)) : 12
    const barHeight = horizontal ? 12 : Math.max(1, Math.min(140, plot.height - fontSize * 2 - 24))
    const y = horizontal ? position === 'top' ? plot.y - chartHeatmapLegendBand(frame.width, scale.type, fontSize) / 2 - 6
      : frame.y + frame.height - chartHeatmapLegendBand(frame.width, scale.type, fontSize) / 2 - 6 : plot.y + (plot.height - barHeight) / 2
    const bar: RenderRect = { x: centerX - barWidth / 2, y, width: barWidth, height: barHeight }
    legend.bar = bar
    // DrawCommand intentionally has no backend-owned gradient object. A bounded
    // ramp gives WebGPU, Canvas2D and ECharts graphic the same sample colors.
    const count = 128
    for (let index = 0; index < count; index++) {
      const ratio = (index + 0.5) / count, rect = horizontal ? { x: bar.x + index / count * bar.width, y: bar.y, width: bar.width / count, height: bar.height }
        : { x: bar.x, y: bar.y + (count - index - 1) / count * bar.height, width: bar.width, height: bar.height / count }
      chartRect(commands, rect, chartHeatmapRatioColor(scale, scale.extent.min === scale.extent.max ? 0.5 : ratio), undefined, frame)
    }
    const minText = chartHeatmapBoundaryText(scale.extent.min, chart.options?.heatmap?.unit), maxText = chartHeatmapBoundaryText(scale.extent.max, chart.options?.heatmap?.unit)
    if (horizontal) {
      const available = Math.max(1, (frame.width - bar.width) / 2 - 20)
      text(minText, bar.x - 12, bar.y + 6, available, 'right'); text(maxText, bar.x + bar.width + 12, bar.y + 6, available, 'left')
    } else {
      text(maxText, centerX, bar.y - 12 - fontSize / 2, sideWidth - 8)
      text(minText, centerX, bar.y + bar.height + 12 + fontSize / 2, sideWidth - 8)
    }
  } else {
    const count = scale.pieces.length, columns = horizontal ? Math.max(1, Math.min(count, Math.floor((frame.width - 16) / 108))) : 1
    const rows = Math.ceil(count / columns), step = horizontal ? fontSize + 8 : Math.min(fontSize + 8, plot.height / Math.max(1, rows))
    const itemWidth = horizontal ? Math.max(1, (frame.width - 16) / columns) : position === 'left' ? Math.max(1, plot.x - frame.x - 62) : Math.max(1, frame.x + frame.width - plot.x - plot.width - 20)
    const startX = horizontal || position === 'left' ? frame.x + 8 : plot.x + plot.width + 12
    const startY = horizontal ? position === 'top' ? plot.y - chartHeatmapLegendBand(frame.width, scale.type, fontSize) + 8 : frame.y + frame.height - rows * step - 4
      : plot.y + Math.max(0, (plot.height - rows * step) / 2)
    for (let visibleIndex = 0; visibleIndex < count; visibleIndex++) {
      const pieceIndex = horizontal ? visibleIndex : count - visibleIndex - 1, piece = scale.pieces[pieceIndex]!
      const x = startX + visibleIndex % columns * itemWidth, y = startY + Math.floor(visibleIndex / columns) * step
      const rect = { x, y, width: 12, height: Math.min(12, step) }, hit = { x, y: y - 2, width: itemWidth, height: step }
      legend.pieces.push({ pieceIndex, rect, hit })
      chartRect(commands, rect, piece.color, undefined, frame)
      // Fixed piecewise labels do not append the continuous boundary unit.
      text(piece.text, x + 18, y + rect.height / 2, itemWidth - 24, 'left')
    }
  }
  return legend
}
