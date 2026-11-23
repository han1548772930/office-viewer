import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartModel, ChartSeriesFillPattern, ChartSeriesOptions } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartData, ChartLayout, ChartPoint } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { deriveChartData, chartBubblePointSize, chartCandlestickPointStyle, chartHeatmapExtent, chartHeatmapPointStyle, chartHierarchyPointStyle, chartPointStyle, type ChartComputedStatistics, type ChartDerivedData } from './chart-derived'
import { chartHierarchyItems } from './chart-data'
import { normalizeChartColor } from './chart-color'
import { chartInterpolate, chartLinearRatio } from './chart-math'
import { layoutDerivedChartData } from './chart-layout'
import {
  chartArc, chartAxisRatio, chartAxisTicks, chartBezier, chartCircle, chartColorOpacity, chartFont, chartLabelText, chartLine,
  chartPolar, chartPolygon, chartRect, chartSymbol, formatChartNumber, type ChartVector,
} from './chart-geometry'
import { chartPalette, chartSeriesColorIsAuthored, chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { chart3DScene } from './chart-3d'
import { chartFunnelLayout, chartGaugeTiles, chartPieLayout } from './chart-pie-layout'
import { chartGaugeAxisLabelAlignment, chartGaugeAxisLabelRadius, chartGaugeColorStops, chartGaugeDetailText, chartGaugePointerPoints, chartGaugeRatio, chartGaugeSweep, resolveChartGaugeOptions } from './chart-gauge'
import { chartRadarAxisLabel, chartRadarAxisLine, chartRadarGeometry } from './chart-radar'
import { chartBoxplotGeometry, chartBoxplotStyle } from './chart-boxplot'
import { chartAxisTitleCommand, chartAxisTitleFont, chartAxisTitleSpace } from './chart-axis-title'
import { chartDateAxisLabel } from './chart-date-axis'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartChordLayout, chartRelationLayout, chartRibbon, chartSankeyLayout, chartSunburstLayout, chartTreemapLayout, chartWordCloudLayout } from './chart-special-layout'
import type { ChartRenderTheme } from './chart-theme'

export function drawChartData(source: ChartData, chart: ChartModel, theme: ChartRenderTheme = {}, computed?: ChartComputedStatistics): DrawCommand[] {
  const palette = chartPalette(chart, theme.palette), derived = deriveChartData(source, chart, computed)
  return drawDerivedChartData(derived, chart, theme, palette)
}

export function drawDerivedChartData(derived: ChartDerivedData, chart: ChartModel, theme: ChartRenderTheme = {}, palette = chartPalette(chart, theme.palette), preparedLayout?: ChartLayout): DrawCommand[] {
  const data = derived.data, layout = preparedLayout ?? layoutDerivedChartData(derived, chart, palette), descriptor = resolveChartType(chart.type)
  const background = chart.options?.appearance?.background ?? theme.backgroundColor ?? '#ffffff', textColor = chart.options?.appearance?.font?.color ?? theme.textColor ?? '#64748b'
  const border = chart.options?.appearance?.border
  const commands: DrawCommand[] = [{ kind: 'rect', ...layout.frame, fill: background,
    ...(border?.visible === false ? {} : { stroke: border?.color ?? theme.axisColor ?? '#cbd5e1', strokeWidth: border?.width ?? 1 }) }]
  const titleFont = { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(chart.options?.titleStyle) }, subtitleFont = { ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(chart.options?.subtitleStyle) }
  const titleSize = titleFont.fontSize ?? 14, subtitleSize = subtitleFont.fontSize ?? 11, titleBand = chart.title ? Math.max(28, titleSize + 14) : 0
  if (chart.title) commands.push({ kind: 'text', x: layout.frame.x + layout.frame.width / 2, y: layout.frame.y + Math.max(20, 12 + titleSize / 2), text: chart.title, font: chartFont({ bold: true, ...titleFont }, 14), color: titleFont.color ?? textColor, align: 'center', baseline: 'middle', maxWidth: Math.max(1, layout.frame.width - 20), clip: layout.frame })
  if (chart.subtitle) commands.push({ kind: 'text', x: layout.frame.x + layout.frame.width / 2, y: layout.frame.y + titleBand + Math.max(20, 12 + subtitleSize / 2), text: chart.subtitle, font: chartFont(subtitleFont, 11), color: subtitleFont.color ?? textColor, align: 'center', baseline: 'middle', maxWidth: Math.max(1, layout.frame.width - 20), clip: layout.frame })
  if (!data.categories.length || !data.series.length || descriptor.family === 'gauge' && !data.series[0]!.values.some(Number.isFinite)) {
    commands.push({ kind: 'text', x: layout.plot.x + layout.plot.width / 2, y: layout.plot.y + layout.plot.height / 2,
      text: '暂无可绘制数据', font: chartFont(chart.options?.appearance?.font, 11), color: textColor, align: 'center', baseline: 'middle', maxWidth: layout.plot.width })
    return normalizeChartCommands(commands)
  }
  if (descriptor.threeD) {
    commands.push(...chart3DScene(data, chart, layout, textColor, palette).commands)
    drawPointLabels(commands, data, chart, layout, textColor, palette)
  } else if (descriptor.family === 'pie') drawPie(commands, data, chart, layout, textColor, palette)
  else if (descriptor.family === 'radar') {
    drawRadar(commands, data, chart, layout, textColor, palette, theme)
    drawPointLabels(commands, data, chart, layout, textColor, palette)
  }
  else if (descriptor.family === 'relation' || descriptor.family === 'sankey') drawGraph(commands, data, chart, layout, textColor, palette)
  else if (descriptor.family === 'chord') drawChord(commands, data, chart, layout, textColor, palette)
  else if (descriptor.family === 'wordCloud') drawWordCloud(commands, data, chart, layout, palette)
  else if (descriptor.family === 'treemap') drawTreemap(commands, data, chart, layout, textColor, palette)
  else if (descriptor.family === 'sunburst') drawSunburst(commands, data, chart, layout, textColor, palette)
  else if (descriptor.family === 'gauge') drawGauge(commands, data, chart, layout, textColor, palette, theme)
  else if (descriptor.family === 'funnel') drawFunnel(commands, data, chart, layout, textColor, palette)
  else {
    drawAxes(commands, data, chart, layout, textColor, theme)
    if (descriptor.family === 'boxplot') drawBoxplot(commands, derived, chart, layout, textColor, palette)
    else if (descriptor.family === 'candlestick') drawCandlestick(commands, data, chart, layout, palette)
    else if (descriptor.family === 'waterfall') drawWaterfall(commands, derived, chart, layout, palette)
    else if (descriptor.family === 'heatmap') drawHeatmap(commands, data, chart, layout, textColor, palette)
    else drawCartesian(commands, data, chart, layout, palette)
    if (derived.pareto) {
      const color = palette[1 % palette.length]!, points = layout.series.at(-1)!.points
      for (let index = 1; index < points.length; index++) chartLine(commands, points[index - 1]!, points[index]!, { width: 2, color }, color, layout.plot)
      points.forEach(point => chartCircle(commands, point, 2, color, undefined, layout.plot))
    }
    if (descriptor.family !== 'boxplot') drawPointLabels(commands, data, chart, layout, textColor, palette)
    drawTrendlinesAndErrors(commands, derived, chart, layout, textColor, palette)
  }
  drawLegend(commands, data, chart, layout, textColor, palette)
  return normalizeChartCommands(commands)
}

/** The same generated scene can be submitted to either backend, including named and HSL author colors. */
function normalizeChartCommands(commands: DrawCommand[]): DrawCommand[] {
  for (const command of commands) {
    if (command.kind === 'rect') {
      command.fill = normalizeChartColor(command.fill)
      if (command.stroke !== undefined) command.stroke = normalizeChartColor(command.stroke)
    } else if (command.kind === 'polygon') command.fill = normalizeChartColor(command.fill)
    else if (command.kind === 'line' || command.kind === 'text') command.color = normalizeChartColor(command.color)
  }
  return commands
}

function drawAxes(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, theme: ChartRenderTheme): void {
  const { plot } = layout, horizontal = layout.horizontal, descriptor = resolveChartType(chart.type), numericX = layout.xMin !== undefined
  const gridColor = theme.splitLineColor ?? '#e5e7eb', axisColor = theme.axisColor ?? '#64748b'
  const yAxis = chart.yAxis, xAxis = chart.xAxis
  for (const value of chartAxisTicks(layout.min, layout.max, yAxis)) {
    const ratio = chartAxisRatio(value, layout.min, layout.max, yAxis), coordinate = horizontal ? plot.x + ratio * plot.width : plot.y + (1 - ratio) * plot.height
    const start = horizontal ? { x: coordinate, y: plot.y } : { x: plot.x, y: coordinate }, end = horizontal ? { x: coordinate, y: plot.y + plot.height } : { x: plot.x + plot.width, y: coordinate }
    chartLine(commands, start, end, yAxis.gridLine, gridColor, plot)
    if (yAxis.label?.visible !== false && chart.type !== 'heatmap') commands.push({ kind: 'text', x: horizontal ? coordinate : plot.x - 8, y: horizontal ? plot.y + plot.height + 14 : coordinate,
      text: `${formatChartNumber(value, yAxis.format)}${descriptor.percent && !yAxis.format ? '%' : ''}`, font: chartFont({ ...chart.options?.appearance?.font, ...yAxis.label }, 10), color: yAxis.label?.color ?? color,
      align: horizontal ? 'center' : 'right', baseline: 'middle', rotation: yAxis.label?.rotate })
    if (yAxis.tick?.visible) {
      const length = (yAxis.tick.length ?? 4) * (yAxis.tick.position === 'inside' ? 1 : -1)
      chartLine(commands, horizontal ? { x: coordinate, y: plot.y + plot.height } : { x: plot.x, y: coordinate }, horizontal ? { x: coordinate, y: plot.y + plot.height - length } : { x: plot.x + length, y: coordinate }, yAxis.tick, axisColor)
    }
  }
  const zeroRatio = chartAxisRatio(yAxis.scaleType === 'log' ? layout.min : 0, layout.min, layout.max, yAxis)
  const zero = Math.max(0, Math.min(1, zeroRatio))
  chartLine(commands, horizontal ? { x: plot.x + zero * plot.width, y: plot.y } : { x: plot.x, y: plot.y + (1 - zero) * plot.height },
    horizontal ? { x: plot.x + zero * plot.width, y: plot.y + plot.height } : { x: plot.x + plot.width, y: plot.y + (1 - zero) * plot.height }, { ...xAxis.line, width: xAxis.line?.width ?? 1.25 }, axisColor)
  chartLine(commands, horizontal ? { x: plot.x, y: plot.y + plot.height } : { x: plot.x, y: plot.y },
    horizontal ? { x: plot.x + plot.width, y: plot.y + plot.height } : { x: plot.x, y: plot.y + plot.height }, yAxis.line, axisColor)
  const categories = chart.type === 'heatmap' ? data.categories : layout.categories
  if (numericX) for (const value of chartAxisTicks(layout.xMin!, layout.xMax!, xAxis)) {
    const ratio = chartAxisRatio(value, layout.xMin!, layout.xMax!, xAxis), coordinate = horizontal ? plot.y + ratio * plot.height : plot.x + ratio * plot.width
    chartLine(commands, horizontal ? { x: plot.x, y: coordinate } : { x: coordinate, y: plot.y }, horizontal ? { x: plot.x + plot.width, y: coordinate } : { x: coordinate, y: plot.y + plot.height }, xAxis.gridLine, gridColor, plot)
    if (xAxis.label?.visible !== false) commands.push({ kind: 'text', x: horizontal ? plot.x - 8 : coordinate, y: horizontal ? coordinate : plot.y + plot.height + 14,
      text: layout.dateAxis ? chartDateAxisLabel(value, xAxis, layout.date1904, layout.xMax! - layout.xMin!) : formatChartNumber(value, xAxis.format),
      font: chartFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(xAxis.label) }, 10), color: xAxis.label?.color ?? color, align: horizontal ? 'right' : 'center', baseline: 'middle', rotation: xAxis.label?.rotate })
    if (xAxis.tick?.visible) {
      const length = (xAxis.tick.length ?? 4) * (xAxis.tick.position === 'inside' ? -1 : 1)
      chartLine(commands, horizontal ? { x: plot.x, y: coordinate } : { x: coordinate, y: plot.y + plot.height },
        horizontal ? { x: plot.x - length, y: coordinate } : { x: coordinate, y: plot.y + plot.height + length }, xAxis.tick, axisColor)
    }
  } else categories.forEach((category, index) => {
    const ratio = xAxis.reverse ? 1 - (index + 0.5) / Math.max(1, categories.length) : (index + 0.5) / Math.max(1, categories.length)
    const coordinate = horizontal ? plot.y + ratio * plot.height : plot.x + ratio * plot.width
    if (xAxis.gridLine?.visible) chartLine(commands, horizontal ? { x: plot.x, y: coordinate } : { x: coordinate, y: plot.y },
      horizontal ? { x: plot.x + plot.width, y: coordinate } : { x: coordinate, y: plot.y + plot.height }, xAxis.gridLine, gridColor, plot)
    if (xAxis.tick?.visible) {
      const length = (xAxis.tick.length ?? 4) * (xAxis.tick.position === 'inside' ? -1 : 1)
      const start = horizontal ? { x: plot.x + zero * plot.width, y: coordinate } : { x: coordinate, y: plot.y + (1 - zero) * plot.height }
      chartLine(commands, start, { x: start.x - (horizontal ? length : 0), y: start.y + (horizontal ? 0 : length) }, xAxis.tick, axisColor)
    }
    if (xAxis.label?.visible !== false) commands.push({ kind: 'text', x: horizontal ? plot.x - 8 : plot.x + ratio * plot.width, y: horizontal ? plot.y + ratio * plot.height : plot.y + plot.height + 14,
      text: category, font: chartFont({ ...chart.options?.appearance?.font, ...xAxis.label }, 10), color: xAxis.label?.color ?? color, align: horizontal ? 'right' : 'center', baseline: 'middle', rotation: xAxis.label?.rotate,
      maxWidth: horizontal ? Math.max(1, plot.x - layout.frame.x - 12) : Math.max(1, plot.width / Math.max(1, categories.length) - 4) })
  })
  if (chart.type === 'heatmap' && yAxis.label?.visible !== false) data.series.forEach((series, index) => commands.push({ kind: 'text', x: plot.x - 8, y: plot.y + (yAxis.reverse ? data.series.length - index - 0.5 : index + 0.5) / Math.max(1, data.series.length) * plot.height, text: chartSeriesStyle(chart, series.field, index, undefined, series.sourceName ?? series.name).name ?? series.name, font: chartFont(yAxis.label, 10), color: yAxis.label?.color ?? color, align: 'right', baseline: 'middle' }))
  if (layout.rightMin !== undefined && layout.rightMax !== undefined) {
    const axis = chart.options?.rightYAxis ?? { scaleType: 'linear' }
    chartLine(commands, horizontal ? { x: plot.x, y: plot.y } : { x: plot.x + plot.width, y: plot.y },
      { x: plot.x + plot.width, y: horizontal ? plot.y : plot.y + plot.height }, axis.line, axisColor)
    for (const value of chartAxisTicks(layout.rightMin, layout.rightMax, axis)) {
      const ratio = chartAxisRatio(value, layout.rightMin, layout.rightMax, axis)
      const point = horizontal ? { x: plot.x + ratio * plot.width, y: plot.y } : { x: plot.x + plot.width, y: plot.y + (1 - ratio) * plot.height }
      if (axis.gridLine?.visible) chartLine(commands, point, horizontal ? { x: point.x, y: plot.y + plot.height } : { x: plot.x, y: point.y }, axis.gridLine, gridColor, plot)
      if (axis.tick?.visible) {
        const length = (axis.tick.length ?? 4) * (axis.tick.position === 'inside' ? -1 : 1)
        chartLine(commands, point, { x: point.x + (horizontal ? 0 : length), y: point.y - (horizontal ? length : 0) }, axis.tick, axisColor)
      }
      if (axis.label?.visible !== false) commands.push({ kind: 'text', x: point.x + (horizontal ? 0 : 8), y: point.y - (horizontal ? 12 : 0),
        text: `${formatChartNumber(value, axis.format)}${chart.type === 'pareto' && !axis.format ? '%' : ''}`, font: chartFont({ ...chart.options?.appearance?.font, ...axis.label }, 10), color: axis.label?.color ?? color,
        align: horizontal ? 'center' : 'left', baseline: 'middle', rotation: axis.label?.rotate })
    }
    const titleOffset = chartAxisTitleSpace(chart, axis) / 2
    const title = chartAxisTitleCommand(chart, axis,
      horizontal ? { x: plot.x, y: plot.y - 40 - titleOffset } : { x: plot.x + plot.width + 54 + titleOffset, y: plot.y + plot.height },
      horizontal ? { x: plot.x + plot.width, y: plot.y - 40 - titleOffset } : { x: plot.x + plot.width + 54 + titleOffset, y: plot.y },
      color, horizontal ? 0 : 90, layout.frame)
    if (title) commands.push(title)
  }
  const bottomAxis = horizontal ? yAxis : xAxis, sideAxis = horizontal ? xAxis : yAxis
  const titleY = plot.y + plot.height + 30 + (chartAxisTitleFont(chart, bottomAxis).fontSize ?? 11) / 2
  const titleX = plot.x - 54 - chartAxisTitleSpace(chart, sideAxis) / 2
  const bottomTitle = chartAxisTitleCommand(chart, bottomAxis, { x: plot.x, y: titleY }, { x: plot.x + plot.width, y: titleY }, color, 0, layout.frame)
  const sideTitle = chartAxisTitleCommand(chart, sideAxis, { x: titleX, y: horizontal ? plot.y : plot.y + plot.height },
    { x: titleX, y: horizontal ? plot.y + plot.height : plot.y }, color, -90, layout.frame)
  if (bottomTitle) commands.push(bottomTitle)
  if (sideTitle) commands.push(sideTitle)
}

function drawCartesian(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const { plot } = layout, horizontal = layout.horizontal, categoryStep = layout.categoryBandwidth ?? (horizontal ? plot.height : plot.width) / Math.max(1, layout.categories.length)
  const bars = layout.series.filter(series => series.type === 'column' || series.type === 'bar')
  const barGroups = [...new Set(bars.map(series => series.stacked ? `stack:${series.rightAxis ? 'right' : 'left'}` : `field:${series.field}`))]
  const gap = chart.type === 'histogram' ? chart.options?.histogram?.binGap ?? 0.05 : (chart.options?.bar?.gapWidth ?? 150) / 100
  const groupWidth = categoryStep * (chart.type === 'histogram' ? 1 - gap : 1 / (1 + gap)), overlap = (chart.options?.bar?.overlap ?? 0) / 100
  const barWidth = groupWidth / Math.max(1, barGroups.length - (barGroups.length - 1) * overlap)
  const maxBubbleSize = data.series.reduce((max, source) => source.sizeValues?.reduce<number>((maximum, size) => Math.max(maximum, size ?? 0), max) ?? max, 1)
  layout.series.slice(0, data.series.length).forEach((series, seriesIndex) => {
    const source = data.series[seriesIndex]!, field = source.field, style = chartSeriesStyle(chart, field, seriesIndex, palette, source.sourceName ?? source.name)
    const isBar = series.type === 'column' || series.type === 'bar' || chart.type === 'histogram' || chart.type === 'pareto'
    if (isBar) {
      const group = series.stacked ? `stack:${series.rightAxis ? 'right' : 'left'}` : `field:${field}`, groupIndex = Math.max(0, barGroups.indexOf(group))
      for (const point of series.points) {
        const pointStyle = chartPointStyle(style, point.categoryIndex), color = chartColorOpacity(pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1)
        const center = horizontal ? point.y : point.x, offset = center - groupWidth / 2 + groupIndex * barWidth * (1 - overlap)
        const rect = horizontal ? { x: Math.min(point.x, point.baseX!), y: offset, width: Math.abs(point.x - point.baseX!), height: Math.max(1, barWidth - 1) }
          : { x: offset, y: Math.min(point.y, point.baseY!), width: Math.max(1, barWidth - 1), height: Math.abs(point.y - point.baseY!) }
        chartRect(commands, rect, color, pointStyle.border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1))
      }
      return
    }
    if (chart.type === 'scatter' || chart.type === 'bubble') {
      for (const point of series.points) {
        const pointStyle = chartPointStyle(style, point.categoryIndex)
        const size = chart.type === 'bubble' ? chartBubblePointSize(pointStyle, point.size ?? 0, maxBubbleSize, chart.options, Math.min(plot.width, plot.height)) : pointStyle.point?.size ?? 8
        if (chart.type === 'bubble' && (point.size === undefined || point.size <= 0)) continue
        chartSymbol(commands, point, size, chartColorOpacity(pointStyle.point?.color ?? pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1), pointStyle.point?.shape ?? 'circle', pointStyle.border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1))
      }
      return
    }
    for (let index = 1; index < series.points.length; index++) {
      const previous = series.points[index - 1]!, point = series.points[index]!
      if (point.categoryIndex !== previous.categoryIndex + 1 && chart.options?.invalidValueStrategy !== 'link') continue
      const styleAtPoint = chartPointStyle(style, point.categoryIndex), color = styleAtPoint.color ?? series.color
      const path = series.type === 'area' && chart.options?.area?.smooth ? smoothChartSegment(series.points, index) : [previous, point]
      for (let step = 1; step < path.length; step++) {
        const a = path[step - 1]!, b = path[step]!
        if (series.type === 'area') {
          const t0 = (step - 1) / (path.length - 1), t1 = step / (path.length - 1)
          const base0 = previous.baseY! + (point.baseY! - previous.baseY!) * t0, base1 = previous.baseY! + (point.baseY! - previous.baseY!) * t1
          const fill = chartColorOpacity(color, styleAtPoint.fillOpacity ?? 0.28), pattern = chartFillPatternPaint(styleAtPoint.fillPattern, styleAtPoint.fillOpacity ?? 0.28)
          if ((a.y - base0) * (b.y - base1) < 0) {
            const ratio = Math.abs(a.y - base0) / (Math.abs(a.y - base0) + Math.abs(b.y - base1)), crossing = { x: a.x + (b.x - a.x) * ratio, y: base0 + (base1 - base0) * ratio }
            chartPolygon(commands, [{ x: a.x, y: base0 }, a, crossing], fill, undefined, plot, pattern)
            chartPolygon(commands, [crossing, b, { x: b.x, y: base1 }], fill, undefined, plot, pattern)
          } else chartPolygon(commands, [{ x: a.x, y: base0 }, a, b, { x: b.x, y: base1 }], fill, undefined, plot, pattern)
        }
        chartLine(commands, a, b, { ...styleAtPoint.border, color: styleAtPoint.border?.color ?? color, width: styleAtPoint.border?.width ?? 2 }, color, plot)
      }
    }
    series.points.forEach(point => {
      const pointStyle = chartPointStyle(style, point.categoryIndex)
      chartSymbol(commands, point, pointStyle.point?.size ?? 6, chartColorOpacity(pointStyle.point?.color ?? pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1), pointStyle.point?.shape ?? 'circle', pointStyle.border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1))
    })
  })
}

function smoothChartSegment(points: ChartPoint[], index: number): ChartVector[] {
  const a = points[index - 1]!, b = points[index]!, previous = points[Math.max(0, index - 2)]!, next = points[Math.min(points.length - 1, index + 1)]!
  return chartBezier(a, { x: a.x + (b.x - previous.x) / 6, y: a.y + (b.y - previous.y) / 6 }, { x: b.x - (next.x - a.x) / 6, y: b.y - (next.y - a.y) / 6 }, b, 12)
}

function drawPointLabels(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  for (const [index, series] of layout.series.entries()) {
    const source = data.series[index]
    const style = chartSeriesStyle(chart, source?.field ?? series.field ?? index, index, palette, source?.sourceName ?? source?.name)
    if (!style.label?.visible) continue
    series.points.forEach(point => {
      const label = chartPointStyle(style, point.categoryIndex).label, inside = label?.position === 'inside' || label?.position === 'center'
      const x = inside && layout.horizontal ? (point.x + (point.baseX ?? point.x)) / 2 : point.x + (label?.position === 'left' ? -9 : label?.position === 'right' ? 9 : 0)
      const y = inside && !layout.horizontal ? (point.y + (point.baseY ?? point.y)) / 2 : point.y + (inside || label?.position === 'left' || label?.position === 'right' ? 0 : label?.position === 'bottom' ? 10 : -9)
      commands.push({ kind: 'text', x, y, text: chartLabelText(label, point.value, point.category, series.name), font: chartFont({ ...chart.options?.appearance?.font, ...label }), color: label?.color ?? color,
        align: label?.position === 'left' ? 'right' : label?.position === 'right' ? 'left' : 'center', baseline: 'middle', clip: layout.plot })
    })
  }
}

function drawTrendlinesAndErrors(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const { plot } = layout
  const position = (x: number, y: number, seriesPosition: number): ChartVector => {
    const source = derived.data.series[seriesPosition]
    if (!source) throw new RangeError('统计图元对应的系列位置已经失效。')
    const style = chartSeriesStyle(chart, source.field, seriesPosition, palette, source.sourceName ?? source.name)
    const axis = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const min = style.rightAxis ? layout.rightMin! : layout.min, max = style.rightAxis ? layout.rightMax! : layout.max
    const xRatio = layout.xMin === undefined ? (x - 0.5) / Math.max(1, layout.categories.length) : chartAxisRatio(x, layout.xMin, layout.xMax!, chart.xAxis)
    const ratio = chartAxisRatio(y, min, max, axis)
    const categoryRatio = layout.xMin === undefined && chart.xAxis.reverse ? 1 - xRatio : xRatio
    return layout.horizontal ? { x: plot.x + ratio * plot.width, y: plot.y + categoryRatio * plot.height }
      : { x: plot.x + categoryRatio * plot.width, y: plot.y + (1 - ratio) * plot.height }
  }
  let annotation = 0
  for (const trend of derived.trends) {
    if (trend.config.visible === false) continue
    const points = trend.statistics.points.map(([x, y]) => position(x, y, trend.seriesPosition))
    for (let index = 1; index < points.length; index++) chartLine(commands, points[index - 1]!, points[index]!, { color: trend.config.color ?? '#7c3aed', width: trend.config.width ?? 1.5, ...trend.config }, '#7c3aed', plot)
    const text = [trend.config.showFormula ? trend.statistics.formula : undefined, trend.config.showR2 && trend.statistics.r2 !== undefined ? `R² = ${formatChartNumber(trend.statistics.r2)}` : undefined].filter(Boolean).join('   ')
    if (text) commands.push({ kind: 'text', x: plot.x + 6, y: plot.y + 12 + annotation++ * 14, text, font: chartFont({}, 10), color: trend.config.color ?? color, baseline: 'middle', maxWidth: plot.width - 12, clip: plot })
  }
  for (const error of derived.errors) {
    if (error.config.line?.visible === false) continue
    const series = layout.series[error.seriesPosition]
    for (const interval of error.intervals) {
      const point = series?.stacked ? series.points.find(point => point.categoryIndex === interval.categoryIndex) : undefined
      const value = interval.y + (series?.stacked ? point?.baseValue ?? 0 : 0)
      const x = layout.xMin === undefined ? interval.categoryIndex + 1 : interval.x
      const start = position(error.config.direction === 'x' ? x - interval.minus : x, error.config.direction === 'y' ? value - interval.minus : value, error.seriesPosition)
      const end = position(error.config.direction === 'x' ? x + interval.plus : x, error.config.direction === 'y' ? value + interval.plus : value, error.seriesPosition)
      chartLine(commands, start, end, error.config.line, '#334155', plot)
      if (!error.config.noEndCap) {
        const vertical = (error.config.direction === 'y') !== Boolean(layout.horizontal)
        for (const anchor of [start, end]) chartLine(commands, vertical ? { x: anchor.x - 4, y: anchor.y } : { x: anchor.x, y: anchor.y - 4 }, vertical ? { x: anchor.x + 4, y: anchor.y } : { x: anchor.x, y: anchor.y + 4 }, error.config.line, '#334155', plot)
      }
    }
  }
}

function drawRadar(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[], theme: ChartRenderTheme): void {
  const { plot } = layout, count = data.categories.length, geometry = chartRadarGeometry(data, chart, plot), { center, radius, vertex } = geometry
  const axisLabel = chartRadarAxisLabel(chart), axisLine = chartRadarAxisLine(chart), tick = chart.yAxis.tick
  for (const value of geometry.ticks) {
    const ratio = chartAxisRatio(value, geometry.min, geometry.max, chart.yAxis)
    if (ratio <= 0) continue
    if (chart.options?.radar?.shape === 'circle') {
      const ring = chartArc(center, radius * ratio, 0, Math.PI * 2)
      for (let index = 1; index < ring.length; index++) chartLine(commands, ring[index - 1]!, ring[index]!, chart.yAxis.gridLine, theme.splitLineColor ?? '#e2e8f0')
    } else data.categories.forEach((_category, index) => chartLine(commands, vertex(index, ratio), vertex((index + 1) % count, ratio), chart.yAxis.gridLine, theme.splitLineColor ?? '#e2e8f0'))
  }
  data.categories.forEach((category, index) => {
    chartLine(commands, center, vertex(index, 1), axisLine, theme.axisColor ?? '#cbd5e1')
    const angle = geometry.angle(index), perpendicular = { x: -Math.sin(angle), y: Math.cos(angle) }
    for (const value of geometry.ticks) {
      const ratio = chartAxisRatio(value, geometry.min, geometry.max, chart.yAxis)
      if (ratio <= 0) continue
      const point = vertex(index, ratio), length = tick?.length ?? 4, direction = tick?.position === 'inside' ? -1 : 1
      if (tick?.visible) chartLine(commands, point, { x: point.x + perpendicular.x * length * direction, y: point.y + perpendicular.y * length * direction }, tick, theme.axisColor ?? '#cbd5e1')
      if (axisLabel.visible) {
        const distance = 6 + (tick?.visible && direction > 0 ? length : 0)
        commands.push({ kind: 'text', x: point.x + perpendicular.x * distance, y: point.y + perpendicular.y * distance,
          text: formatChartNumber(value, axisLabel.format), font: chartFont(axisLabel, 10), color: axisLabel.color ?? color,
          align: perpendicular.x > 1e-8 ? 'left' : perpendicular.x < -1e-8 ? 'right' : 'center', baseline: 'middle', rotation: chart.yAxis.label?.rotate, clip: layout.frame })
      }
    }
    if (chart.xAxis.label?.visible !== false) commands.push({ kind: 'text', ...vertex(index, 1.22), text: category,
      font: chartFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(chart.xAxis.label) }, 10), color: chart.xAxis.label?.color ?? color,
      align: 'center', baseline: 'middle', rotation: chart.xAxis.label?.rotate, clip: layout.frame })
  })
  for (const [index, series] of layout.series.entries()) {
    const source = data.series[index]!, style = chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), points = new Map(series.points.map(point => [point.categoryIndex, point]))
    const polygon = data.categories.map((_category, categoryIndex) => points.get(categoryIndex) ?? center)
    if (chart.options?.radar?.fill) chartPolygon(commands, polygon, chartColorOpacity(style.color ?? series.color, style.fillOpacity ?? 0.25), undefined, plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 0.25))
    polygon.forEach((point, categoryIndex) => chartLine(commands, point, polygon[(categoryIndex + 1) % polygon.length]!, { color: series.color, width: 2, ...style.border }, series.color))
    series.points.forEach(point => {
      const pointStyle = chartPointStyle(style, point.categoryIndex)
      chartSymbol(commands, point, pointStyle.point?.size ?? 4, chartColorOpacity(pointStyle.point?.color ?? pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1),
        pointStyle.point?.shape ?? 'none', pointStyle.border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1))
    })
  }
  const titleY = plot.y + plot.height + 24 + Math.max(0, (chartAxisTitleFont(chart, chart.xAxis).fontSize ?? 11) - 11) / 2
  const titleX = plot.x - 54 - chartAxisTitleSpace(chart, chart.yAxis) / 2
  const xTitle = chartAxisTitleCommand(chart, chart.xAxis, { x: plot.x, y: titleY }, { x: plot.x + plot.width, y: titleY }, color, 0, layout.frame)
  const yTitle = chartAxisTitleCommand(chart, chart.yAxis, { x: titleX, y: plot.y + plot.height }, { x: titleX, y: plot.y }, color, -90, layout.frame)
  if (xTitle) commands.push(xTitle)
  if (yTitle) commands.push(yTitle)
}

export function drawChartSector(commands: DrawCommand[], center: ChartVector, inner: number, outer: number, start: number, end: number, color: string, border?: ChartSeriesOptions['border'], clip?: RenderRect, pattern?: ChartSeriesFillPattern): void {
  const count = Math.max(2, Math.ceil(Math.abs(end - start) / (Math.PI / 48)))
  for (let index = 0; index < count; index++) {
    const a = start + (end - start) * index / count, b = start + (end - start) * (index + 1) / count
    chartPolygon(commands, inner > 0 ? [chartPolar(center, inner, a), chartPolar(center, outer, a), chartPolar(center, outer, b), chartPolar(center, inner, b)] : [center, chartPolar(center, outer, a), chartPolar(center, outer, b)], color, undefined, clip, pattern)
  }
  if (border && border.visible !== false) {
    const outerArc = chartArc(center, outer, start, end), innerArc = inner > 0 ? chartArc(center, inner, start, end) : [center]
    for (const arc of [outerArc, innerArc]) for (let index = 1; index < arc.length; index++) chartLine(commands, arc[index - 1]!, arc[index]!, border, color, clip)
    chartLine(commands, outerArc[0]!, innerArc[0]!, border, color, clip); chartLine(commands, outerArc.at(-1)!, innerArc.at(-1)!, border, color, clip)
  }
}

function drawPie(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const pie = chartPieLayout(data, chart, layout.plot), field = data.series[0]!.field
  for (const [start, end] of pie.connectors) chartLine(commands, start, end, { color, width: 1 }, color)
  for (const sector of pie.sectors) {
    const style = chartPointStyle(chartSeriesStyle(chart, field, sector.part.index, palette, data.series[0]!.sourceName ?? data.series[0]!.name), sector.part.index)
    drawChartSector(commands, sector.center, sector.inner, sector.outer, sector.start, sector.end, chartColorOpacity(style.color ?? palette[sector.part.index % palette.length]!, style.fillOpacity ?? 1), style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    const label = { ...style.label, ...chart.options?.pie?.label }
    if (label.visible !== false) {
      const outside = label.position !== 'inside' && label.position !== 'center', point = chartPolar(sector.center, outside ? sector.outer * 1.12 : (sector.inner + sector.outer) / 2, (sector.start + sector.end) / 2)
      if (outside && chart.options?.pie?.showLabelLine !== false) chartLine(commands, chartPolar(sector.center, sector.outer, (sector.start + sector.end) / 2), point, { color, width: 0.8 }, color)
      commands.push({ kind: 'text', ...point, text: chartLabelText({ contentType: 'category', ...label }, sector.part.value, sector.part.category, data.series[0]!.name, sector.percentage), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot })
    }
  }
  for (const bar of pie.bars) {
    const style = chartPointStyle(chartSeriesStyle(chart, field, bar.part.index, palette, data.series[0]!.sourceName ?? data.series[0]!.name), bar.part.index)
    chartRect(commands, bar.rect, chartColorOpacity(style.color ?? palette[bar.part.index % palette.length]!, style.fillOpacity ?? 1), style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    const label = { ...style.label, ...chart.options?.pie?.label }
    if (label.visible !== false) commands.push({ kind: 'text', x: bar.rect.x + bar.rect.width / 2, y: bar.rect.y + bar.rect.height / 2, text: chartLabelText({ contentType: 'category', ...label }, bar.part.value, bar.part.category, data.series[0]!.name, bar.percentage), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', maxWidth: bar.rect.width, clip: layout.plot })
  }
}

function drawBoxplot(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, textColor: string, palette: string[]): void {
  const { plot } = layout, boxes = derived.boxes!
  boxes.forEach((box, index) => {
    if (!box.statistics) return
    const style = chartBoxplotStyle(chart, box, index, palette), statistics = box.statistics, color = style.color ?? palette[box.seriesPosition % palette.length]!
    const geometry = chartBoxplotGeometry(chart, layout, style, index, boxes.length), { x, width } = geometry, summary = geometry.summary(statistics)
    if (summary) {
      const top = Math.min(summary.q1, summary.q3), bottom = Math.max(summary.q1, summary.q3), border = { color, width: 1.5, ...style.border }
      chartRect(commands, { x: x - width / 2, y: top, width, height: bottom - top }, chartColorOpacity(color, style.fillOpacity ?? 0.45), undefined, plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 0.45))
      chartLine(commands, { x, y: summary.min }, { x, y: summary.q1 }, border, color, plot)
      chartLine(commands, { x, y: summary.q3 }, { x, y: summary.max }, border, color, plot)
      for (const y of [summary.min, summary.max]) chartLine(commands, { x: x - width / 2, y }, { x: x + width / 2, y }, border, color, plot)
      const rectangle = [{ x: x - width / 2, y: top }, { x: x + width / 2, y: top }, { x: x + width / 2, y: bottom }, { x: x - width / 2, y: bottom }]
      rectangle.forEach((point, position) => chartLine(commands, point, rectangle[(position + 1) % rectangle.length]!, border, color, plot))
      chartLine(commands, { x: x - width / 2, y: summary.median }, { x: x + width / 2, y: summary.median }, border, color, plot)
      if (style.label?.visible) {
        const label = style.label, inside = label.position === 'inside' || label.position === 'center', side = label.position === 'left' || label.position === 'right'
        const labelX = x + (label.position === 'left' ? -width / 2 - 9 : label.position === 'right' ? width / 2 + 9 : 0)
        const labelY = inside ? (top + bottom) / 2 : side ? summary.median : label.position === 'bottom' ? bottom + 10 : top - 9
        commands.push({ kind: 'text', x: labelX, y: labelY, text: chartLabelText(label, statistics.median, box.name, style.name ?? box.name),
          font: chartFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(label) }), color: label.color ?? textColor,
          align: label.position === 'left' ? 'right' : label.position === 'right' ? 'left' : 'center', baseline: 'middle', clip: plot })
      }
    }
    if (chart.options?.boxplot?.showOutliers !== false) statistics.outliers.forEach(value => {
      const y = geometry.y(value), size = style.point?.size ?? 4, shape = style.point?.shape ?? 'circle'
      if (y === undefined || size <= 0 || shape === 'none') return
      const markerColor = style.point?.color ?? color, fill = chartColorOpacity(markerColor, style.fillOpacity ?? 1), border = { color: markerColor, width: 1.5, ...style.border }
      if (shape === 'rect') {
        const half = size / 2, corners = [{ x: x - half, y: y - half }, { x: x + half, y: y - half }, { x: x + half, y: y + half }, { x: x - half, y: y + half }]
        chartRect(commands, { x: x - half, y: y - half, width: size, height: size }, fill, undefined, plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
        corners.forEach((point, position) => chartLine(commands, point, corners[(position + 1) % corners.length]!, border, markerColor, plot))
      } else chartSymbol(commands, { x, y }, size, fill, shape, border, plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    })
  })
}

function drawCandlestick(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const { plot } = layout, step = layout.categoryBandwidth ?? plot.width / Math.max(1, data.categories.length), width = Math.min(step * 0.8, chart.options?.candlestick?.candleWidth ?? step * 0.5)
  const y = (value: number): number => plot.y + (1 - chartAxisRatio(value, layout.min, layout.max, chart.yAxis)) * plot.height
  for (const candle of data.candles ?? []) {
    if (!candle) continue
    const index = candle.categoryIndex, ratio = layout.categoryPositions ? layout.categoryPositions[index] : chart.xAxis.reverse ? 1 - (index + 0.5) / Math.max(1, data.categories.length) : (index + 0.5) / Math.max(1, data.categories.length)
    if (ratio == null) continue
    const x = plot.x + ratio * plot.width
    const rising = candle.open === undefined ? true : candle.close >= candle.open, style = chartCandlestickPointStyle(chart, data.series[0]!.field, index, rising, palette, data.series[0]!.sourceName ?? data.series[0]!.name)
    const color = chartColorOpacity(style.color!, style.fillOpacity ?? 1), border = { color: style.color, ...style.border }
    chartLine(commands, { x, y: y(candle.high) }, { x, y: y(candle.low) }, { width: 1.5, ...border }, color, plot)
    if (candle.open === undefined) chartLine(commands, { x, y: y(candle.close) }, { x: x + width / 2, y: y(candle.close) }, { width: 2, ...border }, color, plot)
    else chartRect(commands, { x: x - width / 2, y: Math.min(y(candle.open), y(candle.close)), width, height: Math.max(1, Math.abs(y(candle.close) - y(candle.open))) },
      chart.options?.candlestick?.hollow && rising ? 'rgba(0, 0, 0, 0)' : color, { ...border, width: border.width ?? 1 }, plot,
      chart.options?.candlestick?.hollow && rising ? undefined : chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
  }
}

function drawWaterfall(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const { plot } = layout, step = plot.width / Math.max(1, layout.categories.length), groupWidth = step / (1 + (chart.options?.bar?.gapWidth ?? 150) / 100), overlap = (chart.options?.bar?.overlap ?? 0) / 100
  const width = groupWidth / Math.max(1, layout.series.length - (layout.series.length - 1) * overlap), options = chart.options?.waterfall
  layout.series.forEach((series, seriesIndex) => series.points.forEach((point, index) => {
    const source = derived.data.series[seriesIndex]!, role = derived.waterfalls![seriesIndex]![index]!.role
    const style = chartPointStyle(chartSeriesStyle(chart, source.field, seriesIndex, palette, source.sourceName ?? source.name), point.categoryIndex)
    const roleColor = role === 'total' ? options?.totalColor ?? palette[seriesIndex % palette.length]! : point.value >= 0 ? options?.positiveColor ?? '#059669' : options?.negativeColor ?? '#dc2626'
    const color = style.dataPoints?.[point.categoryIndex]?.color ?? (chartSeriesColorIsAuthored(chart, source.field, source.sourceName ?? source.name) ? style.color! : roleColor)
    const x = point.x - groupWidth / 2 + seriesIndex * width * (1 - overlap)
    chartRect(commands, { x, y: Math.min(point.y, point.baseY!), width: Math.max(1, width - 1), height: Math.abs(point.y - point.baseY!) }, chartColorOpacity(color, style.fillOpacity ?? 1), style.border, plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    const next = series.points[index + 1]
    if (next && options?.connector !== false) chartLine(commands, { x: x + (chart.xAxis.reverse ? 0 : width), y: point.y }, { x: next.x - groupWidth / 2 + seriesIndex * width * (1 - overlap) + (chart.xAxis.reverse ? width : 0), y: point.y }, { color: '#94a3b8', dashType: 'dashed', width: 1 }, '#94a3b8', plot)
  }))
}

function drawHeatmap(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const extent = chartHeatmapExtent(data, chart)
  const width = layout.plot.width / Math.max(1, data.categories.length), height = layout.plot.height / Math.max(1, data.series.length)
  data.series.forEach((series, row) => series.values.forEach((value, col) => {
    if (value === null) return
    const style = chartHeatmapPointStyle(chart, series.field, row, col, value, extent, palette, series.sourceName ?? series.name)
    const fill = chartColorOpacity(style.color!, style.fillOpacity ?? 1)
    const x = layout.plot.x + (chart.xAxis.reverse ? data.categories.length - col - 1 : col) * width, y = layout.plot.y + (chart.yAxis.reverse ? data.series.length - row - 1 : row) * height
    chartRect(commands, { x, y, width, height }, fill, style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    const label = { ...style.label, ...chart.options?.heatmap?.label }
    if (label.visible) commands.push({ kind: 'text', x: x + width / 2, y: y + height / 2, text: `${chartLabelText(label, value, data.categories[col]!, series.name)}${chart.options?.heatmap?.unit ?? ''}`, font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', maxWidth: Math.max(1, width - 4), clip: layout.plot })
  }))
}

function drawGraph(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const sankey = chart.type === 'sankey', graph = sankey ? chartSankeyLayout(data.graph!, layout.plot) : chartRelationLayout(data.graph!, chart, layout.plot)
  for (const link of graph.links) {
    const style = chartPointStyle(chartSeriesStyle(chart, data.series[0]!.field, link.colorIndex, palette, data.series[0]!.sourceName ?? data.series[0]!.name), link.colorIndex)
    const fill = chartColorOpacity(style.color!, (sankey ? 0.55 : 0.6) * (style.fillOpacity ?? 1))
    const pattern = chartFillPatternPaint(style.fillPattern, (sankey ? 0.55 : 0.6) * (style.fillOpacity ?? 1))
    chartPolygon(commands, chartRibbon(link.points, link.width), fill, undefined, layout.plot, pattern)
    if (!sankey) {
      const end = link.points.at(-1)!, previous = link.points[Math.max(0, link.points.length - 2)]!, dx = end.x - previous.x, dy = end.y - previous.y, distance = Math.max(1e-9, Math.hypot(dx, dy))
      const target = graph.nodes.find(node => node.name === link.target)!, anchor = { x: end.x - dx / distance * target.width / 2, y: end.y - dy / distance * target.height / 2 }
      chartPolygon(commands, [anchor, { x: anchor.x - dx / distance * 6 - dy / distance * 3, y: anchor.y - dy / distance * 6 + dx / distance * 3 }, { x: anchor.x - dx / distance * 6 + dy / distance * 3, y: anchor.y - dy / distance * 6 - dx / distance * 3 }], fill, undefined, layout.plot, pattern)
    }
  }
  for (const node of graph.nodes) {
    const style = chartPointStyle(chartSeriesStyle(chart, data.series[0]!.field, node.index, palette, data.series[0]!.sourceName ?? data.series[0]!.name), node.index), fill = chartColorOpacity(style.color!, style.fillOpacity ?? 1)
    if (sankey) chartRect(commands, { x: node.x, y: node.y, width: node.width, height: node.height }, fill, style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    else chartSymbol(commands, node, style.point?.size ?? node.width, chartColorOpacity(style.point?.color ?? style.color!, style.fillOpacity ?? 1), style.point?.shape ?? 'circle', style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    if (style.label?.visible !== false) commands.push({ kind: 'text', x: node.x + (sankey ? node.width + 4 : 0), y: node.y + (sankey ? node.height / 2 : node.height / 2 + 10),
      text: chartLabelText({ contentType: 'category', ...style.label }, node.value, node.name, data.series[0]!.name), font: chartFont(style.label, 10), color: style.label?.color ?? color, align: sankey ? 'left' : 'center', baseline: 'middle', clip: layout.plot })
  }
}

function drawChord(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const chord = chartChordLayout(data.graph!, chart, layout.plot)
  for (const link of chord.links) {
    const source = chartArc(chord.center, chord.innerRadius, link.sourceStart, link.sourceEnd), target = chartArc(chord.center, chord.innerRadius, link.targetStart, link.targetEnd)
    const first = chartBezier(source.at(-1)!, chord.center, chord.center, target[0]!), second = chartBezier(target.at(-1)!, chord.center, chord.center, source[0]!)
    const style = chartPointStyle(chartSeriesStyle(chart, data.series[0]!.field, link.sourceIndex, palette, data.series[0]!.sourceName ?? data.series[0]!.name), link.sourceIndex)
    chartPolygon(commands, [...source, ...first.slice(1), ...target.slice(1), ...second.slice(1)], chartColorOpacity(style.color!, (chart.options?.chord?.ribbonOpacity ?? 0.55) * (style.fillOpacity ?? 1)), undefined, layout.plot,
      chartFillPatternPaint(style.fillPattern, (chart.options?.chord?.ribbonOpacity ?? 0.55) * (style.fillOpacity ?? 1)))
  }
  for (const node of chord.nodes) {
    const style = chartPointStyle(chartSeriesStyle(chart, data.series[0]!.field, node.index, palette, data.series[0]!.sourceName ?? data.series[0]!.name), node.index), label = { contentType: 'category' as const, ...style.label, ...chart.options?.chord?.label }
    drawChartSector(commands, chord.center, chord.innerRadius, chord.outerRadius, node.start, node.end, chartColorOpacity(style.color!, style.fillOpacity ?? 1), style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    if (label.visible !== false) commands.push({ kind: 'text', ...chartPolar(chord.center, chord.outerRadius + 10, (node.start + node.end) / 2), text: chartLabelText(label, node.value, node.name, data.series[0]!.name), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot })
  }
}

function drawWordCloud(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const words = chartWordCloudLayout(data, chart.options?.wordCloud, layout.plot)
  for (const word of words) {
    const style = chartPointStyle(chartSeriesStyle(chart, data.series[0]!.field, word.categoryIndex, palette, data.series[0]!.sourceName ?? data.series[0]!.name), word.categoryIndex)
    commands.push({ kind: 'text', x: word.x + word.width / 2, y: word.y + word.height / 2, text: word.text, font: chartFont({ ...chart.options?.appearance?.font, ...style.label, fontSize: word.fontSize }), color: chartColorOpacity(style.color ?? palette[word.categoryIndex % palette.length]!, style.fillOpacity ?? 1), align: 'center', baseline: 'middle', rotation: word.rotation, clip: layout.plot })
  }
}

function drawTreemap(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const options = chart.options?.treemap, banner = options?.parentLabelLayout === undefined || options.parentLabelLayout === 'banner'
  const rectangles = chartTreemapLayout(data.hierarchy!, layout.plot, banner), field = data.series[0]!.field
  const scale = data.hierarchy!.reduce((max, node) => Math.max(max, node.value), 0) || 1, total = data.hierarchy!.reduce((sum, node) => sum + node.value / scale, 0)
  for (const rect of rectangles) {
    const style = chartHierarchyPointStyle(chart, field, rect, palette, data.series[0]!.sourceName ?? data.series[0]!.name)
    chartRect(commands, { x: rect.x, y: rect.y, width: rect.width, height: rect.height }, chartColorOpacity(style.color!, style.fillOpacity ?? 1),
      { ...style.border, color: style.border?.color ?? chart.options?.appearance?.background ?? '#ffffff', width: style.border?.width ?? 1 }, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
  }
  for (const rect of rectangles) {
    const parent = Boolean(rect.node.children?.length), style = chartHierarchyPointStyle(chart, field, rect, palette, data.series[0]!.sourceName ?? data.series[0]!.name), label = { contentType: 'category' as const, ...style.label, ...options?.label }
    if (label.visible === false || parent && options?.parentLabelLayout === 'none') continue
    const y = parent && banner ? rect.y + 10 : rect.y + rect.height / 2
    commands.push({ kind: 'text', x: rect.x + rect.width / 2, y, text: chartLabelText(label, rect.node.value, rect.node.name, data.series[0]!.name, total ? (rect.node.value / scale) / total * 100 : 0), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', maxWidth: Math.max(1, rect.width - 6), clip: rect })
  }
}

function drawSunburst(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const sectors = chartSunburstLayout(data.hierarchy!, layout.plot), center = { x: layout.plot.x + layout.plot.width / 2, y: layout.plot.y + layout.plot.height / 2 }
  const scale = data.hierarchy!.reduce((max, node) => Math.max(max, node.value), 0) || 1, total = data.hierarchy!.reduce((sum, node) => sum + node.value / scale, 0)
  for (const sector of sectors) {
    const style = chartHierarchyPointStyle(chart, data.series[0]!.field, sector, palette, data.series[0]!.sourceName ?? data.series[0]!.name), label = { contentType: 'category' as const, ...style.label, ...chart.options?.sunburst?.label }
    drawChartSector(commands, center, sector.inner, sector.outer, sector.start, sector.end, chartColorOpacity(style.color!, style.fillOpacity ?? 1), { color: chart.options?.appearance?.background ?? '#ffffff', width: 0.7, ...style.border }, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    if (label.visible !== false && sector.end - sector.start > 0.06) commands.push({ kind: 'text', ...chartPolar(center, (sector.inner + sector.outer) / 2, (sector.start + sector.end) / 2), text: chartLabelText(label, sector.node.value, sector.node.name, data.series[0]!.name, total ? (sector.node.value / scale) / total * 100 : 0), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', maxWidth: Math.max(1, (sector.end - sector.start) * (sector.inner + sector.outer) / 2), clip: layout.plot })
  }
  if (chart.options?.sunburst?.centerLabel?.visible) commands.push({ kind: 'text', ...center, text: data.series[0]!.name, font: chartFont(chart.options.sunburst.centerLabel.font, 11), color: chart.options.sunburst.centerLabel.font?.color ?? color, align: 'center', baseline: 'middle', maxWidth: sectors[0] ? sectors[0].inner * 1.5 : layout.plot.width, clip: layout.plot })
}

function drawGauge(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[], theme: ChartRenderTheme): void {
  const authored = chart.options?.gauge, source = data.series[0]!
  for (const tile of chartGaugeTiles(data, layout.plot)) {
    const index = tile.categoryIndex, value = source.values[index]!, name = data.categories[index]!
    const style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index), label = style.label
    const options = resolveChartGaugeOptions(authored, tile.radius, chart.options?.appearance?.font, label), sweep = chartGaugeSweep(options)
    const ratio = chartGaugeRatio(value, options.min, options.max), span = sweep.end - sweep.start, angle = sweep.start + span * ratio
    const pointerColor = chartColorOpacity(options.pointer.color ?? style.color!, style.fillOpacity ?? 1), progressColor = chartColorOpacity(options.progress.color ?? style.color!, style.fillOpacity ?? 1)
    if (options.axisLine.visible && options.axisLine.width > 0 && span > 0) {
      let from = sweep.start
      for (const [end, fill] of chartGaugeColorStops(authored, options.axisLine.color ?? theme.splitLineColor ?? '#e2e8f0')) {
        const to = sweep.start + span * end
        if (to > from) drawChartSector(commands, tile.center, tile.radius - options.axisLine.width, tile.radius, from, to, fill, undefined, layout.plot)
        from = to
      }
    }
    if (options.progress.visible && options.progress.width > 0 && ratio > 0 && span > 0) {
      const width = options.progress.width
      drawChartSector(commands, tile.center, tile.radius - width, tile.radius, sweep.start, angle, progressColor, style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
      if (options.progress.rounded && angle - sweep.start < Math.PI * 2 - 1e-8) {
        for (const mark of [sweep.start, angle]) chartCircle(commands, chartPolar(tile.center, tile.radius - width / 2, mark), width / 2, progressColor, style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
      }
    }
    const innerRadius = tile.radius - options.axisLine.width, axisColor = theme.axisColor ?? color
    for (let index = 0; index <= options.splitNumber; index++) {
      const tickAngle = sweep.start + span * index / options.splitNumber
      if (options.splitLine.visible) chartLine(commands, chartPolar(tile.center, innerRadius, tickAngle), chartPolar(tile.center, innerRadius - options.splitLine.length, tickAngle), options.splitLine, axisColor, layout.plot)
      if (options.axisLabel.visible) commands.push({ kind: 'text', ...chartPolar(tile.center, chartGaugeAxisLabelRadius(options, tile.radius), tickAngle),
        text: formatChartNumber(chartInterpolate(options.min, options.max, index / options.splitNumber), options.axisLabel.format), font: chartFont(options.axisLabel),
        color: options.axisLabel.color ?? axisColor, ...chartGaugeAxisLabelAlignment(tickAngle), clip: layout.plot })
      if (options.axisTick.visible && index < options.splitNumber) for (let minor = 0; minor <= options.axisTick.splitNumber; minor++) {
        const minorAngle = sweep.start + span * (index + minor / options.axisTick.splitNumber) / options.splitNumber
        chartLine(commands, chartPolar(tile.center, innerRadius, minorAngle), chartPolar(tile.center, innerRadius - options.axisTick.length, minorAngle), options.axisTick, axisColor, layout.plot)
      }
    }
    if (options.anchor.visible) chartCircle(commands, tile.center, options.anchor.size / 2,
      options.anchor.color ?? (options.mode === 'pointer' ? pointerColor : chart.options?.appearance?.background ?? theme.backgroundColor ?? '#ffffff'),
      { color: options.anchor.borderColor ?? pointerColor, width: options.anchor.borderWidth }, layout.plot)
    if (options.pointer.visible && options.pointer.width > 0 && options.pointer.length > 0) {
      chartPolygon(commands, chartGaugePointerPoints(tile.center, tile.radius * options.pointer.length, options.pointer.width, angle, options.pointer.shape), pointerColor, style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    }
    if (options.title.visible) commands.push({ kind: 'text', x: tile.center.x, y: tile.center.y + tile.radius * options.title.offsetY, text: name,
      font: chartFont(options.title), color: options.title.color ?? color, align: 'center', baseline: 'middle', maxWidth: tile.radius * 1.7, clip: layout.plot })
    if (options.detail.visible) commands.push({ kind: 'text', x: tile.center.x, y: tile.center.y + tile.radius * options.detail.offsetY,
      text: chartGaugeDetailText(value, name, style.name ?? source.name, { ...label, ...options.detail }, data.records?.[index]?.[source.field]),
      font: chartFont(options.detail), color: options.detail.color ?? color, align: 'center', baseline: 'middle', maxWidth: tile.radius * 1.7, clip: layout.plot })
  }
}

function drawFunnel(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const source = data.series[0]!
  for (const step of chartFunnelLayout(data, chart, layout.plot)) {
    const style = chartPointStyle(chartSeriesStyle(chart, source.field, step.categoryIndex, palette, source.sourceName ?? source.name), step.categoryIndex)
    chartPolygon(commands, step.points, chartColorOpacity(style.color ?? palette[step.categoryIndex % palette.length]!, style.fillOpacity ?? 1), style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    if (style.label?.visible !== false) commands.push({ kind: 'text', ...step.center, text: chartLabelText({ contentType: 'category', ...style.label }, step.value, step.category, source.name), font: chartFont(style.label, 10), color: style.label?.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot })
  }
}

function drawLegend(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  if (!chart.legend.visible) return
  const family = resolveChartType(chart.type).family
  const categoryLegend = ['pie', 'funnel', 'gauge', 'chord', 'relation', 'sankey', 'treemap', 'sunburst'].includes(family)
  const roots = data.hierarchy ? chartHierarchyItems(data.hierarchy).filter(item => item.depth === 0) : undefined
  const items = categoryLegend ? data.categories.flatMap((name, index) => {
    if (family === 'gauge' && !Number.isFinite(data.series[0]?.values[index])) return []
    const style = roots?.[index] ? chartHierarchyPointStyle(chart, data.series[0]!.field, roots[index]!, palette, data.series[0]!.sourceName ?? data.series[0]!.name)
      : chartPointStyle(chartSeriesStyle(chart, data.series[0]!.field, index, palette, data.series[0]!.sourceName ?? data.series[0]!.name), index)
    return [{ name, color: chartColorOpacity(family === 'relation' ? style.point?.color ?? style.color! : style.color!, style.fillOpacity ?? 1), pattern: chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1) }]
  }) : chart.type === 'boxplot' ? [{ name: '样本分布', color: layout.series[0]?.color ?? palette[0]!, pattern: chartFillPatternPaint(chartPointStyle(chartSeriesStyle(chart, layout.series[0]?.field ?? data.series[0]!.field, 0, palette, data.series[0]!.sourceName ?? data.series[0]!.name), 0).fillPattern, 0.45) }, ...(chart.options?.boxplot?.showOutliers === false ? [] : [{ name: '异常值', color: chartSeriesStyle(chart, layout.series[0]?.field ?? data.series[0]!.field, 0, palette, data.series[0]!.sourceName ?? data.series[0]!.name).point?.color ?? layout.series[0]?.color ?? palette[0]!, pattern: undefined }])]
    : layout.series.map((series, index) => {
      const source = data.series[index], style = source ? chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name) : undefined
      return { name: series.name, color: series.color, pattern: chartFillPatternPaint(style?.fillPattern, style?.fillOpacity ?? 1) }
    })
  const font: ChartFontOptions = { ...chart.options?.appearance?.font, ...chart.legend.font }, fontSize = font.fontSize ?? 11
  const horizontal = chart.legend.position === 'top' || chart.legend.position === 'bottom', itemWidth = Math.max(72, Math.min(140, layout.frame.width / Math.max(1, items.length)))
  const perRow = Math.max(1, Math.floor((layout.frame.width - 24) / itemWidth))
  items.forEach((item, index) => {
    const x = horizontal ? layout.frame.x + 12 + index % perRow * itemWidth : chart.legend.position === 'left' ? layout.frame.x + 8 : layout.plot.x + layout.plot.width + 12
    const y = horizontal ? chart.legend.position === 'top' ? layout.plot.y - 14 + Math.floor(index / perRow) * (fontSize + 4) : layout.frame.y + layout.frame.height - 12 - Math.floor(index / perRow) * (fontSize + 4)
      : layout.plot.y + 10 + index * (fontSize + 6)
    chartRect(commands, { x, y: y - 5, width: 10, height: 10 }, item.color, undefined, layout.frame, item.pattern)
    commands.push({ kind: 'text', x: x + 15, y, text: item.name, font: chartFont(font, 11), color: font.color ?? color, baseline: 'middle', maxWidth: horizontal ? itemWidth - 20 : Math.max(1, layout.frame.width * 0.24 - 22), clip: layout.frame })
  })
}
