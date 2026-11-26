import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartLineOptions, ChartModel, ChartSeriesOptions } from '../../domain/model'
import type { DrawCommand, RenderRect, TextCommand } from '../../rendering/contracts'
import { ComputeError } from '../../compute/errors'
import type { ChartData, ChartLayout, ChartPoint } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { deriveChartData, chartBubblePointSize, chartPointStyle, type ChartComputedStatistics, type ChartDerivedData } from './chart-derived'
import { isChartLinearGradient, normalizeChartColor } from './chart-color'
import { chartInterpolate, chartLinearRatio } from './chart-math'
import { layoutDerivedChartData } from './chart-layout'
import {
  chartArc, chartAxisRatio, chartAxisTicks, chartCircle, chartColorOpacity, chartFilledMarkBorder, chartFont, chartLabelText, chartLine,
  chartPolar, chartPolygon, chartRect, chartSymbol, formatChartNumber, type ChartVector,
} from './chart-geometry'
import { chartPalette, chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { chart3DScene } from './chart-3d'
import { chartFunnelLayout, chartGaugeTiles, chartPieLayout } from './chart-pie-layout'
import { chartPieAggregateStyle, chartPieConnector, chartPieSliceStyle, isChartCompoundPie } from './chart-pie'
import { chartGaugeAxisLabelAlignment, chartGaugeAxisLabelRadius, chartGaugeColorStops, chartGaugeDetailText, chartGaugePointerPoints, chartGaugeRatio, chartGaugeSweep, resolveChartGaugeOptions } from './chart-gauge'
import { chartRadarAxisLabel, chartRadarAxisLine, chartRadarGeometry } from './chart-radar'
import { chartBoxplotGeometry, chartBoxplotOutliers, chartBoxplotOutlierStyle, chartBoxplotStyle } from './chart-boxplot'
import { chartRectangularAxisTitleCommands } from './chart-axis-title'
import { chartDateCoordinate } from './chart-date-axis'
import { chartDateAxisCommands } from './chart-date-axis-draw'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartSeriesGradientBounds, createChartMarkGradient } from './chart-gradient'
import { CHART_PARETO_CUMULATIVE_NAME, chartParetoBarStyle, chartParetoCumulativeMarkerStyle, chartParetoCumulativePointStyle, chartParetoCumulativeStyle, chartParetoValuePercentages } from './chart-pareto'
import { chartPointLabelAnchor } from './chart-labels'
import { createChartSankeyScene, type ChartSankeyScene } from './chart-sankey-draw'
import { CHART_WATERFALL_GLOBAL_SERIES } from './chart-statistics'
import { chartWaterfallConnector, chartWaterfallRoleStyles } from './chart-waterfall'
import type { ChartRenderTheme } from './chart-theme'
import { chartHeatmapExtent, chartHeatmapLegendBand, createChartHeatmapScale, type ChartHeatmapScale } from './chart-heatmap'
import { createChartHeatmapLegend, drawChartHeatmapAxes, drawChartHeatmapCells } from './chart-heatmap-draw'
import { chartMarkerShape } from './chart-markers'
import { chartHistogramBinGap } from './chart-histogram'
import { drawChartCandlesticks } from './chart-candlestick-draw'
import { drawChartTreemap } from './chart-treemap-draw'
import { drawChartSunburst } from './chart-sunburst-draw'
import { drawChartSector } from './chart-sector-draw'
import { createChartChordScene, type ChartChordScene } from './chart-chord-draw'
import { createChartRelationScene, type ChartRelationScene } from './chart-relation-draw'
import { chartHasLegend, chartLegendDataCount, chartLegendGeometry, chartLegendItems } from './chart-legend'
import { chartTitleCommands, chartTitleGeometry } from './chart-titles'
import { resolveChartAreaLineStyle } from './chart-line-and-area'
import { chartLinePointsConnected, chartLineSegment } from './chart-line-path'
import { chartCanvasBoxplotBindings, chartCanvasParetoMark, chartCanvasPieSliceMark, chartCanvasSeriesMark, chartCanvasSeriesBindings, tagChartCanvasElementCommands, withChartCanvasElement, type ChartCanvasElement } from './chart-element-interaction'
import type { ChartEditPanelAxis } from './chart-edit-panel-route'
import { pivotChartCategoryAxis } from './chart-pivot-axis-layout.ts'

export { drawChartSector } from './chart-sector-draw'

export function drawChartData(source: ChartData, chart: ChartModel, theme: ChartRenderTheme = {}, computed?: ChartComputedStatistics): DrawCommand[] {
  const palette = chartPalette(chart, theme.palette), derived = deriveChartData(source, chart, computed)
  return drawDerivedChartData(derived, chart, theme, palette)
}

export function drawDerivedChartData(derived: ChartDerivedData, chart: ChartModel, theme: ChartRenderTheme = {}, palette = chartPalette(chart, theme.palette), preparedLayout?: ChartLayout, preparedSankeyScene?: ChartSankeyScene, preparedChordScene?: ChartChordScene, preparedRelationScene?: ChartRelationScene): DrawCommand[] {
  const data = derived.data, layout = preparedLayout ?? layoutDerivedChartData(derived, chart, palette), descriptor = resolveChartType(chart.type)
  const heatmapScale = chart.type === 'heatmap' ? createChartHeatmapScale(chartHeatmapExtent(data, chart), palette, chart.options?.heatmap?.visualMapType) : undefined
  const background = chart.options?.appearance?.background ?? theme.backgroundColor ?? '#ffffff', textColor = chart.options?.appearance?.font?.color ?? theme.textColor ?? (chart.type === 'heatmap' ? '#334155' : '#64748b')
  const border = chart.options?.appearance?.border
  const commands: DrawCommand[] = [{ kind: 'rect', ...layout.frame, fill: background,
    ...(border?.visible === false ? {} : { stroke: border?.color ?? theme.axisColor ?? '#cbd5e1', strokeWidth: border?.width ?? 1 }) }]
  const titleSeed = chartTitleGeometry(chart).bands, legendWidth = Math.max(1, chart.size.width - titleSeed.left - titleSeed.right)
  const legendBand = chart.type === 'heatmap' ? chartHeatmapLegendBand(legendWidth, chart.options?.heatmap?.visualMapType ?? 'continuous', chart.legend.font?.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 11)
    : chartLegendGeometry(chart, chartLegendDataCount(data, chart)).band
  commands.push(...chartTitleCommands(chart, chart.options?.appearance?.font?.color ?? theme.textColor ?? '#334155', legendBand))
  if (!data.categories.length || !data.series.length || descriptor.family === 'gauge' && !data.series[0]!.values.some(Number.isFinite)) {
    commands.push({ kind: 'text', x: layout.plot.x + layout.plot.width / 2, y: layout.plot.y + layout.plot.height / 2,
      text: data.hierarchyDiagnostic ?? data.graph?.diagnostic ?? '暂无可绘制数据', font: chartFont(chart.options?.appearance?.font, 11), color: textColor, align: 'center', baseline: 'middle', maxWidth: layout.plot.width })
    return normalizeChartCommands(commands)
  }
  if (descriptor.threeD) {
    for (const command of chart3DScene(data, chart, layout, textColor, palette).commands) commands.push(command)
    drawPointLabels(commands, data, chart, layout, textColor, palette)
  } else if (descriptor.family === 'pie') drawPie(commands, data, chart, layout, textColor, palette)
  else if (descriptor.family === 'radar') {
    drawRadar(commands, data, chart, layout, textColor, palette, theme)
    drawPointLabels(commands, data, chart, layout, textColor, palette)
  }
  else if (descriptor.family === 'sankey') {
    if (!layout.sankey) throw new TypeError('桑基图绘制需要匹配的流带布局。')
    const scene = preparedSankeyScene ?? createChartSankeyScene(data, chart, layout.sankey, textColor, palette, layout.frame)
    for (const command of scene.commands) commands.push(command)
  }
  else if (descriptor.family === 'relation') {
    if (!layout.relation) throw new TypeError('关系图绘制需要匹配的 Worker 节点布局。')
    const scene = preparedRelationScene ?? createChartRelationScene(data, chart, layout.relation, theme, palette, layout.frame)
    for (const command of scene.commands) commands.push(command)
  }
  else if (descriptor.family === 'chord') {
    if (!layout.chord) throw new TypeError('和弦图绘制需要匹配的角度布局。')
    const scene = preparedChordScene ?? createChartChordScene(data, chart, layout.chord, theme, palette, layout.frame)
    for (const command of scene.commands) commands.push(command)
  }
  else if (descriptor.family === 'wordCloud') drawWordCloud(commands, data, chart, layout, palette)
  else if (descriptor.family === 'treemap') drawChartTreemap(commands, data, chart, layout, textColor, palette, background)
  else if (descriptor.family === 'sunburst') drawChartSunburst(commands, data, chart, layout, textColor, palette, background)
  else if (descriptor.family === 'gauge') drawGauge(commands, data, chart, layout, textColor, palette, theme)
  else if (descriptor.family === 'funnel') drawFunnel(commands, data, chart, layout, textColor, palette)
  else {
    drawAxes(commands, data, chart, layout, textColor, theme)
    if (descriptor.family === 'boxplot') drawBoxplot(commands, derived, chart, layout, textColor, palette)
    else if (descriptor.family === 'candlestick') drawChartCandlesticks(commands, data, chart, layout, palette, textColor)
    else if (descriptor.family === 'waterfall') drawWaterfall(commands, data, chart, layout, textColor, palette)
    else if (heatmapScale) drawChartHeatmapCells(commands, data, chart, layout, textColor, palette, heatmapScale)
    else drawCartesian(commands, data, chart, layout, palette)
    if (derived.pareto) drawParetoCumulative(commands, chart, layout, textColor, palette)
    if (descriptor.family !== 'boxplot' && descriptor.family !== 'waterfall' && descriptor.family !== 'heatmap') drawPointLabels(commands, data, chart, layout, textColor, palette,
      derived.pareto ? chartParetoValuePercentages(derived.pareto.values) : undefined)
    drawTrendlinesAndErrors(commands, derived, chart, layout, textColor, palette)
  }
  drawLegend(commands, derived, chart, layout, textColor, palette, heatmapScale)
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
  if (chart.type === 'heatmap') { drawChartHeatmapAxes(commands, data, chart, layout, color, theme); return }
  const { plot } = layout, horizontal = layout.horizontal, descriptor = resolveChartType(chart.type), numericX = layout.xMin !== undefined
  const gridColor = theme.splitLineColor ?? '#e5e7eb', axisColor = theme.axisColor ?? '#64748b'
  const yAxis = chart.yAxis, xAxis = chart.xAxis
  const line = (axisId: ChartEditPanelAxis, kind: 'axis' | 'gridLine' | 'tick', start: ChartVector, end: ChartVector,
    style: ChartLineOptions | undefined, color: string, clip?: RenderRect): void => {
    const first = commands.length
    chartLine(commands, start, end, style, color, clip)
    tagChartCanvasElementCommands(commands, first, kind === 'axis' ? { kind, axisId, part: 'line' } : { kind, axisId })
  }
  const label = (axisId: ChartEditPanelAxis, command: TextCommand): void => { commands.push(withChartCanvasElement(command, { kind: 'axis', axisId, part: 'label' })) }
  for (const value of chartAxisTicks(layout.min, layout.max, yAxis)) {
    const ratio = chartAxisRatio(value, layout.min, layout.max, yAxis), coordinate = horizontal ? plot.x + ratio * plot.width : plot.y + (1 - ratio) * plot.height
    const start = horizontal ? { x: coordinate, y: plot.y } : { x: plot.x, y: coordinate }, end = horizontal ? { x: coordinate, y: plot.y + plot.height } : { x: plot.x + plot.width, y: coordinate }
    line('yAxis', 'gridLine', start, end, yAxis.gridLine, gridColor, plot)
    if (yAxis.label?.visible !== false) label('yAxis', { kind: 'text', x: horizontal ? coordinate : plot.x - 8, y: horizontal ? plot.y + plot.height + 14 : coordinate,
      text: `${formatChartNumber(value, yAxis.format)}${descriptor.percent && !yAxis.format ? '%' : ''}`, font: chartFont({ ...chart.options?.appearance?.font, ...yAxis.label }, 10), color: yAxis.label?.color ?? color,
      align: horizontal ? 'center' : 'right', baseline: 'middle', rotation: yAxis.label?.rotate })
    if (yAxis.tick?.visible) {
      const length = (yAxis.tick.length ?? 4) * (yAxis.tick.position === 'inside' ? 1 : -1)
      line('yAxis', 'tick', horizontal ? { x: coordinate, y: plot.y + plot.height } : { x: plot.x, y: coordinate }, horizontal ? { x: coordinate, y: plot.y + plot.height - length } : { x: plot.x + length, y: coordinate }, yAxis.tick, axisColor)
    }
  }
  const zeroRatio = chartAxisRatio(yAxis.scaleType === 'log' ? layout.min : 0, layout.min, layout.max, yAxis)
  const zero = Math.max(0, Math.min(1, zeroRatio))
  line('xAxis', 'axis', horizontal ? { x: plot.x + zero * plot.width, y: plot.y } : { x: plot.x, y: plot.y + (1 - zero) * plot.height },
    horizontal ? { x: plot.x + zero * plot.width, y: plot.y + plot.height } : { x: plot.x + plot.width, y: plot.y + (1 - zero) * plot.height }, { ...xAxis.line, width: xAxis.line?.width ?? 1.25 }, axisColor)
  line('yAxis', 'axis', horizontal ? { x: plot.x, y: plot.y + plot.height } : { x: plot.x, y: plot.y },
    horizontal ? { x: plot.x + plot.width, y: plot.y + plot.height } : { x: plot.x, y: plot.y + plot.height }, yAxis.line, axisColor)
  const categories = layout.categories
  const pivotAxis = pivotChartCategoryAxis(data, chart, layout.frame, plot, color)
  if (layout.dateAxis) commands.push(...chartDateAxisCommands(chart, layout, color, theme))
  else if (numericX) for (const value of chartAxisTicks(layout.xMin!, layout.xMax!, xAxis)) {
    const ratio = chartAxisRatio(value, layout.xMin!, layout.xMax!, xAxis), coordinate = horizontal ? plot.y + ratio * plot.height : plot.x + ratio * plot.width
    line('xAxis', 'gridLine', horizontal ? { x: plot.x, y: coordinate } : { x: coordinate, y: plot.y }, horizontal ? { x: plot.x + plot.width, y: coordinate } : { x: coordinate, y: plot.y + plot.height }, xAxis.gridLine, gridColor, plot)
    if (xAxis.label?.visible !== false) label('xAxis', { kind: 'text', x: horizontal ? plot.x - 8 : coordinate, y: horizontal ? coordinate : plot.y + plot.height + 14,
      text: formatChartNumber(value, xAxis.format),
      font: chartFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(xAxis.label) }, 10), color: xAxis.label?.color ?? color, align: horizontal ? 'right' : 'center', baseline: 'middle', rotation: xAxis.label?.rotate })
    if (xAxis.tick?.visible) {
      const length = (xAxis.tick.length ?? 4) * (xAxis.tick.position === 'inside' ? -1 : 1)
      line('xAxis', 'tick', horizontal ? { x: plot.x, y: coordinate } : { x: coordinate, y: plot.y + plot.height },
        horizontal ? { x: plot.x - length, y: coordinate } : { x: coordinate, y: plot.y + plot.height + length }, xAxis.tick, axisColor)
    }
  } else categories.forEach((category, index) => {
    const ratio = xAxis.reverse ? 1 - (index + 0.5) / Math.max(1, categories.length) : (index + 0.5) / Math.max(1, categories.length)
    const coordinate = horizontal ? plot.y + ratio * plot.height : plot.x + ratio * plot.width
    if (xAxis.gridLine?.visible) line('xAxis', 'gridLine', horizontal ? { x: plot.x, y: coordinate } : { x: coordinate, y: plot.y },
      horizontal ? { x: plot.x + plot.width, y: coordinate } : { x: coordinate, y: plot.y + plot.height }, xAxis.gridLine, gridColor, plot)
    if (xAxis.tick?.visible) {
      const length = (xAxis.tick.length ?? 4) * (xAxis.tick.position === 'inside' ? -1 : 1)
      const start = horizontal ? { x: plot.x + zero * plot.width, y: coordinate } : { x: coordinate, y: plot.y + (1 - zero) * plot.height }
      line('xAxis', 'tick', start, { x: start.x - (horizontal ? length : 0), y: start.y + (horizontal ? 0 : length) }, xAxis.tick, axisColor)
    }
    if (!pivotAxis && xAxis.label?.visible !== false) label('xAxis', { kind: 'text', x: horizontal ? plot.x - 8 : plot.x + ratio * plot.width, y: horizontal ? plot.y + ratio * plot.height : plot.y + plot.height + 14,
      text: category, font: chartFont({ ...chart.options?.appearance?.font, ...xAxis.label }, 10), color: xAxis.label?.color ?? color, align: horizontal ? 'right' : 'center', baseline: 'middle', rotation: xAxis.label?.rotate,
      maxWidth: horizontal ? Math.max(1, plot.x - layout.frame.x - 12) : Math.max(1, plot.width / Math.max(1, categories.length) - 4) })
  })
  if (pivotAxis) commands.push(...pivotAxis.labels.map(label => label.command))
  if (layout.rightMin !== undefined && layout.rightMax !== undefined) {
    const axis = chart.options?.rightYAxis ?? { scaleType: 'linear' }
    line('rightYAxis', 'axis', horizontal ? { x: plot.x, y: plot.y } : { x: plot.x + plot.width, y: plot.y },
      { x: plot.x + plot.width, y: horizontal ? plot.y : plot.y + plot.height }, axis.line, axisColor)
    for (const value of chartAxisTicks(layout.rightMin, layout.rightMax, axis)) {
      const ratio = chartAxisRatio(value, layout.rightMin, layout.rightMax, axis)
      const point = horizontal ? { x: plot.x + ratio * plot.width, y: plot.y } : { x: plot.x + plot.width, y: plot.y + (1 - ratio) * plot.height }
      if (axis.gridLine?.visible) line('rightYAxis', 'gridLine', point, horizontal ? { x: point.x, y: plot.y + plot.height } : { x: plot.x, y: point.y }, axis.gridLine, gridColor, plot)
      if (axis.tick?.visible) {
        const length = (axis.tick.length ?? 4) * (axis.tick.position === 'inside' ? -1 : 1)
        line('rightYAxis', 'tick', point, { x: point.x + (horizontal ? 0 : length), y: point.y - (horizontal ? length : 0) }, axis.tick, axisColor)
      }
      if (axis.label?.visible !== false) label('rightYAxis', { kind: 'text', x: point.x + (horizontal ? 0 : 8), y: point.y - (horizontal ? 12 : 0),
        text: `${formatChartNumber(value, axis.format)}${chart.type === 'pareto' && !axis.format ? '%' : ''}`, font: chartFont({ ...chart.options?.appearance?.font, ...axis.label }, 10), color: axis.label?.color ?? color,
        align: horizontal ? 'center' : 'left', baseline: 'middle', rotation: axis.label?.rotate })
    }
  }
  for (const command of chartRectangularAxisTitleCommands(chart, layout, color)) commands.push(command)
}

function drawCartesian(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const bindings = chart.type === 'pareto' ? [] : chartCanvasSeriesBindings(data, chart.options)
  const { plot } = layout, horizontal = layout.horizontal, categoryStep = layout.categoryBandwidth ?? (horizontal ? plot.height : plot.width) / Math.max(1, layout.categories.length)
  const bars = layout.series.filter(series => series.points.length && (series.type === 'column' || series.type === 'bar'))
  const barGroups = [...new Set(bars.map(series => series.stacked ? `stack:${series.rightAxis ? 'right' : 'left'}` : `field:${series.field}`))]
  const gap = chart.type === 'histogram' ? chartHistogramBinGap(chart.options?.histogram) : (chart.options?.bar?.gapWidth ?? 150) / 100
  const groupWidth = categoryStep * (chart.type === 'histogram' ? 1 - gap : 1 / (1 + gap)), overlap = (chart.options?.bar?.overlap ?? 0) / 100
  const barWidth = groupWidth / Math.max(1, barGroups.length - (barGroups.length - 1) * overlap)
  const maxBubbleSize = data.series.reduce((max, source) => source.sizeValues?.reduce<number>((maximum, size) => Math.max(maximum, size ?? 0), max) ?? max, 1)
  layout.series.slice(0, data.series.length).forEach((series, seriesIndex) => {
    const binding = bindings[seriesIndex], source = data.series[seriesIndex]!, field = source.field
    const pointElement = (dataIndex: number): ChartCanvasElement | undefined => chart.type === 'pareto' ? chartCanvasParetoMark('bar', source)
      : binding ? { kind: 'dataPoint', binding, dataIndex } : undefined
    const style = chart.type === 'pareto' ? chartParetoBarStyle(chart, field, seriesIndex, palette, source.sourceName ?? source.name)
      : chartSeriesStyle(chart, field, seriesIndex, palette, source.sourceName ?? source.name)
    const axis = series.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const isBar = series.type === 'column' || series.type === 'bar' || chart.type === 'histogram' || chart.type === 'pareto'
    if (isBar) {
      const group = series.stacked ? `stack:${series.rightAxis ? 'right' : 'left'}` : `field:${field}`, groupIndex = Math.max(0, barGroups.indexOf(group))
      for (const point of series.points) {
        const first = commands.length
        const pointStyle = chartPointStyle(style, point.categoryIndex, chart.type), color = chartColorOpacity(pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1)
        const center = horizontal ? point.y : point.x, offset = center - groupWidth / 2 + groupIndex * barWidth * (1 - overlap)
        const rect = horizontal ? { x: Math.min(point.x, point.baseX!), y: offset, width: Math.abs(point.x - point.baseX!), height: Math.max(1, barWidth - 1) }
          : { x: offset, y: Math.min(point.y, point.baseY!), width: Math.max(1, barWidth - 1), height: Math.abs(point.y - point.baseY!) }
        const gradient = createChartMarkGradient(chart, { ...pointStyle, color: pointStyle.color ?? series.color }, horizontal ? 'bar' : 'column', { negative: point.value < 0, reverse: axis.reverse })
        const border = chart.type === 'pareto' ? chartFilledMarkBorder(pointStyle.color ?? series.color, pointStyle.border) : pointStyle.border
        chartRect(commands, rect, color, border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1), gradient)
        tagChartCanvasElementCommands(commands, first, pointElement(point.categoryIndex))
      }
      return
    }
    if (chart.type === 'scatter' || chart.type === 'bubble') {
      for (const point of series.points) {
        const first = commands.length
        const pointStyle = chartPointStyle(style, point.categoryIndex, chart.type)
        const size = chart.type === 'bubble' ? chartBubblePointSize(pointStyle, point.size ?? 0, maxBubbleSize, chart.options, Math.min(plot.width, plot.height)) : pointStyle.point?.size ?? 8
        if (chart.type === 'bubble' && (point.size === undefined || point.size <= 0)) continue
        chartSymbol(commands, point, size, chartColorOpacity(pointStyle.point?.color ?? pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1), pointStyle.point?.shape ?? 'circle', pointStyle.border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1))
        tagChartCanvasElementCommands(commands, first, pointElement(point.categoryIndex))
      }
      return
    }
    const lineStyle = resolveChartAreaLineStyle(chart, series.type === 'area' ? 'area' : 'line')
    const curvedPoints = lineStyle !== 'line' ? (function* (): Generator<ChartVector> {
      for (let index = 1; index < series.points.length; index++) {
        for (const point of chartLineSegment(series.points, index, lineStyle, chart.options?.invalidValueStrategy)) {
          yield point
          if (series.type === 'area') yield { x: point.x, y: point.baseY }
        }
      }
    })() : undefined
    const seriesBounds = chartSeriesGradientBounds(series, series.type === 'area', curvedPoints)
    const seriesStart = commands.length
    for (let index = 1; index < series.points.length; index++) {
      const previous = series.points[index - 1]!, point = series.points[index]!
      if (!connectedChartPoints(previous, point, chart)) continue
      const styleAtPoint = chartPointStyle(style, point.categoryIndex, chart.type), color = styleAtPoint.color ?? series.color
      const path = chartLineSegment(series.points, index, lineStyle, chart.options?.invalidValueStrategy)
      const areaGradient = series.type === 'area' ? createChartMarkGradient(chart, { ...styleAtPoint, color, fillOpacity: styleAtPoint.fillOpacity ?? 0.28 }, 'area', { reverse: axis.reverse }, isChartLinearGradient(color) ? seriesBounds : plot) : undefined
      const lineColor = styleAtPoint.border?.color ?? color
      const lineGradient = series.type === 'line' ? createChartMarkGradient(chart, { ...styleAtPoint, color: lineColor, fillOpacity: styleAtPoint.border?.opacity ?? 1 }, 'line', { reverse: axis.reverse }, isChartLinearGradient(lineColor) ? seriesBounds : plot) : undefined
      for (let step = 1; step < path.length; step++) {
        const a = path[step - 1]!, b = path[step]!
        if (series.type === 'area') {
          const base0 = a.baseY, base1 = b.baseY
          const fill = chartColorOpacity(color, styleAtPoint.fillOpacity ?? 0.28), pattern = chartFillPatternPaint(styleAtPoint.fillPattern, styleAtPoint.fillOpacity ?? 0.28)
          if ((a.y - base0) * (b.y - base1) < 0) {
            const ratio = Math.abs(a.y - base0) / (Math.abs(a.y - base0) + Math.abs(b.y - base1)), crossing = { x: a.x + (b.x - a.x) * ratio, y: base0 + (base1 - base0) * ratio }
            chartPolygon(commands, [{ x: a.x, y: base0 }, a, crossing], fill, undefined, plot, pattern, areaGradient)
            chartPolygon(commands, [crossing, b, { x: b.x, y: base1 }], fill, undefined, plot, pattern, areaGradient)
          } else chartPolygon(commands, [{ x: a.x, y: base0 }, a, b, { x: b.x, y: base1 }], fill, undefined, plot, pattern, areaGradient)
        }
        chartLine(commands, a, b, { ...styleAtPoint.border, color: styleAtPoint.border?.color ?? color, width: styleAtPoint.border?.width ?? 2 }, color, plot, lineGradient)
      }
    }
    tagChartCanvasElementCommands(commands, seriesStart, binding ? chartCanvasSeriesMark(binding) : undefined)
    series.points.forEach(point => {
      const first = commands.length
      const pointStyle = chartPointStyle(style, point.categoryIndex, chart.type)
      const color = pointStyle.point?.color ?? pointStyle.color ?? series.color
      const gradient = createChartMarkGradient(chart, { ...pointStyle, color: isChartLinearGradient(pointStyle.color) ? pointStyle.color : color }, 'line', { reverse: axis.reverse, symbol: true }, plot)
      chartSymbol(commands, point, pointStyle.point?.size ?? 6, chartColorOpacity(color, pointStyle.fillOpacity ?? 1), pointStyle.point?.shape ?? 'circle', pointStyle.border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1), gradient)
      tagChartCanvasElementCommands(commands, first, pointElement(point.categoryIndex))
    })
  })
}

function connectedChartPoints(previous: ChartPoint, point: ChartPoint, chart: ChartModel): boolean {
  return chartLinePointsConnected(previous, point, chart.options?.invalidValueStrategy)
}

function drawPointLabels(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[], businessPercentages?: readonly number[]): void {
  const bindings = chart.type === 'pareto' ? [] : chartCanvasSeriesBindings(data, chart.options)
  for (const [index, series] of layout.series.entries()) {
    if (chart.type === 'pareto' && index >= data.series.length) continue
    const source = data.series[index]
    const style = chart.type === 'pareto' ? chartParetoBarStyle(chart, source?.field ?? series.field ?? index, index, palette, source?.sourceName ?? source?.name)
      : chartSeriesStyle(chart, source?.field ?? series.field ?? index, index, palette, source?.sourceName ?? source?.name)
    if (!style.label?.visible) continue
    series.points.forEach(point => {
      const current = chartPointStyle(style, point.categoryIndex, chart.type)
      const label = current.label
      if (!label?.visible) return
      const anchor = chartPointLabelAnchor(label.position, point, layout.horizontal, label.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 11)
      const binding = bindings[index]
      commands.push(withChartCanvasElement({ kind: 'text', ...anchor, text: chartLabelText(label, point.value, point.category, series.name, businessPercentages?.[point.categoryIndex], point.size), font: chartFont({ ...chart.options?.appearance?.font, ...label }), color: label?.color ?? color, clip: layout.plot },
        chart.type === 'pareto' ? chartCanvasParetoMark('bar', source)
          : binding ? { kind: 'dataLabel', binding, dataIndex: point.categoryIndex } : undefined))
    })
  }
}

function drawParetoCumulative(commands: DrawCommand[], chart: ChartModel, layout: ChartLayout, textColor: string, palette: string[]): void {
  const first = commands.length
  const series = layout.series.at(-1)!, style = chartParetoCumulativeStyle(chart, palette), bounds = chartSeriesGradientBounds(series)
  const reverse = chart.options?.rightYAxis?.reverse
  for (let index = 1; index < series.points.length; index++) {
    const previous = series.points[index - 1]!, point = series.points[index]!
    if (!connectedChartPoints(previous, point, chart)) continue
    const current = chartParetoCumulativePointStyle(style, point.categoryIndex), color = current.border?.color ?? current.color!
    const gradient = createChartMarkGradient(chart, { ...current, color, fillOpacity: current.border?.opacity ?? 1 }, 'line', { reverse }, isChartLinearGradient(color) ? bounds : layout.plot)
    chartLine(commands, previous, point, { ...current.border, visible: true, color,
      width: current.border?.width !== undefined && current.border.width > 0 ? current.border.width : style.border?.width ?? 2 }, color, layout.plot, gradient)
  }
  for (const point of series.points) {
    const current = chartParetoCumulativeMarkerStyle(style, point.categoryIndex), label = current.label
    chartSymbol(commands, point, current.point?.size ?? 7, chartColorOpacity(current.color!, current.fillOpacity ?? 1),
      current.point?.shape ?? 'circle', current.border, layout.plot, chartFillPatternPaint(current.fillPattern, current.fillOpacity ?? 1))
    if (!label?.visible) continue
    const anchor = chartPointLabelAnchor(label.position, point, false, label.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 11)
    commands.push({ kind: 'text', ...anchor, text: chartLabelText(label, point.value, point.category, CHART_PARETO_CUMULATIVE_NAME, point.value),
      font: chartFont({ ...chart.options?.appearance?.font, ...label }), color: label.color ?? textColor, clip: layout.plot })
  }
  tagChartCanvasElementCommands(commands, first, { kind: 'mark', directEdit: { kind: 'paretoStyle', part: 'line' } })
}

function drawTrendlinesAndErrors(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const { plot } = layout
  const position = (x: number, y: number, seriesPosition: number): ChartVector => {
    const source = derived.data.series[seriesPosition]
    if (!source) throw new RangeError('统计图元对应的系列位置已经失效。')
    const style = chartSeriesStyle(chart, source.field, seriesPosition, palette, source.sourceName ?? source.name)
    const axis = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const min = style.rightAxis ? layout.rightMin! : layout.min, max = style.rightAxis ? layout.rightMax! : layout.max
    const coordinate = layout.dateAxis ? chartDateCoordinate(x, chart.xAxis.baseTimeUnit, layout.date1904) ?? NaN : x
    const xRatio = layout.xMin === undefined ? (x - 0.5) / Math.max(1, layout.categories.length) : chartAxisRatio(coordinate, layout.xMin, layout.xMax!, chart.xAxis)
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
  const bindings = chartCanvasSeriesBindings(data, chart.options)
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
        const distance = 4 + (tick?.visible && direction > 0 ? length : 0)
        commands.push({ kind: 'text', x: point.x + perpendicular.x * distance, y: point.y + perpendicular.y * distance,
          text: formatChartNumber(value, axisLabel.format), font: chartFont(axisLabel, 12), color: axisLabel.color ?? color,
          align: perpendicular.x > 1e-8 ? 'left' : perpendicular.x < -1e-8 ? 'right' : 'center', baseline: 'middle', rotation: chart.yAxis.label?.rotate, clip: layout.frame })
      }
    }
    if (chart.xAxis.label?.visible !== false) commands.push({ kind: 'text', ...vertex(index, 1.22), text: category,
      font: chartFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(chart.xAxis.label) }, 10), color: chart.xAxis.label?.color ?? color,
      align: 'center', baseline: 'middle', rotation: chart.xAxis.label?.rotate, clip: layout.frame })
  })
  for (const [index, series] of layout.series.entries()) {
    const binding = bindings[index], first = commands.length
    const source = data.series[index]!, style = chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), points = new Map(series.points.map(point => [point.categoryIndex, point]))
    const polygon = data.categories.map((_category, categoryIndex) => points.get(categoryIndex) ?? center)
    if (chart.options?.radar?.fill) chartPolygon(commands, polygon, chartColorOpacity(style.color ?? series.color, style.fillOpacity ?? 0.25), undefined, plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 0.25))
    polygon.forEach((point, categoryIndex) => chartLine(commands, point, polygon[(categoryIndex + 1) % polygon.length]!, { color: series.color, width: 2, ...style.border }, series.color))
    tagChartCanvasElementCommands(commands, first, binding ? chartCanvasSeriesMark(binding) : undefined)
    series.points.forEach(point => {
      const pointStart = commands.length
      const pointStyle = chartPointStyle(style, point.categoryIndex, chart.type)
      chartSymbol(commands, point, pointStyle.point?.size ?? 4, chartColorOpacity(pointStyle.point?.color ?? pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1),
        pointStyle.point?.shape ?? 'none', pointStyle.border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1))
      tagChartCanvasElementCommands(commands, pointStart, binding ? { kind: 'dataPoint', binding, dataIndex: point.categoryIndex } : undefined)
    })
  }
  for (const command of chartRectangularAxisTitleCommands(chart, layout, color)) commands.push(command)
}

function drawPie(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const pie = chartPieLayout(data, chart, layout.plot), source = data.series[0]!, field = source.field
  const binding = chartCanvasSeriesBindings(data, chart.options)[0]
  const aggregateIndex = data.categories.length
  const aggregateElement: ChartCanvasElement = { kind: 'mark', directEdit: { kind: 'pieAggregateStyle',
    sourceField: field, sourceName: source.sourceName ?? source.name, aggregateIndex } }
  const labelElement: ChartCanvasElement = { kind: 'mark', directEdit: { kind: 'pieStyle' } }
  const element = (kind: 'dataPoint' | 'dataLabel', dataIndex: number, secondary = false): ChartCanvasElement | undefined => dataIndex >= 0 && dataIndex < source.values.length
    ? kind === 'dataLabel' ? labelElement
      : isChartCompoundPie(chart.type) || chart.options?.pie?.sliceStyles?.[dataIndex] !== undefined
        ? chartCanvasPieSliceMark(source, dataIndex, secondary ? 'secondary' : 'primary')
        : binding ? { kind, binding, dataIndex } : undefined : undefined
  const connector = chartPieConnector(chart, color)
  if (connector) for (const [start, end] of pie.connectors) {
    const first = commands.length
    chartLine(commands, start, end, connector, color, layout.plot)
    tagChartCanvasElementCommands(commands, first, { kind: 'mark', directEdit: { kind: 'pieConnectorStyle' } })
  }
  for (const sector of pie.sectors) {
    const first = commands.length
    const aggregate = isChartCompoundPie(chart.type) && !sector.secondary && sector.part.index === aggregateIndex
    const style = aggregate ? chartPieAggregateStyle(chart, source, aggregateIndex, palette)
      : chartPieSliceStyle(chart, source, sector.part.index, palette)
    const fill = style.color ?? palette[sector.part.index % palette.length]!, gradient = createChartMarkGradient(chart, { ...style, color: fill }, 'pie')
    drawChartSector(commands, sector.center, sector.inner, sector.outer, sector.start, sector.end, chartColorOpacity(fill, style.fillOpacity ?? 1), style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1), gradient)
    tagChartCanvasElementCommands(commands, first, aggregate ? aggregateElement : element('dataPoint', sector.part.index, sector.secondary))
    const label = { ...style.label, ...chart.options?.pie?.label }
    if (label.visible !== false) {
      const outside = label.position !== 'inside' && label.position !== 'center', point = chartPolar(sector.center, outside ? sector.outer * 1.12 : (sector.inner + sector.outer) / 2, (sector.start + sector.end) / 2)
      if (outside && chart.options?.pie?.showLabelLine !== false) {
        const guideStart = commands.length
        chartLine(commands, chartPolar(sector.center, sector.outer, (sector.start + sector.end) / 2), point, { color, width: 0.8 }, color)
        tagChartCanvasElementCommands(commands, guideStart, aggregate ? aggregateElement : labelElement)
      }
      commands.push(withChartCanvasElement({ kind: 'text', ...point, text: chartLabelText({ contentType: 'category', ...label }, sector.part.value, sector.part.category, data.series[0]!.name, sector.percentage), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot }, aggregate ? aggregateElement : element('dataLabel', sector.part.index)))
    }
  }
  for (const bar of pie.bars) {
    const first = commands.length
    const style = chartPieSliceStyle(chart, source, bar.part.index, palette)
    const fill = style.color ?? palette[bar.part.index % palette.length]!, gradient = createChartMarkGradient(chart, { ...style, color: fill }, 'column')
    chartRect(commands, bar.rect, chartColorOpacity(fill, style.fillOpacity ?? 1), style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1), gradient)
    tagChartCanvasElementCommands(commands, first, element('dataPoint', bar.part.index, true))
    const label = { ...style.label, ...chart.options?.pie?.label }
    if (label.visible !== false) commands.push(withChartCanvasElement({ kind: 'text', x: bar.rect.x + bar.rect.width / 2, y: bar.rect.y + bar.rect.height / 2, text: chartLabelText({ contentType: 'category', ...label }, bar.part.value, bar.part.category, data.series[0]!.name, bar.percentage), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', maxWidth: bar.rect.width, clip: layout.plot }, element('dataLabel', bar.part.index)))
  }
}

function drawBoxplot(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, textColor: string, palette: string[]): void {
  const { plot } = layout, boxes = derived.boxes!, bindings = chartCanvasBoxplotBindings(boxes, chart.options)
  boxes.forEach((box, index) => {
    if (!box.statistics) return
    const first = commands.length, binding = bindings[index]
    const style = chartBoxplotStyle(chart, box, index, palette), statistics = box.statistics, color = style.color ?? palette[box.seriesPosition % palette.length]!
    const geometry = chartBoxplotGeometry(chart, layout, style, index, boxes.length), { x, width } = geometry, summary = geometry.summary(statistics)
    if (summary && derived.legendSelection?.['样本分布'] !== false) {
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
    tagChartCanvasElementCommands(commands, first, binding ? chartCanvasSeriesMark(binding) : undefined)
    if (chart.options?.boxplot?.showOutliers !== false && derived.legendSelection?.['异常值'] !== false) for (const outlier of chartBoxplotOutliers(box)) {
      const pointStart = commands.length
      const pointStyle = chartBoxplotOutlierStyle(style, outlier.sourceIndex)
      const y = geometry.y(outlier.value), size = pointStyle.point?.size ?? 4, shape = pointStyle.point?.shape ?? 'circle'
      if (y === undefined || size <= 0 || shape === 'none') continue
      const markerColor = pointStyle.point?.color ?? pointStyle.color ?? color, fill = chartColorOpacity(markerColor, pointStyle.fillOpacity ?? 1),
        border = { color: markerColor, width: 1.5, ...pointStyle.border }, marker = chartMarkerShape(shape)
      if (marker.base === 'rect' && !marker.hollow) {
        const half = size / 2, corners = [{ x: x - half, y: y - half }, { x: x + half, y: y - half }, { x: x + half, y: y + half }, { x: x - half, y: y + half }]
        chartRect(commands, { x: x - half, y: y - half, width: size, height: size }, fill, undefined, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1))
        corners.forEach((point, position) => chartLine(commands, point, corners[(position + 1) % corners.length]!, border, markerColor, plot))
      } else chartSymbol(commands, { x, y }, size, fill, shape, border, plot, chartFillPatternPaint(pointStyle.fillPattern, pointStyle.fillOpacity ?? 1))
      tagChartCanvasElementCommands(commands, pointStart, binding ? { kind: 'dataPoint', binding, dataIndex: outlier.sourceIndex } : undefined)
    }
  })
}

function drawWaterfall(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, textColor: string, palette: string[]): void {
  const scene = layout.waterfall
  if (!scene) throw new ComputeError('COMPUTE_NOT_READY', '瀑布图绘制需要完整的计算布局。')
  const { plot } = layout, step = plot.width / Math.max(1, layout.categories.length)
  const width = Math.max(1, step / (1 + (chart.options?.bar?.gapWidth ?? 150) / 100) - 1), connector = chartWaterfallConnector(chart)
  const roles = data.series.map((source, index) => chartWaterfallRoleStyles(chart, source.field, index, palette, source.sourceName ?? source.name))
  const globalRoles = chartWaterfallRoleStyles(chart, undefined, 0, palette)
  const bindings = chartCanvasSeriesBindings(data, chart.options)
  if (connector) for (const segment of scene.connectors) {
    const direction = Math.sign(segment.to.x - segment.from.x), half = Math.min(width / 2, Math.abs(segment.to.x - segment.from.x) / 2)
    chartLine(commands, { x: segment.from.x + direction * half, y: segment.from.y },
      { x: segment.to.x - direction * half, y: segment.to.y }, connector, connector.color, plot)
  }
  for (const point of scene.points) {
    const first = commands.length, binding = bindings[point.seriesPosition]
    const element: ChartCanvasElement | undefined = point.seriesPosition === CHART_WATERFALL_GLOBAL_SERIES
      ? { kind: 'mark', directEdit: { kind: 'waterfallStyle', part: 'subtotal' } }
      : binding ? { kind: 'mark', directEdit: { kind: 'waterfallStyle', part: point.waterfallRole, binding } } : undefined
    const group = point.seriesPosition === CHART_WATERFALL_GLOBAL_SERIES ? globalRoles : roles[point.seriesPosition]!
    const style = chartPointStyle(group[point.waterfallRole], point.dataPointIndex, chart.type), color = style.color!
    const axis = point.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const gradient = createChartMarkGradient(chart, { ...style, color }, 'column', { negative: point.value < 0, reverse: axis.reverse })
    chartRect(commands, { x: point.x - width / 2, y: Math.min(point.y, point.baseY), width, height: Math.abs(point.y - point.baseY) },
      chartColorOpacity(color, style.fillOpacity ?? 1), style.border, plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1), gradient)
    tagChartCanvasElementCommands(commands, first, element)
    const label = style.label, inside = label?.position === 'inside' || label?.position === 'center'
    if (!label?.visible) continue
    const x = point.x + (label.position === 'left' ? -9 : label.position === 'right' ? 9 : 0)
    const y = inside ? (point.y + point.baseY) / 2 : point.y + (label.position === 'left' || label.position === 'right' ? 0 : label.position === 'bottom' ? 10 : -9)
    commands.push(withChartCanvasElement({ kind: 'text', x, y, text: chartLabelText(label, point.value, point.category, style.name ?? point.category),
      font: chartFont({ ...chart.options?.appearance?.font, ...label }), color: label.color ?? textColor,
      align: label.position === 'left' ? 'right' : label.position === 'right' ? 'left' : 'center', baseline: 'middle', clip: plot }, element))
  }
}

function drawWordCloud(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const words = layout.wordCloud ?? []
  for (const word of words) {
    const style = chartPointStyle(chartSeriesStyle(chart, data.series[0]!.field, word.categoryIndex, palette, data.series[0]!.sourceName ?? data.series[0]!.name), word.categoryIndex, chart.type)
    commands.push({ kind: 'text', x: word.x + word.width / 2, y: word.y + word.height / 2, text: word.text, font: word.font, color: chartColorOpacity(style.color ?? palette[word.categoryIndex % palette.length]!, style.fillOpacity ?? 1), align: 'center', baseline: 'middle', rotation: word.rotation, clip: layout.plot })
  }
}

function drawGauge(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[], theme: ChartRenderTheme): void {
  const authored = chart.options?.gauge, source = data.series[0]!
  for (const tile of chartGaugeTiles(data, layout.plot)) {
    const index = tile.categoryIndex, value = source.values[index]!, name = data.categories[index]!
    const style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index, chart.type), label = style.label
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
  const binding = chartCanvasSeriesBindings(data, chart.options)[0]
  for (const step of chartFunnelLayout(data, chart, layout.plot)) {
    const first = commands.length
    const style = chartPointStyle(chartSeriesStyle(chart, source.field, step.categoryIndex, palette, source.sourceName ?? source.name), step.categoryIndex, chart.type)
    chartPolygon(commands, step.points, chartColorOpacity(style.color ?? palette[step.categoryIndex % palette.length]!, style.fillOpacity ?? 1), style.border, layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
    tagChartCanvasElementCommands(commands, first, binding ? { kind: 'dataPoint', binding, dataIndex: step.categoryIndex } : undefined)
    if (style.label?.visible !== false) commands.push(withChartCanvasElement({ kind: 'text', ...step.center, text: chartLabelText({ contentType: 'category', ...style.label }, step.value, step.category, source.name), font: chartFont(style.label, 10), color: style.label?.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot },
      binding ? { kind: 'dataLabel', binding, dataIndex: step.categoryIndex } : undefined))
  }
}

function drawLegend(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[], heatmapScale?: ChartHeatmapScale): void {
  if (!chart.legend.visible || !chartHasLegend(chart.type)) return
  if (heatmapScale) { for (const command of createChartHeatmapLegend(chart, layout, heatmapScale, color).commands) commands.push(command); return }
  const items = [...new Map(chartLegendItems(derived, chart, layout, palette).map(item => [item.name, item])).values()]
  const font: ChartFontOptions = { ...chart.options?.appearance?.font, ...chart.legend.font }
  const geometry = chartLegendGeometry(chart, items.length), { horizontal, perRow, itemWidth, rowHeight } = geometry
  const titles = chartTitleGeometry(chart, geometry.band).bands
  const limit = horizontal ? geometry.wrap ? items.length : perRow : Math.max(1, Math.floor(layout.plot.height / rowHeight))
  items.slice(0, limit).forEach((item, index) => {
    const x = horizontal ? layout.frame.x + geometry.startX + index % perRow * itemWidth : chart.legend.position === 'left' ? layout.frame.x + titles.left + 8
      : layout.frame.x + layout.frame.width - titles.right - Math.min(96, layout.frame.width * 0.24) + 8
    const y = horizontal ? chart.legend.position === 'top' ? layout.frame.y + titles.top + 12 + Math.floor(index / perRow) * rowHeight
      : layout.frame.y + layout.frame.height - titles.bottom - geometry.band + 12 + Math.floor(index / perRow) * rowHeight : layout.plot.y + 10 + index * rowHeight
    chartRect(commands, { x, y: y - 5, width: 10, height: 10 }, item.color, undefined, layout.frame, item.pattern)
    commands.push({ kind: 'text', x: x + 15, y, text: item.name, font: chartFont(font, 11), color: font.color ?? color, baseline: 'middle', maxWidth: geometry.textWidth, clip: layout.frame })
  })
}
