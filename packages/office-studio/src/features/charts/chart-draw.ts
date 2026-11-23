import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartModel, ChartSeriesOptions } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartData, ChartLayout, ChartPoint } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { deriveChartData, chartPointStyle, type ChartComputedStatistics, type ChartDerivedData } from './chart-derived'
import { layoutDerivedChartData } from './chart-layout'
import {
  chartArc, chartAxisRatio, chartAxisTicks, chartBezier, chartCircle, chartColorOpacity, chartFont, chartLabelText, chartLine,
  chartPolar, chartPolygon, chartShade, chartSymbol, formatChartNumber, type ChartVector,
} from './chart-geometry'
import { chartPalette, chartSeriesStyle } from './chart-options'
import { chart3DScene } from './chart-3d'
import { chartGaugeTiles, chartPieLayout } from './chart-pie-layout'
import { chartChordLayout, chartRelationLayout, chartRibbon, chartSankeyLayout, chartSunburstLayout, chartTreemapLayout, chartWordCloudLayout } from './chart-special-layout'
import type { ChartRenderTheme } from './chart-theme'

export function drawChartData(source: ChartData, chart: ChartModel, theme: ChartRenderTheme = {}, computed?: ChartComputedStatistics): DrawCommand[] {
  const palette = chartPalette(chart, theme.palette), derived = deriveChartData(source, chart, computed)
  return drawDerivedChartData(derived, chart, theme, palette)
}

export function drawDerivedChartData(derived: ChartDerivedData, chart: ChartModel, theme: ChartRenderTheme = {}, palette = chartPalette(chart, theme.palette)): DrawCommand[] {
  const data = derived.data, layout = layoutDerivedChartData(derived, chart, palette), descriptor = resolveChartType(chart.type)
  const background = chart.options?.appearance?.background ?? theme.backgroundColor ?? '#ffffff', textColor = chart.options?.appearance?.font?.color ?? theme.textColor ?? '#64748b'
  const border = chart.options?.appearance?.border
  const commands: DrawCommand[] = [{ kind: 'rect', ...layout.frame, fill: background,
    ...(border?.visible === false ? {} : { stroke: border?.color ?? theme.axisColor ?? '#cbd5e1', strokeWidth: border?.width ?? 1 }) }]
  const titleFont = { ...chart.options?.appearance?.font, ...chart.options?.titleStyle }, subtitleFont = { ...chart.options?.appearance?.font, ...chart.options?.subtitleStyle }
  if (chart.title) commands.push({ kind: 'text', x: layout.frame.x + layout.frame.width / 2, y: layout.frame.y + 20, text: chart.title, font: chartFont({ bold: true, ...titleFont }, 14), color: titleFont.color ?? textColor, align: 'center', baseline: 'middle', maxWidth: layout.frame.width - 20 })
  if (chart.subtitle) commands.push({ kind: 'text', x: layout.frame.x + layout.frame.width / 2, y: layout.frame.y + (chart.title ? 43 : 20), text: chart.subtitle, font: chartFont(subtitleFont, 11), color: subtitleFont.color ?? textColor, align: 'center', baseline: 'middle', maxWidth: layout.frame.width - 20 })
  if (descriptor.threeD) {
    commands.push(...chart3DScene(data, chart, layout, textColor, palette).commands)
    drawPointLabels(commands, data, chart, layout, textColor, palette)
  } else if (descriptor.family === 'pie') drawPie(commands, data, chart, layout, textColor, palette)
  else if (descriptor.family === 'radar') drawRadar(commands, data, chart, layout, textColor, palette, theme)
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
    else if (descriptor.family === 'candlestick') drawCandlestick(commands, data, chart, layout)
    else if (descriptor.family === 'waterfall') drawWaterfall(commands, derived, chart, layout, palette)
    else if (descriptor.family === 'heatmap') drawHeatmap(commands, data, chart, layout, textColor, palette)
    else drawCartesian(commands, data, chart, layout, palette)
    if (derived.pareto) {
      const axis: ChartAxisOptions = chart.options?.rightYAxis ?? { scaleType: 'linear' }, color = palette[1 % palette.length]!
      const points = derived.pareto.cumulativePercent.map((value, index) => ({ x: layout.plot.x + (index + 0.5) / Math.max(1, data.categories.length) * layout.plot.width,
        y: layout.plot.y + (1 - chartAxisRatio(value, layout.rightMin!, layout.rightMax!, axis)) * layout.plot.height }))
      for (let index = 1; index < points.length; index++) chartLine(commands, points[index - 1]!, points[index]!, { width: 2, color }, color, layout.plot)
      points.forEach(point => chartCircle(commands, point, 2, color, undefined, layout.plot))
    }
    drawPointLabels(commands, data, chart, layout, textColor, palette)
    drawTrendlinesAndErrors(commands, derived, chart, layout, textColor, palette)
  }
  drawLegend(commands, data, chart, layout, textColor, palette)
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
  chartLine(commands, { x: plot.x, y: plot.y }, { x: plot.x, y: plot.y + plot.height }, yAxis.line, axisColor)
  const categories = chart.type === 'heatmap' ? data.categories : layout.categories
  if (numericX) for (const value of chartAxisTicks(layout.xMin!, layout.xMax!, xAxis)) {
    const x = plot.x + chartAxisRatio(value, layout.xMin!, layout.xMax!, xAxis) * plot.width
    chartLine(commands, { x, y: plot.y }, { x, y: plot.y + plot.height }, xAxis.gridLine, gridColor, plot)
    if (xAxis.label?.visible !== false) commands.push({ kind: 'text', x, y: plot.y + plot.height + 14, text: formatChartNumber(value, xAxis.format), font: chartFont(xAxis.label, 10), color: xAxis.label?.color ?? color, align: 'center', baseline: 'middle', rotation: xAxis.label?.rotate })
  } else categories.forEach((category, index) => {
    const ratio = xAxis.reverse ? 1 - (index + 0.5) / Math.max(1, categories.length) : (index + 0.5) / Math.max(1, categories.length)
    if (xAxis.label?.visible !== false) commands.push({ kind: 'text', x: horizontal ? plot.x - 8 : plot.x + ratio * plot.width, y: horizontal ? plot.y + ratio * plot.height : plot.y + plot.height + 14,
      text: category, font: chartFont({ ...chart.options?.appearance?.font, ...xAxis.label }, 10), color: xAxis.label?.color ?? color, align: horizontal ? 'right' : 'center', baseline: 'middle', rotation: xAxis.label?.rotate,
      maxWidth: horizontal ? Math.max(1, plot.x - layout.frame.x - 12) : Math.max(1, plot.width / Math.max(1, categories.length) - 4) })
  })
  if (chart.type === 'heatmap') data.series.forEach((series, index) => commands.push({ kind: 'text', x: plot.x - 8, y: plot.y + (index + 0.5) / Math.max(1, data.series.length) * plot.height, text: chartSeriesStyle(chart, series.field, index).name ?? series.name, font: chartFont(yAxis.label, 10), color: yAxis.label?.color ?? color, align: 'right', baseline: 'middle' }))
  if (layout.rightMin !== undefined && layout.rightMax !== undefined) {
    const axis = chart.options?.rightYAxis ?? { scaleType: 'linear' }, x = plot.x + plot.width
    chartLine(commands, { x, y: plot.y }, { x, y: plot.y + plot.height }, axis.line, axisColor)
    if (axis.label?.visible !== false) for (const value of chartAxisTicks(layout.rightMin, layout.rightMax, axis)) commands.push({ kind: 'text', x: x + 8, y: plot.y + (1 - chartAxisRatio(value, layout.rightMin, layout.rightMax, axis)) * plot.height,
      text: `${formatChartNumber(value, axis.format)}${chart.type === 'pareto' && !axis.format ? '%' : ''}`, font: chartFont(axis.label, 10), color: axis.label?.color ?? color, baseline: 'middle' })
    if (axis.title) commands.push({ kind: 'text', x: layout.frame.x + layout.frame.width - 12, y: plot.y + plot.height / 2, text: axis.title, font: chartFont({}, 11), color, rotation: 90, align: 'center', baseline: 'middle' })
  }
  const bottomTitle = horizontal ? yAxis.title : xAxis.title, sideTitle = horizontal ? xAxis.title : yAxis.title
  if (bottomTitle) commands.push({ kind: 'text', x: plot.x + plot.width / 2, y: plot.y + plot.height + 34, text: bottomTitle, font: chartFont({}, 11), color, align: 'center', baseline: 'middle' })
  if (sideTitle) commands.push({ kind: 'text', x: layout.frame.x + 13, y: plot.y + plot.height / 2, text: sideTitle, font: chartFont({}, 11), color, align: 'center', baseline: 'middle', rotation: -90 })
}

function drawCartesian(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const { plot } = layout, horizontal = layout.horizontal, categoryStep = (horizontal ? plot.height : plot.width) / Math.max(1, layout.categories.length)
  const bars = layout.series.filter(series => series.type === 'column' || series.type === 'bar')
  const barGroups = [...new Set(bars.map(series => series.stacked ? `stack:${series.rightAxis ? 'right' : 'left'}` : `field:${series.field}`))]
  const gap = chart.type === 'histogram' ? chart.options?.histogram?.binGap ?? 0.05 : (chart.options?.bar?.gapWidth ?? 150) / 100
  const groupWidth = categoryStep * (chart.type === 'histogram' ? 1 - gap : 1 / (1 + gap)), overlap = (chart.options?.bar?.overlap ?? 0) / 100
  const barWidth = groupWidth / Math.max(1, barGroups.length - (barGroups.length - 1) * overlap)
  layout.series.forEach((series, seriesIndex) => {
    const field = data.series[seriesIndex]!.field, style = chartSeriesStyle(chart, field, seriesIndex, palette)
    const isBar = series.type === 'column' || series.type === 'bar' || chart.type === 'histogram' || chart.type === 'pareto'
    if (isBar) {
      const group = series.stacked ? `stack:${series.rightAxis ? 'right' : 'left'}` : `field:${field}`, groupIndex = Math.max(0, barGroups.indexOf(group))
      for (const point of series.points) {
        const pointStyle = chartPointStyle(style, point.categoryIndex), color = chartColorOpacity(pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1)
        const center = horizontal ? point.y : point.x, offset = center - groupWidth / 2 + groupIndex * barWidth * (1 - overlap)
        const rect = horizontal ? { x: Math.min(point.x, point.baseX!), y: offset, width: Math.abs(point.x - point.baseX!), height: Math.max(1, barWidth - 1) }
          : { x: offset, y: Math.min(point.y, point.baseY!), width: Math.max(1, barWidth - 1), height: Math.abs(point.y - point.baseY!) }
        commands.push({ kind: 'rect', ...rect, fill: color, ...(pointStyle.border?.visible === false ? {} : { stroke: pointStyle.border?.color, strokeWidth: pointStyle.border?.width }), clip: plot })
      }
      return
    }
    if (chart.type === 'scatter' || chart.type === 'bubble') {
      let maxSize = 1
      for (const point of series.points) maxSize = Math.max(maxSize, point.size ?? 1)
      for (const point of series.points) {
        const pointStyle = chartPointStyle(style, point.categoryIndex), valueSize = (point.size ?? 1) / maxSize
        const size = chart.type === 'bubble' ? Math.max(0, (chart.options?.bubble?.scale ?? 1) * Math.min(54, Math.min(plot.width, plot.height) * 0.2) * (chart.options?.bubble?.sizeRepresents === 'width' ? valueSize : Math.sqrt(valueSize))) : pointStyle.point?.size ?? 6
        if (chart.type === 'bubble' && (point.size === undefined || point.size <= 0)) continue
        chartSymbol(commands, point, size, chartColorOpacity(pointStyle.point?.color ?? pointStyle.color ?? series.color, pointStyle.fillOpacity ?? 1), chart.type === 'bubble' ? 'circle' : pointStyle.point?.shape ?? 'rect', pointStyle.border, plot)
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
          const fill = chartColorOpacity(color, styleAtPoint.fillOpacity === 1 ? 0.28 : styleAtPoint.fillOpacity ?? 0.28)
          if ((a.y - base0) * (b.y - base1) < 0) {
            const ratio = Math.abs(a.y - base0) / (Math.abs(a.y - base0) + Math.abs(b.y - base1)), crossing = { x: a.x + (b.x - a.x) * ratio, y: base0 + (base1 - base0) * ratio }
            chartPolygon(commands, [{ x: a.x, y: base0 }, a, crossing], fill, undefined, plot)
            chartPolygon(commands, [crossing, b, { x: b.x, y: base1 }], fill, undefined, plot)
          } else chartPolygon(commands, [{ x: a.x, y: base0 }, a, b, { x: b.x, y: base1 }], fill, undefined, plot)
        }
        chartLine(commands, a, b, { ...styleAtPoint.border, color: styleAtPoint.border?.color ?? color, width: styleAtPoint.border?.width ?? 2 }, color, plot)
      }
    }
    series.points.forEach(point => { const pointStyle = chartPointStyle(style, point.categoryIndex); chartSymbol(commands, point, pointStyle.point?.size ?? (series.type === 'area' ? 4 : 6), pointStyle.point?.color ?? pointStyle.color ?? series.color, pointStyle.point?.shape ?? 'rect', undefined, plot) })
  })
}

function smoothChartSegment(points: ChartPoint[], index: number): ChartVector[] {
  const a = points[index - 1]!, b = points[index]!, previous = points[Math.max(0, index - 2)]!, next = points[Math.min(points.length - 1, index + 1)]!
  return chartBezier(a, { x: a.x + (b.x - previous.x) / 6, y: a.y + (b.y - previous.y) / 6 }, { x: b.x - (next.x - a.x) / 6, y: b.y - (next.y - a.y) / 6 }, b, 12)
}

function drawPointLabels(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  for (const [index, series] of layout.series.entries()) {
    const style = chartSeriesStyle(chart, data.series[index]?.field ?? series.field ?? index, index, palette)
    if (!style.label?.visible) continue
    series.points.forEach(point => {
      const label = chartPointStyle(style, point.categoryIndex).label, inside = label?.position === 'inside' || label?.position === 'center'
      const x = inside && layout.horizontal ? (point.x + (point.baseX ?? point.x)) / 2 : point.x
      const y = inside && !layout.horizontal ? (point.y + (point.baseY ?? point.y)) / 2 : point.y - (label?.position === 'bottom' ? -10 : 9)
      commands.push({ kind: 'text', x, y, text: chartLabelText(label, point.value, point.category, series.name), font: chartFont({ ...chart.options?.appearance?.font, ...label }), color: label?.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot })
    })
  }
}

function drawTrendlinesAndErrors(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const { plot } = layout
  const position = (x: number, y: number, field: number): ChartVector => {
    const seriesIndex = derived.data.series.findIndex(series => series.field === field), style = chartSeriesStyle(chart, field, Math.max(0, seriesIndex), palette)
    const axis = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const min = style.rightAxis ? layout.rightMin! : layout.min, max = style.rightAxis ? layout.rightMax! : layout.max
    const xRatio = layout.xMin === undefined ? (x - 0.5) / Math.max(1, layout.categories.length) : chartAxisRatio(x, layout.xMin, layout.xMax!, chart.xAxis)
    const ratio = chartAxisRatio(y, min, max, axis)
    return layout.horizontal ? { x: plot.x + ratio * plot.width, y: plot.y + (chart.xAxis.reverse ? 1 - xRatio : xRatio) * plot.height }
      : { x: plot.x + (layout.xMin === undefined && chart.xAxis.reverse ? 1 - xRatio : xRatio) * plot.width, y: plot.y + (1 - ratio) * plot.height }
  }
  let annotation = 0
  for (const trend of derived.trends) {
    if (trend.config.visible === false) continue
    const points = trend.statistics.points.map(([x, y]) => position(x, y, trend.field))
    for (let index = 1; index < points.length; index++) chartLine(commands, points[index - 1]!, points[index]!, { color: trend.config.color ?? '#7c3aed', width: trend.config.width ?? 1.5, ...trend.config }, '#7c3aed', plot)
    const text = [trend.config.showFormula ? trend.statistics.formula : undefined, trend.config.showR2 && trend.statistics.r2 !== undefined ? `R² = ${formatChartNumber(trend.statistics.r2)}` : undefined].filter(Boolean).join('   ')
    if (text) commands.push({ kind: 'text', x: plot.x + 6, y: plot.y + 12 + annotation++ * 14, text, font: chartFont({}, 10), color: trend.config.color ?? color, baseline: 'middle', maxWidth: plot.width - 12, clip: plot })
  }
  for (const error of derived.errors) {
    if (error.config.line?.visible === false) continue
    const series = layout.series.find(series => series.field === error.field)
    for (const interval of error.intervals) {
      const point = series?.points.find(point => point.categoryIndex === interval.categoryIndex)
      const value = interval.y + (series?.stacked ? point?.baseValue ?? 0 : 0)
      const x = layout.xMin === undefined ? interval.categoryIndex + 1 : interval.x
      const start = position(error.config.direction === 'x' ? x - interval.minus : x, error.config.direction === 'y' ? value - interval.minus : value, error.field)
      const end = position(error.config.direction === 'x' ? x + interval.plus : x, error.config.direction === 'y' ? value + interval.plus : value, error.field)
      chartLine(commands, start, end, error.config.line, '#334155', plot)
      if (!error.config.noEndCap) {
        const vertical = (error.config.direction === 'y') !== Boolean(layout.horizontal)
        for (const anchor of [start, end]) chartLine(commands, vertical ? { x: anchor.x - 4, y: anchor.y } : { x: anchor.x, y: anchor.y - 4 }, vertical ? { x: anchor.x + 4, y: anchor.y } : { x: anchor.x, y: anchor.y + 4 }, error.config.line, '#334155', plot)
      }
    }
  }
}

function drawRadar(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[], theme: ChartRenderTheme): void {
  const { plot } = layout, count = data.categories.length, center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }, radius = Math.min(plot.width, plot.height) * 0.36
  const vertex = (index: number, ratio: number): ChartVector => chartPolar(center, radius * ratio, -Math.PI / 2 + index * Math.PI * 2 / Math.max(1, count))
  for (const ratio of [0.25, 0.5, 0.75, 1]) {
    if (chart.options?.radar?.shape === 'circle') {
      const ring = chartArc(center, radius * ratio, 0, Math.PI * 2)
      for (let index = 1; index < ring.length; index++) chartLine(commands, ring[index - 1]!, ring[index]!, chart.yAxis.gridLine, theme.splitLineColor ?? '#e2e8f0')
    } else data.categories.forEach((_category, index) => chartLine(commands, vertex(index, ratio), vertex((index + 1) % count, ratio), chart.yAxis.gridLine, theme.splitLineColor ?? '#e2e8f0'))
  }
  data.categories.forEach((category, index) => { chartLine(commands, center, vertex(index, 1), chart.xAxis.line, theme.axisColor ?? '#cbd5e1'); if (chart.xAxis.label?.visible !== false) commands.push({ kind: 'text', ...vertex(index, 1.22), text: category, font: chartFont(chart.xAxis.label, 10), color: chart.xAxis.label?.color ?? color, align: 'center', baseline: 'middle' }) })
  for (const [index, series] of layout.series.entries()) {
    const style = chartSeriesStyle(chart, data.series[index]!.field, index, palette), points = new Map(series.points.map(point => [point.categoryIndex, point]))
    const polygon = data.categories.map((_category, categoryIndex) => points.get(categoryIndex) ?? center)
    if (chart.options?.radar?.fill) chartPolygon(commands, polygon, chartColorOpacity(style.color ?? series.color, style.fillOpacity === 1 ? 0.25 : style.fillOpacity ?? 0.25))
    polygon.forEach((point, categoryIndex) => chartLine(commands, point, polygon[(categoryIndex + 1) % polygon.length]!, { color: series.color, width: 2, ...style.border }, series.color))
    series.points.forEach(point => chartSymbol(commands, point, style.point?.size ?? 4, style.point?.color ?? series.color, style.point?.shape ?? 'rect'))
  }
  if (chart.xAxis.title) commands.push({ kind: 'text', x: plot.x + plot.width / 2, y: plot.y + plot.height + 24, text: chart.xAxis.title, font: chartFont({}, 11), color, align: 'center', baseline: 'middle' })
  if (chart.yAxis.title) commands.push({ kind: 'text', x: layout.frame.x + 13, y: plot.y + plot.height / 2, text: chart.yAxis.title, font: chartFont({}, 11), color, rotation: -90, align: 'center', baseline: 'middle' })
}

export function drawChartSector(commands: DrawCommand[], center: ChartVector, inner: number, outer: number, start: number, end: number, color: string, border?: ChartSeriesOptions['border'], clip?: RenderRect): void {
  const count = Math.max(2, Math.ceil(Math.abs(end - start) / (Math.PI / 48)))
  for (let index = 0; index < count; index++) {
    const a = start + (end - start) * index / count, b = start + (end - start) * (index + 1) / count
    chartPolygon(commands, inner > 0 ? [chartPolar(center, inner, a), chartPolar(center, outer, a), chartPolar(center, outer, b), chartPolar(center, inner, b)] : [center, chartPolar(center, outer, a), chartPolar(center, outer, b)], color, undefined, clip)
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
    const style = chartPointStyle(chartSeriesStyle(chart, field, sector.part.index, palette), sector.part.index)
    drawChartSector(commands, sector.center, sector.inner, sector.outer, sector.start, sector.end, chartColorOpacity(style.color ?? palette[sector.part.index % palette.length]!, style.fillOpacity ?? 1), style.border, layout.plot)
    const label = { ...style.label, ...chart.options?.pie?.label }
    if (label.visible !== false) {
      const outside = label.position !== 'inside' && label.position !== 'center', point = chartPolar(sector.center, outside ? sector.outer * 1.12 : (sector.inner + sector.outer) / 2, (sector.start + sector.end) / 2)
      if (outside && chart.options?.pie?.showLabelLine !== false) chartLine(commands, chartPolar(sector.center, sector.outer, (sector.start + sector.end) / 2), point, { color, width: 0.8 }, color)
      commands.push({ kind: 'text', ...point, text: chartLabelText({ contentType: 'category', ...label }, sector.part.value, sector.part.category, data.series[0]!.name, sector.percentage), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot })
    }
  }
  for (const bar of pie.bars) {
    const style = chartPointStyle(chartSeriesStyle(chart, field, bar.part.index, palette), bar.part.index)
    commands.push({ kind: 'rect', ...bar.rect, fill: chartColorOpacity(style.color ?? palette[bar.part.index % palette.length]!, style.fillOpacity ?? 1), stroke: style.border?.visible === false ? undefined : style.border?.color, strokeWidth: style.border?.width, clip: layout.plot })
    const label = { ...style.label, ...chart.options?.pie?.label }
    if (label.visible !== false) commands.push({ kind: 'text', x: bar.rect.x + bar.rect.width / 2, y: bar.rect.y + bar.rect.height / 2, text: chartLabelText({ contentType: 'category', ...label }, bar.part.value, bar.part.category, data.series[0]!.name, bar.percentage), font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', maxWidth: bar.rect.width, clip: layout.plot })
  }
}

function drawBoxplot(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, _color: string, palette: string[]): void {
  const { plot } = layout, boxes = derived.boxes!, step = plot.width / Math.max(1, boxes.length), width = step * 0.48
  const y = (value: number): number => plot.y + (1 - chartAxisRatio(value, layout.min, layout.max, chart.yAxis)) * plot.height
  boxes.forEach((box, index) => {
    if (!box.statistics) return
    const style = chartSeriesStyle(chart, box.field, index, palette), statistics = box.statistics, color = style.color ?? palette[index % palette.length]!, x = plot.x + (index + 0.5) * step
    chartLine(commands, { x, y: y(statistics.min) }, { x, y: y(statistics.max) }, { color, width: 1.5, ...style.border }, color, plot)
    for (const value of [statistics.min, statistics.max]) chartLine(commands, { x: x - width / 3, y: y(value) }, { x: x + width / 3, y: y(value) }, { color, width: 1.5, ...style.border }, color, plot)
    commands.push({ kind: 'rect', x: x - width / 2, y: Math.min(y(statistics.q1), y(statistics.q3)), width, height: Math.abs(y(statistics.q1) - y(statistics.q3)), fill: chartColorOpacity(color, style.fillOpacity === 1 ? 0.45 : style.fillOpacity ?? 0.45), stroke: color, strokeWidth: style.border?.width ?? 1.5, clip: plot })
    chartLine(commands, { x: x - width / 2, y: y(statistics.median) }, { x: x + width / 2, y: y(statistics.median) }, { color, width: 2 }, color, plot)
    if (chart.options?.boxplot?.showOutliers !== false) statistics.outliers.forEach(value => chartCircle(commands, { x, y: y(value) }, 2, color, undefined, plot))
  })
}

function drawCandlestick(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout): void {
  const { plot } = layout, step = plot.width / Math.max(1, data.categories.length), width = Math.min(step * 0.8, chart.options?.candlestick?.candleWidth ?? step * 0.5)
  const y = (value: number): number => plot.y + (1 - chartAxisRatio(value, layout.min, layout.max, chart.yAxis)) * plot.height
  for (const candle of data.candles ?? []) {
    if (!candle) continue
    const index = candle.categoryIndex, ratio = chart.xAxis.reverse ? 1 - (index + 0.5) / Math.max(1, data.categories.length) : (index + 0.5) / Math.max(1, data.categories.length), x = plot.x + ratio * plot.width
    const rising = candle.open === undefined ? true : candle.close >= candle.open, color = rising ? chart.options?.candlestick?.risingColor ?? '#dc2626' : chart.options?.candlestick?.fallingColor ?? '#059669'
    chartLine(commands, { x, y: y(candle.high) }, { x, y: y(candle.low) }, { color, width: 1.5 }, color, plot)
    if (candle.open === undefined) chartLine(commands, { x, y: y(candle.close) }, { x: x + width / 2, y: y(candle.close) }, { color, width: 2 }, color, plot)
    else commands.push({ kind: 'rect', x: x - width / 2, y: Math.min(y(candle.open), y(candle.close)), width, height: Math.max(1, Math.abs(y(candle.close) - y(candle.open))), fill: chart.options?.candlestick?.hollow && rising ? chart.options?.appearance?.background ?? 'transparent' : color, stroke: color, strokeWidth: 1, clip: plot })
  }
}

function drawWaterfall(commands: DrawCommand[], derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const { plot } = layout, step = plot.width / Math.max(1, layout.categories.length), width = step * 0.65 / Math.max(1, layout.series.length), options = chart.options?.waterfall
  layout.series.forEach((series, seriesIndex) => series.points.forEach((point, index) => {
    const role = derived.waterfalls![seriesIndex]![index]!.role, style = chartSeriesStyle(chart, derived.data.series[seriesIndex]!.field, seriesIndex, palette)
    const color = role === 'total' ? options?.totalColor ?? palette[seriesIndex % palette.length]! : point.value >= 0 ? options?.positiveColor ?? '#059669' : options?.negativeColor ?? '#dc2626'
    const x = point.x - step * 0.325 + seriesIndex * width
    commands.push({ kind: 'rect', x, y: Math.min(point.y, point.baseY!), width: Math.max(1, width - 1), height: Math.abs(point.y - point.baseY!), fill: chartColorOpacity(color, style.fillOpacity ?? 1), stroke: style.border?.color, strokeWidth: style.border?.width, clip: plot })
    const next = series.points[index + 1]
    if (next && options?.connector !== false) chartLine(commands, { x: x + width, y: point.y }, { x: next.x - step * 0.325 + seriesIndex * width, y: point.y }, { color: '#94a3b8', dashType: 'dashed', width: 1 }, '#94a3b8', plot)
  }))
}

function drawHeatmap(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const values = data.series.flatMap(series => series.values.filter((value): value is number => value !== null)), min = chart.options?.heatmap?.min ?? Math.min(0, ...values), max = chart.options?.heatmap?.max ?? Math.max(1, ...values)
  const width = layout.plot.width / Math.max(1, data.categories.length), height = layout.plot.height / Math.max(1, data.series.length)
  data.series.forEach((series, row) => series.values.forEach((value, col) => {
    if (value === null) return
    const ratio = Math.max(0, Math.min(1, (value - min) / Math.max(1e-9, max - min))), colorIndex = Math.min(palette.length - 1, Math.floor(ratio * palette.length))
    const fill = chart.options?.heatmap?.visualMapType === 'piecewise' ? palette[colorIndex]! : chartShade(palette[Math.min(palette.length - 1, Math.floor(ratio * (palette.length - 1)))]!, 0.5 + ratio * 0.7)
    const x = layout.plot.x + (chart.xAxis.reverse ? data.categories.length - col - 1 : col) * width, y = layout.plot.y + (chart.yAxis.reverse ? data.series.length - row - 1 : row) * height
    commands.push({ kind: 'rect', x, y, width, height, fill, stroke: chart.options?.allSeriesStyle?.border?.color, strokeWidth: chart.options?.allSeriesStyle?.border?.width, clip: layout.plot })
    const label = chart.options?.heatmap?.label
    if (label?.visible) commands.push({ kind: 'text', x: x + width / 2, y: y + height / 2, text: `${formatChartNumber(value, label.format)}${chart.options?.heatmap?.unit ?? ''}`, font: chartFont(label, 10), color: label.color ?? color, align: 'center', baseline: 'middle', maxWidth: width - 4, clip: layout.plot })
  }))
}

function drawGraph(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const sankey = chart.type === 'sankey', graph = sankey ? chartSankeyLayout(data.graph!, layout.plot) : chartRelationLayout(data.graph!, chart, layout.plot)
  for (const link of graph.links) {
    const fill = chartColorOpacity(palette[link.colorIndex % palette.length]!, sankey ? 0.55 : 0.6)
    chartPolygon(commands, chartRibbon(link.points, link.width), fill, undefined, layout.plot)
    if (!sankey) {
      const end = link.points.at(-1)!, previous = link.points[Math.max(0, link.points.length - 2)]!, dx = end.x - previous.x, dy = end.y - previous.y, distance = Math.max(1e-9, Math.hypot(dx, dy))
      const target = graph.nodes.find(node => node.name === link.target)!, anchor = { x: end.x - dx / distance * target.width / 2, y: end.y - dy / distance * target.height / 2 }
      chartPolygon(commands, [anchor, { x: anchor.x - dx / distance * 6 - dy / distance * 3, y: anchor.y - dy / distance * 6 + dx / distance * 3 }, { x: anchor.x - dx / distance * 6 + dy / distance * 3, y: anchor.y - dy / distance * 6 - dx / distance * 3 }], fill, undefined, layout.plot)
    }
  }
  for (const node of graph.nodes) {
    const fill = palette[node.index % palette.length]!
    if (sankey) commands.push({ kind: 'rect', x: node.x, y: node.y, width: node.width, height: node.height, fill, clip: layout.plot })
    else chartCircle(commands, node, node.width / 2, fill, undefined, layout.plot)
    commands.push({ kind: 'text', x: node.x + (sankey ? node.width + 4 : 0), y: node.y + (sankey ? node.height / 2 : node.height / 2 + 10), text: node.name, font: chartFont(chart.options?.allSeriesStyle?.label, 10), color, align: sankey ? 'left' : 'center', baseline: 'middle', clip: layout.plot })
  }
}

function drawChord(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const chord = chartChordLayout(data.graph!, chart, layout.plot), label = chart.options?.chord?.label
  for (const link of chord.links) {
    const source = chartArc(chord.center, chord.innerRadius, link.sourceStart, link.sourceEnd), target = chartArc(chord.center, chord.innerRadius, link.targetStart, link.targetEnd)
    const first = chartBezier(source.at(-1)!, chord.center, chord.center, target[0]!), second = chartBezier(target.at(-1)!, chord.center, chord.center, source[0]!)
    chartPolygon(commands, [...source, ...first.slice(1), ...target.slice(1), ...second.slice(1)], chartColorOpacity(palette[link.sourceIndex % palette.length]!, chart.options?.chord?.ribbonOpacity ?? 0.45), undefined, layout.plot)
  }
  for (const node of chord.nodes) {
    drawChartSector(commands, chord.center, chord.innerRadius, chord.outerRadius, node.start, node.end, palette[node.index % palette.length]!, undefined, layout.plot)
    if (label?.visible !== false) commands.push({ kind: 'text', ...chartPolar(chord.center, chord.outerRadius + 10, (node.start + node.end) / 2), text: chartLabelText({ contentType: 'category', ...label }, node.value, node.name, data.series[0]!.name), font: chartFont(label, 10), color: label?.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot })
  }
}

function drawWordCloud(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, palette: string[]): void {
  const words = chartWordCloudLayout(data, chart.options?.wordCloud, layout.plot)
  for (const word of words) {
    const style = chartPointStyle(chartSeriesStyle(chart, data.series[0]!.field, word.categoryIndex, palette), word.categoryIndex)
    commands.push({ kind: 'text', x: word.x + word.width / 2, y: word.y + word.height / 2, text: word.text, font: chartFont({ ...chart.options?.appearance?.font, ...style.label, fontSize: word.fontSize }), color: style.color ?? palette[word.categoryIndex % palette.length]!, align: 'center', baseline: 'middle', rotation: word.rotation, clip: layout.plot })
  }
}

function drawTreemap(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const options = chart.options?.treemap, rectangles = chartTreemapLayout(data.hierarchy!, layout.plot, options?.parentLabelLayout !== 'overlapping')
  for (const rect of rectangles) commands.push({ kind: 'rect', x: rect.x, y: rect.y, width: rect.width, height: rect.height, fill: chartShade(palette[rect.colorIndex % palette.length]!, 1 - Math.min(0.4, rect.depth * 0.1)), stroke: chart.options?.appearance?.background ?? '#ffffff', strokeWidth: 1, clip: layout.plot })
  for (const rect of rectangles) {
    const parent = Boolean(rect.node.children?.length)
    if (options?.label?.visible === false || parent && options?.parentLabelLayout === 'none') continue
    const y = parent && options?.parentLabelLayout !== 'overlapping' ? rect.y + 10 : rect.y + rect.height / 2
    commands.push({ kind: 'text', x: rect.x + rect.width / 2, y, text: chartLabelText({ contentType: 'category', ...options?.label }, rect.node.value, rect.node.name, data.series[0]!.name), font: chartFont(options?.label, 10), color: options?.label?.color ?? color, align: 'center', baseline: 'middle', maxWidth: Math.max(1, rect.width - 6), clip: rect })
  }
}

function drawSunburst(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const sectors = chartSunburstLayout(data.hierarchy!, layout.plot), center = { x: layout.plot.x + layout.plot.width / 2, y: layout.plot.y + layout.plot.height / 2 }, label = chart.options?.sunburst?.label
  const total = data.hierarchy!.reduce((sum, node) => sum + node.value, 0)
  for (const sector of sectors) {
    drawChartSector(commands, center, sector.inner, sector.outer, sector.start, sector.end, chartShade(palette[sector.colorIndex % palette.length]!, 1 - sector.depth * 0.08), { color: chart.options?.appearance?.background ?? '#ffffff', width: 0.7 }, layout.plot)
    if (label?.visible !== false && sector.end - sector.start > 0.06) commands.push({ kind: 'text', ...chartPolar(center, (sector.inner + sector.outer) / 2, (sector.start + sector.end) / 2), text: chartLabelText({ contentType: 'category', ...label }, sector.node.value, sector.node.name, data.series[0]!.name, total ? sector.node.value * 100 / total : 0), font: chartFont(label, 10), color: label?.color ?? color, align: 'center', baseline: 'middle', maxWidth: Math.max(1, (sector.end - sector.start) * (sector.inner + sector.outer) / 2), clip: layout.plot })
  }
  if (chart.options?.sunburst?.centerLabel?.visible) commands.push({ kind: 'text', ...center, text: data.series[0]!.name, font: chartFont(chart.options.sunburst.centerLabel.font, 11), color: chart.options.sunburst.centerLabel.font?.color ?? color, align: 'center', baseline: 'middle', maxWidth: sectors[0] ? sectors[0].inner * 1.5 : layout.plot.width, clip: layout.plot })
}

function drawGauge(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[], theme: ChartRenderTheme): void {
  const options = chart.options?.gauge, min = options?.min ?? 0, max = options?.max ?? 100, start = (options?.startAngle ?? 225) * Math.PI / 180, end = (options?.endAngle ?? -45) * Math.PI / 180
  for (const tile of chartGaugeTiles(data, layout.plot)) {
    const value = data.series[0]!.values[tile.categoryIndex]
    if (value === null || value === undefined) continue
    const stroke = tile.radius * 0.1, angle = start + (end - start) * Math.max(0, Math.min(1, (value - min) / (max - min)))
    drawChartSector(commands, tile.center, tile.radius - stroke, tile.radius, -start, -end, theme.splitLineColor ?? '#e2e8f0', undefined, layout.plot)
    for (const range of options?.ranges ?? []) drawChartSector(commands, tile.center, tile.radius - stroke, tile.radius, -(start + (end - start) * Math.max(0, (range.from - min) / (max - min))), -(start + (end - start) * Math.min(1, (range.to - min) / (max - min))), range.color, undefined, layout.plot)
    const progress = options?.mode === 'progress' || options?.progress?.visible
    if (progress) {
      const width = options?.progress?.width ?? stroke
      drawChartSector(commands, tile.center, Math.max(0, tile.radius - width), tile.radius, -start, -angle, options?.progress?.color ?? palette[tile.categoryIndex % palette.length]!, undefined, layout.plot)
      if (options?.progress?.rounded) for (const mark of [-start, -angle]) chartCircle(commands, chartPolar(tile.center, tile.radius - width / 2, mark), width / 2, options.progress.color ?? palette[tile.categoryIndex % palette.length]!, undefined, layout.plot)
    }
    if (options?.mode !== 'progress' && options?.pointer?.visible !== false) {
      const tip = chartPolar(tile.center, tile.radius * (options?.pointer?.length ?? 0.7), -angle), width = options?.pointer?.width ?? 4
      const sideA = chartPolar(tile.center, width, -angle + Math.PI / 2), sideB = chartPolar(tile.center, width, -angle - Math.PI / 2)
      chartPolygon(commands, [sideA, tip, sideB], options?.pointer?.color ?? palette[tile.categoryIndex % palette.length]!, undefined, layout.plot)
      chartCircle(commands, tile.center, width * 1.2, options?.pointer?.color ?? palette[tile.categoryIndex % palette.length]!, undefined, layout.plot)
    }
    const divisions = options?.splitNumber ?? 5
    for (let index = 0; index <= divisions; index++) {
      const tickAngle = -(start + (end - start) * index / divisions)
      chartLine(commands, chartPolar(tile.center, tile.radius * 0.86, tickAngle), chartPolar(tile.center, tile.radius * 0.95, tickAngle), { color, width: 1 }, color, layout.plot)
      commands.push({ kind: 'text', ...chartPolar(tile.center, tile.radius * 0.72, tickAngle), text: formatChartNumber(min + (max - min) * index / divisions), font: chartFont({}, 8), color, align: 'center', baseline: 'middle', clip: layout.plot })
    }
    commands.push({ kind: 'text', x: tile.center.x, y: tile.center.y + tile.radius * 0.56, text: data.categories[tile.categoryIndex]!, font: chartFont(options?.label, 10), color: options?.label?.color ?? color, align: 'center', baseline: 'middle', maxWidth: tile.radius * 1.7, clip: layout.plot })
    if (options?.label?.visible !== false) commands.push({ kind: 'text', x: tile.center.x, y: tile.center.y + tile.radius * 0.3, text: formatChartNumber(value, options?.label?.format), font: chartFont({ bold: true, ...options?.label }, 13), color: options?.label?.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot })
  }
}

function drawFunnel(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  const source = data.series[0]!, options = chart.options?.funnel, parts = data.categories.map((category, index) => ({ category, index, value: options?.useAbsoluteValue ? Math.abs(source.values[index] ?? 0) : source.values[index] ?? 0 }))
  if (options?.sort !== 'none') parts.sort((a, b) => options?.sort === 'ascending' ? a.value - b.value || a.index - b.index : b.value - a.value || a.index - b.index)
  let max = 1
  for (const part of parts) max = Math.max(max, part.value)
  const height = layout.plot.height / Math.max(1, parts.length), gap = Math.max(0, options?.gap ?? 2)
  parts.forEach((part, index) => {
    const width = Math.max(0, part.value / max * layout.plot.width), next = parts[index + 1], nextWidth = next ? Math.max(0, next.value / max * layout.plot.width) : 0
    const cx = layout.plot.x + layout.plot.width / 2, top = layout.plot.y + index * height, bottom = top + Math.max(0, height - gap)
    const style = chartPointStyle(chartSeriesStyle(chart, source.field, part.index, palette), part.index)
    chartPolygon(commands, [{ x: cx - width / 2, y: top }, { x: cx + width / 2, y: top }, { x: cx + nextWidth / 2, y: bottom }, { x: cx - nextWidth / 2, y: bottom }], chartColorOpacity(style.color ?? palette[part.index % palette.length]!, style.fillOpacity ?? 1), style.border, layout.plot)
    if (style.label?.visible !== false) commands.push({ kind: 'text', x: cx, y: (top + bottom) / 2, text: chartLabelText({ contentType: 'category', ...style.label }, part.value, part.category, source.name), font: chartFont(style.label, 10), color: style.label?.color ?? color, align: 'center', baseline: 'middle', clip: layout.plot })
  })
}

function drawLegend(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout, color: string, palette: string[]): void {
  if (!chart.legend.visible) return
  const family = resolveChartType(chart.type).family
  const categoryLegend = ['pie', 'funnel', 'gauge', 'chord', 'relation', 'sankey', 'treemap', 'sunburst'].includes(family)
  const names = categoryLegend ? data.categories : data.series.map((series, index) => chartSeriesStyle(chart, series.field, index, palette).name ?? series.name)
  const font: ChartFontOptions = { ...chart.options?.appearance?.font, ...chart.legend.font }, fontSize = font.fontSize ?? 11
  const horizontal = chart.legend.position === 'top' || chart.legend.position === 'bottom', itemWidth = Math.max(72, Math.min(140, layout.frame.width / Math.max(1, names.length)))
  const perRow = Math.max(1, Math.floor((layout.frame.width - 24) / itemWidth))
  names.forEach((name, index) => {
    const x = horizontal ? layout.frame.x + 12 + index % perRow * itemWidth : chart.legend.position === 'left' ? layout.frame.x + 8 : layout.plot.x + layout.plot.width + 12
    const y = horizontal ? chart.legend.position === 'top' ? layout.plot.y - 14 + Math.floor(index / perRow) * (fontSize + 4) : layout.frame.y + layout.frame.height - 12 - Math.floor(index / perRow) * (fontSize + 4)
      : layout.plot.y + 10 + index * (fontSize + 6)
    commands.push({ kind: 'rect', x, y: y - 5, width: 10, height: 10, fill: categoryLegend ? palette[index % palette.length]! : chartSeriesStyle(chart, data.series[index]!.field, index, palette).color ?? palette[index % palette.length]! })
    commands.push({ kind: 'text', x: x + 15, y, text: name, font: chartFont(font, 11), color: font.color ?? color, baseline: 'middle', maxWidth: horizontal ? itemWidth - 20 : Math.max(1, layout.frame.width * 0.24 - 22), clip: layout.frame })
  })
}
