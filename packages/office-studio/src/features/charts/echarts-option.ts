import type {
  BarSeriesOption, BoxplotSeriesOption, CandlestickSeriesOption, FunnelSeriesOption, GaugeSeriesOption, GraphSeriesOption,
  HeatmapSeriesOption, LineSeriesOption, PieSeriesOption, RadarSeriesOption, SankeySeriesOption, ScatterSeriesOption, SunburstSeriesOption, TreemapSeriesOption,
} from 'echarts/charts'
import type {
  GraphicComponentOption, GridComponentOption, LegendComponentOption, RadarComponentOption, TitleComponentOption, TooltipComponentOption, VisualMapComponentOption,
} from 'echarts/components'
import type { ComposeOption } from 'echarts/core'
import type { XAxisComponentOption, YAxisComponentOption } from 'echarts'
import type { ChartAxisOptions, ChartModel, WorkbookModel } from '../../domain/model'
import type { DrawCommand } from '../../rendering/contracts'
import { radarValueExtent, readChartData, type ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartPointStyle, chartSeriesKind, deriveChartData, type ChartComputedStatistics } from './chart-derived'
import { chartPalette, chartSeriesStyle } from './chart-options'
import { chartEChartsAxisStyle, chartEChartsFont, chartEChartsItemStyle, chartEChartsLabel, chartEChartsLineStyle, chartEChartsPointLabel } from './echarts-style'
import { chartSpecialEChartsOption, type ChartEChartsContext } from './echarts-special-option'
import { drawDerivedChartData } from './chart-draw'
import { chartPlotFrame } from './chart-layout'
import type { ChartRenderTheme } from './chart-theme'

export type ChartEChartsOption = ComposeOption<
  | BarSeriesOption | LineSeriesOption | PieSeriesOption | RadarSeriesOption | ScatterSeriesOption | BoxplotSeriesOption
  | CandlestickSeriesOption | FunnelSeriesOption | GaugeSeriesOption | GraphSeriesOption | HeatmapSeriesOption | SankeySeriesOption | SunburstSeriesOption | TreemapSeriesOption
  | GraphicComponentOption | GridComponentOption | LegendComponentOption | RadarComponentOption | TitleComponentOption | TooltipComponentOption
  | VisualMapComponentOption | XAxisComponentOption | YAxisComponentOption
> & {
  /** Authored dependencies used by the bitmap cache, including formatter closures. Removed before ECharts.setOption. */
  officeChartKey?: string
}

export type ChartEChartsTheme = ChartRenderTheme

/** Converts authored semantics and derived statistics to a Canvas option without mutating the document. */
export function chartEChartsOption(workbook: WorkbookModel, chart: ChartModel, theme: ChartEChartsTheme = {}, computed?: ChartComputedStatistics): ChartEChartsOption {
  const source = readChartData(workbook, chart), derived = deriveChartData(source, chart, computed), data = derived.data
  const descriptor = resolveChartType(chart.type), color = chart.options?.appearance?.font?.color ?? theme.textColor ?? '#334155'
  const axisColor = theme.axisColor ?? '#94a3b8', splitLineColor = theme.splitLineColor ?? '#e2e8f0', palette = chartPalette(chart, theme.palette)
  const officeChartKey = JSON.stringify([chart, theme])
  const localChart = { ...chart, position: { x: 0, y: 0 } }
  const frame = chartPlotFrame(localChart), grid = {
    left: frame.plot.x, right: Math.max(0, chart.size.width - frame.plot.x - frame.plot.width),
    top: frame.plot.y, bottom: Math.max(0, chart.size.height - frame.plot.y - frame.plot.height), containLabel: true,
  }
  if (requiresGraphicChart(chart, data)) {
    return {
      animation: false, backgroundColor: chart.options?.appearance?.background ?? theme.backgroundColor ?? 'transparent',
      graphic: { elements: chartGraphics(drawDerivedChartData(derived, localChart, theme, palette), chart.id) }, officeChartKey,
    }
  }
  const isPie = descriptor.family === 'pie', isScatter = chart.type === 'scatter' || chart.type === 'bubble', isRadar = chart.type === 'radar', horizontal = descriptor.family === 'bar'
  const context: ChartEChartsContext = { palette, color, axisColor, splitLineColor, grid }
  const seriesNames = data.series.map((series, index) => chartSeriesStyle(chart, series.field, index, palette).name ?? series.name)
  const titleHeight = (chart.title ? 28 : 0) + (chart.subtitle ? 18 : 0), title: TitleComponentOption[] = []
  if (chart.title || chart.subtitle) title.push({ text: chart.title ?? '', subtext: chart.subtitle, left: 'center', top: 8,
    textStyle: chartEChartsFont({ bold: true, ...chart.options?.titleStyle }, color, 14),
    subtextStyle: chartEChartsFont(chart.options?.subtitleStyle, color, 11) })
  if (isRadar && chart.xAxis.title) title.push({ text: chart.xAxis.title, left: 'center', bottom: grid.bottom - 30, textStyle: chartEChartsFont(undefined, color, 11) })
  if (isRadar && chart.yAxis.title) title.push({ text: chart.yAxis.title, left: grid.left - 44, top: grid.top, textStyle: chartEChartsFont(undefined, color, 11) })
  const categoryLegend = ['pie', 'funnel', 'gauge', 'treemap', 'sunburst'].includes(descriptor.family)
  const right = data.series.some((series, index) => chartSeriesStyle(chart, series.field, index, palette).rightAxis)
  const valueAxis = numericAxis(chart.yAxis, axisColor, splitLineColor, horizontal ? 'x' : 'y', descriptor.percent ?? false)
  const secondAxis: ChartAxisOptions = chart.options?.rightYAxis ?? { scaleType: 'linear' }
  const option: ChartEChartsOption = {
    animation: false, backgroundColor: chart.options?.appearance?.background ?? theme.backgroundColor ?? 'transparent',
    color: palette, textStyle: chartEChartsFont(chart.options?.appearance?.font, color),
    ...(title.length ? { title: title.length === 1 ? title[0] : title } : {}),
    legend: { show: chart.legend.visible, data: categoryLegend ? data.categories : seriesNames, ...legendPosition(chart.legend.position, titleHeight),
      textStyle: chartEChartsFont(chart.legend.font, color, 11) },
    tooltip: { trigger: isPie || isScatter || isRadar || descriptor.dataShape === 'hierarchy' ? 'item' : 'axis' },
    ...(isPie || isRadar ? {} : {
      grid,
      xAxis: horizontal ? valueAxis : isScatter ? numericAxis(chart.xAxis, axisColor, splitLineColor, 'x') : categoryAxis(data.categories, chart.xAxis, axisColor, splitLineColor),
      yAxis: horizontal ? categoryAxis(data.categories, chart.xAxis, axisColor, splitLineColor)
        : right ? [valueAxis, { ...numericAxis(secondAxis, axisColor, splitLineColor, 'y'), position: 'right', splitLine: { show: false } }] : valueAxis,
    }),
    ...(isRadar ? { radar: radarComponent(data, chart, context) } : {}),
    series: isPie ? [pieSeries(data, chart, context)] : isScatter ? scatterSeries(data, chart, context)
      : isRadar ? radarSeries(data, chart, context) : categorySeries(data, chart, context),
    officeChartKey,
  }
  const special = chartSpecialEChartsOption(data, derived, chart, context)
  if (special) {
    const previousLegend = option.legend
    Object.assign(option, special)
    if (special.legend && previousLegend && !Array.isArray(previousLegend) && !Array.isArray(special.legend)) option.legend = { ...previousLegend, ...special.legend }
  }
  if (derived.pareto) {
    const cumulative: LineSeriesOption = { type: 'line', name: '累计百分比', yAxisIndex: 1, data: derived.pareto.cumulativePercent, showSymbol: true, lineStyle: { color: palette[1 % palette.length] } }
    option.series = [...categorySeries(data, chart, context), cumulative]
    option.yAxis = [numericAxis(chart.yAxis, axisColor, splitLineColor, 'y'), { ...numericAxis(secondAxis, axisColor, splitLineColor, 'y'), min: secondAxis.min ?? 0, max: secondAxis.max ?? 100,
      position: 'right', axisLabel: { ...chartEChartsFont(secondAxis.label, axisColor, 10), formatter: '{value}%' }, splitLine: { show: false } }]
    option.legend = { ...(!Array.isArray(option.legend) ? option.legend : {}), data: [...seriesNames, '累计百分比'] }
  }
  if (chart.options?.appearance?.border?.visible) {
    const border = chart.options.appearance.border, graphic: NonNullable<GraphicComponentOption['elements']>[number] = { type: 'rect',
      shape: { x: 0.5, y: 0.5, width: chart.size.width - 1, height: chart.size.height - 1 }, style: { fill: 'transparent', stroke: border.color ?? axisColor, lineWidth: border.width ?? 1, opacity: border.opacity ?? 1 }, silent: true, z: -10 }
    const current = option.graphic
    option.graphic = { elements: [graphic, ...(!Array.isArray(current) ? current?.elements ?? [] : [])] }
  }
  return option
}

function requiresGraphicChart(chart: ChartModel, data: ChartData): boolean {
  return Boolean(resolveChartType(chart.type).threeD || ['pieOfPie', 'barOfPie', 'wordCloud', 'chord', 'waterfall', 'sankey', 'relation'].includes(chart.type)
    || chart.type === 'candlestick' && data.candles?.some(candle => candle !== null && candle.open === undefined)
    || chart.options?.trendlines?.length || chart.options?.errorBars?.length)
}

export function chartGraphics(commands: DrawCommand[], chartId: string): NonNullable<GraphicComponentOption['elements']> {
  return commands.map((command, index): NonNullable<GraphicComponentOption['elements']>[number] => {
    const id = `${chartId}-${index}`, common = { id, silent: true, z: index }
    let element: NonNullable<GraphicComponentOption['elements']>[number]
    if (command.kind === 'rect') element = { ...common, type: 'rect', shape: { x: command.x, y: command.y, width: command.width, height: command.height }, style: { fill: command.fill, stroke: command.stroke, lineWidth: command.strokeWidth ?? 1 } }
    else if (command.kind === 'line') element = { ...common, type: 'line', shape: { x1: command.x1, y1: command.y1, x2: command.x2, y2: command.y2 }, style: { stroke: command.color, lineWidth: command.width ?? 1 } }
    else if (command.kind === 'polygon') element = { ...common, type: 'polygon', shape: { points: command.points.map(point => [point.x, point.y]) }, style: { fill: command.fill } }
    else if (command.kind === 'text') element = { ...common, type: 'text', x: command.x, y: command.y, rotation: -(command.rotation ?? 0) * Math.PI / 180,
      style: { text: command.text, font: command.font, fill: command.color, align: command.align ?? 'left', verticalAlign: command.baseline === 'middle' ? 'middle' : command.baseline === 'bottom' ? 'bottom' : 'top',
        ...(command.maxWidth === undefined ? {} : { width: command.maxWidth, overflow: 'truncate' as const }) } }
    else element = { ...common, type: 'image', style: { image: command.source, x: command.x, y: command.y, width: command.width, height: command.height, opacity: command.opacity ?? 1 } }
    return command.clip ? { type: 'group', id: `${id}-clip`, silent: true, children: [element],
      clipPath: { type: 'rect', shape: { x: command.clip.x, y: command.clip.y, width: command.clip.width, height: command.clip.height } } } : element
  })
}

function radarComponent(data: ChartData, chart: ChartModel, context: ChartEChartsContext): RadarComponentOption {
  const extent = radarValueExtent(data), { grid, color, axisColor, splitLineColor } = context
  return { center: [(chart.size.width + grid.left - grid.right) / 2, (chart.size.height + grid.top - grid.bottom) / 2],
    radius: Math.max(1, Math.min(chart.size.width - grid.left - grid.right, chart.size.height - grid.top - grid.bottom) * 0.37),
    clockwise: true, shape: chart.options?.radar?.shape ?? 'polygon', indicator: data.categories.map(name => ({ name, ...extent })),
    axisName: { show: chart.xAxis.label?.visible !== false, ...chartEChartsFont(chart.xAxis.label, color) },
    axisLine: { show: chart.xAxis.line?.visible !== false, lineStyle: { color: chart.xAxis.line?.color ?? axisColor } },
    axisLabel: { show: chart.yAxis.label?.visible !== false, ...chartEChartsFont(chart.yAxis.label, axisColor) },
    splitLine: { show: chart.yAxis.gridLine?.visible !== false, lineStyle: { color: chart.yAxis.gridLine?.color ?? splitLineColor, type: chart.yAxis.gridLine?.dashType ?? 'solid' } }, splitArea: { show: false } }
}

function radarSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): RadarSeriesOption[] {
  return data.series.map((source, index) => {
    const style = chartSeriesStyle(chart, source.field, index, context.palette)
    return { type: 'radar', name: style.name ?? source.name, symbol: style.point?.shape ?? 'none', symbolSize: style.point?.size ?? 4,
      lineStyle: chartEChartsLineStyle(style), itemStyle: chartEChartsItemStyle(style),
      ...(chart.options?.radar?.fill ? { areaStyle: { color: style.color, opacity: style.fillOpacity === 1 ? 0.25 : style.fillOpacity ?? 0.25 } } : {}),
      data: [{ name: style.name ?? source.name, value: source.values }] }
  })
}

function legendPosition(position: ChartModel['legend']['position'], titleHeight: number): Pick<LegendComponentOption, 'top' | 'right' | 'bottom' | 'left' | 'orient'> {
  switch (position) {
    case 'top': return { top: 8 + titleHeight, left: 'center', orient: 'horizontal' }
    case 'right': return { right: 8, top: 'middle', orient: 'vertical' }
    case 'left': return { left: 8, top: 'middle', orient: 'vertical' }
    case 'bottom': return { bottom: 8, left: 'center', orient: 'horizontal' }
  }
}

function categoryAxis(categories: string[], axis: ChartAxisOptions, axisColor: string, splitLineColor: string): XAxisComponentOption {
  return { type: 'category', data: categories, ...chartEChartsAxisStyle(axis, axisColor, splitLineColor, false), nameGap: 28 }
}

function numericAxis(axis: ChartAxisOptions, axisColor: string, splitLineColor: string, dimension: 'x' | 'y', percent = false): XAxisComponentOption & YAxisComponentOption {
  const style = chartEChartsAxisStyle(axis, axisColor, splitLineColor, true)
  return { type: axis.scaleType === 'log' ? 'log' : 'value', ...(axis.scaleType === 'log' ? { logBase: axis.logBase ?? 10 } : {}),
    ...style, min: axis.min, max: axis.max, nameGap: dimension === 'x' ? 28 : 42,
    ...(percent && !axis.format ? { axisLabel: { ...style.axisLabel, formatter: '{value}%' } } : {}),
    ...(dimension === 'x' && axis.gridLine?.visible === undefined ? { splitLine: { show: false } } : {}) }
}

function pieSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): PieSeriesOption {
  const source = data.series[0]!, options = chart.options?.pie, { palette, color, grid } = context
  const total = source.values.reduce<number>((sum, value) => sum + Math.max(0, value ?? 0), 0), start = -(options?.startAngle ?? -90)
  return { type: 'pie', name: source.name, radius: chart.type === 'doughnut' ? options?.doughnutHole === undefined ? ['52%', '74%'] : [`${options.doughnutHole * 74}%`, '74%'] : '74%',
    center: ['50%', '50%'], left: grid.left, right: grid.right, top: grid.top, bottom: Math.max(8, grid.bottom - 12),
    startAngle: start, ...(options?.isHalfPie ? { endAngle: start - 180 } : {}), ...(options?.rosePie ? { roseType: 'radius' } : {}),
    ...(options?.explosion ? { selectedMode: 'multiple', selectedOffset: Math.min(chart.size.width, chart.size.height) * 0.37 * options.explosion / 100 } : {}),
    avoidLabelOverlap: true, stillShowZeroSum: false, showEmptyCircle: false, label: { color }, labelLine: { show: options?.showLabelLine !== false },
    data: data.categories.map((name, index) => {
      const value = Math.max(0, source.values[index] ?? 0), style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette), index), label = { ...style.label, ...options?.label }
      return { name, value, selected: Boolean(options?.explosion), itemStyle: chartEChartsItemStyle(style),
        label: chartEChartsPointLabel({ visible: true, contentType: 'category', position: 'outside', ...label }, color, value, name, source.name, total ? value * 100 / total : 0) }
    }) }
}

function scatterSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): ScatterSeriesOption[] {
  let maxSize = 1
  for (const source of data.series) for (const size of source.sizeValues ?? []) if (size !== null) maxSize = Math.max(maxSize, size)
  return data.series.map((source, seriesIndex) => {
    const style = chartSeriesStyle(chart, source.field, seriesIndex, context.palette), name = style.name ?? source.name, bubble = chart.type === 'bubble'
    const axis: ChartAxisOptions = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    return { type: 'scatter', name, xAxisIndex: 0, yAxisIndex: style.rightAxis ? 1 : 0, symbol: bubble ? 'circle' : style.point?.shape ?? 'circle', symbolSize: style.point?.size ?? 8,
      itemStyle: chartEChartsItemStyle(style), label: chartEChartsLabel(style.label, context.color),
      data: source.values.flatMap((value, index) => {
        const x = source.xValues?.[index], size = source.sizeValues?.[index]
        if (value === null || x == null || !Number.isFinite(value) || !Number.isFinite(x) || chart.xAxis.scaleType === 'log' && x <= 0 || axis.scaleType === 'log' && value <= 0 || bubble && (size == null || size <= 0)) return []
        const pointStyle = chartPointStyle(style, index)
        if (!bubble && !style.dataPoints?.[index] && !style.label?.visible) return [[x, value]]
        const diameter = bubble ? (chart.options?.bubble?.scale ?? 1) * Math.min(54, Math.min(chart.size.width - context.grid.left - context.grid.right, chart.size.height - context.grid.top - context.grid.bottom) * 0.2)
          * (chart.options?.bubble?.sizeRepresents === 'width' ? size! / maxSize : Math.sqrt(size! / maxSize)) : pointStyle.point?.size ?? style.point?.size ?? 8
        return [{ name: data.categories[index], value: [x, value, ...(bubble ? [size!] : [])], symbolSize: diameter, itemStyle: chartEChartsItemStyle(pointStyle),
          ...(pointStyle.label?.visible ? { label: chartEChartsPointLabel(pointStyle.label, context.color, value, data.categories[index] ?? '', name) } : {}) }]
      }) }
  })
}

function categorySeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): Array<BarSeriesOption | LineSeriesOption> {
  return data.series.map((source, index) => {
    const kind = chartSeriesKind(chart, source.field, index), style = chartSeriesStyle(chart, source.field, index, context.palette), name = style.name ?? source.name
    const axis: ChartAxisOptions = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const values = source.values.map((value, categoryIndex) => {
      if (value === null || axis.scaleType === 'log' && value <= 0) return null
      if (!style.dataPoints?.[categoryIndex] && !style.label?.visible) return value
      const point = chartPointStyle(style, categoryIndex)
      return { value, itemStyle: chartEChartsItemStyle(point), ...(point.label?.visible ? { label: chartEChartsPointLabel(point.label, context.color, value, data.categories[categoryIndex] ?? '', name) } : {}) }
    })
    const stacked = kind.stacked ? chart.type === 'combination' ? `${kind.type}-${style.rightAxis ? 'right' : 'left'}` : style.rightAxis ? 'total-right' : 'total' : undefined
    if (kind.type === 'bar' || kind.type === 'column' || chart.type === 'histogram' || chart.type === 'pareto') return {
      type: 'bar', name, data: values, stack: stacked, ...(style.rightAxis ? { yAxisIndex: 1 } : {}), itemStyle: chartEChartsItemStyle(style), label: chartEChartsLabel(style.label, context.color),
      barCategoryGap: chart.type === 'histogram' ? `${(chart.options?.histogram?.binGap ?? 0.05) * 100}%` : `${chart.options?.bar?.gapWidth ?? 150}%`,
      ...(chart.options?.bar?.overlap === undefined ? {} : { barGap: `${-chart.options.bar.overlap}%` }),
    }
    return { type: 'line', name, data: values, stack: stacked, yAxisIndex: style.rightAxis ? 1 : 0, showSymbol: style.point?.shape !== 'none',
      symbol: style.point?.shape ?? 'circle', symbolSize: style.point?.size ?? 6, connectNulls: chart.options?.invalidValueStrategy === 'link',
      smooth: kind.type === 'area' && chart.options?.area?.smooth, lineStyle: chartEChartsLineStyle(style), itemStyle: chartEChartsItemStyle(style), label: chartEChartsLabel(style.label, context.color),
      ...(kind.type === 'area' ? { areaStyle: { color: style.color, opacity: style.fillOpacity === 1 ? 0.28 : style.fillOpacity ?? 0.28 } } : {}) }
  })
}
