import type {
  BarSeriesOption, BoxplotSeriesOption, CandlestickSeriesOption, FunnelSeriesOption, GaugeSeriesOption, GraphSeriesOption,
  HeatmapSeriesOption, LineSeriesOption, PieSeriesOption, RadarSeriesOption, SankeySeriesOption, ScatterSeriesOption, SunburstSeriesOption, TreemapSeriesOption,
} from 'echarts/charts'
import type {
  AxisPointerComponentOption, GraphicComponentOption, GridComponentOption, LegendComponentOption, RadarComponentOption, TitleComponentOption, TooltipComponentOption, VisualMapComponentOption,
} from 'echarts/components'
import type { ComposeOption } from 'echarts/core'
import type { XAxisComponentOption, YAxisComponentOption } from 'echarts'
import type { ChartAxisOptions, ChartFontOptions, ChartModel, WorkbookModel } from '../../domain/model'
import type { DrawCommand } from '../../rendering/contracts'
import { readChartData, type ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartBubblePointSize, chartPointStyle, chartSeriesKind, deriveChartData, type ChartComputedStatistics } from './chart-derived'
import { chartPalette, chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { chartEChartsAxisStyle, chartEChartsFont, chartEChartsItemStyle, chartEChartsLabel, chartEChartsLineStyle, chartEChartsMarkerStyle, chartEChartsPointLabel } from './echarts-style'
import { chartSpecialEChartsOption, type ChartEChartsContext } from './echarts-special-option'
import { drawDerivedChartData } from './chart-draw'
import { chartPlotFrame, layoutDerivedChartData } from './chart-layout'
import { chartPieOuterRadius } from './chart-pie-layout'
import { chartRadarAxisLabel, chartRadarAxisLine, chartRadarExtent, chartRadarGeometry } from './chart-radar'
import { formatChartNumber } from './chart-geometry'
import { chartAxisTitleCommand, chartAxisTitleFont, chartAxisTitleSpace } from './chart-axis-title'
import { chartDateAxis, chartDateAxisLabel, type ChartDateAxis } from './chart-date-axis'
import { chartAxisPointerGeometry, chartAxisPointerPresentation, supportsChartAxisPointer, type ChartAxisPointerRuntime } from './chart-axis-pointer'
import type { ChartRenderTheme } from './chart-theme'

export type ChartEChartsOption = ComposeOption<
  | BarSeriesOption | LineSeriesOption | PieSeriesOption | RadarSeriesOption | ScatterSeriesOption | BoxplotSeriesOption
  | CandlestickSeriesOption | FunnelSeriesOption | GaugeSeriesOption | GraphSeriesOption | HeatmapSeriesOption | SankeySeriesOption | SunburstSeriesOption | TreemapSeriesOption
  | GraphicComponentOption | GridComponentOption | LegendComponentOption | RadarComponentOption | TitleComponentOption | TooltipComponentOption
  | AxisPointerComponentOption | VisualMapComponentOption | XAxisComponentOption | YAxisComponentOption
> & {
  /** Authored dependencies used by the bitmap cache, including formatter closures. Removed before ECharts.setOption. */
  officeChartKey?: string
  /** Transient surface interaction metadata. Removed before signature generation and setOption. */
  officeChartPointer?: ChartAxisPointerRuntime
}

export type ChartEChartsTheme = ChartRenderTheme

/** Converts authored semantics and derived statistics to a Canvas option without mutating the document. */
export function chartEChartsOption(workbook: WorkbookModel, chart: ChartModel, theme: ChartEChartsTheme = {}, computed?: ChartComputedStatistics): ChartEChartsOption {
  const source = computed?.preparedData ?? readChartData(workbook, chart), derived = deriveChartData(source, chart, computed), data = derived.data
  const descriptor = resolveChartType(chart.type), color = chart.options?.appearance?.font?.color ?? theme.textColor ?? '#334155'
  const axisColor = theme.axisColor ?? '#94a3b8', splitLineColor = theme.splitLineColor ?? '#e2e8f0', palette = chartPalette(chart, theme.palette)
  const localChart = { ...chart, position: { x: 0, y: 0 } }
  const officeChartKey = JSON.stringify([localChart, theme, Boolean(data.date1904)])
  const pointerEnabled = supportsChartAxisPointer(chart) && data.categories.length > 0 && data.series.length > 0
  const frame = chartPlotFrame(localChart, derived.boxes ?? data.series), grid = {
    left: frame.plot.x, right: Math.max(0, chart.size.width - frame.plot.x - frame.plot.width),
    top: frame.plot.y, bottom: Math.max(0, chart.size.height - frame.plot.y - frame.plot.height), containLabel: true,
  }
  if (!data.categories.length || !data.series.length || descriptor.family === 'gauge' && !data.series[0]!.values.some(Number.isFinite) || requiresGraphicChart(chart, data, derived.boxes)) {
    const layout = layoutDerivedChartData(derived, localChart, palette)
    const pointerRuntime: ChartAxisPointerRuntime | undefined = pointerEnabled
      ? { mode: 'graphic', geometry: chartAxisPointerGeometry(data, localChart, layout, theme) } : undefined
    return {
      animation: false, backgroundColor: chart.options?.appearance?.background ?? theme.backgroundColor ?? 'transparent',
      graphic: { elements: chartGraphics(drawDerivedChartData(derived, localChart, theme, palette, layout), chart.id) }, officeChartKey,
      ...(pointerRuntime ? { officeChartPointer: pointerRuntime } : {}),
    }
  }
  const isPie = descriptor.family === 'pie', isScatter = chart.type === 'scatter' || chart.type === 'bubble', isRadar = chart.type === 'radar', horizontal = descriptor.family === 'bar'
  const dates = chartDateAxis(data, chart), inheritedFont = chart.options?.appearance?.font
  const context: ChartEChartsContext = { palette, color, axisColor, splitLineColor, grid, dateAxis: dates, backgroundColor: chart.options?.appearance?.background ?? theme.backgroundColor }
  const seriesNames = data.series.map((series, index) => chartSeriesStyle(chart, series.field, index, palette, series.sourceName ?? series.name).name ?? series.name)
  const titleHeight = (chart.title ? Math.max(28, (chart.options?.titleStyle?.fontSize ?? inheritedFont?.fontSize ?? 14) + 14) : 0)
    + (chart.subtitle ? Math.max(18, (chart.options?.subtitleStyle?.fontSize ?? inheritedFont?.fontSize ?? 11) + 7) : 0), title: TitleComponentOption[] = []
  if (chart.title || chart.subtitle) title.push({ text: chart.title ?? '', subtext: chart.subtitle, left: 'center', top: 8,
    textStyle: chartEChartsFont({ bold: true, ...definedChartStyleFields(inheritedFont), ...definedChartStyleFields(chart.options?.titleStyle) }, color, 14),
    subtextStyle: chartEChartsFont({ ...definedChartStyleFields(inheritedFont), ...definedChartStyleFields(chart.options?.subtitleStyle) }, color, 11) })
  const categoryLegend = ['pie', 'funnel', 'gauge', 'treemap', 'sunburst'].includes(descriptor.family)
  const right = data.series.some((series, index) => chartSeriesStyle(chart, series.field, index, palette, series.sourceName ?? series.name).rightAxis)
  const valueAxis = numericAxis(chart.yAxis, axisColor, splitLineColor, horizontal ? 'x' : 'y', descriptor.percent ?? false, inheritedFont)
  const baseAxis = dates ? dateCategoryAxis(dates, chart.xAxis, axisColor, splitLineColor, inheritedFont)
    : categoryAxis(data.categories, chart.xAxis, axisColor, splitLineColor, inheritedFont)
  const secondAxis: ChartAxisOptions = chart.options?.rightYAxis ?? { scaleType: 'linear' }
  const pointer = chartAxisPointerPresentation(chart, theme)
  const tooltip: TooltipComponentOption = { trigger: isPie || isScatter || isRadar || descriptor.dataShape === 'hierarchy' ? 'item' : 'axis', showContent: false, triggerOn: 'none' }
  const pointerConfig: AxisPointerComponentOption | undefined = pointerEnabled && !isRadar ? { show: true, type: 'line', triggerOn: 'none', animation: false,
    lineStyle: { color: pointer.indicatorLineColor, type: pointer.indicatorLineType, width: 1 },
    label: { show: true, ...chartEChartsFont(pointer.font, pointer.indicatorLabelTextColor, 10), color: pointer.indicatorLabelTextColor, backgroundColor: pointer.indicatorLabelColor } } : undefined
  // ECharts tooltip.axisPointer has priority over the component defaults for both cross axes.
  if (pointerConfig) tooltip.axisPointer = { type: 'cross', lineStyle: pointerConfig.lineStyle, crossStyle: pointerConfig.lineStyle,
    label: pointerConfig.label, animation: false }
  const pointerRuntime: ChartAxisPointerRuntime | undefined = pointerEnabled ? isRadar
    ? { mode: 'graphic', geometry: chartAxisPointerGeometry(data, localChart, layoutDerivedChartData(derived, localChart, palette), theme) }
    : { mode: 'native' } : undefined
  const option: ChartEChartsOption = {
    animation: false, backgroundColor: chart.options?.appearance?.background ?? theme.backgroundColor ?? 'transparent',
    color: palette, textStyle: chartEChartsFont(chart.options?.appearance?.font, color),
    ...(title.length ? { title: title.length === 1 ? title[0] : title } : {}),
    legend: { show: chart.legend.visible, data: categoryLegend ? data.categories : seriesNames, ...legendPosition(chart.legend.position, titleHeight),
      textStyle: chartEChartsFont(chart.legend.font, color, 11) },
    tooltip,
    ...(pointerConfig ? { axisPointer: pointerConfig } : {}),
    ...(isPie || isRadar ? {} : {
      grid,
      xAxis: horizontal ? right ? [valueAxis, { ...numericAxis(secondAxis, axisColor, splitLineColor, 'x', false, inheritedFont), position: 'top', splitLine: { show: false } }] : valueAxis
        : dates ? baseAxis : isScatter ? numericAxis(chart.xAxis, axisColor, splitLineColor, 'x', false, inheritedFont) : baseAxis,
      yAxis: horizontal ? { ...baseAxis, inverse: !chart.xAxis.reverse }
        : right ? [valueAxis, { ...numericAxis(secondAxis, axisColor, splitLineColor, 'y', false, inheritedFont), position: 'right', splitLine: { show: false } }] : valueAxis,
    }),
    ...(isRadar ? { radar: radarComponent(data, chart, context) } : {}),
    series: isPie ? [pieSeries(data, chart, context)] : isScatter ? scatterSeries(data, chart, context)
      : isRadar ? radarSeries(data, chart, context) : categorySeries(data, chart, context),
    officeChartKey,
    ...(pointerRuntime ? { officeChartPointer: pointerRuntime } : {}),
  }
  const special = chartSpecialEChartsOption(data, derived, chart, context)
  if (special) {
    const previousLegend = option.legend
    Object.assign(option, special)
    if (special.legend && previousLegend && !Array.isArray(previousLegend) && !Array.isArray(special.legend)) option.legend = { ...previousLegend, ...special.legend }
  }
  if (derived.pareto) {
    const cumulative: LineSeriesOption = { type: 'line', name: '累计百分比', yAxisIndex: 1,
      data: dates ? derived.pareto.cumulativePercent.map((value, index) => dates.values[index] == null ? null : [dates.values[index]!, value]) : derived.pareto.cumulativePercent,
      showSymbol: true, lineStyle: { color: palette[1 % palette.length] } }
    option.series = [...categorySeries(data, chart, context), cumulative]
    option.yAxis = [numericAxis(chart.yAxis, axisColor, splitLineColor, 'y', false, inheritedFont), { ...numericAxis(secondAxis, axisColor, splitLineColor, 'y', true, inheritedFont), min: secondAxis.min ?? 0, max: secondAxis.max ?? 100,
      position: 'right', splitLine: { show: false } }]
    option.legend = { ...(!Array.isArray(option.legend) ? option.legend : {}), data: [...seriesNames, '累计百分比'] }
  }
  if (isRadar) {
    const { plot } = frame, titleY = plot.y + plot.height + 24 + Math.max(0, (chartAxisTitleFont(chart, chart.xAxis).fontSize ?? 11) - 11) / 2
    const titleX = plot.x - 54 - chartAxisTitleSpace(chart, chart.yAxis) / 2
    const titles = [chartAxisTitleCommand(chart, chart.xAxis, { x: plot.x, y: titleY }, { x: plot.x + plot.width, y: titleY }, color, 0, frame.frame),
      chartAxisTitleCommand(chart, chart.yAxis, { x: titleX, y: plot.y + plot.height }, { x: titleX, y: plot.y }, color, -90, frame.frame)]
      .filter((command): command is DrawCommand => command !== undefined)
    if (titles.length) option.graphic = { elements: chartGraphics(titles, `${chart.id}:axis-titles`) }
  }
  if (chart.options?.appearance?.border?.visible) {
    const border = chart.options.appearance.border, graphic: NonNullable<GraphicComponentOption['elements']>[number] = { type: 'rect',
      shape: { x: 0.5, y: 0.5, width: chart.size.width - 1, height: chart.size.height - 1 }, style: { fill: 'transparent', stroke: border.color ?? axisColor, lineWidth: border.width ?? 1, opacity: border.opacity ?? 1 }, silent: true, z: -10 }
    const current = option.graphic
    option.graphic = { elements: [graphic, ...(!Array.isArray(current) ? current?.elements ?? [] : [])] }
  }
  // Specialized adapters can supply their own tooltip; raster surfaces receive pointer actions explicitly.
  if (option.tooltip && !Array.isArray(option.tooltip)) option.tooltip = { ...option.tooltip, showContent: false, triggerOn: 'none',
    ...(pointerConfig ? { axisPointer: tooltip.axisPointer } : {}) }
  return option
}

function requiresGraphicChart(chart: ChartModel, data: ChartData, boxes?: readonly { field: number; seriesPosition: number; sourceName: string }[]): boolean {
  const radar = chart.type === 'radar' ? chartRadarExtent(data, chart.yAxis) : undefined
  return Boolean(resolveChartType(chart.type).threeD || ['pieOfPie', 'barOfPie', 'wordCloud', 'chord', 'waterfall', 'sankey', 'relation'].includes(chart.type)
    // Scatter and Heatmap do not support native decals; every patterned mark uses the shared geometry.
    || data.series.some((source, index) => {
      const style = chartSeriesStyle(chart, source.field, index, undefined, source.sourceName ?? source.name)
      return style.fillPattern || Object.values(style.dataPoints ?? {}).some(point => point?.fillPattern)
    })
    || chart.type === 'radar' && data.series.some((source, index) => {
      const style = chartSeriesStyle(chart, source.field, index, undefined, source.sourceName ?? source.name)
      return style.label?.visible || Object.keys(style.dataPoints ?? {}).length > 0
    })
    || chart.type === 'radar' && (chart.yAxis.reverse || Boolean(chart.xAxis.label?.rotate))
    || radar && !Number.isFinite(radar.max - radar.min)
    || chart.type === 'boxplot' && boxes?.some(box => {
      const style = chartSeriesStyle(chart, box.field, box.seriesPosition, undefined, box.sourceName)
      return style.label?.visible || style.rightAxis
    })
    || chart.type === 'gauge' && !Number.isFinite((chart.options?.gauge?.max ?? 100) - (chart.options?.gauge?.min ?? 0))
    || chart.type === 'candlestick' && data.candles?.some(candle => candle !== null && candle.open === undefined)
    || chart.options?.mapping?.useDateAxis && ['bar', 'boxplot', 'candlestick'].includes(resolveChartType(chart.type).family)
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
    else throw new TypeError('图表几何不能携带位图，请使用图表位图适配层。')
    return command.clip ? { type: 'group', id: `${id}-clip`, silent: true, children: [element],
      clipPath: { type: 'rect', shape: { x: command.clip.x, y: command.clip.y, width: command.clip.width, height: command.clip.height } } } : element
  })
}

function radarComponent(data: ChartData, chart: ChartModel, context: ChartEChartsContext): RadarComponentOption {
  const { grid, color, axisColor, splitLineColor } = context
  const geometry = chartRadarGeometry(data, chart, { x: grid.left, y: grid.top, width: chart.size.width - grid.left - grid.right, height: chart.size.height - grid.top - grid.bottom })
  const label = chartRadarAxisLabel(chart), line = chartRadarAxisLine(chart), axis = chartEChartsAxisStyle({ ...chart.yAxis, line }, axisColor, splitLineColor, true, chart.options?.appearance?.font)
  return { center: [geometry.center.x, geometry.center.y], radius: geometry.radius, startAngle: 90, splitNumber: geometry.ticks.length - 1,
    clockwise: !chart.xAxis.reverse, shape: chart.options?.radar?.shape ?? 'polygon', indicator: data.categories.map(name => ({ name, min: geometry.min, max: geometry.max })),
    axisNameGap: geometry.radius * 0.22, axisName: { show: chart.xAxis.label?.visible !== false, ...chartEChartsFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(chart.xAxis.label) }, color, 10) },
    axisLine: axis.axisLine, axisTick: axis.axisTick,
    axisLabel: { show: label.visible ?? false, showMinLabel: false, rotate: chart.yAxis.label?.rotate ?? 0, ...chartEChartsFont(label, color, 10),
      formatter: (value: string | number) => formatChartNumber(Number(value), label.format) },
    splitLine: axis.splitLine, splitArea: { show: false } }
}

function radarSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): RadarSeriesOption[] {
  return data.series.map((source, index) => {
    const style = chartSeriesStyle(chart, source.field, index, context.palette, source.sourceName ?? source.name)
    return { type: 'radar', name: style.name ?? source.name, symbol: style.point?.shape ?? 'none', symbolSize: style.point?.size ?? 4,
      lineStyle: chartEChartsLineStyle(style), itemStyle: chartEChartsMarkerStyle(style),
      ...(chart.options?.radar?.fill ? { areaStyle: { color: style.color, opacity: style.fillOpacity ?? 0.25 } } : {}),
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

function categoryAxis(categories: string[], axis: ChartAxisOptions, axisColor: string, splitLineColor: string, inheritedFont?: ChartFontOptions): XAxisComponentOption {
  return { type: 'category', data: categories, ...chartEChartsAxisStyle(axis, axisColor, splitLineColor, false, inheritedFont) }
}

function numericAxis(axis: ChartAxisOptions, axisColor: string, splitLineColor: string, dimension: 'x' | 'y', percent = false, inheritedFont?: ChartFontOptions): XAxisComponentOption & YAxisComponentOption {
  const style = chartEChartsAxisStyle(axis, axisColor, splitLineColor, true, inheritedFont)
  return { type: axis.scaleType === 'log' ? 'log' : 'value', ...(axis.scaleType === 'log' ? { logBase: axis.logBase ?? 10 } : {}),
    ...style, min: axis.min, max: axis.max, nameGap: Math.max(style.nameGap, dimension === 'x' ? 28 : 42),
    axisPointer: { label: { formatter: params => `${formatChartNumber(Number(params.value), axis.format)}${percent && !axis.format ? '%' : ''}` } },
    ...(percent && !axis.format ? { axisLabel: { ...style.axisLabel, formatter: '{value}%' } } : {}),
    ...(dimension === 'x' && axis.gridLine?.visible === undefined ? { splitLine: { show: false } } : {}) }
}

function dateCategoryAxis(dates: ChartDateAxis, axis: ChartAxisOptions, axisColor: string, splitLineColor: string, inheritedFont?: ChartFontOptions): XAxisComponentOption {
  const style = chartEChartsAxisStyle(axis, axisColor, splitLineColor, false, inheritedFont)
  const label = (value: string | number): string => chartDateAxisLabel(Number(value), axis, dates.date1904, dates.max - dates.min)
  return { ...style, type: 'value', min: dates.min, max: dates.max,
    axisLabel: { ...style.axisLabel, formatter: label }, axisPointer: { snap: true, label: { formatter: params => label(params.value) } } }
}

function pieSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): PieSeriesOption {
  const source = data.series[0]!, options = chart.options?.pie, { palette, color, grid } = context
  const scale = source.values.reduce<number>((max, value) => Math.max(max, value ?? 0), 0) || 1, total = source.values.reduce<number>((sum, value) => sum + Math.max(0, value ?? 0) / scale, 0), start = -(options?.startAngle ?? -90)
  const radius = chartPieOuterRadius({ x: grid.left, y: grid.top, width: chart.size.width - grid.left - grid.right, height: chart.size.height - grid.top - grid.bottom }, options?.explosion)
  return { type: 'pie', name: source.name, radius: chart.type === 'doughnut' ? [radius * (options?.doughnutHole ?? 0.55), radius] : radius,
    center: ['50%', '50%'], left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
    startAngle: start, ...(options?.isHalfPie ? { endAngle: start - 180 } : {}), ...(options?.rosePie ? { roseType: 'radius' } : {}),
    ...(options?.explosion ? { selectedMode: 'multiple', selectedOffset: radius * options.explosion / 100 } : {}),
    avoidLabelOverlap: true, stillShowZeroSum: false, showEmptyCircle: false, label: { color }, labelLine: { show: options?.showLabelLine !== false },
    data: data.categories.map((name, index) => {
      const value = Math.max(0, source.values[index] ?? 0), style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index), label = { ...style.label, ...options?.label }
      return { name, value, selected: Boolean(options?.explosion), itemStyle: chartEChartsItemStyle(style),
        label: chartEChartsPointLabel({ visible: true, contentType: 'category', position: 'outside', ...label }, color, value, name, source.name, total ? (value / scale) / total * 100 : 0) }
    }) }
}

function scatterSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): ScatterSeriesOption[] {
  let maxSize = 1
  for (const source of data.series) for (const size of source.sizeValues ?? []) if (size !== null) maxSize = Math.max(maxSize, size)
  return data.series.map((source, seriesIndex) => {
    const style = chartSeriesStyle(chart, source.field, seriesIndex, context.palette, source.sourceName ?? source.name), name = style.name ?? source.name, bubble = chart.type === 'bubble'
    const axis: ChartAxisOptions = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    return { type: 'scatter', name, xAxisIndex: 0, yAxisIndex: style.rightAxis ? 1 : 0, symbol: style.point?.shape ?? 'circle', symbolSize: style.point?.size ?? 8,
      itemStyle: chartEChartsMarkerStyle(style), label: chartEChartsLabel(style.label, context.color),
      data: source.values.flatMap<NonNullable<ScatterSeriesOption['data']>[number]>((value, index) => {
        const x = source.xValues?.[index], size = source.sizeValues?.[index]
        if (value === null || x == null || !Number.isFinite(value) || !Number.isFinite(x) || chart.xAxis.scaleType === 'log' && x <= 0 || axis.scaleType === 'log' && value <= 0 || bubble && (size == null || size <= 0)) return []
        const pointStyle = chartPointStyle(style, index)
        if (!bubble && !style.dataPoints?.[index] && !style.label?.visible) return [[x, value]]
        const diameter = bubble ? chartBubblePointSize(pointStyle, size!, maxSize, chart.options, Math.min(chart.size.width - context.grid.left - context.grid.right, chart.size.height - context.grid.top - context.grid.bottom)) : pointStyle.point?.size ?? 8
        return [{ name: data.categories[index], value: [x, value, ...(bubble ? [size!] : [])], symbol: pointStyle.point?.shape ?? 'circle', symbolSize: diameter, itemStyle: chartEChartsMarkerStyle(pointStyle),
          ...(pointStyle.label?.visible ? { label: chartEChartsPointLabel(pointStyle.label, context.color, value, data.categories[index] ?? '', name) } : {}) }]
      }) }
  })
}

function categorySeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): Array<BarSeriesOption | LineSeriesOption> {
  return data.series.map((source, index) => {
    const kind = chartSeriesKind(chart, source.field, index, source.sourceName ?? source.name), style = chartSeriesStyle(chart, source.field, index, context.palette, source.sourceName ?? source.name), name = style.name ?? source.name
    const axis: ChartAxisOptions = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const values = source.values.map((value, categoryIndex) => {
      const date = context.dateAxis?.values[categoryIndex]
      if (value === null || axis.scaleType === 'log' && value <= 0 || context.dateAxis && date == null) return null
      const datum = context.dateAxis ? resolveChartType(chart.type).family === 'bar' ? [value, date!] : [date!, value] : value
      if (!style.dataPoints?.[categoryIndex] && !style.label?.visible) return datum
      const point = chartPointStyle(style, categoryIndex)
      return { value: datum, itemStyle: kind.type === 'column' || kind.type === 'bar' ? chartEChartsItemStyle(point) : chartEChartsMarkerStyle(point), symbol: point.point?.shape, symbolSize: point.point?.size,
        ...(point.label?.visible ? { label: chartEChartsPointLabel(point.label, context.color, value, data.categories[categoryIndex] ?? '', name) } : {}) }
    })
    const stacked = kind.stacked ? chart.type === 'combination' ? `${kind.type}-${style.rightAxis ? 'right' : 'left'}` : style.rightAxis ? 'total-right' : 'total' : undefined
    if (kind.type === 'bar' || kind.type === 'column' || chart.type === 'histogram' || chart.type === 'pareto') return {
      type: 'bar', name, data: values, stack: stacked,
      ...(resolveChartType(chart.type).family === 'bar' ? { xAxisIndex: style.rightAxis ? 1 : 0, yAxisIndex: 0 } : style.rightAxis ? { yAxisIndex: 1 } : {}),
      itemStyle: chartEChartsItemStyle(style), label: chartEChartsLabel(style.label, context.color),
      barCategoryGap: chart.type === 'histogram' ? `${(chart.options?.histogram?.binGap ?? 0.05) * 100}%` : `${chart.options?.bar?.gapWidth ?? 150}%`,
      ...(chart.options?.bar?.overlap === undefined ? {} : { barGap: `${-chart.options.bar.overlap}%` }),
    }
    return { type: 'line', name, data: values, stack: stacked, yAxisIndex: style.rightAxis ? 1 : 0,
      showSymbol: style.point?.shape !== 'none' || Object.values(style.dataPoints ?? {}).some(point => point?.shape !== undefined && point.shape !== 'none'),
      symbol: style.point?.shape ?? 'circle', symbolSize: style.point?.size ?? 6, connectNulls: chart.options?.invalidValueStrategy === 'link',
      smooth: kind.type === 'area' && chart.options?.area?.smooth, lineStyle: chartEChartsLineStyle(style), itemStyle: chartEChartsMarkerStyle(style), label: chartEChartsLabel(style.label, context.color),
      ...(kind.type === 'area' ? { areaStyle: { color: style.color, opacity: style.fillOpacity ?? 0.28 } } : {}) }
  })
}
