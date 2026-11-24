import type {
  BarSeriesOption, BoxplotSeriesOption, CandlestickSeriesOption, FunnelSeriesOption, GaugeSeriesOption, GraphSeriesOption,
  HeatmapSeriesOption, LineSeriesOption, PieSeriesOption, RadarSeriesOption, SankeySeriesOption, ScatterSeriesOption, SunburstSeriesOption, TreemapSeriesOption,
} from 'echarts/charts'
import type {
  AxisPointerComponentOption, GraphicComponentOption, GridComponentOption, LegendComponentOption, RadarComponentOption, TitleComponentOption, TooltipComponentOption, VisualMapComponentOption,
} from 'echarts/components'
import type { ComposeOption } from 'echarts/core'
import type { XAxisComponentOption, YAxisComponentOption } from 'echarts'
import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartModel, WorkbookModel } from '../../domain/model'
import type { DrawCommand } from '../../rendering/contracts'
import { readChartData, type ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartBubblePointSize, chartPointStyle, chartSeriesKind, deriveChartData, type ChartComputedStatistics } from './chart-derived'
import { chartPalette, chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { chartEChartsAxisStyle, chartEChartsFont, chartEChartsIndependentItemStyle, chartEChartsItemStyle, chartEChartsLabel, chartEChartsLineStyle, chartEChartsMarkerStyle, chartEChartsPointLabel } from './echarts-style'
import { chartSpecialEChartsOption, type ChartEChartsContext } from './echarts-special-option'
import { chartBoxplotAxis, chartBoxplotOutliers, chartBoxplotOutlierStyle, chartBoxplotStyle, chartBoxplotValueDrawable, type ChartBoxplotDistribution } from './chart-boxplot'
import { drawDerivedChartData } from './chart-draw'
import { chartPlotFrame, layoutDerivedChartData } from './chart-layout'
import { chartPieLayout, chartPieOuterRadius } from './chart-pie-layout'
import { chartRadarAxisLabel, chartRadarAxisLine, chartRadarExtent, chartRadarGeometry } from './chart-radar'
import { formatChartNumber } from './chart-geometry'
import { chartRectangularAxisTitleCommands } from './chart-axis-title'
import { chartDateAxis, chartDateSeriesOrder, type ChartDateAxis } from './chart-date-axis'
import { chartDateAxisCommands } from './chart-date-axis-draw'
import { chartAxisPointerGeometry, chartAxisPointerPresentation, supportsChartAxisPointer, type ChartAxisPointerRuntime } from './chart-axis-pointer'
import type { ChartRenderTheme } from './chart-theme'
import { chartAutoGradientFill, chartEChartsAutomaticGradient } from './chart-auto-gradient'
import { chartColorRepresentative, isChartLinearGradient } from './chart-color'
import { createChartMarkGradient, supportsChartExplicitGradient } from './chart-gradient'
import { CHART_PARETO_CUMULATIVE_NAME, chartParetoBarStyle, chartParetoCumulativeExtent, chartParetoCumulativeMarkerStyle, chartParetoCumulativeStyle, chartParetoValuePercentages } from './chart-pareto'
import { chartMarkerRequiresSharedGeometry, chartNativeMarkerSymbol } from './chart-markers'
import { createChartSankeyScene, type ChartSankeyScene } from './chart-sankey-draw'
import { createChartSankeyHitIndex, createChartSankeyTopology, type ChartSankeyPointerRuntime } from './chart-sankey-interaction'
import { chartHeatmapExtent } from './chart-heatmap'
import { chartHistogramBinGap } from './chart-histogram'
import { chartCandlestickMode, chartCandlestickState, chartCandlestickStyle } from './chart-candlestick'

export type ChartEChartsOption = ComposeOption<
  | BarSeriesOption | LineSeriesOption | PieSeriesOption | RadarSeriesOption | ScatterSeriesOption | BoxplotSeriesOption
  | CandlestickSeriesOption | FunnelSeriesOption | GaugeSeriesOption | GraphSeriesOption | HeatmapSeriesOption | SankeySeriesOption | SunburstSeriesOption | TreemapSeriesOption
  | GraphicComponentOption | GridComponentOption | LegendComponentOption | RadarComponentOption | TitleComponentOption | TooltipComponentOption
  | AxisPointerComponentOption | VisualMapComponentOption | XAxisComponentOption | YAxisComponentOption
> & {
  /** Authored dependencies used by the bitmap cache, including formatter closures. Removed before ECharts.setOption. */
  officeChartKey?: string
  /** Transient surface interaction metadata. Removed before signature generation and setOption. */
  officeChartPointer?: ChartAxisPointerRuntime | ChartSankeyPointerRuntime
}

export type ChartEChartsTheme = ChartRenderTheme

/** Converts authored semantics and derived statistics to a Canvas option without mutating the document. */
export function chartEChartsOption(workbook: WorkbookModel, chart: ChartModel, theme: ChartEChartsTheme = {}, computed?: ChartComputedStatistics): ChartEChartsOption {
  const source = computed?.preparedData ?? readChartData(workbook, chart), derived = deriveChartData(source, chart, computed), data = derived.data
  const businessPercentages = derived.pareto ? chartParetoValuePercentages(derived.pareto.values) : undefined
  const descriptor = resolveChartType(chart.type), color = chart.options?.appearance?.font?.color ?? theme.textColor ?? '#334155'
  const axisColor = theme.axisColor ?? '#94a3b8', splitLineColor = theme.splitLineColor ?? '#e2e8f0', palette = chartPalette(chart, theme.palette)
  const localChart = { ...chart, position: { x: 0, y: 0 } }
  const officeChartKey = JSON.stringify([localChart, theme, Boolean(data.date1904)])
  const pointerEnabled = supportsChartAxisPointer(chart) && data.categories.length > 0 && data.series.length > 0
  const frame = chartPlotFrame(localChart, derived.boxes ?? data.series), grid = {
    left: frame.plot.x, right: Math.max(0, chart.size.width - frame.plot.x - frame.plot.width),
    // chartPlotFrame already reserves automatic label/title bands; authored bounds are exact.
    top: frame.plot.y, bottom: Math.max(0, chart.size.height - frame.plot.y - frame.plot.height), containLabel: false,
  }
  if (!data.categories.length || !data.series.length || descriptor.family === 'gauge' && !data.series[0]!.values.some(Number.isFinite) || requiresGraphicChart(chart, data, derived.boxes, derived.pareto?.cumulativePercent)) {
    const layout = layoutDerivedChartData(derived, localChart, palette)
    const sankeyScene = layout.sankey ? createChartSankeyScene(data, localChart, layout.sankey,
      chart.options?.appearance?.font?.color ?? theme.textColor ?? '#64748b', palette, layout.frame) : undefined
    const commands = drawDerivedChartData(derived, localChart, theme, palette, layout, sankeyScene)
    const sankeyGraphics = sankeyScene?.commands.length ? chartSankeyGraphics(commands, sankeyScene, chart.id) : undefined
    const pointerRuntime: ChartAxisPointerRuntime | ChartSankeyPointerRuntime | undefined = sankeyGraphics
      ? { mode: 'sankey', groupId: sankeyGraphics.groupId, zOffset: sankeyGraphics.zOffset, scene: sankeyScene!,
        hitIndex: createChartSankeyHitIndex(sankeyScene!.layout), topology: createChartSankeyTopology(sankeyScene!.layout) }
      : pointerEnabled ? { mode: 'graphic', geometry: chartAxisPointerGeometry(data, localChart, layout, theme) } : undefined
    return {
      animation: false, backgroundColor: chart.options?.appearance?.background ?? theme.backgroundColor ?? 'transparent',
      graphic: { elements: sankeyGraphics?.elements ?? chartGraphics(commands, chart.id) }, officeChartKey,
      ...(pointerRuntime ? { officeChartPointer: pointerRuntime } : {}),
    }
  }
  const isPie = descriptor.family === 'pie', isScatter = chart.type === 'scatter' || chart.type === 'bubble', isRadar = chart.type === 'radar', horizontal = descriptor.family === 'bar'
  const dates = chartDateAxis(data, chart), inheritedFont = chart.options?.appearance?.font
  const sharedLayout = dates || chart.type === 'candlestick' || pointerEnabled && (isRadar || chart.type === 'heatmap') ? layoutDerivedChartData(derived, localChart, palette) : undefined
  const context: ChartEChartsContext = { palette, color, axisColor, splitLineColor, grid, dateAxis: dates, backgroundColor: chart.options?.appearance?.background ?? theme.backgroundColor }
  const seriesNames = data.series.map((series, index) => chartSeriesStyle(chart, series.field, index, palette, series.sourceName ?? series.name).name ?? series.name)
  const titleHeight = (chart.title ? Math.max(28, (chart.options?.titleStyle?.fontSize ?? inheritedFont?.fontSize ?? 14) + 14) : 0)
    + (chart.subtitle ? Math.max(18, (chart.options?.subtitleStyle?.fontSize ?? inheritedFont?.fontSize ?? 11) + 7) : 0), title: TitleComponentOption[] = []
  if (chart.title || chart.subtitle) title.push({ text: chart.title ?? '', subtext: chart.subtitle, left: 'center', top: 8,
    textStyle: chartEChartsFont({ bold: true, ...definedChartStyleFields(inheritedFont), ...definedChartStyleFields(chart.options?.titleStyle) }, color, 14),
    subtextStyle: chartEChartsFont({ ...definedChartStyleFields(inheritedFont), ...definedChartStyleFields(chart.options?.subtitleStyle) }, color, 11) })
  const categoryLegend = ['pie', 'funnel', 'gauge', 'treemap', 'sunburst'].includes(descriptor.family)
  const right = data.series.some((series, index) => chartSeriesStyle(chart, series.field, index, palette, series.sourceName ?? series.name).rightAxis)
  const valueAxis = { ...numericAxis(chart.yAxis, axisColor, splitLineColor, horizontal ? 'x' : 'y', descriptor.percent ?? false, inheritedFont),
    ...(dates || chart.type === 'candlestick' ? { min: sharedLayout!.min, max: sharedLayout!.max } : {}) }
  const baseAxis = dates ? dateCategoryAxis(dates, chart.xAxis, axisColor, splitLineColor, inheritedFont)
    : categoryAxis(data.categories, chart.xAxis, axisColor, splitLineColor, inheritedFont)
  const secondAxis: ChartAxisOptions = chart.options?.rightYAxis ?? { scaleType: 'linear' }
  const secondaryValueAxis = { ...numericAxis(secondAxis, axisColor, splitLineColor, horizontal ? 'x' : 'y', false, inheritedFont),
    ...(dates ? { min: sharedLayout!.rightMin, max: sharedLayout!.rightMax } : {}) }
  const pointer = chartAxisPointerPresentation(chart, theme)
  const tooltip: TooltipComponentOption = { trigger: isPie || isScatter || isRadar || descriptor.dataShape === 'hierarchy' ? 'item' : 'axis', showContent: false, triggerOn: 'none' }
  const pointerConfig: AxisPointerComponentOption | undefined = pointerEnabled && !isRadar && chart.type !== 'heatmap' && !dates ? { show: true, type: 'line', triggerOn: 'none', animation: false,
    lineStyle: { color: pointer.indicatorLineColor, type: pointer.indicatorLineType, width: 1 },
    label: { show: true, ...chartEChartsFont(pointer.font, pointer.indicatorLabelTextColor, 10), color: pointer.indicatorLabelTextColor, backgroundColor: pointer.indicatorLabelColor } } : undefined
  // ECharts tooltip.axisPointer has priority over the component defaults for both cross axes.
  if (pointerConfig) tooltip.axisPointer = { type: 'cross', lineStyle: pointerConfig.lineStyle, crossStyle: pointerConfig.lineStyle,
    label: pointerConfig.label, animation: false }
  const pointerRuntime: ChartAxisPointerRuntime | undefined = pointerEnabled ? isRadar || chart.type === 'heatmap' || dates
    ? { mode: 'graphic', geometry: chartAxisPointerGeometry(data, localChart, sharedLayout!, theme) }
    : { mode: 'native' } : undefined
  if (pointerRuntime?.mode === 'graphic' && pointerRuntime.geometry.heatmap) context.heatmapScale = pointerRuntime.geometry.heatmap.scale
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
      xAxis: horizontal ? right ? [valueAxis, { ...secondaryValueAxis, position: 'top', splitLine: { show: false } }] : valueAxis
        : dates ? baseAxis : isScatter ? numericAxis(chart.xAxis, axisColor, splitLineColor, 'x', false, inheritedFont) : baseAxis,
      yAxis: horizontal ? { ...baseAxis, inverse: !chart.xAxis.reverse }
        : right ? [valueAxis, { ...secondaryValueAxis, position: 'right', splitLine: { show: false } }] : valueAxis,
    }),
    ...(isRadar ? { radar: radarComponent(data, chart, context) } : {}),
    series: isPie ? [pieSeries(data, chart, context)] : isScatter ? scatterSeries(data, chart, context)
      : isRadar ? radarSeries(data, chart, context) : categorySeries(data, chart, context, businessPercentages),
    officeChartKey,
    ...(pointerRuntime ? { officeChartPointer: pointerRuntime } : {}),
  }
  const special = chartSpecialEChartsOption(data, derived, chart, context)
  if (special) {
    const previousLegend = option.legend
    Object.assign(option, special)
    if (special.legend && previousLegend && !Array.isArray(previousLegend) && !Array.isArray(special.legend)) option.legend = { ...previousLegend, ...special.legend }
  }
  if (pointerRuntime?.mode === 'graphic' && pointerRuntime.geometry.heatmap) {
    option.graphic = { elements: chartGraphics(pointerRuntime.geometry.heatmap.legend.commands, `${chart.id}:heatmap-legend`) }
  }
  if (derived.pareto) {
    const style = chartParetoCumulativeStyle(chart, palette), name = CHART_PARETO_CUMULATIVE_NAME
    const extent = chartParetoCumulativeExtent(chart, derived.pareto.cumulativePercent)
    const cumulative: LineSeriesOption = { type: 'line', name, yAxisIndex: 1,
      data: (dates ? chartDateSeriesOrder(dates, chart) : derived.pareto.cumulativePercent.map((_value, index) => index)).map(index => {
        const value = derived.pareto!.cumulativePercent[index]!
        if (secondAxis.scaleType === 'log' && value <= 0) return null
        const datum = dates ? [dates.coordinates[index]!, value] : value
        return dates || style.label?.visible ? { ...(dates ? { id: `${chart.id}:pareto-cumulative:${index}`, name: data.categories[index] } : {}), value: datum,
          ...(style.label?.visible ? { label: chartEChartsPointLabel(style.label, color, value, data.categories[index] ?? '', name, value) } : {}) } : datum
      }),
      showSymbol: style.point?.shape !== 'none' && (style.point?.size ?? 7) > 0, symbol: chartNativeMarkerSymbol(style.point?.shape), symbolSize: style.point?.size ?? 7,
      connectNulls: chart.options?.invalidValueStrategy === 'link', lineStyle: chartEChartsLineStyle(style), itemStyle: chartEChartsMarkerStyle(chartParetoCumulativeMarkerStyle(style)), label: chartEChartsLabel(style.label, color) }
    option.series = [...categorySeries(data, chart, context, businessPercentages), cumulative]
    option.yAxis = [{ ...numericAxis(chart.yAxis, axisColor, splitLineColor, 'y', false, inheritedFont), ...(dates ? { min: sharedLayout!.min, max: sharedLayout!.max } : {}) }, { ...numericAxis(secondAxis, axisColor, splitLineColor, 'y', true, inheritedFont), min: extent.min, max: extent.max,
      position: 'right', splitLine: { show: false } }]
    option.legend = { ...(!Array.isArray(option.legend) ? option.legend : {}), data: [...seriesNames, '累计百分比'] }
  }
  if (dates) {
    const commands = chartDateAxisCommands(localChart, sharedLayout!, chart.options?.appearance?.font?.color ?? theme.textColor ?? '#64748b', theme)
    if (commands.length) {
      const current = option.graphic
      option.graphic = { elements: [...(!Array.isArray(current) ? current?.elements ?? [] : []), ...chartGraphics(commands, `${chart.id}:date-axis`, -commands.length)] }
    }
  }
  if ((option.xAxis || option.yAxis || isRadar) && (chart.xAxis.title || chart.yAxis.title || chart.options?.rightYAxis?.title)) {
    const layout = pointerRuntime?.mode === 'graphic' ? pointerRuntime.geometry.layout : layoutDerivedChartData(derived, localChart, palette)
    const titles = chartRectangularAxisTitleCommands(localChart, layout, chart.options?.appearance?.font?.color ?? theme.textColor ?? '#64748b')
    if (titles.length) {
      const current = option.graphic
      option.graphic = { elements: [...(!Array.isArray(current) ? current?.elements ?? [] : []), ...chartGraphics(titles, `${chart.id}:axis-titles`)] }
    }
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

function requiresGraphicLabel(label: ChartLabelOptions | undefined): boolean {
  return Boolean(label?.visible && (label.position === 'auto' || label.position === 'insideBase' || label.position === 'insideEnd'))
}

function requiresGraphicChart(chart: ChartModel, data: ChartData, boxes?: readonly ChartBoxplotDistribution[], cumulativePercent?: readonly number[]): boolean {
  const radar = chart.type === 'radar' ? chartRadarExtent(data, chart.yAxis) : undefined
  const pareto = chart.type === 'pareto' ? chartParetoCumulativeStyle(chart, chartPalette(chart)) : undefined
  const cumulativeAxis = chart.options?.rightYAxis
  const paretoGeometry = pareto && cumulativePercent?.some((value, index) => {
    if (cumulativeAxis?.scaleType === 'log' && value <= 0) return false
    const marker = chartParetoCumulativeMarkerStyle(pareto, index)
    return isChartLinearGradient(pareto.color)
      || Object.values(pareto.dataPoints?.[index] ?? {}).some(entry => entry !== undefined)
      || requiresGraphicLabel(pareto.label)
      || (marker.point?.size ?? 7) > 0 && chartMarkerRequiresSharedGeometry(marker.point?.shape)
  })
  return Boolean(resolveChartType(chart.type).threeD || ['pieOfPie', 'barOfPie', 'wordCloud', 'chord', 'waterfall', 'sankey', 'relation'].includes(chart.type)
    // Scatter and Heatmap do not support native decals; every patterned mark uses the shared geometry.
    || data.series.some((source, index) => {
      // Boxplot data here contains summaries. Its sample styles are evaluated
      // against the original outlier identities in the branch below.
      if (chart.type === 'boxplot') return false
      const style = chart.type === 'pareto' ? chartParetoBarStyle(chart, source.field, index, chartPalette(chart), source.sourceName ?? source.name)
        : chartSeriesStyle(chart, source.field, index, undefined, source.sourceName ?? source.name)
      const kind = chartSeriesKind(chart, source.field, index, source.sourceName ?? source.name)
      const axis = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
      const usesMarkers = kind.type === 'line' || kind.type === 'area' || chart.type === 'scatter' || chart.type === 'bubble' || chart.type === 'radar'
      const sharedActiveMark = source.values.some((value, pointIndex) => {
        if (value === null || !Number.isFinite(value) || axis.scaleType === 'log' && value <= 0) return false
        if (chart.type === 'scatter' || chart.type === 'bubble') {
          const x = source.xValues?.[pointIndex]
          if (x == null || !Number.isFinite(x) || chart.xAxis.scaleType === 'log' && x <= 0
            || chart.type === 'bubble' && (source.sizeValues?.[pointIndex] ?? 0) <= 0) return false
        }
        const point = chartPointStyle(style, pointIndex, chart.type)
        return requiresGraphicLabel(point.label)
          || usesMarkers && (point.point?.size ?? 1) > 0 && chartMarkerRequiresSharedGeometry(point.point?.shape)
      })
      return style.fillPattern || Object.values(style.dataPoints ?? {}).some(point => point?.fillPattern)
        || sharedActiveMark
        // Explicit line/area gradients share the Canvas mark rectangle across
        // fragments; native paints cannot reproduce these segment coordinates.
        || (kind.type === 'line' || kind.type === 'area') && supportsChartExplicitGradient(chart.type)
          && (isChartLinearGradient(style.color) || supportsChartExplicitGradient(chart.type, 'dataPoint') && Object.values(style.dataPoints ?? {}).some(point => isChartLinearGradient(point?.color)))
        // Native line/area paints cannot vary per segment; use the common scene for authored point overrides.
        || chartAutoGradientFill(chart) && (kind.type === 'line' || kind.type === 'area') && Object.values(style.dataPoints ?? {}).some(point => point && (point.color !== undefined || point.fillOpacity !== undefined || point.border?.color !== undefined || point.border?.opacity !== undefined))
    })
    || paretoGeometry
    || chart.type === 'radar' && data.series.some((source, index) => {
      const style = chartSeriesStyle(chart, source.field, index, undefined, source.sourceName ?? source.name)
      return style.label?.visible || Object.keys(style.dataPoints ?? {}).length > 0
    })
    || chart.type === 'radar' && (chart.yAxis.reverse || Boolean(chart.xAxis.label?.rotate))
    || radar && !Number.isFinite(radar.max - radar.min)
    || chart.type === 'boxplot' && boxes?.some(box => {
      const style = chartBoxplotStyle(chart, box, box.seriesPosition, chartPalette(chart)), axis = chartBoxplotAxis(chart, style)
      if (style.label?.visible || style.rightAxis || style.fillPattern) return true
      if (chart.options?.boxplot?.showOutliers === false) return false
      for (const outlier of chartBoxplotOutliers(box)) {
        if (!chartBoxplotValueDrawable(outlier.value, axis)) continue
        const point = chartBoxplotOutlierStyle(style, outlier.sourceIndex)
        if ((point.point?.size ?? 4) > 0 && point.point?.shape !== 'none'
          && (point.fillPattern || chartMarkerRequiresSharedGeometry(point.point?.shape))) return true
      }
      return false
    })
    || chart.type === 'gauge' && !Number.isFinite((chart.options?.gauge?.max ?? 100) - (chart.options?.gauge?.min ?? 0))
    || chart.type === 'heatmap' && (() => { const extent = chartHeatmapExtent(data, chart); return !Number.isFinite(extent.max - extent.min) })()
    || chart.type === 'candlestick' && (chartCandlestickMode(chart, data) === 'hlc'
      || chart.options?.candlestick?.highLowLine !== undefined || chart.options?.candlestick?.closeMarker !== undefined
      || data.candles?.some(candle => {
        if (!candle) return false
        const source = data.series[0]!, style = chartCandlestickStyle(chart, source.field, candle.categoryIndex, chartCandlestickState(candle), chartPalette(chart), source.sourceName ?? source.name)
        return style.label?.visible || style.border?.dashType && style.border.dashType !== 'solid'
      }))
    || chart.options?.mapping?.useDateAxis && ['bar', 'boxplot', 'candlestick'].includes(resolveChartType(chart.type).family)
    || chart.options?.trendlines?.length || chart.options?.errorBars?.length)
}

export function chartGraphics(commands: DrawCommand[], chartId: string, zOffset = 0): NonNullable<GraphicComponentOption['elements']> {
  return commands.map((command, index): NonNullable<GraphicComponentOption['elements']>[number] => {
    const id = `${chartId}-${index}`, common = { id, silent: true, z: index + zOffset }
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

/** Only this group is replaced on hover/drag; the frame, titles and legend
 * keep their existing graphic elements and the authored cache signature. */
function chartSankeyGraphics(commands: DrawCommand[], scene: ChartSankeyScene, chartId: string): {
  elements: NonNullable<GraphicComponentOption['elements']>; groupId: string; zOffset: number
} {
  const start = commands.indexOf(scene.commands[0]!), end = start + scene.commands.length
  if (start < 0) throw new TypeError('桑基图场景与绘制命令不匹配。')
  const groupId = `${chartId}:sankey-content`, elements = chartGraphics(commands.slice(0, start), chartId)
  elements.push({ type: 'group', id: groupId, silent: true, children: chartGraphics(scene.commands, groupId, start) })
  for (const element of chartGraphics(commands.slice(end), `${chartId}:chrome`, end)) elements.push(element)
  return { elements, groupId, zOffset: start }
}

function radarComponent(data: ChartData, chart: ChartModel, context: ChartEChartsContext): RadarComponentOption {
  const { grid, color, axisColor, splitLineColor } = context
  const geometry = chartRadarGeometry(data, chart, { x: grid.left, y: grid.top, width: chart.size.width - grid.left - grid.right, height: chart.size.height - grid.top - grid.bottom })
  const label = chartRadarAxisLabel(chart), line = chartRadarAxisLine(chart), axis = chartEChartsAxisStyle({ ...chart.yAxis, line }, axisColor, splitLineColor, true, chart.options?.appearance?.font)
  return { center: [geometry.center.x, geometry.center.y], radius: geometry.radius, startAngle: 90, splitNumber: geometry.ticks.length - 1,
    clockwise: !chart.xAxis.reverse, shape: chart.options?.radar?.shape ?? 'polygon', indicator: data.categories.map(name => ({ name, min: geometry.min, max: geometry.max })),
    axisNameGap: geometry.radius * 0.22, axisName: { show: chart.xAxis.label?.visible !== false, ...chartEChartsFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(chart.xAxis.label) }, color, 10) },
    axisLine: axis.axisLine, axisTick: axis.axisTick,
    axisLabel: { show: label.visible ?? true, showMinLabel: false, margin: 4, hideOverlap: true, rotate: chart.yAxis.label?.rotate ?? 0, ...chartEChartsFont(label, color, 12),
      formatter: (value: string | number) => formatChartNumber(Number(value), label.format) },
    splitLine: axis.splitLine, splitArea: { show: false } }
}

function radarSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): RadarSeriesOption[] {
  return data.series.map((source, index) => {
    const style = chartSeriesStyle(chart, source.field, index, context.palette, source.sourceName ?? source.name)
    return { type: 'radar', name: style.name ?? source.name, symbol: chartNativeMarkerSymbol(style.point?.shape, 'none'), symbolSize: style.point?.size ?? 4,
      lineStyle: chartEChartsLineStyle(style), itemStyle: chartEChartsMarkerStyle(style),
      ...(chart.options?.radar?.fill ? { areaStyle: { color: style.color === undefined ? undefined : chartColorRepresentative(style.color), opacity: style.fillOpacity ?? 0.25 } } : {}),
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
    ...(axis.majorUnit !== undefined ? { interval: axis.majorUnit, minInterval: axis.majorUnit, maxInterval: axis.majorUnit } : {}),
    ...style, min: axis.min, max: axis.max, nameGap: Math.max(style.nameGap, dimension === 'x' ? 28 : 42),
    axisPointer: { label: { formatter: params => `${formatChartNumber(Number(params.value), axis.format)}${percent && !axis.format ? '%' : ''}` } },
    ...(percent && !axis.format ? { axisLabel: { ...style.axisLabel, formatter: '{value}%' } } : {}),
    ...(dimension === 'x' && axis.gridLine?.visible === undefined ? { splitLine: { show: false } } : {}) }
}

function dateCategoryAxis(dates: ChartDateAxis, axis: ChartAxisOptions, axisColor: string, splitLineColor: string, inheritedFont?: ChartFontOptions): XAxisComponentOption {
  const style = chartEChartsAxisStyle(axis, axisColor, splitLineColor, false, inheritedFont)
  // Native content keeps the exact shared coordinate extent. Calendar chrome
  // is drawn once from dateTicks; IntervalScale cannot substitute elapsed-day
  // ticks for authored month/year ticks, or reinterpret an ordinal as a serial.
  return { ...style, type: 'value', min: dates.coordinateMin, max: dates.coordinateMax, scale: true,
    axisLabel: { ...style.axisLabel, show: false }, axisTick: { ...style.axisTick, show: false }, splitLine: { show: false }, axisPointer: { show: false } }
}

function pieSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): PieSeriesOption {
  const source = data.series[0]!, options = chart.options?.pie, { palette, color, grid } = context
  const scale = source.values.reduce<number>((max, value) => Math.max(max, value ?? 0), 0) || 1, total = source.values.reduce<number>((sum, value) => sum + Math.max(0, value ?? 0) / scale, 0), start = -(options?.startAngle ?? -90)
  const plot = { x: grid.left, y: grid.top, width: chart.size.width - grid.left - grid.right, height: chart.size.height - grid.top - grid.bottom }
  const radius = chartPieOuterRadius(plot, options?.explosion)
  const sectors = new Map(chartPieLayout(data, chart, plot).sectors.map(sector => [sector.part.index, sector]))
  return { type: 'pie', name: source.name, radius: chart.type === 'doughnut' ? [radius * (options?.doughnutHole ?? 0.55), radius] : radius,
    center: ['50%', '50%'], left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
    startAngle: start, ...(options?.isHalfPie ? { endAngle: start - 180 } : {}), ...(options?.rosePie ? { roseType: 'radius' } : {}),
    ...(options?.explosion ? { selectedMode: 'multiple', selectedOffset: radius * options.explosion / 100 } : {}),
    avoidLabelOverlap: true, stillShowZeroSum: false, showEmptyCircle: false, label: { color }, labelLine: { show: options?.showLabelLine !== false },
    data: data.categories.map((name, index) => {
      const value = Math.max(0, source.values[index] ?? 0), style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index, chart.type), label = { ...style.label, ...options?.label }
      const sector = sectors.get(index), gradient = createChartMarkGradient(chart, style, 'pie', {}, sector ? { x: sector.center.x - sector.outer, y: sector.center.y - sector.outer, width: sector.outer * 2, height: sector.outer * 2 } : undefined)
      return { name, value, selected: Boolean(options?.explosion), itemStyle: chartEChartsItemStyle(style, gradient),
        label: chartEChartsPointLabel({ visible: true, contentType: 'category', position: 'outside', ...label }, color, value, name, source.name, total ? (value / scale) / total * 100 : 0) }
    }) }
}

function scatterSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): ScatterSeriesOption[] {
  let maxSize = 1
  for (const source of data.series) for (const size of source.sizeValues ?? []) if (size !== null) maxSize = Math.max(maxSize, size)
  return data.series.map((source, seriesIndex) => {
    const style = chartSeriesStyle(chart, source.field, seriesIndex, context.palette, source.sourceName ?? source.name), name = style.name ?? source.name, bubble = chart.type === 'bubble'
    const axis: ChartAxisOptions = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const order = context.dateAxis ? chartDateSeriesOrder(context.dateAxis, chart) : source.values.map((_value, index) => index)
    return { type: 'scatter', name, xAxisIndex: 0, yAxisIndex: style.rightAxis ? 1 : 0, symbol: chartNativeMarkerSymbol(style.point?.shape), symbolSize: style.point?.size ?? 8,
      itemStyle: chartEChartsMarkerStyle(style), label: chartEChartsLabel(style.label, context.color),
      data: order.flatMap<NonNullable<ScatterSeriesOption['data']>[number]>(index => {
        const value = source.values[index] ?? null, x = context.dateAxis ? context.dateAxis.coordinates[index] : source.xValues?.[index], size = source.sizeValues?.[index]
        if (value === null || x == null || !Number.isFinite(value) || !Number.isFinite(x) || chart.xAxis.scaleType === 'log' && x <= 0 || axis.scaleType === 'log' && value <= 0 || bubble && (size == null || size <= 0)) return []
        const pointStyle = chartPointStyle(style, index, chart.type)
        if (!context.dateAxis && !bubble && !style.dataPoints?.[index] && !style.label?.visible) return [[x, value]]
        const diameter = bubble ? chartBubblePointSize(pointStyle, size!, maxSize, chart.options, Math.min(chart.size.width - context.grid.left - context.grid.right, chart.size.height - context.grid.top - context.grid.bottom)) : pointStyle.point?.size ?? 8
        return [{ ...(context.dateAxis ? { id: `${chart.id}:series:${seriesIndex}:point:${index}` } : {}), name: data.categories[index], value: [x, value, ...(bubble ? [size!] : [])], symbol: chartNativeMarkerSymbol(pointStyle.point?.shape), symbolSize: diameter, itemStyle: chartEChartsMarkerStyle(pointStyle),
          ...(pointStyle.label?.visible ? { label: chartEChartsPointLabel(pointStyle.label, context.color, value, data.categories[index] ?? '', name, undefined, bubble ? size! : undefined) } : {}) }]
      }) }
  })
}

function categorySeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext, businessPercentages?: readonly number[]): Array<BarSeriesOption | LineSeriesOption> {
  const horizontal = resolveChartType(chart.type).family === 'bar', grid = context.grid
  const plot = { x: grid.left, y: grid.top, width: chart.size.width - grid.left - grid.right, height: chart.size.height - grid.top - grid.bottom }
  return data.series.map((source, index) => {
    const kind = chartSeriesKind(chart, source.field, index, source.sourceName ?? source.name), style = chart.type === 'pareto'
      ? chartParetoBarStyle(chart, source.field, index, context.palette, source.sourceName ?? source.name) : chartSeriesStyle(chart, source.field, index, context.palette, source.sourceName ?? source.name), name = style.name ?? source.name
    const itemStyle = chart.type === 'pareto' ? chartEChartsIndependentItemStyle : chartEChartsItemStyle
    const axis: ChartAxisOptions = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const isBar = kind.type === 'column' || kind.type === 'bar' || chart.type === 'histogram' || chart.type === 'pareto'
    const markerGradient = createChartMarkGradient(chart, { ...style, color: isChartLinearGradient(style.color) ? style.color : style.point?.color ?? style.color }, 'line', { reverse: axis.reverse, symbol: true }, plot)
    const order = context.dateAxis ? chartDateSeriesOrder(context.dateAxis, chart) : source.values.map((_value, categoryIndex) => categoryIndex)
    const values = order.map(categoryIndex => {
      const value = source.values[categoryIndex] ?? null, date = context.dateAxis?.coordinates[categoryIndex]
      if (value === null || !Number.isFinite(value) || axis.scaleType === 'log' && value <= 0 || context.dateAxis && date == null) return null
      const datum = context.dateAxis ? horizontal ? [value, date!] : [date!, value] : value
      const point = chartPointStyle(style, categoryIndex, chart.type)
      const gradient = isBar ? createChartMarkGradient(chart, point, horizontal ? 'bar' : 'column', { negative: value < 0, reverse: axis.reverse })
        : createChartMarkGradient(chart, { ...point, color: isChartLinearGradient(point.color) ? point.color : point.point?.color ?? point.color }, 'line', { reverse: axis.reverse, symbol: true }, plot)
      if (!context.dateAxis && !style.dataPoints?.[categoryIndex] && !style.label?.visible && !gradient) return datum
      return { ...(context.dateAxis ? { id: `${chart.id}:series:${index}:point:${categoryIndex}`, name: data.categories[categoryIndex] } : {}), value: datum,
        itemStyle: isBar ? itemStyle(point, gradient) : chartEChartsMarkerStyle(point, gradient), symbol: point.point?.shape === undefined ? undefined : chartNativeMarkerSymbol(point.point.shape), symbolSize: point.point?.size,
        ...(point.label?.visible ? { label: chartEChartsPointLabel(point.label, context.color, value, data.categories[categoryIndex] ?? '', name, businessPercentages?.[categoryIndex]) } : {}) }
    })
    const stacked = kind.stacked ? chart.type === 'combination' ? `${kind.type}-${style.rightAxis ? 'right' : 'left'}` : style.rightAxis ? 'total-right' : 'total' : undefined
    if (isBar) return {
      type: 'bar', name, data: values, stack: stacked,
      ...(resolveChartType(chart.type).family === 'bar' ? { xAxisIndex: style.rightAxis ? 1 : 0, yAxisIndex: 0 } : style.rightAxis ? { yAxisIndex: 1 } : {}),
      itemStyle: itemStyle(style, createChartMarkGradient(chart, style, horizontal ? 'bar' : 'column', { reverse: axis.reverse })), label: chartEChartsLabel(style.label, context.color),
      barCategoryGap: chart.type === 'histogram' ? `${chartHistogramBinGap(chart.options?.histogram) * 100}%` : `${chart.options?.bar?.gapWidth ?? 150}%`,
      ...(chart.options?.bar?.overlap === undefined ? {} : { barGap: `${-chart.options.bar.overlap}%` }),
    }
    const areaGradient = kind.type === 'area' ? createChartMarkGradient(chart, { ...style, fillOpacity: style.fillOpacity ?? 0.28 }, 'area', { reverse: axis.reverse }, plot) : undefined
    const lineGradient = kind.type === 'line' ? createChartMarkGradient(chart, { ...style, color: style.border?.color ?? style.color, fillOpacity: style.border?.opacity ?? 1 }, 'line', { reverse: axis.reverse }, plot) : undefined
    return { type: 'line', name, data: values, stack: stacked, yAxisIndex: style.rightAxis ? 1 : 0,
      showSymbol: style.point?.shape !== 'none' || Object.values(style.dataPoints ?? {}).some(point => point?.shape !== undefined && point.shape !== 'none'),
      symbol: chartNativeMarkerSymbol(style.point?.shape), symbolSize: style.point?.size ?? 6, connectNulls: chart.options?.invalidValueStrategy === 'link',
      smooth: kind.type === 'area' && chart.options?.area?.smooth, lineStyle: chartEChartsLineStyle(style, lineGradient), itemStyle: chartEChartsMarkerStyle(style, markerGradient), label: chartEChartsLabel(style.label, context.color),
      ...(kind.type === 'area' ? { areaStyle: { color: chartEChartsAutomaticGradient(areaGradient) ?? (style.color === undefined ? undefined : chartColorRepresentative(style.color)), opacity: areaGradient ? 1 : style.fillOpacity ?? 0.28 } } : {}) }
  })
}
