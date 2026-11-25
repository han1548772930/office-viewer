import type { BoxplotSeriesOption, FunnelSeriesOption, GaugeSeriesOption, HeatmapSeriesOption, ScatterSeriesOption } from 'echarts/charts'
import type { ChartModel } from '../../domain/model'
import type { ChartDateAxis } from './chart-date-axis'
import type { ChartData } from './chart-engine'
import type { ChartDerivedData } from './chart-derived'
import type { ChartEChartsOption } from './echarts-option'
import { chartHeatmapExtent, chartHeatmapPointStyle, chartPointStyle } from './chart-derived'
import { chartSeriesStyle } from './chart-options'
import { chartEChartsAxisStyle, chartEChartsFont, chartEChartsItemStyle, chartEChartsMarkerStyle, chartEChartsPointLabel } from './echarts-style'
import { chartColorOpacity, chartLabelText, formatChartNumber } from './chart-geometry'
import { chartGaugeTiles } from './chart-pie-layout'
import { chartGaugeAxisLabelRadius, chartGaugeColorStops, chartGaugeDetailText, chartGaugePointerIcon, chartGaugeRatio, chartGaugeSweep, resolveChartGaugeOptions } from './chart-gauge'
import { chartBoxplotAxis, chartBoxplotEChartsStyle, chartBoxplotOutliers, chartBoxplotOutlierStyle, chartBoxplotStyle, chartBoxplotSummaryDrawable, chartBoxplotValueDrawable } from './chart-boxplot'
import { chartNativeMarkerSymbol } from './chart-markers'
import { chartHeatmapCellRect, chartHeatmapFieldNames, chartHeatmapLabelAnchor, chartHeatmapLabelText, chartHeatmapValueVisible,
  createChartHeatmapScale, resolveChartHeatmapLabel, type ChartHeatmapScale } from './chart-heatmap'
import { createChartCandlestickEChartsSeries } from './echarts-candlestick-option'
import { chartTreemapTooltip, createChartTreemapEChartsSeries } from './echarts-treemap-option'
import { chartSunburstCenterGraphic, chartSunburstTooltip, createChartSunburstEChartsSeries } from './echarts-sunburst-option'
import { chartCanvasBoxplotRuntimeId, chartCanvasRuntimeSeriesId } from './chart-element-interaction'

export interface ChartEChartsContext {
  palette: string[]
  color: string
  axisColor: string
  splitLineColor: string
  backgroundColor?: string
  dateAxis?: ChartDateAxis
  heatmapScale?: ChartHeatmapScale
  grid: { left: number; right: number; top: number; bottom: number; containLabel: boolean }
}

/** A one-value native Gauge still assigns overlap-progress z2 from its value. Split
 * its visual layers so the progress cannot cover ticks, the anchor, pointer or text.
 * The hidden axis width in the progress layer is its authored annular width. */
function chartGaugeEChartsLayers(series: GaugeSeriesOption, id: string): GaugeSeriesOption[] {
  const layers: GaugeSeriesOption[] = []
  const hidden = { axisTick: { show: false }, splitLine: { show: false }, axisLabel: { show: false }, pointer: { show: false },
    anchor: { show: false }, title: { show: false }, detail: { show: false }, silent: true, tooltip: { show: false } }
  if (series.axisLine?.show) layers.push({ ...series, ...hidden, id: `${id}:axis`, z: 1, progress: { show: false } })
  if (series.progress?.show) layers.push({ ...series, ...hidden, id: `${id}:progress`, z: 2,
    axisLine: { show: false, lineStyle: { width: series.progress.width } }, progress: { ...series.progress, overlap: false } })
  layers.push({ ...series, id: `${id}:foreground`, z: 3, axisLine: { ...series.axisLine, show: false }, progress: { ...series.progress, show: false } })
  return layers
}

export function chartSpecialEChartsOption(data: ChartData, derived: ChartDerivedData, chart: ChartModel, context: ChartEChartsContext): Partial<ChartEChartsOption> | undefined {
  if (!data.categories.length || !data.series.length) return { xAxis: undefined, yAxis: undefined, grid: undefined, legend: { show: false }, series: [],
    ...(data.hierarchyDiagnostic ? { graphic: { elements: [{ type: 'text', x: chart.size.width / 2, y: chart.size.height / 2,
      style: { text: data.hierarchyDiagnostic, fill: context.color, fontSize: 12, align: 'center', verticalAlign: 'middle' } }] } } : {}) }
  const { palette, color, grid } = context
  switch (chart.type) {
    case 'funnel': {
      const source = data.series[0]!, options = chart.options?.funnel
      const series: FunnelSeriesOption = { type: 'funnel', id: chartCanvasRuntimeSeriesId(chart.id, 0), name: source.name, left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
        sort: options?.sort ?? 'descending', gap: options?.gap ?? 2, label: { show: true, ...chartEChartsFont(undefined, color, 10) },
        data: data.categories.map((name, index) => {
          const style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index, chart.type), value = options?.useAbsoluteValue ? Math.abs(source.values[index] ?? 0) : source.values[index] ?? 0
          return { name, value, itemStyle: chartEChartsItemStyle(style), label: chartEChartsPointLabel({ visible: true, contentType: 'category', ...style.label }, color, value, name, source.name) }
        }),
      }
      return { xAxis: undefined, yAxis: undefined, grid: undefined, series: [series] }
    }
    case 'heatmap': {
      const options = chart.options?.heatmap, extent = chartHeatmapExtent(data, chart), scale = context.heatmapScale ?? createChartHeatmapScale(extent, palette, options?.visualMapType)
      const plot = { x: grid.left, y: grid.top, width: Math.max(1, chart.size.width - grid.left - grid.right), height: Math.max(1, chart.size.height - grid.top - grid.bottom) }
      const series: HeatmapSeriesOption = { type: 'heatmap', name: data.series[0]?.name, progressive: 1000,
        data: data.series.flatMap((source, seriesPosition) => source.values.flatMap((value, categoryIndex) => {
          if (value === null || !Number.isFinite(value) || categoryIndex >= data.categories.length) return []
          const style = chartHeatmapPointStyle(chart, source.field, seriesPosition, categoryIndex, value, extent, palette, source.sourceName ?? source.name, scale)
          const label = resolveChartHeatmapLabel(chart, style.label, color), visible = chartHeatmapValueVisible(scale, value)
          const rect = chartHeatmapCellRect(data, chart, plot, seriesPosition, categoryIndex)!
          return [{ value: [seriesPosition, categoryIndex, value], visualMap: false,
            // ECharts visualSolution explicitly honors datum.visualMap=false.
            // This keeps shared eight-piece colors and author overrides intact.
            itemStyle: { ...chartEChartsItemStyle(style), color: chartColorOpacity(style.color!, style.fillOpacity ?? 1), opacity: visible ? 1 : 0 },
            label: { show: label.visible && visible, ...chartEChartsFont(label, color, 12), position: chartHeatmapLabelAnchor(rect, label).nativePosition,
              distance: 5, width: Math.max(1, rect.width - 4), overflow: 'truncate' as const, formatter: chartHeatmapLabelText(value, source.formattedValues?.[categoryIndex], label) } }]
        })), emphasis: { itemStyle: { borderColor: color, borderWidth: 1 } } }
      const inheritedFont = chart.options?.appearance?.font
      const xStyle = chartEChartsAxisStyle(chart.xAxis, context.axisColor, context.splitLineColor, false, inheritedFont)
      const yStyle = chartEChartsAxisStyle(chart.yAxis, context.axisColor, context.splitLineColor, false, inheritedFont)
      return { xAxis: { type: 'category', data: chartHeatmapFieldNames(data, chart), boundaryGap: true,
        ...xStyle, axisLabel: { ...xStyle.axisLabel, color: chart.xAxis.label?.color ?? color },
        nameTextStyle: chartEChartsFont({ ...inheritedFont, ...chart.xAxis.titleStyle }, color, 11), inverse: chart.xAxis.reverse === true },
      yAxis: { type: 'category', data: data.categories, boundaryGap: true,
        ...yStyle, axisLabel: { ...yStyle.axisLabel, color: chart.yAxis.label?.color ?? color },
        nameTextStyle: chartEChartsFont({ ...inheritedFont, ...chart.yAxis.titleStyle }, color, 11), inverse: chart.yAxis.reverse === true },
      legend: { show: false },
      // Visible color legend and hover forwarding use the shared Canvas scene.
      // This component supplies the native heatmap's required visualMap model.
      visualMap: { type: scale.type, min: extent.min, max: extent.max, dimension: 2, seriesIndex: 0, calculable: false, realtime: false,
        selectedMode: false, hoverLink: false, show: false, inRange: { color: scale.palette },
        ...(scale.type === 'piecewise' ? { splitNumber: 8, pieces: scale.pieces.map(piece => ({ ...(piece.lowerClosed ? { gte: piece.lower } : { gt: piece.lower }),
          lte: piece.upper, color: piece.color, label: piece.text })) } : {}) }, series: [series] }
    }
    case 'boxplot': {
      const boxes = derived.boxes!, series: BoxplotSeriesOption = { type: 'boxplot', id: chartCanvasBoxplotRuntimeId(chart.id, 'summary'), name: '样本分布', boxWidth: ['48%', '48%'], data: boxes.map((box, index) => {
        const statistics = box.statistics, style = chartBoxplotStyle(chart, box, index, palette)
        return { name: box.name, value: statistics && chartBoxplotSummaryDrawable(statistics, chartBoxplotAxis(chart, style)) ? [statistics.min, statistics.q1, statistics.median, statistics.q3, statistics.max] : [], itemStyle: chartBoxplotEChartsStyle(style) }
      }) }
      const outliers: ScatterSeriesOption = { type: 'scatter', id: chartCanvasBoxplotRuntimeId(chart.id, 'outliers'), name: '异常值', data: boxes.flatMap((box, index) => {
        if (chart.options?.boxplot?.showOutliers === false) return []
        const style = chartBoxplotStyle(chart, box, index, palette), axis = chartBoxplotAxis(chart, style)
        const items: NonNullable<ScatterSeriesOption['data']> = []
        for (const outlier of chartBoxplotOutliers(box)) {
          const pointStyle = chartBoxplotOutlierStyle(style, outlier.sourceIndex), size = pointStyle.point?.size ?? 4, shape = pointStyle.point?.shape ?? 'circle'
          if (!chartBoxplotValueDrawable(outlier.value, axis) || size <= 0 || shape === 'none') continue
          const markerStyle = chartBoxplotEChartsStyle({ ...pointStyle, color: pointStyle.point?.color ?? pointStyle.color, fillOpacity: pointStyle.fillOpacity ?? 1 })
          items.push({ id: `${chart.id}:box:${box.seriesPosition}:outlier:${outlier.sourceIndex}`, name: box.name, value: [index, outlier.value],
            itemStyle: markerStyle, symbol: chartNativeMarkerSymbol(shape), symbolSize: size })
        }
        return items
      }), symbolSize: 4 }
      return { legend: { show: chart.legend.visible, data: ['样本分布', ...(chart.options?.boxplot?.showOutliers === false ? [] : ['异常值'])] }, series: [series, outliers] }
    }
    case 'candlestick': {
      return { series: [createChartCandlestickEChartsSeries(data, chart, context)] }
    }
    case 'treemap': {
      return { xAxis: undefined, yAxis: undefined, grid: undefined, animation: false, legend: { show: false },
        tooltip: chartTreemapTooltip(data, context), series: [createChartTreemapEChartsSeries(data, chart, context)] }
    }
    case 'sunburst': {
      const graphic = chartSunburstCenterGraphic(data, chart, context)
      return { xAxis: undefined, yAxis: undefined, grid: undefined, animation: false, legend: { show: false },
        tooltip: chartSunburstTooltip(data, chart, context), series: [createChartSunburstEChartsSeries(data, chart, context)], ...(graphic ? { graphic } : {}) }
    }
    case 'gauge': {
      const authored = chart.options?.gauge, source = data.series[0]!
      const tiles = chartGaugeTiles(data, { x: grid.left, y: grid.top, width: Math.max(1, chart.size.width - grid.left - grid.right), height: Math.max(1, chart.size.height - grid.top - grid.bottom) })
      const series: GaugeSeriesOption[] = tiles.flatMap(tile => {
        const index = tile.categoryIndex, value = source.values[index]!
        const name = data.categories[index]!, radius = tile.radius, style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index, chart.type), label = style.label
        const options = resolveChartGaugeOptions(authored, radius, chart.options?.appearance?.font, label)
        const ratio = chartGaugeRatio(value, options.min, options.max), sweep = chartGaugeSweep(options)
        const pointerColor = chartColorOpacity(options.pointer.color ?? style.color!, style.fillOpacity ?? 1), progressColor = chartColorOpacity(options.progress.color ?? style.color!, style.fillOpacity ?? 1)
        const detailText = chartGaugeDetailText(value, name, style.name ?? source.name, { ...label, ...options.detail }, data.records?.[index]?.[source.field])
        const result: GaugeSeriesOption = { type: 'gauge', name, itemStyle: chartEChartsItemStyle(style),
          center: [tile.center.x, tile.center.y], radius,
          min: options.min, max: options.max, startAngle: options.startAngle, endAngle: options.endAngle, clockwise: true, splitNumber: options.splitNumber,
          axisLine: { show: options.axisLine.visible && options.axisLine.width > 0 && sweep.end > sweep.start, lineStyle: { color: chartGaugeColorStops(authored, options.axisLine.color ?? context.splitLineColor), width: options.axisLine.width } },
          axisTick: { show: options.axisTick.visible, splitNumber: options.axisTick.splitNumber, distance: 0, length: options.axisTick.length,
            lineStyle: { color: options.axisTick.color ?? context.axisColor, width: options.axisTick.width } },
          splitLine: { show: options.splitLine.visible, distance: 0, length: options.splitLine.length, lineStyle: { color: options.splitLine.color ?? context.axisColor, width: options.splitLine.width } },
          axisLabel: { show: options.axisLabel.visible, ...chartEChartsFont(options.axisLabel, context.axisColor), distance: radius - options.splitLine.length - chartGaugeAxisLabelRadius(options, radius),
            formatter: (current: number) => formatChartNumber(current, options.axisLabel.format) },
          pointer: { show: options.pointer.visible && options.pointer.width > 0 && options.pointer.length > 0, length: `${options.pointer.length * 100}%`, width: options.pointer.width,
            icon: chartGaugePointerIcon(options.pointer.shape), keepAspect: false, showAbove: false, itemStyle: { ...chartEChartsItemStyle(style), opacity: 1, color: pointerColor } },
          anchor: { show: options.anchor.visible && options.anchor.size > 0, size: options.anchor.size, showAbove: false, itemStyle: {
            color: options.anchor.color ?? (options.mode === 'pointer' ? pointerColor : chart.options?.appearance?.background ?? context.backgroundColor ?? '#ffffff'),
            borderColor: options.anchor.borderColor ?? pointerColor, borderWidth: options.anchor.borderWidth } },
          progress: { show: options.progress.visible && options.progress.width > 0 && ratio > 0 && sweep.end > sweep.start, width: options.progress.width, roundCap: options.progress.rounded, clip: true,
            itemStyle: { ...chartEChartsItemStyle(style), opacity: 1, color: progressColor } },
          title: { show: options.title.visible, ...chartEChartsFont(options.title, color), offsetCenter: [0, `${options.title.offsetY * 100}%`], width: radius * 1.7, overflow: 'truncate' },
          detail: { show: options.detail.visible, ...chartEChartsFont(options.detail, color), valueAnimation: false, offsetCenter: [0, `${options.detail.offsetY * 100}%`],
            width: radius * 1.7, overflow: 'truncate', formatter: () => detailText },
          tooltip: { formatter: () => `${name}: ${detailText}` },
          data: [{ name, value: Math.min(options.max, Math.max(options.min, value)) }],
        }
        return chartGaugeEChartsLayers(result, `${chart.id}:gauge:${index}`)
      })
      return { xAxis: undefined, yAxis: undefined, grid: undefined, legend: { data: tiles.map(tile => data.categories[tile.categoryIndex]!) }, tooltip: { trigger: 'item', renderMode: 'richText' }, series }
    }
    default: return undefined
  }
}
