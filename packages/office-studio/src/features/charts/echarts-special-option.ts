import type { BoxplotSeriesOption, CandlestickSeriesOption, FunnelSeriesOption, GaugeSeriesOption, HeatmapSeriesOption, ScatterSeriesOption, SunburstSeriesOption, TreemapSeriesOption } from 'echarts/charts'
import type { ChartModel } from '../../domain/model'
import type { ChartDateAxis } from './chart-date-axis'
import type { ChartData } from './chart-engine'
import type { ChartDerivedData } from './chart-derived'
import type { ChartEChartsOption } from './echarts-option'
import { chartCandlestickPointStyle, chartHeatmapExtent, chartHeatmapPointStyle, chartHierarchyPointStyle, chartPointStyle } from './chart-derived'
import { chartHierarchyItems, type ChartHierarchyNode } from './chart-data'
import { chartSeriesStyle } from './chart-options'
import { chartEChartsFont, chartEChartsItemStyle, chartEChartsMarkerStyle, chartEChartsPointLabel } from './echarts-style'
import { chartColorOpacity, chartLabelText, formatChartNumber } from './chart-geometry'
import { chartGaugeTiles } from './chart-pie-layout'
import { chartGaugeAxisLabelRadius, chartGaugeColorStops, chartGaugeDetailText, chartGaugePointerIcon, chartGaugeRatio, chartGaugeSweep, resolveChartGaugeOptions } from './chart-gauge'
import { chartBoxplotAxis, chartBoxplotEChartsStyle, chartBoxplotStyle, chartBoxplotSummaryDrawable, chartBoxplotValueDrawable } from './chart-boxplot'

export interface ChartEChartsContext {
  palette: string[]
  color: string
  axisColor: string
  splitLineColor: string
  backgroundColor?: string
  dateAxis?: ChartDateAxis
  grid: { left: number; right: number; top: number; bottom: number; containLabel: boolean }
}

interface HierarchyEChartsNode {
  id: string
  name: string
  value: number
  children?: HierarchyEChartsNode[]
  itemStyle: ReturnType<typeof chartEChartsItemStyle> & { borderColor: string; borderWidth: number; gapWidth: number }
  label: ReturnType<typeof chartEChartsFont> & { show: boolean; formatter: string }
  upperLabel?: ReturnType<typeof chartEChartsFont> & { show: boolean; formatter: string; position: 'inside' | 'insideTop'; height: number }
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

/** ECharts does not inherit Sunburst parent styles; explicitly style every source node. */
function chartHierarchyEChartsData(data: ChartData, chart: ChartModel, context: ChartEChartsContext): HierarchyEChartsNode[] {
  const hierarchy = data.hierarchy ?? [], items = new Map(chartHierarchyItems(hierarchy).map(item => [item.node, item])), source = data.series[0]!
  const scale = hierarchy.reduce((max, node) => Math.max(max, node.value), 0) || 1, total = hierarchy.reduce((sum, node) => sum + node.value / scale, 0)
  const treemap = chart.type === 'treemap', options = treemap ? chart.options?.treemap?.label : chart.options?.sunburst?.label
  function nodeOption(node: ChartHierarchyNode): HierarchyEChartsNode {
    const item = items.get(node)!, style = chartHierarchyPointStyle(chart, source.field, item, context.palette, source.sourceName ?? source.name)
    const label = { contentType: 'category' as const, ...style.label, ...options }
    const formatter = chartLabelText(label, node.value, node.name, source.name, total ? (node.value / scale) / total * 100 : 0)
    const font = chartEChartsFont(label, context.color, 10)
    return { id: `${chart.id}:node:${item.pointIndex}`, name: node.name, value: node.value,
      itemStyle: { ...chartEChartsItemStyle(style), borderColor: chartColorOpacity(style.border?.color ?? chart.options?.appearance?.background ?? 'transparent', style.border?.opacity ?? 1), borderWidth: style.border?.visible === false ? 0 : style.border?.width ?? 1, gapWidth: 2 },
      label: { ...font, show: label.visible !== false, formatter },
      ...(treemap ? { upperLabel: { ...font, show: label.visible !== false && chart.options?.treemap?.parentLabelLayout !== 'none', formatter,
        position: chart.options?.treemap?.parentLabelLayout === 'overlapping' ? 'inside' as const : 'insideTop' as const, height: 18 } } : {}),
      ...(node.children?.length ? { children: node.children.map(nodeOption) } : {}),
    }
  }
  return hierarchy.map(nodeOption)
}

export function chartSpecialEChartsOption(data: ChartData, derived: ChartDerivedData, chart: ChartModel, context: ChartEChartsContext): Partial<ChartEChartsOption> | undefined {
  if (!data.categories.length || !data.series.length) return { xAxis: undefined, yAxis: undefined, grid: undefined, legend: { show: false }, series: [] }
  const { palette, color, grid } = context
  switch (chart.type) {
    case 'funnel': {
      const source = data.series[0]!, options = chart.options?.funnel
      const series: FunnelSeriesOption = { type: 'funnel', name: source.name, left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
        sort: options?.sort ?? 'descending', gap: options?.gap ?? 2, label: { show: true, ...chartEChartsFont(undefined, color, 10) },
        data: data.categories.map((name, index) => {
          const style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index), value = options?.useAbsoluteValue ? Math.abs(source.values[index] ?? 0) : source.values[index] ?? 0
          return { name, value, itemStyle: chartEChartsItemStyle(style), label: chartEChartsPointLabel({ visible: true, contentType: 'category', ...style.label }, color, value, name, source.name) }
        }),
      }
      return { xAxis: undefined, yAxis: undefined, grid: undefined, series: [series] }
    }
    case 'heatmap': {
      const options = chart.options?.heatmap, extent = chartHeatmapExtent(data, chart)
      const series: HeatmapSeriesOption = { type: 'heatmap', name: data.series[0]?.name, data: data.series.flatMap((source, row) => source.values.flatMap((value, column) => {
        if (value === null) return []
        const style = chartHeatmapPointStyle(chart, source.field, row, column, value, extent, palette, source.sourceName ?? source.name), label = { ...style.label, ...options?.label }
        return [{ value: [column, row, value], itemStyle: { ...chartEChartsItemStyle(style), color: chartColorOpacity(style.color!, style.fillOpacity ?? 1), opacity: 1 },
          label: { show: label.visible ?? false, ...chartEChartsFont(label, color, 10), formatter: `${chartLabelText(label, value, data.categories[column]!, source.name)}${options?.unit ?? ''}` } }]
      })), emphasis: { itemStyle: { borderColor: context.axisColor, borderWidth: 1 } } }
      return { yAxis: { type: 'category', data: data.series.map((series, index) => chartSeriesStyle(chart, series.field, index, palette, series.sourceName ?? series.name).name ?? series.name), inverse: chart.yAxis.reverse !== true,
        axisLabel: { ...chartEChartsFont(chart.yAxis.label, context.axisColor, 10), show: chart.yAxis.label?.visible !== false }, axisLine: { show: chart.yAxis.line?.visible !== false, lineStyle: { color: chart.yAxis.line?.color ?? context.axisColor } }, splitArea: { show: false } },
      visualMap: { type: options?.visualMapType ?? 'continuous', min: extent.min, max: extent.max, calculable: false, show: false, inRange: { color: palette } }, series: [series] }
    }
    case 'boxplot': {
      const boxes = derived.boxes!, series: BoxplotSeriesOption = { type: 'boxplot', name: '样本分布', boxWidth: ['48%', '48%'], data: boxes.map((box, index) => {
        const statistics = box.statistics, style = chartBoxplotStyle(chart, box, index, palette)
        return { name: box.name, value: statistics && chartBoxplotSummaryDrawable(statistics, chartBoxplotAxis(chart, style)) ? [statistics.min, statistics.q1, statistics.median, statistics.q3, statistics.max] : [], itemStyle: chartBoxplotEChartsStyle(style) }
      }) }
      const outliers: ScatterSeriesOption = { type: 'scatter', name: '异常值', data: boxes.flatMap((box, index) => {
        if (chart.options?.boxplot?.showOutliers === false) return []
        const style = chartBoxplotStyle(chart, box, index, palette), axis = chartBoxplotAxis(chart, style)
        const markerStyle = chartBoxplotEChartsStyle({ ...style, color: style.point?.color ?? style.color, fillOpacity: style.fillOpacity ?? 1 })
        return (box.statistics?.outliers ?? []).filter(value => chartBoxplotValueDrawable(value, axis)).map(value => ({ name: box.name, value: [index, value], itemStyle: markerStyle, symbol: style.point?.shape ?? 'circle', symbolSize: style.point?.size ?? 4 }))
      }), symbolSize: 4 }
      return { legend: { show: chart.legend.visible, data: ['样本分布', ...(chart.options?.boxplot?.showOutliers === false ? [] : ['异常值'])] }, series: [series, outliers] }
    }
    case 'candlestick': {
      const options = chart.options?.candlestick, series: CandlestickSeriesOption = { type: 'candlestick', name: data.series[0]!.name,
        data: (data.candles ?? []).map(candle => {
          if (!candle || candle.open === undefined) return []
          const rising = candle.close >= candle.open, style = chartCandlestickPointStyle(chart, data.series[0]!.field, candle.categoryIndex, rising, palette, data.series[0]!.sourceName ?? data.series[0]!.name)
          const fill = chartColorOpacity(style.color!, style.fillOpacity ?? 1), border = chartColorOpacity(style.border?.color ?? style.color!, style.border?.opacity ?? 1)
          return { value: [candle.open, candle.close, candle.low, candle.high], itemStyle: { ...chartEChartsItemStyle(style), opacity: 1,
            color: options?.hollow && rising ? 'transparent' : fill, color0: fill, borderColor: border, borderColor0: border } }
        }),
        ...(options?.candleWidth === undefined ? {} : { barWidth: options.candleWidth }),
        itemStyle: { color: options?.hollow ? 'transparent' : options?.risingColor ?? '#dc2626', color0: options?.fallingColor ?? '#059669', borderColor: options?.risingColor ?? '#dc2626', borderColor0: options?.fallingColor ?? '#059669' },
      }
      return { series: [series] }
    }
    case 'treemap': {
      const options = chart.options?.treemap, series: TreemapSeriesOption = { type: 'treemap', name: data.series[0]!.name, left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
        data: chartHierarchyEChartsData(data, chart, context), roam: false, nodeClick: false, breadcrumb: { show: false },
        label: { show: options?.label?.visible !== false, ...chartEChartsFont(options?.label, color, 10), formatter: options?.label?.contentType === 'value' ? '{c}' : '{b}' },
        upperLabel: { show: options?.parentLabelLayout !== 'none', position: options?.parentLabelLayout === 'overlapping' ? 'inside' : 'insideTop', height: 18, ...chartEChartsFont(options?.label, color, 10) },
        itemStyle: { borderWidth: 1, gapWidth: 2, borderColor: chart.options?.appearance?.background ?? 'transparent' },
      }
      return { xAxis: undefined, yAxis: undefined, grid: undefined, series: [series] }
    }
    case 'sunburst': {
      const options = chart.options?.sunburst, series: SunburstSeriesOption = { type: 'sunburst', name: data.series[0]!.name, data: chartHierarchyEChartsData(data, chart, context),
        center: [(chart.size.width + grid.left - grid.right) / 2, (chart.size.height + grid.top - grid.bottom) / 2], radius: ['16%', '74%'], sort: null, nodeClick: false,
        label: { show: options?.label?.visible !== false, rotate: 'radial', ...chartEChartsFont(options?.label, color, 10), formatter: options?.label?.contentType === 'value' ? '{c}' : '{b}' },
        itemStyle: { borderWidth: 1, borderColor: chart.options?.appearance?.background ?? 'transparent' },
      }
      return { xAxis: undefined, yAxis: undefined, grid: undefined, series: [series], ...(options?.centerLabel?.visible ? { graphic: { elements: [{ type: 'text', x: (chart.size.width + grid.left - grid.right) / 2, y: (chart.size.height + grid.top - grid.bottom) / 2,
        style: { text: data.series[0]!.name, fill: options.centerLabel.font?.color ?? color, fontSize: options.centerLabel.font?.fontSize ?? 11, fontFamily: options.centerLabel.font?.family, align: 'center', verticalAlign: 'middle' } }] } } : {}) }
    }
    case 'gauge': {
      const authored = chart.options?.gauge, source = data.series[0]!
      const tiles = chartGaugeTiles(data, { x: grid.left, y: grid.top, width: Math.max(1, chart.size.width - grid.left - grid.right), height: Math.max(1, chart.size.height - grid.top - grid.bottom) })
      const series: GaugeSeriesOption[] = tiles.flatMap(tile => {
        const index = tile.categoryIndex, value = source.values[index]!
        const name = data.categories[index]!, radius = tile.radius, style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index), label = style.label
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
