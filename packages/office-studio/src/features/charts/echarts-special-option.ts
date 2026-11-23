import type { BoxplotSeriesOption, CandlestickSeriesOption, FunnelSeriesOption, GaugeSeriesOption, HeatmapSeriesOption, ScatterSeriesOption, SunburstSeriesOption, TreemapSeriesOption } from 'echarts/charts'
import type { ChartModel } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { ChartDerivedData } from './chart-derived'
import type { ChartEChartsOption } from './echarts-option'
import { chartPointStyle } from './chart-derived'
import { chartSeriesStyle } from './chart-options'
import { chartEChartsFont, chartEChartsItemStyle, chartEChartsPointLabel } from './echarts-style'
import { formatChartNumber } from './chart-geometry'

export interface ChartEChartsContext {
  palette: string[]
  color: string
  axisColor: string
  splitLineColor: string
  grid: { left: number; right: number; top: number; bottom: number; containLabel: boolean }
}

export function chartSpecialEChartsOption(data: ChartData, derived: ChartDerivedData, chart: ChartModel, context: ChartEChartsContext): Partial<ChartEChartsOption> | undefined {
  const { palette, color, grid } = context
  switch (chart.type) {
    case 'funnel': {
      const source = data.series[0]!, options = chart.options?.funnel
      const series: FunnelSeriesOption = { type: 'funnel', name: source.name, left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
        sort: options?.sort ?? 'descending', gap: options?.gap ?? 2, label: { show: true, ...chartEChartsFont(undefined, color, 10) },
        data: data.categories.map((name, index) => {
          const style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette), index), value = options?.useAbsoluteValue ? Math.abs(source.values[index] ?? 0) : source.values[index] ?? 0
          return { name, value, itemStyle: chartEChartsItemStyle(style), label: chartEChartsPointLabel({ visible: true, contentType: 'category', ...style.label }, color, value, name, source.name) }
        }),
      }
      return { xAxis: undefined, yAxis: undefined, grid: undefined, series: [series] }
    }
    case 'heatmap': {
      const options = chart.options?.heatmap, valid = data.series.flatMap(series => series.values.filter((value): value is number => value !== null))
      let min = 0, max = 1
      for (const value of valid) { min = Math.min(min, value); max = Math.max(max, value) }
      const series: HeatmapSeriesOption = { type: 'heatmap', name: data.series[0]?.name, data: data.series.flatMap((source, row) => source.values.flatMap((value, column) => value === null ? [] : [{
        value: [column, row, value], label: { show: options?.label?.visible ?? false, ...chartEChartsFont(options?.label, color, 10), formatter: `${formatChartNumber(value, options?.label?.format)}${options?.unit ?? ''}` },
      }])), emphasis: { itemStyle: { borderColor: context.axisColor, borderWidth: 1 } } }
      return { yAxis: { type: 'category', data: data.series.map((series, index) => chartSeriesStyle(chart, series.field, index, palette).name ?? series.name), inverse: chart.yAxis.reverse !== true,
        axisLabel: { ...chartEChartsFont(chart.yAxis.label, context.axisColor, 10), show: chart.yAxis.label?.visible !== false }, axisLine: { show: chart.yAxis.line?.visible !== false, lineStyle: { color: chart.yAxis.line?.color ?? context.axisColor } }, splitArea: { show: false } },
      visualMap: { type: options?.visualMapType ?? 'continuous', min: options?.min ?? min, max: options?.max ?? max, calculable: false, show: false, inRange: { color: palette } }, series: [series] }
    }
    case 'boxplot': {
      const boxes = derived.boxes!, series: BoxplotSeriesOption = { type: 'boxplot', name: '样本分布', data: boxes.map((box, index) => {
        const statistics = box.statistics, style = chartSeriesStyle(chart, box.field, index, palette)
        return { name: box.name, value: statistics ? [statistics.min, statistics.q1, statistics.median, statistics.q3, statistics.max] : [], itemStyle: chartEChartsItemStyle(style) }
      }) }
      const outliers: ScatterSeriesOption = { type: 'scatter', name: '异常值', data: boxes.flatMap((box, index) => chart.options?.boxplot?.showOutliers === false ? [] : (box.statistics?.outliers ?? []).map(value => ({ value: [index, value], itemStyle: { color: chartSeriesStyle(chart, box.field, index, palette).color } }))), symbolSize: 4 }
      return { legend: { show: chart.legend.visible, data: ['样本分布', ...(chart.options?.boxplot?.showOutliers === false ? [] : ['异常值'])] }, series: [series, outliers] }
    }
    case 'candlestick': {
      const options = chart.options?.candlestick, series: CandlestickSeriesOption = { type: 'candlestick', name: data.series[0]!.name,
        data: (data.candles ?? []).map(candle => candle?.open === undefined ? [] : [candle.open, candle.close, candle.low, candle.high]),
        ...(options?.candleWidth === undefined ? {} : { barWidth: options.candleWidth }),
        itemStyle: { color: options?.hollow ? 'transparent' : options?.risingColor ?? '#dc2626', color0: options?.fallingColor ?? '#059669', borderColor: options?.risingColor ?? '#dc2626', borderColor0: options?.fallingColor ?? '#059669' },
      }
      return { series: [series] }
    }
    case 'treemap': {
      const options = chart.options?.treemap, series: TreemapSeriesOption = { type: 'treemap', name: data.series[0]!.name, left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
        data: data.hierarchy, roam: false, nodeClick: false, breadcrumb: { show: false },
        label: { show: options?.label?.visible !== false, ...chartEChartsFont(options?.label, color, 10), formatter: options?.label?.contentType === 'value' ? '{c}' : '{b}' },
        upperLabel: { show: options?.parentLabelLayout !== 'none', position: options?.parentLabelLayout === 'overlapping' ? 'inside' : 'insideTop', height: 18, ...chartEChartsFont(options?.label, color, 10) },
        itemStyle: { borderWidth: 1, gapWidth: 2, borderColor: chart.options?.appearance?.background ?? 'transparent' },
      }
      return { xAxis: undefined, yAxis: undefined, grid: undefined, series: [series] }
    }
    case 'sunburst': {
      const options = chart.options?.sunburst, series: SunburstSeriesOption = { type: 'sunburst', name: data.series[0]!.name, data: data.hierarchy,
        center: [(chart.size.width + grid.left - grid.right) / 2, (chart.size.height + grid.top - grid.bottom) / 2], radius: ['16%', '74%'], sort: null, nodeClick: false,
        label: { show: options?.label?.visible !== false, rotate: 'radial', ...chartEChartsFont(options?.label, color, 10), formatter: options?.label?.contentType === 'value' ? '{c}' : '{b}' },
        itemStyle: { borderWidth: 1, borderColor: chart.options?.appearance?.background ?? 'transparent' },
      }
      return { xAxis: undefined, yAxis: undefined, grid: undefined, series: [series], ...(options?.centerLabel?.visible ? { graphic: { elements: [{ type: 'text', x: (chart.size.width + grid.left - grid.right) / 2, y: (chart.size.height + grid.top - grid.bottom) / 2,
        style: { text: data.series[0]!.name, fill: options.centerLabel.font?.color ?? color, fontSize: options.centerLabel.font?.fontSize ?? 11, fontFamily: options.centerLabel.font?.family, align: 'center', verticalAlign: 'middle' } }] } } : {}) }
    }
    case 'gauge': {
      const options = chart.options?.gauge, min = options?.min ?? 0, max = options?.max ?? 100, columns = Math.max(1, Math.ceil(Math.sqrt(data.categories.length * chart.size.width / Math.max(1, chart.size.height)))), rows = Math.max(1, Math.ceil(data.categories.length / columns))
      const width = Math.max(1, chart.size.width - grid.left - grid.right), height = Math.max(1, chart.size.height - grid.top - grid.bottom), radius = Math.min(width / columns, height / rows) * 0.4
      const ranges: Array<[number, string]> = [], sorted = [...(options?.ranges ?? [])].sort((a, b) => a.from - b.from)
      let cursor = min
      for (const range of sorted) {
        if (range.from > cursor) ranges.push([Math.max(0, Math.min(1, (range.from - min) / (max - min))), context.splitLineColor])
        ranges.push([Math.max(0, Math.min(1, (range.to - min) / (max - min))), range.color]); cursor = range.to
      }
      if (cursor < max) ranges.push([1, context.splitLineColor])
      const series: GaugeSeriesOption[] = data.categories.map((name, index) => ({ type: 'gauge', name,
        center: [grid.left + (index % columns + 0.5) * width / columns, grid.top + (Math.floor(index / columns) + 0.5) * height / rows], radius,
        min, max, startAngle: options?.startAngle ?? 225, endAngle: options?.endAngle ?? -45, splitNumber: options?.splitNumber ?? 5,
        axisLine: { lineStyle: { color: ranges.length ? ranges : [[1, context.splitLineColor]], width: Math.max(2, radius * 0.1) } },
        axisLabel: { color: context.axisColor, fontSize: Math.min(10, radius * 0.16) },
        pointer: { show: options?.pointer?.visible !== false && options?.mode !== 'progress', length: `${(options?.pointer?.length ?? 0.7) * 100}%`, width: options?.pointer?.width ?? 4, itemStyle: { color: options?.pointer?.color ?? palette[index % palette.length] } },
        progress: { show: options?.progress?.visible ?? options?.mode === 'progress', width: options?.progress?.width ?? Math.max(2, radius * 0.1), roundCap: options?.progress?.rounded, itemStyle: { color: options?.progress?.color ?? palette[index % palette.length] } },
        title: { show: true, ...chartEChartsFont(options?.label, color, 10), offsetCenter: [0, '60%'] },
        detail: { show: options?.label?.visible !== false, ...chartEChartsFont(options?.label, color, Math.min(18, radius * 0.28)), valueAnimation: false, offsetCenter: [0, '30%'], formatter: (value: number) => formatChartNumber(value, options?.label?.format) },
        data: data.series[0]!.values[index] === null ? [] : [{ name, value: data.series[0]!.values[index] ?? 0 }],
      }))
      return { xAxis: undefined, yAxis: undefined, grid: undefined, series }
    }
    default: return undefined
  }
}
