import type { CandlestickSeriesOption } from 'echarts/charts'
import type { ChartModel } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { ChartEChartsContext } from './echarts-special-option'
import { chartCandlestickState, chartCandlestickStyle, chartCandlestickWick, chartCandlestickWidth, type ChartCandlestickResolvedStyle } from './chart-candlestick'
import { chartColorOpacity } from './chart-geometry'
import { chartEChartsItemStyle } from './echarts-style'

/** Native ECharts rewrites color/outline by sign. Every datum carries both
 * signs and borderColorDoji so equal-price rows retain their explicit state. */
export function chartCandlestickEChartsStyle(style: ChartCandlestickResolvedStyle): NonNullable<CandlestickSeriesOption['itemStyle']> {
  const fill = style.hollow ? chartColorOpacity(style.color!, 0) : chartColorOpacity(style.color!, style.fillOpacity ?? 1)
  const border = chartColorOpacity(style.border?.color ?? style.color!, style.border?.opacity ?? 1)
  return { ...chartEChartsItemStyle(style), opacity: 1, color: fill, color0: fill, borderColor: border, borderColor0: border,
    borderColorDoji: border, borderWidth: style.border?.visible === false ? 0 : style.border?.width ?? 1 }
}

export function createChartCandlestickEChartsSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): CandlestickSeriesOption {
  const source = data.series[0]!, sourceName = source.sourceName ?? source.name
  const rising = chartCandlestickStyle(chart, source.field, undefined, 'rising', context.palette, sourceName)
  const falling = chartCandlestickEChartsStyle(chartCandlestickStyle(chart, source.field, undefined, 'falling', context.palette, sourceName))
  const doji = chartCandlestickEChartsStyle(chartCandlestickStyle(chart, source.field, undefined, 'doji', context.palette, sourceName))
  const plotWidth = chart.size.width - context.grid.left - context.grid.right
  return { type: 'candlestick', name: rising.name ?? source.name, clip: true, large: false,
    barWidth: chartCandlestickWidth(plotWidth / Math.max(1, data.categories.length), chart.options?.candlestick),
    data: (data.candles ?? []).map(candle => {
      if (!candle || candle.open === undefined || [candle.open, candle.high, candle.low, candle.close]
        .some(value => !Number.isFinite(value) || chart.yAxis.scaleType === 'log' && value <= 0)) return []
      const style = chartCandlestickStyle(chart, source.field, candle.categoryIndex, chartCandlestickState(candle), context.palette, sourceName)
      const wick = chartCandlestickWick(candle, chart.options?.candlestick)
      return { value: [candle.open, candle.close, wick.low, wick.high], itemStyle: chartCandlestickEChartsStyle(style) }
    }),
    itemStyle: { ...chartCandlestickEChartsStyle(rising), color0: falling.color0, borderColor0: falling.borderColor0,
      borderColorDoji: doji.borderColorDoji },
  }
}
