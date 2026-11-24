import type { ChartModel } from '../../domain/model'
import type { DrawCommand } from '../../rendering/contracts'
import type { ChartData, ChartLayout } from './chart-engine'
import { chartCandlestickOverlayOptions, chartCandlestickState, chartCandlestickStyle, chartCandlestickWick, chartCandlestickWidth } from './chart-candlestick'
import { chartAxisRatio, chartColorOpacity, chartFont, chartLabelText, chartLine, chartRect, chartSymbol } from './chart-geometry'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { definedChartStyleFields } from './chart-options'

/** Both HLC range marks and OHLC overlays submit the same clipped Canvas scene.
 * HLC has no synthetic open/body; missing prices preserve an empty sample slot. */
export function drawChartCandlesticks(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout,
  palette: readonly string[], textColor: string): void {
  const source = data.series[0]
  if (!source) return
  const { plot } = layout, count = Math.max(1, data.categories.length)
  const width = chartCandlestickWidth(layout.categoryBandwidth ?? plot.width / count, chart.options?.candlestick)
  const y = (value: number): number => plot.y + (1 - chartAxisRatio(value, layout.min, layout.max, chart.yAxis)) * plot.height
  for (const candle of data.candles ?? []) {
    if (!candle || [candle.high, candle.low, candle.close, ...(candle.open === undefined ? [] : [candle.open])]
      .some(value => !Number.isFinite(value) || chart.yAxis.scaleType === 'log' && value <= 0)) continue
    const index = candle.categoryIndex, ratio = layout.categoryPositions ? layout.categoryPositions[index]
      : chart.xAxis.reverse ? 1 - (index + 0.5) / count : (index + 0.5) / count
    if (ratio == null || !Number.isFinite(ratio)) continue
    const x = plot.x + ratio * plot.width, close = y(candle.close), mode = candle.open === undefined ? 'hlc' : 'ohlc'
    const overlays = chartCandlestickOverlayOptions(chart.options?.candlestick, mode)
    const style = chartCandlestickStyle(chart, source.field, index, chartCandlestickState(candle), palette, source.sourceName ?? source.name)
    if (candle.open !== undefined) {
      const open = y(candle.open), wick = chartCandlestickWick(candle, chart.options?.candlestick)
      const top = Math.min(open, close), bottom = Math.max(open, close), border = style.border!
      chartLine(commands, { x, y: y(wick.high) }, { x, y: y(Math.max(candle.open, candle.close)) }, border, style.color, plot)
      chartLine(commands, { x, y: y(Math.min(candle.open, candle.close)) }, { x, y: y(wick.low) }, border, style.color, plot)
      if (candle.open === candle.close) chartLine(commands, { x: x - width / 2, y: close }, { x: x + width / 2, y: close }, border, style.color, plot)
      else chartRect(commands, { x: x - width / 2, y: top, width, height: bottom - top },
        style.hollow ? 'transparent' : chartColorOpacity(style.color!, style.fillOpacity), border, plot,
        style.hollow ? undefined : chartFillPatternPaint(style.fillPattern, style.fillOpacity))
    }
    if (overlays.highLowLine.visible) chartLine(commands, { x, y: y(candle.high) }, { x, y: y(candle.low) },
      overlays.highLowLine.border, overlays.highLowLine.border.color, plot)
    const marker = overlays.closeMarker
    if (marker.visible && marker.point.shape !== 'none' && marker.point.size > 0) {
      if (marker.point.shape === 'dot' || marker.point.shape === 'dash') {
        const half = marker.point.size * (marker.point.shape === 'dot' ? 0.5 : 1) / 2
        chartLine(commands, { x: x - half, y: close }, { x: x + half, y: close },
          { color: marker.point.color, width: marker.lineWidth }, marker.point.color, plot)
      } else chartSymbol(commands, { x, y: close }, marker.point.size, marker.point.color, marker.point.shape,
        { color: marker.point.color, width: marker.lineWidth }, plot)
    }
    const label = style.label
    if (label?.visible) {
      const side = label.position === 'left' || label.position === 'right', inside = label.position === 'inside' || label.position === 'center'
      const top = candle.open === undefined ? close : Math.min(y(candle.open), close), bottom = candle.open === undefined ? close : Math.max(y(candle.open), close)
      commands.push({ kind: 'text', x: x + (label.position === 'left' ? -width / 2 - 9 : label.position === 'right' ? width / 2 + 9 : 0),
        y: inside ? (top + bottom) / 2 : side ? close : label.position === 'bottom' ? bottom + 10 : top - 9,
        text: chartLabelText(label, candle.close, data.categories[index] ?? '', style.name ?? source.name),
        font: chartFont({ ...definedChartStyleFields(chart.options?.appearance?.font), ...definedChartStyleFields(label) }), color: label.color ?? textColor,
        align: label.position === 'left' ? 'right' : label.position === 'right' ? 'left' : 'center', baseline: 'middle', clip: plot })
    }
  }
}
