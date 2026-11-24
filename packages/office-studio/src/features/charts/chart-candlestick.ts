import type { ChartCandlestickMappingOptions, ChartCandlestickOptions, ChartLineOptions, ChartModel, ChartPointShape, ChartSeriesOptions } from '../../domain/model'
import type { ChartCandle } from './chart-data'
import type { ChartConfig, ChartData } from './chart-engine'
import { chartSeriesColorIsAuthored, chartSeriesStyle, definedChartStyleFields } from './chart-options'
import { chartColorRepresentative, isChartLinearGradient } from './chart-color'

export type ChartCandlestickState = 'rising' | 'falling' | 'doji'
export type ChartCandlestickMode = 'hlc' | 'ohlc'
export const CHART_CANDLESTICK_STATE_CHOICES: ReadonlyArray<readonly [ChartCandlestickState, string]> = [
  ['rising', '上涨'], ['falling', '下跌'], ['doji', '十字线'],
]
/** Fixed v1.6.0 default-theme tokens: green.500 / red.500 / gray.500. */
export const CHART_CANDLESTICK_STATE_COLORS: Readonly<Record<ChartCandlestickState, string>> = {
  rising: '#0DA471', falling: '#F05252', doji: '#4E5565',
}
export const CHART_CANDLESTICK_MIN_AUTO_WIDTH = 8
export const CHART_CANDLESTICK_MAX_AUTO_WIDTH = 32
export const CHART_CANDLESTICK_CLOSE_MARKER_SIZE = 8
export const CHART_CANDLESTICK_CLOSE_MARKER_LINE_WIDTH = 1
export interface ChartCandlestickResolvedStyle extends ChartSeriesOptions { color?: string; hollow: boolean }
export interface ChartCandlestickOverlayOptions {
  highLowLine: { visible: boolean; border: ChartLineOptions }
  closeMarker: { visible: boolean; point: { color: string; size: number; shape: ChartPointShape }; lineWidth: number }
}

export function chartCandlestickState(candle: Pick<ChartCandle, 'open' | 'close'>): ChartCandlestickState {
  return candle.open === undefined || candle.close === candle.open ? 'doji' : candle.close > candle.open ? 'rising' : 'falling'
}

/** An explicit HLC mapping remains HLC even when older valueFields contains
 * four fields. This helper only reads draft configuration. */
export function chartCandlestickSettingsMapping(config: Pick<ChartConfig, 'options' | 'categoryField' | 'valueFields'>,
  fieldCount = 4, sourceValueFields?: readonly number[]): ChartCandlestickMappingOptions {
  if (config.options?.mapping?.candlestick) return { ...config.options.mapping.candlestick }
  const defaults = sourceValueFields ?? Array.from({ length: fieldCount }, (_unused, index) => index)
  const fields = config.valueFields?.length ? config.valueFields : defaults.filter(index => index !== (config.categoryField ?? 0))
  return fields.length >= 4 ? { openIndex: fields[0], highIndex: fields[1]!, lowIndex: fields[2]!, closeIndex: fields[3]! }
    : { highIndex: fields[0]!, lowIndex: fields[1]!, closeIndex: fields[2]! }
}

export function chartCandlestickMode(config: Pick<ChartConfig, 'options' | 'categoryField' | 'valueFields'>, data?: ChartData): ChartCandlestickMode {
  const mapping = config.options?.mapping?.candlestick ?? data?.candleFields
    ?? chartCandlestickSettingsMapping(config, data?.fields?.length ?? 4, data?.sourceValueFields)
  return mapping.openIndex === undefined ? 'hlc' : 'ohlc'
}

/** Shared nominal width; automatic candles use half a category band bounded
 * by the release's 8..32px policy. Authored positive widths remain absolute. */
export function chartCandlestickWidth(bandwidth: number, options?: ChartCandlestickOptions): number {
  if (options?.candleWidth !== undefined) return options.candleWidth
  return Math.max(CHART_CANDLESTICK_MIN_AUTO_WIDTH, Math.min(CHART_CANDLESTICK_MAX_AUTO_WIDTH, bandwidth / 2))
}

/** State groups override legacy snapshot fields and generic series defaults.
 * Existing per-point overrides are resolved last in the same source indices. */
export function chartCandlestickStyle(chart: Pick<ChartModel, 'options'>, field: number, index: number | undefined, state: ChartCandlestickState,
  palette: readonly string[] = [], sourceName?: string): ChartCandlestickResolvedStyle {
  const base = chartSeriesStyle(chart, field, 0, palette.length ? palette : undefined, sourceName), options = chart.options?.candlestick
  const authored = options?.[state], legacy = state === 'rising' ? options?.risingColor : state === 'falling' ? options?.fallingColor : undefined
  const genericColor = chartSeriesColorIsAuthored(chart, field, sourceName), defaultColor = CHART_CANDLESTICK_STATE_COLORS[state]
  const inheritedColor = genericColor ? chartColorRepresentative(base.color!) : legacy ?? defaultColor
  const rawPoint = index === undefined ? undefined : base.dataPoints?.[index]
  const point = rawPoint && { ...rawPoint, color: isChartLinearGradient(rawPoint.color) ? undefined : rawPoint.color }
  const color = authored?.color ?? inheritedColor
  return { ...base, color, fillOpacity: authored?.fillOpacity ?? base.fillOpacity ?? 1,
    hollow: authored?.hollow ?? (state === 'rising' ? options?.hollow : undefined) ?? false,
    ...definedChartStyleFields(point),
    border: { ...base.border, color: authored?.border?.color ?? base.border?.color ?? chartColorRepresentative(inheritedColor),
      width: authored?.border?.width ?? base.border?.width ?? 1, opacity: authored?.border?.opacity ?? base.border?.opacity ?? 1,
      ...definedChartStyleFields(point?.border) },
    point: { ...base.point, ...(point?.color === undefined ? {} : { color: chartColorRepresentative(point.color) }),
      ...(point?.size === undefined ? {} : { size: point.size }), ...(point?.shape === undefined ? {} : { shape: point.shape }) } }
}

export function chartCandlestickOverlayOptions(options: ChartCandlestickOptions | undefined, mode: ChartCandlestickMode): ChartCandlestickOverlayOptions {
  const color = CHART_CANDLESTICK_STATE_COLORS.doji
  return {
    highLowLine: { visible: options?.highLowLine?.visible ?? (mode === 'hlc' || options?.highLowLine !== undefined),
      border: { color, width: 1, opacity: 1, dashType: 'solid', ...definedChartStyleFields(options?.highLowLine?.border) } },
    closeMarker: { visible: options?.closeMarker?.visible ?? (mode === 'hlc' || options?.closeMarker !== undefined),
      lineWidth: CHART_CANDLESTICK_CLOSE_MARKER_LINE_WIDTH,
      point: { color, size: CHART_CANDLESTICK_CLOSE_MARKER_SIZE, shape: 'dot', ...definedChartStyleFields(options?.closeMarker?.point) } },
  }
}

/** Presence of the high/low overlay changes the ordinary OHLC wick even when
 * that overlay is explicitly hidden. Clearing the object restores full wicks. */
export function chartCandlestickWick(candle: ChartCandle, options?: ChartCandlestickOptions): { low: number; high: number } {
  return candle.open !== undefined && options?.highLowLine !== undefined
    ? { low: Math.min(candle.open, candle.close), high: Math.max(candle.open, candle.close) } : { low: candle.low, high: candle.high }
}
