import type { ChartCandlestickCloseMarkerOptions, ChartCandlestickHighLowLineOptions, ChartCandlestickMappingOptions,
  ChartCandlestickStateStyleOptions, ChartNullablePatch, ChartType } from '../../domain/model'
import { chartNullableValues } from './chart-cartesian-api'
import { cloneChartOptions } from './chart-options'

export function assertChartCandlestickType(type: ChartType): void {
  if (type !== 'candlestick') throw new TypeError('K 线专属设置仅适用于 K 线图。')
}

function assertFields(value: unknown, fields: readonly string[], label: string): void {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype
    || Object.getOwnPropertySymbols(value).length || Object.keys(value).some(key => !fields.includes(key))) throw new TypeError(label + '包含无效字段。')
}

export function chartCandlestickFields(value: ChartCandlestickMappingOptions): ChartCandlestickMappingOptions {
  assertFields(value, ['openIndex', 'highIndex', 'lowIndex', 'closeIndex'], 'K 线字段映射')
  for (const key of ['openIndex', 'highIndex', 'lowIndex', 'closeIndex'] as const) {
    const index = value[key]
    if (key === 'openIndex' && index === undefined) continue
    if (!Number.isSafeInteger(index) || index === undefined || index < 0) throw new RangeError('K 线字段索引需要非负安全整数。')
  }
  return { ...value }
}

export function chartCandlestickStateStyleValues(value: ChartNullablePatch<ChartCandlestickStateStyleOptions>): ChartCandlestickStateStyleOptions {
  assertFields(value, ['color', 'fillOpacity', 'border', 'hollow'], 'K 线状态样式')
  if (value.border) assertFields(value.border, ['color', 'width', 'opacity'], 'K 线状态边框')
  const style = chartNullableValues<ChartCandlestickStateStyleOptions>(value)
  return cloneChartOptions({ candlestick: { rising: style } })!.candlestick!.rising!
}

export function chartCandlestickHighLowLineValues(value: ChartNullablePatch<ChartCandlestickHighLowLineOptions>): ChartCandlestickHighLowLineOptions {
  assertFields(value, ['visible', 'border'], 'K 线高低线样式')
  const highLowLine = chartNullableValues<ChartCandlestickHighLowLineOptions>(value)
  return cloneChartOptions({ candlestick: { highLowLine } })!.candlestick!.highLowLine!
}

export function chartCandlestickCloseMarkerValues(value: ChartNullablePatch<ChartCandlestickCloseMarkerOptions>): ChartCandlestickCloseMarkerOptions {
  assertFields(value, ['visible', 'point'], 'K 线收盘标记样式')
  const closeMarker = chartNullableValues<ChartCandlestickCloseMarkerOptions>(value)
  return cloneChartOptions({ candlestick: { closeMarker } })!.candlestick!.closeMarker!
}

export function chartCandleWidth(value: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) throw new RangeError('K 线宽度需要有限正数。')
  return value
}
