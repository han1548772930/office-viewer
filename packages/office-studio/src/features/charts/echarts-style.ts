import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartSeriesOptions } from '../../domain/model'
import { chartLabelText, formatChartNumber } from './chart-geometry'

export function chartEChartsFont(font: ChartFontOptions | undefined, fallback: string, size?: number) {
  return { color: font?.color ?? fallback, fontSize: font?.fontSize ?? size, fontWeight: font?.bold ? 'bold' as const : 'normal' as const,
    fontStyle: font?.italic ? 'italic' as const : 'normal' as const, fontFamily: font?.family }
}

export function chartEChartsItemStyle(style: ChartSeriesOptions) {
  return { color: style.color, opacity: style.fillOpacity ?? 1, borderColor: style.border?.color,
    borderWidth: style.border?.visible === false ? 0 : style.border?.width, borderType: style.border?.dashType, borderJoin: 'round' as const }
}

export function chartEChartsLineStyle(style: ChartSeriesOptions) {
  return { color: style.border?.color ?? style.color, width: style.border?.visible === false ? 0 : style.border?.width ?? 2,
    opacity: style.border?.opacity ?? 1, type: style.border?.dashType ?? 'solid' as const }
}

export function chartEChartsLabel(label: ChartLabelOptions | undefined, color: string) {
  return { show: label?.visible ?? false, position: label?.position ?? 'top', ...chartEChartsFont(label, color, 10) }
}

export function chartEChartsPointLabel(label: ChartLabelOptions | undefined, color: string, value: number, category: string, series: string, percentage?: number) {
  return { ...chartEChartsLabel(label, color), formatter: chartLabelText(label, value, category, series, percentage) }
}

export function chartEChartsAxisStyle(axis: ChartAxisOptions, color: string, gridColor: string, numeric: boolean) {
  return {
    inverse: axis.reverse,
    name: axis.title, nameLocation: 'middle' as const,
    axisLine: { show: axis.line?.visible ?? true, lineStyle: { color: axis.line?.color ?? color, width: axis.line?.width ?? 1, opacity: axis.line?.opacity ?? 1, type: axis.line?.dashType ?? 'solid' as const } },
    axisTick: { show: axis.tick?.visible ?? false, inside: axis.tick?.position === 'inside', length: axis.tick?.length ?? 4,
      lineStyle: { color: axis.tick?.color ?? color, width: axis.tick?.width ?? 1, type: axis.tick?.dashType ?? 'solid' as const } },
    axisLabel: { show: axis.label?.visible ?? true, rotate: axis.label?.rotate ?? 0, ...chartEChartsFont(axis.label, color, 10),
      ...(numeric && axis.format ? { formatter: (value: string | number) => formatChartNumber(Number(value), axis.format) } : {}) },
    splitLine: { show: axis.gridLine?.visible ?? numeric, lineStyle: { color: axis.gridLine?.color ?? gridColor, width: axis.gridLine?.width ?? 1, opacity: axis.gridLine?.opacity ?? 1, type: axis.gridLine?.dashType ?? 'solid' as const } },
  }
}
