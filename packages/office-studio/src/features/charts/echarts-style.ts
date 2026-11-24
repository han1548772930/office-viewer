import type { ChartAxisOptions, ChartFontOptions, ChartLabelOptions, ChartSeriesOptions } from '../../domain/model'
import { chartColorOpacity, chartFilledMarkBorder, chartLabelText, formatChartNumber } from './chart-geometry'
import { definedChartStyleFields } from './chart-options'
import { chartEChartsAutomaticGradient, type ChartAutomaticGradient } from './chart-auto-gradient'
import { chartColorRepresentative } from './chart-color'

export function chartEChartsFont(font: ChartFontOptions | undefined, fallback: string, size?: number) {
  return { color: font?.color ?? fallback, fontSize: font?.fontSize ?? size, fontWeight: font?.bold ? 'bold' as const : 'normal' as const,
    fontStyle: font?.italic ? 'italic' as const : 'normal' as const, fontFamily: font?.family }
}

export function chartEChartsItemStyle(style: ChartSeriesOptions, gradient?: ChartAutomaticGradient) {
  return { color: chartEChartsAutomaticGradient(gradient) ?? (style.color === undefined ? undefined : chartColorRepresentative(style.color)), opacity: gradient ? 1 : style.fillOpacity ?? 1, borderColor: style.border?.color === undefined ? undefined : chartColorOpacity(style.border.color, style.border.opacity ?? 1),
    borderWidth: style.border?.visible === false ? 0 : style.border?.width, borderType: style.border?.dashType, borderJoin: 'round' as const }
}

/** Native item opacity also fades its outline. Pareto business bars author
 * fill and border alpha independently, so embed only the fill multiplier. */
export function chartEChartsIndependentItemStyle(style: ChartSeriesOptions, gradient?: ChartAutomaticGradient) {
  const item = chartEChartsItemStyle(style.color === undefined ? style : { ...style, border: chartFilledMarkBorder(style.color, style.border) }, gradient)
  return gradient || style.fillOpacity === undefined || style.fillOpacity === 1 ? item
    : { ...item, color: style.color === undefined ? undefined : chartColorOpacity(style.color, style.fillOpacity), opacity: 1 }
}

export function chartEChartsMarkerStyle(style: ChartSeriesOptions, gradient?: ChartAutomaticGradient) {
  return chartEChartsItemStyle({ ...style, color: style.point?.color ?? style.color }, gradient?.authored ? undefined : gradient)
}

export function chartEChartsLineStyle(style: ChartSeriesOptions, gradient?: ChartAutomaticGradient) {
  return { color: chartEChartsAutomaticGradient(gradient) ?? style.border?.color ?? (style.color === undefined ? undefined : chartColorRepresentative(style.color)), width: style.border?.visible === false ? 0 : style.border?.width ?? 2,
    opacity: gradient ? 1 : style.border?.opacity ?? 1, type: style.border?.dashType ?? 'solid' as const }
}

export function chartEChartsLabel(label: ChartLabelOptions | undefined, color: string) {
  const position = label?.position === 'auto' ? 'top' : label?.position === 'insideBase' || label?.position === 'insideEnd' || label?.position === 'center' ? 'inside' : label?.position ?? 'top'
  return { show: label?.visible ?? false, position, ...chartEChartsFont(label, color, 10) }
}

export function chartEChartsPointLabel(label: ChartLabelOptions | undefined, color: string, value: number, category: string, series: string, percentage?: number, bubbleSize?: number) {
  return { ...chartEChartsLabel(label, color), formatter: chartLabelText(label, value, category, series, percentage, bubbleSize) }
}

export function chartEChartsAxisStyle(axis: ChartAxisOptions, color: string, gridColor: string, numeric: boolean, inheritedFont?: ChartFontOptions) {
  const titleFont = { ...definedChartStyleFields(inheritedFont), ...definedChartStyleFields(axis.titleStyle) }
  return {
    inverse: axis.reverse,
    // Titles use the shared physical-box graphic geometry, independently of
    // ECharts nameLocation's value-axis direction and reverse behavior.
    name: undefined, nameLocation: axis.titleAlignment === 'start' ? 'start' as const : axis.titleAlignment === 'end' ? 'end' as const : 'middle' as const,
    nameTextStyle: chartEChartsFont(titleFont, color, 11), nameGap: Math.max(28, (titleFont.fontSize ?? 11) + 17),
    axisLine: { show: axis.line?.visible ?? true, lineStyle: { color: axis.line?.color ?? color, width: axis.line?.width ?? 1, opacity: axis.line?.opacity ?? 1, type: axis.line?.dashType ?? 'solid' as const } },
    axisTick: { show: axis.tick?.visible ?? false, inside: axis.tick?.position === 'inside', length: axis.tick?.length ?? 4,
      lineStyle: { color: axis.tick?.color ?? color, width: axis.tick?.width ?? 1, type: axis.tick?.dashType ?? 'solid' as const } },
    axisLabel: { show: axis.label?.visible ?? true, rotate: axis.label?.rotate ?? 0, ...chartEChartsFont({ ...definedChartStyleFields(inheritedFont), ...definedChartStyleFields(axis.label) }, color, 10),
      ...(numeric && axis.format ? { formatter: (value: string | number) => formatChartNumber(Number(value), axis.format) } : {}) },
    splitLine: { show: axis.gridLine?.visible ?? numeric, lineStyle: { color: axis.gridLine?.color ?? gridColor, width: axis.gridLine?.width ?? 1, opacity: axis.gridLine?.opacity ?? 1, type: axis.gridLine?.dashType ?? 'solid' as const } },
  }
}
