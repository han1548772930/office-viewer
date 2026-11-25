import type { ChartAreaLineStyle, ChartInvalidValueStrategy, ChartModel, ChartSemanticOptions, ChartType } from '../../domain/model'
import { resolveChartType } from './chart-catalog'

/** Published @univerjs-pro/engine-chart@1.6.0 AreaLineStyle values. */
export const ChartAreaLineStyles = Object.freeze({ Line: 'line', Smooth: 'smooth', Step: 'step' } as const)
export const ChartInvalidValueTypes = Object.freeze({ Zero: 'zero', Break: 'break', Link: 'link' } as const)
export const CHART_AREA_LINE_STYLE_DEFAULT: ChartAreaLineStyle = 'line'
export const CHART_INVALID_VALUE_STRATEGY_DEFAULT: ChartInvalidValueStrategy = 'break'

export function supportsChartAreaLineStyle(type: ChartType): boolean {
  return ['line', 'area', 'combination'].includes(resolveChartType(type).family)
}

/** Scatter shares the empty-value section, but has no area line-shape control. */
export function supportsChartLineAndAreaSection(type: ChartType): boolean {
  return supportsChartAreaLineStyle(type) || resolveChartType(type).family === 'scatter'
}

export function chartAreaLineStyleValue(value: unknown): ChartAreaLineStyle {
  if (value !== 'line' && value !== 'smooth' && value !== 'step') throw new TypeError('图表线型须为直线、平滑或阶梯。')
  return value
}

export function chartInvalidValueStrategyValue(value: unknown): ChartInvalidValueStrategy {
  if (value !== 'zero' && value !== 'break' && value !== 'link') throw new TypeError('空值处理须为断开、连接或按零值。')
  return value
}

export function assertChartAreaLineStyleType(type: ChartType): void {
  if (!supportsChartAreaLineStyle(type)) throw new TypeError('该图型不支持折线或面积线型设置。')
}

/** Legacy smooth never changes an existing line series. New lineStyle is shared. */
export function resolveChartAreaLineStyle(chart: Pick<ChartModel, 'type' | 'options'>, seriesType?: 'line' | 'area'): ChartAreaLineStyle {
  if (!supportsChartAreaLineStyle(chart.type)) return CHART_AREA_LINE_STYLE_DEFAULT
  const area = chart.options?.area
  if (area?.lineStyle !== undefined) return chartAreaLineStyleValue(area.lineStyle)
  const family = seriesType ?? (resolveChartType(chart.type).family === 'area' ? 'area' : 'line')
  return family === 'area' && area?.smooth === true ? 'smooth' : CHART_AREA_LINE_STYLE_DEFAULT
}

/** Explicit clears must not resurrect the legacy area switch after merge/reopen. */
export function chartAreaLineStylePatch(type: ChartType, value: ChartAreaLineStyle | undefined): ChartSemanticOptions {
  assertChartAreaLineStyleType(type)
  return { area: { lineStyle: value === undefined ? undefined : chartAreaLineStyleValue(value), smooth: undefined } }
}
