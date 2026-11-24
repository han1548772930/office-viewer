import type { ChartModel, ChartNullablePatch, ChartSemanticOptions, ChartSeriesOptions, ChartSeriesSelector, ChartType, ChartWaterfallConnectorOptions, ChartWaterfallSeriesStyle, ChartWaterfallStyles, ChartWaterfallStyleTarget } from '../../domain/model'
import { chartNullableValues, patchChartSeriesOverride } from './chart-cartesian-api'
import { chartNamedSeriesStyle, chartSeriesOverride, validateChartSeriesSelector } from './chart-series-selector'
import { chartSeriesColorIsAuthored, chartSeriesStyle, cloneChartOptions, definedChartStyleFields, mergeChartOptions } from './chart-options'

export const ChartWaterfallStyleTargets = Object.freeze({ Positive: 'positive', Negative: 'negative', Subtotal: 'subtotal' } as const)
export const CHART_WATERFALL_ROLE_NAMES: Readonly<Record<ChartWaterfallStyleTarget, string>> = Object.freeze({ positive: '增加', negative: '减少', subtotal: '合计' })

export function assertChartWaterfallType(type: ChartType): void {
  if (type !== 'waterfall') throw new TypeError('此配置仅适用于瀑布图。')
}

export function assertChartWaterfallStyleTarget(target: unknown): asserts target is ChartWaterfallStyleTarget {
  if (target !== 'positive' && target !== 'negative' && target !== 'subtotal') throw new TypeError('无效的瀑布图样式目标。')
}

/** Role patches cannot recursively create a new waterfall style tree. Null
 * resets are the same local extension used by the Cartesian style API. */
export function chartWaterfallStyleValues(value: ChartNullablePatch<ChartWaterfallSeriesStyle> | undefined): ChartWaterfallSeriesStyle | undefined {
  if (value === undefined) return undefined
  const style = chartNullableValues<ChartWaterfallSeriesStyle>(value)
  if (Object.hasOwn(style, 'waterfallStyles')) throw new TypeError('瀑布图角色样式不能嵌套 waterfallStyles。')
  return cloneChartOptions({ allSeriesStyle: style })!.allSeriesStyle
}

/** Source selectors and all other role targets retain their authored state. */
export function patchChartWaterfallStyle(current: ChartSemanticOptions | undefined, selector: ChartSeriesSelector,
  target: ChartWaterfallStyleTarget, value: ChartNullablePatch<ChartWaterfallSeriesStyle> | undefined): ChartSemanticOptions {
  validateChartSeriesSelector(selector)
  assertChartWaterfallStyleTarget(target)
  const role = chartWaterfallStyleValues(value), before = chartSeriesOverride(current, selector)?.waterfallStyles?.[target]
  const next = role === undefined ? undefined
    : mergeChartOptions({ allSeriesStyle: before }, { allSeriesStyle: role })!.allSeriesStyle
  return patchChartSeriesOverride(current, selector, { waterfallStyles: { [target]: next } })
}

export function chartWaterfallConnectorValues(value: boolean | ChartNullablePatch<ChartWaterfallConnectorOptions> | undefined): boolean | ChartWaterfallConnectorOptions | undefined {
  if (value === undefined || typeof value === 'boolean') return value
  const connector = chartNullableValues<ChartWaterfallConnectorOptions>(value)
  return cloneChartOptions({ waterfall: { connector } })!.waterfall!.connector
}

/** Runtime defaults: omission/false hides it; true and an object enable
 * a neutral 1px dashed connector with full opacity. */
export function chartWaterfallConnector(chart: Pick<ChartModel, 'options'>, themeColor = '#8C8C8C'): Required<ChartWaterfallConnectorOptions> | undefined {
  const authored = chart.options?.waterfall?.connector
  if (authored === undefined || authored === false) return undefined
  return { color: themeColor, width: 1, opacity: 1, dashType: 'dashed', ...(authored === true ? {} : definedChartStyleFields(authored)) }
}

export function chartWaterfallRole(role: 'change' | 'total', value: number): ChartWaterfallStyleTarget {
  return role === 'total' ? 'subtotal' : value < 0 ? 'negative' : 'positive'
}

export function chartWaterfallUsesSecondaryAxis(chart: Pick<ChartModel, 'type' | 'options'>,
  sources?: readonly { field: number; name: string; sourceName?: string }[]): boolean {
  if (chart.type !== 'waterfall') return false
  const roles = Object.values(ChartWaterfallStyleTargets)
  if (chart.options?.waterfall?.stackType === 'stacked' && chart.options.waterfall.useSubtotal !== false
    && chartWaterfallRoleStyle(chart, undefined, 0, 'subtotal', []).rightAxis) return true
  if (sources) return sources.some((source, position) => roles.some(target => chartWaterfallRoleStyle(chart,
    source.field, position, target, [], source.sourceName ?? source.name).rightAxis))
  const fields = [...new Set([0, ...Object.keys(chart.options?.series ?? {}).map(Number)])]
  if (fields.some(field => roles.some(target => chartWaterfallRoleStyle(chart, field, 0, target, []).rightAxis))) return true
  return Boolean(chart.options?.namedSeries?.some(entry => roles.some(target => chartWaterfallRoleStyle(chart, 0, 0, target, [], entry.name).rightAxis)))
}

/** Business-series styles form the base. Global role, shared role, source-field
 * role and exact-name role layers follow it; individual role dataPoints are
 * merged before the materialized point is resolved. Old snapshot colors stay
 * as fallback values, never replace newly authored role colors. */
export function chartWaterfallRoleStyle(chart: Pick<ChartModel, 'options'>, field: number | undefined, position: number,
  target: ChartWaterfallStyleTarget, palette: readonly string[], sourceName?: string): ChartWaterfallSeriesStyle {
  assertChartWaterfallStyleTarget(target)
  const options = chart.options, base = field === undefined ? { options: { allSeriesStyle: options?.allSeriesStyle } } : chart
  const common = chartSeriesStyle(base, field ?? 0, position, palette, sourceName), prefix = common.name ?? sourceName
  const legacy = target === 'positive' ? options?.waterfall?.positiveColor ?? palette[position * 2 % Math.max(1, palette.length)] ?? '#059669'
    : target === 'negative' ? options?.waterfall?.negativeColor ?? palette[(position * 2 + 1) % Math.max(1, palette.length)] ?? '#dc2626'
      : options?.waterfall?.totalColor ?? '#CDD0D8'
  let style: ChartSeriesOptions = { ...common, name: prefix ? `${prefix}-(${CHART_WATERFALL_ROLE_NAMES[target]})` : CHART_WATERFALL_ROLE_NAMES[target],
    color: chartSeriesColorIsAuthored(base, field ?? 0, sourceName) ? common.color : legacy }
  const layers: Array<ChartWaterfallStyles | undefined> = [options?.waterfall?.styles,
    options?.allSeriesStyle?.waterfallStyles, field === undefined ? undefined : options?.series?.[field]?.waterfallStyles,
    field === undefined ? undefined : chartNamedSeriesStyle(options, sourceName)?.waterfallStyles]
  for (const group of layers) if (group?.[target]) style = chartSeriesStyle({ options: { allSeriesStyle: style, series: { [field ?? 0]: group[target] } } }, field ?? 0, position, palette)
  const result = { ...style }
  delete result.waterfallStyles
  return result
}

/** Resolve each layer once per series, including role data-point maps. */
export function chartWaterfallRoleStyles(chart: Pick<ChartModel, 'options'>, field: number | undefined, position: number,
  palette: readonly string[], sourceName?: string): Readonly<Record<ChartWaterfallStyleTarget, ChartWaterfallSeriesStyle>> {
  return { positive: chartWaterfallRoleStyle(chart, field, position, 'positive', palette, sourceName),
    negative: chartWaterfallRoleStyle(chart, field, position, 'negative', palette, sourceName),
    subtotal: chartWaterfallRoleStyle(chart, field, position, 'subtotal', palette, sourceName) }
}

/** Keeps descriptions detached without making renderer defaults persistent. */
export function chartWaterfallAuthoredStyles(chart: Pick<ChartModel, 'options'>, field: number, sourceName?: string): ChartWaterfallStyles | undefined {
  const groups = [chart.options?.waterfall?.styles, chart.options?.allSeriesStyle?.waterfallStyles,
    chart.options?.series?.[field]?.waterfallStyles, chartNamedSeriesStyle(chart.options, sourceName)?.waterfallStyles]
  if (!groups.some(group => group !== undefined)) return undefined
  const roles: ChartWaterfallStyles = {}
  for (const target of Object.values(ChartWaterfallStyleTargets)) for (const group of groups) if (group?.[target]) {
    roles[target] = mergeChartOptions({ allSeriesStyle: roles[target] }, { allSeriesStyle: group[target] })!.allSeriesStyle
  }
  return cloneChartOptions({ waterfall: { styles: roles } })!.waterfall!.styles
}
