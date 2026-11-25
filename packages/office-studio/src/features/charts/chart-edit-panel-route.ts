import type { ChartSeriesSelector, ChartWaterfallStyleTarget } from '../../domain/model'
import type { ChartConfig, ChartData } from './chart-engine'
import { resolveChartType, supportsChartStatistics } from './chart-catalog'
import { chartHasLegend } from './chart-legend'
import { chartSeriesStyle } from './chart-options'
import { supportsChartAxisPointer } from './chart-axis-pointer'
import { chartWaterfallUsesSecondaryAxis } from './chart-waterfall'
import { supportsChartAreaLineStyle, supportsChartLineAndAreaSection } from './chart-line-and-area'

/** Public v1.6.0 route values; UI navigation is separate from authored content. */
export const CHART_EDIT_PANEL_TABS = ['style', 'data'] as const
export const CHART_EDIT_PANEL_SECTIONS = [
  'chartStyle', 'chartAndAxisTitles', 'legend', 'horizontalAxis', 'verticalAxis', 'rightVerticalAxis', 'series',
  'pieStyle', 'paretoBarSeries', 'paretoLineSeries', 'waterfallSeries', 'relation', 'wordCloud', 'radar', 'funnel',
  'heatmap', 'lineAndArea', 'gridlinesAndTicks', 'indicatorLine', 'trendline', 'errorBars', 'candlestick',
  'histogram', 'treemap', 'sunburst', 'gauge', 'chord',
] as const
export const CHART_EDIT_PANEL_AXES = ['xAxis', 'yAxis', 'rightYAxis'] as const
export const CHART_EDIT_PANEL_TITLES = ['title', 'subtitle', 'xAxisTitle', 'yAxisTitle', 'rightYAxisTitle'] as const
export type ChartEditPanelTab = typeof CHART_EDIT_PANEL_TABS[number]
export type ChartEditPanelSection = typeof CHART_EDIT_PANEL_SECTIONS[number]
export type ChartEditPanelAxis = typeof CHART_EDIT_PANEL_AXES[number]
export type ChartEditPanelTitle = typeof CHART_EDIT_PANEL_TITLES[number]
/** Local, transient focus hints refine an existing public section. They are
 * validated control DTOs, never author fields or arbitrary DOM selectors. */
export type ChartEditPanelFocus =
  | { kind: 'waterfallRole'; role: ChartWaterfallStyleTarget }
  | { kind: 'pieAggregate' | 'pieConnector' | 'pieLabel' }
  | { kind: 'pieSlice'; sourcePointIndex: number }
export interface ChartEditPanelRoute {
  tab: ChartEditPanelTab
  section?: ChartEditPanelSection
  axisId?: ChartEditPanelAxis
  seriesId?: string
  titleKey?: ChartEditPanelTitle
  focus?: ChartEditPanelFocus
}
export interface ChartEditPanelNavigation {
  route: ChartEditPanelRoute
  panels: string[]
  fonts: string[]
  axes: string[]
  series: string[]
  optionKey?: string
  seriesTarget?: string
  paretoGroup?: 'barStyle' | 'cumulativeLineStyle'
  statisticsKind?: 'trendlines' | 'errorBars'
  waterfallRole?: ChartWaterfallStyleTarget
  pieSourcePointIndex?: number
}
const axisSection: Record<ChartEditPanelAxis, ChartEditPanelSection> = { xAxis: 'horizontalAxis', yAxis: 'verticalAxis', rightYAxis: 'rightVerticalAxis' }
const axisKey: Record<ChartEditPanelAxis, string> = { xAxis: 'x', yAxis: 'y', rightYAxis: 'right' }
const titleAxis: Partial<Record<ChartEditPanelTitle, ChartEditPanelAxis>> = { xAxisTitle: 'xAxis', yAxisTitle: 'yAxis', rightYAxisTitle: 'rightYAxis' }
const routeKeys = new Set(['tab', 'section', 'axisId', 'seriesId', 'titleKey', 'focus'])
function member<T extends string>(value: unknown, values: readonly T[], label: string): T {
  if (typeof value !== 'string' || !values.includes(value as T)) throw new TypeError(`${label} 无效。`)
  return value as T
}
function normalizeFocus(value: unknown): ChartEditPanelFocus | undefined {
  if (value === undefined) return undefined
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype
    || Object.getOwnPropertySymbols(value).length) throw new TypeError('图表字段定位无效。')
  const dto = value as Record<string, unknown>, keys = Object.keys(dto)
  if (dto.kind === 'waterfallRole' && keys.length === 2 && keys.includes('kind') && keys.includes('role'))
    return { kind: dto.kind, role: member(dto.role, ['positive', 'negative', 'subtotal'], '瀑布图角色定位') }
  if (dto.kind === 'pieSlice' && keys.length === 2 && keys.includes('kind') && keys.includes('sourcePointIndex')
    && typeof dto.sourcePointIndex === 'number' && Number.isSafeInteger(dto.sourcePointIndex) && dto.sourcePointIndex >= 0)
    return { kind: dto.kind, sourcePointIndex: dto.sourcePointIndex }
  if (keys.length === 1 && keys[0] === 'kind' && (dto.kind === 'pieAggregate' || dto.kind === 'pieConnector' || dto.kind === 'pieLabel')) return { kind: dto.kind }
  throw new TypeError('图表字段定位无效。')
}

/** Validate both the DTO and combinations which would otherwise be ignored. */
export function normalizeChartEditPanelRoute(input: unknown): ChartEditPanelRoute {
  if (!input || typeof input !== 'object' || Array.isArray(input) || Object.getPrototypeOf(input) !== Object.prototype
    || Object.getOwnPropertySymbols(input).length || Object.keys(input).some(key => !routeKeys.has(key))) throw new TypeError('图表面板路由必须仅包含公开字段和受支持的字段定位。')
  const dto = input as Record<string, unknown>, tab = member(dto.tab, CHART_EDIT_PANEL_TABS, '图表面板标签页')
  const section = dto.section === undefined ? undefined : member(dto.section, CHART_EDIT_PANEL_SECTIONS, '图表面板分组')
  const axisId = dto.axisId === undefined ? undefined : member(dto.axisId, CHART_EDIT_PANEL_AXES, '图表面板坐标轴')
  const titleKey = dto.titleKey === undefined ? undefined : member(dto.titleKey, CHART_EDIT_PANEL_TITLES, '图表面板标题')
  const seriesId = dto.seriesId
  const focus = normalizeFocus(dto.focus)
  if (seriesId !== undefined && (typeof seriesId !== 'string' || !seriesId.trim() || seriesId.length > 16_384)) throw new TypeError('图表面板系列标识无效。')
  if (tab === 'data') {
    if (section || axisId || titleKey || seriesId !== undefined || focus) throw new TypeError('数据标签页不能指定样式分组、轴、系列、标题或字段定位。')
    return { tab }
  }
  const resolvedSection = section ?? (titleKey ? 'chartAndAxisTitles' : axisId ? axisSection[axisId] : seriesId ? 'series' : 'chartStyle')
  if (titleKey && resolvedSection !== 'chartAndAxisTitles') throw new TypeError('标题标识必须指向图表与坐标轴标题分组。')
  if (axisId && !['horizontalAxis', 'verticalAxis', 'rightVerticalAxis', 'chartAndAxisTitles', 'gridlinesAndTicks'].includes(resolvedSection)) throw new TypeError('坐标轴标识与图表面板分组不一致。')
  if (seriesId !== undefined && resolvedSection !== 'series' && resolvedSection !== 'waterfallSeries') throw new TypeError('系列标识必须指向系列或瀑布图系列分组。')
  if (focus && (focus.kind === 'waterfallRole' ? resolvedSection !== 'waterfallSeries' : resolvedSection !== 'pieStyle'))
    throw new TypeError('图表字段定位与编辑分组不一致。')
  const expectedAxis = resolvedSection === 'horizontalAxis' ? 'xAxis' : resolvedSection === 'verticalAxis' ? 'yAxis' : resolvedSection === 'rightVerticalAxis' ? 'rightYAxis' : undefined
  if (axisId && expectedAxis && axisId !== expectedAxis || axisId && titleKey && titleAxis[titleKey] !== axisId) throw new TypeError('坐标轴、标题与分组的目标不一致。')
  const normalizedAxis = axisId ?? expectedAxis ?? (titleKey ? titleAxis[titleKey] : undefined)
  return { tab, section: resolvedSection, ...(normalizedAxis ? { axisId: normalizedAxis } : {}),
    ...(titleKey ? { titleKey } : {}), ...(seriesId !== undefined ? { seriesId: seriesId as string } : {}), ...(focus ? { focus } : {}) }
}

/** Reuse the editor's field/name scopes; these IDs never use display labels or indexes. */
export function chartEditPanelSeriesId(selector: ChartSeriesSelector): string {
  if (typeof selector === 'number') {
    if (!Number.isSafeInteger(selector) || selector < 0 || selector > 1_048_575) throw new RangeError('图表系列字段无效。')
    return `field:${selector}`
  }
  if (!selector || typeof selector.name !== 'string') throw new TypeError('图表系列名称选择器无效。')
  return `name:${JSON.stringify(selector.name)}`
}
export function resolveChartEditPanelSeries(seriesId: string, data: Pick<ChartData, 'series'> | undefined): ChartSeriesSelector {
  if (!data) throw new Error('图表系列数据尚未就绪，请完成当前图表数据准备后再定位系列。')
  if (/^field:(0|[1-9]\d*)$/.test(seriesId)) {
    const field = Number(seriesId.slice(6))
    if (Number.isSafeInteger(field) && field <= 1_048_575 && data.series.some(series => series.field === field)) return field
  } else if (seriesId.startsWith('name:')) {
    let name: unknown
    try { name = JSON.parse(seriesId.slice(5)) } catch { throw new TypeError('图表系列名称标识无效。') }
    if (typeof name === 'string' && chartEditPanelSeriesId({ name }) === seriesId) {
      const matches = data.series.filter(series => (series.sourceName ?? series.name) === name)
      if (matches.length === 1) return { name }
      if (matches.length > 1) throw new Error('图表源系列名称重复，请使用明确的字段系列标识。')
    }
  }
  throw new Error('指定系列不存在或系列标识不属于当前图表。')
}

/** Capability predicates are shared with the settings view, not a second engine. */
export function chartEditPanelCapabilities(config: ChartConfig, data?: ChartData) {
  const descriptor = resolveChartType(config.type), family = descriptor.family
  const hasAxes = ['line', 'column', 'bar', 'area', 'scatter', 'bubble', 'combination', 'waterfall', 'pareto', 'heatmap', 'boxplot', 'candlestick', 'histogram', 'radar'].includes(family)
  const secondaryAxis = !descriptor.threeD && ['line', 'column', 'bar', 'area', 'scatter', 'bubble', 'combination', 'waterfall', 'boxplot'].includes(family)
  const rightAxis = family === 'pareto' || chartWaterfallUsesSecondaryAxis(config, data?.series) || secondaryAxis && Boolean(data
    ? data.series.some((series, index) => chartSeriesStyle(config, series.field, index, undefined, series.sourceName ?? series.name).rightAxis)
    : config.options?.allSeriesStyle?.rightAxis || Object.values(config.options?.series ?? {}).some(style => style.rightAxis) || config.options?.namedSeries?.some(entry => entry.style.rightAxis))
  const indicator = supportsChartAxisPointer(config), statistics = supportsChartStatistics(config)
  const sections = new Set<ChartEditPanelSection>(['chartStyle', 'chartAndAxisTitles', 'series'])
  if (chartHasLegend(config.type)) sections.add('legend')
  if (hasAxes) for (const section of ['horizontalAxis', 'verticalAxis', 'gridlinesAndTicks'] as const) sections.add(section)
  if (rightAxis) sections.add('rightVerticalAxis')
  if (indicator) sections.add('indicatorLine')
  if (statistics) { sections.add('trendline'); sections.add('errorBars') }
  const specific: Partial<Record<ChartEditPanelSection, boolean>> = {
    pieStyle: family === 'pie', paretoBarSeries: family === 'pareto', paretoLineSeries: family === 'pareto',
    waterfallSeries: family === 'waterfall', relation: family === 'relation', wordCloud: family === 'wordCloud',
    radar: family === 'radar', funnel: family === 'funnel', heatmap: family === 'heatmap',
    lineAndArea: supportsChartLineAndAreaSection(config.type), candlestick: family === 'candlestick',
    histogram: family === 'histogram', treemap: family === 'treemap', sunburst: family === 'sunburst', gauge: family === 'gauge', chord: family === 'chord',
  }
  for (const [section, available] of Object.entries(specific)) if (available) sections.add(section as ChartEditPanelSection)
  return { descriptor, hasAxes, secondaryAxis, rightAxis, indicator, statistics, sections }
}

/** The existing form remains authoritative; the plan opens and focuses its real controls. */
export function planChartEditPanelRoute(config: ChartConfig, input: unknown, data?: ChartData): ChartEditPanelNavigation {
  const route = normalizeChartEditPanelRoute(input)
  const result: ChartEditPanelNavigation = { route, panels: [], fonts: [], axes: [], series: [] }
  if (route.tab === 'data') { result.panels = ['data']; result.optionKey = 'source'; return result }
  const section = route.section!, capabilities = chartEditPanelCapabilities(config, data)
  if (!capabilities.sections.has(section)) throw new Error('当前图表类型没有指定的编辑分组。')
  const axis = route.axisId ?? (section === 'gridlinesAndTicks' ? 'yAxis' : undefined)
  if (axis && (!capabilities.hasAxes || axis === 'rightYAxis' && !capabilities.rightAxis)) throw new Error('指定坐标轴不属于当前图表。')
  if (route.seriesId !== undefined) result.seriesTarget = chartEditPanelSeriesId(resolveChartEditPanelSeries(route.seriesId, data))
  if (section === 'chartStyle') { result.panels = ['common', 'fonts']; result.fonts = ['text', 'border']; result.optionKey = 'appearance.background' }
  else if (section === 'chartAndAxisTitles') {
    if (axis) { result.panels = ['axes']; result.axes = [axisKey[axis]]; result.optionKey = `${axis}.title` }
    else { result.panels = ['common', 'fonts']; result.fonts = [route.titleKey === 'subtitle' ? 'subtitle' : 'title']; result.optionKey = route.titleKey === 'subtitle' ? 'subtitle' : 'title' }
  } else if (section === 'legend') { result.panels = ['common', 'fonts']; result.fonts = ['legend']; result.optionKey = 'legend.visible' }
  else if (axis) { result.panels = ['axes']; result.axes = [axisKey[axis]]; result.optionKey = `${axis}.${section === 'gridlinesAndTicks' ? 'gridLine' : 'title'}` }
  else if (section === 'series') { result.panels = ['series']; result.series = [route.seriesId ? 'one' : 'all'] }
  else if (section === 'lineAndArea') {
    const hasShape = supportsChartAreaLineStyle(config.type)
    result.panels = hasShape ? ['common', 'type', 'series'] : ['common']
    if (hasShape) result.series = ['all']
    result.optionKey = hasShape ? 'area.lineStyle' : 'invalidValueStrategy'
  }
  else if (section === 'indicatorLine') { result.panels = ['indicator']; result.optionKey = 'axisPointer' }
  else if (section === 'trendline' || section === 'errorBars') { result.panels = ['statistics']; result.statisticsKind = section === 'trendline' ? 'trendlines' : 'errorBars' }
  else {
    result.panels = ['type']
    if (section === 'paretoBarSeries' || section === 'paretoLineSeries') result.paretoGroup = section === 'paretoBarSeries' ? 'barStyle' : 'cumulativeLineStyle'
    const keys: Partial<Record<ChartEditPanelSection, string>> = { pieStyle: 'pie', radar: 'radar', funnel: 'funnel', heatmap: 'heatmap',
      relation: 'relation', wordCloud: 'wordCloud',
      candlestick: 'candlestick', histogram: 'histogram', treemap: 'treemap', sunburst: 'sunburst', gauge: 'gauge', chord: 'chord' }
    result.optionKey = keys[section]
  }
  const focus = route.focus
  if (focus?.kind === 'waterfallRole') {
    result.waterfallRole = focus.role
    const selector = route.seriesId === undefined ? undefined : resolveChartEditPanelSeries(route.seriesId, data)
    const prefix = selector === undefined ? 'waterfall.styles' : typeof selector === 'number' ? `series.${selector}.waterfallStyles`
      : `namedSeries.${JSON.stringify(selector.name)}.waterfallStyles`
    result.optionKey = `${prefix}.${focus.role}.color`
    if (config.options?.waterfall?.stackType === 'stacked' && focus.role === 'subtotal'
      && (route.seriesId !== undefined || config.options?.waterfall?.useSubtotal === false)) throw new Error('当前堆积瀑布图没有可定位的来源合计或自动合计。')
  } else if (focus) {
    if (capabilities.descriptor.threeD) throw new Error('当前三维图表没有可定位的独立饼图标记字段。')
    const compound = config.type === 'pieOfPie' || config.type === 'barOfPie'
    if ((focus.kind === 'pieAggregate' || focus.kind === 'pieConnector') && !compound) throw new Error('当前饼图没有主副图合计或连接线。')
    if (focus.kind === 'pieSlice') {
      if (!data || focus.sourcePointIndex >= data.categories.length || !data.series.length) throw new Error('饼图来源扇区不存在或尚未就绪。')
      result.pieSourcePointIndex = focus.sourcePointIndex
      result.optionKey = `pie.sliceStyles.${focus.sourcePointIndex}.color`
    } else result.optionKey = focus.kind === 'pieAggregate' ? 'pie.aggregateStyle.color'
      : focus.kind === 'pieConnector' ? 'pie.secondaryPlot.connector.visible' : 'pie.label.visible'
  }
  return result
}

export type ChartElementPanelTarget =
  | { scope: 'chartArea' }
  | { scope: 'title'; titleKey: 'title' | 'subtitle' }
  | { scope: 'axisTitle'; axisId: ChartEditPanelAxis; titleKey: 'xAxisTitle' | 'yAxisTitle' | 'rightYAxisTitle' }
  | { scope: 'axis' | 'gridLine' | 'tick'; axisId: ChartEditPanelAxis }
  | { scope: 'legend' }
  | { scope: 'series' | 'dataPoint'; seriesId: string }
  | { scope: 'dataLabel'; seriesId?: string }
  | { scope: 'mark'; seriesId?: string; panel?: 'pieStyle' | 'paretoBarSeries' | 'paretoLineSeries' | 'waterfallSeries' | 'relation' | 'wordCloud' | 'data';
      directEdit?: { kind: 'seriesStyle' | 'waterfallStyle'; seriesId: string } | { kind: 'pieStyle' | 'pieSliceStyle' | 'pieAggregateStyle' | 'pieConnectorStyle' } | { kind: 'paretoStyle'; part: 'bar' | 'line' } }

/** Observable v1.6.0 context → route mapping, rewritten around local selector IDs. */
export function resolveChartEditPanelRoute(target: ChartElementPanelTarget): ChartEditPanelRoute | undefined {
  switch (target.scope) {
    case 'chartArea': return { tab: 'style', section: 'chartStyle' }
    case 'title': return { tab: 'style', section: 'chartAndAxisTitles', titleKey: target.titleKey }
    case 'axisTitle': return { tab: 'style', section: 'chartAndAxisTitles', axisId: target.axisId, titleKey: target.titleKey }
    case 'axis': case 'gridLine': case 'tick': return { tab: 'style', section: axisSection[target.axisId], axisId: target.axisId }
    case 'legend': return { tab: 'style', section: 'legend' }
    case 'series': case 'dataPoint': case 'dataLabel': return target.seriesId ? { tab: 'style', section: 'series', seriesId: target.seriesId } : undefined
    case 'mark': {
      const direct = target.directEdit
      if (direct?.kind === 'seriesStyle') return { tab: 'style', section: 'series', seriesId: direct.seriesId }
      if (direct?.kind === 'waterfallStyle') return { tab: 'style', section: 'waterfallSeries', seriesId: direct.seriesId }
      if (direct?.kind === 'pieStyle') return { tab: 'style', section: 'pieStyle' }
      if (direct?.kind === 'paretoStyle') return { tab: 'style', section: direct.part === 'bar' ? 'paretoBarSeries' : 'paretoLineSeries' }
      if (target.panel === 'data') return { tab: 'data' }
      if (target.panel) return { tab: 'style', section: target.panel, ...(target.panel === 'waterfallSeries' && target.seriesId ? { seriesId: target.seriesId } : {}) }
      return target.seriesId ? { tab: 'style', section: 'series', seriesId: target.seriesId } : undefined
    }
  }
}
