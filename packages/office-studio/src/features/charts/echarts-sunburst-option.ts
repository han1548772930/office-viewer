import type { SunburstSeriesOption } from 'echarts/charts'
import type { GraphicComponentOption, TooltipComponentOption } from 'echarts/components'
import type { ChartModel, ChartSeriesFillPattern } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { ChartEChartsContext } from './echarts-special-option'
import { chartHierarchyPointStyle } from './chart-derived'
import { chartColorOpacity } from './chart-geometry'
import { chartEChartsAutomaticGradient } from './chart-auto-gradient'
import { chartEChartsFont, chartEChartsItemStyle } from './echarts-style'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartLabelBackground, chartLabelBackgroundPaint } from './chart-label-background'
import { CHART_SUNBURST_INNER_RADIUS_RATIO, CHART_SUNBURST_OUTER_RADIUS_RATIO, chartSunburstCenterText,
  chartSunburstLabelText, chartSunburstLabelVisible, chartSunburstRenderNodes, chartSunburstTotal,
  resolveChartSunburstCenterLabel, resolveChartSunburstLabel, type ChartSunburstRenderNode } from './chart-sunburst'
import { CHART_SUNBURST_BACKGROUND_KEY, CHART_SUNBURST_FIT_KEY } from './echarts-sunburst-layout'
import { CHART_HIERARCHY_PATTERN_KEY } from './echarts-hierarchy-pattern'
import { chartContentTooltipText, chartSunburstContentTooltip, type ChartHierarchyTooltipRuntime } from './chart-content-tooltip'

type NativeSunburstItem = NonNullable<SunburstSeriesOption['data']>[number]
export interface ChartSunburstEChartsNode extends Omit<NativeSunburstItem, 'children' | 'value'> {
  id: string; name: string; value: number; formattedValue?: string
  semanticPath: string[]; semanticSelfValue: number; semanticChildValue: number; semanticParentId?: string; isSunburstResidual?: boolean
  label: NonNullable<NativeSunburstItem['label']> & { show: boolean; formatter: () => string }
  [CHART_SUNBURST_BACKGROUND_KEY]?: ReturnType<typeof chartEChartsAutomaticGradient>
  [CHART_HIERARCHY_PATTERN_KEY]?: ChartSeriesFillPattern
  children?: ChartSunburstEChartsNode[]
}
export interface ChartSunburstPointerRuntime extends ChartHierarchyTooltipRuntime { mode: 'sunburst' }

export function createChartSunburstPointerRuntime(data: ChartData, chart: ChartModel): ChartSunburstPointerRuntime {
  const context = { fieldName: data.hierarchyValueFieldName ?? data.series[0]?.name ?? '',
    total: chartSunburstTotal(data.hierarchy ?? []), format: resolveChartSunburstLabel(chart).format, date1904: data.date1904 }
  return { mode: 'sunburst', contentTooltip: item => chartSunburstContentTooltip(item, context) }
}

/** The option contains flat semantic facts and one native child tree. It does
 * not serialize a duplicate semantic subtree into every descendant. */
export function chartSunburstEChartsData(data: ChartData, chart: ChartModel, context: ChartEChartsContext): ChartSunburstEChartsNode[] {
  const roots = chartSunburstRenderNodes(data.hierarchy ?? []), total = chartSunburstTotal(data.hierarchy ?? []), source = data.series[0]!
  const result = new Map<ChartSunburstRenderNode, ChartSunburstEChartsNode>(), pending = roots.map(node => ({ node, visited: false })).reverse()
  while (pending.length) {
    const entry = pending.pop()!, node = entry.node
    if (!entry.visited) {
      pending.push({ node, visited: true })
      for (let index = (node.children?.length ?? 0) - 1; index >= 0; index--) pending.push({ node: node.children![index]!, visited: false })
      continue
    }
    const style = chartHierarchyPointStyle(chart, source.field, node, context.palette, source.sourceName ?? source.name), label = resolveChartSunburstLabel(chart, style.label)
    const text = chartSunburstLabelText(node, label, data.date1904), visible = chartSunburstLabelVisible(node, total, label) && text.length > 0
    const background = chartLabelBackground(label.backgroundColor), paint = chartLabelBackgroundPaint(background, { x: 0, y: 0, width: 1, height: 1 })
    const gradient = paint && chartEChartsAutomaticGradient({ ...paint, bounds: undefined }), pattern = chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1)
    result.set(node, { id: node.id, name: node.name, value: node.value, formattedValue: node.formattedValue,
      semanticPath: node.path, semanticSelfValue: node.node.selfValue, semanticChildValue: node.node.value - node.node.selfValue,
      ...(node.isSunburstResidual ? { semanticParentId: node.semanticParentId, isSunburstResidual: true } : {}),
      itemStyle: { ...chartEChartsItemStyle(style), borderColor: chartColorOpacity(style.border?.color ?? chart.options?.appearance?.background ?? context.backgroundColor ?? '#ffffff', style.border?.opacity ?? 1),
        borderWidth: style.border?.visible === false ? 0 : style.border?.width ?? 2 },
      label: { ...chartEChartsFont(label, context.color, 12), show: visible, formatter: () => text, position: label.position,
        rotate: node.depth === 0 ? 'tangential' : 'radial', lineHeight: label.fontSize * 1.2, backgroundColor: typeof background === 'string' ? background : undefined },
      emphasis: { label: { show: visible } }, select: { label: { show: visible } }, blur: { label: { show: visible } },
      ...(gradient ? { [CHART_SUNBURST_BACKGROUND_KEY]: gradient } : {}), ...(pattern ? { [CHART_HIERARCHY_PATTERN_KEY]: pattern } : {}),
      ...(node.children?.length ? { children: node.children.map(child => result.get(child)!) } : {}) })
  }
  return roots.map(root => result.get(root)!)
}

/** Residual tooltips show their own positive fact but retain the owning
 * parent's path and signed own/children facts, matching fixed XOe. */
export function chartSunburstTooltip(data: ChartData, chart: ChartModel, context: ChartEChartsContext): TooltipComponentOption {
  const pointer = createChartSunburstPointerRuntime(data, chart)
  return { show: true, showContent: false, trigger: 'item', triggerOn: 'none', renderMode: 'richText',
    showDelay: 0, hideDelay: 0, transitionDuration: 0, confine: true, backgroundColor: context.backgroundColor ?? '#ffffff',
    borderColor: context.axisColor, textStyle: { color: context.color }, formatter: params => {
      const node = (Array.isArray(params) ? params[0] : params)?.data
      return chartContentTooltipText(pointer.contentTooltip?.(node))
    } }
}

export function createChartSunburstEChartsSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): SunburstSeriesOption {
  const { grid } = context, width = Math.max(0, chart.size.width - grid.left - grid.right), height = Math.max(0, chart.size.height - grid.top - grid.bottom)
  const radius = Math.min(width, height) / 2, label = resolveChartSunburstLabel(chart)
  const series: SunburstSeriesOption & { [CHART_SUNBURST_FIT_KEY]: boolean } = {
    type: 'sunburst', name: data.series[0]!.name, data: chartSunburstEChartsData(data, chart, context), animation: false, stateAnimation: { duration: 0 },
    center: [grid.left + width / 2, grid.top + height / 2], radius: [radius * CHART_SUNBURST_INNER_RADIUS_RATIO, radius * CHART_SUNBURST_OUTER_RADIUS_RATIO],
    clockwise: true, startAngle: 90, minAngle: 0, nodeClick: false, stillShowZeroSum: false,
    label: { ...chartEChartsFont(label, context.color, 12), show: label.visible, lineHeight: label.fontSize * 1.2 },
    itemStyle: { borderWidth: 2, borderColor: chart.options?.appearance?.background ?? context.backgroundColor ?? '#ffffff' },
    emphasis: { focus: 'ancestor' }, [CHART_SUNBURST_FIT_KEY]: true,
  }
  // The locked runtime explicitly checks sort != null. Its public type omits
  // null, so apply the fixed no-sort value at this native adapter boundary.
  Object.assign(series, { sort: null })
  return series
}

export function chartSunburstCenterGraphic(data: ChartData, chart: ChartModel, context: ChartEChartsContext): GraphicComponentOption | undefined {
  const label = resolveChartSunburstCenterLabel(chart)
  if (!label.visible) return undefined
  const { grid } = context
  return { elements: [{ id: `${chart.id}:sunburst-center`, type: 'text', silent: true,
    x: (chart.size.width + grid.left - grid.right) / 2, y: (chart.size.height + grid.top - grid.bottom) / 2,
    style: { text: chartSunburstCenterText(data.hierarchy ?? [], label, data.date1904), fill: label.color ?? context.color, fontFamily: label.family,
      fontSize: label.fontSize, fontWeight: label.bold ? 'bold' : 'normal', fontStyle: label.italic ? 'italic' : 'normal', align: 'center', verticalAlign: 'middle' } }] }
}
