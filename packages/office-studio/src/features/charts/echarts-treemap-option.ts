import type { TreemapSeriesOption } from 'echarts/charts'
import type { TooltipComponentOption } from 'echarts/components'
import type { ChartModel, ChartSeriesFillPattern } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { ChartEChartsContext } from './echarts-special-option'
import { chartHierarchyPointStyle } from './chart-derived'
import { chartColorOpacity } from './chart-geometry'
import { chartEChartsAutomaticGradient } from './chart-auto-gradient'
import { chartEChartsFont, chartEChartsIndependentItemStyle } from './echarts-style'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartTreemapBackgroundPaint, chartTreemapHeaderHeight, chartTreemapLabelBackground, chartTreemapLabelText,
  chartTreemapParentLayout, chartTreemapRenderNodes, chartTreemapTextColor, resolveChartTreemapLabel, type ChartTreemapRenderNode } from './chart-treemap'
import { CHART_TREEMAP_BACKGROUND_KEY, CHART_TREEMAP_FIT_KEY, CHART_TREEMAP_OVERLAP_KEY } from './echarts-treemap-layout'
import { CHART_HIERARCHY_PATTERN_KEY } from './echarts-hierarchy-pattern'
import { chartContentTooltipText, chartTreemapContentTooltip, type ChartHierarchyTooltipRuntime } from './chart-content-tooltip'

interface NativeTreemapNode {
  id: string; name: string; value: number; semanticNode: ChartTreemapRenderNode['node']; semanticParentId?: string; isTreemapResidual?: boolean
  itemStyle: ReturnType<typeof chartEChartsIndependentItemStyle> & { gapWidth: number }
  label: ReturnType<typeof chartEChartsFont> & { show: boolean; formatter: string; overflow: 'break'; lineHeight: number; backgroundColor?: string }
  [CHART_TREEMAP_BACKGROUND_KEY]?: ReturnType<typeof chartEChartsAutomaticGradient>
  [CHART_HIERARCHY_PATTERN_KEY]?: ChartSeriesFillPattern
  upperLabel: ReturnType<typeof chartEChartsFont> & { show: boolean; formatter: string; overflow: 'break'; height: number; lineHeight: number }
  children?: NativeTreemapNode[]
}
export interface ChartTreemapPointerRuntime extends ChartHierarchyTooltipRuntime { mode: 'treemap' }

export function createChartTreemapPointerRuntime(data: ChartData): ChartTreemapPointerRuntime {
  const context = { fieldName: data.hierarchyValueFieldName ?? data.series[0]?.name ?? '' }
  return { mode: 'treemap', contentTooltip: item => chartTreemapContentTooltip(item, context) }
}

export function chartTreemapEChartsData(data: ChartData, chart: ChartModel, context: ChartEChartsContext): NativeTreemapNode[] {
  const roots = chartTreemapRenderNodes(data.hierarchy ?? []), source = data.series[0]!, mode = chartTreemapParentLayout(chart.options?.treemap)
  const result = new Map<ChartTreemapRenderNode, NativeTreemapNode>(), pending = roots.map(node => ({ node, visited: false })).reverse()
  while (pending.length) {
    const entry = pending.pop()!, node = entry.node
    if (!entry.visited) {
      pending.push({ node, visited: true })
      for (let index = (node.children?.length ?? 0) - 1; index >= 0; index--) pending.push({ node: node.children![index]!, visited: false })
      continue
    }
    const style = chartHierarchyPointStyle(chart, source.field, node, context.palette, source.sourceName ?? source.name)
    const label = resolveChartTreemapLabel(chart, style.label), textColor = chartTreemapTextColor(style.color!, label.color, context.color)
    const font = chartEChartsFont(label, textColor, 12), background = chartTreemapLabelBackground(label.backgroundColor)
    const backgroundPaint = chartTreemapBackgroundPaint(background, { x: 0, y: 0, width: 1, height: 1 })
    const gradient = backgroundPaint ? chartEChartsAutomaticGradient({ ...backgroundPaint, bounds: undefined }) : undefined
    const pattern = chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1)
    const gap = node.depth === 0 ? 2 : 1
    result.set(node, { id: node.id, name: node.name, value: node.value, semanticNode: node.node,
      ...(node.isTreemapResidual ? { isTreemapResidual: true, semanticParentId: node.semanticParentId } : {}),
      // The locked TreemapView consumes fill color but omits itemStyle.opacity
      // from its shape mapper. Embed only the fill multiplier into that color,
      // preserving the separate background Rect's authored outline alpha.
      itemStyle: { ...chartEChartsIndependentItemStyle(style), borderWidth: style.border?.visible === false ? 0 : style.border?.width ?? gap,
        borderColor: chartColorOpacity(style.border?.color ?? chart.options?.appearance?.background ?? context.backgroundColor ?? '#ffffff', style.border?.opacity ?? 1), gapWidth: gap },
      label: { ...font, show: label.visible, formatter: chartTreemapLabelText(node, label, data.date1904), overflow: 'break', lineHeight: label.fontSize * 1.2,
        backgroundColor: typeof background === 'string' ? background : undefined },
      ...(gradient ? { [CHART_TREEMAP_BACKGROUND_KEY]: gradient } : {}),
      ...(pattern ? { [CHART_HIERARCHY_PATTERN_KEY]: pattern } : {}),
      upperLabel: { ...font, show: mode !== 'none', formatter: node.node.name, overflow: 'break', height: chartTreemapHeaderHeight(mode, label.fontSize), lineHeight: label.fontSize },
      ...(node.children?.length ? { children: node.children.map(child => result.get(child)!) } : {}) })
  }
  return roots.map(root => result.get(root)!)
}

/** Tooltip identity is semantic: residual leaves retain the owning parent path
 * and signed self value, rather than exposing synthetic IDs as categories. */
export function chartTreemapTooltip(data: ChartData, context?: ChartEChartsContext): TooltipComponentOption {
  const pointer = createChartTreemapPointerRuntime(data)
  return { show: true, showContent: false, trigger: 'item', triggerOn: 'none', renderMode: 'richText',
    showDelay: 0, hideDelay: 0, transitionDuration: 0, confine: true, backgroundColor: context?.backgroundColor ?? '#ffffff',
    borderColor: context?.axisColor, textStyle: { color: context?.color ?? '#334155' }, formatter: params => {
    const item = (Array.isArray(params) ? params[0] : params)?.data
    return chartContentTooltipText(pointer.contentTooltip?.(item))
  } }
}

export function createChartTreemapEChartsSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): TreemapSeriesOption {
  const { grid } = context, mode = chartTreemapParentLayout(chart.options?.treemap), label = resolveChartTreemapLabel(chart)
  const series: TreemapSeriesOption & { [CHART_TREEMAP_OVERLAP_KEY]: boolean; [CHART_TREEMAP_FIT_KEY]: boolean } = {
    type: 'treemap', name: data.series[0]!.name, left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
    animation: false, stateAnimation: { duration: 0 }, data: chartTreemapEChartsData(data, chart, context), nodeClick: false, roam: false, breadcrumb: { show: false }, visibleMin: 0,
    // The shared render tree is already sorted stably. Native desc reverses
    // equal-value data indices, which would change the fixed source order.
    sort: false, label: { show: label.visible, ...chartEChartsFont(label, context.color, 12), overflow: 'break', lineHeight: label.fontSize * 1.2 },
    upperLabel: { show: mode !== 'none', fontSize: label.fontSize, height: chartTreemapHeaderHeight(mode, label.fontSize), lineHeight: label.fontSize },
    itemStyle: { borderWidth: 1, gapWidth: 1, borderColor: chart.options?.appearance?.background ?? context.backgroundColor ?? '#ffffff' },
    levels: [{ itemStyle: { borderWidth: 2, gapWidth: 2 } }, { itemStyle: { borderWidth: 1, gapWidth: 1 } }], emphasis: { focus: 'ancestor' },
    [CHART_TREEMAP_OVERLAP_KEY]: mode === 'overlapping', [CHART_TREEMAP_FIT_KEY]: true,
  }
  return series
}
