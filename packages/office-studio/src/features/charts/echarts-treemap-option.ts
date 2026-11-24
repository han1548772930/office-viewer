import type { TreemapSeriesOption } from 'echarts/charts'
import type { TooltipComponentOption } from 'echarts/components'
import type { ChartModel } from '../../domain/model'
import type { ChartData } from './chart-engine'
import type { ChartEChartsContext } from './echarts-special-option'
import { chartHierarchyPointStyle } from './chart-derived'
import { chartColorOpacity } from './chart-geometry'
import { chartEChartsAutomaticGradient } from './chart-auto-gradient'
import { chartEChartsFont, chartEChartsItemStyle } from './echarts-style'
import { chartTreemapBackgroundPaint, chartTreemapHeaderHeight, chartTreemapLabelBackground, chartTreemapLabelText,
  chartTreemapParentLayout, chartTreemapRenderNodes, chartTreemapTextColor, resolveChartTreemapLabel, type ChartTreemapRenderNode } from './chart-treemap'
import { CHART_TREEMAP_FIT_KEY, CHART_TREEMAP_OVERLAP_KEY } from './echarts-treemap-layout'

interface NativeTreemapNode {
  id: string; name: string; value: number; semanticNode: ChartTreemapRenderNode['node']; semanticParentId?: string; isTreemapResidual?: boolean
  itemStyle: ReturnType<typeof chartEChartsItemStyle> & { gapWidth: number }
  label: ReturnType<typeof chartEChartsFont> & { show: boolean; formatter: string; overflow: 'break'; lineHeight: number; backgroundColor?: ReturnType<typeof chartEChartsAutomaticGradient> | string }
  upperLabel: ReturnType<typeof chartEChartsFont> & { show: boolean; formatter: string; overflow: 'break'; height: number; lineHeight: number }
  children?: NativeTreemapNode[]
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
    const gap = node.depth === 0 ? 2 : 1
    result.set(node, { id: node.id, name: node.name, value: node.value, semanticNode: node.node,
      ...(node.isTreemapResidual ? { isTreemapResidual: true, semanticParentId: node.semanticParentId } : {}),
      itemStyle: { ...chartEChartsItemStyle(style), borderWidth: style.border?.visible === false ? 0 : style.border?.width ?? gap,
        borderColor: chartColorOpacity(style.border?.color ?? chart.options?.appearance?.background ?? context.backgroundColor ?? '#ffffff', style.border?.opacity ?? 1), gapWidth: gap },
      label: { ...font, show: label.visible, formatter: chartTreemapLabelText(node, label, data.date1904), overflow: 'break', lineHeight: label.fontSize * 1.2,
        backgroundColor: backgroundPaint ? chartEChartsAutomaticGradient({ ...backgroundPaint, bounds: undefined }) : typeof background === 'string' ? background : undefined },
      upperLabel: { ...font, show: mode !== 'none', formatter: node.node.name, overflow: 'break', height: chartTreemapHeaderHeight(mode, label.fontSize), lineHeight: label.fontSize },
      ...(node.children?.length ? { children: node.children.map(child => result.get(child)!) } : {}) })
  }
  return roots.map(root => result.get(root)!)
}

/** Tooltip identity is semantic: residual leaves retain the owning parent path
 * and signed self value, rather than exposing synthetic IDs as categories. */
export function chartTreemapTooltip(data: ChartData): TooltipComponentOption {
  return { show: true, showContent: true, trigger: 'item', triggerOn: 'mousemove|click', renderMode: 'richText', formatter: params => {
    const item = (Array.isArray(params) ? params[0] : params)?.data
    const node = item && typeof item === 'object' && 'semanticNode' in item ? item.semanticNode as NativeTreemapNode['semanticNode'] : undefined
    if (!node) return ''
    const lines = [node.path?.join(' / ') ?? node.name, `${data.hierarchyValueFieldName ?? data.series[0]?.name ?? ''}: ${node.value}`]
    if (node.children?.length && node.selfValue !== 0) lines.push(`自身数值: ${node.selfValue}`, `子节点合计: ${node.value - node.selfValue}`)
    return lines.join('\n')
  } }
}

export function createChartTreemapEChartsSeries(data: ChartData, chart: ChartModel, context: ChartEChartsContext): TreemapSeriesOption {
  const { grid } = context, mode = chartTreemapParentLayout(chart.options?.treemap), label = resolveChartTreemapLabel(chart)
  const series: TreemapSeriesOption & { [CHART_TREEMAP_OVERLAP_KEY]: boolean; [CHART_TREEMAP_FIT_KEY]: boolean } = {
    type: 'treemap', name: data.series[0]!.name, left: grid.left, right: grid.right, top: grid.top, bottom: grid.bottom,
    animation: false, data: chartTreemapEChartsData(data, chart, context), nodeClick: false, roam: false, breadcrumb: { show: false }, visibleMin: 0,
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
