import { chartTreemapValueText } from './chart-treemap'

/** Presentation facts shared by the editor's Ant Tooltip host. These values
 * belong to a rendered view, never to content history or the saved model. */
export interface ChartContentTooltip { title: string; lines: string[] }
export interface ChartHierarchyContentTooltip extends ChartContentTooltip { id: string; kind: 'hierarchy-node'; index: number }
export interface ChartHierarchyTooltipContext { fieldName: string; format?: string; date1904?: boolean; total?: number }
export interface ChartHierarchyTooltipRuntime {
  mode: 'treemap' | 'sunburst'
  contentTooltip?: (item: unknown) => ChartContentTooltip | undefined
}

function record(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value))
}
function finite(value: unknown): value is number { return typeof value === 'number' && Number.isFinite(value) }
function path(value: unknown): value is string[] { return Array.isArray(value) && value.every(part => typeof part === 'string') }
function optionalText(value: unknown): value is string | undefined { return value === undefined || typeof value === 'string' }

/** A Treemap residual carries the owning semantic node. Keep its signed
 * total/path and formatted own value, rather than its synthetic leaf ID. */
export function chartTreemapContentTooltip(item: unknown, context: ChartHierarchyTooltipContext): ChartContentTooltip | undefined {
  if (!record(item) || !record(item.semanticNode)) return undefined
  const node = item.semanticNode
  if (typeof node.name !== 'string' || !finite(node.value) || !optionalText(node.formattedValue)
    || node.path !== undefined && !path(node.path) || node.selfValue !== undefined && !finite(node.selfValue)
    || !optionalText(node.formattedSelfValue) || node.children !== undefined && !Array.isArray(node.children)) return undefined
  const lines = [`${context.fieldName}: ${chartTreemapValueText(node.value, node.formattedValue)}`]
  if (Array.isArray(node.children) && node.children.length && finite(node.selfValue) && node.selfValue !== 0) {
    const childValue = node.value - node.selfValue
    if (!Number.isFinite(childValue)) return undefined
    lines.push(`自身数值: ${chartTreemapValueText(node.selfValue, node.formattedSelfValue)}`, `子节点合计: ${childValue}`)
  }
  return { title: path(node.path) ? node.path.join(' / ') : node.name, lines }
}

/** Sunburst residuals report their rendered positive own fact and retain the
 * owning parent's path and signed own/children facts, matching fixed XOe. */
export function chartSunburstContentTooltip(item: unknown, context: ChartHierarchyTooltipContext): ChartContentTooltip | undefined {
  if (!record(item) || !finite(item.value) || !path(item.semanticPath) || !finite(item.semanticSelfValue)
    || !finite(item.semanticChildValue) || !optionalText(item.formattedValue)) return undefined
  const percentage = context.total && context.total > 0 ? item.value / context.total * 100 : 0
  if (!Number.isFinite(percentage)) return undefined
  const lines = [`${context.fieldName}: ${chartTreemapValueText(item.value, item.formattedValue, context.format, context.date1904)} (${percentage.toFixed(1)}%)`]
  if (item.semanticChildValue !== 0 && item.semanticSelfValue !== 0) {
    lines.push(`自身数值: ${item.semanticSelfValue}`, `子节点合计: ${item.semanticChildValue}`)
  }
  return { title: item.semanticPath.join(' / '), lines }
}

export function chartContentTooltipText(tooltip?: ChartContentTooltip): string {
  return tooltip ? [tooltip.title, ...tooltip.lines].join('\n') : ''
}
