import { registerLayout, registerPostUpdate } from 'echarts/core'
import { chartTreemapLabelFits } from './chart-treemap'

export const CHART_TREEMAP_OVERLAP_KEY = '__officeTreemapOverlappingParentLabel'
export const CHART_TREEMAP_FIT_KEY = '__officeTreemapCompleteLabelFit'
interface NodeRect { x: number; y: number; width: number; height: number; borderWidth?: number; upperLabelHeight?: number }
interface RuntimeNode {
  dataIndex: number; viewChildren?: RuntimeNode[]
  getLayout(): NodeRect | undefined
  setLayout(rect: Partial<NodeRect>, merge: boolean): void
  getModel(): { get(path: string[]): unknown }
}
interface RuntimeGraphic {
  z2?: number; ignore?: boolean
  shape?: { width?: number; height?: number }
  style?: { padding?: number | number[] }
  beforeUpdate?: (...args: unknown[]) => void
  getTextContent?(): RuntimeGraphic | undefined
  getBoundingRect?(): { width: number; height: number }
  setStyle?(style: Record<string, unknown>): void
  ensureState?(name: string): { ignore?: boolean }
  traverse?(visit: (graphic: RuntimeGraphic) => void): void
  markRedraw?(): void
}
interface RuntimeData { tree: { root: RuntimeNode }; getItemGraphicEl(index: number): RuntimeGraphic | undefined }
interface RuntimeSeries { option: Record<string, unknown>; getData(): RuntimeData }
interface RuntimeModel { eachSeriesByType(type: string, visit: (series: RuntimeSeries) => void): void }
let registered = false

function nodes(root: RuntimeNode): RuntimeNode[] {
  const result: RuntimeNode[] = [], pending = [root]
  while (pending.length) {
    const node = pending.pop()!; result.push(node)
    for (const child of node.viewChildren ?? []) pending.push(child)
  }
  return result
}

/** ECharts reserves upperLabel.height. Refill that strip for the explicit
 * overlapping mode, then let the normal Treemap view draw the adjusted layout. */
function overlap(root: RuntimeNode): void {
  const pending = [root]
  while (pending.length) {
    const node = pending.pop()!, children = node.viewChildren ?? [], parent = node.getLayout()
    if (node.dataIndex >= 0 && children.length && parent) {
      let first = Infinity, last = -Infinity
      for (const child of children) {
        const rect = child.getLayout()
        if (rect) { first = Math.min(first, rect.y); last = Math.max(last, rect.y + rect.height) }
      }
      const span = last - first, gap = Number(node.getModel().get(['itemStyle', 'gapWidth']) ?? 0), inset = (parent.borderWidth ?? 0) - gap / 2
      const filled = last - inset, factor = filled / span
      if (span > 0 && filled > span && Number.isFinite(factor)) {
        for (const child of children) {
          const rect = child.getLayout()
          if (!rect) continue
          child.setLayout({ y: inset + (rect.y - first) * factor, height: rect.height * factor }, true)
          const descendants = [...(child.viewChildren ?? [])]
          while (descendants.length) {
            const descendant = descendants.pop()!, box = descendant.getLayout()
            if (box) descendant.setLayout({ y: box.y * factor, height: box.height * factor }, true)
            for (const next of descendant.viewChildren ?? []) descendants.push(next)
          }
        }
      }
    }
    for (let index = children.length - 1; index >= 0; index--) pending.push(children[index]!)
  }
}

function labels(data: RuntimeData, overlapping: boolean): void {
  const hierarchy = nodes(data.tree.root), subtreeZ = new Map<RuntimeNode, number>()
  for (let index = hierarchy.length - 1; index >= 0; index--) {
    const node = hierarchy[index]!, children = node.viewChildren ?? []
    if (node.dataIndex < 0) continue
    const graphic = data.getItemGraphicEl(node.dataIndex), rect = node.getLayout(), parent = children.length > 0
    let layer = Number.isFinite(graphic?.z2) ? graphic!.z2! : 0
    graphic?.traverse?.(entry => { if (Number.isFinite(entry.z2)) layer = Math.max(layer, entry.z2!) })
    for (const child of children) layer = Math.max(layer, subtreeZ.get(child) ?? 0)
    const hosts: RuntimeGraphic[] = []
    if (graphic?.getTextContent?.()) hosts.push(graphic)
    // TreemapView registers parent backgrounds directly, but leaf graphics are
    // groups with text attached to their content rectangle.
    if (!parent) graphic?.traverse?.(entry => { if (entry !== graphic && entry.getTextContent?.()) hosts.push(entry) })
    for (const host of hosts) {
      const text = host.getTextContent?.()
      if (!text || !rect) continue
      // Stock TreemapView hardcodes lineOverflow=truncate and a height. Clear
      // those clamps and its beforeUpdate setter before measuring the complete
      // wrapped string; otherwise the next zrender refresh restores the clamp.
      const area = parent ? { width: Math.max(0, rect.width - 2 * (rect.borderWidth ?? 0)), height: Math.min(rect.height, rect.upperLabelHeight ?? rect.height) }
        : { width: host.shape?.width ?? rect.width, height: host.shape?.height ?? rect.height }
      const raw = text.style?.padding, padding = typeof raw === 'number' ? [raw, raw, raw, raw] : raw ?? []
      const horizontal = padding.length === 1 ? 2 * padding[0]! : padding.length < 4 ? 2 * (padding[1] ?? 0) : padding[1]! + padding[3]!
      text.beforeUpdate = undefined
      text.setStyle?.({ width: Math.max(0, area.width - horizontal), lineOverflow: undefined, height: undefined, overflow: 'break' })
      const bounds = text.getBoundingRect?.()
      const visible = node.getModel().get([parent ? 'upperLabel' : 'label', 'show']) !== false
      const hidden = !visible || !chartTreemapLabelFits(area, bounds)
      text.ignore = hidden
      for (const state of ['emphasis', 'select', 'blur']) { const target = text.ensureState?.(state); if (target) target.ignore = hidden }
      if (parent && overlapping) text.z2 = layer + 1
      if (Number.isFinite(text.z2)) layer = Math.max(layer, text.z2!)
      text.markRedraw?.()
    }
    subtreeZ.set(node, layer)
  }
}

/** Adapter for the locked ECharts 6.1.0 Tree layout/view contract. Only marked
 * Office treemaps use it; there is no upstream SDK or private server dependency. */
export function registerChartTreemapLayout(): void {
  if (registered) return
  registered = true
  registerLayout(model => (model as unknown as RuntimeModel).eachSeriesByType('treemap', series => {
    if (series.option[CHART_TREEMAP_OVERLAP_KEY]) overlap(series.getData().tree.root)
  }))
  registerPostUpdate(model => (model as unknown as RuntimeModel).eachSeriesByType('treemap', series => {
    if (series.option[CHART_TREEMAP_FIT_KEY]) labels(series.getData(), Boolean(series.option[CHART_TREEMAP_OVERLAP_KEY]))
  }))
}
