import { registerPostUpdate } from 'echarts/core'
import { chartSunburstLabelFits } from './chart-sunburst'
import { applyChartHierarchyPattern, CHART_HIERARCHY_PATTERN_KEY } from './echarts-hierarchy-pattern'

export const CHART_SUNBURST_FIT_KEY = '__officeSunburstCompleteLabelFit'
export const CHART_SUNBURST_BACKGROUND_KEY = '__officeSunburstLabelBackgroundGradient'
interface RuntimeRect { width: number; height: number }
interface RuntimeNode { dataIndex: number; children?: RuntimeNode[]; getModel(): { get(path: string[]): unknown } }
interface RuntimeGraphic {
  ignore?: boolean; rotation?: number
  getTextContent?(): RuntimeGraphic | undefined
  getBoundingRect?(): RuntimeRect
  getComputedTransform?(): ArrayLike<number> | null | undefined
  setStyle?(style: Record<string, unknown>): void
  ensureState?(name: string): { ignore?: boolean; style?: Record<string, unknown> }
  markRedraw?(): void
}
interface RuntimeData { tree: { root: RuntimeNode }; getItemGraphicEl(index: number): RuntimeGraphic | undefined }
interface RuntimeSeries { option: Record<string, unknown>; getData(): RuntimeData }
interface RuntimeModel { eachSeriesByType(type: string, visit: (series: RuntimeSeries) => void): void }
let registered = false

function transformedSize(graphic: RuntimeGraphic, bounds: RuntimeRect): RuntimeRect {
  const transform = graphic.getComputedTransform?.(), angle = graphic.rotation ?? 0
  const a = transform?.[0] ?? Math.cos(angle), b = transform?.[1] ?? Math.sin(angle)
  const c = transform?.[2] ?? -Math.sin(angle), d = transform?.[3] ?? Math.cos(angle)
  return { width: Math.abs(a) * bounds.width + Math.abs(c) * bounds.height, height: Math.abs(b) * bounds.width + Math.abs(d) * bounds.height }
}

function completeLabels(data: RuntimeData, devicePixelRatio: number): void {
  const pending = [data.tree.root]
  while (pending.length) {
    const node = pending.pop()!
    for (const child of node.children ?? []) pending.push(child)
    if (node.dataIndex < 0) continue
    const host = data.getItemGraphicEl(node.dataIndex)
    if (!host) continue
    const model = node.getModel()
    applyChartHierarchyPattern(host, model.get([CHART_HIERARCHY_PATTERN_KEY]), devicePixelRatio)
    const text = host.getTextContent?.()
    if (!text) continue
    const background = model.get([CHART_SUNBURST_BACKGROUND_KEY])
    // The public label type permits string/image backgrounds. The locked
    // zrender Text runtime also paints gradients; keep this at that boundary.
    text.setStyle?.({ width: undefined, height: undefined, overflow: undefined, lineOverflow: undefined,
      ...(background ? { backgroundColor: background } : {}) })
    const area = host.getBoundingRect?.(), local = text.getBoundingRect?.()
    const rect = area && transformedSize(host, area), labelRect = local && transformedSize(text, local)
    const hidden = model.get(['label', 'show']) === false || !chartSunburstLabelFits(rect, labelRect)
    text.ignore = hidden
    for (const name of ['emphasis', 'select', 'blur']) {
      const state = text.ensureState?.(name)
      if (state) { state.ignore = hidden; if (background) state.style = { ...state.style, backgroundColor: background } }
    }
    text.markRedraw?.()
  }
}

/** Fixed QOe uses a hide field that ECharts 6.1.0's public LabelLayoutOption
 * does not implement. This marked adapter applies the same complete rotated
 * rectangle comparison, and keeps hidden labels hidden in every view state. */
export function registerChartSunburstLayout(): void {
  if (registered) return
  registered = true
  registerPostUpdate((model, api) => (model as unknown as RuntimeModel).eachSeriesByType('sunburst', series => {
    if (series.option[CHART_SUNBURST_FIT_KEY]) completeLabels(series.getData(), api.getDevicePixelRatio())
  }))
}
