import type { ECharts, EChartsOption } from 'echarts'
import type { RenderRect } from '../../rendering/contracts.ts'
import type { ChartVector } from './chart-geometry.ts'
import { readPivotChartAxisLabelInfo, type PivotChartCategoryLabelInfo } from './chart-pivot-axis-layout.ts'
import { transformedBounds } from './echarts-element-interaction.ts'

type PaintedElement = ReturnType<ReturnType<ECharts['getZr']>['storage']['getDisplayList']>[number]
export interface PivotChartCanvasAxisTarget { key: string; label: PivotChartCategoryLabelInfo; bounds: RenderRect }

const record = (value: unknown): Record<string, unknown> | undefined => value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : undefined

export function pivotChartAxisGraphicLabels(option: EChartsOption): Map<string, PivotChartCategoryLabelInfo> {
  const labels = new Map<string, PivotChartCategoryLabelInfo>()
  const visit = (elements: unknown): void => {
    if (!Array.isArray(elements)) return
    for (const raw of elements) {
      const element = record(raw), label = readPivotChartAxisLabelInfo(record(element?.info)?.officePivotAxisLabel)
      if (element && typeof element.id === 'string' && label) labels.set(element.id, label)
      visit(element?.children)
    }
  }
  if (Array.isArray(option.graphic)) for (const graphic of option.graphic) visit(graphic.elements)
  else visit(option.graphic?.elements)
  return labels
}

/** Uses the exact painted text, native transforms, clipping and z-order. No estimated hit rectangles. */
export class ChartNativePivotAxisIndex {
  private readonly descriptors = new WeakMap<object, PivotChartCategoryLabelInfo>()
  private readonly targets = new Map<string, PivotChartCanvasAxisTarget>()
  constructor(private readonly instance: ECharts, labels: ReadonlyMap<string, PivotChartCategoryLabelInfo>, width: number, height: number) {
    for (const root of instance.getZr().storage.getRoots()) root.traverse(element => {
      const label = labels.get(String(element.id))
      if (label) this.descriptors.set(element, label)
    })
    const frame = { x: 0, y: 0, width, height }
    for (const painted of instance.getZr().storage.getDisplayList(true)) {
      if (!painted.shouldBePainted(width, height, true, true)) continue
      const label = this.descriptor(painted), bounds = label && transformedBounds(painted, frame)
      if (!label || !bounds || bounds.width <= 0 || bounds.height <= 0) continue
      // ZRender can split one Text into several painted spans. Keep all of the
      // native geometry associated with the same authored label.
      const previous = this.targets.get(label.key)?.bounds
      const x = Math.min(previous?.x ?? bounds.x, bounds.x), y = Math.min(previous?.y ?? bounds.y, bounds.y)
      this.targets.set(label.key, { key: label.key, label, bounds: { x, y,
        width: Math.max(previous ? previous.x + previous.width : bounds.x + bounds.width, bounds.x + bounds.width) - x,
        height: Math.max(previous ? previous.y + previous.height : bounds.y + bounds.height, bounds.y + bounds.height) - y } })
    }
  }
  private descriptor(element: PaintedElement): PivotChartCategoryLabelInfo | undefined {
    let current: object | undefined = element
    while (current) {
      const label = this.descriptors.get(current)
      if (label) return label
      const parent: unknown = '__hostTarget' in current && current.__hostTarget || 'parent' in current && current.parent
      current = parent && typeof parent === 'object' ? parent : undefined
    }
    return undefined
  }
  hit(point?: ChartVector): PivotChartCanvasAxisTarget | undefined {
    if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y)) return undefined
    const painted = this.instance.getZr().findHover(point.x, point.y).topTarget
    const label = painted && this.descriptor(painted)
    return label && this.targets.get(label.key)
  }
  target(key: string): PivotChartCanvasAxisTarget | undefined { return this.targets.get(key) }
  list(): PivotChartCanvasAxisTarget[] { return [...this.targets.values()] }
}
