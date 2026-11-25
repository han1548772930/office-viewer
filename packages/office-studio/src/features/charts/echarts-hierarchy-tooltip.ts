import type { ChartHierarchyContentTooltip, ChartHierarchyTooltipRuntime } from './chart-content-tooltip'

type MouseHandler = (event: unknown) => void
export interface ChartHierarchyMouseEvents {
  on(name: string, handler: MouseHandler): unknown
  off(name: string, handler: MouseHandler): unknown
}

/** Capture params.data from real native item events. The renderer resets the
 * capture before forwarding a point, so a blank hit cannot reuse the last
 * node. Each binding belongs to one ECharts instance and releases its facts
 * and exactly its own handlers on replacement/disposal. */
export class ChartNativeHierarchyTooltipBinding {
  private item?: { mode: ChartHierarchyTooltipRuntime['mode']; data: unknown; index: number; seriesIndex: number }
  private disposed = false
  private readonly move: MouseHandler = event => {
    this.item = undefined
    if (this.disposed || !event || typeof event !== 'object') return
    const params = event as Record<string, unknown>
    if (params.componentType === 'series' && (params.seriesType === 'treemap' || params.seriesType === 'sunburst')
      && typeof params.dataIndex === 'number' && Number.isSafeInteger(params.dataIndex) && params.dataIndex >= 0) {
      this.item = { mode: params.seriesType, data: params.data, index: params.dataIndex,
        seriesIndex: typeof params.seriesIndex === 'number' && Number.isSafeInteger(params.seriesIndex) && params.seriesIndex >= 0 ? params.seriesIndex : 0 }
    }
  }
  private readonly leave: MouseHandler = () => { this.clear() }
  private readonly handlers: ReadonlyArray<readonly [string, MouseHandler]> = [
    ['mouseover', this.move], ['mousemove', this.move], ['mouseout', this.leave], ['globalout', this.leave],
  ]

  constructor(private readonly events: ChartHierarchyMouseEvents) {
    const bound: Array<readonly [string, MouseHandler]> = []
    try {
      for (const [name, handler] of this.handlers) { bound.push([name, handler]); events.on(name, handler) }
    } catch (error) {
      this.disposed = true
      this.clear()
      this.releaseHandlers(bound)
      throw error
    }
  }

  resolve(runtime: ChartHierarchyTooltipRuntime): ChartHierarchyContentTooltip | undefined {
    const item = this.item, content = !this.disposed && item?.mode === runtime.mode ? runtime.contentTooltip?.(item.data) : undefined
    return content && item ? { ...content, id: `${item.mode}:${item.seriesIndex}:${item.index}`, kind: 'hierarchy-node', index: item.index } : undefined
  }
  clear(): void { this.item = undefined }
  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.clear()
    const failures = this.releaseHandlers(this.handlers)
    if (failures.length) throw failures[0]
  }
  private releaseHandlers(handlers: ReadonlyArray<readonly [string, MouseHandler]>): unknown[] {
    const failures: unknown[] = []
    for (const [name, handler] of handlers) {
      try { this.events.off(name, handler) }
      catch (error) { failures.push(error) }
    }
    return failures
  }
}
