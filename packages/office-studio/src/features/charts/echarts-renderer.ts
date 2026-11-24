import * as echarts from 'echarts/core'
import { BarChart, BoxplotChart, CandlestickChart, FunnelChart, GaugeChart, GraphChart, HeatmapChart, LineChart, PieChart, RadarChart, SankeyChart, ScatterChart, SunburstChart, TreemapChart } from 'echarts/charts'
import { AxisPointerComponent, GraphicComponent, GridComponent, LegendComponent, RadarComponent, TitleComponent, TooltipComponent, VisualMapComponent } from 'echarts/components'
import { CanvasRenderer } from 'echarts/renderers'
import type { ECharts, EChartsOption } from 'echarts'
import type { GraphicComponentOption } from 'echarts/components'
import { chartGraphics, type ChartEChartsTheme } from './echarts-option'
import { chartAxisPointerCommands, type ChartAxisPointerRuntime } from './chart-axis-pointer'
import type { ChartVector } from './chart-geometry'
import { moveChartSankeyScene, type ChartSankeyScene } from './chart-sankey-draw'
import { chartSankeyFocusedCommands, chartSankeyTrajectory, createChartSankeyHitIndex, hitChartSankey,
  type ChartSankeyHit, type ChartSankeyPointerRuntime } from './chart-sankey-interaction'
import { registerChartTreemapLayout } from './echarts-treemap-layout'

echarts.use([
  BarChart, BoxplotChart, CandlestickChart, FunnelChart, GaugeChart, GraphChart, HeatmapChart, LineChart, PieChart, RadarChart, SankeyChart, ScatterChart, SunburstChart, TreemapChart,
  AxisPointerComponent, GraphicComponent, GridComponent, LegendComponent, RadarComponent, TitleComponent, TooltipComponent, VisualMapComponent, CanvasRenderer,
])
registerChartTreemapLayout()

export interface EChartsCanvasSize {
  width: number
  height: number
  devicePixelRatio: number
}

export interface EChartsCanvasFrame {
  source: HTMLCanvasElement
  sourceVersion: number
}

export interface EChartsCanvasContentDragStart {
  kind: 'sankey-node'
  nodeIndex: number
  frame?: EChartsCanvasFrame
}

interface SankeyContentDrag {
  nodeIndex: number
  offset: ChartVector
  originalScene: ChartSankeyScene
}

interface ChartEntry {
  instance: ECharts
  canvas: HTMLCanvasElement
  signature: string
  sourceVersion: number
  devicePixelRatio: number
  width: number
  height: number
  pointer?: ChartAxisPointerRuntime | ChartSankeyPointerRuntime
  pointerSignature: string
  contentDrag?: SankeyContentDrag
}

/** Owns the raster surfaces used by spreadsheet chart draw commands. */
export class EChartsCanvasRenderer {
  private readonly charts = new Map<string, ChartEntry>()

  render(chartId: string, option: EChartsOption & { officeChartKey?: string; officeChartPointer?: ChartAxisPointerRuntime | ChartSankeyPointerRuntime }, size: EChartsCanvasSize, theme: ChartEChartsTheme = {}): EChartsCanvasFrame {
    if (![size.width, size.height, size.devicePixelRatio].every(value => Number.isFinite(value) && value > 0)) {
      throw new RangeError('图表尺寸和像素比例必须为有效正数。')
    }
    const { officeChartKey, officeChartPointer, ...renderOption } = option
    const signature = JSON.stringify([officeChartKey, renderOption, size.width, size.height, size.devicePixelRatio, theme])
    let entry = this.charts.get(chartId)
    if (entry?.signature === signature) {
      // The retained Sankey scene owns view-only node positions. Preparing the
      // same authored option again must not replace it with the automatic seed.
      if (entry.pointer?.mode !== 'sankey' || officeChartPointer?.mode !== 'sankey') entry.pointer = officeChartPointer
      return { source: entry.canvas, sourceVersion: entry.sourceVersion }
    }
    const previousVersion = entry?.sourceVersion ?? 0
    if (entry && entry.devicePixelRatio !== size.devicePixelRatio) {
      entry.instance.dispose()
      entry.canvas.width = 1
      entry.canvas.height = 1
      this.charts.delete(chartId)
      entry = undefined
    }
    if (!entry) {
      const canvas = document.createElement('canvas')
      const instance = echarts.init(canvas, null, {
        renderer: 'canvas',
        width: size.width,
        height: size.height,
        devicePixelRatio: size.devicePixelRatio,
      })
      entry = {
        instance,
        canvas,
        signature: '',
        sourceVersion: previousVersion,
        devicePixelRatio: size.devicePixelRatio,
        width: size.width,
        height: size.height,
        pointerSignature: '',
      }
      this.charts.set(chartId, entry)
    }
    try {
      if (entry.width !== size.width || entry.height !== size.height) {
        entry.instance.resize({ width: size.width, height: size.height, silent: true, animation: { duration: 0 } })
      }
      entry.instance.setOption({ ...renderOption, animation: false }, { notMerge: true, lazyUpdate: false, silent: true })
      entry.instance.getZr().refreshImmediately()
      entry.signature = signature
      entry.sourceVersion += 1
      entry.width = size.width
      entry.height = size.height
      entry.pointer = officeChartPointer
      entry.pointerSignature = ''
      entry.contentDrag = undefined
    } catch (error) {
      entry.instance.dispose()
      entry.canvas.width = 1
      entry.canvas.height = 1
      this.charts.delete(chartId)
      throw error
    }
    return { source: entry.canvas, sourceVersion: entry.sourceVersion }
  }

  /** Update only the transient indicator. Content preparation and setOption signatures stay cached. */
  updatePointer(chartId: string, point?: ChartVector): EChartsCanvasFrame | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer
    if (!entry || !pointer) return undefined
    const finitePoint = point && Number.isFinite(point.x) && Number.isFinite(point.y) ? point : undefined
    if (pointer.mode === 'sankey') {
      const hit: ChartSankeyHit | undefined = entry.contentDrag ? { kind: 'node', index: entry.contentDrag.nodeIndex }
        : hitChartSankey(pointer.hitIndex, finitePoint)
      return this.updateSankeyGraphic(entry, pointer, hit)
    }
    const commands = pointer.mode === 'graphic' ? chartAxisPointerCommands(pointer.geometry, finitePoint) : []
    const visible = pointer.mode === 'graphic' ? commands.length > 0
      : Boolean(finitePoint && entry.instance.containPixel({ gridIndex: 0 }, [finitePoint.x, finitePoint.y]))
    const signature = !visible ? '' : pointer.mode === 'graphic' ? JSON.stringify(commands) : `${finitePoint!.x},${finitePoint!.y}`
    if (signature === entry.pointerSignature) return undefined
    if (pointer.mode === 'native') entry.instance.dispatchAction({ type: 'updateAxisPointer', currTrigger: visible ? 'mousemove' : 'leave',
      ...(visible ? { x: finitePoint!.x, y: finitePoint!.y } : {}) })
    else {
      const id = `${chartId}:axis-pointer`, element: NonNullable<GraphicComponentOption['elements']>[number] = visible
        ? { type: 'group', id, $action: 'replace', silent: true, z: 100_000, children: chartGraphics(commands, id) }
        : { id, $action: 'remove' }
      entry.instance.setOption({ graphic: { elements: [element] } }, { notMerge: false, lazyUpdate: false, silent: true })
    }
    entry.instance.getZr().refreshImmediately()
    entry.pointerSignature = signature
    entry.sourceVersion += 1
    return { source: entry.canvas, sourceVersion: entry.sourceVersion }
  }

  /** Editor-host coordinates are local CSS pixels, matching updatePointer. */
  getContentHit(chartId: string, point?: ChartVector): ChartSankeyHit | undefined {
    const pointer = this.charts.get(chartId)?.pointer
    return pointer?.mode === 'sankey' ? hitChartSankey(pointer.hitIndex, point) : undefined
  }

  beginContentDrag(chartId: string, point: ChartVector): EChartsCanvasContentDragStart | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer
    if (!entry || pointer?.mode !== 'sankey' || entry.contentDrag) return undefined
    const hit = hitChartSankey(pointer.hitIndex, point)
    if (hit?.kind !== 'node') return undefined
    const node = pointer.scene.layout.nodes[hit.index]!
    entry.contentDrag = { nodeIndex: hit.index, offset: { x: point.x - node.x, y: point.y - node.y }, originalScene: pointer.scene }
    const frame = this.updateSankeyGraphic(entry, pointer, hit)
    return { kind: 'sankey-node', nodeIndex: hit.index, ...(frame ? { frame } : {}) }
  }

  updateContentDrag(chartId: string, point: ChartVector): EChartsCanvasFrame | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer, drag = entry?.contentDrag
    if (!entry || pointer?.mode !== 'sankey' || !drag || !Number.isFinite(point.x) || !Number.isFinite(point.y)) return undefined
    const scene = moveChartSankeyScene(drag.originalScene, drag.nodeIndex, { x: point.x - drag.offset.x, y: point.y - drag.offset.y })
    const before = pointer.scene.layout.nodes[drag.nodeIndex]!, after = scene.layout.nodes[drag.nodeIndex]!
    if (before.x === after.x && before.y === after.y) return undefined
    pointer.scene = scene; pointer.hitIndex = createChartSankeyHitIndex(scene.layout)
    return this.updateSankeyGraphic(entry, pointer, { kind: 'node', index: drag.nodeIndex }, true)
  }

  /** Commit retains positions only in this surface owner. Cancel restores
   * the exact pre-drag scene; neither path changes chart data or content history. */
  endContentDrag(chartId: string, cancelled = false): EChartsCanvasFrame | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer, drag = entry?.contentDrag
    if (!entry || pointer?.mode !== 'sankey' || !drag) return undefined
    if (cancelled) { pointer.scene = drag.originalScene; pointer.hitIndex = createChartSankeyHitIndex(pointer.scene.layout) }
    entry.contentDrag = undefined
    return this.updateSankeyGraphic(entry, pointer, undefined, true)
  }

  private updateSankeyGraphic(entry: ChartEntry, pointer: ChartSankeyPointerRuntime, hit?: ChartSankeyHit, force = false): EChartsCanvasFrame | undefined {
    const signature = hit ? `${hit.kind}:${hit.index}` : ''
    if (!force && signature === entry.pointerSignature) return undefined
    const commands = chartSankeyFocusedCommands(pointer.scene, chartSankeyTrajectory(pointer.topology, hit), hit)
    const element: NonNullable<GraphicComponentOption['elements']>[number] = { type: 'group', id: pointer.groupId,
      $action: 'replace', silent: true, children: chartGraphics(commands, pointer.groupId, pointer.zOffset) }
    entry.instance.setOption({ graphic: { elements: [element] } }, { notMerge: false, lazyUpdate: false, silent: true })
    entry.instance.getZr().refreshImmediately()
    entry.pointerSignature = signature
    entry.sourceVersion += 1
    return { source: entry.canvas, sourceVersion: entry.sourceVersion }
  }

  invalidate(): void {
    for (const entry of this.charts.values()) entry.signature = ''
  }

  retain(chartIds: ReadonlySet<string>): void {
    for (const [chartId, entry] of this.charts) {
      if (chartIds.has(chartId)) continue
      entry.instance.dispose()
      entry.canvas.width = 1
      entry.canvas.height = 1
      this.charts.delete(chartId)
    }
  }

  dispose(): void {
    for (const entry of this.charts.values()) {
      entry.instance.dispose()
      entry.canvas.width = 1
      entry.canvas.height = 1
    }
    this.charts.clear()
  }
}
