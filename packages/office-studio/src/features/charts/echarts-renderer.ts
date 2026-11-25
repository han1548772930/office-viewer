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
import type { ChartTreemapPointerRuntime } from './echarts-treemap-option'
import { registerChartSunburstLayout } from './echarts-sunburst-layout'
import type { ChartSunburstPointerRuntime } from './echarts-sunburst-option'
import { chartChordContentTooltip, chartChordFocusedCommands, chartChordFocusSelection, hitChartChord,
  type ChartChordContentTooltip, type ChartChordHit, type ChartChordPointerRuntime } from './chart-chord-interaction'
import type { TextCommand } from '../../rendering/contracts'

echarts.use([
  BarChart, BoxplotChart, CandlestickChart, FunnelChart, GaugeChart, GraphChart, HeatmapChart, LineChart, PieChart, RadarChart, SankeyChart, ScatterChart, SunburstChart, TreemapChart,
  AxisPointerComponent, GraphicComponent, GridComponent, LegendComponent, RadarComponent, TitleComponent, TooltipComponent, VisualMapComponent, CanvasRenderer,
])
registerChartTreemapLayout()
registerChartSunburstLayout()

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
  pointer?: ChartAxisPointerRuntime | ChartSankeyPointerRuntime | ChartTreemapPointerRuntime | ChartSunburstPointerRuntime | ChartChordPointerRuntime
  pointerSignature: string
  contentDrag?: SankeyContentDrag
}

/** Owns the raster surfaces used by spreadsheet chart draw commands. */
export class EChartsCanvasRenderer {
  private readonly charts = new Map<string, ChartEntry>()

  render(chartId: string, option: EChartsOption & { officeChartKey?: string; officeChartPointer?: ChartAxisPointerRuntime | ChartSankeyPointerRuntime | ChartTreemapPointerRuntime | ChartSunburstPointerRuntime | ChartChordPointerRuntime }, size: EChartsCanvasSize, theme: ChartEChartsTheme = {}): EChartsCanvasFrame {
    if (![size.width, size.height, size.devicePixelRatio].every(value => Number.isFinite(value) && value > 0)) {
      throw new RangeError('图表尺寸和像素比例必须为有效正数。')
    }
    const { officeChartKey, officeChartPointer, ...renderOption } = option
    const signature = JSON.stringify([officeChartKey, renderOption, size.width, size.height, size.devicePixelRatio, theme])
    let entry = this.charts.get(chartId)
    if (entry?.signature === signature) {
      // The retained Sankey scene owns view-only node positions. Preparing the
      // same authored option again must not replace it with the automatic seed.
      if (entry.pointer?.mode === 'chord' && officeChartPointer?.mode === 'chord') {
        // Equal paint/geometry may still have new aggregate values or formats.
        // Keep the measured labels and current hover bitmap, but refresh the
        // tooltip facts/base scene so a proportional source change is visible.
        officeChartPointer.hitIndex.labels = entry.pointer.hitIndex.labels
        entry.pointer = officeChartPointer
      } else if (entry.pointer?.mode !== 'sankey' || officeChartPointer?.mode !== 'sankey') entry.pointer = officeChartPointer
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
      if (officeChartPointer?.mode === 'chord') this.measureChordLabels(entry, officeChartPointer)
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
    if (pointer.mode === 'treemap' || pointer.mode === 'sunburst') return this.updateHierarchyPointer(entry, finitePoint)
    if (pointer.mode === 'chord') return this.updateChordGraphic(entry, pointer, hitChartChord(pointer.hitIndex, finitePoint))
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

  /** The bitmap lives outside the editor's event target. Forward local pointer
   * coordinates so native ancestor emphasis and rich-text tooltip remain live.
   * The locked ECharts frame listener applies pending hover states before the
   * raster owner publishes this surface version; no source task is rerun. */
  private updateHierarchyPointer(entry: ChartEntry, point?: ChartVector): EChartsCanvasFrame | undefined {
    const inside = point && point.x >= 0 && point.y >= 0 && point.x < entry.width && point.y < entry.height ? point : undefined
    const signature = inside ? `${inside.x},${inside.y}` : ''
    if (signature === entry.pointerSignature) return undefined
    const zr = entry.instance.getZr()
    if (inside) {
      zr.handler.dispatch('mousemove', { type: 'mousemove', zrX: inside.x, zrY: inside.y })
      entry.instance.dispatchAction({ type: 'showTip', x: inside.x, y: inside.y })
    } else {
      zr.handler.dispatch('mousemove', { type: 'mousemove', zrX: -1, zrY: -1 })
      zr.handler.dispatch('mouseout', { type: 'mouseout', zrX: -1, zrY: -1, zrEventControl: 'only_globalout' })
      entry.instance.dispatchAction({ type: 'hideTip' })
    }
    zr.animation.trigger('frame', 0)
    zr.refreshImmediately()
    entry.pointerSignature = signature; entry.sourceVersion++
    return { source: entry.canvas, sourceVersion: entry.sourceVersion }
  }

  /** Editor-host coordinates are local CSS pixels, matching updatePointer. */
  getContentHit(chartId: string, point?: ChartVector): ChartSankeyHit | undefined {
    const pointer = this.charts.get(chartId)?.pointer
    return pointer?.mode === 'sankey' ? hitChartSankey(pointer.hitIndex, point) : undefined
  }

  /** UI tooltips use the shared antdv-next host. Reading this result neither
   * changes paint nor dispatches a content command/history operation. */
  getContentTooltip(chartId: string, point?: ChartVector): ChartChordContentTooltip | undefined {
    const pointer = this.charts.get(chartId)?.pointer
    return pointer?.mode === 'chord' ? chartChordContentTooltip(pointer.scene, hitChartChord(pointer.hitIndex, point)) : undefined
  }

  private measureChordLabels(entry: ChartEntry, pointer: ChartChordPointerRuntime): void {
    const labels = new Map<string, { index: number; command: TextCommand }>()
    let commandIndex = 0
    for (const mark of pointer.scene.marks) for (const command of mark.commands) {
      if (command.kind === 'text' && mark.kind === 'node') labels.set(`${pointer.groupId}-${commandIndex}`, { index: mark.index, command })
      commandIndex++
    }
    pointer.hitIndex.labels = []
    // ZRender owns the real font/truncation bounds. Text is not reconstructed
    // from an estimated character count or a DOM-backed source model.
    for (const root of entry.instance.getZr().storage.getRoots()) root.traverse(element => {
      const label = labels.get(String(element.id))
      if (!label || !(element instanceof echarts.graphic.Text)) return
      const bounds = element.getBoundingRect(), transform = element.getComputedTransform()
      if (![bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite) || bounds.width <= 0 || bounds.height <= 0) return
      const polygon = [{ x: bounds.x, y: bounds.y }, { x: bounds.x + bounds.width, y: bounds.y },
        { x: bounds.x + bounds.width, y: bounds.y + bounds.height }, { x: bounds.x, y: bounds.y + bounds.height }]
      if (transform) for (const point of polygon) {
        const x = point.x, y = point.y
        point.x = transform[0] * x + transform[2] * y + transform[4]
        point.y = transform[1] * x + transform[3] * y + transform[5]
      }
      if (polygon.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))) {
        pointer.hitIndex.labels.push({ index: label.index, polygon, ...(label.command.clip ? { clip: { ...label.command.clip } } : {}) })
      }
    })
  }

  private updateChordGraphic(entry: ChartEntry, pointer: ChartChordPointerRuntime, hit?: ChartChordHit): EChartsCanvasFrame | undefined {
    const selection = chartChordFocusSelection(pointer.topology, pointer.scene.focus, hit)
    const signature = selection && hit ? `${hit.kind}:${hit.index}` : ''
    if (signature === entry.pointerSignature) return undefined
    const commands = chartChordFocusedCommands(pointer.scene, selection)
    const element: NonNullable<GraphicComponentOption['elements']>[number] = { type: 'group', id: pointer.groupId,
      $action: 'replace', silent: true, children: chartGraphics(commands, pointer.groupId, pointer.zOffset) }
    entry.instance.setOption({ graphic: { elements: [element] } }, { notMerge: false, lazyUpdate: false, silent: true })
    entry.instance.getZr().refreshImmediately()
    entry.pointerSignature = signature; entry.sourceVersion++
    return { source: entry.canvas, sourceVersion: entry.sourceVersion }
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
