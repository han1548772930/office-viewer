import * as echarts from 'echarts/core'
import { BarChart, BoxplotChart, CandlestickChart, FunnelChart, GaugeChart, GraphChart, HeatmapChart, LineChart, PieChart, RadarChart, SankeyChart, ScatterChart, SunburstChart, TreemapChart } from 'echarts/charts'
import { AxisPointerComponent, GraphicComponent, GridComponent, LegendComponent, LegendScrollComponent, RadarComponent, TitleComponent, TooltipComponent, VisualMapComponent } from 'echarts/components'
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
import { transformChartRelationScene } from './chart-relation-draw'
import { chartRelationContentTooltip, chartRelationFocusedCommands, chartRelationFocusSelection,
  createChartRelationHitIndex, hitChartRelation, type ChartRelationContentTooltip, type ChartRelationHit, type ChartRelationPointerRuntime } from './chart-relation-interaction'
import { reconcileChartLegendSelection, toggleChartLegendSelection } from './chart-legend'
import { chartNativeLegendHit, chartNativeLegendScrollIndex, chartNativeLegendSelection,
  type ChartCanvasLegendHit, type ChartCanvasLegendRuntime } from './echarts-legend'
import type { ChartEChartsPointerRuntime } from './echarts-option'
import type { ChartHierarchyContentTooltip } from './chart-content-tooltip'
import { ChartNativeHierarchyTooltipBinding } from './echarts-hierarchy-tooltip'
import { chartTitleGraphicLayouts, createChartTitleTargets, hitChartTitle, type ChartCanvasTitleTarget, type ChartMeasuredTitleText } from './chart-title-interaction'
import type { ChartTitleKind } from './chart-titles'
import { hitChartPlotArea, type ChartCanvasPlotAreaTarget } from './chart-plot-area-interaction'
import { ChartNativeElementIndex } from './echarts-element-interaction'
import type { ChartCanvasElement, ChartCanvasElementRuntime, ChartCanvasElementTarget } from './chart-element-interaction'

echarts.use([
  BarChart, BoxplotChart, CandlestickChart, FunnelChart, GaugeChart, GraphChart, HeatmapChart, LineChart, PieChart, RadarChart, SankeyChart, ScatterChart, SunburstChart, TreemapChart,
  AxisPointerComponent, GraphicComponent, GridComponent, LegendComponent, LegendScrollComponent, RadarComponent, TitleComponent, TooltipComponent, VisualMapComponent, CanvasRenderer,
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
  originalBaseScene?: ChartSankeyScene
}

interface ChartEntry {
  instance: ECharts
  canvas: HTMLCanvasElement
  signature: string
  sourceVersion: number
  devicePixelRatio: number
  width: number
  height: number
  pointer?: ChartAxisPointerRuntime | ChartSankeyPointerRuntime | ChartTreemapPointerRuntime | ChartSunburstPointerRuntime | ChartChordPointerRuntime | ChartRelationPointerRuntime
  pointerSignature: string
  legend?: ChartCanvasLegendRuntime
  legendSelection: Record<string, boolean>
  legendScrollDataIndex: number
  legendRecoverySignature?: string
  contentDrag?: SankeyContentDrag
  hierarchyTooltip?: ChartNativeHierarchyTooltipBinding
  titles: ChartCanvasTitleTarget[]
  plotArea?: ChartCanvasPlotAreaTarget
  elementRuntime?: ChartCanvasElementRuntime
  elementOption?: EChartsOption
  elements?: ChartNativeElementIndex
}

/** Owns the raster surfaces used by spreadsheet chart draw commands. */
export class EChartsCanvasRenderer {
  private readonly charts = new Map<string, ChartEntry>()

  render(chartId: string, option: EChartsOption & { officeChartKey?: string; officeChartPointer?: ChartEChartsPointerRuntime; officeChartLegend?: ChartCanvasLegendRuntime; officeChartPlotArea?: ChartCanvasPlotAreaTarget; officeChartElements?: ChartCanvasElementRuntime }, size: EChartsCanvasSize, theme: ChartEChartsTheme = {}): EChartsCanvasFrame {
    if (![size.width, size.height, size.devicePixelRatio].every(value => Number.isFinite(value) && value > 0)) {
      throw new RangeError('图表尺寸和像素比例必须为有效正数。')
    }
    const { officeChartKey, officeChartPointer: authoredPointer, officeChartLegend, officeChartPlotArea, officeChartElements, ...renderOption } = option
    const signature = JSON.stringify([officeChartKey, renderOption, officeChartElements, size.width, size.height, size.devicePixelRatio, theme])
    let entry = this.charts.get(chartId)
    const legendSelection = officeChartLegend ? reconcileChartLegendSelection(officeChartLegend.names, officeChartLegend.selectMode,
      entry?.legend?.selectMode === officeChartLegend.selectMode ? entry.legendSelection : {}) : {}
    const oldScrollName = entry?.legend?.names[entry.legendScrollDataIndex]
    const legendScrollDataIndex = oldScrollName !== undefined && officeChartLegend ? Math.max(0, officeChartLegend.names.indexOf(oldScrollName)) : 0
    const view = officeChartLegend?.rebuild && (Object.values(legendSelection).some(selected => selected === false) || entry?.legendRecoverySignature === signature)
      ? officeChartLegend.rebuild(legendSelection, entry && (entry.signature === signature || entry.legendRecoverySignature === signature) ? entry.pointer : undefined) : undefined
    const officeChartPointer = view ? view.pointer : officeChartLegend?.preparePointer?.(legendSelection) ?? authoredPointer
    if (entry?.signature === signature) {
      entry.legend = officeChartLegend
      entry.legendSelection = legendSelection
      entry.plotArea = view?.plotArea ?? officeChartPlotArea
      // The retained Sankey scene owns view-only node positions. Preparing the
      // same authored option again must not replace it with the automatic seed.
      if (entry.pointer?.mode === 'chord' && officeChartPointer?.mode === 'chord') {
        // Equal paint/geometry may still have new aggregate values or formats.
        // Keep the measured labels and current hover bitmap, but refresh the
        // tooltip facts/base scene so a proportional source change is visible.
        officeChartPointer.hitIndex.labels = entry.pointer.hitIndex.labels
        entry.pointer = officeChartPointer
      } else if (entry.pointer?.mode === 'relation' && officeChartPointer?.mode === 'relation') {
        officeChartPointer.viewport = entry.pointer.viewport
        officeChartPointer.scene = transformChartRelationScene(officeChartPointer.baseScene, officeChartPointer.viewport)
        officeChartPointer.hitIndex = createChartRelationHitIndex(officeChartPointer.scene, entry.pointer.hitIndex.labels)
        entry.pointer = officeChartPointer
      } else if (entry.pointer?.mode !== 'sankey' || officeChartPointer?.mode !== 'sankey') entry.pointer = officeChartPointer
      this.syncHierarchyTooltip(entry)
      return { source: entry.canvas, sourceVersion: entry.sourceVersion }
    }
    const previousVersion = entry?.sourceVersion ?? 0
    if (entry && entry.devicePixelRatio !== size.devicePixelRatio) {
      this.disposeChartEntry(entry)
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
        legendSelection: {},
        legendScrollDataIndex: 0,
        titles: [],
      }
      this.charts.set(chartId, entry)
    }
    try {
      if (entry.width !== size.width || entry.height !== size.height) {
        entry.instance.resize({ width: size.width, height: size.height, silent: true, animation: { duration: 0 } })
      }
      const paintedOption = officeChartLegend ? chartNativeLegendSelection(view?.option ?? renderOption, legendSelection, legendScrollDataIndex) : renderOption
      entry.instance.setOption({ ...paintedOption, animation: false }, { notMerge: true, lazyUpdate: false, silent: true })
      entry.instance.getZr().refreshImmediately()
      this.measureTitleTexts(entry, paintedOption)
      entry.signature = signature
      entry.sourceVersion += 1
      entry.width = size.width
      entry.height = size.height
      entry.pointer = officeChartPointer
      entry.legend = officeChartLegend
      entry.plotArea = view?.plotArea ?? officeChartPlotArea
      entry.elementRuntime = officeChartElements
      entry.elementOption = paintedOption
      entry.elements = officeChartElements ? new ChartNativeElementIndex(entry.instance, paintedOption, officeChartElements, size.width, size.height) : undefined
      entry.legendSelection = legendSelection
      entry.legendScrollDataIndex = legendScrollDataIndex
      entry.legendRecoverySignature = undefined
      if (officeChartPointer?.mode === 'chord') this.measureChordLabels(entry, officeChartPointer)
      if (officeChartPointer?.mode === 'relation') this.measureContentLabels(entry, officeChartPointer)
      this.syncHierarchyTooltip(entry)
      entry.hierarchyTooltip?.clear()
      entry.pointerSignature = ''
      entry.contentDrag = undefined
    } catch (error) {
      this.disposeChartEntry(entry)
      this.charts.delete(chartId)
      throw error
    }
    return { source: entry.canvas, sourceVersion: entry.sourceVersion }
  }

  /** Real Canvas item and pagination hits have priority over object dragging. */
  getTitleHit(chartId: string, point?: ChartVector): ChartCanvasTitleTarget | undefined {
    const entry = this.charts.get(chartId)
    return entry ? hitChartTitle(entry.titles, point) : undefined
  }

  getElementHit(chartId: string, point?: ChartVector, seriesOnly = false): ChartCanvasElementTarget | undefined {
    return this.charts.get(chartId)?.elements?.hit(point, seriesOnly)
  }
  getElementTarget(chartId: string, element: ChartCanvasElement): ChartCanvasElementTarget | undefined {
    return this.charts.get(chartId)?.elements?.target(element)
  }
  getElementTargets(chartId: string): ChartCanvasElementTarget[] { return this.charts.get(chartId)?.elements?.list() ?? [] }

  /** Borrowed surface: copy synchronously before any await, resize or disposal. */
  getImageFrame(chartId: string): EChartsCanvasFrame | undefined {
    const entry = this.charts.get(chartId)
    return entry ? { source: entry.canvas, sourceVersion: entry.sourceVersion } : undefined
  }

  getTitleTarget(chartId: string, kind: ChartTitleKind): ChartCanvasTitleTarget | undefined {
    return this.charts.get(chartId)?.titles.find(target => target.kind === kind)
  }

  getPlotAreaTarget(chartId: string): ChartCanvasPlotAreaTarget | undefined {
    return this.charts.get(chartId)?.plotArea
  }

  getPlotAreaHit(chartId: string, point: ChartVector, tolerance = 5, body = false): ChartCanvasPlotAreaTarget | undefined {
    const target = this.charts.get(chartId)?.plotArea
    return hitChartPlotArea(target, point, tolerance, body) ? target : undefined
  }

  getLegendHit(chartId: string, point: ChartVector): ChartCanvasLegendHit | undefined {
    const entry = this.charts.get(chartId)
    if (!entry?.legend || ![point.x, point.y].every(Number.isFinite) || point.x < 0 || point.y < 0 || point.x >= entry.width || point.y >= entry.height) return undefined
    return chartNativeLegendHit(entry.instance.getZr().findHover(point.x, point.y).topTarget, entry.legend.names)
  }

  /** Toggle/paginate the retained view. No author model, Worker request,
   * mutation, undo record or savepoint is produced. */
  activateLegend(chartId: string, point: ChartVector): EChartsCanvasFrame | undefined {
    const entry = this.charts.get(chartId), hit = this.getLegendHit(chartId, point)
    if (!entry?.legend || !hit || entry.contentDrag) return undefined
    const previousSelection = entry.legendSelection, previousScroll = entry.legendScrollDataIndex, previousPointer = entry.pointer, previousPlotArea = entry.plotArea
    let selected = previousSelection, scrollDataIndex = previousScroll
    try {
      const cleared = this.updatePointer(chartId)
      if (hit.kind === 'page') {
        const previous = chartNativeLegendScrollIndex(entry.instance.getOption())
        const zr = entry.instance.getZr()
        zr.handler.dispatch('click', { type: 'click', zrX: point.x, zrY: point.y })
        scrollDataIndex = chartNativeLegendScrollIndex(entry.instance.getOption())
        if (previous === scrollDataIndex) return cleared
      } else {
        if (entry.legend.selectMode === 'close') return cleared
        selected = toggleChartLegendSelection(entry.legend.names, entry.legend.selectMode, previousSelection, hit.name)
        if (JSON.stringify(selected) === JSON.stringify(previousSelection)) return cleared
        const view = entry.legend.rebuild?.(selected, entry.pointer)
        const nextPointer = view?.pointer ?? entry.legend.preparePointer?.(selected)
        entry.instance.dispatchAction({ type: entry.legend.selectMode === 'single' ? 'legendSelect' : 'legendToggleSelect', name: hit.name })
        if (view) {
          const previous = entry.pointer
          entry.instance.setOption(chartNativeLegendSelection(view.option, selected, previousScroll), { notMerge: true, lazyUpdate: false, silent: true })
          entry.pointer = view.pointer
          if (previous?.mode === 'relation' && entry.pointer?.mode === 'relation') {
            const pointer = entry.pointer
            pointer.viewport = previous.viewport
            pointer.scene = transformChartRelationScene(pointer.baseScene, pointer.viewport)
            pointer.hitIndex = createChartRelationHitIndex(pointer.scene)
            entry.instance.setOption({ graphic: { elements: [{ type: 'group', id: pointer.groupId, $action: 'replace', silent: true,
              children: chartGraphics(pointer.scene.commands, pointer.groupId, pointer.zOffset) }] } }, { notMerge: false, lazyUpdate: false, silent: true })
          }
          this.measureTitleTexts(entry, view.option)
          entry.elementOption = view.option
          entry.plotArea = view.plotArea
        } else if (entry.legend.preparePointer) entry.pointer = nextPointer
        if (entry.pointer?.mode === 'relation') this.measureContentLabels(entry, entry.pointer)
        if (entry.pointer?.mode === 'chord') this.measureChordLabels(entry, entry.pointer)
      }
      entry.instance.getZr().refreshImmediately()
      if (entry.elementRuntime && entry.elementOption) entry.elements = new ChartNativeElementIndex(entry.instance, entry.elementOption, entry.elementRuntime, entry.width, entry.height)
      entry.legendSelection = selected
      entry.legendScrollDataIndex = scrollDataIndex
      entry.pointerSignature = ''
      entry.sourceVersion++
      return { source: entry.canvas, sourceVersion: entry.sourceVersion }
    } catch (error) {
      entry.legendSelection = previousSelection
      entry.legendScrollDataIndex = previousScroll
      entry.pointer = previousPointer
      entry.plotArea = previousPlotArea
      entry.pointerSignature = ''
      // The native action may have painted partially. The next render must
      // replace that frame, while retaining view-only scene positions.
      entry.legendRecoverySignature = entry.signature || entry.legendRecoverySignature
      entry.signature = ''
      throw error
    }
  }

  getLegendView(chartId: string): { selected: Record<string, boolean>; scrollDataIndex: number } | undefined {
    const entry = this.charts.get(chartId)
    return entry?.legend ? { selected: { ...entry.legendSelection }, scrollDataIndex: entry.legendScrollDataIndex } : undefined
  }

  /** Update only the transient indicator. Content preparation and setOption signatures stay cached. */
  updatePointer(chartId: string, point?: ChartVector): EChartsCanvasFrame | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer
    if (!entry || !pointer) return undefined
    const finitePoint = point && Number.isFinite(point.x) && Number.isFinite(point.y) ? point : undefined
    if (pointer.mode === 'treemap' || pointer.mode === 'sunburst') return this.updateHierarchyPointer(entry, finitePoint)
    if (pointer.mode === 'chord') return this.updateChordGraphic(entry, pointer, hitChartChord(pointer.hitIndex, finitePoint))
    if (pointer.mode === 'relation') return this.updateRelationGraphic(entry, pointer, hitChartRelation(pointer.hitIndex, finitePoint))
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
   * coordinates for native ancestor emphasis and the real item event capture.
   * Tooltip content is displayed by the editor's Ant host; the locked native
   * frame listener applies pending emphasis before publishing this surface. */
  private updateHierarchyPointer(entry: ChartEntry, point?: ChartVector): EChartsCanvasFrame | undefined {
    const inside = point && point.x >= 0 && point.y >= 0 && point.x < entry.width && point.y < entry.height ? point : undefined
    const signature = inside ? `${inside.x},${inside.y}` : ''
    if (signature === entry.pointerSignature) return undefined
    const zr = entry.instance.getZr()
    entry.hierarchyTooltip?.clear()
    try {
      if (inside) zr.handler.dispatch('mousemove', { type: 'mousemove', zrX: inside.x, zrY: inside.y })
      else {
        zr.handler.dispatch('mousemove', { type: 'mousemove', zrX: -1, zrY: -1 })
        zr.handler.dispatch('mouseout', { type: 'mouseout', zrX: -1, zrY: -1, zrEventControl: 'only_globalout' })
      }
      zr.animation.trigger('frame', 0)
      zr.refreshImmediately()
    } catch (error) {
      entry.hierarchyTooltip?.clear()
      entry.pointerSignature = ''
      entry.signature = ''
      throw error
    }
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
  getContentTooltip(chartId: string, point?: ChartVector): ChartChordContentTooltip | ChartRelationContentTooltip | ChartHierarchyContentTooltip | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer
    if (entry && (pointer?.mode === 'treemap' || pointer?.mode === 'sunburst')) {
      if (!point || ![point.x, point.y].every(Number.isFinite) || point.x < 0 || point.y < 0
        || point.x >= entry.width || point.y >= entry.height || entry.pointerSignature !== `${point.x},${point.y}`) return undefined
      return entry.hierarchyTooltip?.resolve(pointer)
    }
    return pointer?.mode === 'chord' ? chartChordContentTooltip(pointer.scene, hitChartChord(pointer.hitIndex, point))
      : pointer?.mode === 'relation' ? chartRelationContentTooltip(pointer.scene, hitChartRelation(pointer.hitIndex, point)) : undefined
  }

  private syncHierarchyTooltip(entry: ChartEntry): void {
    if (entry.pointer?.mode === 'treemap' || entry.pointer?.mode === 'sunburst') {
      entry.hierarchyTooltip ??= new ChartNativeHierarchyTooltipBinding(entry.instance)
    } else {
      entry.hierarchyTooltip?.dispose()
      entry.hierarchyTooltip = undefined
    }
  }

  private disposeChartEntry(entry: ChartEntry): void {
    const failures: unknown[] = []
    try { entry.hierarchyTooltip?.dispose() }
    catch (error) { failures.push(error) }
    entry.hierarchyTooltip = undefined
    try { entry.instance.dispose() }
    catch (error) { failures.push(error) }
    entry.canvas.width = 1
    entry.canvas.height = 1
    if (failures.length) throw failures[0]
  }

  private measureChordLabels(entry: ChartEntry, pointer: ChartChordPointerRuntime): void {
    this.measureContentLabels(entry, pointer)
  }

  private measureTitleTexts(entry: ChartEntry, option: EChartsOption): void {
    const graphic = option.graphic
    const elements = Array.isArray(graphic) ? graphic.flatMap(item => item.elements ?? []) : graphic?.elements
    const layouts = chartTitleGraphicLayouts(elements), texts: ChartMeasuredTitleText[] = []
    for (const root of entry.instance.getZr().storage.getRoots()) root.traverse(element => {
      const layout = layouts.get(String(element.id))
      if (!layout || !(element instanceof echarts.graphic.Text)) return
      const bounds = element.getBoundingRect(), transform = element.getComputedTransform()
      if (![bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite) || bounds.width <= 0 || bounds.height <= 0) return
      const polygon = [{ x: bounds.x, y: bounds.y }, { x: bounds.x + bounds.width, y: bounds.y },
        { x: bounds.x + bounds.width, y: bounds.y + bounds.height }, { x: bounds.x, y: bounds.y + bounds.height }]
      if (transform) for (const point of polygon) {
        const x = point.x, y = point.y
        point.x = transform[0] * x + transform[2] * y + transform[4]
        point.y = transform[1] * x + transform[3] * y + transform[5]
      }
      texts.push({ layout, polygon, leading: { x: (transform?.[0] ?? 1) * bounds.x, y: (transform?.[1] ?? 0) * bounds.x } })
    })
    entry.titles = createChartTitleTargets(texts)
  }

  private measureContentLabels(entry: ChartEntry, pointer: ChartChordPointerRuntime | ChartRelationPointerRuntime): void {
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

  private updateRelationGraphic(entry: ChartEntry, pointer: ChartRelationPointerRuntime, hit?: ChartRelationHit, force = false): EChartsCanvasFrame | undefined {
    const selection = chartRelationFocusSelection(pointer.topology, pointer.scene.emphasisEnabled, hit)
    const signature = selection && hit ? `${hit.kind}:${hit.index}` : ''
    if (!force && signature === entry.pointerSignature) return undefined
    const commands = chartRelationFocusedCommands(pointer.scene, selection)
    const element: NonNullable<GraphicComponentOption['elements']>[number] = { type: 'group', id: pointer.groupId,
      $action: 'replace', silent: true, children: chartGraphics(commands, pointer.groupId, pointer.zOffset) }
    entry.instance.setOption({ graphic: { elements: [element] } }, { notMerge: false, lazyUpdate: false, silent: true })
    entry.instance.getZr().refreshImmediately()
    entry.pointerSignature = signature; entry.sourceVersion++
    return { source: entry.canvas, sourceVersion: entry.sourceVersion }
  }

  /** HOe's scale-only roam is handled inside the bitmap owner. Wheel scaling
   * changes no document command, prepared Rust buffer or calculation cache. */
  scaleContent(chartId: string, point: ChartVector, delta: number): EChartsCanvasFrame | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer
    if (!entry || pointer?.mode !== 'relation' || ![point.x, point.y, delta].every(Number.isFinite)) return undefined
    const previous = pointer.viewport, scale = Math.max(0.4, Math.min(2, previous.scale * Math.exp(Math.max(-1, Math.min(1, -delta * 0.0015)))))
    if (scale === previous.scale) return undefined
    const ratio = scale / previous.scale
    pointer.viewport = { scale, x: point.x - (point.x - previous.x) * ratio, y: point.y - (point.y - previous.y) * ratio }
    pointer.scene = transformChartRelationScene(pointer.baseScene, pointer.viewport)
    pointer.hitIndex = createChartRelationHitIndex(pointer.scene)
    const frame = this.updateRelationGraphic(entry, pointer, undefined, true)
    this.measureContentLabels(entry, pointer)
    return frame
  }

  beginContentDrag(chartId: string, point: ChartVector): EChartsCanvasContentDragStart | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer
    if (!entry || pointer?.mode !== 'sankey' || entry.contentDrag) return undefined
    const hit = hitChartSankey(pointer.hitIndex, point)
    if (hit?.kind !== 'node') return undefined
    const node = pointer.scene.layout.nodes[hit.index]!
    entry.contentDrag = { nodeIndex: hit.index, offset: { x: point.x - node.x, y: point.y - node.y }, originalScene: pointer.scene, originalBaseScene: pointer.baseScene }
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
    if (drag.originalBaseScene) pointer.baseScene = moveChartSankeyScene(drag.originalBaseScene, drag.nodeIndex, { x: point.x - drag.offset.x, y: point.y - drag.offset.y })
    return this.updateSankeyGraphic(entry, pointer, { kind: 'node', index: drag.nodeIndex }, true)
  }

  /** Commit retains positions only in this surface owner. Cancel restores
   * the exact pre-drag scene; neither path changes chart data or content history. */
  endContentDrag(chartId: string, cancelled = false): EChartsCanvasFrame | undefined {
    const entry = this.charts.get(chartId), pointer = entry?.pointer, drag = entry?.contentDrag
    if (!entry || pointer?.mode !== 'sankey' || !drag) return undefined
    if (cancelled) { pointer.scene = drag.originalScene; pointer.baseScene = drag.originalBaseScene; pointer.hitIndex = createChartSankeyHitIndex(pointer.scene.layout) }
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
      this.disposeChartEntry(entry)
      this.charts.delete(chartId)
    }
  }

  dispose(): void {
    for (const entry of this.charts.values()) {
      this.disposeChartEntry(entry)
    }
    this.charts.clear()
  }
}
