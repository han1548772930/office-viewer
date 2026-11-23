import * as echarts from 'echarts/core'
import { BarChart, BoxplotChart, CandlestickChart, FunnelChart, GaugeChart, GraphChart, HeatmapChart, LineChart, PieChart, RadarChart, SankeyChart, ScatterChart, SunburstChart, TreemapChart } from 'echarts/charts'
import { AxisPointerComponent, GraphicComponent, GridComponent, LegendComponent, RadarComponent, TitleComponent, TooltipComponent, VisualMapComponent } from 'echarts/components'
import { CanvasRenderer } from 'echarts/renderers'
import type { ECharts, EChartsOption } from 'echarts'
import type { GraphicComponentOption } from 'echarts/components'
import { chartGraphics, type ChartEChartsTheme } from './echarts-option'
import { chartAxisPointerCommands, type ChartAxisPointerRuntime } from './chart-axis-pointer'
import type { ChartVector } from './chart-geometry'

echarts.use([
  BarChart, BoxplotChart, CandlestickChart, FunnelChart, GaugeChart, GraphChart, HeatmapChart, LineChart, PieChart, RadarChart, SankeyChart, ScatterChart, SunburstChart, TreemapChart,
  AxisPointerComponent, GraphicComponent, GridComponent, LegendComponent, RadarComponent, TitleComponent, TooltipComponent, VisualMapComponent, CanvasRenderer,
])

export interface EChartsCanvasSize {
  width: number
  height: number
  devicePixelRatio: number
}

export interface EChartsCanvasFrame {
  source: HTMLCanvasElement
  sourceVersion: number
}

interface ChartEntry {
  instance: ECharts
  canvas: HTMLCanvasElement
  signature: string
  sourceVersion: number
  devicePixelRatio: number
  width: number
  height: number
  pointer?: ChartAxisPointerRuntime
  pointerSignature: string
}

/** Owns the raster surfaces used by spreadsheet chart draw commands. */
export class EChartsCanvasRenderer {
  private readonly charts = new Map<string, ChartEntry>()

  render(chartId: string, option: EChartsOption & { officeChartKey?: string; officeChartPointer?: ChartAxisPointerRuntime }, size: EChartsCanvasSize, theme: ChartEChartsTheme = {}): EChartsCanvasFrame {
    if (![size.width, size.height, size.devicePixelRatio].every(value => Number.isFinite(value) && value > 0)) {
      throw new RangeError('图表尺寸和像素比例必须为有效正数。')
    }
    const { officeChartKey, officeChartPointer, ...renderOption } = option
    const signature = JSON.stringify([officeChartKey, renderOption, size.width, size.height, size.devicePixelRatio, theme])
    let entry = this.charts.get(chartId)
    if (entry?.signature === signature) {
      entry.pointer = officeChartPointer
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
