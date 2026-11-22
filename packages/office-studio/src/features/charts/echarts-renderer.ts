import * as echarts from 'echarts/core'
import { BarChart, LineChart, PieChart, RadarChart, ScatterChart } from 'echarts/charts'
import { GridComponent, LegendComponent, RadarComponent, TitleComponent, TooltipComponent } from 'echarts/components'
import { CanvasRenderer } from 'echarts/renderers'
import type { ECharts, EChartsOption } from 'echarts'
import type { ChartEChartsTheme } from './echarts-option'

echarts.use([
  BarChart, LineChart, PieChart, RadarChart, ScatterChart,
  GridComponent, LegendComponent, RadarComponent, TitleComponent, TooltipComponent, CanvasRenderer,
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
}

/** Owns the raster surfaces used by spreadsheet chart draw commands. */
export class EChartsCanvasRenderer {
  private readonly charts = new Map<string, ChartEntry>()

  render(chartId: string, option: EChartsOption, size: EChartsCanvasSize, theme: ChartEChartsTheme = {}): EChartsCanvasFrame {
    if (![size.width, size.height, size.devicePixelRatio].every(value => Number.isFinite(value) && value > 0)) {
      throw new RangeError('图表尺寸和像素比例必须为有效正数。')
    }
    const signature = JSON.stringify([option, size.width, size.height, size.devicePixelRatio, theme])
    let entry = this.charts.get(chartId)
    if (entry?.signature === signature) return { source: entry.canvas, sourceVersion: entry.sourceVersion }
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
      }
      this.charts.set(chartId, entry)
    }
    try {
      if (entry.width !== size.width || entry.height !== size.height) {
        entry.instance.resize({ width: size.width, height: size.height, silent: true, animation: { duration: 0 } })
      }
      entry.instance.setOption({ ...option, animation: false }, { notMerge: true, lazyUpdate: false, silent: true })
      entry.signature = signature
      entry.sourceVersion += 1
      entry.width = size.width
      entry.height = size.height
    } catch (error) {
      entry.instance.dispose()
      entry.canvas.width = 1
      entry.canvas.height = 1
      this.charts.delete(chartId)
      throw error
    }
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
