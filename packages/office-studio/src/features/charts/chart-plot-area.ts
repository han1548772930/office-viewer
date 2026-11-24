import type { ChartPlotAreaOptions, ChartType } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import { resolveChartType } from './chart-catalog'

/** v1.6.0 chart-ui Yi/Wi: plot geometry belongs to these coordinate hosts. */
export function supportsChartPlotArea(type: ChartType): boolean {
  return ['column', 'bar', 'line', 'area', 'scatter', 'bubble', 'combination', 'pie', 'radar'].includes(resolveChartType(type).family)
}

export function validateChartPlotArea(value: ChartPlotAreaOptions | undefined): void {
  if (value === undefined) return
  if (!value || typeof value !== 'object' || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new TypeError('绘图区须为普通数据对象。')
  for (const key of ['left', 'top', 'width', 'height'] as const) {
    const number = value[key]
    if (typeof number !== 'number' || !Number.isFinite(number) || number < 0 || number > 1 || (key === 'width' || key === 'height') && number === 0) {
      throw new RangeError('绘图区 ' + key + ' 须为 0 到 1 的有限比例，宽高须大于 0。')
    }
  }
  if (value.left + value.width > 1 + 1e-12 || value.top + value.height > 1 + 1e-12) throw new RangeError('绘图区须完整位于图表范围内。')
}

export function chartManualPlotFrame(type: ChartType, frame: RenderRect, area: ChartPlotAreaOptions | undefined): RenderRect | undefined {
  if (!area || !supportsChartPlotArea(type)) return undefined
  validateChartPlotArea(area)
  return { x: frame.x + frame.width * area.left, y: frame.y + frame.height * area.top,
    width: frame.width * area.width, height: frame.height * area.height }
}

/** Convert one automatic frame to a stable authoring seed, independent of position. */
export function chartPlotAreaFromFrame(frame: RenderRect, plot: RenderRect): ChartPlotAreaOptions {
  if (!(frame.width > 0) || !(frame.height > 0) || !Number.isFinite(frame.width) || !Number.isFinite(frame.height)) throw new RangeError('图表尺寸须为正的有限数。')
  const minimum = 1e-6
  const left = Math.min(1 - minimum, Math.max(0, (plot.x - frame.x) / frame.width))
  const top = Math.min(1 - minimum, Math.max(0, (plot.y - frame.y) / frame.height))
  return { left, top, width: Math.max(minimum, Math.min(1 - left, plot.width / frame.width)),
    height: Math.max(minimum, Math.min(1 - top, plot.height / frame.height)) }
}
