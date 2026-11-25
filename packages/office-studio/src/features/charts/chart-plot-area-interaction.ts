import type { ChartPlotAreaOptions, ChartType } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartVector } from './chart-geometry'
import { chartPlotAreaFromFrame, supportsChartPlotArea, validateChartPlotArea } from './chart-plot-area'

export type ChartPlotAreaCoordinate = keyof ChartPlotAreaOptions
export type ChartPlotAreaResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'
export const CHART_PLOT_AREA_HANDLES: readonly { handle: ChartPlotAreaResizeHandle; label: string; cursor: string }[] = [
  { handle: 'nw', label: '调整绘图区左上角', cursor: 'nwse-resize' },
  { handle: 'n', label: '调整绘图区顶部', cursor: 'ns-resize' },
  { handle: 'ne', label: '调整绘图区右上角', cursor: 'nesw-resize' },
  { handle: 'e', label: '调整绘图区右侧', cursor: 'ew-resize' },
  { handle: 'se', label: '调整绘图区右下角', cursor: 'nwse-resize' },
  { handle: 's', label: '调整绘图区底部', cursor: 'ns-resize' },
  { handle: 'sw', label: '调整绘图区左下角', cursor: 'nesw-resize' },
  { handle: 'w', label: '调整绘图区左侧', cursor: 'ew-resize' },
]

/** Surface-only geometry from the same layout that paints the plot. It is
 * never inferred from the chart frame, series marks or displayed labels. */
export interface ChartCanvasPlotAreaTarget {
  frame: RenderRect
  plot: RenderRect
  bounds?: RenderRect
  area: ChartPlotAreaOptions
  automatic: ChartPlotAreaOptions
  manual: boolean
}

export function createChartPlotAreaTarget(type: ChartType, frame: RenderRect, plot: RenderRect, automaticPlot: RenderRect,
  manual: boolean): ChartCanvasPlotAreaTarget | undefined {
  if (!supportsChartPlotArea(type)) return undefined
  for (const rect of [frame, plot, automaticPlot]) {
    if (![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite) || rect.width <= 0 || rect.height <= 0) throw new RangeError('绘图区交互需要有效的绘制矩形。')
  }
  const x = Math.max(frame.x, plot.x), y = Math.max(frame.y, plot.y)
  const right = Math.min(frame.x + frame.width, plot.x + plot.width), bottom = Math.min(frame.y + frame.height, plot.y + plot.height)
  const area = manual ? { left: (plot.x - frame.x) / frame.width, top: (plot.y - frame.y) / frame.height,
    width: plot.width / frame.width, height: plot.height / frame.height } : chartPlotAreaFromFrame(frame, plot)
  validateChartPlotArea(area)
  return { frame: { ...frame }, plot: { ...plot }, ...(right > x && bottom > y ? { bounds: { x, y, width: right - x, height: bottom - y } } : {}),
    area, automatic: chartPlotAreaFromFrame(frame, automaticPlot), manual }
}

/** An unselected plot accepts its painted edge. Its body becomes draggable
 * after explicit selection, preserving normal series/axis-pointer input. */
export function hitChartPlotArea(target: ChartCanvasPlotAreaTarget | undefined, point: ChartVector | undefined, tolerance = 5, body = false): boolean {
  if (!target?.bounds || !point || ![point.x, point.y, tolerance].every(Number.isFinite) || tolerance < 0) return false
  const { frame, bounds } = target, right = bounds.x + bounds.width, bottom = bounds.y + bounds.height
  if (point.x < frame.x || point.y < frame.y || point.x > frame.x + frame.width || point.y > frame.y + frame.height) return false
  if (body && point.x >= bounds.x && point.x <= right && point.y >= bounds.y && point.y <= bottom) return true
  return point.x >= bounds.x - tolerance && point.x <= right + tolerance && point.y >= bounds.y - tolerance && point.y <= bottom + tolerance
    && Math.min(Math.abs(point.x - bounds.x), Math.abs(point.x - right), Math.abs(point.y - bounds.y), Math.abs(point.y - bottom)) <= tolerance
}

export function chartPlotAreaHandlePoints(target: ChartCanvasPlotAreaTarget): { handle: ChartPlotAreaResizeHandle; x: number; y: number }[] {
  const bounds = target.bounds
  if (!bounds) return []
  const left = bounds.x, right = bounds.x + bounds.width, top = bounds.y, bottom = bounds.y + bounds.height
  return CHART_PLOT_AREA_HANDLES.map(({ handle }) => ({ handle,
    x: handle.includes('w') ? left : handle.includes('e') ? right : (left + right) / 2,
    y: handle.includes('n') ? top : handle.includes('s') ? bottom : (top + bottom) / 2 }))
}

export function chartPlotAreaOutlineCommands(target: ChartCanvasPlotAreaTarget, position: ChartVector, color: string, width = 1): DrawCommand[] {
  const bounds = target.bounds
  if (!bounds) return []
  const points = [{ x: bounds.x, y: bounds.y }, { x: bounds.x + bounds.width, y: bounds.y },
    { x: bounds.x + bounds.width, y: bounds.y + bounds.height }, { x: bounds.x, y: bounds.y + bounds.height }]
  const clip = { ...target.frame, x: target.frame.x + position.x, y: target.frame.y + position.y }
  return points.map((point, index) => {
    const next = points[(index + 1) % points.length]!
    return { kind: 'line', x1: position.x + point.x, y1: position.y + point.y,
      x2: position.x + next.x, y2: position.y + next.y, color, width, clip }
  })
}

/** Resizing fixes the opposite edge; moving keeps both dimensions. A tiny
 * existing rectangle never snaps larger merely because a drag begins. */
export function chartPlotAreaDragLayout(target: ChartCanvasPlotAreaTarget, size: { width: number; height: number }, delta: ChartVector,
  handle?: ChartPlotAreaResizeHandle): ChartPlotAreaOptions {
  if (![size.width, size.height, delta.x, delta.y].every(Number.isFinite) || size.width <= 0 || size.height <= 0) throw new RangeError('绘图区拖动需要有效尺寸和坐标。')
  validateChartPlotArea(target.area)
  const area = target.area, dx = delta.x / size.width, dy = delta.y / size.height
  const clamp = (value: number, minimum: number, maximum: number) => Math.min(maximum, Math.max(minimum, value))
  if (!handle) return { ...area, left: clamp(area.left + dx, 0, 1 - area.width), top: clamp(area.top + dy, 0, 1 - area.height) }
  const minWidth = Math.min(area.width, 8 / size.width), minHeight = Math.min(area.height, 8 / size.height)
  let left = area.left, right = area.left + area.width, top = area.top, bottom = area.top + area.height
  if (handle.includes('w')) left = clamp(left + dx, 0, right - minWidth)
  if (handle.includes('e')) right = clamp(right + dx, left + minWidth, 1)
  if (handle.includes('n')) top = clamp(top + dy, 0, bottom - minHeight)
  if (handle.includes('s')) bottom = clamp(bottom + dy, top + minHeight, 1)
  return { left, top, width: right - left, height: bottom - top }
}

export function chartPlotAreasEqual(left: ChartPlotAreaOptions | undefined, right: ChartPlotAreaOptions | undefined): boolean {
  return left === right || !!left && !!right && (['left', 'top', 'width', 'height'] as const).every(key => Math.abs(left[key] - right[key]) <= 1e-10)
}
