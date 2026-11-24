import type { ChartModel } from '../../domain/model'
import type { DrawCommand } from '../../rendering/contracts'
import type { ChartData, ChartLayout } from './chart-engine'
import type { ChartVector } from './chart-geometry'
import { formatChartNumber } from './chart-geometry'
import { chartInterpolate } from './chart-math'
import { createChartHeatmapLegend, type ChartHeatmapLegend } from './chart-heatmap-draw'
import { chartHeatmapCellRect, chartHeatmapExtent, chartHeatmapFieldNames, chartHeatmapPieceIndex, chartHeatmapValueRatio,
  chartHeatmapValueVisible, createChartHeatmapScale, hitChartHeatmapCell, type ChartHeatmapScale } from './chart-heatmap'

export interface ChartHeatmapPointerGeometry {
  data: ChartData
  chart: ChartModel
  layout: ChartLayout
  fieldNames: string[]
  scale: ChartHeatmapScale
  legend: ChartHeatmapLegend
  /** Packed source identities, one admission per visible value. The bounded
   * value bins avoid a source sort or a full source scan on every pointer. */
  bands: number[][]
  color: string
}
export type ChartHeatmapHit = { kind: 'cell'; seriesPosition: number; categoryIndex: number; value: number }
  | { kind: 'piece'; pieceIndex: number }
  | { kind: 'continuous'; ratio: number; value: number; lower: number; upper: number }

export function createChartHeatmapPointerGeometry(data: ChartData, chart: ChartModel, layout: ChartLayout,
  palette: readonly string[], color: string): ChartHeatmapPointerGeometry {
  const scale = createChartHeatmapScale(chartHeatmapExtent(data, chart), palette, chart.options?.heatmap?.visualMapType)
  const legend = createChartHeatmapLegend(chart, layout, scale, color), count = scale.type === 'piecewise' ? scale.pieces.length : 256
  const bands = Array.from({ length: count }, () => [] as number[]), rows = data.categories.length
  data.series.forEach((series, seriesPosition) => series.values.forEach((value, categoryIndex) => {
    if (categoryIndex >= rows || value === null || !chartHeatmapValueVisible(scale, value)) return
    const index = scale.type === 'piecewise' ? chartHeatmapPieceIndex(scale, value) : Math.min(count - 1, Math.floor(chartHeatmapValueRatio(value, scale.extent) * count))
    bands[index]!.push(seriesPosition * rows + categoryIndex)
  }))
  return { data, chart, layout, fieldNames: chartHeatmapFieldNames(data, chart), scale, legend, bands, color }
}

export function hitChartHeatmap(geometry: ChartHeatmapPointerGeometry, point?: ChartVector): ChartHeatmapHit | undefined {
  if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y) || !inside(point, geometry.layout.frame)) return undefined
  const { legend, scale } = geometry
  if (legend.visible) {
    if (legend.bar && inside(point, legend.bar)) {
      const length = legend.horizontal ? legend.bar.width : legend.bar.height
      const ratio = Math.max(0, Math.min(1, legend.horizontal ? (point.x - legend.bar.x) / length : 1 - (point.y - legend.bar.y) / length))
      // Fixed ContinuousView uses a 12px hover band by default, including
      // unbounded source values when that band crosses either colorbar end.
      const half = 6 / length, lowerRatio = ratio - half, upperRatio = ratio + half
      return { kind: 'continuous', ratio, value: chartInterpolate(scale.extent.min, scale.extent.max, ratio),
        lower: lowerRatio < 0 ? -Infinity : chartInterpolate(scale.extent.min, scale.extent.max, lowerRatio),
        upper: upperRatio > 1 ? Infinity : chartInterpolate(scale.extent.min, scale.extent.max, upperRatio) }
    }
    const piece = legend.pieces.find(piece => inside(point, piece.hit))
    if (piece) return { kind: 'piece', pieceIndex: piece.pieceIndex }
  }
  const cell = hitChartHeatmapCell(geometry.data, geometry.chart, geometry.layout.plot, point)
  if (!cell) return undefined
  const value = geometry.data.series[cell.seriesPosition]!.values[cell.categoryIndex]
  return value == null || !chartHeatmapValueVisible(scale, value) ? undefined : { kind: 'cell', ...cell, value }
}

/** Source identities are returned in deterministic band/source order.
 * Hover is view state and never writes min/max, a selected range or history. */
export function chartHeatmapFocusedCells(geometry: ChartHeatmapPointerGeometry, hit?: ChartHeatmapHit): number[] {
  if (!hit) return []
  const rows = geometry.data.categories.length
  if (hit.kind === 'cell') return [hit.seriesPosition * rows + hit.categoryIndex]
  if (hit.kind === 'piece') return geometry.bands[hit.pieceIndex]?.slice() ?? []
  const { scale, bands } = geometry
  const lower = hit.lower === -Infinity ? 0 : Math.min(bands.length - 1, Math.floor(chartHeatmapValueRatio(hit.lower, scale.extent) * bands.length))
  const upper = hit.upper === Infinity ? bands.length - 1 : Math.min(bands.length - 1, Math.floor(chartHeatmapValueRatio(hit.upper, scale.extent) * bands.length))
  const cells: number[] = []
  for (let index = lower; index <= upper; index++) for (const cell of bands[index]!) {
    const value = geometry.data.series[Math.floor(cell / rows)]!.values[cell % rows]!
    if (value !== null && value >= hit.lower && value <= hit.upper) cells.push(cell)
  }
  return cells
}

export function chartHeatmapPointerCommands(geometry: ChartHeatmapPointerGeometry, point?: ChartVector): DrawCommand[] {
  const hit = hitChartHeatmap(geometry, point)
  if (!hit) return []
  const { layout, scale, legend, color, data, chart } = geometry, commands: DrawCommand[] = [], rows = data.categories.length
  for (const cell of chartHeatmapFocusedCells(geometry, hit)) {
    const rect = chartHeatmapCellRect(data, chart, layout.plot, Math.floor(cell / rows), cell % rows)!
    commands.push({ kind: 'rect', ...rect, fill: 'transparent', stroke: color, strokeWidth: 1, clip: layout.plot })
  }
  if (!legend.visible) return commands
  if (scale.type === 'piecewise') {
    const pieceIndex = hit.kind === 'piece' ? hit.pieceIndex : hit.kind === 'cell' ? chartHeatmapPieceIndex(scale, hit.value) : -1
    const piece = legend.pieces.find(piece => piece.pieceIndex === pieceIndex)
    if (piece) commands.push({ kind: 'rect', x: piece.rect.x - 2, y: piece.rect.y - 2, width: piece.rect.width + 4, height: piece.rect.height + 4,
      fill: 'transparent', stroke: color, strokeWidth: 1, clip: layout.frame })
  } else if (legend.bar && hit.kind !== 'piece') {
    const ratio = hit.kind === 'cell' ? chartHeatmapValueRatio(hit.value, scale.extent) : hit.ratio, bar = legend.bar
    const x = legend.horizontal ? bar.x + ratio * bar.width : bar.x + bar.width
    const y = legend.horizontal ? bar.y + bar.height : bar.y + (1 - ratio) * bar.height
    commands.push({ kind: 'polygon', points: legend.horizontal ? [{ x, y }, { x: x - 3, y: y + 4 }, { x: x + 3, y: y + 4 }]
      : [{ x, y }, { x: x + 4, y: y - 3 }, { x: x + 4, y: y + 3 }], fill: color, clip: layout.frame })
    const text = formatChartNumber(hit.value), size = legend.font.fontSize ?? 11
    commands.push({ kind: 'text', x: legend.horizontal ? x : x + 8, y: legend.horizontal ? bar.y - size / 2 - 4 : y,
      text, font: `${legend.font.italic ? 'italic ' : ''}${legend.font.bold ? 'bold ' : ''}${size}px ${legend.font.family ?? 'sans-serif'}`, color: legend.color,
      align: legend.horizontal ? 'center' : 'left', baseline: 'middle', maxWidth: legend.horizontal ? Math.max(1, bar.width) : Math.max(1, layout.frame.x + layout.frame.width - x - 10), clip: layout.frame })
  }
  return commands
}

function inside(point: ChartVector, rect: { x: number; y: number; width: number; height: number }): boolean {
  return point.x >= rect.x && point.x <= rect.x + rect.width && point.y >= rect.y && point.y <= rect.y + rect.height
}
