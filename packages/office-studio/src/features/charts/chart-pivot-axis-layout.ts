import type { CellScalar, ChartModel } from '../../domain/model.ts'
import type { DrawCommand, RenderRect, TextCommand } from '../../rendering/contracts.ts'
import type { ChartData } from './chart-engine.ts'
import { resolveChartType } from './chart-catalog.ts'
import { chartFont } from './chart-geometry.ts'

export interface PivotChartCategoryLabelInfo {
  key: string
  fieldId: string
  fieldName: string
  item: string
  level: number
  startIndex: number
  endIndex: number
  path: CellScalar[]
  text: string
  depth: number
}
export interface PivotChartCategoryTextCommand extends TextCommand { pivotChartAxisLabel: PivotChartCategoryLabelInfo }
export interface PivotChartCategoryLabel extends PivotChartCategoryLabelInfo { command: PivotChartCategoryTextCommand }
export interface PivotChartCategoryAxis {
  depth: number
  horizontal: boolean
  labels: PivotChartCategoryLabel[]
}

/** Category metadata must describe this exact projection, including synthetic value rows. */
export function pivotChartCategoryAxisDepth(data: ChartData, chart: Pick<ChartModel, 'type' | 'categoryField' | 'options' | 'xAxis'>): number {
  const axis = data.pivotAxis, descriptor = resolveChartType(chart.type)
  if (!axis || axis.rowFields.length < 2 || axis.recordRows.length !== data.categories.length || !data.categories.length
    || descriptor.threeD || !['column', 'bar', 'line', 'area', 'combination', 'waterfall'].includes(descriptor.family)
    || chart.options?.mapping?.useDateAxis || chart.xAxis.label?.visible === false) return 0
  const categories = chart.options?.mapping?.categoryIndexes
  if (categories && (categories.length !== data.sourceCategoryFields?.length
    || categories.some((field, index) => field !== data.sourceCategoryFields![index]))) return 0
  if (!categories && chart.categoryField !== 0 && chart.categoryField !== data.sourceCategoryFields?.[0]) return 0
  return axis.rowFields.length
}

export function pivotChartCategoryAxisBand(chart: Pick<ChartModel, 'options' | 'xAxis'>): { lineHeight: number; columnWidth: number } {
  const fontSize = chart.xAxis.label?.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 10
  return { lineHeight: Math.max(18, fontSize + 8), columnWidth: Math.max(48, fontSize * 6) }
}

/** Contiguous typed prefixes are authoritative. Equal displayed text cannot merge different members. */
export function pivotChartCategoryAxis(data: ChartData, chart: ChartModel, frame: RenderRect, plot: RenderRect,
  color: string): PivotChartCategoryAxis | undefined {
  const depth = pivotChartCategoryAxisDepth(data, chart)
  if (!depth) return undefined
  const axis = data.pivotAxis!, horizontal = resolveChartType(chart.type).family === 'bar'
  const { lineHeight, columnWidth } = pivotChartCategoryAxisBand(chart), labels: PivotChartCategoryLabel[] = []
  const count = axis.recordRows.length, font = chartFont({ ...chart.options?.appearance?.font, ...chart.xAxis.label }, 10)
  const labelColor = chart.xAxis.label?.color ?? color
  for (let level = 0; level < depth; level++) {
    let previousCoordinate = -Infinity
    for (let index = 0; index < count;) {
      const row = axis.recordRows[index]!
      if (level >= row.labels.length) { index++; continue }
      const path = row.labels.slice(0, level + 1), identity = JSON.stringify(path), startIndex = index
      while (++index < count && level < axis.recordRows[index]!.labels.length
        && JSON.stringify(axis.recordRows[index]!.labels.slice(0, level + 1)) === identity) { /* one visible prefix span */ }
      const endIndex = index - 1, extent = (horizontal ? plot.height : plot.width) * (index - startIndex) / count
      const center = (startIndex + index) / 2 / count, ratio = chart.xAxis.reverse ? 1 - center : center
      const coordinate = horizontal ? plot.y + ratio * plot.height : plot.x + ratio * plot.width
      // Native axes also skip dense labels. Only painted text receives a hit target.
      const orderedCoordinate = chart.xAxis.reverse ? -coordinate : coordinate
      if (extent < 8 || orderedCoordinate - previousCoordinate < (horizontal ? lineHeight : 14)) continue
      previousCoordinate = orderedCoordinate
      const field = axis.rowFields[level]!, item = String(path[level] ?? ''), text = item || '（空白）'
      const info: PivotChartCategoryLabelInfo = { key: JSON.stringify([field.id, path, startIndex]), fieldId: field.id, fieldName: field.name,
        item, level, startIndex, endIndex, path, text, depth }
      const command: PivotChartCategoryTextCommand = { kind: 'text', text, font, color: labelColor, baseline: 'middle',
        x: horizontal ? plot.x - 8 - (depth - 1 - level) * columnWidth : coordinate,
        y: horizontal ? coordinate : plot.y + plot.height + 14 + (depth - 1 - level) * lineHeight,
        align: horizontal ? 'right' : 'center', maxWidth: horizontal ? columnWidth - 8 : Math.max(1, extent - 4),
        rotation: chart.xAxis.label?.rotate, clip: { ...frame }, pivotChartAxisLabel: info }
      labels.push({ ...info, command })
    }
  }
  return { depth, horizontal, labels }
}

export function readPivotChartAxisCommand(command: DrawCommand): PivotChartCategoryLabelInfo | undefined {
  return command.kind === 'text' && 'pivotChartAxisLabel' in command ? readPivotChartAxisLabelInfo(command.pivotChartAxisLabel) : undefined
}

export function readPivotChartAxisLabelInfo(value: unknown): PivotChartCategoryLabelInfo | undefined {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined
  const item = value as Record<string, unknown>
  if (typeof item.key !== 'string' || typeof item.fieldId !== 'string' || !item.fieldId || typeof item.fieldName !== 'string'
    || typeof item.item !== 'string' || typeof item.text !== 'string' || !Number.isSafeInteger(item.level) || Number(item.level) < 0
    || !Number.isSafeInteger(item.depth) || Number(item.depth) <= Number(item.level) || !Number.isSafeInteger(item.startIndex) || Number(item.startIndex) < 0
    || !Number.isSafeInteger(item.endIndex) || Number(item.endIndex) < Number(item.startIndex) || !Array.isArray(item.path)
    || item.path.length !== Number(item.level) + 1 || item.path.some(value => value !== null && typeof value !== 'string'
      && typeof value !== 'boolean' && !(typeof value === 'number' && Number.isFinite(value)))) return undefined
  return { key: item.key, fieldId: item.fieldId, fieldName: item.fieldName, item: item.item, text: item.text,
    level: Number(item.level), depth: Number(item.depth), startIndex: Number(item.startIndex), endIndex: Number(item.endIndex), path: [...item.path] }
}
