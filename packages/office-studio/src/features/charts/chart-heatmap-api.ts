import type { ChartLabelOptions, ChartNullablePatch, ChartType } from '../../domain/model'
import { chartNullableValues } from './chart-cartesian-api'
import { cloneChartOptions } from './chart-options'

/** Fixed v1.6.0 ChartVisualMapType and IChartHeatmapLabelSpec. */
export const ChartVisualMapTypes = Object.freeze({ Continuous: 'continuous', Piecewise: 'piecewise' } as const)
export type ChartHeatmapVisualMapType = typeof ChartVisualMapTypes[keyof typeof ChartVisualMapTypes]
export type ChartHeatmapLabelOptions = Omit<ChartLabelOptions, 'contentType'>
const CELL_LABEL_FIELDS = new Set(['visible', 'position', 'format', 'family', 'fontSize', 'color', 'bold', 'italic'])

export function assertChartHeatmapType(type: ChartType): void {
  if (type !== 'heatmap') throw new TypeError('此操作仅适用于热力图。')
}
export function chartHeatmapVisualMapType(value: unknown): ChartHeatmapVisualMapType {
  if (value !== ChartVisualMapTypes.Continuous && value !== ChartVisualMapTypes.Piecewise) throw new TypeError('热力图色阶类型无效。')
  return value
}
export function chartHeatmapValueRange(min: unknown, max: unknown): { min: number; max: number } {
  if (typeof min !== 'number' || typeof max !== 'number' || !Number.isFinite(min) || !Number.isFinite(max)) throw new RangeError('热力图色阶范围须为有限数值。')
  if (min > max) throw new RangeError('热力图色阶最小值不能大于最大值。')
  return { min, max }
}
/** Only the public value-label fields are newly authored here. Existing generic
 * label extensions in retained models are not discarded by a type switch. */
export function chartHeatmapCellLabelValues(value: ChartNullablePatch<ChartHeatmapLabelOptions>): ChartHeatmapLabelOptions {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype
    || Object.getOwnPropertySymbols(value).length || Object.keys(value).some(key => !CELL_LABEL_FIELDS.has(key))) throw new TypeError('热力图单元格标签包含无效字段。')
  const label = chartNullableValues<ChartHeatmapLabelOptions>(value)
  return cloneChartOptions({ heatmap: { label } })!.heatmap!.label!
}
export function chartHeatmapValueUnit(value: unknown): string {
  if (typeof value !== 'string') throw new TypeError('热力图数值单位须为文本。')
  return value
}
