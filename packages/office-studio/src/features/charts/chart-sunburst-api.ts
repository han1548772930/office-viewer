import type { ChartNullablePatch, ChartSunburstCenterLabelOptions, ChartSunburstLabelOptions, ChartType } from '../../domain/model'
import { chartNullableValues } from './chart-cartesian-api'
import { chartHierarchyFields, chartHierarchyWeightField } from './chart-hierarchy-api'
import { cloneChartOptions } from './chart-options'

const FONT_FIELDS = ['visible', 'color', 'fontSize', 'bold', 'italic', 'family', 'format']
function labelInput(value: object, fields: readonly string[], name: string): void {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype
    || Object.getOwnPropertySymbols(value).length || Object.keys(value).some(key => !fields.includes(key))) {
    throw new TypeError(`${name}包含无效字段。`)
  }
}

export function assertChartSunburstType(type: ChartType): void {
  if (type !== 'sunburst') throw new TypeError('旭日图专属设置仅适用于旭日图。')
}
export function chartSunburstHierarchyFields(indexes: readonly number[]): number[] { return chartHierarchyFields(indexes) }
export function chartSunburstValueField(index: number): number { return chartHierarchyWeightField(index) }

export function chartSunburstLabelValues(label: ChartNullablePatch<ChartSunburstLabelOptions>): ChartSunburstLabelOptions {
  labelInput(label, [...FONT_FIELDS, 'contentType', 'position', 'backgroundColor'], '旭日图标签')
  if (label.position !== null && label.position !== undefined && label.position !== 'inside' && label.position !== 'outside') {
    throw new TypeError('旭日图标签位置需要 inside、outside 或 null（自动）。')
  }
  const values = chartNullableValues<ChartSunburstLabelOptions>(label)
  return cloneChartOptions({ sunburst: { label: values } })!.sunburst!.label!
}

export function chartSunburstCenterLabelValues(label: ChartNullablePatch<ChartSunburstCenterLabelOptions>): ChartSunburstCenterLabelOptions {
  // The persisted legacy .font shape remains readable, but the fixed public
  // setter authors direct font properties rather than another nested API.
  labelInput(label, FONT_FIELDS, '旭日图中心标签')
  const values = chartNullableValues<ChartSunburstCenterLabelOptions>(label)
  return cloneChartOptions({ sunburst: { centerLabel: values } })!.sunburst!.centerLabel!
}
