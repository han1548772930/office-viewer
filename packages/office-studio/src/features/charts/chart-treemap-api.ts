import type { ChartNullablePatch, ChartTreemapLabelOptions, ChartTreemapOptions, ChartType } from '../../domain/model'
import { chartNullableValues } from './chart-cartesian-api'
import { cloneChartOptions } from './chart-options'
import { chartHierarchyFields, chartHierarchyWeightField } from './chart-hierarchy-api'

export function assertChartTreemapType(type: ChartType): void {
  if (type !== 'treemap') throw new TypeError('树图专属设置仅适用于树图。')
}

export function chartTreemapHierarchyFields(indexes: readonly number[]): number[] {
  return chartHierarchyFields(indexes)
}

export function chartTreemapValueField(index: number): number {
  return chartHierarchyWeightField(index)
}

export function chartTreemapParentLabelLayout(layout: NonNullable<ChartTreemapOptions['parentLabelLayout']>): NonNullable<ChartTreemapOptions['parentLabelLayout']> {
  if (!['banner', 'overlapping', 'none'].includes(layout)) throw new TypeError('树图父级标签布局需要 banner、overlapping 或 none。')
  return layout
}

export function chartTreemapLabelValues(label: ChartNullablePatch<ChartTreemapLabelOptions>): ChartTreemapLabelOptions {
  if (!label || typeof label !== 'object' || Array.isArray(label) || Object.getPrototypeOf(label) !== Object.prototype
    || Object.getOwnPropertySymbols(label).length || Object.keys(label).some(key => !['visible', 'contentType', 'color', 'fontSize', 'bold', 'italic', 'family', 'format', 'backgroundColor'].includes(key))) {
    throw new TypeError('树图叶标签包含无效字段。')
  }
  const values = chartNullableValues<ChartTreemapLabelOptions>(label)
  return cloneChartOptions({ treemap: { label: values } })!.treemap!.label!
}
