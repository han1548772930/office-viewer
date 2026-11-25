import type { ChartConfig } from './chart-engine'
import { chartHierarchyFields, chartHierarchyWeightField } from './chart-hierarchy-api'

export interface ChartHierarchySettingsMapping { hierarchyIndexes: number[]; valueIndex: number }
export function chartHierarchySettingsMapping(config: ChartConfig, fieldCount: number): ChartHierarchySettingsMapping {
  const hierarchyIndexes = config.options?.mapping?.hierarchyIndexes ?? [config.categoryField ?? 0]
  const used = new Set(hierarchyIndexes)
  let fallback = 0
  while (fallback < fieldCount && used.has(fallback)) fallback++
  return { hierarchyIndexes: [...hierarchyIndexes], valueIndex: config.valueFields?.[0] ?? (fallback < fieldCount ? fallback : fieldCount > 1 ? 1 : 0) }
}

function assertMapping(hierarchyIndexes: readonly number[], valueIndex: number, fieldCount: number): ChartHierarchySettingsMapping {
  const fields = chartHierarchyFields(hierarchyIndexes), weight = chartHierarchyWeightField(valueIndex)
  if (!Number.isSafeInteger(fieldCount) || fieldCount < 1 || [...fields, weight].some(index => index >= fieldCount)) throw new RangeError('层级和权重字段须位于当前数据区域内。')
  if (fields.includes(weight)) throw new RangeError('权重字段不能同时作为层级字段。')
  return { hierarchyIndexes: fields, valueIndex: weight }
}

/** Validate both roles before changing a draft; never publish a half-updated
 * hierarchy or let field order depend on the Select's option ordering. */
export function setChartHierarchySettingsMapping(config: ChartConfig, hierarchyIndexes: readonly number[], valueIndex: number, fieldCount: number): void {
  const next = assertMapping(hierarchyIndexes, valueIndex, fieldCount)
  ((config.options ??= {}).mapping ??= {}).hierarchyIndexes = next.hierarchyIndexes
  config.valueFields = [next.valueIndex]
}

export function addChartHierarchyField(config: ChartConfig, field: number, fieldCount: number): void {
  const current = chartHierarchySettingsMapping(config, fieldCount)
  setChartHierarchySettingsMapping(config, [...current.hierarchyIndexes, field], current.valueIndex, fieldCount)
}
export function replaceChartHierarchyField(config: ChartConfig, position: number, field: number, fieldCount: number): void {
  const current = chartHierarchySettingsMapping(config, fieldCount)
  if (!Number.isSafeInteger(position) || position < 0 || position >= current.hierarchyIndexes.length) throw new RangeError('层级字段位置无效。')
  current.hierarchyIndexes[position] = field
  setChartHierarchySettingsMapping(config, current.hierarchyIndexes, current.valueIndex, fieldCount)
}
export function moveChartHierarchyField(config: ChartConfig, position: number, direction: -1 | 1, fieldCount: number): void {
  const current = chartHierarchySettingsMapping(config, fieldCount), target = position + direction
  if ((direction !== -1 && direction !== 1) || !Number.isSafeInteger(position) || position < 0 || position >= current.hierarchyIndexes.length || target < 0 || target >= current.hierarchyIndexes.length) throw new RangeError('无法移动到该层级位置。')
  const field = current.hierarchyIndexes.splice(position, 1)[0]!
  current.hierarchyIndexes.splice(target, 0, field)
  setChartHierarchySettingsMapping(config, current.hierarchyIndexes, current.valueIndex, fieldCount)
}
export function removeChartHierarchyField(config: ChartConfig, position: number, fieldCount: number): void {
  const current = chartHierarchySettingsMapping(config, fieldCount)
  if (current.hierarchyIndexes.length <= 1) throw new RangeError('至少保留一个层级字段。')
  if (!Number.isSafeInteger(position) || position < 0 || position >= current.hierarchyIndexes.length) throw new RangeError('层级字段位置无效。')
  current.hierarchyIndexes.splice(position, 1)
  setChartHierarchySettingsMapping(config, current.hierarchyIndexes, current.valueIndex, fieldCount)
}
export function setChartHierarchyWeightField(config: ChartConfig, field: number, fieldCount: number): void {
  setChartHierarchySettingsMapping(config, chartHierarchySettingsMapping(config, fieldCount).hierarchyIndexes, field, fieldCount)
}
