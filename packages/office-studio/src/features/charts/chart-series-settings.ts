import type { ChartSeriesSelector } from '../../domain/model'
import type { ChartConfig, ChartData } from './chart-engine'
import type { ChartSettingsField } from './chart-settings-draft'

export interface ChartSeriesSettingTarget {
  value: string
  label: string
  selector: ChartSeriesSelector
  source?: { field: number; position: number; name: string }
  ambiguous?: boolean
  disabled?: boolean
}

/** Field and name scopes remain separate even when Bubble groups share the same Y field. */
export function chartSeriesSettingTargets(config: ChartConfig, fields: readonly ChartSettingsField[], data?: ChartData): ChartSeriesSettingTarget[] {
  const targets: ChartSeriesSettingTarget[] = fields.map(field => ({ value: `field:${field.value}`, label: `字段 · ${field.label}`, selector: field.value }))
  const names = new Map<string, Array<{ field: number; position: number; name: string }>>()
  for (const [position, series] of (data?.series ?? []).entries()) {
    const name = series.sourceName ?? series.name, matches = names.get(name) ?? []
    matches.push({ field: series.field, position, name }); names.set(name, matches)
  }
  const authoredNames = new Set(config.options?.namedSeries?.map(entry => entry.name))
  for (const name of authoredNames) if (!names.has(name)) names.set(name, [])
  for (const [name, matches] of names) {
    const authored = authoredNames.has(name)
    const suffix = matches.length > 1 ? authored ? '（源名称重复，请恢复该名称覆盖）' : '（源名称重复，请按字段设置）'
      : matches.length ? '' : '（当前数据中不存在）'
    targets.push({ value: `name:${JSON.stringify(name)}`, label: `名称 · ${JSON.stringify(name)}${suffix}`, selector: { name },
      ...(matches.length === 1 ? { source: matches[0]! } : {}), ...(matches.length > 1 ? { ambiguous: true, disabled: !authored } : {}) })
  }
  return targets
}

export function chartSeriesSettingPoints(config: ChartConfig, selector: ChartSeriesSelector, categories: readonly string[], fields: readonly ChartSettingsField[], data?: ChartData): ChartSettingsField[] {
  const items = categories.map((category, value) => ({ value, label: `${value + 1} · ${category}` }))
  const matches = typeof selector === 'number' ? [] : data?.series.filter(series => (series.sourceName ?? series.name) === selector.name) ?? []
  if (typeof selector !== 'number' && matches.length !== 1) return []
  if (config.type === 'boxplot') return items.filter(item => {
    const source = data?.series[item.value]
    return typeof selector === 'number' ? (source?.field ?? fields[item.value]?.value) === selector
      : source !== undefined && (source.sourceName ?? source.name) === selector.name
  })
  if (typeof selector === 'number') return items
  if (config.type !== 'bubble') return items
  const groupField = config.options?.mapping?.bubble?.seriesIndex
  if (groupField === undefined) return items
  const source = matches[0]!
  return items.filter(item => data?.records?.[item.value]
    ? String(data.records[item.value]![groupField] ?? '') === selector.name
    : source.values[item.value] != null || source.xValues?.[item.value] != null || source.sizeValues?.[item.value] != null)
}
