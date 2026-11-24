import type { ChartSeriesSelector } from '../../domain/model'
import type { ChartConfig, ChartData } from './chart-engine'
import type { ChartSettingsField } from './chart-settings-draft'
import { chartSeriesOverride } from './chart-series-selector'

export interface ChartSeriesSettingTarget {
  value: string
  label: string
  selector: ChartSeriesSelector
  source?: { field: number; position: number; name: string }
  ambiguous?: boolean
  disabled?: boolean
}

export const CHART_BOXPLOT_SAMPLE_CHOICES_LIMIT = 16_384
export const CHART_CATEGORY_POINT_CHOICES_LIMIT = 16_384
type SettingSource = { field: number; position: number; name: string }

/** Field maps and named-series arrays replace atomically at the settings
 * command boundary. Removing one authored point resumes the earlier layers. */
export function clearChartSeriesSettingPoint(config: ChartConfig, selector: ChartSeriesSelector, index: number): void {
  if (!Number.isSafeInteger(index) || index < 0) throw new RangeError('图表数据项目序号须为非负安全整数。')
  const points = chartSeriesOverride(config.options, selector)?.dataPoints
  if (points) delete points[index]
}

/** Field and name scopes remain separate even when Bubble groups share the same Y field. */
export function chartSeriesSettingTargets(config: ChartConfig, fields: readonly ChartSettingsField[], data?: ChartData): ChartSeriesSettingTarget[] {
  const names = new Map<string, SettingSource[]>(), sourceFields = new Map<number, SettingSource[]>()
  for (const [position, series] of (data?.series ?? []).entries()) {
    const name = series.sourceName ?? series.name, source = { field: series.field, position, name }, matches = names.get(name) ?? [], fieldMatches = sourceFields.get(series.field) ?? []
    matches.push(source); names.set(name, matches)
    fieldMatches.push(source); sourceFields.set(series.field, fieldMatches)
  }
  const targets: ChartSeriesSettingTarget[] = fields.map(field => {
    const sources = sourceFields.get(field.value)
    return { value: `field:${field.value}`, label: `字段 · ${field.label}`, selector: field.value, ...(sources?.length === 1 ? { source: sources[0]! } : {}) }
  })
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

function boxplotSettingSources(selector: ChartSeriesSelector, data?: ChartData): ChartData['series'] {
  const matches = data?.series.filter(series => typeof selector === 'number' ? series.field === selector : (series.sourceName ?? series.name) === selector.name) ?? []
  return typeof selector === 'number' || matches.length === 1 ? matches : []
}

function boxplotSettingSample(sources: ChartData['series'], index: number, data?: ChartData): ChartSettingsField | undefined {
  if (!Number.isSafeInteger(index) || index < 0) return undefined
  for (const source of sources) {
    const value = source.values[index]
    if (value === null || value === undefined || !Number.isFinite(value)) continue
    const category = data?.categories[index], text = source.formattedValues?.[index] ?? String(value)
    return { value: index, label: `${index + 1} · ${category ? `${category} · ` : ''}${text}` }
  }
  return undefined
}

/** Direct lookup makes every source sample reachable without an unbounded option list. */
export function chartSeriesSettingSample(selector: ChartSeriesSelector, index: number, data?: ChartData): ChartSettingsField | undefined {
  return boxplotSettingSample(boxplotSettingSources(selector, data), index, data)
}

export function chartSeriesSettingSampleCount(selector: ChartSeriesSelector, data?: ChartData): number {
  return boxplotSettingSources(selector, data).reduce((count, source) => Math.max(count, source.values.length), 0)
}

export function chartCategorySettingPoint(categories: readonly string[], index: number): ChartSettingsField | undefined {
  if (!Number.isSafeInteger(index) || index < 0 || index >= categories.length) return undefined
  return { value: index, label: `${index + 1} · ${categories[index]}` }
}

/** The direct number field makes every materialized mark reachable. Only the
 * bounded initial choices and the selected mark become Select options. */
export function chartCategorySettingPoints(categories: readonly string[], selectedPoint?: number): ChartSettingsField[] {
  const count = Math.min(CHART_CATEGORY_POINT_CHOICES_LIMIT, categories.length), items: ChartSettingsField[] = []
  for (let index = 0; index < count; index++) items.push(chartCategorySettingPoint(categories, index)!)
  if (selectedPoint !== undefined && selectedPoint >= count) {
    const item = chartCategorySettingPoint(categories, selectedPoint)
    if (item) items.push(item)
  }
  return items
}

function categorySettingMembership(config: ChartConfig, selector: ChartSeriesSelector, data?: ChartData): ((index: number) => boolean) | undefined {
  if (typeof selector === 'number') return () => true
  const matches = data?.series.filter(series => (series.sourceName ?? series.name) === selector.name) ?? []
  if (matches.length !== 1) return undefined
  const groupField = config.type === 'bubble' ? config.options?.mapping?.bubble?.seriesIndex : undefined
  if (groupField === undefined) return () => true
  const source = matches[0]!
  return index => data?.records?.[index] ? String(data.records[index]![groupField] ?? '') === selector.name
    : source.values[index] != null || source.xValues?.[index] != null || source.sizeValues?.[index] != null
}

export function chartSeriesSettingPoint(config: ChartConfig, selector: ChartSeriesSelector, categories: readonly string[], data: ChartData | undefined, index: number): ChartSettingsField | undefined {
  if (config.type === 'boxplot') return chartSeriesSettingSample(selector, index, data)
  const item = chartCategorySettingPoint(categories, index)
  return item && categorySettingMembership(config, selector, data)?.(index) ? item : undefined
}

export function chartSeriesSettingPoints(config: ChartConfig, selector: ChartSeriesSelector, categories: readonly string[], _fields: readonly ChartSettingsField[], data?: ChartData, selectedSample?: number): ChartSettingsField[] {
  if (config.type === 'boxplot') {
    const sources = boxplotSettingSources(selector, data), count = Math.min(CHART_BOXPLOT_SAMPLE_CHOICES_LIMIT, sources.reduce((length, source) => Math.max(length, source.values.length), 0)), items: ChartSettingsField[] = []
    for (let index = 0; index < count; index++) { const item = boxplotSettingSample(sources, index, data); if (item) items.push(item) }
    if (selectedSample !== undefined && selectedSample >= count) { const item = boxplotSettingSample(sources, selectedSample, data); if (item) items.push(item) }
    return items
  }
  const membership = categorySettingMembership(config, selector, data)
  return membership ? chartCategorySettingPoints(categories, selectedSample).filter(item => membership(item.value)) : []
}
