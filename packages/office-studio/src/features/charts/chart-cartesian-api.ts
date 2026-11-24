import type { ChartAxisOptions, ChartAxisPointerOptions, ChartNullablePatch, ChartSemanticOptions, ChartSeriesOptions, ChartSeriesSelector } from '../../domain/model'
import { cloneChartOptions, mergeChartOptions } from './chart-options'
import { chartSeriesOverride, validateChartSeriesSelector } from './chart-series-selector'
import { chartLinearGradient, isChartLinearGradient } from './chart-color'

export const ChartAllSeriesStyleTargets = Object.freeze({ Border: 'border', Color: 'color', FillOpacity: 'fillOpacity', Label: 'label', Point: 'point', RightAxis: 'rightAxis' } as const)
export const ChartSeriesClearTargets = Object.freeze({ ...ChartAllSeriesStyleTargets, DataPoints: 'dataPoints', WaterfallStyles: 'waterfallStyles' } as const)
export const ChartAxisTargets = Object.freeze({ Min: 'min', Max: 'max', Reverse: 'reverse', Line: 'line', Tick: 'tick', Format: 'format', Label: 'label', GridLine: 'gridLine' } as const)
export const ChartAxisTitleAlignments = Object.freeze({ Start: 'start', Center: 'center', End: 'end' } as const)
export const ChartAxisPointerTargets = Object.freeze({ IndicatorLabelColor: 'indicatorLabelColor', IndicatorLabelTextColor: 'indicatorLabelTextColor', IndicatorLineColor: 'indicatorLineColor', IndicatorLineType: 'indicatorLineType' } as const)
export type ChartAllSeriesStyleTarget = typeof ChartAllSeriesStyleTargets[keyof typeof ChartAllSeriesStyleTargets]
export type ChartSeriesClearTarget = typeof ChartSeriesClearTargets[keyof typeof ChartSeriesClearTargets]
export type ChartAxisTarget = typeof ChartAxisTargets[keyof typeof ChartAxisTargets]
export type ChartAxisPointerTarget = typeof ChartAxisPointerTargets[keyof typeof ChartAxisPointerTargets]

/** The public patch is detached before null resets become explicit undefined in the model. */
export function chartNullableValues<T extends object>(value: ChartNullablePatch<T>): Partial<T> {
  const normalized = normalizePatchObject(value)
  const patch: Partial<T> = {}
  Object.assign(patch, normalized)
  return patch
}

function normalizePatchObject(value: unknown, nullResets = true): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype) throw new TypeError('图表配置 patch 须为普通对象。')
  const entries: Array<[string, unknown]> = []
  for (const [key, entry] of Object.entries(value)) {
    if (['__proto__', 'constructor', 'prototype'].includes(key)) throw new TypeError('图表配置 patch 包含无效属性。')
    if (entry === null) entries.push([key, key === 'fillPattern' || !nullResets ? null : undefined])
    else if (isChartLinearGradient(entry)) entries.push([key, chartLinearGradient(entry)])
    else if (Array.isArray(entry)) entries.push([key, detachPatchArray(entry)])
    else entries.push([key, entry && typeof entry === 'object' ? normalizePatchObject(entry, nullResets) : entry])
  }
  return Object.fromEntries(entries)
}

function detachPatchArray(value: readonly unknown[]): unknown[] {
  const array: unknown[] = []
  for (let index = 0; index < value.length; index++) {
    if (!Object.hasOwn(value, index)) throw new TypeError('图表配置 patch 不接受稀疏数组。')
    const item: unknown = value[index]
    array.push(Array.isArray(item) ? detachPatchArray(item) : item && typeof item === 'object' ? normalizePatchObject(item, false) : item)
  }
  return array
}

export function chartSeriesStyleValues(value: ChartNullablePatch<ChartSeriesOptions> | undefined): ChartSeriesOptions | undefined {
  return value === undefined ? undefined : cloneChartOptions({ allSeriesStyle: chartNullableValues<ChartSeriesOptions>(value) })!.allSeriesStyle
}

/** The selected collection is detached and replaced; all other selectors retain their authored state. */
export function patchChartSeriesOverride(current: ChartSemanticOptions | undefined, selector: ChartSeriesSelector, style: ChartSeriesOptions | undefined): ChartSemanticOptions {
  validateChartSeriesSelector(selector)
  const merged = style === undefined ? undefined
    : mergeChartOptions({ allSeriesStyle: chartSeriesOverride(current, selector) }, { allSeriesStyle: style })!.allSeriesStyle!
  if (typeof selector === 'number') {
    const series = cloneChartOptions({ series: current?.series })?.series ?? {}
    if (merged === undefined) delete series[selector]
    else series[selector] = merged
    return { series }
  }
  const namedSeries = cloneChartOptions({ namedSeries: current?.namedSeries })?.namedSeries ?? []
  const index = namedSeries.findIndex(entry => entry.name === selector.name)
  if (merged === undefined) { if (index >= 0) namedSeries.splice(index, 1) }
  else if (index >= 0) namedSeries[index] = { ...namedSeries[index]!, style: merged }
  else namedSeries.push({ name: selector.name, style: merged })
  return { namedSeries }
}

export function clearChartSeriesStyle(current: ChartSeriesOptions | undefined, target?: ChartSeriesClearTarget): ChartSeriesOptions | undefined {
  if (target === undefined) return undefined
  if (!Object.values(ChartSeriesClearTargets).includes(target)) throw new TypeError('无效的图表系列恢复目标。')
  const style = cloneChartOptions({ allSeriesStyle: current })?.allSeriesStyle ?? {}
  Object.assign(style, { [target]: undefined })
  return style
}

export function clearChartAllSeriesStyle(current: ChartSeriesOptions | undefined, target?: ChartAllSeriesStyleTarget): ChartSeriesOptions | undefined {
  if (target !== undefined && !Object.values(ChartAllSeriesStyleTargets).includes(target)) throw new TypeError('无效的图表全系列恢复目标。')
  return clearChartSeriesStyle(current, target)
}

/** Explicit linear intent clears a previously authored logarithm base. */
export function chartCartesianAxisPatch(current: Partial<ChartAxisOptions> | undefined, patch: ChartNullablePatch<ChartAxisOptions>): ChartAxisOptions {
  const source: ChartAxisOptions = { scaleType: 'linear', ...current }
  const values = chartNullableValues<ChartAxisOptions>(patch)
  const next: ChartAxisOptions = { ...source, ...values, scaleType: values.scaleType ?? source.scaleType }
  for (const key of ['line', 'tick', 'label', 'gridLine', 'titleStyle'] as const) {
    if (values[key]) Object.assign(next, { [key]: { ...source[key], ...values[key] } })
  }
  if (Object.hasOwn(values, 'scaleType') && values.scaleType === 'linear') next.logBase = undefined
  return cloneChartOptions({ rightYAxis: next })!.rightYAxis!
}

export function clearChartAxis(current: Partial<ChartAxisOptions> | undefined, target?: ChartAxisTarget): ChartAxisOptions {
  if (target !== undefined) {
    if (!Object.values(ChartAxisTargets).includes(target)) throw new TypeError('无效的图表坐标轴恢复目标。')
    const patch: Partial<ChartAxisOptions> = {}
    Object.assign(patch, { [target]: undefined })
    return chartCartesianAxisPatch(current, patch)
  }
  return chartCartesianAxisPatch(current, { scaleType: 'linear', logBase: undefined, min: undefined, max: undefined,
    baseTimeUnit: undefined, majorTimeUnit: undefined, majorUnit: undefined,
    title: undefined, titleStyle: undefined, titleAlignment: undefined, reverse: undefined, line: undefined, tick: undefined, format: undefined, label: undefined, gridLine: undefined })
}

export function chartAxisPointerValues(value: ChartNullablePatch<ChartAxisPointerOptions> | undefined): ChartAxisPointerOptions | undefined {
  return value === undefined ? undefined : cloneChartOptions({ axisPointer: chartNullableValues<ChartAxisPointerOptions>(value) })!.axisPointer
}

export function clearChartAxisPointer(current: ChartAxisPointerOptions | undefined, target?: ChartAxisPointerTarget): ChartAxisPointerOptions | undefined {
  if (target === undefined) return undefined
  if (!Object.values(ChartAxisPointerTargets).includes(target)) throw new TypeError('无效的图表轴指示器恢复目标。')
  return cloneChartOptions({ axisPointer: { ...current, [target]: undefined } })!.axisPointer
}
