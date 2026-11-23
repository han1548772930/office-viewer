import type { ChartGaugeNullablePatch, ChartGaugeOptions, ChartGaugeRangeOptions, ChartGaugeTicksPatch } from '../../domain/model'
import { cloneChartOptions } from './chart-options'

/** Resolve a partial semantic patch against the existing scale before validating its defaults.
 * A negative-domain scale can therefore update just max without being checked against min=0. */
export function patchChartGaugeOptions(current: ChartGaugeOptions | undefined, patch: ChartGaugeOptions | undefined): ChartGaugeOptions | undefined {
  if (patch === undefined) return undefined
  plainGaugePatch(patch)
  const previous = cloneChartOptions({ gauge: current })?.gauge
  const next: ChartGaugeOptions = { ...previous, ...patch }
  for (const key of ['axisLine', 'axisTick', 'splitLine', 'axisLabel', 'pointer', 'anchor', 'progress', 'title', 'detail', 'label'] as const) {
    const value = patch[key]
    if (value && typeof value === 'object' && !Array.isArray(value)) {
      plainGaugePatch(value)
      if (previous?.[key]) Object.assign(next, { [key]: { ...previous[key], ...value } })
    }
  }
  if (Array.isArray(patch.ranges)) for (const range of patch.ranges) plainGaugePatch(range)
  return cloneChartOptions({ gauge: next })!.gauge
}

export function chartGaugeNullableValues<T extends object>(value: ChartGaugeNullablePatch<T>): Partial<T> {
  plainGaugePatch(value)
  const patch: Partial<T> = {}
  for (const [key, entry] of Object.entries(value)) {
    Object.assign(patch, { [key]: entry === null ? undefined : entry })
  }
  return patch
}

export function chartGaugeTicksOptions(value: ChartGaugeTicksPatch): ChartGaugeOptions {
  plainGaugePatch(value)
  const patch: ChartGaugeOptions = {}
  if (Object.hasOwn(value, 'majorDivisions')) patch.splitNumber = value.majorDivisions ?? undefined
  if (Object.hasOwn(value, 'minorDivisionsPerMajor')) patch.axisTick = { splitNumber: value.minorDivisionsPerMajor ?? undefined }
  return patch
}

export function chartGaugeRangesOptions(value: readonly ChartGaugeRangeOptions[]): ChartGaugeOptions {
  if (!Array.isArray(value)) throw new TypeError('仪表颜色区间须为数组。')
  return { ranges: Array.from(value, range => { plainGaugePatch(range); return { ...range } }) }
}

function plainGaugePatch(value: unknown): void {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.getPrototypeOf(value) !== Object.prototype) throw new TypeError('仪表配置须为普通对象。')
  if (['__proto__', 'constructor', 'prototype'].some(key => Object.hasOwn(value, key))) throw new TypeError('仪表配置包含无效属性。')
}
