import type { ChartColor, ChartLinearGradient } from '../../domain/model'
import { CHART_GRADIENT_MAX_STOPS, chartColorRepresentative, chartLinearGradient, isChartLinearGradient, validateChartColor } from './chart-color'

export type ChartColorMode = 'inherit' | 'solid' | 'gradient'
export function chartColorMode(color: ChartColor | undefined): ChartColorMode {
  return color === undefined ? 'inherit' : isChartLinearGradient(color) ? 'gradient' : 'solid'
}

/** Mode changes are detached edits. The local editor starts a horizontal fade
 * from an effective solid color; inherited gradients keep their authored vector. */
export function chartColorForMode(color: ChartColor | undefined, mode: ChartColorMode, fallback: ChartColor = '#2563eb'): ChartColor | undefined {
  if (mode === 'inherit') return undefined
  const effective = color ?? fallback
  validateChartColor(effective)
  if (mode === 'solid') return chartColorRepresentative(effective)
  if (mode !== 'gradient') throw new TypeError('无效的图表颜色填充方式。')
  if (isChartLinearGradient(effective)) return chartLinearGradient(effective)
  return chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 0 },
    stops: [{ offset: 0, color: effective }, { offset: 1, color: effective, opacity: 0 }] })
}

export function patchChartGradientCoordinate(color: ChartLinearGradient, endpoint: 'start' | 'end', coordinate: 'x' | 'y', value: number): ChartLinearGradient {
  if (!['start', 'end'].includes(endpoint) || !['x', 'y'].includes(coordinate)) throw new TypeError('无效的图表渐变坐标。')
  const next = chartLinearGradient(color)
  next[endpoint][coordinate] = value
  return chartLinearGradient(next)
}

export function patchChartGradientStop(color: ChartLinearGradient, index: number, patch: Partial<ChartLinearGradient['stops'][number]>): ChartLinearGradient {
  const next = chartLinearGradient(color)
  checkedStop(next, index)
  if (!patch || typeof patch !== 'object' || Array.isArray(patch) || Object.getPrototypeOf(patch) !== Object.prototype
    || Reflect.ownKeys(patch).some(key => typeof key !== 'string' || !['offset', 'color', 'opacity'].includes(key))) throw new TypeError('无效的图表渐变色标更新。')
  next.stops[index] = { ...next.stops[index]!, ...patch }
  return chartLinearGradient(next)
}

/** Append rather than reorder: duplicate offsets can represent a hard edge. */
export function appendChartGradientStop(color: ChartLinearGradient): ChartLinearGradient {
  const next = chartLinearGradient(color)
  if (next.stops.length >= CHART_GRADIENT_MAX_STOPS) throw new RangeError(`图表渐变最多接受 ${CHART_GRADIENT_MAX_STOPS} 个色标。`)
  next.stops.push({ ...next.stops.at(-1)! })
  return chartLinearGradient(next)
}

export function removeChartGradientStop(color: ChartLinearGradient, index: number): ChartLinearGradient {
  const next = chartLinearGradient(color)
  checkedStop(next, index)
  if (next.stops.length <= 2) throw new RangeError('图表渐变至少需要两个色标。')
  next.stops.splice(index, 1)
  return chartLinearGradient(next)
}

function checkedStop(color: ChartLinearGradient, index: number): void {
  if (!Number.isSafeInteger(index) || index < 0 || index >= color.stops.length) throw new RangeError('无效的图表渐变色标序号。')
}
