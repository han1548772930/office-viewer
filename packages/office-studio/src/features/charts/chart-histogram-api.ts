import type { ChartType } from '../../domain/model'

export function assertChartHistogramType(type: ChartType): void {
  if (type !== 'histogram') throw new TypeError('直方图专属设置仅适用于直方图。')
}

/** Ineffective finite authors remain persisted; non-finite public calls are no-op. */
export function isChartHistogramAuthoredNumber(value: number): boolean {
  return typeof value === 'number' && Number.isFinite(value)
}
