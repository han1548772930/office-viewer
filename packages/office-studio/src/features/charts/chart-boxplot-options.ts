import type { ChartSemanticOptions } from '../../domain/model'

/** Fixed v1.6.0 uses the median of each half (excluding the odd middle
 * sample) and full-sample whiskers. Retain authored interpolated/IQR models
 * from earlier snapshots when their legacy options are present. */
export function chartBoxplotStatisticOptions(options: ChartSemanticOptions['boxplot'] = {}): {
  quartileMethod: 'medianOfHalves' | 'inclusive' | 'exclusive'
  whiskerMethod: 'range' | 'iqr'
  whiskerMultiplier: number
} {
  const quartileMethod = options.quartileMethod ?? 'medianOfHalves'
  const legacyIqr = options.quartileMethod !== undefined && options.quartileMethod !== 'medianOfHalves'
    || options.whiskerMultiplier !== undefined
  return { quartileMethod, whiskerMethod: options.whiskerMethod ?? (legacyIqr ? 'iqr' : 'range'),
    whiskerMultiplier: options.whiskerMultiplier ?? 1.5 }
}
