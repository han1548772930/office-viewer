/** Stable finite-number operations shared by layout and the small-data statistics path. */
export function finiteChartValue(value: number, label = '图表计算结果'): number {
  if (!Number.isFinite(value)) throw new RangeError(`${label}超出有限数值范围。`)
  return value
}

export function chartInterpolate(first: number, last: number, ratio: number): number {
  if (ratio <= 0) return first
  if (ratio >= 1) return last
  const difference = last - first
  return finiteChartValue(Number.isFinite(difference) ? first + difference * ratio : first * (1 - ratio) + last * ratio)
}

export function chartLinearRatio(value: number, min: number, max: number): number {
  if (min === max) return 0
  const difference = max - min, offset = value - min
  if (Number.isFinite(difference) && Number.isFinite(offset)) return offset / difference
  const scale = Math.max(Math.abs(value), Math.abs(min), Math.abs(max))
  return (value / scale - min / scale) / (max / scale - min / scale)
}

export function chartFiniteSum(values: readonly number[], label = '图表合计'): number {
  let scale = 0
  for (const value of values) scale = Math.max(scale, Math.abs(finiteChartValue(value, label)))
  if (!scale) return 0
  let total = 0, compensation = 0
  for (const value of values) {
    const addend = value / scale - compensation, next = total + addend
    compensation = next - total - addend; total = next
  }
  return finiteChartValue(total * scale, label)
}

export function chartScaledMoments(samples: readonly number[]): { mean: number; deviation: number; scale: number; scaledMean: number; scaledVariance: number } {
  let scale = 0
  for (const value of samples) scale = Math.max(scale, Math.abs(finiteChartValue(value)))
  if (!samples.length || !scale) return { mean: 0, deviation: 0, scale: 1, scaledMean: 0, scaledVariance: 0 }
  const scaledMean = samples.reduce((sum, value) => sum + value / scale / samples.length, 0)
  const scaledVariance = samples.reduce((sum, value) => sum + (value / scale - scaledMean) ** 2 / samples.length, 0)
  return { mean: finiteChartValue(scaledMean * scale), deviation: finiteChartValue(Math.sqrt(scaledVariance) * scale), scale, scaledMean, scaledVariance }
}
