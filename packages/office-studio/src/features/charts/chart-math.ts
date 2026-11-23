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

/** Combine finite factors without overflowing or underflowing an intermediate
 * product. Only the final representable result is rounded to binary64. */
export function chartFiniteProduct(factors: readonly number[], label = '图表计算结果'): number {
  let sign = 1, mantissa = 1, exponent = 0, zero = false
  for (const factor of factors) {
    finiteChartValue(factor, label)
    if (factor === 0) { zero = true; continue }
    if (factor < 0) sign = -sign
    const absolute = Math.abs(factor)
    let power = Math.min(1023, Math.floor(Math.log2(absolute))), part = absolute / 2 ** power
    // log2 can round a value immediately below a power of two upward.
    if (part < 1) { part *= 2; power-- }
    mantissa *= part; exponent += power
    if (mantissa >= 2) { mantissa /= 2; exponent++ }
  }
  if (zero) return 0
  if (exponent > 1023) throw new RangeError(`${label}超出有限数值范围。`)
  const result = exponent < -1022 ? mantissa * 2 ** -1022 * 2 ** (exponent + 1022) : mantissa * 2 ** exponent
  return finiteChartValue(sign * result, label)
}

export interface ChartScaledMoments { mean: number; scale: number; scaledMean: number; scaledVariance: number }

/** Population moments in normalized space. Callers construct the requested
 * deviation only when needed, after applying its divisor and multiplier. */
export function chartScaledMoments(samples: readonly number[]): ChartScaledMoments {
  let scale = 0
  for (const value of samples) scale = Math.max(scale, Math.abs(finiteChartValue(value)))
  if (!samples.length || !scale) return { mean: 0, scale: 1, scaledMean: 0, scaledVariance: 0 }
  let total = 0, compensation = 0
  for (const value of samples) {
    const addend = value / scale - compensation, next = total + addend
    compensation = next - total - addend; total = next
  }
  const scaledMean = Math.max(-1, Math.min(1, total / samples.length))
  total = 0; compensation = 0
  for (const value of samples) {
    const addend = (value / scale - scaledMean) ** 2 - compensation, next = total + addend
    compensation = next - total - addend; total = next
  }
  const scaledVariance = Math.max(0, Math.min(1, total / samples.length))
  return { mean: finiteChartValue(scaledMean * scale), scale, scaledMean, scaledVariance }
}
