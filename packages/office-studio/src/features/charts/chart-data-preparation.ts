import { CHART_PREPARATION_CHUNK } from './chart-compute-lifecycle'

/** Projection writes yield at fixed work boundaries, without a second engine. */
export function* chartMapSteps<T, R>(source: readonly T[], read: (value: T, index: number) => R): Generator<void, R[], void> {
  const result: R[] = []
  for (let index = 0; index < source.length; index++) {
    result.push(read(source[index]!, index))
    if ((index + 1) % CHART_PREPARATION_CHUNK === 0) yield
  }
  return result
}

export function* chartNumericFieldSteps<T>(records: readonly (readonly T[])[], fieldCount: number,
  numeric: (value: T | undefined) => number | null): Generator<void, number[], void> {
  const result: number[] = []
  let work = 0
  for (let field = 0; field < fieldCount; field++) {
    for (const record of records) {
      const value = numeric(record[field])
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
      if (value !== null) { result.push(field); break }
    }
  }
  return result
}

/** Conservative source/reference allowance, not a measured heap figure. */
export function chartProjectionSourceBytes(rows: number, fields: number): number {
  return rows * (fields * 64 + 192) + 8 * 1024 * 1024
}
