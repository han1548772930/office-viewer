import type { CellScalar } from './model.ts'

/** Public raw identities shared by filters, Pivot caches and Slicer queries. */
export const RecordValueType = { Blank: 'blank', String: 'string', Number: 'number', Date: 'date', Boolean: 'boolean', Error: 'error' } as const
export const FilterSelectionMode = { All: 'all', Include: 'include', Exclude: 'exclude' } as const
export type RecordValueType = typeof RecordValueType[keyof typeof RecordValueType]
export type FilterSelectionMode = typeof FilterSelectionMode[keyof typeof FilterSelectionMode]
export type RecordValue = { type: 'blank' } | { type: 'string' | 'error'; value: string }
  | { type: 'number' | 'date'; value: number } | { type: 'boolean'; value: boolean }
export type RecordSelection = { mode: 'all' } | { mode: 'include' | 'exclude'; values: RecordValue[] }
export type PivotRecordValue = CellScalar | RecordValue

const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)

export function normalizeRecordValue(input: unknown): RecordValue {
  if (!object(input)) throw new TypeError('Invalid record value')
  switch (input.type) {
    case 'blank':
      if ('value' in input) throw new TypeError('Blank record values cannot contain a value')
      return { type: 'blank' }
    case 'string': case 'error':
      if (typeof input.value !== 'string') throw new TypeError('Record text values require a string')
      return { type: input.type, value: input.value }
    case 'boolean':
      if (typeof input.value !== 'boolean') throw new TypeError('Record boolean values require a boolean')
      return { type: 'boolean', value: input.value }
    case 'number': case 'date':
      if (typeof input.value !== 'number') throw new TypeError('Record numeric values require a number')
      if (!Number.isFinite(input.value)) throw new RangeError('Record numeric values must be finite')
      return { type: input.type, value: input.value === 0 ? 0 : input.value }
    default: throw new TypeError('Unknown record value type')
  }
}

export function getRecordValueKey(input: RecordValue): string {
  const value = normalizeRecordValue(input)
  return JSON.stringify(value.type === 'blank' ? ['blank'] : [value.type, value.value])
}

export function recordValueFromScalar(value: CellScalar): RecordValue {
  if (value === null) return { type: 'blank' }
  if (typeof value === 'string') return { type: 'string', value }
  if (typeof value === 'boolean') return { type: 'boolean', value }
  return normalizeRecordValue({ type: 'number', value })
}

export function recordValueScalar(value: RecordValue): CellScalar { return value.type === 'blank' ? null : value.value }
export function isRecordValue(value: PivotRecordValue): value is RecordValue { return value !== null && typeof value === 'object' }
export function pivotRecordScalar(value: PivotRecordValue): CellScalar { return isRecordValue(value) ? recordValueScalar(value) : value }

/** Legacy scalar rules retain their raw-value semantics; typed rules also retain Date/Error identity. */
export function pivotRecordValueKey(value: PivotRecordValue): string {
  return `${isRecordValue(value) ? 'typed' : 'scalar'}:${getRecordValueKey(isRecordValue(value) ? value : recordValueFromScalar(value))}`
}

export function normalizeRecordSelection(input: unknown): RecordSelection {
  if (!object(input) || typeof input.mode !== 'string' || !['all', 'include', 'exclude'].includes(input.mode)
    || Object.keys(input).some(key => !['mode', 'values'].includes(key))) throw new TypeError('Invalid record selection')
  if (input.mode === 'all') {
    if ('values' in input) throw new TypeError('All record selection cannot contain values')
    return { mode: 'all' }
  }
  if (!Array.isArray(input.values)) throw new TypeError('Record selection requires a values array')
  const seen = new Set<string>(), values: RecordValue[] = []
  for (const inputValue of input.values) {
    const value = normalizeRecordValue(inputValue), key = getRecordValueKey(value)
    if (!seen.has(key)) { seen.add(key); values.push(value) }
  }
  return { mode: input.mode === 'include' ? 'include' : 'exclude', values }
}

export function compileRecordSelection(input: RecordSelection): (value: RecordValue) => boolean {
  const selection = normalizeRecordSelection(input)
  if (selection.mode === 'all') return () => true
  const selected = new Set(selection.values.map(getRecordValueKey))
  return value => selected.has(getRecordValueKey(value)) === (selection.mode === 'include')
}
