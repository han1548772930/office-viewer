import { isRecord } from './protocol.ts'
import { isUnsignedCount } from '../platform/wasm/task-protocol.ts'
import type { ValidationListQueryOutput, ValidationListQueryParams } from './validation-list-query-protocol.ts'

export function readValidationListQueryOutput(value: unknown, params: ValidationListQueryParams): ValidationListQueryOutput | null {
  if (!isRecord(value) || value.kind !== 'validationListQuery' || !isUnsignedCount(value.total) || value.total > params.rowCount
    || !(value.indices instanceof Uint32Array) || !(value.indices.buffer instanceof ArrayBuffer)
    || value.indices.byteOffset !== 0 || value.indices.byteLength !== value.indices.buffer.byteLength
    || value.indices.length !== Math.min(Math.max(0, value.total - params.start), params.count)) return null
  try { new Uint8Array(value.indices.buffer) } catch { return null }
  if (value.indices.length && (value.indices[0]! >= params.rowCount || value.indices.at(-1)! >= params.rowCount
    || value.indices.length > 1 && value.indices[0]! >= value.indices.at(-1)!)) return null
  // The kernel checks all membership/order. The host maps each returned index
  // with bounds checking while yielding; it never publishes a partial window.
  return { kind: 'validationListQuery', indices: value.indices, total: value.total }
}
