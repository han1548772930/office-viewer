import type { ComputeErrorCode, ComputeErrorData } from './protocol.ts'
import { isRecord } from './protocol.ts'
import { serializeWasmError } from '../platform/wasm/task-protocol.ts'

const CODES = new Set<string>([
  'COMPUTE_INVALID_INPUT', 'COMPUTE_MEMORY_LIMIT', 'COMPUTE_UNSUPPORTED_TYPE',
  'COMPUTE_WASM_LOAD_FAILED', 'COMPUTE_WASM_API_MISSING', 'COMPUTE_PROTOCOL_MISMATCH',
  'COMPUTE_FAILED', 'COMPUTE_CANCELLED', 'COMPUTE_TIMEOUT', 'COMPUTE_STALE_RESULT',
  'COMPUTE_WORKER_FAILED', 'COMPUTE_WASM_TRAP', 'COMPUTE_DISPOSED', 'COMPUTE_NOT_READY',
  'COMPUTE_RESULT_CONSUMED',
])
const BRIDGE_CODES: Readonly<Record<string, ComputeErrorCode>> = {
  WASM_CONTROL_INVALID: 'COMPUTE_INVALID_INPUT',
  WASM_INPUT_ALIAS: 'COMPUTE_INVALID_INPUT',
  WASM_MEMORY_LIMIT: 'COMPUTE_MEMORY_LIMIT',
  WASM_PROTOCOL_MISMATCH: 'COMPUTE_PROTOCOL_MISMATCH',
  WASM_OUTPUT_ALIAS: 'COMPUTE_PROTOCOL_MISMATCH',
  WASM_CONTROL_SERIALIZATION: 'COMPUTE_PROTOCOL_MISMATCH',
  WASM_API_MISSING: 'COMPUTE_WASM_API_MISSING',
  WASM_LOAD_FAILED: 'COMPUTE_WASM_LOAD_FAILED',
  WASM_TRAP: 'COMPUTE_WASM_TRAP',
}
export class ComputeError extends Error {
  readonly code: ComputeErrorCode
  readonly details?: string
  constructor(code: ComputeErrorCode, message: string, details?: string) {
    super(message); this.name = 'ComputeError'; this.code = code; this.details = details
  }
  toJSON(): ComputeErrorData {
    return { code: this.code, message: this.message, ...(this.details ? { details: this.details } : {}) }
  }
}
export function isErrorData(value: unknown): value is ComputeErrorData {
  return isRecord(value) && typeof value.code === 'string' && CODES.has(value.code)
    && typeof value.message === 'string' && (value.details === undefined || typeof value.details === 'string')
}
export function normalizeComputeError(error: unknown, fallback: ComputeErrorCode = 'COMPUTE_FAILED'): ComputeError {
  if (error instanceof ComputeError) return error
  const value = serializeWasmError(error, { code: fallback, message: '计算未能完成' })
  if (isErrorData(value)) return new ComputeError(value.code, value.message, value.details)
  const code = BRIDGE_CODES[value.code] ?? fallback
  return new ComputeError(code, value.message, value.details ?? (value.code !== code ? value.code : undefined))
}
