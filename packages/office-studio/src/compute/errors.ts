import type { ComputeErrorCode, ComputeErrorData } from './protocol.ts'
import { isRecord } from './protocol.ts'

const CODES = new Set<string>([
  'COMPUTE_INVALID_INPUT', 'COMPUTE_MEMORY_LIMIT', 'COMPUTE_UNSUPPORTED_TYPE',
  'COMPUTE_WASM_LOAD_FAILED', 'COMPUTE_WASM_API_MISSING', 'COMPUTE_PROTOCOL_MISMATCH',
  'COMPUTE_FAILED', 'COMPUTE_CANCELLED', 'COMPUTE_TIMEOUT', 'COMPUTE_STALE_RESULT',
  'COMPUTE_WORKER_FAILED', 'COMPUTE_WASM_TRAP', 'COMPUTE_DISPOSED', 'COMPUTE_NOT_READY',
])

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
  if (error instanceof WebAssembly.RuntimeError) return new ComputeError('COMPUTE_WASM_TRAP', '计算 WASM 执行异常', error.message)
  if (error instanceof RangeError) return new ComputeError('COMPUTE_MEMORY_LIMIT', '计算内存分配失败', error.message)
  if (typeof error === 'string') {
    try {
      const parsed: unknown = JSON.parse(error)
      if (isErrorData(parsed)) return new ComputeError(parsed.code, parsed.message, parsed.details)
    } catch { /* A non-JSON WASM trap remains available in details below. */ }
  }
  if (isErrorData(error)) return new ComputeError(error.code, error.message, error.details)
  return new ComputeError(fallback, '计算未能完成', error instanceof Error ? error.message : String(error))
}

