import { WASM_PROTOCOL_VERSION } from '../platform/wasm/task-protocol.ts'
import type { BoundaryCopyStats, WasmTaskControl, WorkerTransportStats } from '../platform/wasm/task-protocol.ts'
export const COMPUTE_PROTOCOL_VERSION = WASM_PROTOCOL_VERSION
export const COMPUTE_LIMITS = {
  maxElements: 16_000_000,
  maxRows: 10_000_000,
  maxKeys: 32,
  defaultMemoryBudget: 256 * 1024 * 1024,
  maxPendingTasks: 32,
} as const

/** 0=empty, 1=number, 2=other nonempty value, 3=spreadsheet error. */
export const CELL_STATE = { empty: 0, number: 1, other: 2, error: 3 } as const
export const FILTER_OPERATOR = {
  eq: 0, ne: 1, lt: 2, le: 3, gt: 4, ge: 5,
  between: 6, notBetween: 7, isEmpty: 8, isNotEmpty: 9, isError: 10,
} as const

export interface NumericColumn { values: Float64Array; states: Uint8Array }
export interface AggregateTask extends NumericColumn { kind: 'aggregate' }
/** Column-major: key k / row r is values[k * rowCount + r]. */
export interface SortTask extends NumericColumn {
  kind: 'sort'; rowCount: number; directions: Uint8Array; nullsLast: boolean
}
export interface FilterTask extends NumericColumn {
  kind: 'filter'; rowCount: number; operators: Uint8Array
  lower: Float64Array; upper: Float64Array; match: 'all' | 'any'
}
export type ComputeTask = AggregateTask | SortTask | FilterTask
export type ComputeErrorCode =
  | 'COMPUTE_INVALID_INPUT' | 'COMPUTE_MEMORY_LIMIT' | 'COMPUTE_UNSUPPORTED_TYPE'
  | 'COMPUTE_WASM_LOAD_FAILED' | 'COMPUTE_WASM_API_MISSING' | 'COMPUTE_PROTOCOL_MISMATCH'
  | 'COMPUTE_FAILED' | 'COMPUTE_CANCELLED' | 'COMPUTE_TIMEOUT' | 'COMPUTE_STALE_RESULT'
  | 'COMPUTE_WORKER_FAILED' | 'COMPUTE_WASM_TRAP' | 'COMPUTE_DISPOSED' | 'COMPUTE_NOT_READY'
  | 'COMPUTE_RESULT_CONSUMED'

export interface ComputeErrorData { code: ComputeErrorCode; message: string; details?: string }
export interface ComputeContext { requestId: string; documentId: string; modelRevision: number }
export interface ComputeProgress extends ComputeContext {
  type: 'progress'; phase: 'initializing' | 'computing' | 'completed'
  ratio: number; processed: number; total: number
  unit: 'values' | 'rows' | 'merge-writes'
}
export type ComputeWorkerProgress = Omit<ComputeProgress, keyof ComputeContext> & { context: WasmTaskControl }
export interface NumericStatistics {
  numericCount: number; nonEmptyCount: number; emptyCount: number
  errorCount: number; invalidNumericCount: number
  sum: number | null; mean: number | null; min: number | null; max: number | null
  populationVariance: number | null; sampleVariance: number | null
  sampleStandardDeviation: number | null; overflowed: boolean
}
export interface AggregateOutput { kind: 'aggregate'; statistics: NumericStatistics }
export interface IndexOutput { kind: 'sort' | 'filter'; indices: Uint32Array }
export type ComputeOutput = AggregateOutput | IndexOutput
export interface ComputeMetrics {
  elapsedMs: number; inputBytes: number; estimatedPeakBytes: number
  boundary: BoundaryCopyStats; transport: WorkerTransportStats
  /** Copy/ownership preparation and scheduling are added by the client. */
  preparedInputCopies?: number; preparedInputBytes?: number; ownershipTransfers?: number
  queueMs?: number; totalMs?: number
}
export interface ComputeResult extends ComputeContext {
  type: 'result'; context: WasmTaskControl; engineVersion: string; output: ComputeOutput; metrics: ComputeMetrics
}
export type ComputeWorkerResult = Omit<ComputeResult, keyof ComputeContext>
export type ComputeWorkerRequest =
  | { type: 'init'; protocolVersion: number; moduleUrl: string }
  | { type: 'run'; context: WasmTaskControl; task: ComputeTask }
  | { type: 'cancel'; context: WasmTaskControl }
export type ComputeWorkerEvent =
  | { type: 'ready'; protocolVersion: number; engineVersion: string }
  | { type: 'init-error'; error: ComputeErrorData }
  | ComputeWorkerProgress | ComputeWorkerResult
  | { type: 'cancelled'; context: WasmTaskControl }
  | { type: 'error'; context: WasmTaskControl; error: ComputeErrorData }

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
export function isContext(value: Record<string, unknown>): value is Record<string, unknown> & ComputeContext {
  return typeof value.requestId === 'string' && value.requestId.trim().length > 0 && typeof value.documentId === 'string' && value.documentId.trim().length > 0
    && Number.isSafeInteger(value.modelRevision) && Number(value.modelRevision) >= 0
}
