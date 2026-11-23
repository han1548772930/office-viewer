/** One protocol shared by Rust exports and every Office Worker.
 * Version changes are coordinated source changes; older artifacts are rejected.
 */
export const WASM_PROTOCOL_VERSION = 2 as const
export const WASM_TASK_KINDS = ['parse', 'export', 'aggregate', 'sort', 'filter', 'chartStatistics', 'layout'] as const
export type WasmTaskKind = typeof WASM_TASK_KINDS[number]

export interface WasmTaskBudget {
  maxInputBytes: number
  maxOutputBytes: number
  maxWorkingBytes: number
}
export const DEFAULT_WASM_TASK_BUDGET: Readonly<WasmTaskBudget> = {
  maxInputBytes: 256 * 1024 * 1024,
  maxOutputBytes: 256 * 1024 * 1024,
  maxWorkingBytes: 768 * 1024 * 1024,
}
export interface WasmTaskControl {
  protocolVersion: typeof WASM_PROTOCOL_VERSION
  requestId: string
  unitId: string
  baseRevision: number
  kind: WasmTaskKind
  budget: WasmTaskBudget
}

/** Measured JS/WASM boundaries only. Object conversion, transport and GPU
 * uploads do not imply end-to-end zero-copy or bound the browser heap.
 */
export interface BoundaryCopyStats {
  inputCopies: number
  inputBytes: number
  outputCopies: number
  outputBytes: number
  rustBufferClones: number
  controlDtoConversions: number
  jsonConversions: number
}
export interface WorkerTransportStats {
  inputTransfers: number
  inputBytes: number
  outputTransfers: number
  outputBytes: number
  /** Large structured objects have no claimed binary byte measurement. */
  structuredResultMessages: number
}
export interface WasmTaskMetrics {
  elapsedMs: number
  boundary: BoundaryCopyStats
  transport: WorkerTransportStats
  totalMs?: number
}
export interface WasmErrorData { code: string; message: string; details?: string | null }

export function isWasmRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
export function isUnsignedCount(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}
const isU32 = (value: unknown): value is number => isUnsignedCount(value) && value <= 0xffff_ffff

export function isWasmTaskBudget(value: unknown): value is WasmTaskBudget {
  return isWasmRecord(value) && isU32(value.maxInputBytes) && isU32(value.maxOutputBytes)
    && isU32(value.maxWorkingBytes) && value.maxWorkingBytes > 0
    && Object.keys(value).every((key) => ['maxInputBytes', 'maxOutputBytes', 'maxWorkingBytes'].includes(key))
}
export function isWasmTaskControl(value: unknown): value is WasmTaskControl {
  return isWasmRecord(value) && value.protocolVersion === WASM_PROTOCOL_VERSION
    && typeof value.requestId === 'string' && value.requestId.trim().length > 0
    && typeof value.unitId === 'string' && value.unitId.trim().length > 0
    && isUnsignedCount(value.baseRevision)
    && WASM_TASK_KINDS.some((kind) => kind === value.kind) && isWasmTaskBudget(value.budget)
    && Object.keys(value).every((key) => ['protocolVersion', 'requestId', 'unitId', 'baseRevision', 'kind', 'budget'].includes(key))
}
export function matchesWasmTaskControl(value: unknown, expected: WasmTaskControl): value is WasmTaskControl {
  return isWasmTaskControl(value) && value.requestId === expected.requestId && value.unitId === expected.unitId
    && value.baseRevision === expected.baseRevision && value.kind === expected.kind
    && value.budget.maxInputBytes === expected.budget.maxInputBytes
    && value.budget.maxOutputBytes === expected.budget.maxOutputBytes
    && value.budget.maxWorkingBytes === expected.budget.maxWorkingBytes
}
export function createWasmTaskControl(
  kind: WasmTaskKind, requestId: string, unitId: string, baseRevision: number,
  budget: Partial<WasmTaskBudget> = {},
): WasmTaskControl {
  const context = { protocolVersion: WASM_PROTOCOL_VERSION, requestId, unitId, baseRevision, kind,
    budget: { ...DEFAULT_WASM_TASK_BUDGET, ...budget } }
  if (!isWasmTaskControl(context)) throw new WasmProtocolError('WASM_CONTROL_INVALID', '任务身份、版本或内存预算无效')
  return context
}
export function requireWasmTaskControl(value: unknown, expected: WasmTaskKind): WasmTaskControl {
  if (!isWasmTaskControl(value) || value.kind !== expected) {
    throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', 'Worker 任务协议或计算域不匹配')
  }
  return value
}

const COPY_KEYS = ['inputCopies', 'inputBytes', 'outputCopies', 'outputBytes', 'rustBufferClones', 'controlDtoConversions', 'jsonConversions'] as const
export function isBoundaryCopyStats(value: unknown): value is BoundaryCopyStats {
  return isWasmRecord(value) && COPY_KEYS.every((key) => isUnsignedCount(value[key]))
    && Object.keys(value).every((key) => COPY_KEYS.some((known) => known === key))
}
export function readBoundaryCopyStats(value: unknown): BoundaryCopyStats {
  if (!isBoundaryCopyStats(value)) throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', 'WASM 复制统计结构不合法')
  return value
}
export function isWasmTaskMetrics(value: unknown): value is WasmTaskMetrics {
  if (!isWasmRecord(value) || typeof value.elapsedMs !== 'number' || !Number.isFinite(value.elapsedMs)
    || value.elapsedMs < 0 || !isBoundaryCopyStats(value.boundary) || !isWasmRecord(value.transport)) return false
  const transport = value.transport
  return ['inputTransfers', 'inputBytes', 'outputTransfers', 'outputBytes', 'structuredResultMessages']
    .every((key) => isUnsignedCount(transport[key]))
}
export function isWasmErrorData(value: unknown): value is WasmErrorData {
  return isWasmRecord(value) && typeof value.code === 'string' && value.code.trim().length > 0
    && typeof value.message === 'string' && (value.details === undefined || value.details === null || typeof value.details === 'string')
}
export class WasmProtocolError extends Error {
  readonly code: string
  readonly details?: string
  constructor(code: string, message: string, details?: string) {
    super(message); this.name = 'WasmProtocolError'; this.code = code; this.details = details
  }
}
export function serializeWasmError(error: unknown, fallback: WasmErrorData): WasmErrorData {
  // Rust rejects with a structured { ok: false, error } DTO. There is no
  // JSON-string error adapter for old generated bindings.
  const structured = isWasmRecord(error) && error.ok === false ? error.error : error
  if (isWasmErrorData(structured)) return { code: structured.code, message: structured.message,
    ...(structured.details == null ? {} : { details: structured.details }) }
  if (typeof WebAssembly !== 'undefined' && error instanceof WebAssembly.RuntimeError) {
    return { code: 'WASM_TRAP', message: 'WASM 执行异常', details: error.message }
  }
  if (error instanceof RangeError) return { code: 'WASM_MEMORY_LIMIT', message: 'WASM 内存分配失败', details: error.message }
  return { ...fallback, details: error instanceof Error ? error.message : String(error) }
}
export function requireMatchingWasmResult(value: unknown, expected: WasmTaskControl): Record<string, unknown> {
  if (!isWasmRecord(value) || !matchesWasmTaskControl(value.context, expected)) {
    throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', 'WASM 返回了不匹配的任务上下文')
  }
  return value
}
export function requireIndependentOutput(view: Uint8Array | Uint32Array, memory: WebAssembly.Memory): ArrayBuffer {
  if (!(view.buffer instanceof ArrayBuffer) || view.buffer === memory.buffer
    || view.byteOffset !== 0 || view.byteLength !== view.buffer.byteLength) {
    throw new WasmProtocolError('WASM_OUTPUT_ALIAS', '结果必须使用独立且只覆盖有效区间的 ArrayBuffer')
  }
  return view.buffer
}
export function requireWasmMemory(value: unknown): WebAssembly.Memory {
  if (!isWasmRecord(value) || !(value.memory instanceof WebAssembly.Memory)) {
    throw new WasmProtocolError('WASM_API_MISSING', 'WASM 初始化结果缺少内存句柄')
  }
  return value.memory
}
export function assertWasmProtocol(version: number, domain: string): void {
  if (version !== WASM_PROTOCOL_VERSION) throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH',
    domain + ' WASM 产物与当前源码协议不匹配，请重新构建')
}
