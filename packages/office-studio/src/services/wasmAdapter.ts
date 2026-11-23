import type { ImportedOfficeModel, ParseProgress, ParseRequest } from '../parser/contracts.ts'
import { ParserError, serializeParserError } from '../parser/errors.ts'
import {
  createWasmTaskControl, isWasmErrorData, isWasmRecord, isWasmTaskMetrics, matchesWasmTaskControl,
} from '../platform/wasm/task-protocol.ts'
import type { WasmTaskBudget, WasmTaskControl } from '../platform/wasm/task-protocol.ts'
export { ParserError } from '../parser/errors.ts'

export interface ImportOptions {
  signal?: AbortSignal
  onProgress?: (progress: ParseProgress) => void
  timeoutMs?: number
  unitId?: string
  baseRevision?: number
  budget?: Partial<WasmTaskBudget>
  isCurrent?: (context: WasmTaskControl) => boolean
}
interface Pending {
  context: WasmTaskControl
  inputBytes: number
  fileName: string
  startedAt: number
  options: ImportOptions
  resolve(result: ImportedOfficeModel): void
  reject(error: ParserError): void
  release(): void
}
let worker: Worker | undefined
let pending: Pending | undefined

function parserFailure(error: unknown): ParserError {
  if (error instanceof ParserError) return error
  const data = serializeParserError(error)
  return new ParserError(data.code, data.message, [{ code: data.code, severity: 'error',
    message: data.message, details: data.details ?? undefined }])
}
function terminate(reason: ParserError, expected?: Pending): void {
  if (expected && pending !== expected) return
  const task = pending
  pending = undefined
  if (worker) { worker.onmessage = null; worker.onerror = null; worker.onmessageerror = null; worker.terminate() }
  worker = undefined
  task?.release()
  task?.reject(reason)
}
function isCurrent(task: Pending): boolean {
  return pending === task && !task.options.signal?.aborted && (task.options.isCurrent?.(task.context) ?? true)
}
function isProgress(value: unknown): value is ParseProgress {
  return isWasmRecord(value) && typeof value.message === 'string'
    && ['loading', 'reading', 'parsing', 'normalizing', 'ready'].some((stage) => stage === value.stage)
}
function isImportedModel(value: unknown): value is ImportedOfficeModel {
  if (!isWasmRecord(value) || (value.kind !== 'workbook' && value.kind !== 'document')
    || !isWasmRecord(value.model) || value.model.kind !== value.kind || typeof value.model.id !== 'string'
    || typeof value.model.name !== 'string' || typeof value.format !== 'string' || typeof value.fileName !== 'string'
    || typeof value.parserVersion !== 'string' || !Number.isSafeInteger(value.size) || Number(value.size) < 0
    || !Array.isArray(value.diagnostics) || !isWasmRecord(value.stats) || !Array.isArray(value.capabilities)
    || !value.capabilities.every((item) => typeof item === 'string')) return false
  return value.kind === 'workbook' ? Array.isArray(value.model.sheets) : Array.isArray(value.model.blocks)
}
function acceptMessage(data: unknown): void {
  const task = pending
  if (!task) return
  try {
    if (!isWasmRecord(data)) throw new ParserError('WASM_PROTOCOL_MISMATCH', '解析 Worker 消息结构无效')
    if (data.type === 'protocol-error') {
      if (!isWasmErrorData(data.error)) throw new ParserError('WASM_PROTOCOL_MISMATCH', '解析 Worker 错误结构无效')
      terminate(parserFailure(data.error), task); return
    }
    // Reused Worker instances can deliver late events for a previous request.
    if (isWasmRecord(data.context) && typeof data.context.requestId === 'string'
      && data.context.requestId !== task.context.requestId) return
    if (!matchesWasmTaskControl(data.context, task.context)) throw new ParserError('WASM_PROTOCOL_MISMATCH', '解析任务身份或版本不匹配')
    if (!isCurrent(task)) { terminate(new ParserError('PARSE_STALE_RESULT', '导入上下文已变化，旧解析结果已丢弃'), task); return }
    if (data.type === 'progress') {
      if (!isProgress(data.progress)) throw new ParserError('WASM_PROTOCOL_MISMATCH', '解析阶段消息无效')
      task.options.onProgress?.(data.progress); return
    }
    if (data.type === 'error') {
      if (!isWasmErrorData(data.error)) throw new ParserError('WASM_PROTOCOL_MISMATCH', '解析错误结构无效')
      const error = parserFailure(data.error)
      if (error.code === 'WASM_TRAP') { terminate(error, task); return }
      pending = undefined; task.release(); task.reject(error); return
    }
    if (data.type !== 'success' || !isImportedModel(data.result) || !isWasmTaskMetrics(data.metrics)
      || data.result.model.id !== task.context.unitId || data.result.size !== task.inputBytes || data.result.fileName !== task.fileName
      || data.metrics.boundary.inputBytes !== task.inputBytes || data.metrics.transport.inputBytes !== task.inputBytes
      || data.metrics.transport.inputTransfers !== 1 || data.metrics.transport.outputTransfers !== 0
      || data.metrics.transport.structuredResultMessages !== 1) {
      throw new ParserError('WASM_PROTOCOL_MISMATCH', '解析结果或任务计量不匹配')
    }
    const result: ImportedOfficeModel = { ...data.result,
      task: { context: task.context, metrics: { ...data.metrics, totalMs: performance.now() - task.startedAt } } }
    pending = undefined; task.release(); task.resolve(result)
  } catch (error) { terminate(parserFailure(error), task) }
}
function ensureWorker(): Worker {
  if (worker) return worker
  const instance = new Worker(new URL('../parser/parser.worker.ts', import.meta.url), { type: 'module', name: 'office-parser' })
  worker = instance
  instance.onmessage = ({ data }: MessageEvent<unknown>) => { if (worker === instance) acceptMessage(data) }
  instance.onerror = (event) => {
    if (worker === instance) terminate(new ParserError('WORKER_FAILED', '解析任务异常：' + event.message))
  }
  instance.onmessageerror = () => {
    if (worker === instance) terminate(new ParserError('WORKER_MESSAGE_FAILED', '解析结果无法传输'))
  }
  return instance
}

export async function parseOfficeFile(file: File, options: ImportOptions = {}): Promise<ImportedOfficeModel> {
  if (options.signal?.aborted) throw new ParserError('CANCELLED', '已取消解析')
  if (!Number.isSafeInteger(file.size) || file.size <= 0) throw new ParserError('EMPTY_FILE', '文件为空或大小无效')
  const timeoutMs = options.timeoutMs ?? 120_000
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0 || timeoutMs > 2_147_483_647) throw new ParserError('PARSE_INVALID_TIMEOUT', '解析超时时间无效')
  const requestId = crypto.randomUUID()
  let context: WasmTaskControl
  try { context = createWasmTaskControl('parse', requestId, options.unitId ?? 'import-' + requestId, options.baseRevision ?? 0, options.budget) }
  catch (error) { throw parserFailure(error) }
  if (file.size > context.budget.maxInputBytes || file.size > context.budget.maxWorkingBytes) {
    throw new ParserError('WASM_MEMORY_LIMIT', '文件超过解析内存预算')
  }
  if (options.isCurrent && !options.isCurrent(context)) throw new ParserError('PARSE_STALE_RESULT', '导入上下文已失效')
  if (pending) terminate(new ParserError('SUPERSEDED', '已被新的解析任务替代'))
  return new Promise<ImportedOfficeModel>((resolve, reject) => {
    const task: Pending = {
      context, inputBytes: file.size, fileName: file.name, startedAt: performance.now(), options: { ...options }, resolve, reject,
      release: () => { clearTimeout(timeout); options.signal?.removeEventListener('abort', abort) },
    }
    const abort = () => terminate(new ParserError('CANCELLED', '已取消解析'), task)
    const timeout = setTimeout(() => terminate(new ParserError('PARSE_TIMEOUT', '解析超时，请检查文件或重试'), task), timeoutMs)
    pending = task
    options.signal?.addEventListener('abort', abort, { once: true })
    if (options.signal?.aborted) { abort(); return }
    const start = async () => {
      try {
        task.options.onProgress?.({ stage: 'reading', message: '正在读取 ' + file.name })
        if (pending !== task) return
        if (!isCurrent(task)) { terminate(new ParserError('PARSE_STALE_RESULT', '导入上下文已变化'), task); return }
        const bytes = await file.arrayBuffer()
        if (pending !== task) return
        if (!isCurrent(task)) { terminate(new ParserError('PARSE_STALE_RESULT', '导入上下文已变化'), task); return }
        if (bytes.byteLength !== file.size) throw new ParserError('FILE_READ_FAILED', '读取的文件长度与声明大小不一致')
        const request: ParseRequest = { type: 'parse', context, bytes, fileName: file.name,
          moduleUrl: new URL(import.meta.env.BASE_URL + 'wasm/office_parser.js', window.location.href).href }
        ensureWorker().postMessage(request, [bytes])
      } catch (error) { terminate(parserFailure(error), task) }
    }
    void start()
  })
}
export function resetParser(): void { terminate(new ParserError('CANCELLED', '解析器已重置')) }
