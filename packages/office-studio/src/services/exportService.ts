import type { DocumentModel, OfficeKind, WorkbookModel } from '../domain/model.ts'
import { isExportWarnings } from '../parser/export-contracts.ts'
import type { ExportOutcome, ExportProgress, ExportRequest } from '../parser/export-contracts.ts'
import { assertWorkbookConditionalFormatBudget } from '../features/conditional-formatting/conditional-format-workbook-budget.ts'
import {
  createWasmTaskControl, isWasmErrorData, isWasmRecord, isWasmTaskMetrics, matchesWasmTaskControl, serializeWasmError,
} from '../platform/wasm/task-protocol.ts'
import type { WasmTaskBudget, WasmTaskControl } from '../platform/wasm/task-protocol.ts'

export interface ExportOptions {
  onProgress?: (progress: ExportProgress) => void
  timeoutMs?: number
  signal?: AbortSignal
  modelRevision?: number
  budget?: Partial<WasmTaskBudget>
  isCurrent?: (context: WasmTaskControl) => boolean
}
export class OfficeExportError extends Error {
  readonly code: string
  readonly details?: string
  constructor(code: string, message: string, details?: string) {
    super(message); this.code = code; this.details = details; this.name = 'OfficeExportError'
  }
}
interface PendingExport {
  context: WasmTaskControl
  options: ExportOptions
  startedAt: number
  resolve(outcome: ExportOutcome): void
  reject(error: OfficeExportError): void
  release(): void
}
let worker: Worker | undefined
let pending: PendingExport | undefined

function exportFailure(error: unknown): OfficeExportError {
  if (error instanceof OfficeExportError) return error
  const data = serializeWasmError(error, { code: 'EXPORT_FAILED', message: 'Office 文件导出失败' })
  return new OfficeExportError(data.code, data.message, data.details ?? undefined)
}
function fail(error: OfficeExportError, expected?: PendingExport): void {
  if (expected && pending !== expected) return
  const task = pending
  pending = undefined
  if (worker) { worker.onmessage = null; worker.onerror = null; worker.onmessageerror = null; worker.terminate() }
  worker = undefined
  task?.release(); task?.reject(error)
}
function current(task: PendingExport): boolean {
  return pending === task && !task.options.signal?.aborted && (task.options.isCurrent?.(task.context) ?? true)
}
function acceptMessage(data: unknown): void {
  const task = pending
  if (!task) return
  try {
    if (!isWasmRecord(data)) throw new OfficeExportError('WASM_PROTOCOL_MISMATCH', '写出 Worker 消息结构无效')
    if (data.type === 'protocol-error') {
      if (!isWasmErrorData(data.error)) throw new OfficeExportError('WASM_PROTOCOL_MISMATCH', '写出错误结构无效')
      fail(exportFailure(data.error), task); return
    }
    if (isWasmRecord(data.context) && typeof data.context.requestId === 'string'
      && data.context.requestId !== task.context.requestId) return
    if (!matchesWasmTaskControl(data.context, task.context)) throw new OfficeExportError('WASM_PROTOCOL_MISMATCH', '写出任务身份或版本不匹配')
    if (!current(task)) { fail(new OfficeExportError('EXPORT_STALE_RESULT', '文档已变化，旧导出结果已丢弃'), task); return }
    if (data.type === 'progress') {
      if (!isWasmRecord(data.progress) || (data.progress.stage !== 'loading' && data.progress.stage !== 'writing')
        || typeof data.progress.message !== 'string') throw new OfficeExportError('WASM_PROTOCOL_MISMATCH', '写出阶段消息无效')
      task.options.onProgress?.({ stage: data.progress.stage, message: data.progress.message }); return
    }
    if (data.type === 'error') {
      if (!isWasmErrorData(data.error)) throw new OfficeExportError('WASM_PROTOCOL_MISMATCH', '写出错误结构无效')
      const error = exportFailure(data.error)
      if (error.code === 'WASM_TRAP') { fail(error, task); return }
      pending = undefined; task.release(); task.reject(error); return
    }
    if (data.type !== 'success' || !(data.bytes instanceof Uint8Array) || !data.bytes.byteLength
      || !(data.bytes.buffer instanceof ArrayBuffer) || data.bytes.byteOffset !== 0
      || data.bytes.byteLength !== data.bytes.buffer.byteLength || data.bytes.byteLength > task.context.budget.maxOutputBytes
      || !isExportWarnings(data.warnings) || !isWasmTaskMetrics(data.metrics)
      || data.metrics.boundary.outputCopies !== 1 || data.metrics.boundary.outputBytes !== data.bytes.byteLength
      || data.metrics.transport.outputTransfers !== 1 || data.metrics.transport.outputBytes !== data.bytes.byteLength
      || data.metrics.transport.inputTransfers !== 0) {
      throw new OfficeExportError('WASM_PROTOCOL_MISMATCH', '写出结果、警告或任务计量不匹配')
    }
    const outcome: ExportOutcome = { bytes: data.bytes, warnings: data.warnings, context: task.context,
      metrics: { ...data.metrics, totalMs: performance.now() - task.startedAt } }
    pending = undefined; task.release(); task.resolve(outcome)
  } catch (error) { fail(exportFailure(error), task) }
}
function ensureWorker(): Worker {
  if (worker) return worker
  const instance = new Worker(new URL('../parser/export.worker.ts', import.meta.url), { type: 'module', name: 'office-writer' })
  worker = instance
  instance.onmessage = ({ data }: MessageEvent<unknown>) => { if (worker === instance) acceptMessage(data) }
  instance.onerror = (event) => { if (worker === instance) fail(new OfficeExportError('EXPORT_WORKER_FAILED', 'Office 写出任务异常：' + event.message)) }
  instance.onmessageerror = () => { if (worker === instance) fail(new OfficeExportError('EXPORT_MESSAGE_FAILED', 'Office 写出结果无法传输')) }
  return instance
}

export async function exportOfficeModel(kind: OfficeKind, model: WorkbookModel | DocumentModel, options: ExportOptions = {}): Promise<ExportOutcome> {
  if (model.kind !== kind) throw new OfficeExportError('EXPORT_INVALID_INPUT', '导出类型与当前模型不一致')
  if (options.signal?.aborted) throw new OfficeExportError('CANCELLED', 'Office 文件导出已取消')
  const timeoutMs = options.timeoutMs ?? 120_000
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0 || timeoutMs > 2_147_483_647) throw new OfficeExportError('EXPORT_INVALID_TIMEOUT', '写出超时时间无效')
  if (model.kind === 'workbook') assertWorkbookConditionalFormatBudget(model.sheets)
  const requestId = crypto.randomUUID()
  let context: WasmTaskControl
  try { context = createWasmTaskControl('export', requestId, model.id, options.modelRevision ?? 0, options.budget) }
  catch (error) { throw exportFailure(error) }
  if (options.isCurrent && !options.isCurrent(context)) throw new OfficeExportError('EXPORT_STALE_RESULT', '文档版本已失效')
  const startedAt = performance.now()
  const snapshot = JSON.stringify(model)
  if (snapshot.length > context.budget.maxInputBytes || snapshot.length > context.budget.maxWorkingBytes) throw new OfficeExportError('WASM_MEMORY_LIMIT', 'Office 快照超过写出内存预算')
  if (options.isCurrent && !options.isCurrent(context)) throw new OfficeExportError('EXPORT_STALE_RESULT', '准备快照期间文档已变化')
  if (options.signal?.aborted) throw new OfficeExportError('CANCELLED', 'Office 文件导出已取消')
  if (pending) fail(new OfficeExportError('CANCELLED', '已被新的导出任务替代'))
  return new Promise<ExportOutcome>((resolve, reject) => {
    const task: PendingExport = { context, options: { ...options }, startedAt, resolve, reject, release: () => {
      clearTimeout(timeout); options.signal?.removeEventListener('abort', abort)
    } }
    const abort = () => fail(new OfficeExportError('CANCELLED', 'Office 文件导出已取消'), task)
    const remaining = timeoutMs - (performance.now() - startedAt)
    const timeout = setTimeout(() => fail(new OfficeExportError('EXPORT_TIMEOUT', 'Office 文件导出超时'), task), Math.max(0, remaining))
    pending = task
    options.signal?.addEventListener('abort', abort, { once: true })
    if (options.signal?.aborted) { abort(); return }
    if (remaining <= 0) { fail(new OfficeExportError('EXPORT_TIMEOUT', 'Office 文件导出超时'), task); return }
    try {
      const request: ExportRequest = { type: 'export', context, kind, data: snapshot,
        moduleUrl: new URL(import.meta.env.BASE_URL + 'wasm-writer/office_writer.js', window.location.href).href }
      ensureWorker().postMessage(request)
    } catch (error) { fail(exportFailure(error), task) }
  })
}
export function resetOfficeWriter(): void { fail(new OfficeExportError('CANCELLED', 'Office 写出器已重置')) }

export function officeDownloadName(model: WorkbookModel | DocumentModel, kind: OfficeKind): string {
  const base = model.name.replace(/\.[^.]+$/, '') || (kind === 'workbook' ? '未命名工作簿' : '未命名文档')
  return base + (kind === 'workbook' ? '.xlsx' : '.docx')
}
export function downloadOfficeBytes(bytes: Uint8Array, fileName: string, kind: OfficeKind): void {
  const mime = kind === 'workbook'
    ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  let buffer: ArrayBuffer
  if (bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) buffer = bytes.buffer
  else { const owned = new Uint8Array(bytes.byteLength); owned.set(bytes); buffer = owned.buffer }
  // Blob takes its own immutable snapshot; it is a separate browser boundary.
  const blob = new Blob([buffer], { type: mime })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url; anchor.download = fileName; anchor.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}
