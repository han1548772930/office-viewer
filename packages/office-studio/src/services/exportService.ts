import type { DocumentModel, OfficeKind, WorkbookModel } from '../domain/model'
import type { ExportOutcome, ExportProgress, ExportRequest, ExportResponse, ExportWarning } from '../parser/export-contracts'

interface PendingExport {
  resolve: (outcome: ExportOutcome) => void
  reject: (error: Error) => void
  progress?: (progress: ExportProgress) => void
  release: () => void
}

let worker: Worker | undefined
let pending: PendingExport | undefined
let activeRequest: string | undefined

function fail(error: Error): void {
  const task = pending
  pending = undefined
  activeRequest = undefined
  worker?.terminate()
  worker = undefined
  task?.release()
  task?.reject(error)
}

function ensureWorker(): Worker {
  if (worker) return worker
  worker = new Worker(new URL('../parser/export.worker.ts', import.meta.url), { type: 'module', name: 'office-writer' })
  worker.onmessage = ({ data }: MessageEvent<ExportResponse>) => {
    if (!pending || data.requestId !== activeRequest) return
    if (data.type === 'progress') { pending.progress?.(data.progress); return }
    const task = pending
    pending = undefined
    activeRequest = undefined
    task.release()
    if (data.type === 'success') task.resolve({ bytes: data.bytes, warnings: parseWarnings(data.warnings) })
    else task.reject(new Error(`Office 文件导出失败：${data.error.message}`))
  }
  worker.onerror = (error) => fail(new Error(`Office 写出任务异常：${error.message}`))
  worker.onmessageerror = () => fail(new Error('Office 写出结果无法传输'))
  return worker
}

export interface ExportOptions { onProgress?: (progress: ExportProgress) => void; timeoutMs?: number; signal?: AbortSignal }

export class OfficeExportError extends Error {
  readonly code: 'CANCELLED' | 'EXPORT_FAILED'

  constructor(code: 'CANCELLED' | 'EXPORT_FAILED', message: string) {
    super(message)
    this.code = code
    this.name = 'OfficeExportError'
  }
}

export async function exportOfficeModel(kind: OfficeKind, model: WorkbookModel | DocumentModel, options: ExportOptions = {}): Promise<ExportOutcome> {
  if (model.kind !== kind) throw new Error('导出类型与当前模型不一致')
  if (options.signal?.aborted) throw new OfficeExportError('CANCELLED', 'Office 文件导出已取消')
  if (pending) fail(new OfficeExportError('CANCELLED', '已被新的导出任务替代'))
  const requestId = crypto.randomUUID()
  activeRequest = requestId
  const data: ExportRequest = {
    requestId, kind, data: JSON.stringify(model),
    baseUrl: new URL(import.meta.env.BASE_URL, window.location.href).href,
  }
  const writer = ensureWorker()
  return new Promise((resolve, reject) => {
    const timeout = window.setTimeout(() => fail(new OfficeExportError('EXPORT_FAILED', 'Office 文件导出超时')), options.timeoutMs ?? 120_000)
    const abort = () => fail(new OfficeExportError('CANCELLED', 'Office 文件导出已取消'))
    pending = { resolve, reject, progress: options.onProgress, release: () => {
      window.clearTimeout(timeout)
      options.signal?.removeEventListener('abort', abort)
    } }
    options.signal?.addEventListener('abort', abort, { once: true })
    if (options.signal?.aborted) { abort(); return }
    try { writer.postMessage(data) } catch (error: unknown) { fail(error instanceof Error ? error : new Error(String(error))) }
  })
}

function parseWarnings(value: string): ExportWarning[] {
  try {
    const parsed: unknown = JSON.parse(value)
    if (!Array.isArray(parsed)) return []
    return parsed.filter((warning): warning is ExportWarning =>
      typeof warning === 'object' && warning !== null &&
      'code' in warning && typeof warning.code === 'string' &&
      'message' in warning && typeof warning.message === 'string')
  } catch {
    return []
  }
}

export function officeDownloadName(model: WorkbookModel | DocumentModel, kind: OfficeKind): string {
  const base = model.name.replace(/\.[^.]+$/, '') || (kind === 'workbook' ? '未命名工作簿' : '未命名文档')
  return `${base}.${kind === 'workbook' ? 'xlsx' : 'docx'}`
}

export function downloadOfficeBytes(bytes: Uint8Array, fileName: string, kind: OfficeKind): void {
  const mime = kind === 'workbook'
    ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  const blobBytes = new Uint8Array(bytes.byteLength)
  blobBytes.set(bytes)
  const blob = new Blob([blobBytes.buffer], { type: mime })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}
