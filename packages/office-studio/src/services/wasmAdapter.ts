import type { ImportedOfficeModel, ParseProgress, ParseResponse } from '../parser/contracts'
import { ParserError } from '../parser/errors'
export { ParserError } from '../parser/errors'

interface ImportOptions { signal?: AbortSignal; onProgress?: (progress: ParseProgress) => void; timeoutMs?: number }
interface Pending {
  resolve: (result: ImportedOfficeModel) => void
  reject: (error: ParserError) => void
  progress?: (progress: ParseProgress) => void
  release: () => void
}
let worker: Worker | undefined
let pending: Pending | undefined
let activeRequest: string | undefined
function terminate(reason: ParserError): void {
  const task = pending
  pending = undefined
  activeRequest = undefined
  worker?.terminate()
  worker = undefined
  task?.release()
  task?.reject(reason)
}
function ensureWorker(): Worker {
  if (worker) return worker
  worker = new Worker(new URL('../parser/parser.worker.ts', import.meta.url), { type: 'module', name: 'office-parser' })
  worker.onmessage = ({ data }: MessageEvent<ParseResponse>) => {
    if (!pending || data.requestId !== activeRequest) return
    if (data.type === 'progress') { pending.progress?.(data.progress); return }
    const task = pending
    pending = undefined
    activeRequest = undefined
    task.release()
    if (data.type === 'success') task.resolve(data.result)
    else task.reject(new ParserError(data.error.code, '文件解析失败：'+data.error.message, [{
      code: data.error.code, severity: 'error', message: data.error.message, details: data.error.details ?? undefined,
    }]))
  }
  worker.onerror = (error) => terminate(new ParserError('WORKER_FAILED', '解析任务异常：'+error.message))
  worker.onmessageerror = () => terminate(new ParserError('WORKER_MESSAGE_FAILED', '解析结果无法传输'))
  return worker
}
export async function parseOfficeFile(file: File, options: ImportOptions = {}): Promise<ImportedOfficeModel> {
  if (options.signal?.aborted) throw new ParserError('CANCELLED', '已取消解析')
  if (!file.size) throw new ParserError('EMPTY_FILE', '文件为空')
  terminate(new ParserError('SUPERSEDED', '已被新的解析任务替代'))
  const requestId = crypto.randomUUID()
  activeRequest = requestId
  options.onProgress?.({ stage: 'reading', message: '正在读取 '+file.name })
  const bytes = await file.arrayBuffer()
  if (activeRequest !== requestId || options.signal?.aborted) throw new ParserError('CANCELLED', '已取消解析')
  const parserWorker = ensureWorker()
  return new Promise((resolve, reject) => {
    const abort = () => terminate(new ParserError('CANCELLED', '已取消解析'))
    const timeout = setTimeout(() => terminate(new ParserError('PARSE_TIMEOUT', '解析超时，请检查文件或重试')), options.timeoutMs ?? 120_000)
    pending = { resolve, reject, progress: options.onProgress, release: () => {
      clearTimeout(timeout)
      options.signal?.removeEventListener('abort', abort)
    } }
    options.signal?.addEventListener('abort', abort, { once: true })
    const baseUrl = new URL(import.meta.env.BASE_URL, window.location.href).href
    try { parserWorker.postMessage({ requestId, fileName: file.name, baseUrl, bytes }, [bytes]) }
    catch { terminate(new ParserError('WORKER_MESSAGE_FAILED', '文件无法传输到解析任务')) }
  })
}
export function resetParser(): void { terminate(new ParserError('CANCELLED', '解析器已重置')) }
