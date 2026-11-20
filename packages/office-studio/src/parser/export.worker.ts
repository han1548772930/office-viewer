import type { ExportRequest, ExportResponse } from './export-contracts'

interface WriterResult { bytes: Uint8Array; warnings: string; free: () => void }
interface WriterModule { default: () => Promise<unknown>; write_office: (kind: string, data: string) => WriterResult }

const scope = self as unknown as {
  postMessage: (message: ExportResponse, transfer?: Transferable[]) => void
  onmessage: ((event: MessageEvent<ExportRequest>) => void) | null
}
let loading: Promise<WriterModule> | undefined

scope.onmessage = async ({ data }) => {
  const { requestId, kind, data: modelData, baseUrl } = data
  const progress = (stage: 'loading' | 'writing', message: string) =>
    scope.postMessage({ requestId, type: 'progress', progress: { stage, message } })
  try {
    progress('loading', '正在加载 Office 写出器')
    loading ??= import(/* @vite-ignore */ `${baseUrl}wasm-writer/office_writer.js`).then(async (module: WriterModule) => {
      if (typeof module.default !== 'function' || typeof module.write_office !== 'function') throw new Error('Office writer WASM API 不完整，请先构建 office-writer')
      await module.default()
      return module
    }).catch((error: unknown) => { loading = undefined; throw error })
    const writer = await loading
    progress('writing', kind === 'workbook' ? '正在生成 XLSX 工作簿' : '正在生成 DOCX 文档')
    const result = writer.write_office(kind, modelData)
    try {
      const bytes = result.bytes
      if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0) throw new Error('Office writer 没有生成文件内容')
      scope.postMessage({ requestId, type: 'success', bytes, warnings: result.warnings }, [bytes.buffer])
    } finally { result.free() }
  } catch (error: unknown) {
    let message = error instanceof Error ? error.message : String(error)
    try {
      const parsed: unknown = JSON.parse(message)
      if (parsed && typeof parsed === 'object' && 'error' in parsed) {
        const detail = parsed.error
        if (detail && typeof detail === 'object' && 'message' in detail && typeof detail.message === 'string') message = detail.message
      }
    } catch { /* retain normal wasm-bindgen and worker errors */ }
    scope.postMessage({ requestId, type: 'error', error: { code: 'EXPORT_FAILED', message } })
  }
}
