import { isExportWarnings } from './export-contracts.ts'
import type { ExportRequest, ExportResponse } from './export-contracts.ts'
import { isWriterWasmResult, loadWriterWasm } from './wasm-modules.ts'
import {
  isWasmRecord, readBoundaryCopyStats, requireIndependentOutput, requireMatchingWasmResult,
  requireWasmTaskControl, serializeWasmError, WasmProtocolError,
} from '../platform/wasm/task-protocol.ts'
import type { WasmTaskControl } from '../platform/wasm/task-protocol.ts'

const scope = self as unknown as {
  postMessage(message: ExportResponse, transfer?: Transferable[]): void
  onmessage: ((event: MessageEvent<unknown>) => void) | null
}
let loading: ReturnType<typeof loadWriterWasm> | undefined
let moduleUrl: string | undefined
let active: WasmTaskControl | undefined
const failure = (error: unknown) => serializeWasmError(error, { code: 'EXPORT_FAILED', message: 'Office 文件写出失败' })

async function write(request: ExportRequest): Promise<void> {
  const context = request.context
  if (active) {
    scope.postMessage({ context, type: 'error', error: { code: 'EXPORT_BUSY', message: '写出器正在执行其他任务' } })
    return
  }
  active = context
  const started = performance.now()
  const progress = (stage: 'loading' | 'writing', message: string) =>
    scope.postMessage({ context, type: 'progress', progress: { stage, message } })
  try {
    progress('loading', '正在加载 Office 写出器')
    if (moduleUrl !== undefined && moduleUrl !== request.moduleUrl) {
      throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '同一写出 Worker 不能切换 WASM 模块')
    }
    moduleUrl = request.moduleUrl
    loading ??= loadWriterWasm(request.moduleUrl).catch((error: unknown) => { loading = undefined; throw error })
    const writer = await loading
    progress('writing', request.kind === 'workbook' ? '正在生成 XLSX 工作簿' : '正在生成 DOCX 文档')
    const result: unknown = await writer.module.write_office(request.kind, request.data, context)
    if (!isWriterWasmResult(result)) throw new WasmProtocolError('WASM_API_MISSING', '写出器返回的结果缺少协议 2 所有权接口')
    try {
      const bytes = result.copy_bytes()
      if (!(bytes instanceof Uint8Array) || !bytes.byteLength) throw new WasmProtocolError('EXPORT_EMPTY_RESULT', '写出器没有生成文件内容')
      const buffer = requireIndependentOutput(bytes, writer.memory)
      // copy_bytes owns its temporary view internally. Metadata may allocate
      // safely because only an independent JS buffer remains here.
      const metadata = requireMatchingWasmResult(result.metadata(), context)
      const boundary = readBoundaryCopyStats(metadata.transfer)
      if (metadata.byteLength !== bytes.byteLength || bytes.byteLength > context.budget.maxOutputBytes
        || boundary.outputCopies !== 1 || boundary.outputBytes !== bytes.byteLength
        || boundary.jsonConversions !== 1 || !isExportWarnings(metadata.warnings)) {
        throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '写出结果的长度、警告或复制统计不匹配')
      }
      scope.postMessage({ context, type: 'success', bytes, warnings: metadata.warnings, metrics: {
        elapsedMs: performance.now() - started, boundary,
        transport: { inputTransfers: 0, inputBytes: 0, outputTransfers: 1,
          outputBytes: bytes.byteLength, structuredResultMessages: 0 },
      } }, [buffer])
    } finally { result.free() }
  } catch (error: unknown) {
    scope.postMessage({ context, type: 'error', error: failure(error) })
  } finally { active = undefined }
}

scope.onmessage = ({ data }) => {
  let context: WasmTaskControl | undefined
  try {
    if (!isWasmRecord(data) || data.type !== 'export') throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '写出请求必须使用协议 2')
    context = requireWasmTaskControl(data.context, 'export')
    if ((data.kind !== 'workbook' && data.kind !== 'document') || typeof data.data !== 'string'
      || typeof data.moduleUrl !== 'string' || !data.moduleUrl.trim()) {
      throw new WasmProtocolError('WASM_CONTROL_INVALID', '写出请求的模型或模块地址无效')
    }
    // JS UTF-16 length is a lower bound on UTF-8 size. Rust checks the exact
    // UTF-8 budget; no extra full-string TextEncoder allocation is needed.
    if (data.data.length > context.budget.maxInputBytes || data.data.length > context.budget.maxWorkingBytes) {
      throw new WasmProtocolError('WASM_MEMORY_LIMIT', 'Office 快照超过写出内存预算')
    }
    void write({ type: 'export', context, kind: data.kind, data: data.data, moduleUrl: data.moduleUrl })
  } catch (error) {
    const errorData = failure(error)
    scope.postMessage(context ? { type: 'error', context, error: errorData } : { type: 'protocol-error', error: errorData })
  }
}
