import type { ParseRequest, ParseResponse } from './contracts.ts'
import { serializeParserError } from './errors.ts'
import { normalizeOfficeEnvelope } from './normalizer.ts'
import { validateEnvelope } from './validation.ts'
import { loadParserWasm } from './wasm-modules.ts'
import {
  isWasmRecord, readBoundaryCopyStats, requireMatchingWasmResult, requireWasmTaskControl,
  WasmProtocolError,
} from '../platform/wasm/task-protocol.ts'
import type { WasmTaskControl } from '../platform/wasm/task-protocol.ts'

const scope = self as unknown as {
  postMessage(message: ParseResponse): void
  onmessage: ((event: MessageEvent<unknown>) => void) | null
}
let loading: ReturnType<typeof loadParserWasm> | undefined
let moduleUrl: string | undefined
let active: WasmTaskControl | undefined

async function parse(request: ParseRequest): Promise<void> {
  const context = request.context
  if (active) {
    scope.postMessage({ context, type: 'error', error: { code: 'PARSE_BUSY', message: '解析器正在执行其他任务' } })
    return
  }
  active = context
  const started = performance.now()
  const inputBytes = request.bytes.byteLength
  const progress = (stage: 'loading' | 'parsing' | 'normalizing', message: string) =>
    scope.postMessage({ context, type: 'progress', progress: { stage, message } })
  try {
    progress('loading', '正在加载文档解析器')
    if (moduleUrl !== undefined && moduleUrl !== request.moduleUrl) {
      throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '同一解析 Worker 不能切换 WASM 模块')
    }
    moduleUrl = request.moduleUrl
    loading ??= loadParserWasm(request.moduleUrl).catch((error: unknown) => { loading = undefined; throw error })
    const parser = await loading
    progress('parsing', '正在解析文件内容')
    const raw = requireMatchingWasmResult(await parser.module.parse_office(new Uint8Array(request.bytes), request.fileName, context), context)
    const boundary = readBoundaryCopyStats(raw.transfer)
    const envelope = validateEnvelope(raw.envelope)
    if (envelope.fileName !== request.fileName || envelope.size !== inputBytes
      || boundary.inputBytes !== inputBytes || boundary.inputCopies !== (inputBytes > 0 ? 1 : 0)
      || boundary.outputCopies !== 0 || boundary.outputBytes !== 0 || boundary.jsonConversions !== 0) {
      throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '解析结果的来源或边界统计不匹配')
    }
    progress('normalizing', '正在生成编辑模型')
    const result = normalizeOfficeEnvelope(envelope)
    // The imported unit gets the task identity chosen by its owner. Source parts
    // remain in their format provenance; they do not define a second unit ID.
    result.model.id = context.unitId
    scope.postMessage({ context, type: 'success', result, metrics: {
      elapsedMs: performance.now() - started, boundary,
      transport: { inputTransfers: 1, inputBytes, outputTransfers: 0, outputBytes: 0, structuredResultMessages: 1 },
    } })
  } catch (error: unknown) {
    scope.postMessage({ context, type: 'error', error: serializeParserError(error) })
  } finally { active = undefined }
}

scope.onmessage = ({ data }) => {
  let context: WasmTaskControl | undefined
  try {
    if (!isWasmRecord(data) || data.type !== 'parse') throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '解析请求必须使用协议 2')
    context = requireWasmTaskControl(data.context, 'parse')
    if (!(data.bytes instanceof ArrayBuffer) || typeof data.fileName !== 'string' || !data.fileName.trim()
      || typeof data.moduleUrl !== 'string' || !data.moduleUrl.trim()) {
      throw new WasmProtocolError('WASM_CONTROL_INVALID', '解析请求缺少文件缓冲区、名称或模块地址')
    }
    // Constructing a view rejects a detached buffer, including detached empties.
    const view = new Uint8Array(data.bytes)
    if (!view.byteLength) throw new WasmProtocolError('EMPTY_FILE', '文件为空')
    if (view.byteLength > context.budget.maxInputBytes || view.byteLength > context.budget.maxWorkingBytes) {
      throw new WasmProtocolError('WASM_MEMORY_LIMIT', '文件超过解析内存预算')
    }
    void parse({ type: 'parse', context, bytes: data.bytes, fileName: data.fileName, moduleUrl: data.moduleUrl })
  } catch (error) {
    const failure = serializeParserError(error)
    scope.postMessage(context ? { type: 'error', context, error: failure } : { type: 'protocol-error', error: failure })
  }
}
