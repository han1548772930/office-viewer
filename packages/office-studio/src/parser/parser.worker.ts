import type { ParseRequest, ParseResponse } from './contracts'
import { serializeParserError } from './errors'
import { normalizeOfficeEnvelope } from './normalizer'
import { validateEnvelope } from './validation'

interface WasmModule { default: () => Promise<unknown>; parse_office: (bytes: Uint8Array, fileName: string) => string }
const scope = self as unknown as {
  postMessage: (message: ParseResponse) => void
  onmessage: ((event: MessageEvent<ParseRequest>) => void) | null
}
let loading: Promise<WasmModule> | undefined
scope.onmessage = async ({ data }) => {
  const { requestId } = data
  const progress = (stage: 'loading' | 'parsing' | 'normalizing', message: string) =>
    scope.postMessage({ requestId, type: 'progress', progress: { stage, message } })
  try {
    progress('loading', '正在加载文档解析器')
    loading ??= import(/* @vite-ignore */ data.baseUrl+'wasm/office_parser.js').then(async (module: WasmModule) => {
      if (typeof module.default !== 'function' || typeof module.parse_office !== 'function') throw new Error('WASM_API_MISSING')
      await module.default()
      return module
    }).catch((error: unknown) => { loading = undefined; throw error })
    const parser = await loading
    progress('parsing', '正在解析文件内容')
    const raw: unknown = JSON.parse(parser.parse_office(new Uint8Array(data.bytes), data.fileName))
    progress('normalizing', '正在生成编辑模型')
    const result = normalizeOfficeEnvelope(validateEnvelope(raw))
    scope.postMessage({ requestId, type: 'success', result })
  } catch (error: unknown) {
    scope.postMessage({ requestId, type: 'error', error: serializeParserError(error) })
  }
}
