import type { ParseEnvelope, ParserDiagnostic } from '../domain/model'

interface WasmParserModule {
  default?: (moduleOrPath?: unknown) => Promise<unknown>
  parse_office?: (bytes: Uint8Array, fileName: string) => string
}

export class ParserError extends Error {
  readonly code: string
  readonly diagnostics: ParserDiagnostic[]

  constructor(code: string, message: string, diagnostics: ParserDiagnostic[] = []) {
    super(message)
    this.name = 'ParserError'
    this.code = code
    this.diagnostics = diagnostics
  }
}

let parserModule: WasmParserModule | null = null
let loading: Promise<WasmParserModule> | null = null

async function loadParser(): Promise<WasmParserModule> {
  if (parserModule) return parserModule
  if (!loading) {
    const wasmEntry = '/wasm/office_parser.js'
    loading = import(/* @vite-ignore */ wasmEntry)
      .then(async (module) => {
        const parser = module as WasmParserModule
        if (parser.default) await parser.default()
        if (typeof parser.parse_office !== 'function') {
          throw new ParserError('WASM_API_MISSING', 'WASM 解析模块缺少 parse_office 接口')
        }
        parserModule = parser
        return parser
      })
      .catch((error: unknown) => {
        loading = null
        if (error instanceof ParserError) throw error
        throw new ParserError('WASM_LOAD_FAILED', 'WASM 解析模块加载失败，请先构建解析模块', [
          { code: 'WASM_LOAD_FAILED', severity: 'error', message: String(error) },
        ])
      })
  }
  return loading
}

export async function parseOfficeFile(file: File): Promise<ParseEnvelope> {
  const parser = await loadParser()
  const bytes = new Uint8Array(await file.arrayBuffer())
  try {
    const raw = parser.parse_office?.(bytes, file.name)
    if (!raw) throw new ParserError('EMPTY_PARSE_RESULT', '解析器没有返回结果')
    const result = JSON.parse(raw) as Partial<ParseEnvelope>
    if (typeof result.kind !== 'string' || result.payload === undefined) {
      throw new ParserError('INVALID_PARSE_RESULT', '解析器返回的数据结构无效')
    }
    return {
      schemaVersion: typeof result.schemaVersion === 'number' ? result.schemaVersion : 1,
      kind: result.kind,
      name: result.name ?? file.name,
      size: result.size ?? file.size,
      payload: result.payload,
      diagnostics: Array.isArray(result.diagnostics) ? result.diagnostics : [],
    }
  } catch (error: unknown) {
    if (error instanceof ParserError) throw error
    throw new ParserError('PARSE_FAILED', '文件解析失败', [{
      code: 'PARSE_FAILED',
      severity: 'error',
      message: error instanceof Error ? error.message : String(error),
    }])
  }
}

export function resetParser(): void {
  parserModule = null
  loading = null
}
