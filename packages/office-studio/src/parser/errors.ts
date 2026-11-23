import type { ParserDiagnostic } from '../domain/model.ts'
import type { RawParserError } from './contracts.ts'
import { serializeWasmError } from '../platform/wasm/task-protocol.ts'

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

export function serializeParserError(error: unknown): RawParserError {
  if (error instanceof ParserError) return { code: error.code, message: error.message }
  return serializeWasmError(error, { code: 'PARSE_FAILED', message: '文件解析失败' })
}
