import type { ParserDiagnostic } from '../domain/model.ts'
import type { RawParserError } from './contracts.ts'

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
  if (typeof error === 'string') {
    try {
      const value: unknown = JSON.parse(error)
      if (value && typeof value === 'object' && 'error' in value) {
        const nested: unknown = value.error
        if (nested && typeof nested === 'object' && 'code' in nested && 'message' in nested) {
          return {
            code: String(nested.code),
            message: String(nested.message),
            details: 'details' in nested ? String(nested.details ?? '') : undefined,
          }
        }
      }
    } catch { /* A non-JSON throw is retained as diagnostic text. */ }
  }
  return { code: 'PARSE_FAILED', message: error instanceof Error ? error.message : String(error) }
}
