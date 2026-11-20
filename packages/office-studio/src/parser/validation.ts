import type { ParseEnvelope, RawParserError } from './contracts.ts'
import { ParserError } from './errors.ts'

export const record = (value: unknown): Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value)
    ? value as Record<string, unknown> : {}
export const array = (value: unknown): unknown[] => Array.isArray(value) ? value : []
export const text = (value: unknown, fallback = ''): string => typeof value === 'string' ? value : fallback
export const finite = (value: unknown, fallback = 0): number =>
  typeof value === 'number' && Number.isFinite(value) ? value : fallback
export const positive = (value: unknown, fallback: number): number => {
  const parsed = finite(value, fallback)
  return parsed > 0 ? parsed : fallback
}

export function validateEnvelope(value: unknown): ParseEnvelope {
  const raw = record(value)
  if (raw.schemaVersion !== 1) throw new ParserError('UNSUPPORTED_SCHEMA', '不支持此解析结果版本')
  const payload = record(raw.payload)
  if ((raw.model !== 'spreadsheet' && raw.model !== 'document') || payload.modelType !== raw.model ||
      !payload.data || typeof payload.data !== 'object' || Array.isArray(payload.data) ||
      typeof raw.fileName !== 'string' || typeof raw.format !== 'string' ||
      typeof raw.parserVersion !== 'string' || typeof raw.size !== 'number' || !Number.isFinite(raw.size) || raw.size < 0) {
    throw new ParserError('INVALID_PARSE_RESULT', '解析器返回的数据结构无效')
  }
  const diagnostics = record(raw.diagnostics)
  const errors: RawParserError[] = array(diagnostics.errors).map((source) => {
    const error = record(source)
    return { code: text(error.code, 'PARSER_FAILURE'), message: text(error.message, '解析错误'), details: text(error.details) }
  })
  return {
    schemaVersion: 1, parserVersion: raw.parserVersion, format: raw.format,
    fileName: raw.fileName, size: raw.size, model: raw.model,
    payload: { modelType: raw.model, data: record(payload.data) },
    diagnostics: {
      warnings: array(diagnostics.warnings).filter((item): item is string => typeof item === 'string'),
      errors, stats: record(diagnostics.stats),
      capabilities: array(diagnostics.capabilities).filter((item): item is string => typeof item === 'string'),
    },
  }
}
