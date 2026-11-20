import type { ImportedOfficeModel, ParseEnvelope } from './contracts.ts'
import { normalizeDocument } from './documentNormalizer.ts'
import { normalizeWorkbook } from './workbookNormalizer.ts'
import { ParserError } from './errors.ts'

export function normalizeOfficeEnvelope(envelope: ParseEnvelope): ImportedOfficeModel {
  if (envelope.diagnostics.errors.length) {
    const error = envelope.diagnostics.errors[0]!
    throw new ParserError(error.code, error.message)
  }
  const provenance = {
    format: envelope.format, fileName: envelope.fileName, size: envelope.size,
    parserVersion: envelope.parserVersion, stats: envelope.diagnostics.stats, capabilities: envelope.diagnostics.capabilities,
    diagnostics: envelope.diagnostics.warnings.map((message) => ({ code: 'PARSER_WARNING', severity: 'warning' as const, message })),
  }
  return envelope.model === 'spreadsheet'
    ? { ...provenance, kind: 'workbook', model: normalizeWorkbook(envelope) }
    : { ...provenance, kind: 'document', model: normalizeDocument(envelope) }
}
