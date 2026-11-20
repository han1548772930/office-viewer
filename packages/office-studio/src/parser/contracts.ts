import type { DocumentModel, ParserDiagnostic, WorkbookModel } from '../domain/model.ts'

export interface RawParserError {
  code: string
  message: string
  details?: string | null
}

export interface ParseEnvelope {
  schemaVersion: 1
  parserVersion: string
  format: string
  fileName: string
  size: number
  model: 'spreadsheet' | 'document'
  payload: { modelType: 'spreadsheet' | 'document'; data: Record<string, unknown> }
  diagnostics: {
    warnings: string[]
    errors: RawParserError[]
    stats: Record<string, unknown>
    capabilities: string[]
  }
}

interface ImportProvenance {
  format: string
  parserVersion: string
  fileName: string
  size: number
  diagnostics: ParserDiagnostic[]
  stats: Record<string, unknown>
  capabilities: string[]
}

export type ImportedOfficeModel = ImportProvenance & (
  { kind: 'workbook'; model: WorkbookModel } | { kind: 'document'; model: DocumentModel }
)

export type ParseStage = 'loading' | 'reading' | 'parsing' | 'normalizing' | 'ready'
export interface ParseProgress { stage: ParseStage; message: string }
export interface ParseRequest { requestId: string; bytes: ArrayBuffer; fileName: string; baseUrl: string }
export type ParseResponse =
  | { requestId: string; type: 'progress'; progress: ParseProgress }
  | { requestId: string; type: 'success'; result: ImportedOfficeModel }
  | { requestId: string; type: 'error'; error: RawParserError }
