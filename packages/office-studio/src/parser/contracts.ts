import type { DocumentModel, ParserDiagnostic, WorkbookModel } from '../domain/model.ts'

export interface RawParserError {
  code: string
  message: string
  details?: string | null
}

export interface DocumentOfficeXmlPartDto {
  partName: string
  xml: string
}

export type DocumentOfficeSourceResourceDto =
  | { kind: 'xml'; partName: string; xml: string }
  | { kind: 'binary'; streamName: string; offset: number; dataHex: string }

export interface DocumentOfficeSourceOriginDto {
  format: 'doc' | 'odt'
  synthesized: boolean
  resources: DocumentOfficeSourceResourceDto[]
  /** Historical diagnostics do not replace the current Word default-style chain. */
  resourceErrors?: string[]
}

export interface DocumentOfficeSourceDto {
  version: 1
  styles?: DocumentOfficeXmlPartDto | null
  theme?: DocumentOfficeXmlPartDto | null
  resourceErrors?: string[]
  /** DOCX omits origin; non-OOXML resources never become Word styles implicitly. */
  origin?: DocumentOfficeSourceOriginDto
  /** DOC/ODT resources retained across DOCX export and reopening. */
  legacyOrigins?: DocumentOfficeSourceOriginDto[]
}

export interface RawParserPayloadData extends Record<string, unknown> {
  officeSource?: DocumentOfficeSourceDto | null
}

export interface ParseEnvelope {
  schemaVersion: 1
  parserVersion: string
  format: string
  fileName: string
  size: number
  model: 'spreadsheet' | 'document'
  payload: { modelType: 'spreadsheet' | 'document'; data: RawParserPayloadData }
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
