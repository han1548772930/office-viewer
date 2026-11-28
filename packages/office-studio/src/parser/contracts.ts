import type { DocumentCustomXmlPart, DocumentModel, ParserDiagnostic, WorkbookModel } from '../domain/model.ts'
import type { WasmTaskControl, WasmTaskMetrics } from '../platform/wasm/task-protocol.ts'

export interface RawParserError {
  code: string
  message: string
  details?: string | null
}

export interface DocumentOfficeXmlPartDto {
  partName: string
  xml: string
}

export interface DocumentNumberingMediaDto {
  partName: string
  contentType: string
  dataBase64: string
}

export interface DocumentNumberingSourceDto extends DocumentOfficeXmlPartDto {
  relationshipsXml?: string | null
  resources?: DocumentNumberingMediaDto[]
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
  numbering?: DocumentNumberingSourceDto | null
  resourceErrors?: string[]
  /** DOCX omits origin; non-OOXML resources never become Word styles implicitly. */
  origin?: DocumentOfficeSourceOriginDto
  /** DOC/ODT resources retained across DOCX export and reopening. */
  legacyOrigins?: DocumentOfficeSourceOriginDto[]
}

export interface RawParserPayloadData extends Record<string, unknown> {
  officeSource?: DocumentOfficeSourceDto | null
  customXmlParts?: DocumentCustomXmlPart[]
  nativeContentControls?: NativeDocumentContentControlDto[]
  formulaCalendarSource?: FormulaCalendarSourceDto | null
}

/** This DTO contains only sources admitted against the native XLSX package. */
export interface FormulaCalendarSourceDto {
  version: 1
  date1904: boolean
  sheets: FormulaCalendarSheetSourceDto[]
  helpers: FormulaCalendarHelperDto[]
  tables: FormulaCalendarTableDto[]
}
export interface FormulaCalendarSheetSourceDto {
  name: string
  sheetId: string
  officeSheetId: number
  import?: Record<string, unknown>
  cells: FormulaCalendarCellSourceDto[]
}
export interface FormulaCalendarCellSourceDto {
  row: number
  col: number
  formula?: string
  projectedFormula?: string
  formulaSource?: { version: 1; formula: string; date1904: boolean; dateResult: boolean; aliasSheetId?: string } & Record<string, unknown>
  dateSource?: { value: number | string; date1904: boolean; targetValue?: number | string; targetDate1904?: boolean; targetFormula?: string | null } & Record<string, unknown>
  calculationSheet?: string
  calculationFormula?: string
}
export interface FormulaCalendarHelperDto {
  name: string
  sheetId: string
  officeSheetId: number
  date1904: boolean
  sourceSheetId?: string
  fingerprint: string
  reclaim?: boolean
}
export interface FormulaCalendarTableDto {
  name: string
  tableId: string
  officeTableId: number
  sheetId: string
  /** Older provenance parts omit columns; native IDs never depend on position. */
  columns?: FormulaCalendarTableColumnDto[]
}
export interface FormulaCalendarTableColumnDto {
  columnId: string
  officeColumnId: number
}

/** Native parser addresses precede stable domain table/row ID normalization. */
export interface NativeDocumentContentControlDto {
  id: string
  parentId?: string
  properties: Record<string, unknown>
  textRanges: Array<{ blockId: string; startOffset: number; endOffset: number; startNoteReferenceCount?: number; endNoteReferenceCount?: number }>
  hasObjectContent?: boolean
  objectContentCount?: number
  blockIds?: string[]
  tableRowRange?: { tableBlockId: string; rowIndexes: number[] }
  partName: string
  story: string
  noteId?: number
  sourceXml: string
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
  /** Task provenance stays outside the persistent Office model. */
  task?: { context: WasmTaskControl; metrics: WasmTaskMetrics }
}

export type ImportedOfficeModel = ImportProvenance & (
  { kind: 'workbook'; model: WorkbookModel } | { kind: 'document'; model: DocumentModel }
)

export type ParseStage = 'loading' | 'reading' | 'parsing' | 'normalizing' | 'ready'
export interface ParseProgress { stage: ParseStage; message: string }
export interface ParseRequest { type: 'parse'; context: WasmTaskControl; bytes: ArrayBuffer; fileName: string; moduleUrl: string }
export type ParseResponse =
  | { context: WasmTaskControl; type: 'progress'; progress: ParseProgress }
  | { context: WasmTaskControl; type: 'success'; result: ImportedOfficeModel; metrics: WasmTaskMetrics }
  | { context: WasmTaskControl; type: 'error'; error: RawParserError }
  | { type: 'protocol-error'; error: RawParserError }
