export type OfficeKind = 'workbook' | 'document'

export type CellScalar = string | number | boolean | null

export interface FontStyle {
  family: string
  size: number
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  color: string
}

export interface CellStyle {
  font: FontStyle
  fill: string
  horizontal: 'left' | 'center' | 'right' | 'justify'
  vertical: 'top' | 'middle' | 'bottom'
  wrap: boolean
  numberFormat: string
}

export interface SpreadsheetCell {
  id: string
  row: number
  col: number
  value: CellScalar
  formula: string | null
  display: string
  style: CellStyle
}

export interface CellRange {
  startRow: number
  startCol: number
  endRow: number
  endCol: number
}

export interface SpreadsheetSheet {
  id: string
  name: string
  visible: boolean
  cells: Record<string, SpreadsheetCell>
  mergedRanges: CellRange[]
  columnWidths: Record<number, number>
  rowHeights: Record<number, number>
  frozenRows: number
  frozenCols: number
  showGridlines: boolean
}

export interface WorkbookModel {
  kind: 'workbook'
  name: string
  sheets: SpreadsheetSheet[]
  activeSheetId: string
  date1904: boolean
  modified: boolean
}

export type DocumentBlockType = 'paragraph' | 'heading' | 'bullet' | 'numbered' | 'quote' | 'divider' | 'table' | 'image' | 'pageBreak'

export interface InlineRun {
  id: string
  text: string
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  color: string
  fontFamily: string
  fontSize: number
}

export interface DocumentParagraphStyle {
  align: 'left' | 'center' | 'right' | 'justify'
  indent: number
  lineHeight: number
  spaceBefore: number
  spaceAfter: number
}

export interface DocumentBlock {
  id: string
  type: DocumentBlockType
  level: number
  runs: InlineRun[]
  style: DocumentParagraphStyle
  metadata: Record<string, unknown>
}

export interface DocumentPageSetup {
  width: number
  height: number
  marginTop: number
  marginRight: number
  marginBottom: number
  marginLeft: number
}

export interface DocumentModel {
  kind: 'document'
  name: string
  blocks: DocumentBlock[]
  page: DocumentPageSetup
  modified: boolean
  wordCount: number
}

export interface ParserDiagnostic {
  code: string
  severity: 'info' | 'warning' | 'error'
  message: string
  details?: string
}

export interface ParseEnvelope<TPayload = unknown> {
  schemaVersion: number
  kind: string
  name: string
  size: number
  payload: TPayload
  diagnostics: ParserDiagnostic[]
}

export interface EditorDocumentState {
  activeKind: OfficeKind
  workbook: WorkbookModel
  document: DocumentModel
  fileName: string
  parserStatus: 'idle' | 'loading' | 'ready' | 'error'
  parserMessage: string
  diagnostics: ParserDiagnostic[]
  selectedRange: CellRange
}

export const defaultCellStyle = (): CellStyle => ({
  font: {
    family: 'Aptos',
    size: 11,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color: '#1f2937',
  },
  fill: '#ffffff',
  horizontal: 'left',
  vertical: 'middle',
  wrap: false,
  numberFormat: 'General',
})

export const createCell = (row: number, col: number, value: CellScalar = ''): SpreadsheetCell => ({
  id: `${row}:${col}`,
  row,
  col,
  value,
  formula: null,
  display: value === null ? '' : String(value),
  style: defaultCellStyle(),
})

export const createBlankSheet = (name = 'Sheet1'): SpreadsheetSheet => ({
  id: `sheet-${crypto.randomUUID?.() ?? Date.now()}`,
  name,
  visible: true,
  cells: {},
  mergedRanges: [],
  columnWidths: {},
  rowHeights: {},
  frozenRows: 0,
  frozenCols: 0,
  showGridlines: true,
})

export const createBlankWorkbook = (): WorkbookModel => {
  const sheet = createBlankSheet()
  return {
    kind: 'workbook',
    name: '未命名工作簿',
    sheets: [sheet],
    activeSheetId: sheet.id,
    date1904: false,
    modified: false,
  }
}

export const createBlankDocument = (): DocumentModel => ({
  kind: 'document',
  name: '未命名文档',
  blocks: [{
    id: `block-${Date.now()}`,
    type: 'paragraph',
    level: 0,
    runs: [{
      id: `run-${Date.now()}`,
      text: '',
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      color: '#1f2937',
      fontFamily: 'Aptos',
      fontSize: 11,
    }],
    style: {
      align: 'left',
      indent: 0,
      lineHeight: 1.5,
      spaceBefore: 0,
      spaceAfter: 8,
    },
    metadata: {},
  }],
  page: {
    width: 794,
    height: 1123,
    marginTop: 72,
    marginRight: 72,
    marginBottom: 72,
    marginLeft: 72,
  },
  modified: false,
  wordCount: 0,
})
