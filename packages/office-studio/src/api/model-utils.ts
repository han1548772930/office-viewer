import {
  createBlankDocument,
  createBlankSheet,
  createBlankWorkbook,
  createCell,
  defaultCellStyle,
  type CellScalar,
  type CellStyle,
  type DocumentModel,
  type SpreadsheetCell,
  type SpreadsheetSheet,
  type WorkbookModel,
} from '../domain/model'
import { cloneModel } from '../domain/serialization'

export type ApiCellData = {
  v?: CellScalar
  f?: string | null
  s?: Partial<CellStyle> & { font?: Partial<CellStyle['font']> }
  t?: string
}

export type ApiCellValue = CellScalar | ApiCellData

export function clone<T>(value: T): T {
  return cloneModel(value)
}

export function ensureWorkbookId(model: WorkbookModel): WorkbookModel {
  const next = clone(model)
  next.id ??= randomId('workbook')
  next.sheets = next.sheets.length ? next.sheets : [createBlankSheet()]
  next.sheets.forEach((sheet, index) => {
    sheet.id ||= randomId(`sheet-${index + 1}`)
    sheet.name ||= `Sheet${index + 1}`
    sheet.cells ||= {}
    sheet.mergedRanges ||= []
    sheet.columnWidths ||= {}
    sheet.rowHeights ||= {}
    sheet.frozenRows ||= 0
    sheet.frozenCols ||= 0
    sheet.visible ??= true
    sheet.showGridlines ??= true
  })
  next.activeSheetId = next.sheets.some((sheet) => sheet.id === next.activeSheetId) ? next.activeSheetId : next.sheets[0]!.id
  next.pivotTables ??= []
  return next
}

export function ensureDocumentId(model: DocumentModel): DocumentModel {
  const next = clone(model)
  next.id ??= randomId('document')
  return next
}

export function workbookFromInput(input: Partial<WorkbookModel> & Record<string, unknown> = {}): WorkbookModel {
  if (input.kind === 'workbook' && Array.isArray(input.sheets)) return ensureWorkbookId(input as WorkbookModel)
  const base = createBlankWorkbook()
  base.id = typeof input.id === 'string' ? input.id : randomId('workbook')
  base.name = typeof input.name === 'string' ? input.name : typeof input.title === 'string' ? input.title : base.name
  const sourceSheets = input.sheets
  if (Array.isArray(sourceSheets)) {
    base.sheets = sourceSheets.map((sheet) => normalizeSheet(sheet as Partial<SpreadsheetSheet> & Record<string, unknown>))
  } else if (sourceSheets && typeof sourceSheets === 'object') {
    const sheetMap = sourceSheets as Record<string, Record<string, unknown>>
    const order = Array.isArray(input.sheetOrder) ? input.sheetOrder.filter((id): id is string => typeof id === 'string') : Object.keys(sheetMap)
    base.sheets = order.map((id, index) => {
      const source = sheetMap[id] ?? {}
      const name = typeof source.name === 'string' ? source.name : `Sheet${index + 1}`
      return normalizeSheet({ ...source, id, name })
    })
  }
  if (!base.sheets.length) base.sheets = [createBlankSheet()]
  base.activeSheetId = typeof input.activeSheetId === 'string' ? input.activeSheetId : base.sheets[0]!.id
  return ensureWorkbookId(base)
}

export function documentFromInput(input: Partial<DocumentModel> & Record<string, unknown> = {}): DocumentModel {
  if (input.kind === 'document' && Array.isArray(input.blocks)) return ensureDocumentId(input as DocumentModel)
  const base = createBlankDocument()
  base.id = typeof input.id === 'string' ? input.id : randomId('document')
  base.name = typeof input.name === 'string' ? input.name : typeof input.title === 'string' ? input.title : base.name
  const body = input.body
  const dataStream = body && typeof body === 'object' && typeof (body as Record<string, unknown>).dataStream === 'string'
    ? String((body as Record<string, unknown>).dataStream)
    : typeof input.text === 'string' ? input.text : undefined
  if (dataStream != null) {
    const lines = dataStream.replace(/\r\n?/g, '\n').split('\n').filter((line, index, all) => index < all.length - 1 || line.length > 0)
    base.blocks = lines.length ? lines.map((line) => ({ ...base.blocks[0]!, id: randomId('paragraph'), runs: [{ ...base.blocks[0]!.runs[0]!, id: randomId('run'), text: line }] })) : base.blocks
  }
  return ensureDocumentId(base)
}

function normalizeSheet(input: Partial<SpreadsheetSheet> & Record<string, unknown>): SpreadsheetSheet {
  const sheet = createBlankSheet(typeof input.name === 'string' ? input.name : 'Sheet1')
  sheet.id = typeof input.id === 'string' ? input.id : randomId('sheet')
  sheet.name = typeof input.name === 'string' ? input.name : sheet.name
  sheet.rowCount = numberOrUndefined(input.rowCount)
  sheet.columnCount = numberOrUndefined(input.columnCount)
  sheet.visible = input.visible !== false
  sheet.showGridlines = input.showGridlines !== false
  sheet.frozenRows = numberOrZero(input.frozenRows ?? (input.freeze as Record<string, unknown> | undefined)?.rows)
  sheet.frozenCols = numberOrZero(input.frozenCols ?? (input.freeze as Record<string, unknown> | undefined)?.columns)
  if (input.cells && typeof input.cells === 'object') sheet.cells = clone(input.cells as Record<string, SpreadsheetCell>)
  const cellData = input.cellData
  if (cellData && typeof cellData === 'object') {
    for (const [rowKey, columns] of Object.entries(cellData as Record<string, unknown>)) {
      if (!columns || typeof columns !== 'object') continue
      for (const [columnKey, raw] of Object.entries(columns as Record<string, unknown>)) {
        const row = Number(rowKey), col = Number(columnKey)
        if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) continue
        sheet.cells[`${row}:${col}`] = cellFromApi(row, col, raw)
      }
    }
  }
  return sheet
}

export function cellFromApi(row: number, col: number, raw: unknown): SpreadsheetCell {
  const source = raw && typeof raw === 'object' ? raw as ApiCellData : undefined
  const scalar = source ? source.v ?? null : (raw as CellScalar)
  const cell = createCell(row, col, scalar === undefined ? null : scalar)
  if (source?.f) {
    cell.formula = source.f
    cell.value = source.v ?? null
    cell.display = source.v == null ? source.f : String(source.v)
  }
  if (source?.s) cell.style = mergeCellStyle(cell.style, source.s)
  return cell
}

export function mergeCellStyle(style: CellStyle, patch: Partial<CellStyle> & { font?: Partial<CellStyle['font']> }): CellStyle {
  return {
    ...style,
    ...patch,
    font: { ...style.font, ...(patch.font ?? {}) },
  }
}

export function applyCellValue(cell: SpreadsheetCell | undefined, row: number, col: number, value: ApiCellValue): SpreadsheetCell {
  const next = cell ? clone(cell) : createCell(row, col, null)
  const source = value && typeof value === 'object' && !Array.isArray(value) ? value as ApiCellData : undefined
  const scalar = source ? source.v ?? null : value as CellScalar
  next.formula = source ? source.f ?? null : typeof scalar === 'string' && scalar.startsWith('=') ? scalar : null
  next.value = next.formula ? (source?.v ?? null) : scalar
  next.display = next.formula ? (source?.v == null ? next.formula : String(source.v)) : scalar == null ? '' : String(scalar)
  if (source?.s) next.style = mergeCellStyle(next.style, source.s)
  return next
}

export function emptyCellStyle(): CellStyle {
  return defaultCellStyle()
}

function numberOrUndefined(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? Math.max(0, Math.floor(value)) : undefined
}

function numberOrZero(value: unknown): number {
  return numberOrUndefined(value) ?? 0
}

export function randomId(prefix: string): string {
  const uuid = globalThis.crypto?.randomUUID?.()
  return `${prefix}-${uuid ?? `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`}`
}

