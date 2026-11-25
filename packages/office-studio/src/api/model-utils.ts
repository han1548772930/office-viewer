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
import { inferSheetCellValueType, isForcedSheetCellValueType, normalizeSheetCellValueType, parseSheetCellText, type SheetCellTypeInput } from '../domain/sheet-cell-value.ts'
import { normalizeDocument } from '../features/document/model/document-model.ts'
import { setWorksheetProtectionRule } from '../features/sheet-protection/range-protection.ts'
import { formatSheetValue } from '../features/sheet-workbench/sheet-number-format.ts'
import { assertWorkbookConditionalFormatBudget } from '../features/conditional-formatting/conditional-format-workbook-budget.ts'
import { normalizeDocumentCustomXmlParts } from '../domain/document-custom-xml'
import { documentFromPublicInput } from './document-input'

export type ApiCellData = {
  v?: CellScalar
  f?: string | null
  s?: Partial<CellStyle> & { font?: Partial<CellStyle['font']> }
  t?: SheetCellTypeInput
}

export type ApiCellValue = CellScalar | ApiCellData

export function clone<T>(value: T): T {
  return cloneModel(value)
}

export function ensureWorkbookId(model: WorkbookModel): WorkbookModel {
  assertWorkbookConditionalFormatBudget(model.sheets)
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
    sheet.rangeProtectionRules ??= []
    if (sheet.worksheetProtection) setWorksheetProtectionRule(sheet, sheet.worksheetProtection)
  })
  next.activeSheetId = next.sheets.some((sheet) => sheet.id === next.activeSheetId) ? next.activeSheetId : next.sheets[0]!.id
  next.calculationMode ??= 'automatic'
  next.pivotTables ??= []
  return next
}

export function ensureDocumentId(model: DocumentModel): DocumentModel {
  const customXmlParts = model.customXmlParts === undefined ? undefined : normalizeDocumentCustomXmlParts(model.customXmlParts)
  const next = clone(model)
  if (customXmlParts !== undefined) next.customXmlParts = customXmlParts
  next.id ??= randomId('document')
  return next
}

export function workbookFromInput(input: Partial<WorkbookModel> & Record<string, unknown> = {}): WorkbookModel {
  if (input.kind === 'workbook' && Array.isArray(input.sheets)) return ensureWorkbookId(input as WorkbookModel)
  assertWorkbookConditionalFormatBudget(input.sheets)
  const base = createBlankWorkbook()
  base.id = typeof input.id === 'string' ? input.id : randomId('workbook')
  base.name = typeof input.name === 'string' ? input.name : typeof input.title === 'string' ? input.title : base.name
  base.date1904 = input.date1904 === true
  base.calculationMode = input.calculationMode === 'manual' ? 'manual' : 'automatic'
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
  if (input.kind === 'document' && Array.isArray(input.blocks)) return normalizeDocument(ensureDocumentId(input as DocumentModel))
  const base = createBlankDocument()
  base.id = typeof input.id === 'string' ? input.id : randomId('document')
  base.name = typeof input.name === 'string' ? input.name : typeof input.title === 'string' ? input.title : base.name
  return documentFromPublicInput(input, ensureDocumentId(base))
}

function normalizeSheet(input: Partial<SpreadsheetSheet> & Record<string, unknown>): SpreadsheetSheet {
  const sheet = createBlankSheet(typeof input.name === 'string' ? input.name : 'Sheet1')
  sheet.id = typeof input.id === 'string' ? input.id : randomId('sheet')
  sheet.name = typeof input.name === 'string' ? input.name : sheet.name
  if (Array.isArray(input.conditionalFormats)) sheet.conditionalFormats = clone({ conditionalFormats: input.conditionalFormats }).conditionalFormats
  if (Array.isArray(input.dataValidations)) sheet.dataValidations = clone(input.dataValidations)
  if (Array.isArray(input.rangeProtectionRules)) sheet.rangeProtectionRules = clone(input.rangeProtectionRules)
  if (input.worksheetProtection && typeof input.worksheetProtection === 'object') {
    setWorksheetProtectionRule(sheet, input.worksheetProtection)
  }
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
  const source = apiCellSource(raw)
  const scalar = apiCellScalar(source ? source.v : raw, source?.t)
  const cell = createCell(row, col, scalar)
  if (source?.f) {
    cell.formula = source.f
    cell.display = scalar === null ? source.f : String(scalar)
  }
  applyApiCellTypes(cell, source)
  if (source?.s) cell.style = mergeCellStyle(cell.style, source.s)
  if (!cell.formula || cell.value !== null) cell.display = formatSheetValue(cell.value, cell.style.numberFormat) ?? cell.display
  return cell
}

export function mergeCellStyle(style: CellStyle, patch: Partial<CellStyle> & { font?: Partial<CellStyle['font']> }): CellStyle {
  return {
    ...style,
    ...patch,
    font: { ...style.font, ...(patch.font ?? {}) },
  }
}

export function applyCellValue(cell: SpreadsheetCell | undefined, row: number, col: number, value: unknown): SpreadsheetCell {
  const next = cell ? clone(cell) : createCell(row, col, null)
  const source = apiCellSource(value)
  const scalar = apiCellScalar(source ? source.v : value, source?.t)
  next.formula = source ? source.f ?? null : typeof scalar === 'string' && scalar.startsWith('=') ? scalar : null
  next.value = next.formula && !source ? null : scalar
  next.display = next.formula && next.value === null ? next.formula : next.value === null ? '' : String(next.value)
  applyApiCellTypes(next, source)
  if (next.metadata && 'formulaError' in next.metadata) delete next.metadata.formulaError
  if (source?.s) next.style = mergeCellStyle(next.style, source.s)
  if (!next.formula || next.value !== null) next.display = formatSheetValue(next.value, next.style.numberFormat) ?? next.display
  return next
}

function apiCellSource(value: unknown): ApiCellData | undefined {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined
  const source = value as ApiCellData
  if (source.f !== undefined && source.f !== null && typeof source.f !== 'string') throw new TypeError('单元格公式必须是文本或空值。')
  if (source.t !== undefined && source.t !== null && typeof source.t !== 'string' && ![1, 2, 3, 4].includes(source.t)) {
    throw new TypeError('单元格数据类型无效。')
  }
  return source
}

function apiCellScalar(value: unknown, type?: SheetCellTypeInput): CellScalar {
  if (value === null || value === undefined) return null
  if (typeof value !== 'string' && typeof value !== 'boolean' && (typeof value !== 'number' || !Number.isFinite(value))) {
    throw new TypeError('单元格值必须是文本、有限数值、布尔值或空值。')
  }
  const kind = normalizeSheetCellValueType(type)
  if (kind === 'text' || kind === 'error') return String(value)
  return typeof value === 'string' && (kind === 'number' || kind === 'boolean') ? parseSheetCellText(value, kind) : value
}

function applyApiCellTypes(cell: SpreadsheetCell, source: ApiCellData | undefined): void {
  const declared = normalizeSheetCellValueType(source?.t)
  const type = declared === 'formula' ? undefined : declared ?? (typeof source?.t === 'string' ? source.t : undefined)
  cell.dataType = cell.formula ? 'formula' : isForcedSheetCellValueType(source?.t) ? 'force_string' : type ?? inferSheetCellValueType(cell.value)
  cell.cachedDataType = cell.formula ? type ?? (cell.value === null ? null : inferSheetCellValueType(cell.value)) : null
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

