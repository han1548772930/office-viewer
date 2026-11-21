import { randomId } from '../../api/model-utils.ts'
import type { CellRange, SheetTableModel, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'

export interface CreateSheetTableInput {
  sheetId: string
  range: CellRange
  name: string
  id?: string
}

export function createSheetTable(workbook: WorkbookModel, input: CreateSheetTableInput): WorkbookModel {
  const sheet = requireSheet(workbook, input.sheetId)
  validateRange(input.range)
  if (input.range.endRow <= input.range.startRow) throw new RangeError('A table needs a header row and at least one data row')
  if (workbook.tables?.some(table => table.name.toLocaleLowerCase() === input.name.trim().toLocaleLowerCase())) {
    throw new RangeError('Table names must be unique in a workbook')
  }
  const tableName = input.name.trim()
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(tableName) || tableName.toLocaleLowerCase() === 'r' || tableName.toLocaleLowerCase() === 'c') {
    throw new TypeError('Table name must start with a letter or underscore and contain only letters, numbers, underscores, or periods')
  }
  const table: SheetTableModel = {
    id: input.id ?? randomId('table'),
    name: tableName,
    sheetId: sheet.id,
    range: { ...input.range },
    columns: readUniqueHeaders(sheet, input.range),
    showHeaderRow: true,
    showTotalsRow: false,
    styleName: 'TableStyleMedium2',
  }
  ensureNoOverlap(workbook.tables ?? [], table)
  return { ...workbook, tables: [...(workbook.tables ?? []), table], modified: true }
}

export function resizeSheetTable(workbook: WorkbookModel, tableId: string, range: CellRange): WorkbookModel {
  validateRange(range)
  const table = requireTable(workbook, tableId)
  if (range.endRow <= range.startRow) throw new RangeError('A table needs a header row and at least one data row')
  if (range.endCol - range.startCol + 1 !== table.columns.length) {
    throw new RangeError('Resizing a table cannot change its column count')
  }
  const next = { ...table, range: { ...range } }
  ensureNoOverlap(workbook.tables ?? [], next, tableId)
  return replaceTable(workbook, next)
}

export function setSheetTableTotalsRow(workbook: WorkbookModel, tableId: string, visible: boolean): WorkbookModel {
  const table = requireTable(workbook, tableId)
  if (table.showTotalsRow === visible) return workbook
  return replaceTable(workbook, { ...table, showTotalsRow: visible })
}

export function deleteSheetTable(workbook: WorkbookModel, tableId: string): WorkbookModel {
  const tables = workbook.tables ?? []
  if (!tables.some(table => table.id === tableId)) return workbook
  return { ...workbook, tables: tables.filter(table => table.id !== tableId), modified: true }
}

export function validateSheetTableName(name: string): void {
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name.trim()) || ['r', 'c'].includes(name.trim().toLocaleLowerCase())) {
    throw new TypeError('Invalid table name')
  }
}

function requireSheet(workbook: WorkbookModel, sheetId: string): SpreadsheetSheet {
  const sheet = workbook.sheets.find(candidate => candidate.id === sheetId)
  if (!sheet) throw new RangeError(`Unknown worksheet: ${sheetId}`)
  return sheet
}

function requireTable(workbook: WorkbookModel, tableId: string): SheetTableModel {
  const table = workbook.tables?.find(candidate => candidate.id === tableId)
  if (!table) throw new RangeError(`Unknown table: ${tableId}`)
  return table
}

function validateRange(range: CellRange): void {
  if (![range.startRow, range.startCol, range.endRow, range.endCol].every(Number.isSafeInteger) ||
      range.startRow < 0 || range.startCol < 0 || range.endRow < range.startRow || range.endCol < range.startCol) {
    throw new RangeError('Invalid table range')
  }
}

function readUniqueHeaders(sheet: SpreadsheetSheet, range: CellRange): SheetTableModel['columns'] {
  const used = new Set<string>()
  return Array.from({ length: range.endCol - range.startCol + 1 }, (_, offset) => {
    const cell = Object.values(sheet.cells).find(candidate => candidate.row === range.startRow && candidate.col === range.startCol + offset)
    const rawName = typeof cell?.value === 'string' ? cell.value.trim() : ''
    const base = rawName || `Column${offset + 1}`
    let name = base
    for (let suffix = 2; used.has(name.toLocaleLowerCase()); suffix += 1) name = `${base}_${suffix}`
    used.add(name.toLocaleLowerCase())
    return { id: randomId('column'), name }
  })
}

function ensureNoOverlap(tables: readonly SheetTableModel[], candidate: SheetTableModel, excludingId?: string): void {
  const overlaps = tables.some(table => table.id !== excludingId && table.sheetId === candidate.sheetId &&
    candidate.range.startRow <= table.range.endRow && candidate.range.endRow >= table.range.startRow &&
    candidate.range.startCol <= table.range.endCol && candidate.range.endCol >= table.range.startCol)
  if (overlaps) throw new RangeError('Tables cannot overlap on a worksheet')
}

function replaceTable(workbook: WorkbookModel, next: SheetTableModel): WorkbookModel {
  return { ...workbook, tables: (workbook.tables ?? []).map(table => table.id === next.id ? next : table), modified: true }
}
