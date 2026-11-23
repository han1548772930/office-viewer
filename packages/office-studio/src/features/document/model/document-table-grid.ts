import type { DocumentCell, DocumentTable, TextTarget } from './document-model'
import { assertDocumentTableContinuationRows } from './document-table-continuations'

/** Inclusive logical row/column range, independent of compact storage indices. */
export interface DocumentTableCellRange {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

export interface DocumentTableGridCell {
  cell: DocumentCell
  row: number
  column: number
  cellIndex: number
  rowSpan: number
  colSpan: number
}

export interface DocumentTableGrid {
  table: DocumentTable
  rowCount: number
  columnCount: number
  anchors: DocumentTableGridCell[]
  /** Each row contains nonoverlapping column intervals, including rowSpan owners. */
  rows: DocumentTableGridCell[][]
  cellAt(row: number, column: number): DocumentTableGridCell | undefined
  storedCell(row: number, cellIndex: number): DocumentTableGridCell | undefined
}

function positiveInteger(value: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(`表格${label}必须为正整数`)
  return value
}

function nextFreeColumn(intervals: readonly DocumentTableGridCell[], start: number, span: number): number {
  let column = start
  for (const owner of intervals) {
    if (owner.column + owner.colSpan <= column) continue
    if (column + span <= owner.column) break
    column = owner.column + owner.colSpan
  }
  return column
}

/**
 * Compact anchors define one authoritative grid. Missing legacy columnIndex
 * values skip earlier rowSpan owners; explicit overlaps or holes are rejected.
 * The reader never changes the imported table or fabricates missing cells.
 */
export function documentTableGrid(table: DocumentTable): DocumentTableGrid {
  if (!Array.isArray(table.rows) || !table.rows.length) throw new RangeError('表格至少需要一行')
  if (table.rows.length > 10_000) throw new RangeError('表格超过 10000 行的解析与写出预算，原表格已保留')
  const rows: DocumentTableGridCell[][] = Array.from({ length: table.rows.length }, () => [])
  const anchors: DocumentTableGridCell[] = []
  const stored = new Map<string, DocumentTableGridCell>()
  const rowIds = new Set<string>(), cellIds = new Set<string>()
  let columnCount = 0
  table.rows.forEach((sourceRow, row) => {
    if (typeof sourceRow.id !== 'string' || !sourceRow.id || rowIds.has(sourceRow.id) || !Array.isArray(sourceRow.cells)) {
      throw new TypeError('表格行标识或单元格集合无效')
    }
    rowIds.add(sourceRow.id)
    let cursor = 0
    sourceRow.cells.forEach((cell, cellIndex) => {
      if (typeof cell.id !== 'string' || !cell.id || cellIds.has(cell.id) || !Array.isArray(cell.blocks)) throw new TypeError('表格单元格标识或内容集合无效')
      cellIds.add(cell.id)
      const rowSpan = positiveInteger(cell.rowSpan, '跨行数'), colSpan = positiveInteger(cell.colSpan, '跨列数')
      if (row + rowSpan > table.rows.length) throw new RangeError('表格跨行合并越过实际行数，原表格已保留')
      assertDocumentTableContinuationRows(table, cell, row)
      rows[row]!.sort((left, right) => left.column - right.column)
      const column = cell.columnIndex === undefined ? nextFreeColumn(rows[row]!, cursor, colSpan) : cell.columnIndex
      if (!Number.isSafeInteger(column) || column < 0 || !Number.isSafeInteger(column + colSpan)) throw new RangeError('表格逻辑列无效')
      if (column + colSpan > 512 || table.rows.length * (column + colSpan) > 1_000_000) throw new RangeError('表格超过 512 列或 1000000 逻辑单元格的解析与写出预算，原表格已保留')
      const owner = { cell, row, column, cellIndex, rowSpan, colSpan }
      for (let coveredRow = row; coveredRow < row + rowSpan; coveredRow++) {
        if (rows[coveredRow]!.some(previous => previous.column < column + colSpan && column < previous.column + previous.colSpan)) {
          throw new RangeError('表格单元格跨度重叠，原表格已保留')
        }
        rows[coveredRow]!.push(owner)
      }
      cursor = column + colSpan
      columnCount = Math.max(columnCount, cursor)
      anchors.push(owner)
      stored.set(`${row}:${cellIndex}`, owner)
    })
  })
  if (!columnCount) throw new RangeError('表格至少需要一列')
  for (const intervals of rows) {
    intervals.sort((left, right) => left.column - right.column)
    let column = 0
    for (const owner of intervals) {
      if (owner.column !== column) throw new RangeError('表格逻辑网格包含空洞，原表格已保留')
      column += owner.colSpan
    }
    if (column !== columnCount) throw new RangeError('表格各行未覆盖相同逻辑列，原表格已保留')
  }
  if (table.columnWidths?.length && (table.columnWidths.length !== columnCount
    || table.columnWidths.some(width => !Number.isFinite(width) || width <= 0))) {
    throw new RangeError('表格列宽与逻辑网格不一致或数值无效，原表格已保留')
  }
  if (table.headerRowCount !== undefined && (!Number.isSafeInteger(table.headerRowCount)
    || table.headerRowCount < 0 || table.headerRowCount > table.rows.length)) throw new RangeError('表格表头行数无效')
  anchors.sort((left, right) => left.row - right.row || left.column - right.column)
  return { table, rowCount: table.rows.length, columnCount, anchors, rows,
    cellAt: (row, column) => Number.isSafeInteger(row) && Number.isSafeInteger(column) && row >= 0 && column >= 0
      ? rows[row]?.find(owner => column >= owner.column && column < owner.column + owner.colSpan) : undefined,
    storedCell: (row, cellIndex) => stored.get(`${row}:${cellIndex}`),
  }
}

export function documentTableWholeRange(grid: DocumentTableGrid): DocumentTableCellRange {
  return { startRow: 0, endRow: grid.rowCount - 1, startColumn: 0, endColumn: grid.columnCount - 1 }
}

export function assertDocumentTableRange(grid: DocumentTableGrid, range: DocumentTableCellRange): void {
  if (!range || ![range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isSafeInteger)
    || range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn
    || range.endRow >= grid.rowCount || range.endColumn >= grid.columnCount) throw new RangeError('表格选区超出逻辑网格')
}

export function documentTableCellsInRange(grid: DocumentTableGrid, range: DocumentTableCellRange): DocumentTableGridCell[] {
  assertDocumentTableRange(grid, range)
  return grid.anchors.filter(owner => owner.row <= range.endRow && owner.row + owner.rowSpan > range.startRow
    && owner.column <= range.endColumn && owner.column + owner.colSpan > range.startColumn)
}

/** A rectangular selection includes the complete extent of every merged cell it touches. */
export function expandDocumentTableRange(grid: DocumentTableGrid, source: DocumentTableCellRange): DocumentTableCellRange {
  assertDocumentTableRange(grid, source)
  const range = { ...source }
  while (true) {
    const previous = { ...range }
    for (const owner of documentTableCellsInRange(grid, range)) {
      range.startRow = Math.min(range.startRow, owner.row)
      range.endRow = Math.max(range.endRow, owner.row + owner.rowSpan - 1)
      range.startColumn = Math.min(range.startColumn, owner.column)
      range.endColumn = Math.max(range.endColumn, owner.column + owner.colSpan - 1)
    }
    if (range.startRow === previous.startRow && range.endRow === previous.endRow
      && range.startColumn === previous.startColumn && range.endColumn === previous.endColumn) return range
  }
}

export function documentTableCellTarget(blockId: string, owner: DocumentTableGridCell, cellBlockId?: string): TextTarget {
  return { blockId, row: owner.row, col: owner.cellIndex, ...(cellBlockId ? { cellBlockId } : {}) }
}
