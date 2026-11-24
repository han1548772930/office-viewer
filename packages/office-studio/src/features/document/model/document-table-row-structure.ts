import { type DocumentCell, type DocumentTable } from './document-model'
import { documentTableGrid, type DocumentTableGrid } from './document-table-grid'
import { rebindDocumentTableContinuations } from './document-table-continuations'

export interface DocumentTableRowCellPlacement { cell: DocumentCell; row: number; column: number; rowSpan: number; colSpan: number }

export function rebuildDocumentTableRows(table: DocumentTable, cells: DocumentTableRowCellPlacement[], rows = table.rows): void {
  table.rows = rows
  for (const row of rows) row.cells = []
  cells.sort((left, right) => left.row - right.row || left.column - right.column)
  for (const owner of cells) {
    owner.cell.columnIndex = owner.column
    owner.cell.rowSpan = owner.rowSpan
    owner.cell.colSpan = owner.colSpan
    rows[owner.row]!.cells.push(owner.cell)
    rebindDocumentTableContinuations(table, owner.cell, owner.row)
  }
}

/** Remove physical rows and shorten/re-anchor surviving vertical merges using the same algorithm as table commands. */
export function deleteDocumentTableRows(table: DocumentTable, grid: DocumentTableGrid, start: number, end: number): void {
  const amount = end - start + 1
  const cells = grid.anchors.flatMap(owner => {
    const overlap = Math.max(0, Math.min(owner.row + owner.rowSpan, end + 1) - Math.max(owner.row, start))
    if (overlap === owner.rowSpan) return []
    return [{ cell: owner.cell, row: owner.row < start ? owner.row : owner.row > end ? owner.row - amount : start,
      column: owner.column, rowSpan: owner.rowSpan - overlap, colSpan: owner.colSpan }]
  })
  const header = table.headerRowCount ?? 0
  table.headerRowCount = header - Math.max(0, Math.min(end + 1, header) - start)
  if (start === 0) table.titleRow = false
  if (table.rowStyles) table.rowStyles.splice(start, amount)
  const rows = [...table.rows]; rows.splice(start, amount)
  rebuildDocumentTableRows(table, cells, rows)
}

/** A copied row rectangle is independent; a merge crossing its boundary must be resolved before copying. */
export function insertCopiedDocumentTableRows(table: DocumentTable, grid: DocumentTableGrid, at: number,
  copied: DocumentTable, sourceStart: number): void {
  const amount = copied.rows.length, rows = [...table.rows]
  const copiedGrid = documentTableGrid(copied)
  const cells: DocumentTableRowCellPlacement[] = grid.anchors.map(owner => ({ cell: owner.cell,
    row: owner.row >= at ? owner.row + amount : owner.row, column: owner.column, rowSpan: owner.rowSpan, colSpan: owner.colSpan }))
  if (grid.anchors.some(owner => owner.row < at && owner.row + owner.rowSpan > at)) throw new RangeError('复制行不能切开跨项目纵向合并单元格。')
  rows.splice(at, 0, ...copied.rows)
  for (const owner of copiedGrid.anchors) cells.push({ cell: owner.cell, row: at + owner.row,
    column: owner.column, rowSpan: owner.rowSpan, colSpan: owner.colSpan })
  if (at < (table.headerRowCount ?? 0)) table.headerRowCount = (table.headerRowCount ?? 0) + amount
  if (table.rowStyles) table.rowStyles.splice(at, 0, ...copied.rows.map((_row, offset) => ({ ...(copied.rowStyles?.[offset] ?? table.rowStyles?.[sourceStart + offset] ?? {}) })))
  rebuildDocumentTableRows(table, cells, rows)
}

