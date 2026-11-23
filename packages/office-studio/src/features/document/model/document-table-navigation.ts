import type { DocumentModel } from '../../../domain/model'
import { findDocumentBlock, tableForBlock, type DocumentSelection } from './document-model'
import { documentTableGrid, expandDocumentTableRange, type DocumentTableCellRange } from './document-table-grid'
import { createDocumentTableViewSelection, resolveDocumentTableViewSelection } from './document-table-selection'

export type DocumentTableNavigationDirection = 'left' | 'right' | 'up' | 'down' | 'home' | 'end'
/** Transient keyboard endpoints retain Shift contraction without becoming document data. */
export interface DocumentTableKeyboardCursor {
  tableId: string
  anchor: { row: number; column: number }
  focus: { row: number; column: number }
  range: DocumentTableCellRange
}

const sameRange = (left: DocumentTableCellRange, right: DocumentTableCellRange) => left.startRow === right.startRow && left.endRow === right.endRow
  && left.startColumn === right.startColumn && left.endColumn === right.endColumn

/** Logical navigation skips covered slots and also works in object-only cells. */
export function navigateDocumentTableSelection(model: DocumentModel, selection: DocumentSelection, direction: DocumentTableNavigationDirection,
  extend = false, toEdge = false, previous?: DocumentTableKeyboardCursor): { selection: DocumentSelection; cursor: DocumentTableKeyboardCursor } | undefined {
  if (!selection.tableSelection) return undefined
  const view = resolveDocumentTableViewSelection(model, selection.tableSelection), table = tableForBlock(findDocumentBlock(model, view.tableId)!)!
  const grid = documentTableGrid(table), negative = direction === 'left' || direction === 'up' || direction === 'home'
  const continuing = previous?.tableId === view.tableId && sameRange(previous.range, view.range)
  const start = { row: view.range.startRow, column: view.range.startColumn }, end = { row: view.range.endRow, column: view.range.endColumn }
  const anchor = extend ? continuing ? { ...previous!.anchor } : negative ? end : start : continuing ? { ...previous!.focus } : negative ? start : end
  const focus = continuing ? { ...previous!.focus } : negative ? start : end
  const owner = grid.cellAt(focus.row, focus.column)!
  if (direction === 'left') focus.column = toEdge ? 0 : Math.max(0, owner.column - 1)
  if (direction === 'right') focus.column = toEdge ? grid.columnCount - 1 : Math.min(grid.columnCount - 1, owner.column + owner.colSpan)
  if (direction === 'up') focus.row = toEdge ? 0 : Math.max(0, owner.row - 1)
  if (direction === 'down') focus.row = toEdge ? grid.rowCount - 1 : Math.min(grid.rowCount - 1, owner.row + owner.rowSpan)
  if (direction === 'home' || direction === 'end') {
    focus.column = direction === 'home' ? 0 : grid.columnCount - 1
    if (toEdge) focus.row = direction === 'home' ? 0 : grid.rowCount - 1
  }
  const destination = grid.cellAt(focus.row, focus.column)!
  const origin = extend ? anchor : { row: destination.row, column: destination.column }
  const range = expandDocumentTableRange(grid, {
    startRow: Math.min(origin.row, destination.row), endRow: Math.max(origin.row, destination.row + destination.rowSpan - 1),
    startColumn: Math.min(origin.column, destination.column), endColumn: Math.max(origin.column, destination.column + destination.colSpan - 1),
  })
  const nextSelection = createDocumentTableViewSelection(model, view.tableId, range, extend ? 'range' : 'cell', undefined, selection)
  return { selection: nextSelection, cursor: { tableId: view.tableId, anchor: origin, focus, range: { ...range } } }
}
