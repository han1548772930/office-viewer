import type { DocumentModel } from '../../../domain/model'
import { findDocumentBlock, tableForBlock, type DocumentSelection, type DocumentTableViewSelection } from '../model/document-model'
import { documentTableGrid, documentTableWholeRange, type DocumentTableCellRange } from '../model/document-table-grid'
import { createDocumentTableViewSelection } from '../model/document-table-selection'
import type { DocumentLayout, LayoutCell, LayoutRect, LayoutTableFragment, LayoutTableRowMetric } from './document-layout-types'
import { intersectTableRect } from './document-table-measurement'

export interface DocumentTableHit { cell: LayoutCell; fragment: LayoutTableFragment }
export interface DocumentTableDragPointer { clientX: number; clientY: number; pointerId: number; ctrlKey: boolean; metaKey: boolean }
export interface DocumentTableGesturePreview {
  rectangle?: LayoutRect
  guide?: { x1: number; y1: number; x2: number; y2: number }
  available: boolean
}

export const documentTableFragmentKey = (fragment: LayoutTableFragment) => `${fragment.tableId}:${fragment.page}:${fragment.column ?? 0}:${fragment.sourceStart}:${fragment.repeatedHeader ? 1 : 0}:${fragment.y}`
export function documentTableVisibleRect(rectangle: LayoutRect & { clip?: LayoutRect }): LayoutRect {
  return rectangle.clip ? intersectTableRect(rectangle, rectangle.clip) : rectangle
}
const contains = (rectangle: LayoutRect, x: number, y: number) => rectangle.width > 0 && rectangle.height > 0
  && x >= rectangle.x && x < rectangle.x + rectangle.width && y >= rectangle.y && y < rectangle.y + rectangle.height

/** Choose the deepest visible owner, retaining logical coordinates and physical-page affinity. */
export function hitTestDocumentTable(layout: DocumentLayout, x: number, y: number, tableId?: string): DocumentTableHit | undefined {
  const cells = layout.cells.filter(cell => cell.tableId && cell.cellId && (!tableId || cell.tableId === tableId) && contains(documentTableVisibleRect(cell), x, y))
    .sort((left, right) => (right.depth ?? 0) - (left.depth ?? 0) || left.width * left.height - right.width * right.height)
  for (const cell of cells) {
    const fragment = layout.tableFragments?.find(candidate => candidate.tableId === cell.tableId && candidate.page === cell.page
      && candidate.depth === (cell.depth ?? 0) && candidate.repeatedHeader === !!cell.repeatedHeader && contains(documentTableVisibleRect(candidate), x, y))
    if (fragment) return { cell, fragment }
  }
}

export function documentTableFragmentColumns(fragment: LayoutTableFragment): Array<{ column: number; start: number; end: number; width: number }> {
  let x = fragment.x
  return fragment.columnWidths.map((width, column) => { const start = x; x += width; return { column, start, end: x, width } })
}

export function documentTableFragmentRows(fragment: LayoutTableFragment): Array<LayoutTableRowMetric & { visibleTop: number; visibleBottom: number; ownsBottom: boolean }> {
  const clip = documentTableVisibleRect(fragment)
  return fragment.rows.filter(row => row.y < clip.y + clip.height && row.y + row.height > clip.y).map(row => ({ ...row,
    visibleTop: Math.max(row.y, clip.y), visibleBottom: Math.min(row.y + row.height, clip.y + clip.height),
    ownsBottom: row.y + row.height <= clip.y + clip.height + 1e-6 }))
}

export function documentTableLogicalHit(fragment: LayoutTableFragment, x: number, y: number): { row: number; column: number } | undefined {
  const clip = documentTableVisibleRect(fragment)
  if (!contains(clip, x, y)) return undefined
  const row = documentTableFragmentRows(fragment).find(row => y >= row.visibleTop && y < row.visibleBottom)
  const column = documentTableFragmentColumns(fragment).find(column => x >= column.start && x < column.end)
  return row && column ? { row: row.row, column: column.column } : undefined
}

export function documentTableSelectionAtHit(model: DocumentModel, hit: DocumentTableHit, fallback: DocumentSelection): DocumentSelection {
  const row = hit.cell.row!, column = hit.cell.columnIndex!
  return createDocumentTableViewSelection(model, hit.fragment.tableId, { startRow: row, endRow: row + (hit.cell.rowSpan ?? 1) - 1,
    startColumn: column, endColumn: column + (hit.cell.colSpan ?? 1) - 1 }, 'cell', hit.fragment.page, fallback)
}

/** Prefer the actual logical focus cell, then the selected cells on its derived page. */
export function documentTableSelectionFocusCell(layout: DocumentLayout, selection: DocumentSelection,
  focus?: { row: number; column: number }): LayoutCell | undefined {
  const view = selection.tableSelection
  if (!view) return undefined
  const row = focus?.row ?? view.range.endRow, column = focus?.column ?? view.range.endColumn
  const cells = layout.cells.filter(cell => cell.tableId === view.tableId && cell.row !== undefined && cell.columnIndex !== undefined
    && cell.row <= view.range.endRow && cell.row + (cell.rowSpan ?? 1) > view.range.startRow
    && cell.columnIndex <= view.range.endColumn && cell.columnIndex + (cell.colSpan ?? 1) > view.range.startColumn
    && documentTableVisibleRect(cell).height > 0 && documentTableVisibleRect(cell).width > 0)
  const focused = cells.filter(cell => cell.row! <= row && cell.row! + (cell.rowSpan ?? 1) > row
    && cell.columnIndex! <= column && cell.columnIndex! + (cell.colSpan ?? 1) > column)
  const candidates = focused.length ? focused : cells
  return candidates.find(cell => cell.page === view.pageAffinity) ?? candidates[0]
}

export function documentTableAxisSelection(model: DocumentModel, tableId: string, kind: DocumentTableViewSelection['kind'],
  index: number | undefined, page: number, fallback: DocumentSelection, extend = false): DocumentSelection {
  const block = findDocumentBlock(model, tableId), table = block && tableForBlock(block)
  if (!table) throw new RangeError('表格已不存在')
  const grid = documentTableGrid(table), range = documentTableWholeRange(grid)
  if (kind === 'row' && index !== undefined) range.startRow = range.endRow = index
  if (kind === 'column' && index !== undefined) range.startColumn = range.endColumn = index
  const previous = fallback.tableSelection
  if (extend && previous?.tableId === tableId && previous.kind === kind) {
    range.startRow = Math.min(range.startRow, previous.range.startRow); range.endRow = Math.max(range.endRow, previous.range.endRow)
    range.startColumn = Math.min(range.startColumn, previous.range.startColumn); range.endColumn = Math.max(range.endColumn, previous.range.endColumn)
  }
  return createDocumentTableViewSelection(model, tableId, range, kind, page, fallback)
}

export function documentTableRangeBetweenHits(model: DocumentModel, anchor: DocumentTableHit, focus: DocumentTableHit, fallback: DocumentSelection): DocumentSelection | undefined {
  if (anchor.fragment.tableId !== focus.fragment.tableId) return undefined
  const range: DocumentTableCellRange = { startRow: Math.min(anchor.cell.row!, focus.cell.row!),
    endRow: Math.max(anchor.cell.row! + (anchor.cell.rowSpan ?? 1) - 1, focus.cell.row! + (focus.cell.rowSpan ?? 1) - 1),
    startColumn: Math.min(anchor.cell.columnIndex!, focus.cell.columnIndex!),
    endColumn: Math.max(anchor.cell.columnIndex! + (anchor.cell.colSpan ?? 1) - 1, focus.cell.columnIndex! + (focus.cell.colSpan ?? 1) - 1) }
  return createDocumentTableViewSelection(model, anchor.fragment.tableId, range, 'range', focus.fragment.page, fallback)
}
