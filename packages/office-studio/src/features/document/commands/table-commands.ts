import type { DocumentModel } from '../../../domain/model'
import { findDocumentBlock, tableForBlock, type DocumentSelection } from '../model/document-model'
import {
  assertDocumentTableRange, documentTableCellsInRange, documentTableGrid, documentTableWholeRange, expandDocumentTableRange,
  type DocumentTableCellRange,
} from '../model/document-table-grid'
import type { DocumentPermissionContext } from '../permissions/document-permissions'
import { assertDocumentTableMutation, prepareDocumentTableMutation, type DocumentTableMutationTarget } from './table-mutations'
import type { DocumentTableCommand } from './table-command-types'

export type { DocumentTableCommand } from './table-command-types'

export interface DocumentTableSelection extends DocumentTableCellRange {
  blockId: string
  rowCount: number
  columnCount: number
  hasMergedCells: boolean
  canMergeCells: boolean
  canUnmergeCells: boolean
}

/** TextTarget.col addresses storage; this context always reports logical columns. */
export function documentTableSelection(model: DocumentModel, selection: DocumentSelection): DocumentTableSelection | undefined {
  const { anchor, focus } = selection
  if (!selection.tableSelection && (anchor.blockId !== focus.blockId || anchor.row == null || focus.row == null || anchor.col == null || focus.col == null)) return undefined
  const block = findDocumentBlock(model, selection.tableSelection?.tableId ?? focus.blockId), table = block && tableForBlock(block)
  if (!table?.rows.length) return undefined
  try {
    const grid = documentTableGrid(table), first = anchor.row === undefined || anchor.col === undefined ? undefined : grid.storedCell(anchor.row, anchor.col)
    const last = focus.row === undefined || focus.col === undefined ? undefined : grid.storedCell(focus.row, focus.col)
    if (!selection.tableSelection && (!first || !last)) return undefined
    const range = expandDocumentTableRange(grid, selection.tableSelection?.range ?? {
      startRow: Math.min(first!.row, last!.row), endRow: Math.max(first!.row + first!.rowSpan - 1, last!.row + last!.rowSpan - 1),
      startColumn: Math.min(first!.column, last!.column), endColumn: Math.max(first!.column + first!.colSpan - 1, last!.column + last!.colSpan - 1),
    })
    const selected = documentTableCellsInRange(grid, range)
    return { blockId: block!.id, ...range, rowCount: grid.rowCount, columnCount: grid.columnCount,
      hasMergedCells: grid.anchors.some(owner => owner.rowSpan > 1 || owner.colSpan > 1),
      canMergeCells: selected.length > 1, canUnmergeCells: selected.some(owner => owner.rowSpan > 1 || owner.colSpan > 1) }
  } catch { return undefined }
}

export function resolveDocumentTableCommandTarget(model: DocumentModel, selection: DocumentSelection, command: DocumentTableCommand): DocumentTableMutationTarget {
  const selected = documentTableSelection(model, selection), blockId = command.tableId ?? selected?.blockId
  const block = blockId && findDocumentBlock(model, blockId), table = block && tableForBlock(block)
  if (!block || !table) throw new RangeError('请先将光标或选区放在同一个表格内')
  const grid = documentTableGrid(table)
  let range = command.range ?? (selected?.blockId === block.id ? {
    startRow: selected.startRow, endRow: selected.endRow, startColumn: selected.startColumn, endColumn: selected.endColumn,
  } : documentTableWholeRange(grid))
  if (command.id === 'doc.command.table-set-cell-text') range = { startRow: command.row, endRow: command.row, startColumn: command.column, endColumn: command.column }
  assertDocumentTableRange(grid, range)
  return { blockId: block.id, range: { ...range } }
}

export interface DocumentTableActionState { available: boolean; reason?: string }

/** UI and API use the same grid and scope rules; execution also checks the complete detached result. */
export function documentTableCommandState(model: DocumentModel, selection: DocumentSelection, command: DocumentTableCommand, permissionContext?: DocumentPermissionContext): DocumentTableActionState {
  try {
    const target = resolveDocumentTableCommandTarget(model, selection, command)
    const context = permissionContext ?? { documentId: model.id ?? '', userId: 'local-user' }
    assertDocumentTableMutation(model, selection, command, target, context)
    return { available: true }
  } catch (error) { return { available: false, reason: error instanceof Error ? error.message : String(error) } }
}

/** Publish only after model, anchors and permissions are ready; history remains document-owned. */
export function executeDocumentTableCommand(model: DocumentModel, selection: DocumentSelection, command: DocumentTableCommand, permissionContext?: DocumentPermissionContext): DocumentSelection {
  const target = resolveDocumentTableCommandTarget(model, selection, command)
  const plan = prepareDocumentTableMutation(model, selection, command, target, permissionContext ?? { documentId: model.id ?? '', userId: 'local-user' })
  if (plan.changed) Object.assign(model, plan.model)
  return plan.selection
}
