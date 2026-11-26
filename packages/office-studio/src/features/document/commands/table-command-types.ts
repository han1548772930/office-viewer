import type { DocumentTableBorder, DocumentTableCellMargin, DocumentTableColumnTypeConfig, DocumentTableStyle, RunPatch } from '../model/document-model'
import type { DocumentTableCellRange } from '../model/document-table-grid'
import type { DocumentTableBorderPreset, DocumentTableRowPropertiesPatch } from '../model/document-table-style'

export interface DocumentTableCommandTarget { tableId?: string; range?: DocumentTableCellRange }

export interface DocumentTableBorderOptions {
  preset?: DocumentTableBorderPreset
  border?: DocumentTableBorder
  color?: string | null
  width?: number
  style?: DocumentTableBorder['style']
  /** Preserve each edge's unmodified properties, including mixed values. */
  patch?: Partial<DocumentTableBorder>
}

/** One style operation is one history item, including background and borders. */
export interface DocumentTableStylePatch {
  backgroundColor?: string | null
  border?: DocumentTableBorderOptions
  align?: DocumentTableStyle['align']
  indent?: number
  offsetY?: number
  width?: number
  /** Measured widths freeze an auto layout without restoring stale saved proportions. */
  columnWidths?: number[]
  layout?: DocumentTableStyle['layout']
  cellMargin?: DocumentTableCellMargin | null
}

export interface DocumentTableInsertOptions {
  tableId?: string
  position?: 'selection' | 'documentEnd' | 'offset'
  /** Facade body offsets use the project's synthetic-newline text contract. */
  offset?: number
  segmentId?: string
  width?: number
  columnWidths?: number[]
  rowHeights?: number[]
  cellMargin?: DocumentTableCellMargin
  titleRow?: boolean
  headerRowCount?: number
  values?: string[][]
  columnTypes?: Record<number, DocumentTableColumnTypeConfig>
  style?: DocumentTableStylePatch
  metadata?: Record<string, unknown>
}

export type DocumentTableInsertCommand = DocumentTableInsertOptions & { id: 'doc.command.insert-table'; rows?: number; cols?: number }
export type DocumentTableCreateCommand = DocumentTableInsertOptions & { id: 'doc.command.create-table'; rowCount: number; colCount: number }

export const DOCUMENT_TABLE_COMMAND_IDS = [
  'doc.command.table-insert-row', 'doc.command.table-insert-column', 'doc.command.table-delete-rows', 'doc.command.table-delete-columns', 'doc.command.table-delete',
  'doc.command.table-merge-cells', 'doc.command.table-unmerge-cells', 'doc.command.table-move-rows', 'doc.command.table-move-columns', 'doc.command.table-move',
  'doc.command.table-resize-row', 'doc.command.table-resize-column', 'doc.command.table-distribute-rows', 'doc.command.table-distribute-columns',
  'doc.command.table-set-background', 'doc.command.table-set-border', 'doc.command.table-set-border-color', 'doc.command.table-set-border-width', 'doc.command.table-set-border-style',
  'doc.command.table-set-vertical-align', 'doc.command.table-set-text-align', 'doc.command.table-set-align', 'doc.command.table-set-cell-margin', 'doc.command.table-set-cell-text',
  'doc.command.table-set-header-row-count', 'doc.command.table-insert-title-row', 'doc.command.table-set-column-type', 'doc.command.table-sort',
  'doc.command.table-set-style',
  'doc.command.table-set-range-text',
] as const

export type DocumentTableCommand = DocumentTableCommandTarget & (
  | { id: 'doc.command.table-insert-row'; position: 'before' | 'after'; row?: number; count?: number }
  | { id: 'doc.command.table-insert-column'; position: 'before' | 'after'; column?: number; count?: number }
  | { id: 'doc.command.table-delete-rows' }
  | { id: 'doc.command.table-delete-columns' }
  | { id: 'doc.command.table-delete' }
  | { id: 'doc.command.table-merge-cells' }
  | { id: 'doc.command.table-unmerge-cells' }
  | { id: 'doc.command.table-move-rows'; fromRow?: number; toRow: number; count?: number; position: 'before' | 'after' }
  | { id: 'doc.command.table-move-columns'; fromColumn?: number; toColumn: number; count?: number; position: 'before' | 'after' }
  | { id: 'doc.command.table-move'; targetBlockId?: string; targetOffset?: number; position?: 'before' | 'after'; deltaX?: number; deltaY?: number }
  | ({ id: 'doc.command.table-resize-row'; row?: number } & DocumentTableRowPropertiesPatch)
  | { id: 'doc.command.table-resize-column'; column?: number; width: number; widths?: number[] }
  | { id: 'doc.command.table-distribute-rows'; heights?: number[] }
  | { id: 'doc.command.table-distribute-columns' }
  | { id: 'doc.command.table-set-background'; color: string | null; scope?: 'table' | 'selection' }
  | { id: 'doc.command.table-set-border'; preset: DocumentTableBorderPreset; border?: DocumentTableBorder; borderPatch?: Partial<DocumentTableBorder> }
  | { id: 'doc.command.table-set-border-color'; color: string | null }
  | { id: 'doc.command.table-set-border-width'; width: number }
  | { id: 'doc.command.table-set-border-style'; style: 'solid' | 'dot' | 'dash' }
  | { id: 'doc.command.table-set-vertical-align'; value: 'top' | 'center' | 'bottom' }
  | { id: 'doc.command.table-set-text-align'; value: 'left' | 'center' | 'right' | 'justify' }
  | { id: 'doc.command.table-set-align'; value: 'left' | 'center' | 'right' }
  | { id: 'doc.command.table-set-cell-margin'; margin?: DocumentTableCellMargin | null; marginPatch?: Partial<DocumentTableCellMargin>; scope?: 'table' | 'selection' }
  | { id: 'doc.command.table-set-cell-text'; row: number; column: number; text: string; optionId?: string }
  | { id: 'doc.command.table-set-range-text'; mode: 'clear' | 'replace' | 'matrix'; text?: string; values?: string[][]; marks?: RunPatch }
  | { id: 'doc.command.table-set-header-row-count'; count: number }
  | { id: 'doc.command.table-insert-title-row'; text?: string }
  | { id: 'doc.command.table-set-column-type'; column?: number; config: DocumentTableColumnTypeConfig }
  | { id: 'doc.command.table-sort'; column?: number; direction: 'ascending' | 'descending'; locale?: string }
  | { id: 'doc.command.table-set-style'; patch: DocumentTableStylePatch }
)

export function isDocumentTableCommand(command: { id: string }): command is DocumentTableCommand {
  return (DOCUMENT_TABLE_COMMAND_IDS as readonly string[]).includes(command.id)
}
