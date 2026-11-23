import type { DocumentBlock } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import type { FDocument } from './facade'
import type { DocumentTableCommand, DocumentTableInsertOptions, DocumentTableBorderOptions, DocumentTableStylePatch } from '../features/document/commands/table-command-types'
import { documentTableCommandState, documentTableSelection } from '../features/document/commands/table-commands'
import { documentTableReferenceWidth } from '../features/document/commands/table-mutations'
import {
  documentId, findDocumentBlock, tableForBlock, textTargets, type DocumentSelection, type DocumentTable,
  type DocumentTableCellMargin, type DocumentTableColumnTypeConfig, type DocumentTableViewSelection,
} from '../features/document/model/document-model'
import {
  assertDocumentTableRange, documentTableGrid, documentTableWholeRange, expandDocumentTableRange,
  type DocumentTableCellRange, type DocumentTableGridCell,
} from '../features/document/model/document-table-grid'
import {
  documentTableCellMargin, documentTableColumnWidths, documentTableRowProperties, DOCUMENT_TABLE_DEFAULT_BORDER,
  type DocumentTableRowPropertiesPatch,
} from '../features/document/model/document-table-style'
import { createDocumentTableViewSelection, documentTableDescendantIds } from '../features/document/model/document-table-selection'
import { flattenDocumentBody } from '../features/document/services/document-facade'
import { documentTableCellBlocks, documentTableCellContentContainers } from '../features/document/model/document-table-continuations'

export const DocsTableSelectionKinds = { Table: 'table', Row: 'row', Column: 'column', Cell: 'cell', Range: 'range' } as const
export const DocsTableRowInsertPositions = { Above: 'before', Below: 'after' } as const
export const DocsTableColumnInsertPositions = { Left: 'before', Right: 'after' } as const
export const DocsTableSortDirections = { Asc: 'ascending', Desc: 'descending' } as const
export const DocsTableColumnTypes = { Text: 'text', Number: 'number', Date: 'date', Checkbox: 'checkbox', Dropdown: 'dropdown' } as const
export const DocsTableRowHeightRules = { Auto: 'auto', AtLeast: 'atLeast', Exact: 'exact' } as const
export const DocsTableBorderPresets = { All: 'all', Inner: 'inner', Outer: 'outer', Top: 'top', Bottom: 'bottom',
  Left: 'left', Right: 'right', None: 'none', InnerHorizontal: 'innerHorizontal', InnerVertical: 'innerVertical' } as const

export interface DocumentTableFindQuery { tableId?: string; text?: string; headerText?: string; titleText?: string }
export interface DocumentTableMetadata extends Record<string, unknown> {
  headerRowCount?: number
  titleRow?: boolean
  columnTypes?: Record<number, DocumentTableColumnTypeConfig>
}
export interface DocumentTableInfo { id: string; rowCount: number; columnCount: number; metadata?: DocumentTableMetadata; source?: DocumentTable }
export interface DocumentTableDescription {
  id: string; rowCount: number; columnCount: number; headerRowCount: number; titleRow: boolean
  columns: Array<{ index: number; type?: string; width?: number }>
  sampleRows: string[][]
}
/** Half-open project body offsets; no Univer TABLE/CELL delimiter characters. */
export interface DocumentTableCellContentRange { startOffset: number; endOffset: number; segmentId: string }
export interface DocumentTableCellOffsetRange { row: number; column: number; startOffset: number | null; endOffset: number | null }
export interface DocumentTableOffsetRange {
  tableId: string; startOffset: number; endOffset: number; rowCount: number; columnCount: number; tableSource: DocumentTable
  offsetContract: 'body-synthetic-newline'
  rows: Array<{ row: number; startOffset: number | null; endOffset: number | null; cells: DocumentTableCellOffsetRange[] }>
}

function tableBlocks(blocks: readonly DocumentBlock[]): DocumentBlock[] {
  const result: DocumentBlock[] = []
  const visit = (items: readonly DocumentBlock[]) => {
    for (const block of items) {
      const table = tableForBlock(block)
      if (!table) continue
      result.push(block)
      for (const row of table.rows) for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) visit(content)
    }
  }
  visit(blocks)
  return result
}

function currentSelection(document: FDocument): DocumentSelection {
  const session = document.getInputState()
  if (session?.valid && session.location.kind === 'body') return session.location.selection
  const target = textTargets(document.getModel())[0]
  if (!target) throw new RangeError('文档没有可用的输入定位点')
  return { anchor: { ...target, offset: 0 }, focus: { ...target, offset: 0 } }
}

function interval(start: number, count: number | undefined, size: number, defaultToEnd: boolean): { start: number; end: number } {
  const amount = count ?? (defaultToEnd ? size - start : 1)
  if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(amount) || amount < 1 || start + amount > size) throw new RangeError('表格行列范围无效')
  return { start, end: start + amount - 1 }
}

export function getDocumentTables(document: FDocument, segmentId = ''): FDocumentTable[] {
  if (segmentId) return [] // Header/footer blocks remain an explicit, unimplemented model target.
  return tableBlocks(document.getModel().blocks).map(block => new FDocumentTable(document, block.id))
}

export function findDocumentTables(document: FDocument, query: string | DocumentTableFindQuery, segmentId = ''): FDocumentTable[] {
  const filter = typeof query === 'string' ? { text: query } : query
  if (!filter || typeof filter !== 'object' || [filter.tableId, filter.text, filter.headerText, filter.titleText].some(value => value !== undefined && typeof value !== 'string')) throw new TypeError('表格查询参数无效')
  return getDocumentTables(document, segmentId).filter(table => {
    if (filter.tableId !== undefined && table.getId() !== filter.tableId) return false
    const rows = (from: number, to: number) => {
      const text: string[] = []
      for (let row = from; row < to; row++) for (let column = 0; column < table.getColumnCount(); column++) text.push(table.getCellText(row, column))
      return text.join('\n')
    }
    return (filter.text === undefined || rows(0, table.getRowCount()).includes(filter.text))
      && (filter.headerText === undefined || rows(0, Math.max(table.getHeaderRowCount(), table.hasTitleRow() ? 1 : 0)).includes(filter.headerText))
      && (filter.titleText === undefined || table.hasTitleRow() && rows(0, 1).includes(filter.titleText))
  })
}

/** Initial values, geometry and metadata publish in one document history item. */
export function insertDocumentTable(document: FDocument, rows: number | undefined, columns: number | undefined,
  options: DocumentTableInsertOptions = {}): FDocumentTable | null {
  const tableId = options.tableId ?? documentId('table')
  const result = document.executeCommand({ ...options, id: 'doc.command.insert-table', tableId,
    ...(rows === undefined ? {} : { rows }), ...(columns === undefined ? {} : { cols: columns }) }, currentSelection(document))
  const block = findDocumentBlock(document.getModel(), tableId)
  return result.changed && block && tableForBlock(block) ? new FDocumentTable(document, tableId) : null
}

/** Stable table ID; every getter and command resolves the current snapshot. */
export class FDocumentTable {
  constructor(private readonly document: FDocument, private readonly tableId: string) {}
  getId(): string { return this.tableId }
  getSegmentId(): string { return '' }
  private source(): DocumentTable | undefined { const block = findDocumentBlock(this.document.getModel(), this.tableId); return block && tableForBlock(block) }
  private grid() { const source = this.source(); if (!source) throw new RangeError('表格已不存在'); return documentTableGrid(source) }
  private owner(row: number, column: number): DocumentTableGridCell | undefined {
    const source = this.source()
    return source && documentTableGrid(source).cellAt(row, column)
  }
  private execute(command: DocumentTableCommand): boolean {
    const selection = currentSelection(this.document), requested = { ...command, tableId: this.tableId }
    const state = documentTableCommandState(this.document.getModel(), selection, requested, { documentId: this.document.id, userId: this.document.getCurrentUserId() })
    if (!state.available) return false
    // The production command service repeats preflight after before-command observers.
    // It owns history and input advancement; stale returned selections are never restored here.
    return this.document.executeCommand(requested, selection).changed
  }
  getCommandState(command: DocumentTableCommand) {
    return documentTableCommandState(this.document.getModel(), currentSelection(this.document), { ...command, tableId: this.tableId },
      { documentId: this.document.id, userId: this.document.getCurrentUserId() })
  }
  getSource(): DocumentTable | undefined { const source = this.source(); return source && cloneModel(source) }
  getMetadata(): DocumentTableMetadata | undefined {
    const block = findDocumentBlock(this.document.getModel(), this.tableId), source = block && tableForBlock(block)
    if (!source) return undefined
    const { table: _table, html: _html, ...metadata } = block!.metadata
    return cloneModel({ ...metadata, headerRowCount: source.headerRowCount ?? 0, titleRow: source.titleRow ?? false,
      ...(source.columnTypes ? { columnTypes: source.columnTypes } : {}) })
  }
  getInfo(): DocumentTableInfo { return { id: this.tableId, rowCount: this.getRowCount(), columnCount: this.getColumnCount(), metadata: this.getMetadata(), source: this.getSource() } }
  getRowCount(): number { return this.source()?.rows.length ?? 0 }
  getColumnCount(): number { const source = this.source(); return source ? documentTableGrid(source).columnCount : 0 }
  describe(): DocumentTableDescription {
    const rows = this.getRowCount(), columns = this.getColumnCount(), source = this.source()
    return { id: this.tableId, rowCount: rows, columnCount: columns, headerRowCount: this.getHeaderRowCount(), titleRow: this.hasTitleRow(),
      columns: Array.from({ length: columns }, (_, index) => ({ index, type: source?.columnTypes?.[index]?.type, width: this.getColumnWidth(index) ?? undefined })),
      sampleRows: Array.from({ length: Math.min(5, rows) }, (_, row) => Array.from({ length: columns }, (_, column) => this.getCellText(row, column))) }
  }
  getCell(row: number, column: number): FDocumentTableCell | null {
    const owner = this.owner(row, column)
    return owner ? new FDocumentTableCell(this, owner.cell.id) : null
  }
  getCellById(cellId: string): FDocumentTableCell | null {
    return this.source() && this.grid().anchors.some(owner => owner.cell.id === cellId) ? new FDocumentTableCell(this, cellId) : null
  }
  resolveCellId(cellId: string): DocumentTableGridCell | undefined {
    const source = this.source(), owner = source && documentTableGrid(source).anchors.find(candidate => candidate.cell.id === cellId)
    return owner && cloneModel(owner)
  }
  getCellText(row: number, column: number): string {
    const owner = this.owner(row, column)
    if (!owner) return ''
    const ids = documentTableDescendantIds(documentTableCellBlocks(owner.cell))
    return flattenDocumentBody(this.document.getModel()).segments.filter(segment => ids.has(segment.cellBlockId ?? segment.blockId)).map(segment => segment.text).join('\n')
  }
  getCellMargin(row: number, column: number): DocumentTableCellMargin | null { const owner = this.owner(row, column), source = this.source(); return owner && source ? documentTableCellMargin(source, owner.cell) : null }
  getCellMarginOverride(row: number, column: number): DocumentTableCellMargin | null { const margin = this.owner(row, column)?.cell.style.margin; return margin ? { ...margin } : null }
  getDefaultCellMargin(): DocumentTableCellMargin { return documentTableCellMargin(this.source() ?? { rows: [] }) }
  getDefaultCellMarginOverride(): DocumentTableCellMargin | null { const margin = this.source()?.style?.cellMargin; return margin ? { ...margin } : null }
  getCellContentRange(row: number, column: number): DocumentTableCellContentRange | null {
    const owner = this.owner(row, column)
    if (!owner) return null
    const ids = documentTableDescendantIds(documentTableCellBlocks(owner.cell)), segments = flattenDocumentBody(this.document.getModel()).segments.filter(segment => ids.has(segment.cellBlockId ?? segment.blockId))
    const first = segments[0], last = segments.at(-1)
    return first && last ? { startOffset: first.startOffset, endOffset: last.endOffset, segmentId: '' } : null
  }
  getCellInsertOffset(row: number, column: number): number | null { return this.getCellContentRange(row, column)?.startOffset ?? null }
  getRange(): DocumentTableOffsetRange | null {
    const source = this.source()
    if (!source) return null
    const grid = documentTableGrid(source), ids = documentTableDescendantIds([findDocumentBlock(this.document.getModel(), this.tableId)!])
    const paragraphs = flattenDocumentBody(this.document.getModel()).segments.filter(segment => ids.has(segment.cellBlockId ?? segment.blockId))
    const first = paragraphs[0], last = paragraphs.at(-1)
    if (!first || !last) return null
    return { tableId: this.tableId, startOffset: first.startOffset, endOffset: last.endOffset, rowCount: grid.rowCount, columnCount: grid.columnCount,
      tableSource: cloneModel(source), offsetContract: 'body-synthetic-newline', rows: grid.rows.map((owners, row) => {
        const cells = owners.map(owner => { const range = this.getCellContentRange(owner.row, owner.column); return { row: owner.row, column: owner.column,
          startOffset: range?.startOffset ?? null, endOffset: range?.endOffset ?? null } })
        const ranges = cells.filter((cell): cell is DocumentTableCellOffsetRange & { startOffset: number; endOffset: number } => cell.startOffset !== null && cell.endOffset !== null)
        return { row, startOffset: ranges[0]?.startOffset ?? null, endOffset: ranges.at(-1)?.endOffset ?? null, cells }
      }) }
  }
  getTableRange(): DocumentTableCellRange | null { const source = this.source(); return source ? documentTableWholeRange(documentTableGrid(source)) : null }
  getCellRange(row: number, column: number): DocumentTableCellRange | null {
    const owner = this.owner(row, column)
    return owner ? { startRow: owner.row, endRow: owner.row + owner.rowSpan - 1, startColumn: owner.column, endColumn: owner.column + owner.colSpan - 1 } : null
  }
  getRowRange(row: number): DocumentTableCellRange { return this.getRowsRange(row, 1) }
  getRowsRange(startRow: number, count?: number): DocumentTableCellRange {
    const grid = this.grid(), rows = interval(startRow, count, grid.rowCount, false)
    return { startRow: rows.start, endRow: rows.end, startColumn: 0, endColumn: grid.columnCount - 1 }
  }
  getColumnRange(column: number): DocumentTableCellRange { return this.getColumnsRange(column, 1) }
  getColumnsRange(startColumn: number, count?: number): DocumentTableCellRange {
    const grid = this.grid(), columns = interval(startColumn, count, grid.columnCount, false)
    return { startRow: 0, endRow: grid.rowCount - 1, startColumn: columns.start, endColumn: columns.end }
  }
  selectRange(range: DocumentTableCellRange, kind: DocumentTableViewSelection['kind'] = 'range'): boolean {
    const grid = this.grid(); assertDocumentTableRange(grid, range)
    return this.document.setInputSelection(createDocumentTableViewSelection(this.document.getModel(), this.tableId, expandDocumentTableRange(grid, range), kind,
      currentSelection(this.document).tableSelection?.pageAffinity, currentSelection(this.document)))
  }
  selectTable(): boolean { const range = this.getTableRange(); return range ? this.selectRange(range, 'table') : false }
  selectRow(row: number): boolean { return this.selectRange(this.getRowRange(row), 'row') }
  selectRows(row: number, count?: number): boolean { return this.selectRange(this.getRowsRange(row, count), 'row') }
  selectColumn(column: number): boolean { return this.selectRange(this.getColumnRange(column), 'column') }
  selectColumns(column: number, count?: number): boolean { return this.selectRange(this.getColumnsRange(column, count), 'column') }
  selectCell(row: number, column: number): boolean { const range = this.getCellRange(row, column); return range ? this.selectRange(range, 'cell') : false }
  insertRowBefore(row: number): boolean { return this.insertRowsBefore(row, 1) }
  insertRowAfter(row: number): boolean { return this.insertRowsAfter(row, 1) }
  insertRowsBefore(row: number, count: number): boolean { return this.execute({ id: 'doc.command.table-insert-row', row, count, position: 'before' }) }
  insertRowsAfter(row: number, count: number): boolean { return this.execute({ id: 'doc.command.table-insert-row', row, count, position: 'after' }) }
  appendRow(): boolean { return this.insertRowAfter(this.getRowCount() - 1) }
  insertColumnBefore(column: number): boolean { return this.insertColumnsBefore(column, 1) }
  insertColumnAfter(column: number): boolean { return this.insertColumnsAfter(column, 1) }
  insertColumnsBefore(column: number, count: number): boolean { return this.execute({ id: 'doc.command.table-insert-column', column, count, position: 'before' }) }
  insertColumnsAfter(column: number, count: number): boolean { return this.execute({ id: 'doc.command.table-insert-column', column, count, position: 'after' }) }
  appendColumn(): boolean { return this.insertColumnAfter(this.getColumnCount() - 1) }
  deleteRow(row: number): boolean { return this.deleteRows(row, 1) }
  deleteRows(row: number, count?: number): boolean { return this.execute({ id: 'doc.command.table-delete-rows', range: this.getRowsRange(row, count) }) }
  deleteColumn(column: number): boolean { return this.deleteColumns(column, 1) }
  deleteColumns(column: number, count?: number): boolean { return this.execute({ id: 'doc.command.table-delete-columns', range: this.getColumnsRange(column, count) }) }
  deleteTable(): boolean { return this.execute({ id: 'doc.command.table-delete' }) }
  moveTableBy(deltaX: number, deltaY: number): boolean { return this.execute({ id: 'doc.command.table-move', deltaX, deltaY }) }
  moveTableToOffset(offset: number): boolean { return this.execute({ id: 'doc.command.table-move', targetOffset: offset }) }
  moveRow(row: number, targetRow: number, position: 'before' | 'after' = 'before'): boolean { return this.moveRows(row, 1, targetRow, position) }
  moveRows(startRow: number, count: number, targetRow: number, position: 'before' | 'after' = 'before'): boolean { return this.execute({ id: 'doc.command.table-move-rows', fromRow: startRow, count, toRow: targetRow, position }) }
  moveColumn(column: number, targetColumn: number, position: 'before' | 'after' = 'before'): boolean { return this.moveColumns(column, 1, targetColumn, position) }
  moveColumns(startColumn: number, count: number, targetColumn: number, position: 'before' | 'after' = 'before'): boolean { return this.execute({ id: 'doc.command.table-move-columns', fromColumn: startColumn, count, toColumn: targetColumn, position }) }
  mergeCells(range: DocumentTableCellRange): boolean { return this.execute({ id: 'doc.command.table-merge-cells', range }) }
  unmergeCells(range: DocumentTableCellRange): boolean { return this.execute({ id: 'doc.command.table-unmerge-cells', range }) }
  setCellText(row: number, column: number, text: string): boolean { return this.execute({ id: 'doc.command.table-set-cell-text', row, column, text }) }
  setCellBackground(range: DocumentTableCellRange, color: string | null): boolean { return this.execute({ id: 'doc.command.table-set-background', range, color }) }
  setCellMargin(range: DocumentTableCellRange, margin: DocumentTableCellMargin | null): boolean { return this.execute({ id: 'doc.command.table-set-cell-margin', range, margin }) }
  setCellMarginPatch(range: DocumentTableCellRange, marginPatch: Partial<DocumentTableCellMargin>): boolean { return this.execute({ id: 'doc.command.table-set-cell-margin', range, marginPatch }) }
  setDefaultCellMargin(margin: DocumentTableCellMargin | null): boolean { return this.execute({ id: 'doc.command.table-set-cell-margin', scope: 'table', margin }) }
  setBorder(range: DocumentTableCellRange, options: DocumentTableBorderOptions): boolean {
    if (options.patch !== undefined) {
      if ([options.border, options.color, options.width, options.style].some(value => value !== undefined)) throw new TypeError('表格边框完整值与局部更改不能同时提供')
      return this.execute({ id: 'doc.command.table-set-border', range, preset: options.preset ?? 'all', borderPatch: options.patch })
    }
    return this.execute({ id: 'doc.command.table-set-border', range, preset: options.preset ?? 'all',
      border: { ...DOCUMENT_TABLE_DEFAULT_BORDER, ...options.border, ...(options.color !== undefined ? { color: options.color ?? DOCUMENT_TABLE_DEFAULT_BORDER.color } : {}),
        ...(options.width === undefined ? {} : { width: options.width }), ...(options.style === undefined ? {} : { style: options.style }) } })
  }
  setTableBackground(color: string | null): boolean { return this.execute({ id: 'doc.command.table-set-background', scope: 'table', color }) }
  setTableBorder(options: DocumentTableBorderOptions): boolean { const range = this.getTableRange(); return range ? this.setBorder(range, options) : false }
  setTableStyle(patch: DocumentTableStylePatch): boolean { return this.execute({ id: 'doc.command.table-set-style', patch }) }
  setRowHeight(row: number, height: number, heightRule?: 'auto' | 'atLeast' | 'exact', allowSplit?: boolean): boolean { return this.execute({ id: 'doc.command.table-resize-row', row, height, heightRule, allowSplit }) }
  setRowProperties(row: number, patch: DocumentTableRowPropertiesPatch): boolean { return this.execute({ ...patch, id: 'doc.command.table-resize-row', row, preserveHeightRule: patch.preserveHeightRule ?? true }) }
  getRowHeight(row: number): { height?: number; heightRule: 'auto' | 'atLeast' | 'exact'; allowSplit: boolean } | null {
    const source = this.source(); return source && Number.isSafeInteger(row) && row >= 0 && row < source.rows.length ? documentTableRowProperties(source, row) : null
  }
  getColumnWidth(column: number): number | null { const source = this.source(); return source && Number.isSafeInteger(column) && column >= 0 && column < this.getColumnCount()
    ? documentTableColumnWidths(source, this.getColumnCount(), documentTableReferenceWidth(this.document.getModel(), this.tableId))[column] ?? null : null }
  setColumnWidth(column: number, width: number): boolean { return this.execute({ id: 'doc.command.table-resize-column', column, width }) }
  distributeRows(startRow = 0, count?: number): boolean {
    const grid = this.grid(), rows = interval(startRow, count, grid.rowCount, true)
    return this.execute({ id: 'doc.command.table-distribute-rows', range: this.getRowsRange(rows.start, rows.end - rows.start + 1) })
  }
  distributeColumns(startColumn = 0, count?: number): boolean {
    const grid = this.grid(), columns = interval(startColumn, count, grid.columnCount, true)
    return this.execute({ id: 'doc.command.table-distribute-columns', range: this.getColumnsRange(columns.start, columns.end - columns.start + 1) })
  }
  getHeaderRowCount(): number { return this.source()?.headerRowCount ?? 0 }
  setHeaderRowCount(count: number): boolean { return this.execute({ id: 'doc.command.table-set-header-row-count', count }) }
  pinHeaderRows(count: number): boolean { return this.setHeaderRowCount(count) }
  hasTitleRow(): boolean { return this.source()?.titleRow ?? false }
  insertTitleRow(text = ''): boolean { return this.execute({ id: 'doc.command.table-insert-title-row', text }) }
  setColumnType(column: number, config: DocumentTableColumnTypeConfig): boolean { return this.execute({ id: 'doc.command.table-set-column-type', column, config }) }
  sortByColumn(column: number, direction: 'ascending' | 'descending'): boolean { return this.execute({ id: 'doc.command.table-sort', column, direction }) }
  setCellVerticalAlign(range: DocumentTableCellRange, value: 'top' | 'center' | 'bottom'): boolean { return this.execute({ id: 'doc.command.table-set-vertical-align', range, value }) }
  setCellTextAlign(range: DocumentTableCellRange, value: 'left' | 'center' | 'right' | 'justify'): boolean { return this.execute({ id: 'doc.command.table-set-text-align', range, value }) }
}

/** Stable cell ID follows moves; deletion or merge of another owner expires it. */
export class FDocumentTableCell {
  constructor(private readonly table: FDocumentTable, private readonly cellId: string) {}
  getId(): string { return this.cellId }
  getTable(): FDocumentTable { return this.table }
  private owner() { return this.table.resolveCellId(this.cellId) }
  getRowIndex(): number { return this.owner()?.row ?? -1 }
  getColumnIndex(): number { return this.owner()?.column ?? -1 }
  getText(): string { const owner = this.owner(); return owner ? this.table.getCellText(owner.row, owner.column) : '' }
  getContentRange(): DocumentTableCellContentRange | null { const owner = this.owner(); return owner ? this.table.getCellContentRange(owner.row, owner.column) : null }
  getInsertOffset(): number | null { return this.getContentRange()?.startOffset ?? null }
  getMargin(): DocumentTableCellMargin | null { const owner = this.owner(); return owner ? this.table.getCellMargin(owner.row, owner.column) : null }
  getMarginOverride(): DocumentTableCellMargin | null { const owner = this.owner(); return owner ? this.table.getCellMarginOverride(owner.row, owner.column) : null }
  private range(): DocumentTableCellRange | null { const owner = this.owner(); return owner ? this.table.getCellRange(owner.row, owner.column) : null }
  setMargin(margin: DocumentTableCellMargin): boolean { const range = this.range(); return range ? this.table.setCellMargin(range, margin) : false }
  clearMarginOverride(): boolean { const range = this.range(); return range ? this.table.setCellMargin(range, null) : false }
  setText(text: string): boolean { const owner = this.owner(); return owner ? this.table.setCellText(owner.row, owner.column, text) : false }
  setBackgroundColor(color: string | null): boolean { const range = this.range(); return range ? this.table.setCellBackground(range, color) : false }
  setBorder(border: DocumentTableBorderOptions): boolean { const range = this.range(); return range ? this.table.setBorder(range, border) : false }
  mergeTo(rowSpan: number, columnSpan: number): boolean {
    const owner = this.owner()
    if (!owner || !Number.isSafeInteger(rowSpan) || rowSpan < 1 || !Number.isSafeInteger(columnSpan) || columnSpan < 1) return false
    return this.table.mergeCells({ startRow: owner.row, endRow: owner.row + rowSpan - 1, startColumn: owner.column, endColumn: owner.column + columnSpan - 1 })
  }
  unmerge(): boolean { const range = this.range(); return range ? this.table.unmergeCells(range) : false }
  select(): boolean { const owner = this.owner(); return owner ? this.table.selectCell(owner.row, owner.column) : false }
}

export function getDocumentTableAtSelection(document: FDocument): FDocumentTable | null {
  const session = document.getInputState()
  if (!session?.valid || session.location.kind !== 'body') return null
  const selected = documentTableSelection(document.getModel(), session.location.selection)
  return selected ? new FDocumentTable(document, selected.blockId) : null
}
