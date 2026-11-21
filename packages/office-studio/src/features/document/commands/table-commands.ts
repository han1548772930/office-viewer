import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { assertContentControlBlockRemovable, getDocumentContentControls } from '../services/content-controls'
import {
  clampPoint, collapsedSelection, createParagraph, documentId, tableForBlock, textTargets,
  type DocumentCell, type DocumentSelection, type DocumentTable,
} from '../model/document-model'

export type DocumentTableCommand =
  | { id: 'doc.command.table-insert-row'; position: 'before' | 'after' }
  | { id: 'doc.command.table-insert-column'; position: 'before' | 'after' }
  | { id: 'doc.command.table-delete-rows' }
  | { id: 'doc.command.table-delete-columns' }
  | { id: 'doc.command.table-delete' }

export interface DocumentTableSelection {
  blockId: string
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  rowCount: number
  columnCount: number
  hasMergedCells: boolean
}

/** Table coordinates come from the logical selection, never from a DOM table. */
export function documentTableSelection(model: DocumentModel, selection: DocumentSelection): DocumentTableSelection | undefined {
  const { anchor, focus } = selection
  if (anchor.blockId !== focus.blockId || anchor.row == null || focus.row == null || anchor.col == null || focus.col == null) return undefined
  const block = model.blocks.find((candidate) => candidate.id === focus.blockId && candidate.type === 'table')
  const table = block && tableForBlock(block)
  if (!table?.rows.length) return undefined
  for (const point of [anchor, focus]) {
    if (!Number.isInteger(point.row) || !Number.isInteger(point.col) || !table.rows[point.row!]?.cells[point.col!]) return undefined
  }
  return {
    blockId: block!.id, startRow: Math.min(anchor.row, focus.row), endRow: Math.max(anchor.row, focus.row),
    startColumn: Math.min(anchor.col, focus.col), endColumn: Math.max(anchor.col, focus.col),
    rowCount: table.rows.length, columnCount: Math.max(...table.rows.map((row) => row.cells.length)),
    hasMergedCells: table.rows.some((row) => row.cells.some((cell) => cell.rowSpan !== 1 || cell.colSpan !== 1)),
  }
}

/** One structural mutation; the caller records the complete draft in its document history. */
export function executeDocumentTableCommand(model: DocumentModel, selection: DocumentSelection, command: DocumentTableCommand): DocumentSelection {
  const context = documentTableSelection(model, selection)
  if (!context) throw new RangeError('请先将光标或选区放在同一个表格内')
  const block = model.blocks.find((candidate) => candidate.id === context.blockId)!
  const table = tableForBlock(block)!
  if (command.id === 'doc.command.table-delete') return deleteTable(model, block, table)
  const grid = validateTableGrid(table, context.columnCount)

  switch (command.id) {
    case 'doc.command.table-insert-row': {
      const index = command.position === 'before' ? context.startRow : context.endRow + 1
      const template = table.rows[command.position === 'before' ? context.startRow : context.endRow]!
      const row = { ...template, id: documentId('row'), cells: template.cells.map((cell) => emptyCellFrom(cell, cell.colSpan)) }
      table.rows.splice(index, 0, row)
      if (index < (table.headerRowCount ?? 0)) table.headerRowCount = (table.headerRowCount ?? 0) + 1
      delete block.metadata.html
      return cellSelection(block.id, table, index, context.startColumn)
    }
    case 'doc.command.table-insert-column': {
      const index = command.position === 'before' ? context.startColumn : context.endColumn + 1
      const templateIndex = command.position === 'before' ? context.startColumn : context.endColumn
      for (const row of table.rows) row.cells.splice(index, 0, emptyCellFrom(row.cells[templateIndex]!))
      if (table.columnWidths?.length === grid.columnCount) {
        const widthIndex = grid.columnOffsets[index]!
        const widthSourceIndex = command.position === 'before'
          ? grid.columnOffsets[templateIndex]!
          : grid.columnOffsets[templateIndex + 1]! - 1
        table.columnWidths = resizedColumnWidths(table.columnWidths, widthIndex, widthSourceIndex)
      }
      delete block.metadata.html
      return cellSelection(block.id, table, context.startRow, index)
    }
    case 'doc.command.table-delete-rows': {
      if (context.startRow === 0 && context.endRow === table.rows.length - 1) return deleteTable(model, block, table)
      const removing = table.rows.slice(context.startRow, context.endRow + 1).flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
      removeAnchoredControls(model, removing)
      const oldHeaderCount = table.headerRowCount ?? 0
      const removedHeaders = Math.max(0, Math.min(context.endRow + 1, oldHeaderCount) - context.startRow)
      table.rows.splice(context.startRow, context.endRow - context.startRow + 1)
      table.headerRowCount = Math.min(table.rows.length, oldHeaderCount - removedHeaders)
      delete block.metadata.html
      return cellSelection(block.id, table, Math.min(context.startRow, table.rows.length - 1), context.startColumn)
    }
    case 'doc.command.table-delete-columns': {
      if (context.startColumn === 0 && context.endColumn === context.columnCount - 1) return deleteTable(model, block, table)
      const removing = table.rows.flatMap((row) => row.cells.slice(context.startColumn, context.endColumn + 1).flatMap((cell) => cell.blocks))
      removeAnchoredControls(model, removing)
      const count = context.endColumn - context.startColumn + 1
      for (const row of table.rows) row.cells.splice(context.startColumn, count)
      if (table.columnWidths?.length === grid.columnCount) {
        const widthStart = grid.columnOffsets[context.startColumn]!
        const widthEnd = grid.columnOffsets[context.endColumn + 1]!
        table.columnWidths.splice(widthStart, widthEnd - widthStart)
      }
      delete block.metadata.html
      return cellSelection(block.id, table, context.startRow, Math.min(context.startColumn, context.columnCount - count - 1))
    }
  }
}

function emptyCellFrom(source: DocumentCell, colSpan = 1): DocumentCell {
  const paragraph = createParagraph()
  const sourceParagraph = source.blocks.find((block) => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type))
  if (sourceParagraph) paragraph.style = { ...sourceParagraph.style }
  return { ...source, id: documentId('cell'), blocks: [paragraph], style: JSON.parse(JSON.stringify(source.style)) as Record<string, unknown>, rowSpan: 1, colSpan }
}

interface DocumentTableGrid {
  columnCount: number
  columnOffsets: number[]
}

function validateTableGrid(table: DocumentTable, selectionColumnCount: number): DocumentTableGrid {
  if (!selectionColumnCount || table.rows.some((row) => row.cells.length !== selectionColumnCount)) {
    throw new RangeError('表格行列网格不规则，无法安全调整结构，原表格已保留')
  }
  if (table.rows.some((row) => row.cells.some((cell) => !Number.isInteger(cell.colSpan) || cell.colSpan < 1 || cell.rowSpan !== 1))) {
    throw new RangeError('含跨行合并或无效跨度的表格无法安全调整结构，原表格已保留')
  }
  const spans = table.rows[0]!.cells.map((cell) => cell.colSpan)
  if (table.rows.some((row) => row.cells.some((cell, index) => cell.colSpan !== spans[index]))) {
    throw new RangeError('表格各行的合并列结构不一致，无法安全调整结构，原表格已保留')
  }
  const columnOffsets = [0]
  for (const span of spans) columnOffsets.push(columnOffsets.at(-1)! + span)
  const columnCount = columnOffsets.at(-1)!
  if (table.columnWidths?.length && table.columnWidths.length !== columnCount) {
    throw new RangeError('表格列宽与合并网格不一致，无法安全调整结构，原表格已保留')
  }
  if (table.columnWidths?.some((width) => !Number.isFinite(width) || width <= 0)) {
    throw new RangeError('表格包含无效列宽，无法安全调整结构，原表格已保留')
  }
  return { columnCount, columnOffsets }
}

function resizedColumnWidths(widths: number[], insertAt: number, sourceIndex: number): number[] {
  const tableWidth = widths.reduce((sum, width) => sum + width, 0)
  const insertedSourceWidth = widths[sourceIndex]!
  const scale = tableWidth / (tableWidth + insertedSourceWidth)
  const resized = widths.map((width) => width * scale)
  resized.splice(insertAt, 0, insertedSourceWidth * scale)
  return resized
}

function cellSelection(blockId: string, table: DocumentTable, row: number, col: number): DocumentSelection {
  const cell = table.rows[row]!.cells[col]!
  let paragraph = cell.blocks.find((block) => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type))
  if (!paragraph) { paragraph = createParagraph(); cell.blocks.push(paragraph) }
  return collapsedSelection({ blockId, row, col, cellBlockId: paragraph.id, offset: 0 })
}

function deleteTable(model: DocumentModel, block: DocumentBlock, table: DocumentTable): DocumentSelection {
  const index = model.blocks.indexOf(block)
  removeAnchoredControls(model, [block, ...table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))])
  model.blocks.splice(index, 1)
  const targets = textTargets(model)
  const following = targets.find((target) => model.blocks.findIndex((candidate) => candidate.id === target.blockId) >= index)
  if (following) return collapsedSelection({ ...following, offset: 0 })
  const preceding = targets.at(-1)
  if (preceding) return collapsedSelection(clampPoint(model, { ...preceding, offset: Number.MAX_SAFE_INTEGER }))
  const paragraph = createParagraph()
  model.blocks.splice(index, 0, paragraph)
  return collapsedSelection({ blockId: paragraph.id, offset: 0 })
}

/** Preflight all targets on a detached plan so a later locked control cannot leave partial removals. */
function removeAnchoredControls(model: DocumentModel, blocks: readonly DocumentBlock[]): void {
  if (!model.contentControls?.length) return
  const plan = { ...model, contentControls: getDocumentContentControls(model) }
  const visit = (block: DocumentBlock) => {
    assertContentControlBlockRemovable(plan, block.id)
    for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const child of cell.blocks) visit(child)
  }
  for (const block of blocks) visit(block)
  model.contentControls = plan.contentControls
}
