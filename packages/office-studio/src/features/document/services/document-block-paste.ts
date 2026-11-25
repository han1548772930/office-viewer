import type { DocumentBlock } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentId, tableForBlock } from '../model/document-model'
import { documentTableGrid } from '../model/document-table-grid'
import { documentTableCellContentContainers, rebindDocumentTableContinuations } from '../model/document-table-continuations'
import { documentTextFieldRanges, documentTextFields, remapDocumentTextFieldIds } from '../model/document-text-fields'
import { documentStructuredChildContainers, validateDocumentStructuredBlock } from '../model/document-structural-blocks'

/** Clipboard blocks are borrowed input; IDs and continuation bindings belong to each new paste. */
export function copyDocumentBlocksForPaste(blocks: readonly DocumentBlock[]): DocumentBlock[] {
  let blockCount = 0, cellCount = 0
  const sourceIds = new Set<string>()
  const inspect = (items: readonly DocumentBlock[], depth: number) => {
    if (depth > 64) throw new RangeError('粘贴内容超过文档嵌套预算')
    for (const block of items) {
      if (!block || typeof block.id !== 'string' || !block.id || sourceIds.has(block.id) || !Array.isArray(block.runs)
        || !block.style || !block.metadata) throw new TypeError('复制范围必须包含标识唯一的文档块')
      sourceIds.add(block.id)
      validateDocumentStructuredBlock(block)
      if (++blockCount > 250_000) throw new RangeError('粘贴内容超过单次文档块预算')
      for (const content of documentStructuredChildContainers(block)) inspect(content, depth + 1)
      const table = tableForBlock(block)
      if (!table) continue
      const grid = documentTableGrid(table)
      cellCount += grid.rowCount * grid.columnCount
      if (cellCount > 250_000) throw new RangeError('粘贴内容超过单次单元格预算')
      for (const owner of grid.anchors) for (const content of documentTableCellContentContainers(owner.cell)) inspect(content, depth + 1)
    }
  }
  inspect(blocks, 0)
  documentTextFieldRanges({ blocks: [...blocks] })
  const copies = cloneModel([...blocks]), blockIds = new Map<string, string>(), paragraphs: DocumentBlock[] = []
  const listIds = new Map<string, string>()
  const visit = (container: DocumentBlock[]) => {
    for (const block of container) {
      const nextId = documentId(block.type)
      blockIds.set(block.id, nextId)
      block.id = nextId
      block.runs = block.runs.map(run => ({ ...run, id: documentId('run') }))
      delete block.metadata.html
      if (block.checklist) {
        const previous = block.checklist.listId
        if (!listIds.has(previous)) listIds.set(previous, documentId('checklist'))
        block.checklist.listId = listIds.get(previous)!
      }
      paragraphs.push(block)
      if (block.columns) for (const column of block.columns.columns) column.id = documentId('column')
      for (const content of documentStructuredChildContainers(block)) visit(content)
      const table = tableForBlock(block)
      if (!table) continue
      // Read before changing row/cell IDs so invalid source continuations cannot be silently discarded.
      documentTableGrid(table)
      const rowIds = new Map(table.rows.map(row => [row.id, documentId('row')]))
      for (const row of table.rows) row.id = rowIds.get(row.id)!
      for (const [rowIndex, row] of table.rows.entries()) for (const cell of row.cells) {
        const previousCellId = cell.id, contents = documentTableCellContentContainers(cell)
        cell.id = documentId('cell')
        for (const content of contents) visit(content)
        rebindDocumentTableContinuations(table, cell, rowIndex, previousCellId, rowIds)
      }
    }
  }
  visit(copies)
  const fieldIds = new Map<string, string>()
  for (const block of paragraphs) for (const field of documentTextFields(block)) fieldIds.set(field.id, documentId('field'))
  for (const block of paragraphs) remapDocumentTextFieldIds(block, blockIds, fieldIds)
  documentTextFieldRanges({ blocks: copies })
  return copies
}
