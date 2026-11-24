import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { tableForBlock, type DocumentCell } from './document-model'
import { documentTableCellContentContainers } from './document-table-continuations'
import { documentTextFieldParagraphs, type DocumentTextFieldParagraph } from './document-text-fields'

export interface DocumentTableTextOriginCell {
  tableId: string
  cell: DocumentCell
  /** Source DFS order includes cells with no text paragraphs or blocks. */
  order: number
  /** Number of source text paragraphs before this cell's contents. */
  paragraphIndex: number
}

/** A derived projection of the same stored containers used by the text-field paragraph traversal. */
export function documentTableTextOriginCells(model: Pick<DocumentModel, 'blocks'>,
  paragraphs: readonly DocumentTextFieldParagraph[] = documentTextFieldParagraphs(model)): DocumentTableTextOriginCell[] {
  const textIds = new Set(paragraphs.map(paragraph => paragraph.block.id)), cells: DocumentTableTextOriginCell[] = []
  let paragraphIndex = 0, order = 0
  const visit = (blocks: readonly DocumentBlock[]) => {
    for (const block of blocks) {
      if (textIds.has(block.id)) paragraphIndex++
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) {
        cells.push({ tableId: block.id, cell, order: order++, paragraphIndex })
        for (const content of documentTableCellContentContainers(cell)) visit(content)
      }
    }
  }
  visit(model.blocks)
  return cells
}
