import type { DocumentBlock } from '../domain/model'
import { tableForBlock } from '../features/document/model/document-model'
import { documentTableCellContentContainers } from '../features/document/model/document-table-continuations'
import { copyDocumentBlocksForPaste } from '../features/document/services/document-block-paste'
import { array, record } from './validation'

export type ParsedDocumentBlockNormalizer = (value: unknown, fallbackId: string) => DocumentBlock

/** Body and auxiliary stories retain the same paragraph/table/object order. */
export function normalizeParsedDocumentBlockGroup(values: readonly unknown[], prefix: string, normalize: ParsedDocumentBlockNormalizer, firstFallbackId?: string): DocumentBlock[] {
  return values.flatMap((value, index) => {
    const raw = record(value), block = normalize(raw, index === 0 && firstFallbackId !== undefined ? firstFallbackId : `${prefix}-${index}`)
    return [block, ...(block.type === 'table' ? [] : array(raw.children)
      .filter(child => record(child).type === 'figure')
      .map((child, childIndex) => {
        const object = normalize(child, `${block.id}-object-${childIndex}`)
        // Source identity, rather than an editable target ID: a physical Word
        // paragraph can project to an anchor plus several object blocks.
        object.metadata.officeSourceParagraphId = typeof raw.id === 'string' ? raw.id : block.id
        return object
      }))]
  })
}

function blockIds(blocks: readonly DocumentBlock[]): string[] {
  const result: string[] = []
  const visit = (content: readonly DocumentBlock[]) => {
    for (const block of content) {
      result.push(block.id)
      const table = tableForBlock(block)
      for (const row of table?.rows ?? []) for (const cell of row.cells) {
        for (const container of documentTableCellContentContainers(cell)) visit(container)
      }
    }
  }
  visit(blocks)
  return result
}

export function parsedDocumentBlockOwnership(blocks: readonly DocumentBlock[]): Set<string> {
  return new Set(blockIds(blocks))
}

/** An explicitly reused native part must not give two editable hosts ownership of one ID. */
export function claimParsedDocumentStoryBlocks(blocks: DocumentBlock[], claimed: Set<string>): { blocks: DocumentBlock[]; copied: boolean } {
  const copied = blockIds(blocks).some(id => claimed.has(id))
  const owned = copied ? copyDocumentBlocksForPaste(blocks) : blocks
  for (const id of blockIds(owned)) {
    if (claimed.has(id)) throw new TypeError('导入内容区的块标识重复。')
    claimed.add(id)
  }
  return { blocks: owned, copied }
}
