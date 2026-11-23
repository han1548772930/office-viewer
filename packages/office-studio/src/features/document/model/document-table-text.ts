import type { DocumentBlock } from '../../../domain/model'
import type { DocumentTable } from './document-model'
import { assertDocumentTableRange, documentTableGrid, documentTableWholeRange, expandDocumentTableRange, type DocumentTableCellRange } from './document-table-grid'
import { documentTableCellBlocks } from './document-table-continuations'

function blockText(block: DocumentBlock): string {
  const table = block.metadata.table as DocumentTable | undefined
  if (block.type === 'table' && table?.rows) return documentTableValuesToTsv(documentTableTextMatrix(table))
  return block.runs.map(run => run.text).join('')
}

/** Covered slots remain empty so copy preserves logical rows/columns and spans. */
export function documentTableTextMatrix(table: DocumentTable, requested?: DocumentTableCellRange): string[][] {
  const grid = documentTableGrid(table), range = requested ?? documentTableWholeRange(grid)
  assertDocumentTableRange(grid, range)
  const selected = expandDocumentTableRange(grid, range)
  return Array.from({ length: selected.endRow - selected.startRow + 1 }, (_, y) =>
    Array.from({ length: selected.endColumn - selected.startColumn + 1 }, (_, x) => {
      const row = selected.startRow + y, column = selected.startColumn + x, owner = grid.cellAt(row, column)!
      return owner.row === row && owner.column === column ? documentTableCellBlocks(owner.cell).map(blockText).join('\n') : ''
    }))
}

export function documentTableValuesToTsv(values: readonly (readonly string[])[]): string {
  const cell = (text: string) => /[\t\r\n"]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text
  return values.map(row => row.map(cell).join('\t')).join('\n')
}

/** Office TSV quoting keeps embedded newlines, tabs, quotes, and trailing empty cells. */
export function documentTableValuesFromTsv(text: string): string[][] | undefined {
  if (!text.length) return undefined
  const source = text.replace(/\r\n?/g, '\n'), rows: string[][] = [], row: string[] = []
  let value = '', quoted = false, afterQuote = false
  for (let i = 0; i < source.length; i++) {
    const character = source[i]!
    if (quoted) {
      if (character !== '"') value += character
      else if (source[i + 1] === '"') { value += '"'; i++ }
      else { quoted = false; afterQuote = true }
      continue
    }
    if (character === '\t' || character === '\n') {
      row.push(value); value = ''; afterQuote = false
      if (character === '\n') { rows.push([...row]); row.length = 0 }
    } else if (character === '"' && !value && !afterQuote) quoted = true
    else {
      if (afterQuote) throw new TypeError('TSV 引号结束后只能出现分隔符')
      value += character
    }
  }
  if (quoted) throw new TypeError('TSV 单元格引号没有闭合')
  row.push(value); rows.push(row)
  return rows
}
