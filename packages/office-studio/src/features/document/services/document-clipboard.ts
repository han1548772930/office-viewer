import type { DocumentBlock } from '../../../domain/model'
import type { DocumentCommand } from '../commands/document-commands'
import { blocksFromHtml, runText, tableForBlock } from '../model/document-model'

export interface DocumentClipboardData { text: string; html?: string }
export interface DocumentClipboardReader {
  readText(): Promise<string>
  read?(): Promise<readonly Pick<ClipboardItem, 'types' | 'getType'>[]>
}

/** Preserve HTML at the clipboard boundary; commands receive real runs and blocks. */
export function documentPasteCommand(data: DocumentClipboardData, plainText = false): DocumentCommand | undefined {
  const blocks = data.html ? blocksFromHtml(data.html) : []
  if (!plainText && blocks.length) return { id: 'doc.command.paste-blocks', blocks }
  const text = data.text || (blocks.length ? blocks.map(blockPlainText).join('\n') : '')
  return text ? { id: 'doc.command.insert-text', text } : undefined
}

/** Browsers without HTML clipboard reads still support a real plain-text paste. */
export async function readDocumentClipboard(clipboard: DocumentClipboardReader, plainText = false): Promise<DocumentClipboardData> {
  if (!plainText && clipboard.read) {
    try {
      const items = await clipboard.read()
      const data: DocumentClipboardData = { text: '' }
      for (const item of items) {
        if (item.types.includes('text/html') && data.html === undefined) data.html = await (await item.getType('text/html')).text()
        if (item.types.includes('text/plain') && !data.text) data.text = await (await item.getType('text/plain')).text()
      }
      if (data.html || data.text) return data
    } catch {
      // Permission and browser support differ for read() and readText(); try the text contract too.
    }
  }
  return { text: await clipboard.readText() }
}

function blockPlainText(block: DocumentBlock): string {
  const table = tableForBlock(block)
  return table ? table.rows.map((row) => row.cells.map((cell) => cell.blocks.map(blockPlainText).join('\n')).join('\t')).join('\n') : runText(block.runs)
}
