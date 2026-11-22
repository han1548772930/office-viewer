import type { DocumentBlock, DocumentParagraphStyle } from '../../../domain/model'
import type { DocumentCommand } from '../commands/document-commands'
import { blocksFromHtml, createParagraph, createRun, runText, tableForBlock, type RunPatch } from '../model/document-model'

export interface DocumentClipboardData { text: string; html?: string }
export type DocumentPasteMode = 'source' | 'destination' | 'text'
export interface DocumentPasteDestinationStyle {
  paragraphStyle: DocumentParagraphStyle
  textStyle: RunPatch
}
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

/** Build a paste command from the clipboard without mutating its source blocks or HTML model. */
export function documentPasteSpecialCommand(
  data: DocumentClipboardData,
  mode: DocumentPasteMode,
  destinationStyle?: DocumentPasteDestinationStyle,
): DocumentCommand | undefined {
  if (mode === 'source') return documentPasteCommand(data)

  const sourceBlocks = data.html ? blocksFromHtml(data.html) : []
  const blocks = applyDocumentPasteMode(sourceBlocks.length ? sourceBlocks : paragraphsFromPlainText(data.text), mode, destinationStyle, data.text)
  return blocks.length ? { id: 'doc.command.paste-blocks', blocks } : undefined
}

export function applyDocumentPasteMode(
  sourceBlocks: readonly DocumentBlock[],
  mode: Exclude<DocumentPasteMode, 'source'>,
  destinationStyle?: DocumentPasteDestinationStyle,
  plainText = '',
): DocumentBlock[] {
  if (mode === 'text') {
    const text = sourceBlocks.some(block => tableForBlock(block))
      ? sourceBlocks.map(blockPlainText).join('\n')
      : plainText || sourceBlocks.map(blockPlainText).join('\n')
    const blocks = text ? paragraphsFromPlainText(text) : []
    if (destinationStyle) applyDestinationStyle(blocks, destinationStyle, false)
    return blocks
  }

  if (!destinationStyle) throw new TypeError('Destination formatting requires the current target paragraph and text style.')
  if (!sourceBlocks.length) return []
  const copiedBlocks = structuredClone([...sourceBlocks])
  applyDestinationStyle(copiedBlocks, destinationStyle)
  return copiedBlocks
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

function paragraphsFromPlainText(text: string): DocumentBlock[] {
  if (!text) return []
  return text.replace(/\r\n?/g, '\n').split('\n').map(line => createParagraph(line))
}

function applyDestinationStyle(blocks: DocumentBlock[], destination: DocumentPasteDestinationStyle, preserveSourceEmphasis = true): void {
  for (const block of blocks) {
    const table = tableForBlock(block)
    if (table) {
      for (const row of table.rows) {
        for (const cell of row.cells) applyDestinationStyle(cell.blocks, destination, preserveSourceEmphasis)
      }
      continue
    }
    if (['image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) continue

    const sourceType = block.type
    block.type = sourceType === 'heading' ? 'paragraph' : sourceType
    if (sourceType === 'heading') block.level = 0
    block.style = { ...destination.paragraphStyle }
    delete block.metadata.namedStyleType
    block.runs = block.runs.map((sourceRun) => {
      const overrides: RunPatch = preserveSourceEmphasis ? {
        bold: sourceRun.bold,
        italic: sourceRun.italic,
        underline: sourceRun.underline,
        strike: sourceRun.strike,
        ...(sourceRun.verticalAlign === undefined ? {} : { verticalAlign: sourceRun.verticalAlign }),
      } : {}
      return createRun(sourceRun.text, { ...destination.textStyle, ...overrides })
    })
  }
}
