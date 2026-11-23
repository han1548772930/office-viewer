import type { DocumentBlock, DocumentModel, DocumentParagraphStyle } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import type { DocumentCommand } from '../commands/document-commands'
import { blocksFromHtml, createParagraph, createRun, documentHtml, findDocumentBlock, runText, tableForBlock, type DocumentRun, type DocumentSelection, type DocumentTableViewSelection, type RunPatch } from '../model/document-model'
import { documentTableCellsInRange, documentTableGrid } from '../model/document-table-grid'
import { resolveDocumentTableViewSelection } from '../model/document-table-selection'
import { documentTableCellBorders, documentTableCellFill, documentTableCellMargin } from '../model/document-table-style'
import { documentTableValuesFromTsv, documentTableTextMatrix, documentTableValuesToTsv } from '../model/document-table-text'
import { documentTableCellContentContainers } from '../model/document-table-continuations'

export interface DocumentClipboardData { text: string; html?: string }
export const DOCUMENT_PASTE_OPTIONS = [
  { value: 'source', label: '保留源格式' }, { value: 'destination', label: '匹配目标格式' }, { value: 'text', label: '仅保留文本' },
] as const
export type DocumentPasteMode = typeof DOCUMENT_PASTE_OPTIONS[number]['value']
export interface DocumentPasteDestinationStyle {
  paragraphStyle: DocumentParagraphStyle
  textStyle: RunPatch
}
export interface DocumentClipboardReader {
  readText(): Promise<string>
  read?(): Promise<readonly Pick<ClipboardItem, 'types' | 'getType'>[]>
}
export interface DocumentClipboardWriter {
  writeText(text: string): Promise<void>
  write?(items: ClipboardItem[]): Promise<void>
}

/** Crop compact anchors while materializing effective formatting at the new clipboard edges. */
export function documentTableClipboardData(model: DocumentModel, selection: DocumentTableViewSelection): DocumentClipboardData {
  const view = resolveDocumentTableViewSelection(model, selection), block = findDocumentBlock(model, view.tableId)!
  const table = tableForBlock(block)!, grid = documentTableGrid(table), range = view.range
  const owners = documentTableCellsInRange(grid, range), byId = new Map(owners.map(owner => [owner.cell.id, owner]))
  const copied = cloneModel(block), cropped = tableForBlock(copied)!
  cropped.rows = cropped.rows.slice(range.startRow, range.endRow + 1)
  for (const row of cropped.rows) {
    row.cells = row.cells.filter(cell => byId.has(cell.id))
    for (const cell of row.cells) {
      const original = byId.get(cell.id)!
      cell.columnIndex = original.column - range.startColumn
      cell.style = { ...cell.style, fill: documentTableCellFill(table, original.cell),
        margin: documentTableCellMargin(table, original.cell), borders: documentTableCellBorders(grid, original) }
    }
  }
  cropped.headerRowCount = Math.max(0, Math.min(table.headerRowCount ?? 0, range.endRow + 1) - range.startRow)
  cropped.titleRow = table.titleRow === true && range.startRow === 0
  if (table.rowStyles) cropped.rowStyles = table.rowStyles.slice(range.startRow, range.endRow + 1).map(style => cloneModel(style))
  cropped.columnWidths = table.columnWidths?.slice(range.startColumn, range.endColumn + 1) ?? []
  if (table.columnTypes) cropped.columnTypes = Object.fromEntries(Object.entries(table.columnTypes)
    .filter(([column]) => Number(column) >= range.startColumn && Number(column) <= range.endColumn)
    .map(([column, config]) => [Number(column) - range.startColumn, cloneModel(config)]))
  const width = cropped.columnWidths.length ? cropped.columnWidths.reduce((sum, value) => sum + value, 0)
    : table.style?.width ? table.style.width * (range.endColumn - range.startColumn + 1) / grid.columnCount : undefined
  cropped.style = { ...cropped.style, align: 'left', indent: 0, offsetY: 0, ...(width === undefined ? {} : { width }) }
  delete copied.metadata.html
  documentTableGrid(cropped)
  return { text: documentTableValuesToTsv(documentTableTextMatrix(table, range)), html: documentHtml({ ...model, blocks: [copied] }) }
}

/** Return whether the rich payload was written; callers can preserve a cut when only text succeeded. */
export async function writeDocumentClipboard(clipboard: DocumentClipboardWriter, data: DocumentClipboardData): Promise<boolean> {
  if (data.html && clipboard.write && typeof ClipboardItem === 'function') {
    try {
      await clipboard.write([new ClipboardItem({ 'text/html': new Blob([data.html], { type: 'text/html' }),
        'text/plain': new Blob([data.text], { type: 'text/plain' }) })])
      return true
    } catch {
      // writeText has a separate browser capability and still provides a real copy.
    }
  }
  await clipboard.writeText(data.text)
  return !data.html
}

/** Preserve HTML at the clipboard boundary; commands receive real runs and blocks. */
export function documentPasteCommand(data: DocumentClipboardData, plainText = false): DocumentCommand | undefined {
  const blocks = data.html ? blocksFromHtml(data.html) : []
  if (!plainText && blocks.length) return { id: 'doc.command.paste-blocks', blocks }
  const text = data.text || (blocks.length ? blocks.map(blockPlainText).join('\n') : '')
  return text ? { id: 'doc.command.insert-text', text } : undefined
}

/** Plain and rich clipboard data use the same logical destination origin. */
export function documentTablePasteCommand(model: DocumentModel, selection: DocumentSelection, data: DocumentClipboardData, plainText = false,
  mode: 'source' | 'destination' = 'source'): DocumentCommand | undefined {
  const view = selection.tableSelection
  if (!view) return documentPasteCommand(data, plainText)
  if (data.html && !plainText) {
    const command = documentPasteCommand(data)
    if (command?.id === 'doc.command.paste-blocks') return { ...command, tablePasteMode: mode }
  }
  const text = data.text || (data.html ? blocksFromHtml(data.html).map(blockPlainText).join('\n') : '')
  const values = documentTableValuesFromTsv(text)
  if (!values) return undefined
  const block = findDocumentBlock(model, view.tableId), table = block && tableForBlock(block)
  if (!table) throw new RangeError('表格粘贴目标已不存在')
  const grid = documentTableGrid(table), columns = values.reduce((maximum, row) => Math.max(maximum, row.length), 0)
  const range = { ...view.range,
    endRow: Math.max(view.range.endRow, view.range.startRow + values.length - 1),
    endColumn: Math.max(view.range.endColumn, view.range.startColumn + columns - 1) }
  if (range.endRow >= grid.rowCount || range.endColumn >= grid.columnCount) throw new RangeError('粘贴数据超过现有表格，请先插入足够的行列')
  return { id: 'doc.command.table-set-range-text', tableId: view.tableId, range, mode: 'matrix', values }
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
    if (destinationStyle) applyDocumentPasteDestinationStyle(blocks, destinationStyle, false)
    return blocks
  }

  if (!destinationStyle) throw new TypeError('Destination formatting requires the current target paragraph and text style.')
  if (!sourceBlocks.length) return []
  const copiedBlocks = structuredClone([...sourceBlocks])
  applyDocumentPasteDestinationStyle(copiedBlocks, destinationStyle)
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
  return table ? documentTableValuesToTsv(documentTableTextMatrix(table)) : runText(block.runs)
}

function paragraphsFromPlainText(text: string): DocumentBlock[] {
  if (!text) return []
  return text.replace(/\r\n?/g, '\n').split('\n').map(line => createParagraph(line))
}

export function applyDocumentPasteDestinationStyle(blocks: DocumentBlock[], destination: DocumentPasteDestinationStyle, preserveSourceEmphasis = true): void {
  const { href: _destinationLink, metadata: _destinationMetadata, ...appearance } = destination.textStyle
  for (const block of blocks) {
    const table = tableForBlock(block)
    if (table) {
      for (const row of table.rows) {
        for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) applyDocumentPasteDestinationStyle(content, destination, preserveSourceEmphasis)
      }
      continue
    }
    if (['image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) continue

    const sourceType = block.type
    block.type = sourceType === 'heading' ? 'paragraph' : sourceType
    if (sourceType === 'heading') block.level = 0
    block.style = { ...destination.paragraphStyle }
    delete block.metadata.namedStyleType
    block.runs = block.runs.map((sourceRun: DocumentRun) => {
      const overrides: RunPatch = preserveSourceEmphasis ? {
        bold: sourceRun.bold,
        italic: sourceRun.italic,
        underline: sourceRun.underline,
        strike: sourceRun.strike,
        ...(sourceRun.verticalAlign === undefined ? {} : { verticalAlign: sourceRun.verticalAlign }),
      } : {}
      return createRun(sourceRun.text, { ...appearance, ...overrides,
        ...(preserveSourceEmphasis && sourceRun.href ? { href: sourceRun.href } : {}),
        ...(preserveSourceEmphasis && sourceRun.metadata ? { metadata: sourceRun.metadata } : {}),
      })
    })
  }
}
