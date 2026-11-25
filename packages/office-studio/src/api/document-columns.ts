import type { DocumentColumn, DocumentColumnGroup } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import type { FDocument, FDocumentParagraph, FDocumentTextRange } from './facade'
import { documentId, findDocumentBlock } from '../features/document/model/document-model'
import { documentTextFieldParagraphs } from '../features/document/model/document-text-fields'
import { DOCUMENT_COLUMN_HORIZONTAL_PADDING, DOCUMENT_COLUMN_MIN_HEIGHT } from '../features/document/model/document-structural-blocks'
import { DOCUMENT_BODY_TOKENS as T, type UniverDocumentColumnGroup } from './document-body-types'
import { assertDocumentPublicRange, documentSelectionAtPublicBodyRange } from './document-public-body'
import { FDocumentStructuredBlock, documentStructuredHost, documentStructuredSelection, type DocumentStructuredRange } from './document-structured-blocks'
import type { DocumentStructureCommand } from '../features/document/commands/document-structure-commands'
import type { DocumentTextStylePatch } from '../features/document/services/document-facade'
import type { UniverDocumentTextStyle } from './document-text-style'

export const DocsColumnPositions = { LEFT: 'left', RIGHT: 'right' } as const
export interface DocumentColumnFindQuery { columnGroupId?: string; text?: string }
/** Public lengths use points; the model and Canvas use CSS pixels. */
export interface DocumentColumnInsertOptions { columnGroupId?: string; columnIds?: string[]; gap?: number; horizontalPadding?: number; minHeight?: number; offset?: number; widthRatios?: number[]; segmentId?: string }
export interface DocumentColumnOffsetRange { column: number; columnId: string; startOffset: number; endOffset: number }
export interface DocumentColumnGroupOffsetRange extends DocumentStructuredRange {
  columnGroupId: string; startOffset: number; endOffset: number; columns: DocumentColumnOffsetRange[]; source: UniverDocumentColumnGroup
}
export interface DocumentColumnContentRange { startOffset: number; endOffset: number; segmentId: string }
const columnLengthInPixels = (value: number, label: string): number => {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) throw new RangeError(`${label}必须是非负的有限 pt 数值。`)
  return value * 4 / 3
}

/** Column ranges retain the public inclusive column delimiter / exclusive group end contract. */
export function documentColumnRanges(stream: string, group: Pick<UniverDocumentColumnGroup, 'columnGroupId' | 'startIndex' | 'endIndex'>, columnIds: readonly string[]): DocumentColumnOffsetRange[] {
  if (stream[group.startIndex] !== T.groupStart || stream[group.endIndex] !== T.groupEnd) throw new RangeError('多栏组标记与公开范围不一致。')
  const ranges: DocumentColumnOffsetRange[] = []
  let cursor = group.startIndex + 1
  const skipRangeMarkers = () => { while (cursor < group.endIndex && (stream[cursor] === T.rangeStart || stream[cursor] === T.rangeEnd)) cursor++ }
  for (const [column, columnId] of columnIds.entries()) {
    skipRangeMarkers()
    const startOffset = cursor
    if (stream[cursor] !== T.columnStart) throw new RangeError('栏缺少开始标记。')
    let depth = 1
    cursor++
    while (cursor < group.endIndex && depth) {
      if (stream[cursor] === T.columnStart) depth++
      else if (stream[cursor] === T.columnEnd) depth--
      cursor++
    }
    if (depth) throw new RangeError('栏缺少匹配的结束标记。')
    ranges.push({ column, columnId, startOffset, endOffset: cursor - 1 })
  }
  skipRangeMarkers()
  if (cursor !== group.endIndex) throw new RangeError('栏数量与多栏组公开范围不一致。')
  return ranges
}

export class FDocumentColumnGroup extends FDocumentStructuredBlock {
  constructor(document: FDocument, columnGroupId: string, segmentId = '') { super(document, columnGroupId, 'columns', segmentId) }
  getConfig(): UniverDocumentColumnGroup | undefined {
    const config = this.document.getPublicBodyIndex(this.segmentId).data.body.columnGroups?.find(group => group.columnGroupId === this.blockId)
    return config && cloneModel(config)
  }
  getRange(): DocumentColumnGroupOffsetRange | null {
    const block = this.source()
    if (!block?.columns) return null
    const index = this.document.getPublicBodyIndex(this.segmentId), source = index.data.body.columnGroups?.find(group => group.columnGroupId === this.blockId)
    if (!source) return null
    return { blockId: this.blockId, columnGroupId: this.blockId, startIndex: source.startIndex, endIndex: source.endIndex,
      startOffset: source.startIndex, endOffset: source.endIndex + 1,
      columns: documentColumnRanges(index.data.body.dataStream, source, block.columns.columns.map(column => column.id)), source: cloneModel(source) }
  }
  getColumnCount(): number { return this.source()?.columns?.columns.length ?? 0 }
  getColumns(): FDocumentColumn[] { return this.source()?.columns?.columns.map(column => new FDocumentColumn(this.document, this.blockId, column.id, this.segmentId)) ?? [] }
  getColumn(indexOrId: number | string): FDocumentColumn | null {
    const columns = this.source()?.columns?.columns ?? []
    const column = typeof indexOrId === 'number' ? Number.isSafeInteger(indexOrId) && indexOrId >= 0 ? columns[indexOrId] : undefined : columns.find(column => column.id === indexOrId)
    return column ? new FDocumentColumn(this.document, this.blockId, column.id, this.segmentId) : null
  }
  getWidthRatios(): number[] { return this.source()?.columns?.columns.map(column => column.widthRatio) ?? [] }
  setWidthRatios(widthRatios: number[]): boolean { return this.execute({ id: 'doc.command.update-column-group', blockId: this.blockId, widthRatios: [...widthRatios] }) }
  setGap(gap: number): boolean {
    return this.execute({ id: 'doc.command.update-column-group', blockId: this.blockId, patch: { gap: columnLengthInPixels(gap, '栏间距') } })
  }
  getHorizontalPadding(): number { return (this.source()?.columns?.horizontalPadding ?? DOCUMENT_COLUMN_HORIZONTAL_PADDING) * 3 / 4 }
  setHorizontalPadding(horizontalPadding: number): boolean {
    return this.execute({ id: 'doc.command.update-column-group', blockId: this.blockId, patch: { horizontalPadding: columnLengthInPixels(horizontalPadding, '栏水平内边距') } })
  }
  getMinHeight(): number { return (this.source()?.columns?.minHeight ?? DOCUMENT_COLUMN_MIN_HEIGHT) * 3 / 4 }
  setMinHeight(minHeight: number): boolean {
    return this.execute({ id: 'doc.command.update-column-group', blockId: this.blockId, patch: { minHeight: columnLengthInPixels(minHeight, '栏最小高度') } })
  }
  setLayout(layout: DocumentColumnGroup['layout'], responsive?: DocumentColumnGroup['responsive']): boolean {
    return this.execute({ id: 'doc.command.update-column-group', blockId: this.blockId, patch: { layout, ...(responsive === undefined ? {} : { responsive }) } })
  }
  addColumn(targetColumnId: string, position: 'left' | 'right', columnId = documentId('column')): FDocumentColumn | null {
    if (position !== 'left' && position !== 'right') throw new TypeError('栏插入位置无效。')
    return this.execute({ id: 'doc.command.add-column', blockId: this.blockId, columnId: targetColumnId, newColumnId: columnId, position: position === 'left' ? 'before' : 'after' })
      ? this.getColumn(columnId) : null
  }
  deleteColumn(columnId: string): boolean { return this.execute({ id: 'doc.command.remove-column', blockId: this.blockId, columnId, keepContents: false }) }
  moveColumn(columnId: string, index: number): boolean { return this.execute({ id: 'doc.command.move-column', blockId: this.blockId, columnId, index }) }
  describe() {
    const group = this.source()?.columns
    return { columnGroupId: this.blockId, columnCount: group?.columns.length ?? 0, gap: (group?.gap ?? 0) * 3 / 4,
      horizontalPadding: this.getHorizontalPadding(), minHeight: this.getMinHeight(),
      layout: group?.layout, responsive: group?.responsive, config: this.getConfig(),
      columns: group?.columns.map((column, index) => ({ columnId: column.id, index, text: documentTextFieldParagraphs({ blocks: column.blocks }).map(paragraph => paragraph.text).join('\n'), widthRatio: column.widthRatio })) ?? [] }
  }
}

export class FDocumentColumn {
  constructor(private readonly document: FDocument, private readonly columnGroupId: string, private readonly columnId: string, private readonly segmentId = '') {}
  getId(): string { return this.columnId }
  getColumnGroupId(): string { return this.columnGroupId }
  getIndex(): number { return this.group()?.columns.findIndex(column => column.id === this.columnId) ?? -1 }
  private group(): DocumentColumnGroup | undefined { return findDocumentBlock(this.document.getSegmentModel(this.segmentId), this.columnGroupId)?.columns }
  private source(): DocumentColumn | undefined { return this.group()?.columns.find(column => column.id === this.columnId) }
  getRange(): DocumentColumnOffsetRange | null { return new FDocumentColumnGroup(this.document, this.columnGroupId, this.segmentId).getRange()?.columns.find(column => column.columnId === this.columnId) ?? null }
  getContentRange(): DocumentColumnContentRange | null {
    const range = this.getRange()
    return range ? { startOffset: range.startOffset + 1, endOffset: range.endOffset, segmentId: this.segmentId } : null
  }
  getInsertOffset(): number | null {
    const range = this.getContentRange()
    if (!range) return null
    const stream = this.document.getBody(this.segmentId).dataStream
    return stream[range.endOffset - 1] === T.paragraph ? range.endOffset - 1 : range.endOffset
  }
  getText(): string { const source = this.source(); return source ? documentTextFieldParagraphs({ blocks: source.blocks }).map(paragraph => paragraph.text).join('\n') : '' }
  getParagraphs(): FDocumentParagraph[] {
    const source = this.source(), ids = new Set(source ? documentTextFieldParagraphs({ blocks: source.blocks }).map(paragraph => paragraph.block.id) : [])
    return this.document.getParagraphs(this.segmentId).filter(paragraph => ids.has(paragraph.getId()))
  }
  getTextRange(): FDocumentTextRange | null {
    const range = this.getContentRange()
    return range ? this.document.getTextRange(range.startOffset, range.endOffset, this.segmentId) : null
  }
  private execute(command: DocumentStructureCommand): boolean {
    const block = findDocumentBlock(this.document.getSegmentModel(this.segmentId), this.columnGroupId)
    if (!block?.columns || !this.source()) return false
    const host = documentStructuredHost(this.document, block)
    return this.document.executeCommand({ ...command, ...(host ? { host } : {}) }, documentStructuredSelection(this.document, block, this.segmentId)).changed
  }
  setText(text: string): boolean { return this.execute({ id: 'doc.command.set-column-text', blockId: this.columnGroupId, columnId: this.columnId, text }) }
  setTextStyle(style: UniverDocumentTextStyle | DocumentTextStylePatch): boolean { return this.getTextRange()?.setTextStyle(style) ?? false }
  insertParagraph(index: number, text = ''): FDocumentParagraph {
    const paragraphId = documentId('paragraph')
    const changed = this.execute({ id: 'doc.command.insert-column-paragraph', blockId: this.columnGroupId, columnId: this.columnId, index, text, paragraphId })
    const paragraph = changed && this.document.getParagraph(paragraphId, this.segmentId)
    if (!paragraph) throw new RangeError('栏内段落插入失败。')
    return paragraph
  }
  appendParagraph(text = ''): FDocumentParagraph { return this.insertParagraph(this.getParagraphs().length, text) }
}

/** Facade lengths use points; the domain and Canvas geometry use CSS pixels. */
export function insertDocumentColumnGroup(document: FDocument, columnCount: number, options: DocumentColumnInsertOptions = {}): FDocumentColumnGroup | null {
  const segmentId = options.segmentId ?? '', model = document.getSegmentModel(segmentId), columnGroupId = options.columnGroupId ?? documentId('columns')
  const selection = documentStructuredSelection(document, undefined, segmentId), hostBlock = model.blocks[0]
  const host = segmentId && hostBlock ? documentStructuredHost(document, hostBlock) : undefined
  let insertion: Pick<Extract<DocumentStructureCommand, { id: 'doc.command.insert-column-group' }>, 'afterBlockId' | 'at'>
  if (options.offset !== undefined) {
    const index = document.getPublicBodyIndex(segmentId)
    assertDocumentPublicRange(index, options.offset, options.offset)
    const requested = documentSelectionAtPublicBodyRange(document.getModel(), options.offset, options.offset, host)
    if (requested) insertion = { at: requested.focus }
    else if (options.offset === index.data.body.dataStream.length - 1 && model.blocks.at(-1)) insertion = { afterBlockId: model.blocks.at(-1)!.id }
    else throw new RangeError('多栏组插入位置必须落在真实段落或文档结束边界。')
  } else insertion = model.blocks.at(-1) ? { afterBlockId: model.blocks.at(-1)!.id } : { at: selection.focus }
  const result = document.executeCommand({ id: 'doc.command.insert-column-group', columnGroupId, columnCount, ...insertion,
    ...(options.columnIds === undefined ? {} : { columnIds: [...options.columnIds] }), ...(options.widthRatios === undefined ? {} : { widthRatios: [...options.widthRatios] }),
    ...(options.gap === undefined ? {} : { gap: columnLengthInPixels(options.gap, '栏间距') }),
    ...(options.horizontalPadding === undefined ? {} : { horizontalPadding: columnLengthInPixels(options.horizontalPadding, '栏水平内边距') }),
    ...(options.minHeight === undefined ? {} : { minHeight: columnLengthInPixels(options.minHeight, '栏最小高度') }), ...(host ? { host } : {}) }, selection)
  return result.changed && findDocumentBlock(document.getSegmentModel(segmentId), columnGroupId)?.columns ? new FDocumentColumnGroup(document, columnGroupId, segmentId) : null
}
