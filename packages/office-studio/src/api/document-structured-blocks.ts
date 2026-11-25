import type { DocumentBlock, DocumentCalloutConfig, DocumentCodeConfig } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import type { FDocument, FDocumentParagraph } from './facade'
import type { DocumentStructureCommand } from '../features/document/commands/document-structure-commands'
import { collapsedSelection, documentId, findDocumentBlock, runText, selectionCollapsed, textTargets, type DocumentSelection } from '../features/document/model/document-model'
import { documentBlockHostForId, sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import { documentTextFieldParagraphs } from '../features/document/model/document-text-fields'
import { defaultDocumentCalloutConfig, defaultDocumentCodeConfig, visitDocumentBlocks } from '../features/document/model/document-structural-blocks'
import { codeBlockData } from '../features/document/code-blocks/code-blocks'
import { documentQuoteStyle, documentStructuredTextColor, type DocumentQuoteStyle } from '../features/document/model/document-structured-appearance'
import { assertDocumentPublicRange, documentSelectionAtPublicBodyRange } from './document-public-body'

export type DocumentStructuredKind = 'quote' | 'callout' | 'code' | 'columns'
export interface DocumentStructuredFindQuery { blockId?: string; text?: string }
export interface DocumentStructuredRange { blockId: string; startIndex: number; endIndex: number }
export interface DocumentStructuredInsertOptions { blockId?: string; startOffset?: number; endOffset?: number; segmentId?: string }
export interface DocumentQuoteInsertOptions extends DocumentStructuredInsertOptions { style?: Partial<DocumentQuoteStyle> }
export interface DocumentCalloutInsertOptions extends DocumentStructuredInsertOptions { config?: Partial<DocumentCalloutConfig>; textColor?: string | null }
export interface DocumentCodeInsertOptions extends DocumentStructuredInsertOptions { config?: Partial<DocumentCodeConfig> }
export interface DocumentCalloutBorderStyle { color: string; opacity: number; style: 1 | 2 | 3; width: number }
export type DocumentCalloutBorderPatch = Partial<Omit<DocumentCalloutBorderStyle, 'style'>> & { style?: 1 | 2 | 3 | DocumentCalloutConfig['borderStyle'] }
export interface DocumentCalloutStyle { backgroundColor: string; border: DocumentCalloutBorderStyle; textColor: string }

function matches(block: DocumentBlock, kind: DocumentStructuredKind): boolean {
  return kind === 'code' ? !!codeBlockData(block) : block.type === kind
}
export function documentStructuredBlocks(document: FDocument, kind: DocumentStructuredKind, segmentId = ''): DocumentBlock[] {
  const result: DocumentBlock[] = []
  visitDocumentBlocks(document.getSegmentModel(segmentId).blocks, block => { if (matches(block, kind)) result.push(block) })
  return result
}
export function documentStructuredText(block: DocumentBlock): string {
  return documentTextFieldParagraphs({ blocks: [block] }).map(paragraph => paragraph.text).join('\n')
}
export function findDocumentStructuredBlocks(document: FDocument, kind: DocumentStructuredKind, query: string | DocumentStructuredFindQuery, segmentId = ''): DocumentBlock[] {
  const filter: DocumentStructuredFindQuery = typeof query === 'string' ? { text: query } : query
  if (!filter || typeof filter !== 'object' || Array.isArray(filter)
    || [filter.blockId, filter.text].some(value => value !== undefined && typeof value !== 'string')) throw new TypeError('高级块查询参数无效。')
  return documentStructuredBlocks(document, kind, segmentId).filter(block => (filter.blockId === undefined || block.id === filter.blockId)
    && (filter.text === undefined || documentStructuredText(block).includes(filter.text)))
}
export function documentStructuredBlockAt(document: FDocument, kind: DocumentStructuredKind, offset: number, segmentId = ''): DocumentBlock | undefined {
  const index = document.getPublicBodyIndex(segmentId)
  assertDocumentPublicRange(index, offset, offset)
  const ranges = kind === 'columns' ? index.data.body.columnGroups?.map(range => ({ blockId: range.columnGroupId, ...range })) ?? []
    : index.data.body.blockRanges?.filter(range => range.blockType === kind) ?? []
  const range = ranges.filter(range => offset >= range.startIndex && offset <= range.endIndex)
    .sort((left, right) => left.endIndex - left.startIndex - (right.endIndex - right.startIndex))[0]
  if (range) return findDocumentBlock(index.model, range.blockId)
  return index.positions.find(position => offset >= position.start && offset <= position.end && matches(position.block, kind))?.block
}

export function documentStructuredHost(document: FDocument, block: DocumentBlock): DocumentAuxiliaryBlockHost | undefined {
  return documentBlockHostForId(document.getModel(), block.id)
}
export function documentStructuredSelection(document: FDocument, block?: DocumentBlock, segmentId = ''): DocumentSelection {
  const model = document.getSegmentModel(segmentId), targets = textTargets(model), state = document.getInputState()
  const host = block ? documentStructuredHost(document, block) : segmentId && model.blocks[0] ? documentStructuredHost(document, model.blocks[0]) : undefined
  if (state?.valid && state.location.kind === 'body' && sameDocumentBlockHost(state.location.host, host)) {
    const { selection } = state.location
    const exists = (point: DocumentSelection['focus']) => targets.some(target => (target.cellBlockId ?? target.blockId) === (point.cellBlockId ?? point.blockId))
    if (exists(selection.anchor) && exists(selection.focus)) return cloneModel(selection)
  }
  const ids = block && new Set(documentTextFieldParagraphs({ blocks: [block] }).map(paragraph => paragraph.block.id))
  const target = targets.find(target => !ids || ids.has(target.cellBlockId ?? target.blockId)) ?? targets[0]
  if (!target) throw new RangeError('文档没有真实的输入段落。')
  return collapsedSelection({ ...target, offset: 0 })
}

/** ID-based wrappers resolve the current model and never retain a mutable tree snapshot. */
export class FDocumentStructuredBlock {
  constructor(protected readonly document: FDocument, protected readonly blockId: string, protected readonly kind: DocumentStructuredKind, protected readonly segmentId = '') {}
  getId(): string { return this.blockId }
  getSegmentId(): string { return this.segmentId }
  protected source(): DocumentBlock | undefined {
    const block = findDocumentBlock(this.document.getSegmentModel(this.segmentId), this.blockId)
    return block && matches(block, this.kind) ? block : undefined
  }
  getRange(): DocumentStructuredRange | null {
    if (!this.source()) return null
    const index = this.document.getPublicBodyIndex(this.segmentId), range = index.data.body.blockRanges?.find(range => range.blockId === this.blockId)
    if (range) return { blockId: this.blockId, startIndex: range.startIndex, endIndex: range.endIndex }
    const legacy = index.positions.find(position => position.block.id === this.blockId)
    return legacy ? { blockId: this.blockId, startIndex: legacy.start, endIndex: legacy.end } : null
  }
  getText(): string { const block = this.source(); return block ? documentStructuredText(block) : '' }
  getParagraphs(): FDocumentParagraph[] {
    const block = this.source(), ids = new Set(block ? documentTextFieldParagraphs({ blocks: [block] }).map(paragraph => paragraph.block.id) : [])
    return this.document.getParagraphs(this.segmentId).filter(paragraph => ids.has(paragraph.getId()))
  }
  protected execute(command: DocumentStructureCommand): boolean {
    const block = this.source()
    if (!block) return false
    const host = documentStructuredHost(this.document, block)
    return this.document.executeCommand({ ...command, ...(host ? { host } : {}) }, documentStructuredSelection(this.document, block, this.segmentId)).changed
  }
  unwrap(): boolean { return this.execute({ id: 'doc.command.unwrap-structured-block', blockId: this.blockId }) }
  remove(): boolean { return this.execute({ id: 'doc.command.remove-structured-block', blockId: this.blockId }) }
  exit(): boolean { return this.execute({ id: 'doc.command.exit-structured-block', blockId: this.blockId }) }
}

export class FDocumentQuote extends FDocumentStructuredBlock {
  constructor(document: FDocument, blockId: string, segmentId = '') { super(document, blockId, 'quote', segmentId) }
  getStyle(): DocumentQuoteStyle { const block = this.source(); return block ? documentQuoteStyle(block) : { lineColor: '#8796aa', textColor: '#1f2937' } }
  setStyle(style: Partial<DocumentQuoteStyle>): boolean { return this.execute({ id: 'doc.command.update-quote-style', blockId: this.blockId, patch: cloneModel(style) }) }
  setLineColor(lineColor: string): boolean { return this.setStyle({ lineColor }) }
  setTextColor(textColor: string): boolean { return this.setStyle({ textColor }) }
  describe() { const range = this.getRange(); return range ? { ...range, type: 'quote' as const, text: this.getText(), style: this.getStyle() } : null }
}

export class FDocumentCallout extends FDocumentStructuredBlock {
  constructor(document: FDocument, blockId: string, segmentId = '') { super(document, blockId, 'callout', segmentId) }
  getConfig(): DocumentCalloutConfig { return cloneModel(this.source()?.callout ?? defaultDocumentCalloutConfig()) }
  getStyle(): DocumentCalloutStyle {
    const config = this.getConfig(), block = this.source()
    return { backgroundColor: config.backgroundColor, textColor: block ? documentStructuredTextColor(block) : '#1f2937',
      border: { color: config.borderColor, opacity: config.borderOpacity, width: config.borderWidth, style: config.borderStyle === 'dotted' ? 2 : config.borderStyle === 'dashed' ? 3 : 1 } }
  }
  updateConfig(config: Partial<DocumentCalloutConfig>): boolean { return this.execute({ id: 'doc.command.update-callout', blockId: this.blockId, patch: cloneModel(config) }) }
  setBackgroundColor(backgroundColor: string): boolean { return this.updateConfig({ backgroundColor }) }
  setBorder(border: DocumentCalloutBorderPatch): boolean {
    const style = border.style === undefined ? undefined : border.style === 1 ? 'solid' : border.style === 2 ? 'dotted' : border.style === 3 ? 'dashed' : border.style
    return this.updateConfig({ ...(border.color === undefined ? {} : { borderColor: border.color }), ...(border.opacity === undefined ? {} : { borderOpacity: border.opacity }),
      ...(border.width === undefined ? {} : { borderWidth: border.width }), ...(style === undefined ? {} : { borderStyle: style }) })
  }
  setTextColor(textColor: string): boolean { return this.execute({ id: 'doc.command.update-callout', blockId: this.blockId, patch: {}, textColor }) }
  resetTextColor(): boolean { return this.execute({ id: 'doc.command.update-callout', blockId: this.blockId, patch: {}, textColor: null }) }
  setIcon(icon: string): boolean { return this.updateConfig({ icon }) }
  setIconVisible(showIcon: boolean): boolean { return this.updateConfig({ showIcon }) }
  describe() { const range = this.getRange(); return range ? { ...range, type: 'callout' as const, text: this.getText(), config: this.getConfig(), style: this.getStyle() } : null }
}

export class FDocumentCode extends FDocumentStructuredBlock {
  constructor(document: FDocument, blockId: string, segmentId = '') { super(document, blockId, 'code', segmentId) }
  getConfig(): DocumentCodeConfig {
    const block = this.source(), data = block && codeBlockData(block)
    return data ? { language: data.language, showLineNumbers: data.lineNumbers, tabSize: data.tabSize ?? 4, wrap: data.wrap ?? true } : defaultDocumentCodeConfig()
  }
  updateConfig(options: { config: Partial<DocumentCodeConfig> }): boolean { return this.execute({ id: 'doc.command.update-code-block', blockId: this.blockId, patch: cloneModel(options.config) }) }
  setText(source: string): boolean { return this.execute({ id: 'doc.command.update-code-block', blockId: this.blockId, source }) }
  describe() { const range = this.getRange(); return range ? { ...range, type: 'code' as const, text: this.getText(), config: this.getConfig() } : null }
}

function isParagraph(value: FDocumentParagraph | DocumentStructuredInsertOptions | undefined): value is FDocumentParagraph {
  return !!value && typeof value === 'object' && 'getInfo' in value && typeof value.getInfo === 'function'
}
function insertDecorated(document: FDocument, blockType: 'quote' | 'callout' | 'code', paragraph: FDocumentParagraph | undefined,
  options: DocumentStructuredInsertOptions, config: { callout?: Partial<DocumentCalloutConfig>; code?: Partial<DocumentCodeConfig>; quote?: Partial<DocumentQuoteStyle>; textColor?: string | null }): string | null {
  if (paragraph && paragraph.getDocumentId() !== document.getId()) throw new RangeError('段落属于另一份文档。')
  const segmentId = paragraph?.segmentId ?? options.segmentId ?? '', model = document.getSegmentModel(segmentId)
  if (paragraph && options.segmentId !== undefined && options.segmentId !== paragraph.segmentId) throw new RangeError('段落与指定的内容区不一致。')
  const original = paragraph ? findDocumentBlock(model, paragraph.getId()) : undefined
  if (paragraph && !original) throw new RangeError('要包裹的段落已经不存在。')
  const hostBlock = original ?? model.blocks[0], host = segmentId && hostBlock ? documentStructuredHost(document, hostBlock) : undefined
  let selection = documentStructuredSelection(document, original, segmentId)
  if (paragraph) {
    const target = textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === paragraph.getId())
    if (!target) throw new RangeError('段落缺少真实输入目标。')
    selection = { anchor: { ...target, offset: 0 }, focus: { ...target, offset: runText(original!.runs).length } }
  } else if (options.startOffset !== undefined || options.endOffset !== undefined) {
    const start = options.startOffset ?? options.endOffset!, end = options.endOffset ?? start
    const requested = documentSelectionAtPublicBodyRange(document.getModel(), start, end, host)
    if (!requested) throw new RangeError('高级块插入范围必须落在真实段落中。')
    selection = requested
  }
  const blockId = options.blockId ?? documentId(blockType), wrap = !!paragraph || !selectionCollapsed(selection)
  const command: DocumentStructureCommand = wrap
    ? { id: 'doc.command.wrap-structured-block', blockId, blockType, ...config, ...(paragraph ? { blockIds: [paragraph.getId()] } : {}) }
    : { id: 'doc.command.insert-structured-block', blockId, blockType, at: selection.focus, ...config }
  const result = document.executeCommand({ ...command, ...(host ? { host } : {}) }, selection)
  return result.changed && findDocumentBlock(document.getSegmentModel(segmentId), blockId) ? blockId : null
}
export function insertDocumentQuote(document: FDocument, paragraphOrOptions: FDocumentParagraph | DocumentQuoteInsertOptions = {}, options: DocumentQuoteInsertOptions = {}): FDocumentQuote | null {
  const paragraph = isParagraph(paragraphOrOptions) ? paragraphOrOptions : undefined, value = isParagraph(paragraphOrOptions) ? options : paragraphOrOptions
  const id = insertDecorated(document, 'quote', paragraph, value, { ...(value.style ? { quote: value.style } : {}) })
  return id ? new FDocumentQuote(document, id, paragraph?.segmentId ?? value.segmentId ?? '') : null
}
export function insertDocumentCallout(document: FDocument, paragraphOrOptions: FDocumentParagraph | DocumentCalloutInsertOptions = {}, options: DocumentCalloutInsertOptions = {}): FDocumentCallout | null {
  const paragraph = isParagraph(paragraphOrOptions) ? paragraphOrOptions : undefined, value = isParagraph(paragraphOrOptions) ? options : paragraphOrOptions
  const id = insertDecorated(document, 'callout', paragraph, value, { ...(value.config ? { callout: value.config } : {}), ...(value.textColor === undefined ? {} : { textColor: value.textColor }) })
  return id ? new FDocumentCallout(document, id, paragraph?.segmentId ?? value.segmentId ?? '') : null
}
export function insertDocumentCode(document: FDocument, paragraphOrOptions: FDocumentParagraph | DocumentCodeInsertOptions = {}, options: DocumentCodeInsertOptions = {}): FDocumentCode | null {
  const paragraph = isParagraph(paragraphOrOptions) ? paragraphOrOptions : undefined, value = isParagraph(paragraphOrOptions) ? options : paragraphOrOptions
  const id = insertDecorated(document, 'code', paragraph, value, { ...(value.config ? { code: value.config } : {}) })
  return id ? new FDocumentCode(document, id, paragraph?.segmentId ?? value.segmentId ?? '') : null
}

export type { DocumentQuoteStyle }
