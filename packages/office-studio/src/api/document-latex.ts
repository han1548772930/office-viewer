import { cloneModel } from '../domain/serialization'
import type { FDocument } from './facade'
import { getLatexFormulas, type DocumentLatexFormulaTarget, type LatexFormulaMode } from '../features/document/latex/latex-formula'
import type { DocumentLatexCommand } from '../features/document/commands/document-latex-commands'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host'
import type { DocumentSelection } from '../features/document/model/document-model'
import { assertDocumentPermission, permissionTargetsForSelection } from '../features/document/permissions/document-permissions'
import { documentSegmentIdForHost } from './document-segments'
import { assertDocumentPublicRange, documentSelectionAtPublicBodyRange, type DocumentPublicBodyIndex } from './document-public-body'
import { isDocumentBodyLatexRange } from './document-body-latex'

export interface DocumentLatexFormulaConfig { latex: string; kind: 'inline'; displayMode: 'inline' }
export interface DocumentLatexCreateOptions {
  properties?: Partial<Omit<DocumentLatexFormulaConfig, 'latex'>>
  /** Local block-formula extension; the fixed public plugin is inline-only. */
  mode?: LatexFormulaMode
}
export interface DocumentLatexInsertOptions extends DocumentLatexCreateOptions { segmentId?: string }
export interface DocumentLatexUpdateOptions { latex: string; properties?: Partial<Omit<DocumentLatexFormulaConfig, 'latex'>> }
export interface DocumentLatexFindQuery { latex?: string; rangeId?: string }
export interface DocumentLatexRange { startOffset: number; endOffset: number; rangeId: string; segmentId: string }
export interface DocumentLatexInfo extends DocumentLatexRange {
  type: 'latex'; latex: string; config: DocumentLatexFormulaConfig; paragraphId: string | null
  context: { before: string; after: string }
}

export function validateDocumentLatexProperties(properties?: DocumentLatexCreateOptions['properties']): void {
  if (properties === undefined) return
  if (!properties || typeof properties !== 'object' || Array.isArray(properties)
    || Object.keys(properties).some(key => key !== 'kind' && key !== 'displayMode')
    || properties.kind !== undefined && properties.kind !== 'inline'
    || properties.displayMode !== undefined && properties.displayMode !== 'inline') throw new TypeError('固定版本的公式公开属性只接受 inline。')
}

export function documentLatexFormulaTargets(document: FDocument, segmentId = ''): DocumentLatexFormulaTarget[] {
  const model = document.getSegmentModel(segmentId), items = getLatexFormulas(model)
  for (const item of items) assertDocumentPermission(model.permissionRules ?? [], { documentId: document.id, userId: document.getCurrentUserId() }, 'view',
    permissionTargetsForSelection(model, { anchor: { ...item.target, offset: item.startOffset }, focus: { ...item.target, offset: item.endOffset } }))
  return items
}

export function documentLatexRangeAt(index: DocumentPublicBodyIndex, offset: number): string | undefined {
  assertDocumentPublicRange(index, offset, offset)
  return index.data.body.customRanges?.find(range => isDocumentBodyLatexRange(range) && offset >= range.startIndex && offset <= range.endIndex)?.rangeId
}

export function assertDocumentLatexAtomicRange(index: DocumentPublicBodyIndex, start: number, end: number): void {
  assertDocumentPublicRange(index, start, end)
  for (const range of index.data.body.customRanges ?? []) if (isDocumentBodyLatexRange(range)
    && (start > range.startIndex && start <= range.endIndex || end > range.startIndex && end <= range.endIndex)) throw new RangeError('公式实体内部的位置不能作为插入或替换边界。')
}

/** Wrappers retain IDs, then resolve the current model for every query/command. */
export class FDocumentLatex {
  constructor(private readonly document: FDocument, private readonly rangeId: string, private readonly segmentId = '') {}
  getId(): string { return this.rangeId }
  getSegmentId(): string { return this.segmentId }
  private source(): DocumentLatexFormulaTarget | undefined { return documentLatexFormulaTargets(this.document, this.segmentId).find(item => item.rangeId === this.rangeId) }
  getRange(): DocumentLatexRange | null {
    if (!this.source()) return null
    const range = this.document.getPublicBodyIndex(this.segmentId).data.body.customRanges?.find(range => range.rangeId === this.rangeId && isDocumentBodyLatexRange(range))
    return range ? { rangeId: this.rangeId, segmentId: this.segmentId, startOffset: range.startIndex, endOffset: range.endIndex + 1 } : null
  }
  getLatex(): string { return this.source()?.formula.latex ?? '' }
  getConfig(): DocumentLatexFormulaConfig { return { latex: this.getLatex(), kind: 'inline', displayMode: 'inline' } }
  getMode(): LatexFormulaMode | null { return this.source()?.formula.mode ?? null }
  describe(): DocumentLatexInfo | null {
    const source = this.source(), range = this.getRange()
    if (!source || !range) return null
    const index = this.document.getPublicBodyIndex(this.segmentId), paragraph = index.positions.find(position => position.block.id === source.blockId)
    return { ...range, type: 'latex', latex: source.formula.latex, config: this.getConfig(), paragraphId: source.blockId,
      context: { before: paragraph ? index.data.body.dataStream.slice(paragraph.start, range.startOffset) : '',
        after: paragraph ? index.data.body.dataStream.slice(range.endOffset, paragraph.end) : '' } }
  }
  private execute(command: DocumentLatexCommand): boolean {
    const source = this.source()
    if (!source) return false
    const host = this.document.getSegmentHost(this.segmentId)
    return this.document.executeCommand({ ...command, ...(host ? { host } : {}) },
      { anchor: { ...source.target, offset: source.startOffset }, focus: { ...source.target, offset: source.endOffset } }).changed
  }
  update(options: DocumentLatexUpdateOptions): boolean {
    if (!options || typeof options.latex !== 'string') throw new TypeError('公式更新需要源文本。')
    validateDocumentLatexProperties(options.properties)
    return this.execute({ id: 'doc.command.update-latex-formula', rangeId: this.rangeId, latex: options.latex })
  }
  remove(): boolean { return this.execute({ id: 'doc.command.remove-latex-formula', rangeId: this.rangeId }) }
  replaceWithText(text: string): boolean { return this.execute({ id: 'doc.command.replace-latex-formula-with-text', rangeId: this.rangeId, text }) }
}

export function findDocumentLatexFormulas(document: FDocument, query: string | DocumentLatexFindQuery, segmentId = ''): FDocumentLatex[] {
  const filter = typeof query === 'string' ? { latex: query } : query
  if (!filter || typeof filter !== 'object' || Array.isArray(filter)
    || [filter.latex, filter.rangeId].some(value => value !== undefined && typeof value !== 'string')) throw new TypeError('公式查询参数无效。')
  return documentLatexFormulaTargets(document, segmentId).filter(item => (filter.rangeId === undefined || item.rangeId === filter.rangeId)
    && (filter.latex === undefined || item.formula.latex.includes(filter.latex))).map(item => new FDocumentLatex(document, item.rangeId, segmentId))
}

export function insertDocumentLatexAtRange(document: FDocument, start: number, end: number, latex: string, options: DocumentLatexInsertOptions = {}): FDocumentLatex | null {
  const segmentId = options.segmentId ?? '', index = document.getPublicBodyIndex(segmentId)
  assertDocumentLatexAtomicRange(index, start, end)
  const host = document.getSegmentHost(segmentId), selection = documentSelectionAtPublicBodyRange(document.getModel(), start, end, host)
  return selection ? insertAtSelection(document, selection, latex, options, host, segmentId) : null
}

export function insertDocumentLatexAtSelection(document: FDocument, latex: string, options: DocumentLatexCreateOptions = {}): FDocumentLatex | null {
  const state = document.getInputState()
  if (!state?.valid || state.location.kind !== 'body') return null
  const host = state.location.host, segmentId = host ? documentSegmentIdForHost(document.getModel(), host) : ''
  return insertAtSelection(document, cloneModel(state.location.selection), latex, options, host, segmentId)
}

function insertAtSelection(document: FDocument, selection: DocumentSelection, latex: string, options: DocumentLatexCreateOptions,
  host: DocumentAuxiliaryBlockHost | undefined, segmentId: string): FDocumentLatex | null {
  validateDocumentLatexProperties(options.properties)
  const before = new Set(documentLatexFormulaTargets(document, segmentId).map(item => item.rangeId))
  const state = document.getInputState(), marks = document.getInputMarks(selection, undefined, host)
  const result = document.executeCommand({ id: 'doc.command.insert-latex-formula', latex, mode: options.mode, marks,
    ...(host ? { host } : {}) }, selection)
  if (!result.changed || result.selectionCurrent === false) return null
  const latest = document.getInputState()
  if (state && latest && !sameDocumentBlockHost(host, latest.location.kind === 'body' ? latest.location.host : undefined)) return null
  const item = documentLatexFormulaTargets(document, segmentId).find(item => !before.has(item.rangeId))
  return item ? new FDocumentLatex(document, item.rangeId, segmentId) : null
}
