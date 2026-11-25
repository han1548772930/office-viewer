import type { DocumentBlock, DocumentModel } from '../../../domain/model.ts'
import { blockAtPoint, clampPoint, collapsedSelection, createParagraph, createRun, findDocumentBlock, findDocumentBlockContainer, orderedSelection,
  runText, sameTarget, selectionCollapsed, textTargets, type DocumentSelection, type RunPatch } from '../model/document-model.ts'
import { replaceDocumentRunRange } from '../model/document-run-range'
import { documentDefaultStyle } from '../model/document-default-style'
import { isDocumentTextBoundary } from '../model/document-story'
import { transformContentControlsForBlockEdit, type ContentControlRangeEditOptions } from '../services/content-controls.ts'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments.ts'
import { documentTextFieldParagraphs, transformDocumentTextFieldsForBlockEdit } from '../model/document-text-fields'
import { cloneModel } from '../../../domain/serialization'
import { resolveDocumentContentControlTextInputId, synchronizeDocumentContentControlsAfterEdit } from '../services/content-control-editing'
import { mapDocumentContentControlBindingSelection, synchronizeDocumentContentControlBindings } from '../services/content-control-bindings'
import { documentBlockHostForId, documentEditingBlockHost } from '../model/document-block-host'
import { applyDocumentBlockHostModel, documentModelForBlockHost } from '../services/document-block-host'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, type DocumentPermissionContext } from '../permissions/document-permissions'
import { validateDocumentCodeConfig, visitDocumentBlocks } from '../model/document-structural-blocks'
import { updateDocumentStructuredCodeBlock } from '../services/document-code-editing'
import { documentCodeSourceEdit as codeSourceEdit, documentCodeIndentEdits, normalizeDocumentCodeSource as normalizeCodeSource } from '../model/document-code-source'
import { documentCodeFontFamily, documentCodeUsesFontFamily, setDocumentCodeFontFamily, validateDocumentCodeFontFamily } from './document-code-fonts'

export { tokenizeCode } from './document-code-tokenizer'
export type { DocumentCodeLanguage, CodeTokenKind, CodeToken } from './document-code-tokenizer'

export interface DocumentCodeBlockData {
  language: string
  source: string
  lineNumbers: boolean
  tabSize?: number
  wrap?: boolean
  fontFamily?: string
}

export function createCodeBlock(source = '', language = 'plain', lineNumbers = false): DocumentBlock {
  const normalizedLanguage = normalizeCodeLanguage(language)
  const normalizedSource = normalizeCodeSource(source)
  const block = createParagraph(normalizedSource, { fontFamily: 'Consolas', fontSize: 10 })
  block.metadata.codeBlock = { language: normalizedLanguage, source: normalizedSource, lineNumbers } satisfies DocumentCodeBlockData
  block.style = { ...block.style, align: 'left', indent: 0, lineHeight: 1.35, spaceBefore: 8, spaceAfter: 8 }
  return block
}

export function codeBlockData(block: DocumentBlock): DocumentCodeBlockData | undefined {
  if (block.type === 'code' && block.code && block.children) return {
    source: documentTextFieldParagraphs({ blocks: block.children }).map(paragraph => paragraph.text).join('\n'), language: block.code.language,
    lineNumbers: block.code.showLineNumbers, tabSize: block.code.tabSize, wrap: block.code.wrap, fontFamily: documentCodeFontFamily(block) }
  const value = block.metadata.codeBlock
  if (!isRecord(value) || typeof value.source !== 'string') return undefined
  return {
    source: normalizeCodeSource(value.source),
    language: normalizeCodeLanguage(value.language),
    lineNumbers: value.lineNumbers === true,
    fontFamily: documentCodeFontFamily(block),
    ...(typeof value.tabSize === 'number' ? { tabSize: value.tabSize } : {}), ...(typeof value.wrap === 'boolean' ? { wrap: value.wrap } : {}),
  }
}

export function codeBlocks(model: DocumentModel): DocumentBlock[] {
  const blocks: DocumentBlock[] = []
  visitDocumentBlocks(model.blocks, block => { if (codeBlockData(block)) blocks.push(block) })
  return blocks
}

export function updateCodeBlock(model: DocumentModel, blockId: string, patch: Partial<DocumentCodeBlockData>): boolean {
  return updateCodeBlockWithSelection(model, blockId, patch).changed
}

/** Keyboard indentation shares source editing, control locks, binding updates and one outer history item. */
export function indentCodeBlockWithSelection(model: DocumentModel, blockId: string, selection: DocumentSelection, direction: -1 | 1,
  context: DocumentPermissionContext): DocumentSelection {
  if (selection.tableSelection) throw new RangeError('代码缩进需要文字选区。')
  const host = documentBlockHostForId(model, blockId), scoped = host ? documentModelForBlockHost(model, host, true) : model
  const block = findDocumentBlock(scoped, blockId), data = block && codeBlockData(block)
  if (!block || !data) throw new RangeError('代码块已不存在。')
  const { start, end } = orderedSelection(scoped, selection)
  const paragraphs = block.type === 'code' ? documentTextFieldParagraphs({ blocks: block.children! }) : documentTextFieldParagraphs({ blocks: [block] })
  const sourceOffset = (point: DocumentSelection['focus']) => {
    const paragraph = paragraphs.find(value => value.block.id === (point.cellBlockId ?? point.blockId))
    if (!paragraph) throw new RangeError('代码缩进不能跨出当前代码块。')
    return paragraph.absoluteStart + point.offset
  }
  const edits = documentCodeIndentEdits(data.source, sourceOffset(start), sourceOffset(end), direction, data.tabSize ?? 4)
  if (!edits.length) return cloneModel(selection)
  const candidate = cloneModel(model)
  let nextSelection = cloneModel(selection), source = data.source
  for (const edit of edits) {
    source = source.slice(0, edit.startOffset) + edit.text + source.slice(edit.endOffset)
    const updated = updateCodeBlockWithSelection(candidate, blockId, { source }, nextSelection, context)
    nextSelection = updated.selection ?? nextSelection
    const currentHost = host ? documentModelForBlockHost(candidate, host, true) : candidate
    const currentBlock = findDocumentBlock(currentHost, blockId)
    if (!currentBlock || codeBlockData(currentBlock)?.source !== source) throw new RangeError('内容绑定改变了代码源文本，请重新选择后缩进。')
  }
  const currentHost = host ? documentModelForBlockHost(candidate, host, true) : candidate, currentBlock = findDocumentBlock(currentHost, blockId)!
  const currentParagraphs = documentTextFieldParagraphs({ blocks: currentBlock.type === 'code' ? currentBlock.children! : [currentBlock] })
  const mapPoint = (point: DocumentSelection['focus']) => {
    let offset = sourceOffset(point)
    for (const edit of edits) {
      if (offset >= edit.endOffset) offset += edit.text.length - (edit.endOffset - edit.startOffset)
      else if (offset > edit.startOffset) offset = edit.startOffset + edit.text.length
    }
    const paragraph = currentParagraphs.find(value => value.block.id === (point.cellBlockId ?? point.blockId))
    if (!paragraph) throw new RangeError('代码缩进后的输入段落已不存在。')
    return clampPoint(currentHost, { ...point, offset: Math.max(0, offset - paragraph.absoluteStart) })
  }
  nextSelection = { anchor: mapPoint(selection.anchor), focus: mapPoint(selection.focus) }
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, candidate))
  Object.assign(model, candidate)
  return nextSelection
}

/** Source edits and linked controls share the full document candidate, including auxiliary hosts. */
export function updateCodeBlockWithSelection(model: DocumentModel, blockId: string, patch: Partial<DocumentCodeBlockData>,
  sourceSelection?: DocumentSelection, permissionContext?: DocumentPermissionContext): { changed: boolean; selection?: DocumentSelection } {
  const host = documentBlockHostForId(model, blockId), source = host ? documentModelForBlockHost(model, host, true) : model
  const block = findDocumentBlock(source, blockId)
  if (!block) return { changed: false }
  const current = codeBlockData(block)
  if (!current) return { changed: false }
  if (block.type === 'code') {
    const owned = new Set(documentTextFieldParagraphs({ blocks: block.children! }).map(paragraph => paragraph.block.id))
    const target = textTargets(source).find(target => owned.has(target.cellBlockId ?? target.blockId))
    if (!target) throw new RangeError('代码块缺少真实输入段落。')
    const result = updateDocumentStructuredCodeBlock(model, sourceSelection ?? collapsedSelection({ ...target, offset: 0 }), {
      blockId, ...(patch.source === undefined ? {} : { source: patch.source }),
      ...(patch.fontFamily === undefined ? {} : { fontFamily: patch.fontFamily }),
      patch: { ...(patch.language === undefined ? {} : { language: patch.language }), ...(patch.lineNumbers === undefined ? {} : { showLineNumbers: patch.lineNumbers }),
        ...(patch.tabSize === undefined ? {} : { tabSize: patch.tabSize }), ...(patch.wrap === undefined ? {} : { wrap: patch.wrap }) },
    }, permissionContext ?? { documentId: model.id, userId: 'local-user' })
    return { changed: result.changed, ...(sourceSelection ? { selection: result.selection } : {}) }
  }
  const next: DocumentCodeBlockData = {
    ...current,
    source: patch.source === undefined ? current.source : normalizeCodeSource(patch.source),
    language: patch.language === undefined ? current.language : normalizeCodeLanguage(patch.language),
    lineNumbers: patch.lineNumbers === undefined ? current.lineNumbers : patch.lineNumbers === true,
    ...(patch.tabSize === undefined ? {} : { tabSize: patch.tabSize }), ...(patch.wrap === undefined ? {} : { wrap: patch.wrap }),
    ...(patch.fontFamily === undefined ? {} : { fontFamily: patch.fontFamily }),
  }
  if (patch.fontFamily !== undefined) validateDocumentCodeFontFamily(patch.fontFamily)
  const fontChanged = patch.fontFamily !== undefined && !documentCodeUsesFontFamily(block, patch.fontFamily)
  validateDocumentCodeConfig({ language: next.language, showLineNumbers: next.lineNumbers, tabSize: next.tabSize ?? 4, wrap: next.wrap ?? true })
  if (host || permissionContext || patch.fontFamily !== undefined || current.source !== next.source && model.contentControls?.length) {
    const candidate = cloneModel(model), scoped = host ? documentModelForBlockHost(candidate, host, true) : candidate
    const change = codeSourceEdit(current.source, next.source)
    const target = textTargets(source).find(target => (target.cellBlockId ?? target.blockId) === block.id)
    if (!target) throw new RangeError('代码块文字目标不存在。')
    const paragraph = documentTextFieldParagraphs(source).find(paragraph => paragraph.block.id === block.id)
    if (!paragraph) throw new RangeError('代码块源段落不存在。')
    const selection = { anchor: { ...target, offset: change.startOffset }, focus: { ...target, offset: change.endOffset } }
    const textChanged = current.source !== next.source
    const activeControlId = textChanged ? resolveDocumentContentControlTextInputId(source, selection) : undefined
    const changed = applyCodeBlockUpdate(scoped, findDocumentBlock(scoped, block.id)!, current, next, change, {}, { activeControlId }, fontChanged)
    if (!changed) return { changed: false, ...(sourceSelection ? { selection: cloneModel(sourceSelection) } : {}) }
    if (host) applyDocumentBlockHostModel(candidate, host, scoped)
    const selectionHost = host ?? documentEditingBlockHost(model)
    const mappedSelection = sourceSelection && textChanged ? mapDocumentContentControlBindingSelection(sourceSelection,
      { controlId: activeControlId ?? '', blockId: block.id, startOffset: change.startOffset, endOffset: change.endOffset, text: change.text,
        ...(selectionHost ? { host: selectionHost } : {}) }, selectionHost) : sourceSelection
    const editResult = textChanged ? synchronizeDocumentContentControlsAfterEdit(model, candidate, { origin: 'text-input', selection, activeControlId }) : undefined
    const bindings = textChanged ? synchronizeDocumentContentControlBindings(model, candidate, { sourceSelection: selection,
      selection: mappedSelection ?? collapsedSelection({ ...target, offset: change.startOffset + change.text.length }), host: selectionHost,
      originTextEdits: [{ startOffset: paragraph.absoluteStart + change.startOffset, endOffset: paragraph.absoluteStart + change.endOffset, text: change.text }],
      activeControlId, editResult, permissionContext }) : undefined
    if (permissionContext) assertDocumentPermission(model.permissionRules ?? [], permissionContext, 'edit', permissionTargetsForChangedBlocks(model, candidate))
    Object.assign(model, candidate)
    return { changed, ...(sourceSelection ? { selection: bindings?.selection ?? mappedSelection } : {}) }
  }
  const changed = applyCodeBlockUpdate(model, block, current, next)
  const change = current.source !== next.source ? codeSourceEdit(current.source, next.source) : undefined
  const selection = sourceSelection && change ? mapDocumentContentControlBindingSelection(sourceSelection,
    { controlId: '', blockId: block.id, ...change, ...(documentEditingBlockHost(model) ? { host: documentEditingBlockHost(model) } : {}) }, documentEditingBlockHost(model)) : sourceSelection
  return { changed, ...(selection ? { selection: cloneModel(selection) } : {}) }
}

interface CodeSourceEdit { startOffset: number; endOffset: number; text: string }

function applyCodeBlockUpdate(model: DocumentModel, block: DocumentBlock, current: DocumentCodeBlockData, next: DocumentCodeBlockData, edit?: CodeSourceEdit, marks: RunPatch = {}, controlEdit: ContentControlRangeEditOptions = {}, applyFont = false): boolean {
  if (current.source === next.source && current.language === next.language && current.lineNumbers === next.lineNumbers
    && current.tabSize === next.tabSize && current.wrap === next.wrap && current.fontFamily === next.fontFamily && !applyFont) return false
  if (current.source !== next.source) {
    const change = edit ?? codeSourceEdit(current.source, next.source)
    if (runText(block.runs) !== current.source) throw new Error('代码源文本与格式范围不一致。')
    if (!isDocumentTextBoundary(current.source, change.startOffset) || !isDocumentTextBoundary(current.source, change.endOffset)
      || change.endOffset < change.startOffset || current.source.slice(0, change.startOffset) + change.text + current.source.slice(change.endOffset) !== next.source) throw new RangeError('代码编辑范围无效。')
    const nextRuns = replaceDocumentRunRange(block.runs, change.startOffset, change.endOffset, change.text, documentDefaultStyle(model).text, marks)
    transformContentControlsForBlockEdit(model, block.id, change.startOffset, change.endOffset, change.text.length, undefined, false, controlEdit)
    transformDocumentTextFieldsForBlockEdit(block, change.startOffset, change.endOffset, change.text, undefined, model)
    block.runs = nextRuns
    transformDocumentCommentAnchorsForTextEdit(model, block.id, change.startOffset, change.endOffset, change.text)
  }
  if (applyFont && next.fontFamily !== undefined) setDocumentCodeFontFamily(model, block, next.fontFamily)
  block.metadata = { ...block.metadata, codeBlock: { ...block.metadata.codeBlock as Record<string, unknown>, ...next } }
  return true
}

export function editCodeBlockText(model: DocumentModel, selection: DocumentSelection, edit: { text: string; marks?: RunPatch } | { deleteDirection: -1 | 1 }, controlEdit: ContentControlRangeEditOptions = {}): DocumentSelection | undefined {
  if (selection.tableSelection) return undefined
  const { start, end } = orderedSelection(model, selection)
  if (!sameTarget(start, end)) return undefined
  const block = blockAtPoint(model, start)
  const data = block && codeBlockData(block)
  if (!block || !data || !isDocumentTextBoundary(data.source, start.offset) || !isDocumentTextBoundary(data.source, end.offset)) return undefined
  let startOffset = start.offset
  let endOffset = end.offset
  let insertion = ''
  if ('text' in edit) insertion = edit.text.replace(/\r\n?/g, '\n')
  else if (selectionCollapsed(selection)) {
    const boundaries = graphemeBoundaries(data.source)
    const adjacent = edit.deleteDirection < 0
      ? boundaries.filter((offset) => offset < startOffset).at(-1)
      : boundaries.find((offset) => offset > startOffset)
    if (adjacent === undefined) return selection
    if (edit.deleteDirection < 0) startOffset = adjacent
    else endOffset = adjacent
  }
  const source = `${data.source.slice(0, startOffset)}${insertion}${data.source.slice(endOffset)}`
  const active = controlEdit.activeControlId && model.contentControls?.find(control => control.id === controlEdit.activeControlId)
  const inside = active && active.blockId === block.id && startOffset >= active.startOffset && endOffset <= active.endOffset
  applyCodeBlockUpdate(model, block, data, { ...data, source }, { startOffset, endOffset, text: insertion }, 'text' in edit ? edit.marks : undefined,
    inside ? controlEdit : {})
  const caret = startOffset + insertion.length
  return collapsedSelection({ ...start, offset: caret })
}

export function insertCodeBlock(model: DocumentModel, afterBlockId: string | undefined, source = '', language = 'plain'): DocumentBlock {
  const block = createCodeBlock(source, language)
  const container = afterBlockId ? findDocumentBlockContainer(model, afterBlockId) : model.blocks
  if (!container) throw new RangeError('代码块插入目标已不存在')
  const index = afterBlockId ? container.findIndex((candidate) => candidate.id === afterBlockId) : container.length - 1
  if (afterBlockId && index < 0) throw new RangeError('代码块插入目标已不存在')
  container.splice(index + 1, 0, block)
  return block
}

export function exitCodeBlock(model: DocumentModel, blockId: string): DocumentBlock | undefined {
  const container = findDocumentBlockContainer(model, blockId)
  const index = container?.findIndex((candidate) => candidate.id === blockId) ?? -1
  const block = container?.[index]
  if (!block || !codeBlockData(block)) return undefined
  const paragraph = createParagraph()
  container!.splice(index + 1, 0, paragraph)
  return paragraph
}

export function cancelEmptyCodeParagraph(model: DocumentModel, blockId: string): boolean {
  const container = findDocumentBlockContainer(model, blockId)
  const index = container?.findIndex((candidate) => candidate.id === blockId) ?? -1
  const block = container?.[index]
  const data = block && codeBlockData(block)
  if (!block || !data || data.source.length !== 0 || runText(block.runs).length !== 0) return false
  container!.splice(index, 1)
  if (!container!.length) container!.push(createParagraph())
  return true
}

export function unwrapCodeBlock(block: DocumentBlock): boolean {
  if (block.type === 'code') return false
  const data = codeBlockData(block)
  if (!data) return false
  delete block.metadata.codeBlock
  block.runs = [createRun(data.source)]
  block.style = { ...block.style, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 }
  return true
}

export function codeBlockPlainText(block: DocumentBlock): string {
  return codeBlockData(block)?.source ?? runText(block.runs)
}

export function cloneCodeBlockData(value: DocumentCodeBlockData): DocumentCodeBlockData {
  return { ...value }
}

export function normalizeCodeLanguage(value: unknown): string {
  return typeof value === 'string' && value.trim() ? value : 'plain'
}

function graphemeBoundaries(source: string): number[] {
  const offsets = Array.from(new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(source), (part) => part.index)
  offsets.push(source.length)
  return offsets
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

