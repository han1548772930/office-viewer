import type { DocumentBlock, DocumentModel } from '../../../domain/model.ts'
import { blockAtPoint, collapsedSelection, createParagraph, createRun, findDocumentBlock, findDocumentBlockContainer, orderedSelection,
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
import { visitDocumentBlocks } from '../model/document-structural-blocks'
import { updateDocumentStructuredCodeBlock } from '../services/document-code-editing'
import { documentCodeSourceEdit as codeSourceEdit, normalizeDocumentCodeSource as normalizeCodeSource } from '../model/document-code-source'

export type DocumentCodeLanguage = 'plain' | 'javascript' | 'typescript' | 'json' | 'css' | 'html' | 'python' | 'sql' | 'rust'
export type CodeTokenKind = 'keyword' | 'string' | 'number' | 'comment' | 'operator' | 'punctuation' | 'plain'

export interface DocumentCodeBlockData {
  language: string
  source: string
  lineNumbers: boolean
  tabSize?: number
  wrap?: boolean
}

export interface CodeToken {
  kind: CodeTokenKind
  text: string
  start: number
  end: number
}

const languages = new Set<DocumentCodeLanguage>(['plain', 'javascript', 'typescript', 'json', 'css', 'html', 'python', 'sql', 'rust'])
const jsKeywords = new Set('async await break case catch class const continue debugger default delete do else export extends finally for from function get if import in instanceof let new of return set static super switch this throw try typeof var void while with yield interface type enum namespace implements public private protected readonly'.split(' '))
const pythonKeywords = new Set('and as assert async await break class continue def del elif else except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield True False None'.split(' '))
const sqlKeywords = new Set('add all alter as asc begin by case cast constraint create cross database date desc delete distinct drop else end exists foreign from full group having in index inner insert into is join key left like not null on or order outer primary references right select set table then union unique update values when where'.split(' '))
const rustKeywords = new Set('as async await break const continue crate dyn else enum extern false fn for if impl in let loop match mod move mut pub ref return self static struct super trait true type unsafe use where while'.split(' '))

export function createCodeBlock(source = '', language: DocumentCodeLanguage = 'plain', lineNumbers = false): DocumentBlock {
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
    lineNumbers: block.code.showLineNumbers, tabSize: block.code.tabSize, wrap: block.code.wrap }
  const value = block.metadata.codeBlock
  if (!isRecord(value) || typeof value.source !== 'string') return undefined
  return {
    source: normalizeCodeSource(value.source),
    language: normalizeCodeLanguage(value.language),
    lineNumbers: value.lineNumbers === true,
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
      patch: { ...(patch.language === undefined ? {} : { language: patch.language }), ...(patch.lineNumbers === undefined ? {} : { showLineNumbers: patch.lineNumbers }),
        ...(patch.tabSize === undefined ? {} : { tabSize: patch.tabSize }), ...(patch.wrap === undefined ? {} : { wrap: patch.wrap }) },
    }, permissionContext ?? { documentId: model.id, userId: 'local-user' })
    return { changed: result.changed, ...(sourceSelection ? { selection: result.selection } : {}) }
  }
  const next: DocumentCodeBlockData = {
    source: patch.source === undefined ? current.source : normalizeCodeSource(patch.source),
    language: patch.language === undefined ? current.language : normalizeCodeLanguage(patch.language),
    lineNumbers: patch.lineNumbers === undefined ? current.lineNumbers : patch.lineNumbers === true,
  }
  if (host || permissionContext || current.source !== next.source && model.contentControls?.length) {
    const candidate = cloneModel(model), scoped = host ? documentModelForBlockHost(candidate, host, true) : candidate
    const change = codeSourceEdit(current.source, next.source)
    const target = textTargets(source).find(target => (target.cellBlockId ?? target.blockId) === block.id)
    if (!target) throw new RangeError('代码块文字目标不存在。')
    const paragraph = documentTextFieldParagraphs(source).find(paragraph => paragraph.block.id === block.id)
    if (!paragraph) throw new RangeError('代码块源段落不存在。')
    const selection = { anchor: { ...target, offset: change.startOffset }, focus: { ...target, offset: change.endOffset } }
    const textChanged = current.source !== next.source
    const activeControlId = textChanged ? resolveDocumentContentControlTextInputId(source, selection) : undefined
    const changed = applyCodeBlockUpdate(scoped, findDocumentBlock(scoped, block.id)!, current, next, change, {}, { activeControlId })
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

function applyCodeBlockUpdate(model: DocumentModel, block: DocumentBlock, current: DocumentCodeBlockData, next: DocumentCodeBlockData, edit?: CodeSourceEdit, marks: RunPatch = {}, controlEdit: ContentControlRangeEditOptions = {}): boolean {
  if (current.source === next.source && current.language === next.language && current.lineNumbers === next.lineNumbers) return false
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
  block.metadata = { ...block.metadata, codeBlock: next }
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

export function insertCodeBlock(model: DocumentModel, afterBlockId: string | undefined, source = '', language: DocumentCodeLanguage = 'plain'): DocumentBlock {
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

/** Token ranges are UTF-16 offsets into the authoritative source string. */
export function tokenizeCode(source: string, languageValue: string): CodeToken[] {
  const language = normalizeCodeLanguage(languageValue)
  const text = normalizeCodeSource(source)
  if (language === 'plain') return text ? [{ kind: 'plain', text, start: 0, end: text.length }] : []
  const tokens: CodeToken[] = []
  const pattern = /("(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|`(?:\\.|[^`\\])*`|\/\/[^\n]*|#[^\n]*|\/\*[\s\S]*?\*\/|<!--[\s\S]*?-->|\b\d+(?:\.\d+)?\b|[A-Za-z_$][\w$]*|=>|==?=?|!=?=?|<=?|>=?|\+\+|--|&&|\|\||[{}()[\].,;:]|[-+*/%=&|!<>?])/g
  let cursor = 0
  for (const match of text.matchAll(pattern)) {
    const start = match.index ?? 0
    if (start > cursor) tokens.push({ kind: 'plain', text: text.slice(cursor, start), start: cursor, end: start })
    const value = match[0]
    const kind = tokenKind(value, language)
    tokens.push({ kind, text: value, start, end: start + value.length })
    cursor = start + value.length
  }
  if (cursor < text.length) tokens.push({ kind: 'plain', text: text.slice(cursor), start: cursor, end: text.length })
  return tokens
}

export function cloneCodeBlockData(value: DocumentCodeBlockData): DocumentCodeBlockData {
  return { ...value }
}

export function normalizeCodeLanguage(value: unknown): DocumentCodeLanguage {
  if (typeof value !== 'string') return 'plain'
  const normalized = value.trim().toLowerCase() as DocumentCodeLanguage
  return languages.has(normalized) ? normalized : 'plain'
}

function graphemeBoundaries(source: string): number[] {
  const offsets = Array.from(new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(source), (part) => part.index)
  offsets.push(source.length)
  return offsets
}

function tokenKind(value: string, language: DocumentCodeLanguage): CodeTokenKind {
  if (/^(?:"|'|`)/.test(value) || (language === 'json' && value.startsWith('"'))) return 'string'
  if (/^(?:\/\/|#|\/\*|<!--)/.test(value)) return 'comment'
  if (/^\d/.test(value)) return 'number'
  const keywords = language === 'python' ? pythonKeywords : language === 'sql' ? sqlKeywords : language === 'rust' ? rustKeywords : jsKeywords
  if (keywords.has(value)) return 'keyword'
  if (/^[{}()[\].,;:]$/.test(value)) return 'punctuation'
  if (/^(?:=>|==?=?|!=?=?|<=?|>=?|\+\+|--|&&|\|\||[-+*/%=&|!<>?])$/.test(value)) return 'operator'
  return 'plain'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

