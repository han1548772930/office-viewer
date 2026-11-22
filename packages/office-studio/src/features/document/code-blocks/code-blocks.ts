import type { DocumentBlock, DocumentModel } from '../../../domain/model.ts'
import { collapsedSelection, createParagraph, createRun, orderedSelection, runText, selectionCollapsed, type DocumentSelection, type RunPatch } from '../model/document-model.ts'
import { replaceDocumentRunRange } from '../model/document-run-range'
import { documentDefaultStyle } from '../model/document-default-style'
import { isDocumentTextBoundary } from '../model/document-story'
import { transformContentControlsForBlockEdit } from '../services/content-controls.ts'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments.ts'

export type DocumentCodeLanguage = 'plain' | 'javascript' | 'typescript' | 'json' | 'css' | 'html' | 'python' | 'sql' | 'rust'
export type CodeTokenKind = 'keyword' | 'string' | 'number' | 'comment' | 'operator' | 'punctuation' | 'plain'

export interface DocumentCodeBlockData {
  language: DocumentCodeLanguage
  source: string
  lineNumbers: boolean
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
  const value = block.metadata.codeBlock
  if (!isRecord(value) || typeof value.source !== 'string') return undefined
  return {
    source: normalizeCodeSource(value.source),
    language: normalizeCodeLanguage(value.language),
    lineNumbers: value.lineNumbers === true,
  }
}

export function codeBlocks(model: DocumentModel): DocumentBlock[] {
  return model.blocks.filter((block) => codeBlockData(block) !== undefined)
}

export function updateCodeBlock(model: DocumentModel, blockId: string, patch: Partial<DocumentCodeBlockData>): boolean {
  const block = model.blocks.find((candidate) => candidate.id === blockId)
  if (!block) return false
  const current = codeBlockData(block)
  if (!current) return false
  const next: DocumentCodeBlockData = {
    source: patch.source === undefined ? current.source : normalizeCodeSource(patch.source),
    language: patch.language === undefined ? current.language : normalizeCodeLanguage(patch.language),
    lineNumbers: patch.lineNumbers === undefined ? current.lineNumbers : patch.lineNumbers === true,
  }
  return applyCodeBlockUpdate(model, block, current, next)
}

interface CodeSourceEdit { startOffset: number; endOffset: number; text: string }

function codeSourceEdit(before: string, after: string): CodeSourceEdit {
  let startOffset = 0
  while (startOffset < before.length && startOffset < after.length && before[startOffset] === after[startOffset]) startOffset++
  while (startOffset > 0 && (!isDocumentTextBoundary(before, startOffset) || !isDocumentTextBoundary(after, startOffset))) startOffset--
  let suffix = 0
  while (suffix < before.length - startOffset && suffix < after.length - startOffset
    && before[before.length - suffix - 1] === after[after.length - suffix - 1]) suffix++
  while (suffix > 0 && (!isDocumentTextBoundary(before, before.length - suffix) || !isDocumentTextBoundary(after, after.length - suffix))) suffix--
  return { startOffset, endOffset: before.length - suffix, text: after.slice(startOffset, after.length - suffix) }
}

function applyCodeBlockUpdate(model: DocumentModel, block: DocumentBlock, current: DocumentCodeBlockData, next: DocumentCodeBlockData, edit?: CodeSourceEdit, marks: RunPatch = {}): boolean {
  if (current.source === next.source && current.language === next.language && current.lineNumbers === next.lineNumbers) return false
  if (current.source !== next.source) {
    const change = edit ?? codeSourceEdit(current.source, next.source)
    if (runText(block.runs) !== current.source) throw new Error('代码源文本与格式范围不一致。')
    if (!isDocumentTextBoundary(current.source, change.startOffset) || !isDocumentTextBoundary(current.source, change.endOffset)
      || change.endOffset < change.startOffset || current.source.slice(0, change.startOffset) + change.text + current.source.slice(change.endOffset) !== next.source) throw new RangeError('代码编辑范围无效。')
    const nextRuns = replaceDocumentRunRange(block.runs, change.startOffset, change.endOffset, change.text, documentDefaultStyle(model).text, marks)
    transformContentControlsForBlockEdit(model, block.id, change.startOffset, change.endOffset, change.text.length)
    block.runs = nextRuns
    transformDocumentCommentAnchorsForTextEdit(model, block.id, change.startOffset, change.endOffset, change.text)
  }
  block.metadata = { ...block.metadata, codeBlock: next }
  return true
}

export function editCodeBlockText(model: DocumentModel, selection: DocumentSelection, edit: { text: string; marks?: RunPatch } | { deleteDirection: -1 | 1 }): DocumentSelection | undefined {
  const { start, end } = orderedSelection(model, selection)
  if (start.blockId !== end.blockId || start.row != null || end.row != null) return undefined
  const block = model.blocks.find((candidate) => candidate.id === start.blockId)
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
  applyCodeBlockUpdate(model, block, data, { ...data, source }, { startOffset, endOffset, text: insertion }, 'text' in edit ? edit.marks : undefined)
  const caret = startOffset + insertion.length
  return collapsedSelection({ blockId: block.id, offset: caret })
}

export function insertCodeBlock(model: DocumentModel, afterBlockId: string | undefined, source = '', language: DocumentCodeLanguage = 'plain'): DocumentBlock {
  const block = createCodeBlock(source, language)
  const index = afterBlockId ? model.blocks.findIndex((candidate) => candidate.id === afterBlockId) : model.blocks.length - 1
  model.blocks.splice(index < 0 ? model.blocks.length : index + 1, 0, block)
  return block
}

export function exitCodeBlock(model: DocumentModel, blockId: string): DocumentBlock | undefined {
  const index = model.blocks.findIndex((candidate) => candidate.id === blockId)
  const block = model.blocks[index]
  if (!block || !codeBlockData(block)) return undefined
  const paragraph = createParagraph()
  model.blocks.splice(index + 1, 0, paragraph)
  return paragraph
}

export function cancelEmptyCodeParagraph(model: DocumentModel, blockId: string): boolean {
  const index = model.blocks.findIndex((candidate) => candidate.id === blockId)
  const block = model.blocks[index]
  const data = block && codeBlockData(block)
  if (!block || !data || data.source.length !== 0 || runText(block.runs).length !== 0) return false
  model.blocks.splice(index, 1)
  if (!model.blocks.length) model.blocks.push(createParagraph())
  return true
}

export function unwrapCodeBlock(block: DocumentBlock): boolean {
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
export function tokenizeCode(source: string, language: DocumentCodeLanguage): CodeToken[] {
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
  return { source: value.source, language: value.language, lineNumbers: value.lineNumbers }
}

export function normalizeCodeLanguage(value: unknown): DocumentCodeLanguage {
  if (typeof value !== 'string') return 'plain'
  const normalized = value.trim().toLowerCase() as DocumentCodeLanguage
  return languages.has(normalized) ? normalized : 'plain'
}

function normalizeCodeSource(source: string): string {
  return source.replace(/\r\n?/g, '\n')
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

