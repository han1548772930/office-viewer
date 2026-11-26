import type { DocumentModel, DocumentParagraphStyle, DocumentSectionHeaderFooter } from '../../../domain/model'
import { hasEastAsianLetters } from '../model/document-font'
import { applyRunPatch, createRun, DOCUMENT_NAMED_PARAGRAPH_STYLES, type DocumentNamedParagraphStyle, type DocumentRun, type RunPatch } from '../model/document-model'
import { formatDocumentRunRange, validateDocumentRunPatch } from '../model/document-run-range'
import { DOCUMENT_STORY_KEYS, documentStoryContent, documentStoryDefaultStyle, documentStoryFieldCache, documentStoryFieldDisplayRuns, documentStoryTextStyle, isDocumentTextBoundary, type DocumentStoryRange } from '../model/document-story'
import { assertDocumentPermission, type DocumentPermissionContext } from '../permissions/document-permissions'
import { assertDocumentStoryFormattingEditable } from './document-story-locks'
import { applySectionHeaderFooter, documentSectionIds, resolveSectionHeaderFooter, sectionStoryLinkedToPrevious } from './section-header-footer'
import { storyBuffer, storyRunStyleAt } from './story-editor'

function readRange(model: DocumentModel, request: DocumentStoryRange) {
  if (!request || !DOCUMENT_STORY_KEYS.includes(request.story) || !documentSectionIds(model).includes(request.sectionId)) throw new RangeError('页眉页脚片段不存在。')
  const section = resolveSectionHeaderFooter(model, request.sectionId), buffer = storyBuffer(section, request.story)
  if (!isDocumentTextBoundary(buffer.value, request.start) || !isDocumentTextBoundary(buffer.value, request.end)) throw new RangeError('页眉页脚选区的 UTF-16 位置无效。')
  const range = { ...request, start: Math.min(request.start, request.end), end: Math.max(request.start, request.end) }
  return { section, buffer, range, content: documentStoryContent(section, request.story, model) }
}

function readEditableRange(model: DocumentModel, request: DocumentStoryRange, context: DocumentPermissionContext) {
  const prepared = readRange(model, request)
  const { section, range, content, buffer } = prepared
  if (sectionStoryLinkedToPrevious(section, range.story)) throw new Error('请先取消此页眉页脚与前节的链接，再修改格式。')
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', [
    { scope: 'section', id: range.sectionId }, { scope: 'entity', id: `${range.sectionId}:${range.story}` },
    { scope: 'entity', id: `${range.sectionId}:${range.story.toLowerCase().includes('header') ? 'header' : 'footer'}` },
  ])
  assertDocumentStoryFormattingEditable(content, buffer.value, range.start, range.end)
  return prepared
}

export function documentStoryTextFormattingState(model: DocumentModel, request: DocumentStoryRange, context: DocumentPermissionContext): { available: boolean; reason?: string } {
  try { readEditableRange(model, request, context); return { available: true } }
  catch (error) { return { available: false, reason: error instanceof Error ? error.message : String(error) } }
}

function displayFont(style: RunPatch, text: string): RunPatch {
  return style.eastAsiaFontFamily && hasEastAsianLetters(text) ? { ...style, fontFamily: style.eastAsiaFontFamily } : style
}

function sameTextStyle(left: RunPatch, right: RunPatch): boolean {
  return (['bold', 'italic', 'underline', 'strike', 'color', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'highlight', 'verticalAlign'] as const)
    .every(key => left[key] === right[key])
}

function characterBefore(text: string, offset: number): string {
  const previous = text.charCodeAt(offset - 1)
  return text.slice(Math.max(0, offset - (previous >= 0xDC00 && previous <= 0xDFFF ? 2 : 1)), offset)
}

function fieldCaretCharacter(text: string, offset: number): string {
  if (offset !== 0) return characterBefore(text, text.length)
  const first = text.codePointAt(0)
  return first === undefined ? '' : text.slice(0, first > 0xFFFF ? 2 : 1)
}

/** Read the actual story runs and field caches, never the dormant body selection. */
export function documentStorySelectionFormatting(model: DocumentModel, request: DocumentStoryRange, pending: RunPatch = {}): {
  styles: RunPatch[]; paragraphs: DocumentParagraphStyle[]; hasSelection: boolean; namedStyle: DocumentNamedParagraphStyle
} {
  const { section, buffer, range, content } = readRange(model, request)
  const { start, end } = range, collapsed = start === end
  const styles: RunPatch[] = []
  if (collapsed) {
    const style = applyRunPatch(createRun('', storyRunStyleAt(section, range.story, start, end, model)), pending)
    const preceding = Math.max(0, start - 1), field = buffer.objects.get(preceding)
    const text = field ? fieldCaretCharacter(field.cachedResult ?? '', start) : characterBefore(buffer.value, start)
    styles.push(displayFont(style, text))
  } else {
    const from = buffer.literalOffsets[start]!, to = buffer.literalOffsets[end]!
    let offset = 0
    for (const run of content.runs) {
      const a = Math.max(from, offset), b = Math.min(to, offset + run.text.length)
      if (a < b) styles.push(displayFont(run, run.text.slice(a - offset, b - offset)))
      offset += run.text.length
    }
    const defaults = documentStoryDefaultStyle(model)
    for (const [position, field] of buffer.objects) {
      if (position < start || position >= end) continue
      const display = documentStoryFieldDisplayRuns(field, field.cachedResult ?? '', defaults.text)
      const characters = display.filter(run => run.text.length > 0)
      for (const run of characters) styles.push(displayFont(run, run.text))
      if (!characters.length) styles.push(field.textStyle ?? defaults.text)
    }
  }
  if (!styles.length) styles.push(storyRunStyleAt(section, range.story, start, end, model))
  const from = buffer.literalOffsets[start]!, to = buffer.literalOffsets[end]!
  const paragraphAtStart = content.paragraphs.find(paragraph => paragraph.startOffset <= from && paragraph.endOffset >= from)
  const paragraphs = collapsed || from === to
    ? [paragraphAtStart?.style ?? documentStoryDefaultStyle(model).paragraph]
    : content.paragraphs.filter(paragraph => paragraph.startOffset < to && paragraph.endOffset >= from).map(paragraph => paragraph.style)
  const namedStyle = paragraphAtStart?.metadata?.namedStyleType
  return { styles, paragraphs, hasSelection: !collapsed,
    namedStyle: typeof namedStyle === 'string' && Object.hasOwn(DOCUMENT_NAMED_PARAGRAPH_STYLES, namedStyle) ? namedStyle as DocumentNamedParagraphStyle : 'normal' }
}

/** CJK cache text participates in family selection even though input stores one object marker. */
export function documentStoryFontPatch(model: DocumentModel, request: DocumentStoryRange, patch: RunPatch): RunPatch {
  validateDocumentRunPatch(patch)
  if (!patch.fontFamily || patch.eastAsiaFontFamily !== undefined) return patch
  const { section, buffer, range } = readRange(model, request)
  let text = ''
  if (range.start === range.end) {
    const field = buffer.objects.get(Math.max(0, range.start - 1))
    text = field ? fieldCaretCharacter(field.cachedResult ?? '', range.start) : characterBefore(buffer.value, range.start)
  } else {
    text = buffer.value.slice(range.start, range.end)
    for (const [position, field] of buffer.objects) if (position >= range.start && position < range.end) text += field.cachedResult ?? ''
  }
  // Resolving the story also validates an absent legacy rich-content fallback.
  documentStoryContent(section, range.story, model)
  return hasEastAsianLetters(text) ? { ...patch, eastAsiaFontFamily: patch.fontFamily } : patch
}

export interface DocumentStoryTextFormattingPlan {
  range: DocumentStoryRange
  typingPatch: RunPatch
  /** A caret changes transient input style only. */
  patch?: Partial<DocumentSectionHeaderFooter>
}

/** Prepare every literal and field update before mutating the section. */
export function prepareDocumentStoryTextFormatting(model: DocumentModel, request: DocumentStoryRange, patch: RunPatch, context: DocumentPermissionContext): DocumentStoryTextFormattingPlan {
  const { section, buffer, range, content } = readEditableRange(model, request, context)
  const typingPatch = documentStoryFontPatch(model, range, patch)
  if (range.start === range.end) return { range, typingPatch }
  const defaults = documentStoryDefaultStyle(model)
  const runs = formatDocumentRunRange(content.runs as DocumentRun[], buffer.literalOffsets[range.start]!, buffer.literalOffsets[range.end]!, typingPatch)
  const pageFields = section.pageFields?.map(field => {
    if (field.story !== range.story || ![...buffer.objects].some(([position, object]) => object === field && position >= range.start && position < range.end)) return field
    if (Object.hasOwn(typingPatch, 'href') || Object.hasOwn(typingPatch, 'metadata')) throw new TypeError('页面字段对象不能接受链接或来源对象替换。')
    const originalStyle = documentStoryTextStyle(field.textStyle ?? {}, defaults.text)
    const { id: _id, text: _text, href: _href, metadata: _metadata, ...style } = applyRunPatch(createRun('', originalStyle), typingPatch)
    const cache = documentStoryFieldCache(field)
    if (!cache) return sameTextStyle(style, originalStyle) ? field : { ...field, textStyle: style }
    const displayed = documentStoryFieldDisplayRuns(field, field.cachedResult ?? '', defaults.text)
    const changed = displayed.map(run => applyRunPatch(run, typingPatch))
    if (sameTextStyle(style, originalStyle) && JSON.stringify(changed) === JSON.stringify(displayed)) return field
    return { ...field, textStyle: style, metadata: { ...field.metadata,
      officeStoryStyleBaseline: { ...style }, officeStoryFieldCache: { version: 1, runs: changed } } }
  })
  if (JSON.stringify(runs) === JSON.stringify(content.runs) && JSON.stringify(pageFields) === JSON.stringify(section.pageFields)) return { range, typingPatch }
  return { range, typingPatch, patch: { pageFields, storyContent: { ...section.storyContent, [range.story]: { ...content, runs } } } }
}

export function applyDocumentStoryTextFormatting(model: DocumentModel, request: DocumentStoryRange, patch: RunPatch, context: DocumentPermissionContext): DocumentStoryTextFormattingPlan {
  const plan = prepareDocumentStoryTextFormatting(model, request, patch, context)
  if (plan.patch) applySectionHeaderFooter(model, request.sectionId, plan.patch)
  return plan
}
