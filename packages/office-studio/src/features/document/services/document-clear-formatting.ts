import type { DocumentDefaultStyle, DocumentModel, DocumentSectionHeaderFooter } from '../../../domain/model'
import {
  blockAtPoint, documentId, orderedSelection, runText, runsForTarget, sameTarget, textTargets,
  type DocumentRun, type DocumentSelection, type TextTarget,
} from '../model/document-model'
import { documentDefaultStyle } from '../model/document-default-style'
import { clearBlockNumbering } from '../model/document-numbering'
import { officeNumberingPrefixLength } from '../model/document-office-numbering-marker'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../permissions/document-permissions'
import { assertTextRangeEditable, getDocumentContentControls, transformContentControlsForBlockEdit } from './content-controls'
import { flattenDocumentBody } from './document-facade'
import { documentFormattingWordRange, documentPainterTarget } from './document-format-painter'
import { DOCUMENT_STORY_KEYS, documentStoryContent, documentStoryDefaultStyle, documentStoryFieldCache, isDocumentTextBoundary, type DocumentStoryRange } from '../model/document-story'
import { applySectionHeaderFooter, documentSectionIds, resolveSectionHeaderFooter, sectionStoryLinkedToPrevious } from './section-header-footer'
import { storyBuffer } from './story-editor'
import { assertDocumentStoryFormattingEditable } from './document-story-locks'
import { documentTableSelectedTextTargets } from '../model/document-table-selection'

export interface DocumentClearFormattingPlan {
  range: DocumentSelection
  targets: TextTarget[]
  defaults: DocumentDefaultStyle
}
export interface DocumentClearFormattingState { available: boolean; reason?: string }

/** Resolve the same word and half-open paragraph targets for UI and public commands. */
export function prepareDocumentClearFormatting(model: DocumentModel, selection: DocumentSelection, context: DocumentPermissionContext): DocumentClearFormattingPlan | undefined {
  const range = documentPainterTarget(model, selection)
  if (!range) return undefined
  const { start, end } = orderedSelection(model, range)
  const all = textTargets(model)
  const first = all.findIndex(target => sameTarget(target, start))
  const last = all.findIndex(target => sameTarget(target, end))
  if (first < 0 || last < first) return undefined
  const targets = range.tableSelection ? documentTableSelectedTextTargets(model, range.tableSelection)
    : all.slice(first, last + (last > first && end.offset === 0 ? 0 : 1))
  if (!targets.length) return undefined
  for (const target of targets) {
    const block = blockAtPoint(model, target)
    if (!block || !['paragraph', 'heading', 'quote', 'numbered', 'bullet', 'checklist'].includes(block.type)) return undefined
    if (block.metadata.codeBlock && runText(block.runs) !== Reflect.get(block.metadata.codeBlock, 'source')) throw new Error('代码源文本与格式范围不一致，请重新载入文档后重试。')
  }
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, range))
  // Paragraph defaults affect the entire paragraph, including locked content
  // outside the selected word. Check every paragraph before any mutation.
  if (model.contentControls?.length) {
    const body = flattenDocumentBody(model)
    const controls: DocumentModel = { ...model, contentControls: getDocumentContentControls(model) }
    for (const target of targets) {
      const segment = body.segments.find(segment => sameTarget(segment, target))
      if (segment) assertTextRangeEditable(model, segment.startOffset, segment.endOffset)
      const block = blockAtPoint(model, target)
      const prefix = block?.metadata.officeNumberingMarker ? officeNumberingPrefixLength(block) ?? 0 : 0
      if (prefix) transformContentControlsForBlockEdit(controls, target.cellBlockId ?? target.blockId, 0, prefix, 0)
    }
  }
  return { range, targets, defaults: documentDefaultStyle(model) }
}

export function documentClearFormattingState(model: DocumentModel, selection: DocumentSelection, context: DocumentPermissionContext, story?: DocumentStoryRange): DocumentClearFormattingState {
  try {
    return (story ? prepareDocumentStoryClearFormatting(model, story, context) : prepareDocumentClearFormatting(model, selection, context))
      ? { available: true }
      : { available: false, reason: '请选中文字，或将光标放在要清除格式的词中。' }
  } catch (error) {
    return { available: false, reason: error instanceof Error ? error.message : '此选区无法清除格式。' }
  }
}

export interface DocumentStoryClearFormattingPlan { range: DocumentStoryRange; patch: Partial<DocumentSectionHeaderFooter> }

/** Story offsets are in the input buffer, where each field occupies one object marker. */
export function prepareDocumentStoryClearFormatting(model: DocumentModel, request: DocumentStoryRange, context: DocumentPermissionContext): DocumentStoryClearFormattingPlan | undefined {
  if (!request || !DOCUMENT_STORY_KEYS.includes(request.story) || !documentSectionIds(model).includes(request.sectionId)) throw new RangeError('页眉页脚片段不存在。')
  const section = resolveSectionHeaderFooter(model, request.sectionId)
  if (sectionStoryLinkedToPrevious(section, request.story)) throw new Error('请先取消此页眉页脚与前节的链接，再清除格式。')
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', [
    { scope: 'section', id: request.sectionId }, { scope: 'entity', id: `${request.sectionId}:${request.story}` },
    { scope: 'entity', id: `${request.sectionId}:${request.story.toLowerCase().includes('header') ? 'header' : 'footer'}` },
  ])
  const buffer = storyBuffer(section, request.story)
  if (!isDocumentTextBoundary(buffer.value, request.start) || !isDocumentTextBoundary(buffer.value, request.end)) throw new RangeError('页眉页脚选区的 UTF-16 位置无效。')
  let start = Math.min(request.start, request.end), end = Math.max(request.start, request.end)
  if (start === end) {
    const word = documentFormattingWordRange(buffer.value, start)
    if (!word) return undefined
    start = word.start; end = word.end
  }
  const content = documentStoryContent(section, request.story, model), defaults = documentStoryDefaultStyle(model)
  const paragraphs = content.paragraphs.map((paragraph, index) => {
    const from = buffer.literalOffsets.findIndex(offset => offset === paragraph.startOffset)
    const to = index === content.paragraphs.length - 1 ? buffer.value.length
      : buffer.literalOffsets.findIndex((offset, position) => offset === paragraph.endOffset && buffer.value[position] === '\n')
    if (start >= to + 1 || end <= from) return paragraph
    assertDocumentStoryFormattingEditable(content, buffer.value, from, Math.min(buffer.value.length, to + 1))
    if (paragraph.metadata?.contentLocked === true) throw new Error('选中的页眉页脚段落含已锁定内容。')
    const metadata = { ...paragraph.metadata }
    for (const field of ['namedStyleType', 'styleId', 'numbering', 'list', 'html']) delete metadata[field]
    if (hasStorySourceFormat(metadata)) metadata.officeStoryParagraphFormatCleared = true
    return { ...paragraph, style: { ...defaults.paragraph }, ...(Object.keys(metadata).length ? { metadata } : { metadata: undefined }) }
  })
  const pageFields = section.pageFields?.map(field => {
    if (field.story !== request.story || ![...buffer.objects].some(([offset, object]) => object === field && offset >= start && offset < end)) return field
    const metadata = { ...field.metadata }, cache = documentStoryFieldCache(field)
    // A field is one selected object, but its cached text can contain links and
    // other entities. Clear every cached run without collapsing those ranges.
    if (cache) metadata.officeStoryFieldCache = { version: 1, runs: cache.runs.map(run => clearedRun(run, defaults.text)) }
    if (hasStorySourceFormat(metadata)) metadata.officeStoryTextFormatCleared = true
    return { ...field, textStyle: { ...defaults.text }, ...(Object.keys(metadata).length ? { metadata } : {}) }
  })
  return {
    range: { ...request, start, end },
    patch: { pageFields, storyContent: { ...section.storyContent, [request.story]: {
      ...content,
      runs: clearDocumentFormattingRuns(content.runs, buffer.literalOffsets[start]!, buffer.literalOffsets[end]!, defaults.text), paragraphs,
    } } },
  }
}

export function applyDocumentStoryClearFormatting(model: DocumentModel, story: DocumentStoryRange, context: DocumentPermissionContext): DocumentStoryRange | undefined {
  const plan = prepareDocumentStoryClearFormatting(model, story, context)
  if (plan) applySectionHeaderFooter(model, story.sectionId, plan.patch)
  return plan?.range
}

function hasStorySourceFormat(metadata: Record<string, unknown>): boolean {
  const source = metadata.source
  if (!source || typeof source !== 'object' || Array.isArray(source)) return false
  const details = Reflect.get(source, 'details')
  return Reflect.get(source, 'sourceFormat') !== undefined || !!details && typeof details === 'object' && !Array.isArray(details) && Reflect.get(details, 'sourceFormat') !== undefined
}

function clearedRun(run: DocumentRun, defaults: DocumentDefaultStyle['text'], disableCodeSyntax = false): DocumentRun {
  const next: DocumentRun = { ...run, ...defaults }
  // An omitted East Asian slot already uses the Latin family. Keep that
  // representation stable so clearing an unformatted run stays a no-op.
  if (run.eastAsiaFontFamily === undefined && next.eastAsiaFontFamily === next.fontFamily) delete next.eastAsiaFontFamily
  if (!defaults.verticalAlign || defaults.verticalAlign === 'baseline') delete next.verticalAlign
  if (!defaults.highlight) delete next.highlight
  // Parser provenance and entities stay intact; direct legacy appearance must
  // not restore highlighting or vertical alignment on save and reopen.
  if (run.metadata) {
    const metadata = { ...run.metadata }
    for (const key of ['bold', 'italic', 'underline', 'strike', 'color', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'highlight', 'verticalAlign']) delete metadata[key]
    delete metadata.univerTextStyle
    delete metadata.univerTextStyleResolution
    if (defaults.highlight) metadata.highlight = defaults.highlight
    if (disableCodeSyntax) metadata.codeSyntaxHighlight = false
    if (hasStorySourceFormat(metadata)) metadata.officeStoryTextFormatCleared = true
    if (Object.keys(metadata).length) next.metadata = metadata
    else delete next.metadata
  } else if (defaults.highlight || disableCodeSyntax) next.metadata = { ...(defaults.highlight ? { highlight: defaults.highlight } : {}), ...(disableCodeSyntax ? { codeSyntaxHighlight: false } : {}) }
  return next
}

/** Keep zero-length notes/fields and untouched run identities; split only changed text. */
export function clearDocumentFormattingRuns(runs: readonly DocumentRun[], start: number, end: number, defaults: DocumentDefaultStyle['text'], disableCodeSyntax = false): DocumentRun[] {
  const result: DocumentRun[] = []
  let offset = 0
  for (const run of runs) {
    const from = Math.max(0, start - offset), to = Math.min(run.text.length, end - offset)
    offset += run.text.length
    if (to <= from) { result.push(run); continue }
    const next = clearedRun(run, defaults, disableCodeSyntax)
    if (JSON.stringify(next) === JSON.stringify(run)) { result.push(run); continue }
    if (from > 0) result.push({ ...run, text: run.text.slice(0, from) })
    result.push({ ...next, id: from > 0 ? documentId('run') : run.id, text: run.text.slice(from, to) })
    if (to < run.text.length) result.push({ ...run, id: documentId('run'), text: run.text.slice(to) })
  }
  return result
}

/** The command has already validated the complete plan and removed cached list labels. */
export function applyDocumentClearFormatting(model: DocumentModel, range: DocumentSelection, targets: readonly TextTarget[], defaults: DocumentDefaultStyle): void {
  const { start, end } = orderedSelection(model, range)
  for (const target of targets) {
    const block = blockAtPoint(model, target)
    if (!block) continue
    const runs = runsForTarget(model, target)
    const from = !range.tableSelection && sameTarget(target, start) ? start.offset : 0
    const to = !range.tableSelection && sameTarget(target, end) ? end.offset : runText(runs).length
    block.runs = clearDocumentFormattingRuns(runs, from, to, defaults.text, !!block.metadata.codeBlock)
    block.style = { ...defaults.paragraph }
    if (hasStorySourceFormat(block.metadata)) block.metadata.officeStoryParagraphFormatCleared = true
    block.type = 'paragraph'
    block.level = 0
    clearBlockNumbering(block)
    delete block.checklist
    delete block.metadata.namedStyleType
    delete block.metadata.html
    if (target.row !== undefined) {
      const container = model.blocks.find(candidate => candidate.id === target.blockId)
      if (container) delete container.metadata.html
    }
  }
}
