import type { DocumentModel, DocumentPageField, DocumentSectionHeaderFooter } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { sectionStoryLinkedToPrevious } from '../model/document-section-stories'
import { DOCUMENT_STORY_KEYS, documentStoryContent, documentStoryDefaultStyle, documentStoryFieldCache, documentStoryFieldDisplayRuns, isDocumentTextBoundary, remapDocumentStoryFieldRuns, validateDocumentStoryLocks, type DocumentStoryKey, type DocumentStoryRange, type StoryDefaultsSource } from '../model/document-story'
import { assertDocumentPermission, permissionTargetsForBlock, permissionTargetsForChangedBlocks, resolveDocumentPermission, type DocumentPermissionContext, type DocumentPermissionTarget } from '../permissions/document-permissions'
import { blockAtPoint, runText, sameTarget, sliceRuns, textTargets, type DocumentPoint, type DocumentRun, type DocumentSelection } from '../model/document-model'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentRunStyleAt, splitDocumentRunRange } from '../model/document-run-range'
import { documentFieldAutomaticRefreshBlocked, documentTextFieldParagraphs, documentTextFieldRangeContainsNested, documentTextFieldRangeIntersects, documentTextFieldRanges, documentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockRangeEdit, type DocumentTextFieldParagraph, type DocumentTextFieldRange } from '../model/document-text-fields'
import { getDocumentContentControls, transformContentControlsForBlockEdit, transformContentControlsForBodyRangeEdit } from './content-controls'
import { transformDocumentCommentAnchorsForBodyRangeEdit, transformDocumentCommentAnchorsForTextEdit } from '../comments/document-comment-anchors'
import { resolveDocumentDateTimeFieldResult } from './document-date-time-field'
import { applySectionHeaderFooter, documentSectionIds, resolveSectionHeaderFooter } from './section-header-footer'
import { storyBuffer } from './story-editor'

export interface DocumentStoryFieldRefreshRequest {
  locale: string
  /** One time snapshot for the entire content-history transaction. */
  timestamp: number
  /** Undefined selects all areas unless bodySelection is present; [] selects no story fields. */
  stories?: DocumentStoryRange[]
  /** Selection-only body update; a collapsed caret inside a result selects that field. */
  bodySelection?: DocumentSelection
  /** The caller publishes the same content transaction without adding a history entry. */
  onOpen?: boolean
}

export interface DocumentStoryFieldRefreshPlan {
  sectionHeaderFooters: DocumentModel['sectionHeaderFooters']
  blocks?: DocumentModel['blocks']
  contentControls?: DocumentModel['contentControls']
  comments?: DocumentModel['comments']
  /** Current body caret/range mapped through every changed result. */
  selection?: DocumentSelection
  updatedFields: number
}

/** Keep cached run boundaries when a refreshed result changes within several formats. */
export function refreshDocumentStoryFieldResult(field: DocumentPageField, result: string, model?: StoryDefaultsSource): DocumentPageField {
  if (result === field.cachedResult) return field
  const cache = documentStoryFieldCache(field)
  if (!cache) return cloneModel({ ...field, cachedResult: result })
  const defaults = documentStoryDefaultStyle(model)
  const runs = documentStoryFieldDisplayRuns(field, result, defaults.text)
  const next = cloneModel({ ...field, cachedResult: result, metadata: { ...field.metadata,
    ...(field.textStyle ? { officeStoryStyleBaseline: { ...field.textStyle } } : {}), officeStoryFieldCache: { version: 1, runs } } })
  documentStoryFieldCache(next)
  return next
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function sourceLanguage(style: unknown): string | undefined {
  return isRecord(style) && typeof style.language === 'string' ? style.language : undefined
}

/** Parsed cache styles retain duke-pages' resolved w:lang in their provenance. */
function storyDateLanguage(field: DocumentPageField, fallback: string): string {
  const runs = documentStoryFieldCache(field)?.runs
  const first = runs?.find(run => run.text.length > 0)
  const runSource = first?.metadata?.source, fieldSource = field.metadata?.source
  const emptySource = isRecord(runSource) && runSource.text === ''
  if (field.metadata?.officeStoryTextFormatCleared === true || !emptySource && first?.metadata?.officeStoryTextFormatCleared === true) return fallback
  // Expanding an empty cache retains its source style. That empty source must
  // not become the language authority on the next refresh.
  return sourceLanguage(isRecord(runSource) && !emptySource ? runSource.style : undefined)
    ?? sourceLanguage(isRecord(fieldSource) ? fieldSource.textStyle : undefined)
    ?? fallback
}

/** Explicit date-time pictures are refreshed only in the selected, unlinked story. */
export function refreshDocumentStoryDateFields(section: DocumentSectionHeaderFooter, story: DocumentStoryKey, locale: string, date: Date, model?: StoryDefaultsSource): DocumentPageField[] | undefined {
  if (sectionStoryLinkedToPrevious(section, story)) throw new Error('请先取消此页眉页脚与前节的链接，再更新字段。')
  return section.pageFields?.map(field => {
    if (field.story !== story || field.field !== 'DATE' && field.field !== 'TIME' || documentFieldAutomaticRefreshBlocked(field)) return field
    const result = resolveDocumentDateTimeFieldResult(field, storyDateLanguage(field, locale), date)
    return result === undefined ? field : refreshDocumentStoryFieldResult(field, result, model)
  })
}

function validateStoryRange(buffer: string, range: Pick<DocumentStoryRange, 'start' | 'end'>): void {
  if (!range || !isDocumentTextBoundary(buffer, range.start) || !isDocumentTextBoundary(buffer, range.end) || range.end < range.start) {
    throw new RangeError('字段更新选区的 UTF-16 位置无效。')
  }
}

/** A caret on the object marker selects the same field as a range covering that marker. */
export function documentStoryRangeHasField(section: DocumentSectionHeaderFooter, story: DocumentStoryKey, range: Pick<DocumentStoryRange, 'start' | 'end'>): boolean {
  const buffer = storyBuffer(section, story)
  validateStoryRange(buffer.value, range)
  return [...buffer.objects.keys()].some(position => storyRangeIntersectsField(range, position))
}

function storyRangeIntersectsField(range: Pick<DocumentStoryRange, 'start' | 'end'>, position: number): boolean {
  return range.start === range.end ? range.start === position : range.start <= position && range.end > position
}

function sourceStoryPermissionTargets(sections: readonly DocumentSectionHeaderFooter[], sectionIds: readonly string[], index: number, story: DocumentStoryKey): DocumentPermissionTarget[] {
  const result: DocumentPermissionTarget[] = []
  const area = story.toLowerCase().includes('header') ? 'header' : 'footer'
  // A source edit changes every following linked story until its next override.
  for (let target = index; target < sections.length; target++) {
    if (target > index && !sectionStoryLinkedToPrevious(sections[target]!, story)) break
    const sectionId = sectionIds[target]!
    result.push({ scope: 'section', id: sectionId }, { scope: 'entity', id: `${sectionId}:${story}` }, { scope: 'entity', id: `${sectionId}:${area}` })
  }
  return result
}

function editableStoryDateFields(section: DocumentSectionHeaderFooter, story: DocumentStoryKey, model: StoryDefaultsSource, ranges?: readonly DocumentStoryRange[]): DocumentPageField[] {
  const buffer = storyBuffer(section, story)
  const content = documentStoryContent(section, story, model)
  validateDocumentStoryLocks(content.locks, buffer.value)
  return [...buffer.objects].filter(([position, field]) => (field.field === 'DATE' || field.field === 'TIME')
    && !documentFieldAutomaticRefreshBlocked(field)
    && (!ranges || ranges.some(range => storyRangeIntersectsField(range, position)))
    && !content.locks?.some(lock => lock.lock !== 'sdtLocked' && position < lock.endOffset && lock.startOffset < position + 1)).map(([, field]) => field)
}

/** The global update menu remains usable from a linked or protected current story. */
export function canRefreshDocumentStoryFields(model: DocumentModel, context: DocumentPermissionContext): boolean {
  if (editableBodyDateFields(model, context).length > 0) return true
  const sectionIds = documentSectionIds(model), sections = sectionIds.map(sectionId => resolveSectionHeaderFooter(model, sectionId))
  return sections.some((section, index) => DOCUMENT_STORY_KEYS.some(story => !sectionStoryLinkedToPrevious(section, story)
    && resolveDocumentPermission(model.permissionRules ?? [], context, 'edit', sourceStoryPermissionTargets(sections, sectionIds, index, story)).allowed
    && editableStoryDateFields(section, story, model).length > 0))
}

interface BodySelectionRanges { start: number; end: number }

function bodySelectionRanges(model: DocumentModel, selection: DocumentSelection): BodySelectionRanges {
  if (!selection || !selection.anchor || !selection.focus) throw new TypeError('正文字段更新选区无效。')
  const targets = textTargets(model)
  for (const point of [selection.anchor, selection.focus]) {
    const target = targets.find(target => sameTarget(target, point)), block = target && blockAtPoint(model, target)
    if (!block || !isDocumentTextBoundary(runText(block.runs), point.offset)) throw new RangeError('正文字段更新选区的 UTF-16 位置无效。')
  }
  const paragraphs = new Map(documentTextFieldParagraphs(model).map(paragraph => [paragraph.block.id, paragraph]))
  const anchor = paragraphs.get(selection.anchor.cellBlockId ?? selection.anchor.blockId)
  const focus = paragraphs.get(selection.focus.cellBlockId ?? selection.focus.blockId)
  if (!anchor || !focus) throw new RangeError('正文字段更新选区的段落已不存在。')
  const anchorOffset = anchor.absoluteStart + selection.anchor.offset, focusOffset = focus.absoluteStart + selection.focus.offset
  return { start: Math.min(anchorOffset, focusOffset), end: Math.max(anchorOffset, focusOffset) }
}

/** Availability can be queried between a model update and its mapped selection. */
function availableBodySelectionRanges(model: DocumentModel, selection: DocumentSelection): BodySelectionRanges | undefined {
  try { return bodySelectionRanges(model, selection) }
  catch (error) {
    if (error instanceof RangeError || error instanceof TypeError) return undefined
    throw error
  }
}

export function documentBodySelectionHasField(model: DocumentModel, selection: DocumentSelection): boolean {
  const ranges = availableBodySelectionRanges(model, selection)
  if (!ranges) return false
  return documentTextFieldRanges(model).some(field => documentTextFieldRangeIntersects(field, ranges.start, ranges.end))
}

function bodyFieldContentEditable(model: DocumentModel, range: DocumentTextFieldRange, paragraphs: ReadonlyMap<string, DocumentTextFieldParagraph>): boolean {
  return !(model.contentControls ?? []).some(control => {
    const paragraph = paragraphs.get(control.blockId)
    if (!paragraph || !['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')) return false
    const start = paragraph.absoluteStart + control.startOffset, end = paragraph.absoluteStart + control.endOffset
    const overlaps = range.absoluteStart === range.absoluteEnd
      ? range.absoluteStart >= start && range.absoluteStart <= end : range.absoluteStart < end && start < range.absoluteEnd
    if (!overlaps) return false
    if (control.type === 'group' && (model.contentControls ?? []).some(child => child.parentId === control.id && child.blockId === control.blockId
      && paragraph.absoluteStart + child.startOffset <= range.absoluteStart
      && paragraph.absoluteStart + child.endOffset >= range.absoluteEnd)) return false
    return true
  })
}

function bodyFieldParagraphs(range: DocumentTextFieldRange, paragraphs: readonly DocumentTextFieldParagraph[]): DocumentTextFieldParagraph[] {
  return paragraphs.filter(paragraph => paragraph.absoluteStart >= range.startParagraph.absoluteStart && paragraph.absoluteStart <= range.endParagraph.absoluteStart)
}

function bodyFieldContainsInlineEntity(range: DocumentTextFieldRange, paragraphs: readonly DocumentTextFieldParagraph[]): boolean {
  return bodyFieldParagraphs(range, paragraphs).some(paragraph => {
    let offset = paragraph.absoluteStart
    return paragraph.block.runs.some(run => {
      const start = offset; offset += run.text.length
      const nested = run.metadata?.note || run.metadata?.noteReference || run.metadata?.latexFormula
      return nested && (run.text.length ? start < range.absoluteEnd && offset > range.absoluteStart : start > range.absoluteStart && start < range.absoluteEnd)
    })
  })
}

function bodyFieldCanMerge(range: DocumentTextFieldRange): boolean {
  const first = range.startParagraph, last = range.endParagraph
  return first.container === last.container && first.container.slice(first.index, last.index + 1)
    .every(block => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)
      && (first.block.id === last.block.id || block.metadata.codeBlock === undefined))
}

function editableBodyDateFields(model: DocumentModel, context: DocumentPermissionContext, selection?: BodySelectionRanges): DocumentTextFieldRange[] {
  const paragraphs = documentTextFieldParagraphs(model), byId = new Map(paragraphs.map(paragraph => [paragraph.block.id, paragraph]))
  const ranges = documentTextFieldRanges(model), result: DocumentTextFieldRange[] = []
  for (const range of ranges) {
    const field = range.field
    if (field.field !== 'DATE' && field.field !== 'TIME' || documentFieldAutomaticRefreshBlocked(field)
      || documentTextFieldRangeContainsNested(range, ranges) || bodyFieldContainsInlineEntity(range, paragraphs) || !bodyFieldCanMerge(range)
      || selection && !documentTextFieldRangeIntersects(range, selection.start, selection.end)
      || !bodyFieldContentEditable(model, range, byId)) continue
    const targets = bodyFieldParagraphs(range, paragraphs).flatMap(paragraph => permissionTargetsForBlock(model, paragraph.block.id))
    for (const outer of ranges) if (outer.field.id !== field.id && outer.absoluteStart <= range.absoluteStart && outer.absoluteEnd >= range.absoluteEnd) {
      targets.push(...permissionTargetsForBlock(model, outer.block.id), { scope: 'entity', id: outer.field.id })
    }
    if (!resolveDocumentPermission(model.permissionRules ?? [], context, 'edit', [...targets, { scope: 'entity', id: field.id }]).allowed) continue
    result.push(range)
  }
  return result
}

export function canRefreshSelectedDocumentBodyFields(model: DocumentModel, selection: DocumentSelection, context: DocumentPermissionContext): boolean {
  const ranges = availableBodySelectionRanges(model, selection)
  return !!ranges && editableBodyDateFields(model, context, ranges).length > 0
}

function bodyDateLanguage(range: DocumentTextFieldRange, fallback: string): string {
  const block = range.block, field = range.field
  let offset = 0
  const first = range.absoluteStart < range.absoluteEnd && field.startOffset < range.startParagraph.text.length ? block.runs.find(run => {
    const from = offset; offset += run.text.length
    return from <= field.startOffset && offset > field.startOffset
  }) : undefined
  if (first?.metadata?.officeStoryTextFormatCleared === true || field.metadata?.officeStoryTextFormatCleared === true) return fallback
  const source = first?.metadata?.source, fieldSource = field.metadata?.source
  return (first?.metadata?.officeTextFieldStyleOutsideResult === true ? undefined : sourceLanguage(isRecord(source) ? source.style : undefined)) ?? field.language
    ?? sourceLanguage(isRecord(fieldSource) ? fieldSource.textStyle : undefined) ?? fallback
}

function mappedSelectionPoint(point: DocumentPoint, range: DocumentTextFieldRange, paragraphs: readonly DocumentTextFieldParagraph[], length: number, model: DocumentModel): DocumentPoint {
  const paragraph = paragraphs.find(candidate => candidate.block.id === (point.cellBlockId ?? point.blockId))
  if (!paragraph) return point
  const absolute = paragraph.absoluteStart + point.offset
  if (absolute <= range.absoluteStart || absolute >= range.absoluteEnd && paragraph.block.id !== range.endParagraph.block.id) return point
  const target = textTargets(model).find(candidate => (candidate.cellBlockId ?? candidate.blockId) === range.block.id)
  if (!target) throw new RangeError('字段更新后的正文输入目标不可用。')
  const offset = absolute >= range.absoluteEnd ? range.field.startOffset + length + absolute - range.absoluteEnd
    : range.field.startOffset + Math.min(absolute - range.absoluteStart, length)
  return { ...target, offset }
}

function bodyFieldCacheRuns(range: DocumentTextFieldRange, paragraphs: readonly DocumentTextFieldParagraph[]): DocumentRun[] {
  const cache: DocumentRun[] = []
  const selected = bodyFieldParagraphs(range, paragraphs)
  for (const [index, paragraph] of selected.entries()) {
    const runs = paragraph.block.runs as DocumentRun[]
    const start = index === 0 ? range.field.startOffset : 0, end = index === selected.length - 1 ? range.field.endOffset : paragraph.text.length
    cache.push(...sliceRuns(runs, start, end))
    if (index < selected.length - 1) {
      const inherited = cache.at(-1) ?? documentRunStyleAt(runs, start, end)
      cache.push({ ...inherited, id: `${range.field.id}:paragraph:${index}`, text: '\n',
        metadata: { ...inherited?.metadata, officeTextFieldStyleOutsideResult: true } })
    }
  }
  return cache
}

function refreshedBodyRuns(range: DocumentTextFieldRange, paragraphs: readonly DocumentTextFieldParagraph[], result: string, model: DocumentModel): DocumentRun[] {
  const runs = range.block.runs as DocumentRun[], field = range.field, defaults = documentDefaultStyle(model).text
  const before = splitDocumentRunRange(runs, field.startOffset, range.block.id === range.endParagraph.block.id ? field.endOffset : runText(runs).length).before
  const after = splitDocumentRunRange(range.endParagraph.block.runs as DocumentRun[], range.block.id === range.endParagraph.block.id ? field.startOffset : 0, field.endOffset).after
  const cache = bodyFieldCacheRuns(range, paragraphs)
  const inherited = documentRunStyleAt(runs, field.startOffset, range.block.id === range.endParagraph.block.id ? field.endOffset : runText(runs).length) ?? defaults
  const styled = cache.length ? cache : [{ ...inherited, id: `${field.id}:result`, text: '',
    metadata: { ...inherited.metadata, officeTextFieldStyleOutsideResult: true } } as DocumentRun]
  return [...before, ...remapDocumentStoryFieldRuns(styled, result, defaults), ...after]
}

/** Prepare all changed caches without modifying the live model or creating overrides for linked stories. */
export function prepareDocumentStoryFieldRefresh(model: DocumentModel, request: DocumentStoryFieldRefreshRequest, context: DocumentPermissionContext, selection?: DocumentSelection): DocumentStoryFieldRefreshPlan | undefined {
  if (!request || typeof request.locale !== 'string' || !request.locale.trim() || typeof request.timestamp !== 'number'
    || !Number.isFinite(request.timestamp) || !Number.isFinite(new Date(request.timestamp).getTime())) throw new TypeError('字段更新时间或语言无效。')
  // Reject invalid command parameters; an unsupported source language still preserves its cache.
  try { Intl.getCanonicalLocales(request.locale) } catch { throw new TypeError('字段更新语言无效。') }
  if (request.stories !== undefined && !Array.isArray(request.stories)) throw new TypeError('字段更新选区列表无效。')
  if (request.onOpen !== undefined && typeof request.onOpen !== 'boolean') throw new TypeError('字段自动更新标志必须为布尔值。')
  const bodyRanges = request.bodySelection !== undefined ? bodySelectionRanges(model, request.bodySelection) : undefined
  const currentSelection = selection ?? request.bodySelection
  if (currentSelection) bodySelectionRanges(model, currentSelection)
  const selectedAreas = request.stories !== undefined || request.bodySelection !== undefined
  const bodyFields = !selectedAreas || bodyRanges ? editableBodyDateFields(model, context, bodyRanges) : []
  const sectionIds = documentSectionIds(model)
  const sections = sectionIds.map(sectionId => resolveSectionHeaderFooter(model, sectionId))
  for (const range of request.stories ?? []) {
    if (!range || typeof range.sectionId !== 'string' || !DOCUMENT_STORY_KEYS.includes(range.story)) throw new TypeError('字段更新片段无效。')
    const index = sectionIds.indexOf(range.sectionId)
    if (index < 0) throw new RangeError('字段更新片段所在的节不存在。')
    validateStoryRange(storyBuffer(sections[index]!, range.story).value, range)
  }
  const candidate: DocumentModel = { ...model, sectionHeaderFooters: model.sectionHeaderFooters ? { ...model.sectionHeaderFooters } : undefined,
    ...(bodyFields.length ? { blocks: cloneModel(model.blocks), contentControls: model.contentControls && getDocumentContentControls(model), comments: model.comments && cloneModel(model.comments) } : {}) }
  const date = new Date(request.timestamp)
  let updatedFields = 0
  let nextSelection = currentSelection && cloneModel(currentSelection)
  for (const sourceRange of [...bodyFields].sort((a, b) => b.absoluteStart - a.absoluteStart)) {
    const range = documentTextFieldRanges(candidate).find(current => current.field.id === sourceRange.field.id)!
    const block = range.block, field = range.field, paragraphs = documentTextFieldParagraphs(candidate)
    const language = bodyDateLanguage(range, request.locale)
    const result = resolveDocumentDateTimeFieldResult({ ...field, field: field.field as 'DATE' | 'TIME' }, language, date)
    const visibleResult = runText(bodyFieldCacheRuns(range, paragraphs))
    if (result === undefined || result === field.cachedResult && visibleResult === result) continue
    if (result === visibleResult) {
      setDocumentTextFields(block, documentTextFields(block).map(current => current.id === field.id ? { ...current, cachedResult: result } : current))
      updatedFields++
      continue
    }
    const nextRuns = refreshedBodyRuns(range, paragraphs, result, candidate)
    const codeBlock = block.metadata.codeBlock
    if (isRecord(codeBlock) && typeof codeBlock.source === 'string' && codeBlock.source !== runText(block.runs)) throw new Error('代码源文本与字段结果范围不一致。')
    const crossParagraph = range.startParagraph.block.id !== range.endParagraph.block.id
    if (!crossParagraph) transformContentControlsForBlockEdit(candidate, block.id, field.startOffset, field.endOffset, result.length, undefined, true)
    const mappedSelection = nextSelection && {
      anchor: mappedSelectionPoint(nextSelection.anchor, range, paragraphs, result.length, candidate),
      focus: mappedSelectionPoint(nextSelection.focus, range, paragraphs, result.length, candidate),
    }
    const first = range.startParagraph, last = range.endParagraph
    transformDocumentTextFieldsForBlockRangeEdit(first.container.slice(first.index, last.index + 1), field.startOffset, field.endOffset, result, [block], field.id, candidate)
    block.runs = nextRuns
    if (crossParagraph) first.container.splice(first.index + 1, last.index - first.index)
    if (isRecord(codeBlock) && typeof codeBlock.source === 'string') block.metadata.codeBlock = { ...codeBlock, source: runText(nextRuns) }
    setDocumentTextFields(block, documentTextFields(block).map(current => current.id === field.id ? { ...current, cachedResult: result,
      ...(field.language === undefined && (range.absoluteStart === range.absoluteEnd || field.startOffset === first.text.length) ? { language } : {}) } : current))
    delete block.metadata.html
    if (crossParagraph) {
      transformContentControlsForBodyRangeEdit(candidate, paragraphs, range.absoluteStart, range.absoluteEnd, result.length)
      transformDocumentCommentAnchorsForBodyRangeEdit(candidate, paragraphs, documentTextFieldParagraphs(candidate), range.absoluteStart, range.absoluteEnd, result, request.timestamp)
    } else transformDocumentCommentAnchorsForTextEdit(candidate, block.id, field.startOffset, field.endOffset, result, request.timestamp)
    nextSelection = mappedSelection
    updatedFields++
  }
  const bodyUpdated = updatedFields > 0
  for (const [index, sectionId] of sectionIds.entries()) {
    const section = resolveSectionHeaderFooter(candidate, sectionId)
    const replacements = new Map<DocumentPageField, DocumentPageField>()
    for (const story of DOCUMENT_STORY_KEYS) {
      if (sectionStoryLinkedToPrevious(section, story)) continue
      const ranges = selectedAreas ? (request.stories ?? []).filter(range => range.sectionId === sectionId && range.story === story) : undefined
      if (ranges?.length === 0) continue
      const targets = sourceStoryPermissionTargets(sections, sectionIds, index, story)
      if (!resolveDocumentPermission(model.permissionRules ?? [], context, 'edit', targets).allowed) continue
      for (const field of editableStoryDateFields(section, story, model, ranges)) {
        const result = resolveDocumentDateTimeFieldResult(field, storyDateLanguage(field, request.locale), date)
        if (result === undefined || result === field.cachedResult) continue
        replacements.set(field, refreshDocumentStoryFieldResult(field, result, model))
      }
    }
    if (!replacements.size) continue
    applySectionHeaderFooter(candidate, sectionId, { pageFields: section.pageFields?.map(field => replacements.get(field) ?? field) })
    updatedFields += replacements.size
  }
  if (!updatedFields) return undefined
  // Keep the same final inheritance-aware permission contract as other content edits.
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, candidate))
  return { sectionHeaderFooters: candidate.sectionHeaderFooters, updatedFields, ...(nextSelection ? { selection: nextSelection } : {}),
    ...(bodyUpdated ? { blocks: candidate.blocks, contentControls: candidate.contentControls, comments: candidate.comments } : {}) }
}
