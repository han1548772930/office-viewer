import type { DocumentBlock, DocumentModel, DocumentPageField, DocumentSectionHeaderFooter } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { sectionStoryLinkedToPrevious } from '../model/document-section-stories'
import { DOCUMENT_STORY_KEYS, documentStoryContent, documentStoryDefaultStyle, documentStoryFieldCache, documentStoryFieldDisplayRuns, isDocumentTextBoundary, remapDocumentStoryFieldRuns, validateDocumentStoryLocks, type DocumentStoryKey, type DocumentStoryRange, type StoryDefaultsSource } from '../model/document-story'
import { assertDocumentPermission, permissionTargetsForBlock, permissionTargetsForChangedBlocks, resolveDocumentPermission, type DocumentPermissionContext, type DocumentPermissionTarget } from '../permissions/document-permissions'
import { blockAtPoint, runText, sameTarget, sliceRuns, textTargets, type DocumentPoint, type DocumentRun, type DocumentSelection } from '../model/document-model'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentRunStyleAt, splitDocumentRunRange } from '../model/document-run-range'
import { documentFieldAutomaticRefreshBlocked, documentTextFieldBlocks, documentTextFieldContainsNestedRange, documentTextFieldIntersects, documentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit, type DocumentTextField } from '../model/document-text-fields'
import { getDocumentContentControls, transformContentControlsForBlockEdit } from './content-controls'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments'
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

function bodySelectionRanges(model: DocumentModel, selection: DocumentSelection): Map<string, { start: number; end: number }> {
  if (!selection || !selection.anchor || !selection.focus) throw new TypeError('正文字段更新选区无效。')
  const targets = textTargets(model)
  for (const point of [selection.anchor, selection.focus]) {
    const target = targets.find(target => sameTarget(target, point)), block = target && blockAtPoint(model, target)
    if (!block || !isDocumentTextBoundary(runText(block.runs), point.offset)) throw new RangeError('正文字段更新选区的 UTF-16 位置无效。')
  }
  const anchor = targets.findIndex(target => sameTarget(target, selection.anchor))
  const focus = targets.findIndex(target => sameTarget(target, selection.focus))
  const forward = anchor < focus || anchor === focus && selection.anchor.offset <= selection.focus.offset
  const first = forward ? anchor : focus, last = forward ? focus : anchor
  const start = forward ? selection.anchor : selection.focus, end = forward ? selection.focus : selection.anchor
  return new Map(targets.slice(first, last + 1).map((target, index) => {
    const block = blockAtPoint(model, target)!
    return [block.id, { start: index === 0 ? start.offset : 0, end: first + index === last ? end.offset : runText(block.runs).length }]
  }))
}

export function documentBodySelectionHasField(model: DocumentModel, selection: DocumentSelection): boolean {
  const ranges = bodySelectionRanges(model, selection)
  return documentTextFieldBlocks(model).some(block => {
    const range = ranges.get(block.id)
    return range && documentTextFields(block).some(field => documentTextFieldIntersects(field, range.start, range.end))
  })
}

function bodyFieldContentEditable(model: DocumentModel, block: DocumentBlock, field: DocumentTextField): boolean {
  return !(model.contentControls ?? []).some(control => {
    if (control.blockId !== block.id || !['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')) return false
    if (control.type === 'group' && (model.contentControls ?? []).some(child => child.parentId === control.id
      && child.startOffset <= field.startOffset && child.endOffset >= field.endOffset)) return false
    return field.startOffset === field.endOffset
      ? field.startOffset >= control.startOffset && field.startOffset <= control.endOffset
      : field.startOffset < control.endOffset && control.startOffset < field.endOffset
  })
}

interface BodyDateField { blockId: string; field: DocumentTextField }

function bodyFieldContainsInlineEntity(block: DocumentBlock, field: DocumentTextField): boolean {
  let offset = 0
  return block.runs.some(run => {
    const start = offset; offset += run.text.length
    const nested = run.metadata?.note || run.metadata?.noteReference || run.metadata?.latexFormula
    return nested && (run.text.length ? start < field.endOffset && offset > field.startOffset : start > field.startOffset && start < field.endOffset)
  })
}

function editableBodyDateFields(model: DocumentModel, context: DocumentPermissionContext, selection?: DocumentSelection): BodyDateField[] {
  const ranges = selection && bodySelectionRanges(model, selection)
  const result: BodyDateField[] = []
  for (const block of documentTextFieldBlocks(model)) {
    const fields = documentTextFields(block), range = ranges?.get(block.id)
    if (ranges && !range) continue
    for (const field of fields) {
      if (field.field !== 'DATE' && field.field !== 'TIME' || documentFieldAutomaticRefreshBlocked(field)
        || documentTextFieldContainsNestedRange(field, fields) || bodyFieldContainsInlineEntity(block, field)
        || range && !documentTextFieldIntersects(field, range.start, range.end)
        || !bodyFieldContentEditable(model, block, field)) continue
      if (!resolveDocumentPermission(model.permissionRules ?? [], context, 'edit', [...permissionTargetsForBlock(model, block.id), { scope: 'entity', id: field.id }]).allowed) continue
      result.push({ blockId: block.id, field })
    }
  }
  return result
}

export function canRefreshSelectedDocumentBodyFields(model: DocumentModel, selection: DocumentSelection, context: DocumentPermissionContext): boolean {
  return editableBodyDateFields(model, context, selection).length > 0
}

function bodyDateLanguage(block: DocumentBlock, field: DocumentTextField, fallback: string): string {
  let offset = 0
  const first = block.runs.find(run => {
    const from = offset; offset += run.text.length
    return from <= field.startOffset && offset > field.startOffset
  })
  if (first?.metadata?.officeStoryTextFormatCleared === true || field.metadata?.officeStoryTextFormatCleared === true) return fallback
  const source = first?.metadata?.source, fieldSource = field.metadata?.source
  return sourceLanguage(isRecord(source) ? source.style : undefined) ?? field.language
    ?? sourceLanguage(isRecord(fieldSource) ? fieldSource.textStyle : undefined) ?? fallback
}

function mappedSelectionPoint(point: DocumentPoint, blockId: string, start: number, end: number, length: number): DocumentPoint {
  if ((point.cellBlockId ?? point.blockId) !== blockId || point.offset <= start) return point
  return { ...point, offset: point.offset >= end ? point.offset + length - (end - start) : start + Math.min(point.offset - start, length) }
}

function refreshedBodyRuns(block: DocumentBlock, field: DocumentTextField, result: string, model: DocumentModel): DocumentRun[] {
  const runs = block.runs as DocumentRun[], defaults = documentDefaultStyle(model).text
  const { before, after } = splitDocumentRunRange(runs, field.startOffset, field.endOffset)
  const cache = sliceRuns(runs, field.startOffset, field.endOffset)
  const styled = cache.length ? cache : [{ ...(documentRunStyleAt(runs, field.startOffset, field.endOffset) ?? defaults), id: `${field.id}:result`, text: '' } as DocumentRun]
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
  if (request.bodySelection !== undefined) bodySelectionRanges(model, request.bodySelection)
  const currentSelection = selection ?? request.bodySelection
  if (currentSelection) bodySelectionRanges(model, currentSelection)
  const selectedAreas = request.stories !== undefined || request.bodySelection !== undefined
  const bodyFields = !selectedAreas || request.bodySelection ? editableBodyDateFields(model, context, request.bodySelection) : []
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
  const bodyBlocks = new Map(documentTextFieldBlocks(candidate).map(block => [block.id, block]))
  for (const { blockId, field: sourceField } of [...bodyFields].sort((a, b) => a.blockId.localeCompare(b.blockId) || b.field.startOffset - a.field.startOffset)) {
    const block = bodyBlocks.get(blockId)!, field = documentTextFields(block).find(field => field.id === sourceField.id)!
    const result = resolveDocumentDateTimeFieldResult({ ...field, field: field.field as 'DATE' | 'TIME' }, bodyDateLanguage(block, field, request.locale), date)
    if (result === undefined || result === field.cachedResult && runText(block.runs).slice(field.startOffset, field.endOffset) === result) continue
    const nextRuns = refreshedBodyRuns(block, field, result, model)
    transformContentControlsForBlockEdit(candidate, block.id, field.startOffset, field.endOffset, result.length, undefined, true)
    transformDocumentTextFieldsForBlockEdit(block, field.startOffset, field.endOffset, result, field.id)
    block.runs = nextRuns
    setDocumentTextFields(block, documentTextFields(block).map(current => current.id === field.id ? { ...current, cachedResult: result } : current))
    delete block.metadata.html
    transformDocumentCommentAnchorsForTextEdit(candidate, block.id, field.startOffset, field.endOffset, result)
    if (nextSelection) nextSelection = {
      anchor: mappedSelectionPoint(nextSelection.anchor, block.id, field.startOffset, field.endOffset, result.length),
      focus: mappedSelectionPoint(nextSelection.focus, block.id, field.startOffset, field.endOffset, result.length),
    }
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
