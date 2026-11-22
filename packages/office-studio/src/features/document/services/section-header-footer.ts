import type { DocumentModel, DocumentPageField, DocumentSectionHeaderFooter } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentDefaultStyle } from '../model/document-default-style'
import { DOCUMENT_STORY_KEYS, documentStoryDefaultStyle, documentStoryFieldCache, documentStoryTextEdit, editDocumentStoryLocks, editDocumentStoryParagraphs, isDocumentTextBoundary, validateDocumentStoryContent, validateDocumentStoryLocks } from '../model/document-story'
import { replaceDocumentRunRange } from '../model/document-run-range'
import { storyBuffer } from './story-editor'
import { assertDocumentStoryLocksPreserved } from './document-story-locks'

export const DEFAULT_DOCUMENT_SECTION_ID = 'first'

export const EMPTY_SECTION_HEADER_FOOTER: DocumentSectionHeaderFooter = {
  header: '',
  footer: '',
  firstPageHeader: '',
  firstPageFooter: '',
  evenPageHeader: '',
  evenPageFooter: '',
  differentFirstPage: false,
  differentOddEven: false,
  linkToPrevious: false,
}

export interface SectionHeaderFooterPatch extends Partial<DocumentSectionHeaderFooter> {}

export function documentSectionIds(model: DocumentModel): string[] {
  return [DEFAULT_DOCUMENT_SECTION_ID, ...model.blocks.filter(block => block.type === 'sectionBreak').map(block => block.id)]
}

export function applySectionHeaderFooter(
  model: DocumentModel,
  sectionId: string,
  patch: SectionHeaderFooterPatch,
): void {
  const sectionIds = documentSectionIds(model)
  const sectionIndex = sectionIds.indexOf(sectionId)
  if (sectionIndex < 0) throw new RangeError(`Unknown document section: ${sectionId}`)
  const current = resolveSectionHeaderFooter(model, sectionId)
  const next = { ...current, ...patch, linkToPrevious: patch.linkToPrevious ?? current.linkToPrevious }
  for (const key of ['header', 'footer', 'firstPageHeader', 'firstPageFooter', 'evenPageHeader', 'evenPageFooter'] as const) {
    if (typeof next[key] !== 'string') throw new TypeError(`Invalid section ${key}`)
  }
  for (const key of ['differentFirstPage', 'differentOddEven', 'linkToPrevious'] as const) {
    if (typeof next[key] !== 'boolean') throw new TypeError(`Invalid section ${key}`)
  }
  if (patch.storyContent !== undefined && (!patch.storyContent || typeof patch.storyContent !== 'object' || Array.isArray(patch.storyContent))) throw new TypeError('Invalid section story formatting')
  if (Object.keys(patch.storyContent ?? {}).some(key => !DOCUMENT_STORY_KEYS.some(story => story === key))) throw new TypeError('Unknown section story formatting key')
  next.storyContent = { ...current.storyContent, ...patch.storyContent }
  for (const key of ['header', 'footer', 'firstPageHeader', 'firstPageFooter', 'evenPageHeader', 'evenPageFooter'] as const) {
    const content = next.storyContent[key]
    if (content !== undefined && next[key] !== current[key] && !Object.hasOwn(patch.storyContent ?? {}, key)) {
      let from = 0, to = current[key].length, last = next[key].length
      while (from < to && from < last && current[key][from] === next[key][from]) from++
      while (to > from && last > from && current[key][to - 1] === next[key][last - 1]) { to--; last-- }
      if (!isDocumentTextBoundary(current[key], from)) from--
      if (!isDocumentTextBoundary(current[key], to)) { to++; last++ }
      if (content.paragraphs.some(paragraph => paragraph.metadata?.contentLocked === true && from <= paragraph.endOffset && to >= paragraph.startOffset)) throw new Error('页眉页脚包含已锁定内容。')
      const insert = next[key].slice(from, last)
      next.storyContent[key] = {
        ...content,
        runs: replaceDocumentRunRange(content.runs, from, to, insert, documentStoryDefaultStyle(model).text),
        paragraphs: editDocumentStoryParagraphs(content.paragraphs, current[key], from, to, insert),
      }
    }
    const prepared = next.storyContent[key]
    if (prepared !== undefined) {
      validateDocumentStoryContent(prepared, next[key])
      next.storyContent[key] = { ...prepared, runs: prepared.runs.map(run => ({ ...run, ...(run.metadata ? { metadata: { ...run.metadata } } : {}) })),
        paragraphs: prepared.paragraphs.map(paragraph => ({ ...paragraph, style: { ...paragraph.style }, ...(paragraph.metadata ? { metadata: { ...paragraph.metadata } } : {}) })),
        ...(prepared.locks ? { locks: prepared.locks.map(lock => ({ ...lock })) } : {}), ...(prepared.metadata ? { metadata: { ...prepared.metadata } } : {}) }
    }
  }
  if (!Object.hasOwn(patch, 'pageFields')) {
    next.pageFields = next.pageFields?.map(field => {
      if (field.offset === undefined || next[field.story] === current[field.story]) return field
      const edit = documentStoryTextEdit(current[field.story], next[field.story])
      return { ...field, offset: field.offset >= edit.end ? field.offset + edit.text.length - (edit.end - edit.start)
        : field.offset > edit.start ? edit.start + edit.text.length : field.offset }
    })
  }
  if (!Object.keys(next.storyContent).length) delete next.storyContent
  if (next.pageFields !== undefined && (!Array.isArray(next.pageFields) || next.pageFields.some(field =>
    !['header', 'footer', 'firstPageHeader', 'firstPageFooter', 'evenPageHeader', 'evenPageFooter'].includes(field.story)
      || !['PAGE', 'NUMPAGES', 'DATE', 'TIME'].includes(field.field)
      || (field.offset !== undefined && !isDocumentTextBoundary(next[field.story], field.offset))
      || (field.instruction !== undefined && (typeof field.instruction !== 'string' || field.instruction.trim().split(/\s+/)[0]?.toUpperCase() !== field.field))
      || (field.cachedResult !== undefined && typeof field.cachedResult !== 'string')
      || (field.metadata !== undefined && (!field.metadata || typeof field.metadata !== 'object' || Array.isArray(field.metadata)))))) {
    throw new TypeError('Invalid section page fields')
  }
  next.pageFields = next.pageFields?.map(field => {
    if (field.textStyle !== undefined) documentDefaultStyle({ defaultStyle: { version: 1, text: field.textStyle, paragraph: documentStoryDefaultStyle(model).paragraph } })
    documentStoryFieldCache(field)
    return { ...field, ...(field.textStyle ? { textStyle: { ...field.textStyle } } : {}), ...(field.metadata ? { metadata: cloneModel(field.metadata) } : {}) }
  })
  for (const key of DOCUMENT_STORY_KEYS) {
    const content = next.storyContent?.[key]
    const previousContent = current.storyContent?.[key]
    if (content && previousContent?.locks && !Object.hasOwn(patch.storyContent ?? {}, key)) {
      const before = storyBuffer(current, key).value, after = storyBuffer(next, key).value
      const edit = documentStoryTextEdit(before, after)
      content.locks = editDocumentStoryLocks(previousContent.locks, before, edit.start, edit.end, edit.text.length)
    }
    if (content) validateDocumentStoryLocks(content.locks, storyBuffer(next, key).value)
  }
  if (sectionIndex > 0 && next.linkToPrevious) {
    const previous = resolveSectionHeaderFooter(model, sectionIds[sectionIndex - 1]!)
    const linked = { ...previous, differentFirstPage: next.differentFirstPage, differentOddEven: next.differentOddEven, linkToPrevious: true }
    for (const key of DOCUMENT_STORY_KEYS) assertDocumentStoryLocksPreserved(current, linked, key, model)
    model.sectionHeaderFooters = { ...model.sectionHeaderFooters, [sectionId]: linked }
  } else {
    for (const key of DOCUMENT_STORY_KEYS) assertDocumentStoryLocksPreserved(current, next, key, model)
    model.sectionHeaderFooters = { ...model.sectionHeaderFooters, [sectionId]: next }
  }
}

export function insertSectionPageField(
  model: DocumentModel,
  sectionId: string,
  story: DocumentPageField['story'],
  field: DocumentPageField['field'],
  details: Pick<DocumentPageField, 'offset' | 'cachedResult' | 'instruction'> = {},
): void {
  const current = resolveSectionHeaderFooter(model, sectionId)
  if (current.linkToPrevious) throw new Error('Unlink the section before editing its fields')
  const { offset, cachedResult, instruction } = details
  applySectionHeaderFooter(model, sectionId, { pageFields: [...(current.pageFields ?? []), { story, field, offset, cachedResult, instruction }] })
}

export function renderSectionStory(section: DocumentSectionHeaderFooter, story: DocumentPageField['story'], page: number, total: number): string {
  const text = section[story]
  const fields = (section.pageFields ?? []).filter(field => field.story === story)
    .sort((a, b) => (a.offset ?? Infinity) - (b.offset ?? Infinity))
  let result = '', cursor = 0
  for (const field of fields) {
    const offset = Math.min(text.length, field.offset ?? text.length)
    result += text.slice(cursor, offset) + (field.offset === undefined ? ' ' : '')
    result += field.field === 'PAGE' ? page : field.field === 'NUMPAGES' ? total : field.cachedResult ?? ''
    cursor = offset
  }
  return result + text.slice(cursor)
}

export function resolveSectionHeaderFooter(model: DocumentModel, sectionId: string): DocumentSectionHeaderFooter {
  const sectionIds = documentSectionIds(model)
  const index = sectionIds.indexOf(sectionId)
  if (index < 0) return { ...EMPTY_SECTION_HEADER_FOOTER }
  const section = model.sectionHeaderFooters?.[sectionId]
  if (!section) {
    return index === 0
      ? { ...EMPTY_SECTION_HEADER_FOOTER }
      : { ...resolveSectionHeaderFooter(model, sectionIds[index - 1]!), linkToPrevious: true }
  }
  if (index > 0 && section.linkToPrevious) {
    return { ...resolveSectionHeaderFooter(model, sectionIds[index - 1]!), differentFirstPage: section.differentFirstPage, differentOddEven: section.differentOddEven, linkToPrevious: true }
  }
  return { ...EMPTY_SECTION_HEADER_FOOTER, ...section }
}

export function sectionIdForBlock(model: DocumentModel, blockId: string): string {
  let sectionId = DEFAULT_DOCUMENT_SECTION_ID
  for (const block of model.blocks) {
    if (block.id === blockId) return block.type === 'sectionBreak' ? block.id : sectionId
    if (block.type === 'sectionBreak') sectionId = block.id
  }
  return sectionId
}
