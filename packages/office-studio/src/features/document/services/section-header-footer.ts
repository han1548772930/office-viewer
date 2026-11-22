import type { DocumentModel, DocumentPageField, DocumentSectionHeaderFooter } from '../../../domain/model'

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
  if (next.pageFields !== undefined && (!Array.isArray(next.pageFields) || next.pageFields.some(field =>
    !['header', 'footer', 'firstPageHeader', 'firstPageFooter', 'evenPageHeader', 'evenPageFooter'].includes(field.story)
      || !['PAGE', 'NUMPAGES', 'DATE', 'TIME'].includes(field.field)
      || (field.offset !== undefined && (!Number.isSafeInteger(field.offset) || field.offset < 0))
      || (field.instruction !== undefined && field.instruction.trim().split(/\s+/)[0]?.toUpperCase() !== field.field)
      || (field.cachedResult !== undefined && typeof field.cachedResult !== 'string')))) {
    throw new TypeError('Invalid section page fields')
  }
  next.pageFields = next.pageFields?.map(field => ({ ...field }))
  if (sectionIndex > 0 && next.linkToPrevious) {
    const previous = resolveSectionHeaderFooter(model, sectionIds[sectionIndex - 1]!)
    model.sectionHeaderFooters = { ...model.sectionHeaderFooters, [sectionId]: { ...previous, differentFirstPage: next.differentFirstPage, differentOddEven: next.differentOddEven, linkToPrevious: true } }
  } else {
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
