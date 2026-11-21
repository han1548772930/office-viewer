import type { DocumentModel, DocumentSectionHeaderFooter } from '../../../domain/model'

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
    next[key] = next[key].slice(0, 500)
  }
  for (const key of ['differentFirstPage', 'differentOddEven', 'linkToPrevious'] as const) {
    if (typeof next[key] !== 'boolean') throw new TypeError(`Invalid section ${key}`)
  }
  if (sectionIndex > 0 && next.linkToPrevious) {
    const previous = resolveSectionHeaderFooter(model, sectionIds[sectionIndex - 1]!)
    model.sectionHeaderFooters = { ...model.sectionHeaderFooters, [sectionId]: { ...previous, linkToPrevious: true } }
  } else {
    model.sectionHeaderFooters = { ...model.sectionHeaderFooters, [sectionId]: next }
  }
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
    return { ...resolveSectionHeaderFooter(model, sectionIds[index - 1]!), linkToPrevious: true }
  }
  return { ...EMPTY_SECTION_HEADER_FOOTER, ...section }
}

export function sectionIdForBlock(model: DocumentModel, blockId: string): string {
  let sectionId = DEFAULT_DOCUMENT_SECTION_ID
  for (const block of model.blocks) {
    if (block.id === blockId) return sectionId
    if (block.type === 'sectionBreak') sectionId = block.id
  }
  return sectionId
}
