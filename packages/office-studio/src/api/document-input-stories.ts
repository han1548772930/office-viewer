import type { DocumentModel, DocumentSectionHeaderFooter } from '../domain/model'
import { DOCUMENT_STORY_KEYS } from '../features/document/model/document-story'
import { projectDocumentStoryBlocks } from '../features/document/model/document-story-blocks'
import { EMPTY_SECTION_HEADER_FOOTER, DEFAULT_DOCUMENT_SECTION_ID } from '../features/document/services/section-header-footer'
import { appendDocumentNoteContentControls } from '../features/document/notes/document-note-content'
import { decodeDocumentBody } from './document-body-codec'
import { bodyBoolean, bodyRecord, type UniverBodyRecord } from './document-body-style'
import type { DocumentBodyDecodeOptions, UniverDocumentBodyResources } from './document-body-types'
import { claimPublicDocumentBody, publicDocumentRecord, publicDocumentResources } from './document-input-common'

export function publicDocumentSectionSources(model: DocumentModel): Array<{ localId: string; source: UniverBodyRecord }> {
  return [{ localId: DEFAULT_DOCUMENT_SECTION_ID, source: bodyRecord(model.metadata?.univerBodyFirstSectionSource) },
    ...model.blocks.filter(block => block.type === 'sectionBreak').map(block => ({ localId: block.id, source: bodyRecord(block.metadata.univerBodySectionSource) }))]
}

const storyReferences = [
  ['header', 'defaultHeaderId', 'headers', 'headerId'], ['footer', 'defaultFooterId', 'footers', 'footerId'],
  ['firstPageHeader', 'firstPageHeaderId', 'headers', 'headerId'], ['firstPageFooter', 'firstPageFooterId', 'footers', 'footerId'],
  ['evenPageHeader', 'evenPageHeaderId', 'headers', 'headerId'], ['evenPageFooter', 'evenPageFooterId', 'footers', 'footerId'],
] as const

/** Public segments enter the existing section-story model, including fields,
 * table blocks and SDTs; IDs repeated by shared source parts are claimed once. */
export function applyPublicDocumentStories(model: DocumentModel, input: UniverBodyRecord, resources: UniverDocumentBodyResources,
  options: DocumentBodyDecodeOptions, claimed: Set<string>): void {
  const headers = input.headers === undefined ? {} : publicDocumentRecord(input.headers, 'headers')
  const footers = input.footers === undefined ? {} : publicDocumentRecord(input.footers, 'footers')
  for (const [kind, entries, identity] of [['headers', headers, 'headerId'], ['footers', footers, 'footerId']] as const) {
    for (const [id, value] of Object.entries(entries)) {
      const raw = publicDocumentRecord(value, `${kind}.${id}`)
      if (raw[identity] !== undefined && raw[identity] !== id) throw new TypeError(`页眉页脚资源身份不匹配：${id}`)
      publicDocumentRecord(raw.body, `${kind}.${id}.body`)
    }
  }
  const documentStyle = bodyRecord(input.documentStyle), previous = new Map<string, string | undefined>()
  const sections: NonNullable<DocumentModel['sectionHeaderFooters']> = {}
  for (const [index, binding] of publicDocumentSectionSources(model).entries()) {
    const source = binding.source, section: DocumentSectionHeaderFooter = { ...EMPTY_SECTION_HEADER_FOOTER, storyContent: {}, pageFields: [], storyLinkToPrevious: {} }
    section.differentFirstPage = bodyBoolean(source.useFirstPageHeaderFooter ?? documentStyle.useFirstPageHeaderFooter) ?? false
    section.differentOddEven = bodyBoolean(source.evenAndOddHeaders ?? documentStyle.evenAndOddHeaders) ?? false
    for (const [story, reference, kind] of storyReferences) {
      const value = source[reference] ?? (index === 0 ? documentStyle[reference] : previous.get(story))
      if (value !== undefined && (typeof value !== 'string' || !value.trim())) throw new TypeError(`页眉页脚 ${reference} 标识无效。`)
      const id = typeof value === 'string' ? value : undefined
      const linked = index > 0 && id === previous.get(story)
      section.storyLinkToPrevious![story] = linked
      previous.set(story, id)
      if (linked || id === undefined) continue
      const raw = publicDocumentRecord((kind === 'headers' ? headers : footers)[id], `页眉页脚资源 ${id}`)
      const decoded = claimPublicDocumentBody(decodeDocumentBody(raw.body, publicDocumentResources(raw, resources), { ...options, allowSections: false }), claimed)
      decoded.content.metadata = { ...decoded.content.metadata, univerHeaderFooterSource: { version: 1, id, kind } }
      const projection = projectDocumentStoryBlocks(decoded.content)
      section[story] = projection.text; section.storyContent![story] = projection.content
      section.pageFields!.push(...projection.pageFields.map(field => ({ ...field, story })))
      appendDocumentNoteContentControls(model, decoded.contentControls)
    }
    section.linkToPrevious = DOCUMENT_STORY_KEYS.every(story => section.storyLinkToPrevious![story])
    if (!Object.keys(section.storyContent!).length) delete section.storyContent
    if (!section.pageFields!.length) delete section.pageFields
    sections[binding.localId] = section
  }
  model.sectionHeaderFooters = sections
}
