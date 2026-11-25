import type { DocumentModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentStoryDefaultStyle } from '../features/document/model/document-story'
import { documentSectionIds, resolveSectionHeaderFooter, sectionStoryLinkedToPrevious } from '../features/document/services/section-header-footer'
import { resolveDocumentBlockHost } from '../features/document/services/document-block-host'
import { documentBodyContentControls, encodeDocumentBody } from './document-body-codec'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'
import type { UniverDocumentBodyData } from './document-body-types'

const references = [
  ['header', 'defaultHeaderId', 'headers', 'headerId'], ['footer', 'defaultFooterId', 'footers', 'footerId'],
  ['firstPageHeader', 'firstPageHeaderId', 'headers', 'headerId'], ['firstPageFooter', 'firstPageFooterId', 'footers', 'footerId'],
  ['evenPageHeader', 'evenPageHeaderId', 'headers', 'headerId'], ['evenPageFooter', 'evenPageFooterId', 'footers', 'footerId'],
] as const
type StoryMap = Record<string, UniverDocumentBodyData & UniverBodyRecord>

/** Writes each live outgoing section's segment references. Linked sections
 * reuse the source resource; independent overrides receive separate IDs. */
export function writePublicDocumentStories(model: DocumentModel, source: UniverBodyRecord, sections: UniverBodyRecord[]): {
  headers: StoryMap; footers: StoryMap
} {
  const headers: StoryMap = {}, footers: StoryMap = {}, previous = new Map<string, string | undefined>()
  const used = { headers: new Map<string, string>(), footers: new Map<string, string>() }
  const written = { headers: new Set<string>(), footers: new Set<string>() }
  const sourceMaps = { headers: bodyRecord(source.headers), footers: bodyRecord(source.footers) }
  const modeled = { headers: new Set<string>(), footers: new Set<string>() }
  const oldSections = bodyRecord(bodyRecord(model.metadata?.univerBodySource).body).sectionBreaks
  for (const section of [bodyRecord(source.documentStyle), ...(Array.isArray(oldSections) ? oldSections.map(bodyRecord) : [])]) {
    for (const [, reference, kind] of references) if (typeof section[reference] === 'string') modeled[kind].add(section[reference])
  }
  for (const kind of ['headers', 'footers'] as const) for (const [id, value] of Object.entries(sourceMaps[kind])) {
    if (modeled[kind].has(id)) continue
    const raw = bodyRecord(value)
    if (typeof bodyRecord(raw.body).dataStream !== 'string') throw new TypeError('保留的页眉页脚缺少公开正文：' + id)
    Object.assign(kind === 'headers' ? headers : footers, { [id]: cloneModel(raw) })
  }
  const ids = documentSectionIds(model)
  if (ids.length !== sections.length) throw new RangeError('公开分节描述与当前文档节数量不一致。')
  for (const [index, sectionId] of ids.entries()) {
    const descriptor = sections[index]!, section = resolveSectionHeaderFooter(model, sectionId)
    descriptor.useFirstPageHeaderFooter = Number(section.differentFirstPage)
    descriptor.evenAndOddHeaders = Number(section.differentOddEven)
    for (const [story, reference, kind, identity] of references) {
      delete descriptor[reference]
      if (index > 0 && sectionStoryLinkedToPrevious(section, story)) {
        const id = previous.get(story)
        if (id !== undefined) descriptor[reference] = id
        continue
      }
      const present = section.storyContent?.[story] !== undefined || section[story].length > 0
        || section.pageFields?.some(field => field.story === story)
      if (!present) { previous.set(story, undefined); continue }
      const view = resolveDocumentBlockHost(model, { kind: 'story', sectionId, story })
      const data = encodeDocumentBody(view.content, documentStoryDefaultStyle(model), documentBodyContentControls(view.content, model.contentControls))
      const provenance = bodyRecord(view.content.metadata?.univerHeaderFooterSource)
      const preferred = typeof provenance.id === 'string' && provenance.id ? provenance.id : `office:${identity}:${sectionId}:${story}`
      const signature = JSON.stringify(data), existing = used[kind].get(preferred)
      let id = preferred
      if (existing !== undefined && existing !== signature || (kind === 'headers' ? headers : footers)[id] && !written[kind].has(id)) {
        id = `office:${identity}:${sectionId}:${story}`
      }
      if (used[kind].has(id) && used[kind].get(id) !== signature) throw new RangeError('独立页眉页脚资源身份冲突：' + id)
      used[kind].set(id, signature); written[kind].add(id)
      const raw = bodyRecord(sourceMaps[kind][preferred])
      Object.assign(kind === 'headers' ? headers : footers, { [id]: { ...cloneModel(raw), ...data,
        tableSource: data.tableSource, drawings: data.drawings, drawingsOrder: data.drawingsOrder, lists: data.lists, [identity]: id } })
      previous.set(story, id); descriptor[reference] = id
    }
  }
  return { headers, footers }
}
