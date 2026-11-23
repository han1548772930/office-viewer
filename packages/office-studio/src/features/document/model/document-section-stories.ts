import type { DocumentSectionHeaderFooter } from '../../../domain/model'
import { DOCUMENT_STORY_KEYS, type DocumentStoryKey } from './document-story'

export type DocumentSectionStoryLinks = Record<DocumentStoryKey, boolean>

export function validateSectionStoryLinkMap(value: unknown): asserts value is Partial<DocumentSectionStoryLinks> | undefined {
  if (value === undefined) return
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('Invalid section story associations')
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !DOCUMENT_STORY_KEYS.some(story => story === key)) throw new TypeError('Unknown section story association key')
    if (typeof Reflect.get(value, key) !== 'boolean') throw new TypeError(`Invalid section story association: ${key}`)
  }
}

export function sectionStoryLinkedToPrevious(section: DocumentSectionHeaderFooter, story: DocumentStoryKey): boolean {
  return section.storyLinkToPrevious && Object.hasOwn(section.storyLinkToPrevious, story)
    ? section.storyLinkToPrevious[story]!
    : section.linkToPrevious
}

/** Complete the legacy fallback before computing a mixed section's aggregate flag. */
export function sectionStoryLinks(section: DocumentSectionHeaderFooter): DocumentSectionStoryLinks {
  if (typeof section.linkToPrevious !== 'boolean') throw new TypeError('Invalid section linkToPrevious')
  validateSectionStoryLinkMap(section.storyLinkToPrevious)
  return {
    header: sectionStoryLinkedToPrevious(section, 'header'), footer: sectionStoryLinkedToPrevious(section, 'footer'),
    firstPageHeader: sectionStoryLinkedToPrevious(section, 'firstPageHeader'), firstPageFooter: sectionStoryLinkedToPrevious(section, 'firstPageFooter'),
    evenPageHeader: sectionStoryLinkedToPrevious(section, 'evenPageHeader'), evenPageFooter: sectionStoryLinkedToPrevious(section, 'evenPageFooter'),
  }
}

/** Replace one story in a derived draft. Its caller owns cloning before publication. */
export function replaceSectionStory(target: DocumentSectionHeaderFooter, source: DocumentSectionHeaderFooter, story: DocumentStoryKey): void {
  target[story] = source[story]
  const content = { ...target.storyContent }, replacement = source.storyContent?.[story]
  if (replacement === undefined) delete content[story]
  else content[story] = replacement
  if (Object.keys(content).length) target.storyContent = content
  else delete target.storyContent

  const fields = target.pageFields, inherited = source.pageFields?.filter(field => field.story === story) ?? []
  if (fields === undefined && !inherited.length) return
  const first = fields?.findIndex(field => field.story === story) ?? -1
  const remaining = fields?.filter(field => field.story !== story) ?? []
  const at = first < 0 ? remaining.length : first
  target.pageFields = [...remaining.slice(0, at), ...inherited, ...remaining.slice(at)]
}

export function sameSectionStory(first: DocumentSectionHeaderFooter, second: DocumentSectionHeaderFooter, story: DocumentStoryKey): boolean {
  const snapshot = (section: DocumentSectionHeaderFooter) => ({ text: section[story], content: section.storyContent?.[story],
    fields: section.pageFields?.filter(field => field.story === story) ?? [] })
  const stable = (value: unknown) => JSON.stringify(value, (_key, item: unknown) => item && typeof item === 'object' && !Array.isArray(item)
    ? Object.fromEntries(Object.entries(item).sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0)) : item)
  return stable(snapshot(first)) === stable(snapshot(second))
}
