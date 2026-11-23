import type { DocumentStoryKey } from '../model/document-story'
import type { DocumentLayout } from './document-layout-types'

export interface DocumentStoryPageRequest {
  area: 'header' | 'footer'
  preferredPage?: number
  sectionId?: string
  story?: DocumentStoryKey
}

/** Physical parity filler pages do not own an editable header or footer story. */
export function findDocumentStoryPage(layout: Pick<DocumentLayout, 'pages' | 'pageHeaderFooters'>, request: DocumentStoryPageRequest): number | undefined {
  const matches = (index: number) => {
    const page = layout.pages[index], entry = layout.pageHeaderFooters[index]
    return !!page && !!entry && !page.isParityBlank && !entry.isParityBlank
      && (request.sectionId === undefined || entry.sectionId === request.sectionId)
      && (request.story === undefined || (request.area === 'header' ? entry.headerStory : entry.footerStory) === request.story)
  }
  const preferred = request.preferredPage
  if (preferred !== undefined && matches(preferred)) return preferred
  let result: number | undefined, distance = Infinity
  for (let index = 0; index < layout.pages.length; index++) {
    if (!matches(index)) continue
    const next = preferred === undefined ? index : Math.abs(index - preferred)
    if (next < distance) { result = index; distance = next }
  }
  return result
}
