import type { DocumentPoint } from '../model/document-model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentLayoutForBlockHost } from './document-auxiliary-layout'
import { hitTestDocument } from './document-layout'
import { intersectTableRect } from './document-table-measurement'
import type { DocumentLayout, LayoutNoteGlyphAddress, LayoutRect } from './document-layout-types'

export interface DocumentNoteGlyphHit extends LayoutNoteGlyphAddress { page: number; rect: LayoutRect }
export interface DocumentNoteRegionHit { host: Extract<DocumentAuxiliaryBlockHost, { kind: 'note' }>; page: number; point?: DocumentPoint }

const contains = (rect: LayoutRect, x: number, y: number) => rect.width > 0 && rect.height > 0
  && x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height

/** Only actual visible glyphs are navigation targets; decoration stories are excluded. */
export function hitTestDocumentNoteGlyph(layout: DocumentLayout, x: number, y: number): DocumentNoteGlyphHit | undefined {
  const page = layout.pages.find(page => !page.isParityBlank && contains(page, x, y))
  if (!page) return undefined
  const lines = [...(layout.pageContent[page.index]?.lines ?? []),
    ...(layout.auxiliaryContent ?? []).filter(content => content.page === page.index).flatMap(content => content.lines),
    ...(layout.pageHeaderFooters[page.index]?.headerLayoutLines ?? []), ...(layout.pageHeaderFooters[page.index]?.footerLayoutLines ?? [])]
  for (const line of lines) for (const glyph of line.glyphs) {
    if (!glyph.note) continue
    const clip = 'clip' in line && line.clip ? intersectTableRect(line.clip, page) : page
    const rect = intersectTableRect(glyph, clip)
    if (contains(rect, x, y)) return { ...glyph.note, page: page.index, rect }
  }
  return undefined
}

/** Route a real note column before ordinary body hit-testing can choose a nearby line. */
export function hitTestDocumentNoteRegion(layout: DocumentLayout, x: number, y: number): DocumentNoteRegionHit | undefined {
  for (const content of layout.auxiliaryContent ?? []) {
    if (content.host.kind !== 'note') continue
    const page = layout.pages[content.page]
    if (!page || page.isParityBlank) continue
    const column = content.columns?.[0], clip = column ? intersectTableRect(column, page.content) : page.content
    if (!contains(intersectTableRect(content, clip), x, y)) continue
    return { host: { ...content.host }, page: content.page,
      point: hitTestDocument(documentLayoutForBlockHost(layout, content.host), x, y) }
  }
  return undefined
}

export function hitTestDocumentNoteDecoration(layout: DocumentLayout, x: number, y: number): boolean {
  return (layout.noteDecorations ?? []).some(content => {
    const page = layout.pages[content.page]
    return !!page && !page.isParityBlank && contains(intersectTableRect(content, page.content), x, y)
  })
}
