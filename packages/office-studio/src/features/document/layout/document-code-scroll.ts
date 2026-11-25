import type { DocumentPoint } from '../model/document-model'
import { documentBlockHostId, sameDocumentBlockHost } from '../model/document-block-host'
import { lineForPoint } from './document-layout'
import type { DocumentLayout, LayoutAuxiliaryContent, LayoutContainerDecoration, LayoutLine, LayoutRect } from './document-layout-types'

export type DocumentCodeScrollOffsets = Readonly<Record<string, number>>
export interface DocumentCodeScrollRegion {
  key: string
  container: LayoutContainerDecoration
  clip: LayoutRect
  maxScroll: number
}

export function documentCodeScrollKey(container: Pick<LayoutContainerDecoration, 'blockId' | 'host'>): string {
  return JSON.stringify([container.host ? `${container.host.kind}:${documentBlockHostId(container.host)}` : 'body', container.blockId])
}

function intersect(left: LayoutRect, right: LayoutRect): LayoutRect {
  const x = Math.max(left.x, right.x), y = Math.max(left.y, right.y)
  return { x, y, width: Math.max(0, Math.min(left.x + left.width, right.x + right.width) - x),
    height: Math.max(0, Math.min(left.y + left.height, right.y + right.height) - y) }
}
const allLines = (layout: DocumentLayout) => [...new Set([...layout.lines, ...layout.pageContent.flatMap(content => content.lines),
  ...(layout.auxiliaryContent ?? []).flatMap(content => content.lines), ...(layout.noteDecorations ?? []).flatMap(content => content.lines)])]
const allContainers = (layout: DocumentLayout) => [...new Set([...(layout.containers ?? []), ...layout.pageContent.flatMap(content => content.containers ?? []),
  ...(layout.auxiliaryContent ?? []).flatMap(content => content.containers ?? []), ...(layout.noteDecorations ?? []).flatMap(content => content.containers ?? [])])]

function ownsLine(container: LayoutContainerDecoration, line: LayoutLine): boolean {
  return container.blockId === line.codeBlock?.containerId && container.page === line.page && sameDocumentBlockHost(container.host, line.host)
    && line.y + line.height > container.y && line.y < container.y + container.height
    && line.x >= container.x - 1e-6 && line.x <= container.x + container.width + 1e-6
}

/** Fragment clips use physical geometry; one view offset is shared by the block's fragments. */
export function documentCodeScrollRegions(layout: DocumentLayout): DocumentCodeScrollRegion[] {
  const lines = allLines(layout), regions: DocumentCodeScrollRegion[] = []
  for (const container of allContainers(layout).filter(container => container.kind === 'code')) {
    const owned = lines.filter(line => ownsLine(container, line))
    if (!owned.length) continue
    const leftInset = owned[0]!.codeBlock?.showLineNumbers ? 44 : 12
    const inset = Math.min(leftInset, Math.max(0, container.width - 1))
    const content: LayoutRect = { x: container.x + inset, y: container.y,
      width: Math.max(1, container.width - inset - 12), height: container.height }
    const clip = container.clip ? intersect(content, container.clip) : content
    const right = Math.max(clip.x, ...owned.map(line => Math.max(line.x + line.width,
      ...line.glyphs.map(glyph => glyph.x + glyph.width))))
    regions.push({ key: documentCodeScrollKey(container), container, clip,
      maxScroll: container.codeWrap === false ? Math.max(0, right + 1.3 - clip.x - clip.width) : 0 })
  }
  const maxima = new Map<string, number>()
  for (const region of regions) maxima.set(region.key, Math.max(maxima.get(region.key) ?? 0, region.maxScroll))
  return regions.map(region => ({ ...region, maxScroll: maxima.get(region.key)! }))
}

export function clampDocumentCodeScrollOffsets(layout: DocumentLayout, offsets: DocumentCodeScrollOffsets): Record<string, number> {
  const result: Record<string, number> = {}
  for (const region of documentCodeScrollRegions(layout)) {
    const value = offsets[region.key]
    if (typeof value === 'number' && Number.isFinite(value)) result[region.key] = Math.min(region.maxScroll, Math.max(0, value))
  }
  return result
}

/** The authoritative layout stays unchanged. Render, hit testing and input share this projection. */
export function projectDocumentCodeScroll(layout: DocumentLayout, offsets: DocumentCodeScrollOffsets): DocumentLayout {
  const regions = documentCodeScrollRegions(layout)
  if (!regions.length) return layout
  const projected = new Map<LayoutLine, LayoutLine>()
  const line = (source: LayoutLine): LayoutLine => {
    const existing = projected.get(source)
    if (existing) return existing
    const region = regions.filter(region => ownsLine(region.container, source))
      .sort((left, right) => left.container.width * left.container.height - right.container.width * right.container.height)[0]
    if (!region) return source
    const requested = offsets[region.key] ?? 0, offset = Number.isFinite(requested) ? Math.max(0, Math.min(requested, region.maxScroll)) : 0
    const clip = source.clip ? intersect(region.clip, source.clip) : region.clip
    const gutterClip = source.clip ? intersect(region.container, source.clip) : { ...region.container }
    const result = { ...source, x: source.x - offset, clip,
      ...(source.markerStyle ? { markerStyle: { ...source.markerStyle, x: source.markerStyle.x - offset } } : {}),
      codeBlock: { ...source.codeBlock!, gutterX: source.x, gutterClip },
      glyphs: source.glyphs.map(glyph => ({ ...glyph, x: glyph.x - offset })) }
    projected.set(source, result)
    return result
  }
  const auxiliary = (content: LayoutAuxiliaryContent): LayoutAuxiliaryContent => ({ ...content, lines: content.lines.map(line) })
  return { ...layout, lines: layout.lines.map(line), pageContent: layout.pageContent.map(content => ({ ...content, lines: content.lines.map(line) })),
    ...(layout.auxiliaryContent ? { auxiliaryContent: layout.auxiliaryContent.map(auxiliary) } : {}),
    ...(layout.noteDecorations ? { noteDecorations: layout.noteDecorations.map(content => ({ ...content, lines: content.lines.map(line) })) } : {}),
    pageHeaderFooters: layout.pageHeaderFooters.map(content => ({ ...content,
      ...(content.headerBlockLayout ? { headerBlockLayout: auxiliary(content.headerBlockLayout) } : {}),
      ...(content.footerBlockLayout ? { footerBlockLayout: auxiliary(content.footerBlockLayout) } : {}) })) }
}

/** Use the raw glyph position: caretRect deliberately clamps the visible caret to its clip. */
export function documentCodeCaretScrollOffset(layout: DocumentLayout, point: DocumentPoint, offsets: DocumentCodeScrollOffsets): { key: string; offset: number } | undefined {
  const line = lineForPoint(layout, point)
  if (!line?.codeBlock?.containerId) return undefined
  const region = documentCodeScrollRegions(layout).find(region => ownsLine(region.container, line))
  if (!region || region.clip.width <= 0) return undefined
  const glyphs = line.glyphs.filter(glyph => glyph.note?.role !== 'bodyLabel'), next = glyphs.find(glyph => point.offset <= glyph.offset), last = glyphs.at(-1)
  const x = next?.x ?? (last ? last.x + last.width : line.x), current = offsets[region.key] ?? 0
  const inset = Math.min(8, region.clip.width / 4), left = region.clip.x + inset, right = region.clip.x + region.clip.width - inset - 1.3
  const desired = x - current < left ? x - left : x - current > right ? x - right : current
  return { key: region.key, offset: Math.max(0, Math.min(region.maxScroll, desired)) }
}
