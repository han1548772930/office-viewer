import type { DocumentModel } from '../../../domain/model'
import type { DocumentPoint } from '../model/document-model'
import { createRun, sameTarget, targetKey } from '../model/document-model'
import { sameDocumentBlockHost } from '../model/document-block-host'
import { layoutDocumentWithNotes } from './document-note-layout'
import { intersectTableRect } from './document-table-measurement'
import type { DocumentLayout, LayoutLine, LayoutMeasurer, LayoutRect, LayoutStoryLine } from './document-layout-types'

export type { DocumentLayout, LayoutCell, LayoutColumn, LayoutDivider, LayoutFlowLocation, LayoutGlyph, LayoutImage, LayoutLine, LayoutMeasurer, LayoutPage, LayoutPageContent, LayoutPageHeaderFooter, LayoutRect, LayoutRightTabStop, LayoutSection, LayoutStoryLine, LayoutTableFragment, LayoutTableRowMetric, LayoutTabGlyph, LayoutTabLeader } from './document-layout-types'
export { PAGE_GAP, PAGE_INSET, STORY_LINE_HEIGHT } from './document-layout-types'
export { createTextMeasurer, runFont } from './document-text-layout'

export function layoutDocument(model: DocumentModel, measurer: LayoutMeasurer): DocumentLayout {
  return layoutDocumentWithNotes(model, measurer)
}

export function lineForPoint(layout: DocumentLayout, point: DocumentPoint): LayoutLine | undefined {
  const all = layout.lines.filter((line) => sameTarget(line.target, point))
  const atOffset = (lines: LayoutLine[]) => lines.find((line, index) => point.offset >= line.start && point.offset <= line.end
    && (point.offset < line.end || index === lines.length - 1 || line.end !== lines[index + 1]?.start))
  const preferred = point.pageAffinity === undefined ? [] : all.filter(line => line.page === point.pageAffinity)
  const original = all.filter(line => !line.repeatedHeader)
  return atOffset(preferred) ?? atOffset(original) ?? atOffset(all) ?? original.at(-1) ?? all.at(-1)
}

/** Wrap fragments around a drawing are one physical row in their owning host. */
export function documentPhysicalLineFragments(layout: Pick<DocumentLayout, 'lines'>, line: LayoutLine): LayoutLine[] {
  if (!line.rowKey) return [line]
  return layout.lines.filter(candidate => candidate.rowKey === line.rowKey && candidate.page === line.page
    && candidate.sectionId === line.sectionId && candidate.sectionFragment === line.sectionFragment && candidate.column === line.column
    && candidate.tableCellId === line.tableCellId && !!candidate.repeatedHeader === !!line.repeatedHeader
    && sameTarget(candidate.target, line.target) && sameDocumentBlockHost(candidate.host, line.host))
    .sort((left, right) => (left.fragmentIndex ?? 0) - (right.fragmentIndex ?? 0) || left.start - right.start)
}

export function documentPhysicalLineNavigation(layout: DocumentLayout, point: DocumentPoint, direction: 'up' | 'down' | 'home' | 'end',
  preferredX?: number): DocumentPoint | undefined {
  const line = lineForPoint(layout, point)
  if (!line) return undefined
  const fragments = documentPhysicalLineFragments(layout, line)
  if (direction === 'home' || direction === 'end') {
    const edge = direction === 'home' ? fragments[0]! : fragments.at(-1)!
    return { ...edge.target, offset: direction === 'home' ? edge.start : edge.end, pageAffinity: edge.page }
  }
  const indices = fragments.map(fragment => layout.lines.indexOf(fragment)), step = direction === 'up' ? -1 : 1
  const adjacent = layout.lines[(step < 0 ? Math.min(...indices) : Math.max(...indices)) + step]
  if (!adjacent) return undefined
  const x = preferredX ?? caretRect(layout, point)?.x ?? line.x
  const nextFragments = documentPhysicalLineFragments(layout, adjacent)
  const distance = (candidate: LayoutLine) => x < candidate.x ? candidate.x - x
    : x > candidate.x + candidate.width ? x - candidate.x - candidate.width : 0
  const next = nextFragments.reduce((best, candidate) => distance(candidate) < distance(best) ? candidate : best)
  return hitTestDocumentLine(next, x)
}

export function documentStoryLayoutLines(layout: DocumentLayout, location: { page: number; area: 'header' | 'footer' }): LayoutStoryLine[] {
  const story = layout.pageHeaderFooters[location.page]
  if (story?.isParityBlank) return []
  return (location.area === 'header' ? story?.headerLayoutLines : story?.footerLayoutLines) ?? []
}

export function documentStoryCaretRect(lines: readonly LayoutStoryLine[], offset: number): LayoutRect | undefined {
  // A cached field may wrap across several lines, all carrying the same object
  // range. Its leading caret uses the first glyph and its trailing caret uses
  // the last glyph, rather than stopping at the first wrapped fragment.
  for (const line of lines) {
    const next = line.glyphs.find(glyph => glyph.offset === offset)
    if (next) return { x: next.x, y: line.y, width: 1, height: line.height }
    if (!line.glyphs.length && offset === line.start) return { x: line.x, y: line.y, width: 1, height: line.height }
  }
  for (const line of [...lines].reverse()) {
    const last = [...line.glyphs].reverse().find(glyph => glyph.end === offset)
    if (last) return { x: last.x + last.width, y: line.y, width: 1, height: line.height }
    if (!line.glyphs.length && offset === line.end) return { x: line.x, y: line.y, width: 1, height: line.height }
  }
  const line = lines.find((line, index) => offset >= line.start && (offset < line.end || index === lines.length - 1 || offset === line.end && line.end !== lines[index + 1]?.start)) ?? lines.at(-1)
  if (!line) return undefined
  const next = line.glyphs.find(glyph => glyph.offset >= offset)
  const last = line.glyphs.at(-1)
  return { x: next?.x ?? (last ? last.x + last.width : line.x), y: line.y, width: 1, height: line.height }
}

export function hitTestDocumentStory(lines: readonly LayoutStoryLine[], x: number, y: number): number {
  const line = lines.find(line => y >= line.y && y < line.y + line.height)
    ?? [...lines].sort((a, b) => Math.abs(y - a.y - a.height / 2) - Math.abs(y - b.y - b.height / 2))[0]
  if (!line) return 0
  for (let index = 0; index < line.glyphs.length;) {
    const first = line.glyphs[index]!
    let last = first
    while (index + 1 < line.glyphs.length && line.glyphs[index + 1]!.offset === first.offset && line.glyphs[index + 1]!.end === first.end) last = line.glyphs[++index]!
    if (x < first.x + (last.x + last.width - first.x) / 2) return first.offset
    if (x < last.x + last.width) return first.end
    index++
  }
  return line.end
}

export function caretRect(layout: DocumentLayout, point: DocumentPoint): LayoutRect | undefined {
  const line = lineForPoint(layout, point)
  if (!line) return undefined
  const sourceGlyphs = line.glyphs.filter(glyph => glyph.note?.role !== 'bodyLabel')
  const glyph = sourceGlyphs.find((candidate) => point.offset <= candidate.offset)
  const last = sourceGlyphs.at(-1), label = line.glyphs.find(glyph => glyph.note?.role === 'bodyLabel')
  const rawX = glyph?.x ?? (last ? last.x + last.width : label ? label.x + label.width : line.x)
  const image = line.pictureInputAnchorImageId ? layout.images.find(image => image.blockId === line.pictureInputAnchorImageId
    && image.page === line.page && image.tableCellId === line.tableCellId && !!image.repeatedHeader === !!line.repeatedHeader) : undefined
  const clip = image ? line.clip ? intersectTableRect(image, line.clip) : image : line.clip
  const width = clip ? Math.min(1.3, Math.max(0, clip.width)) : 1.3
  // At an exact-width cell's trailing edge the caret still occupies a visible
  // pixel inside the content clip, rather than intersecting to zero width.
  const x = clip ? Math.max(clip.x, Math.min(rawX, clip.x + clip.width - width)) : rawX
  const inset = Math.min(2, Math.max(0, line.height - 1) / 2)
  const caret = { x, y: line.y + inset, width, height: Math.min(line.height - inset, Math.max(12, line.height - 4)) }
  const clipped = clip ? intersectTableRect(caret, clip) : caret
  return clipped.width > 0 && clipped.height > 0 ? clipped : undefined
}

function visibleLineRect(line: LayoutLine): LayoutRect {
  if (!line.clip) return line
  // A scrolled short/empty code row remains a hit target even when its glyphs are outside the content clip.
  if (line.codeBlock?.gutterX !== undefined) return intersectTableRect({ ...line.clip, y: line.y, height: line.height }, line.clip)
  return intersectTableRect(line, line.clip)
}

export function hitTestDocument(layout: DocumentLayout, x: number, y: number): DocumentPoint | undefined {
  const page = layout.pages.find(page => x >= page.x && x <= page.x + page.width && y >= page.y && y <= page.y + page.height)
  if (page?.isParityBlank) return undefined
  const pageLines = page ? layout.pageContent[page.index]?.lines ?? [] : layout.lines
  const pageCells = page ? layout.pageContent[page.index]?.cells ?? [] : layout.cells
  const cell = pageCells.filter(cell => {
    const visible = cell.clip ? intersectTableRect(cell, cell.clip) : cell
    return x >= visible.x && x <= visible.x + visible.width && y >= visible.y && y <= visible.y + visible.height
  }).sort((left, right) => (right.depth ?? 0) - (left.depth ?? 0) || left.width * left.height - right.width * right.height)[0]
  const visibleLines = pageLines.filter(line => { const visible = visibleLineRect(line); return visible.width > 0 && visible.height > 0 })
  const cellLines = cell?.cellId ? visibleLines.filter(line => line.tableCellId === cell.cellId) : []
  if (cell?.cellId && !cellLines.length) return undefined
  const columnLines = visibleLines.filter(line => {
    const auxiliary = line.host && layout.auxiliaryContent?.find(content => content.page === line.page
      && sameDocumentBlockHost(content.host, line.host) && content.lines.includes(line))
    if (auxiliary) {
      const column = auxiliary.columns?.find(column => column.index === (line.column ?? 0)) ?? auxiliary
      return x >= column.x && x <= column.x + column.width
    }
    const section = layout.pages[line.page]?.sections.find(section => section.fragment === line.sectionFragment)
    const column = section?.columns[line.column ?? 0]
    return column && x >= column.x && x <= column.x + column.width
  })
  let nearest: LayoutLine | undefined
  let best = Number.POSITIVE_INFINITY
  for (const line of cellLines.length ? cellLines : columnLines.length ? columnLines : visibleLines) {
    const visible = visibleLineRect(line)
    const dy = y < visible.y ? visible.y - y : y > visible.y + visible.height ? y - visible.y - visible.height : 0
    const dx = x < visible.x ? visible.x - x : x > visible.x + visible.width ? x - visible.x - visible.width : 0
    const score = dy * 8 + dx
    if (score < best) { best = score; nearest = line }
  }
  if (!nearest) return undefined
  return hitTestDocumentLine(nearest, x)
}

/** Resolve a known visual line without re-entering page/cell hit routing. */
export function hitTestDocumentLine(line: LayoutLine, x: number): DocumentPoint {
  const visibleX = line.clip ? Math.max(line.clip.x, Math.min(x, line.clip.x + line.clip.width)) : x
  const glyph = line.glyphs.find(candidate => candidate.note?.role !== 'bodyLabel' && visibleX < candidate.x + candidate.width / 2)
  return { ...line.target, offset: glyph?.offset ?? line.end, pageAffinity: line.page }
}

export function hitTestChecklist(layout: DocumentLayout, x: number, y: number): string | undefined {
  const line = layout.lines.find(candidate => {
    if (!candidate.checklist) return false
    const bounds = { x: candidate.x - 28, y: candidate.y, width: 22, height: candidate.height }
    const visible = candidate.clip ? intersectTableRect(bounds, candidate.clip) : bounds
    return visible.width > 0 && visible.height > 0 && y >= visible.y && y <= visible.y + visible.height && x >= visible.x && x <= visible.x + visible.width
  })
  return line?.checklist?.blockId
}

export function linesByTarget(layout: DocumentLayout): Map<string, LayoutLine[]> {
  const map = new Map<string, LayoutLine[]>()
  layout.lines.forEach((line) => map.set(targetKey(line.target), [...(map.get(targetKey(line.target)) ?? []), line]))
  return map
}

export const emptyLayout = (): DocumentLayout => ({ width: 850, height: 1179, pages: [], pageContent: [], lines: [], cells: [], images: [], dividers: [], columnSeparators: [], pageHeaderFooters: [], headings: [] })
export const defaultRun = createRun
