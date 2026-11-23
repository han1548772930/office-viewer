import type { DocumentModel } from '../../../domain/model'
import type { DocumentPoint } from '../model/document-model'
import { paragraphListMarkers } from '../model/document-lists'
import { createRun, imageForBlock, sameTarget, targetKey } from '../model/document-model'
import { DocumentSectionFlow } from './document-section-flow'
import { balanceDocumentSection } from './document-column-balance'
import { layoutDocumentParagraph } from './document-paragraph-layout'
import { layoutDocumentTable } from './document-table-layout'
import { arrangeDocumentPages } from './document-page-arrangement'
import { layoutDocumentPageStories } from './document-page-stories'
import { intersectTableRect } from './document-table-measurement'
import type { DocumentLayout, LayoutLine, LayoutMeasurer, LayoutRect, LayoutStoryLine } from './document-layout-types'

export type { DocumentLayout, LayoutCell, LayoutColumn, LayoutDivider, LayoutFlowLocation, LayoutGlyph, LayoutImage, LayoutLine, LayoutMeasurer, LayoutPage, LayoutPageContent, LayoutPageHeaderFooter, LayoutRect, LayoutRightTabStop, LayoutSection, LayoutStoryLine, LayoutTableFragment, LayoutTableRowMetric, LayoutTabGlyph, LayoutTabLeader } from './document-layout-types'
export { PAGE_GAP, PAGE_INSET, STORY_LINE_HEIGHT } from './document-layout-types'
export { createTextMeasurer, runFont } from './document-text-layout'

export function layoutDocument(model: DocumentModel, measurer: LayoutMeasurer): DocumentLayout {
  const flow = new DocumentSectionFlow(model)
  const result: DocumentLayout = { width: 0, height: 0, pages: flow.pages, pageContent: [], lines: [], cells: [],
    images: [], dividers: [], columnSeparators: [], pageHeaderFooters: [], headings: [] }
  const bodyMarkers = paragraphListMarkers(model.blocks)
  for (const block of model.blocks) {
    if (block.type === 'pageBreak') { flow.pageBreak(); continue }
    if (block.type === 'columnBreak') { flow.nextColumn(true); continue }
    if (block.type === 'sectionBreak') {
      const type = block.sectionBreak?.type ?? 'nextPage'
      if (type === 'continuous' && flow.columnsChange(block.id)) balanceDocumentSection(result, flow.section)
      flow.startSection(block.id, type)
      continue
    }
    flow.addSpace(block.style.spaceBefore)
    if (block.type === 'table') layoutDocumentTable(block, bodyMarkers, flow, result, measurer)
    else if (block.type === 'image') {
      const image = imageForBlock(block)
      if (!image) continue
      const nativeWidth = Math.max(1, image.width || 320), nativeHeight = Math.max(1, image.height || nativeWidth * 0.625)
      const size = () => {
        const scale = Math.min(1, flow.column.width / nativeWidth, flow.page.content.height / nativeHeight)
        return { width: nativeWidth * scale, height: nativeHeight * scale }
      }
      let prepared = size()
      while (true) {
        flow.ensureRoom(prepared.height)
        const next = size()
        if (Math.abs(next.width - prepared.width) < 1e-6 && Math.abs(next.height - prepared.height) < 1e-6) break
        prepared = next
      }
      const shift = block.style.align === 'center' ? (flow.column.width - prepared.width) / 2
        : block.style.align === 'right' ? flow.column.width - prepared.width : 0
      result.images.push({ x: flow.column.x + shift, y: flow.cursor, ...prepared, blockId: block.id,
        src: image.src, alt: image.alt, page: flow.page.index, ...flow.location })
      flow.consume(prepared.height)
    } else if (block.type === 'divider') {
      flow.ensureRoom(18)
      result.dividers.push({ x: flow.column.x, y: flow.cursor + 9, width: flow.column.width, height: 1, page: flow.page.index, ...flow.location })
      flow.consume(18)
    } else layoutDocumentParagraph(block, bodyMarkers.get(block.id), flow, result, measurer)
    flow.addSpace(block.style.spaceAfter)
  }
  arrangeDocumentPages(result)
  result.pageHeaderFooters = layoutDocumentPageStories(model, result.pages, flow.settings, measurer)
  result.auxiliaryContent = result.pageHeaderFooters.flatMap(page => [page.headerBlockLayout, page.footerBlockLayout].filter((content): content is NonNullable<typeof content> => !!content))
  return result
}

export function lineForPoint(layout: DocumentLayout, point: DocumentPoint): LayoutLine | undefined {
  const all = layout.lines.filter((line) => sameTarget(line.target, point))
  const atOffset = (lines: LayoutLine[]) => lines.find((line, index) => point.offset >= line.start && point.offset <= line.end
    && (point.offset < line.end || index === lines.length - 1 || line.end !== lines[index + 1]?.start))
  const preferred = point.pageAffinity === undefined ? [] : all.filter(line => line.page === point.pageAffinity)
  const original = all.filter(line => !line.repeatedHeader)
  return atOffset(preferred) ?? atOffset(original) ?? atOffset(all) ?? original.at(-1) ?? all.at(-1)
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
  const glyph = line.glyphs.find((candidate) => point.offset <= candidate.offset)
  const rawX = glyph?.x ?? (line.glyphs.at(-1) ? line.glyphs.at(-1)!.x + line.glyphs.at(-1)!.width : line.x)
  const width = line.clip ? Math.min(1.3, Math.max(0, line.clip.width)) : 1.3
  // At an exact-width cell's trailing edge the caret still occupies a visible
  // pixel inside the content clip, rather than intersecting to zero width.
  const x = line.clip ? Math.max(line.clip.x, Math.min(rawX, line.clip.x + line.clip.width - width)) : rawX
  const caret = { x, y: line.y + 2, width, height: Math.max(12, line.height - 4) }
  const clipped = line.clip ? intersectTableRect(caret, line.clip) : caret
  return clipped.width > 0 && clipped.height > 0 ? clipped : undefined
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
  const visibleLines = pageLines.filter(line => !line.clip || intersectTableRect(line, line.clip).width > 0 && intersectTableRect(line, line.clip).height > 0)
  const cellLines = cell?.cellId ? visibleLines.filter(line => line.tableCellId === cell.cellId) : []
  if (cell?.cellId && !cellLines.length) return undefined
  const columnLines = visibleLines.filter(line => {
    const section = layout.pages[line.page]?.sections.find(section => section.fragment === line.sectionFragment)
    const column = section?.columns[line.column ?? 0]
    return column && x >= column.x && x <= column.x + column.width
  })
  let nearest: LayoutLine | undefined
  let best = Number.POSITIVE_INFINITY
  for (const line of cellLines.length ? cellLines : columnLines.length ? columnLines : visibleLines) {
    const visible = line.clip ? intersectTableRect(line, line.clip) : line
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
  const glyph = line.glyphs.find(candidate => visibleX < candidate.x + candidate.width / 2)
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
