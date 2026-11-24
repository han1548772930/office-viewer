import type { DocumentLayout, LayoutRect, LayoutStoryLine } from './document-layout-types'
import { PAGE_INSET } from './document-layout-types'

/** Center mixed physical page sizes without changing their local content geometry. */
export function arrangeDocumentPages(layout: DocumentLayout): void {
  const maximumWidth = layout.pages.reduce((width, page) => Math.max(width, page.width), 0)
  layout.width = maximumWidth + PAGE_INSET * 2
  const shifts = layout.pages.map(page => (maximumWidth - page.width) / 2)
  const moved = new Set<object>()
  const shiftRect = (item: LayoutRect, dx: number): boolean => {
    if (moved.has(item)) return false
    moved.add(item); item.x += dx
    return true
  }
  const shiftClipped = (item: LayoutRect & { clip?: LayoutRect }, dx: number) => {
    shiftRect(item, dx)
    if (item.clip) shiftRect(item.clip, dx)
  }
  const shiftLine = (line: LayoutStoryLine & { clip?: LayoutRect; markerStyle?: { x: number } }, dx: number) => {
    if (!shiftRect(line, dx)) return
    if (line.clip) shiftRect(line.clip, dx)
    if (line.markerStyle && !moved.has(line.markerStyle)) { moved.add(line.markerStyle); line.markerStyle.x += dx }
    for (const glyph of line.glyphs) shiftRect(glyph, dx)
  }
  for (const page of layout.pages) {
    const dx = shifts[page.index]!
    shiftRect(page, dx); shiftRect(page.content, dx)
    for (const section of page.sections) {
      shiftRect(section, dx)
      for (const column of section.columns) shiftRect(column, dx)
      const height = Math.min(section.height, section.usedHeight)
      if (section.columnSeparator !== 'line' || height <= 0) continue
      section.columns.slice(0, -1).forEach((column, index) => {
        const next = section.columns[index + 1]!
        layout.columnSeparators.push({ x: (column.x + column.width + next.x) / 2, y: section.y, width: 1, height,
          page: page.index, sectionId: section.sectionId, sectionFragment: section.fragment, column: index })
      })
    }
  }
  for (const line of layout.lines) shiftLine(line, shifts[line.page]!)
  for (const item of [...layout.cells, ...layout.images, ...layout.dividers]) {
    shiftClipped(item, shifts[item.page]!)
  }
  for (const table of layout.tableFragments ?? []) shiftClipped(table, shifts[table.page]!)
  for (const content of [...(layout.auxiliaryContent ?? []), ...(layout.noteDecorations ?? [])]) {
    const dx = shifts[content.page]!
    shiftRect(content, dx)
    if ('columns' in content) for (const column of content.columns ?? []) shiftRect(column, dx)
    for (const line of content.lines) shiftLine(line, dx)
    for (const item of [...content.cells, ...content.images, ...content.dividers, ...content.tableFragments]) shiftClipped(item, dx)
  }
  // Block-story lines and headerLayoutLines share objects. Shift each geometry
  // once, including their shared page clips, before publishing page indexes.
  for (const story of layout.pageHeaderFooters) {
    const dx = shifts[story.page]!
    shiftRect(story, dx)
    for (const line of [...story.headerLayoutLines, ...story.footerLayoutLines]) shiftLine(line, dx)
  }
  const last = layout.pages.at(-1)
  layout.height = last ? last.y + last.height + PAGE_INSET : 0
  layout.pageContent = layout.pages.map(() => ({ lines: [], cells: [], images: [], dividers: [], columnSeparators: [] }))
  for (const line of layout.lines) layout.pageContent[line.page]?.lines.push(line)
  for (const cell of layout.cells) layout.pageContent[cell.page]?.cells.push(cell)
  for (const image of layout.images) layout.pageContent[image.page]?.images.push(image)
  for (const divider of layout.dividers) layout.pageContent[divider.page]?.dividers.push(divider)
  for (const separator of layout.columnSeparators) layout.pageContent[separator.page]?.columnSeparators.push(separator)
}
