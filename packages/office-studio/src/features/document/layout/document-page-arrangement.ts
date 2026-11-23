import type { DocumentLayout } from './document-layout-types'
import { PAGE_INSET } from './document-layout-types'

/** Center mixed physical page sizes without changing their local content geometry. */
export function arrangeDocumentPages(layout: DocumentLayout): void {
  const maximumWidth = layout.pages.reduce((width, page) => Math.max(width, page.width), 0)
  layout.width = maximumWidth + PAGE_INSET * 2
  const shifts = layout.pages.map(page => (maximumWidth - page.width) / 2)
  for (const page of layout.pages) {
    const dx = shifts[page.index]!
    page.x += dx; page.content.x += dx
    for (const section of page.sections) {
      section.x += dx
      for (const column of section.columns) column.x += dx
      const height = Math.min(section.height, section.usedHeight)
      if (section.columnSeparator !== 'line' || height <= 0) continue
      section.columns.slice(0, -1).forEach((column, index) => {
        const next = section.columns[index + 1]!
        layout.columnSeparators.push({ x: (column.x + column.width + next.x) / 2, y: section.y, width: 1, height,
          page: page.index, sectionId: section.sectionId, sectionFragment: section.fragment, column: index })
      })
    }
  }
  for (const line of layout.lines) {
    const dx = shifts[line.page]!
    line.x += dx
    if (line.markerStyle) line.markerStyle.x += dx
    for (const glyph of line.glyphs) glyph.x += dx
  }
  for (const item of [...layout.cells, ...layout.images, ...layout.dividers]) item.x += shifts[item.page]!
  const last = layout.pages.at(-1)
  layout.height = last ? last.y + last.height + PAGE_INSET : 0
  layout.pageContent = layout.pages.map(() => ({ lines: [], cells: [], images: [], dividers: [], columnSeparators: [] }))
  for (const line of layout.lines) layout.pageContent[line.page]?.lines.push(line)
  for (const cell of layout.cells) layout.pageContent[cell.page]?.cells.push(cell)
  for (const image of layout.images) layout.pageContent[image.page]?.images.push(image)
  for (const divider of layout.dividers) layout.pageContent[divider.page]?.dividers.push(divider)
  for (const separator of layout.columnSeparators) layout.pageContent[separator.page]?.columnSeparators.push(separator)
}
