import type { DocumentBlock } from '../../../domain/model'
import { tableForBlock, type DocumentRun } from '../model/document-model'
import type { ParagraphListMarker } from '../model/document-lists'
import type { DocumentLayout, LayoutMeasurer } from './document-layout-types'
import type { DocumentSectionFlow } from './document-section-flow'
import { appendListMarker, paragraphGeometry } from './document-paragraph-layout'
import { paragraphLineHeight, positionLine, wrapRuns } from './document-text-layout'

export function layoutDocumentTable(block: DocumentBlock, markers: Map<string, ParagraphListMarker>, flow: DocumentSectionFlow, result: DocumentLayout, measurer: LayoutMeasurer): void {
  const table = tableForBlock(block)
  if (!table?.rows.length) return
  const colCount = Math.max(1, ...table.rows.map(row => row.cells.reduce((sum, cell) => sum + Math.max(1, cell.colSpan), 0)))
  const specifiedWidths = table.columnWidths?.length === colCount && table.columnWidths.every(width => Number.isFinite(width) && width > 0)
    ? table.columnWidths : undefined
  const totalWidth = specifiedWidths?.reduce((sum, width) => sum + width, 0)
  table.rows.forEach((row, rowIndex) => {
    const prepare = (available: number) => {
      const widths = specifiedWidths && totalWidth ? specifiedWidths.map(width => width * available / totalWidth)
        : Array.from({ length: colCount }, () => available / colCount)
      let column = 0
      const cells = row.cells.map((cell, colIndex) => {
        const width = widths.slice(column, column + Math.max(1, cell.colSpan)).reduce((sum, value) => sum + value, 0)
        const xOffset = widths.slice(0, column).reduce((sum, value) => sum + value, 0)
        column += Math.max(1, cell.colSpan)
        const paragraphs = cell.blocks.filter(paragraph => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(paragraph.type)).map(paragraph => {
          const list = markers.get(paragraph.id), geometry = paragraphGeometry(paragraph, list, width - 16, measurer)
          return { paragraph, list, geometry, lines: wrapRuns(paragraph.runs as DocumentRun[], geometry.width, measurer, paragraph.style.wordWrap === true) }
        })
        const height = Math.max(34, 16 + paragraphs.reduce((sum, item) => sum + item.paragraph.style.spaceBefore
          + item.lines.reduce((lineSum, line, index) => lineSum + paragraphLineHeight(Math.max(line.fontHeight, index === 0 ? item.geometry.markerHeight : 0), item.paragraph.style), 0)
          + item.paragraph.style.spaceAfter, 0))
        return { cell, colIndex, width, xOffset, paragraphs, height }
      })
      return { cells, height: Math.max(34, ...cells.map(cell => cell.height)) }
    }
    let available = flow.column.width, prepared = prepare(available)
    while (true) {
      flow.ensureRoom(prepared.height)
      if (Math.abs(available - flow.column.width) <= 1e-6) break
      available = flow.column.width
      prepared = prepare(available)
    }
    for (const { cell, colIndex, width, xOffset, paragraphs } of prepared.cells) {
      const x = flow.column.x + xOffset, target = { blockId: block.id, row: rowIndex, col: colIndex }
      result.cells.push({ x, y: flow.cursor, width, height: prepared.height, page: flow.page.index, target, ...flow.location,
        fill: typeof cell.style.fill === 'string' ? cell.style.fill : typeof cell.style.background === 'string' ? cell.style.background : '#ffffff' })
      let y = flow.cursor + 8
      for (const item of paragraphs) {
        y += item.paragraph.style.spaceBefore
        item.lines.forEach((raw, index) => {
          const height = paragraphLineHeight(Math.max(raw.fontHeight, index === 0 ? item.geometry.markerHeight : 0), item.paragraph.style)
          const indent = index === 0 ? item.geometry.firstIndent : item.geometry.indent
          const line = positionLine(raw, { ...target, cellBlockId: item.paragraph.id }, x + 8 + indent, y,
            item.geometry.width(index), height, flow.page.index, item.paragraph.style.align, index < item.lines.length - 1, index === 0 ? item.list?.marker : undefined)
          Object.assign(line, flow.location)
          if (index === 0) appendListMarker(line, item.paragraph, item.list, item.geometry, x + 8, result.images)
          if (index === 0 && item.paragraph.type === 'checklist') line.checklist = { blockId: item.paragraph.id, checked: item.paragraph.checklist?.checked === true }
          result.lines.push(line)
          y += height
        })
        y += item.paragraph.style.spaceAfter
      }
    }
    flow.consume(prepared.height)
  })
}
