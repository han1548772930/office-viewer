import type { DocumentBlock } from '../../../domain/model'
import { createRun, runText, sliceRuns, type DocumentRun } from '../model/document-model'
import type { ParagraphListMarker } from '../model/document-lists'
import { codeBlockData } from '../code-blocks/code-blocks'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutImage, LayoutLine, LayoutMeasurer } from './document-layout-types'
import type { DocumentSectionFlow } from './document-section-flow'
import { paragraphLineHeight, positionLine, runFont, wrapRuns } from './document-text-layout'
import { documentTocPageTab } from '../toc/document-toc-layout'

export interface ParagraphGeometry {
  indent: number
  firstIndent: number
  rightIndent: number
  markerIndent?: number
  markerFont?: string
  markerHeight: number
  markerWidth: number
  width: (index: number) => number
}

export function paragraphGeometry(block: DocumentBlock, list: ParagraphListMarker | undefined, available: number, measurer: LayoutMeasurer): ParagraphGeometry {
  const legacyIndent = block.style.indent * 24
  const left = block.style.indentStart ?? legacyIndent
  let indent = left, firstIndent = left + (block.style.indentFirstLine ?? 0) - (block.style.hanging ?? 0)
  let rightIndent = block.style.indentEnd ?? 0
  let markerIndent: number | undefined, markerFont: string | undefined, markerHeight = 0, markerWidth = 0
  if (list) {
    const extra = block.style.indentStart === undefined ? legacyIndent : 0
    indent = list.inset + extra
    rightIndent = list.rightIndent
    markerIndent = list.markerIndent + extra
    markerFont = runFont(createRun('', { fontFamily: list.fontFamily, fontSize: list.fontSize, bold: list.bold, italic: list.italic }))
    markerHeight = list.picture ? list.picture.heightPx ?? list.fontSize * 4 / 3 : list.fontSize * 4 / 3
    markerWidth = list.picture ? list.picture.widthPx ?? markerHeight : measurer.measure(list.marker, markerFont)
    if (list.alignment === 'center') markerIndent -= markerWidth / 2
    else if (list.alignment === 'right' || list.alignment === 'end') markerIndent -= markerWidth
    firstIndent = Math.max(indent, markerIndent + markerWidth + (markerWidth ? list.gap : 0))
  } else if (block.type === 'checklist') {
    indent += 28 * (block.level + 1)
    firstIndent = indent
  } else if (block.type === 'quote') {
    indent += 16
    firstIndent += 16
  }
  return { indent, firstIndent, rightIndent, markerIndent, markerFont, markerHeight, markerWidth,
    width: index => Math.max(1, available - (index === 0 ? firstIndent : indent) - rightIndent) }
}

export function appendListMarker(line: LayoutLine, block: DocumentBlock, list: ParagraphListMarker | undefined, geometry: ParagraphGeometry, left: number, images: LayoutImage[]): void {
  if (!list || geometry.markerIndent === undefined) return
  if (list.picture?.src) {
    images.push({ blockId: block.id, src: list.picture.src, alt: '项目符号', x: left + geometry.markerIndent,
      y: line.y + (line.height - geometry.markerHeight) / 2, width: geometry.markerWidth, height: geometry.markerHeight, page: line.page,
      sectionId: line.sectionId, sectionFragment: line.sectionFragment, column: line.column })
  } else if (geometry.markerFont) line.markerStyle = { x: left + geometry.markerIndent, font: geometry.markerFont, color: list.color }
}

export function layoutDocumentParagraph(block: DocumentBlock, list: ParagraphListMarker | undefined, flow: DocumentSectionFlow, result: DocumentLayout, measurer: LayoutMeasurer,
  display?: DocumentTextLayoutDisplay): void {
  const codeData = codeBlockData(block), text = runText(block.runs)
  let available = flow.column.width, geometry = paragraphGeometry(block, list, available, measurer)
  let rawIndex = 0, lineIndex = 0, baseOffset = 0
  const pageTab = documentTocPageTab(block)
  const rightTabs = (index: number) => !pageTab || pageTab.offset < baseOffset ? [] : pageTab.stops.map(stop => ({ ...stop,
    offset: pageTab.offset - baseOffset,
    ...(stop.right === undefined ? {} : { right: Math.max(0, stop.right - (index + lineIndex === 0 ? geometry.firstIndent : geometry.indent)) }),
  }))
  const textDisplay = () => ({ references: display?.references, leading: lineIndex === 0 && baseOffset === 0 ? display?.leadingLabels?.get(block.id) : undefined })
  let lines = wrapRuns(block.runs as DocumentRun[], geometry.width, measurer, block.style.wordWrap === true, undefined, pageTab ? rightTabs : undefined, textDisplay())
  while (rawIndex < lines.length) {
    const raw = lines[rawIndex]!
    const height = paragraphLineHeight(Math.max(raw.fontHeight, lineIndex === 0 ? geometry.markerHeight : 0), block.style)
    flow.advanceToPage(flow.referencePageFloor(raw.glyphs.filter(glyph => glyph.note?.role === 'reference').map(glyph => glyph.run.id)))
    flow.ensureRoom(height)
    if (Math.abs(available - flow.column.width) > 1e-6) {
      // A continuous section can retain the old page's margins until overflow.
      // Rewrap only the unplaced suffix, with its original UTF-16 positions.
      baseOffset += raw.start
      available = flow.column.width
      geometry = paragraphGeometry(block, list, available, measurer)
      const remaining = sliceRuns(block.runs as DocumentRun[], baseOffset, text.length)
      if (!remaining.length && block.runs.length) remaining.push({ ...block.runs.at(-1)!, text: '' })
      lines = wrapRuns(remaining, index => geometry.width(index + lineIndex), measurer, block.style.wordWrap === true, undefined, pageTab ? rightTabs : undefined, textDisplay())
      rawIndex = 0
      continue
    }
    if (lineIndex === 0 && block.type === 'heading') result.headings.push({ id: block.id, text, level: block.level || 1, y: flow.cursor, page: flow.page.index })
    const marker = lineIndex === 0 ? list?.marker ?? (block.type === 'quote' ? '│' : undefined) : undefined
    const indent = lineIndex === 0 ? geometry.firstIndent : geometry.indent
    const line = positionLine({ ...raw, start: baseOffset + raw.start, end: baseOffset + raw.end,
      glyphs: raw.glyphs.map(glyph => ({ ...glyph, offset: baseOffset + glyph.offset, end: baseOffset + glyph.end })) },
    { blockId: block.id }, flow.column.x + indent, flow.cursor, geometry.width(lineIndex), height, flow.page.index,
    block.style.align, rawIndex < lines.length - 1, marker)
    Object.assign(line, flow.location)
    if (lineIndex === 0) appendListMarker(line, block, list, geometry, flow.column.x, result.images)
    if (codeData) {
      line.codeBlock = { language: codeData.language, lineNumber: lineIndex + 1 }
      line.x += 8
      line.width = Math.max(2, line.width - 16)
      for (const glyph of line.glyphs) glyph.x += 8
    }
    if (lineIndex === 0 && block.type === 'checklist') line.checklist = { blockId: block.id, checked: block.checklist?.checked === true }
    result.lines.push(line)
    if (lineIndex === 0 && block.metadata.noteSeparator && typeof block.metadata.noteSeparator === 'object') {
      const separator = block.metadata.noteSeparator, fraction = Reflect.get(separator, 'widthFraction'), color = Reflect.get(separator, 'color')
      if (Reflect.get(separator, 'version') === 1 && typeof fraction === 'number' && Number.isFinite(fraction) && fraction >= 0 && fraction <= 1) {
        result.dividers.push({ x: flow.column.x + indent, y: flow.cursor + Math.max(1, height - 2), width: geometry.width(lineIndex) * fraction,
          height: 1, page: flow.page.index, ...flow.location, ...(typeof color === 'string' ? { color } : {}) })
      }
    }
    flow.consume(height)
    rawIndex++; lineIndex++
  }
}
