import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentBlockHostId, sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { imageForBlock, runText, tableForBlock, type DocumentRun } from '../model/document-model'
import { paragraphListMarkers } from '../model/document-lists'
import { documentFieldAutomaticRefreshBlocked, documentTextFieldParagraphs, documentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit } from '../model/document-text-fields'
import { documentModelForBlockHost, resolveDocumentBlockHost } from '../services/document-block-host'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentTableHorizontalOffset } from '../model/document-table-style'
import { remapDocumentStoryFieldRuns } from '../model/document-story'
import { sliceRuns } from '../model/document-model'
import { splitDocumentRunRange } from '../model/document-run-range'
import { appendListMarker, paragraphGeometry } from './document-paragraph-layout'
import { paragraphLineHeight, positionLine, wrapRuns } from './document-text-layout'
import { intersectTableRect, measureDocumentTable, translateTableLine, translateTableRect } from './document-table-measurement'
import type { DocumentLayout, LayoutAuxiliaryContent, LayoutLine, LayoutMeasurer, LayoutRect } from './document-layout-types'

interface FieldDisplayMap { sourceStart: number; sourceEnd: number; start: number; end: number }

/** PAGE displays vary per physical page while their source cached-result coordinates remain stable. */
function storyPageDisplay(model: DocumentModel, pageNumber: number, total: number): { model: DocumentModel; offsets: Map<string, FieldDisplayMap[]> } {
  const display = cloneModel(model), offsets = new Map<string, FieldDisplayMap[]>()
  for (const paragraph of documentTextFieldParagraphs(display)) {
    const fields = [...documentTextFields(paragraph.block)].filter(field => ['PAGE', 'NUMPAGES'].includes(field.field.toUpperCase())
      && !field.endBlockId && !documentFieldAutomaticRefreshBlocked(field)).sort((a, b) => a.startOffset - b.startOffset)
    const maps: FieldDisplayMap[] = []
    let delta = 0
    for (const field of fields) {
      const result = String(field.field.toUpperCase() === 'PAGE' ? pageNumber : total)
      maps.push({ sourceStart: field.startOffset, sourceEnd: field.endOffset, start: field.startOffset + delta, end: field.startOffset + delta + result.length })
      delta += result.length - (field.endOffset - field.startOffset)
    }
    for (const field of [...fields].reverse()) {
      const result = String(field.field.toUpperCase() === 'PAGE' ? pageNumber : total)
      const { before, after } = splitDocumentRunRange(paragraph.block.runs, field.startOffset, field.endOffset)
      const cache = remapDocumentStoryFieldRuns(sliceRuns(paragraph.block.runs, field.startOffset, field.endOffset), result, documentDefaultStyle(display).text)
      transformDocumentTextFieldsForBlockEdit(paragraph.block, field.startOffset, field.endOffset, result, field.id, display)
      paragraph.block.runs = [...before, ...cache, ...after]
      setDocumentTextFields(paragraph.block, documentTextFields(paragraph.block).map(candidate => candidate.id === field.id ? { ...candidate, cachedResult: result } : candidate))
    }
    if (maps.length) offsets.set(paragraph.block.id, maps)
  }
  return { model: display, offsets }
}

function sourceLineOffsets(line: LayoutLine, offsets: ReadonlyMap<string, readonly FieldDisplayMap[]>): LayoutLine {
  const maps = offsets.get(line.target.cellBlockId ?? line.target.blockId)
  if (!maps) return line
  const offset = (value: number, end = false) => {
    let delta = 0
    for (const field of maps) {
      if (value < field.start) break
      if (value === field.start) return field.sourceStart
      if (value < field.end) return end ? field.sourceEnd : field.sourceStart
      delta += (field.end - field.start) - (field.sourceEnd - field.sourceStart)
    }
    return value - delta
  }
  return { ...line, start: offset(line.start), end: offset(line.end, true),
    glyphs: line.glyphs.map(glyph => ({ ...glyph, offset: offset(glyph.offset), end: offset(glyph.end, true) })) }
}

/** Header/footer blocks share the existing paragraph and nested-table measurement kernel. */
export function layoutDocumentAuxiliaryBlocks(model: DocumentModel, host: DocumentAuxiliaryBlockHost, bounds: LayoutRect, page: number,
  pageNumber: number, total: number, measurer: LayoutMeasurer, bottom = false): LayoutAuxiliaryContent {
  const view = resolveDocumentBlockHost(model, host), scoped = documentModelForBlockHost(model, host)
  const display = host.kind === 'story' ? storyPageDisplay(scoped, pageNumber, total) : { model: scoped, offsets: new Map<string, FieldDisplayMap[]>() }
  const markers = paragraphListMarkers(display.model.blocks)
  const output: LayoutAuxiliaryContent = { ...bounds, host: { ...host }, sourceHost: view.sourceHost, page,
    lines: [], cells: [], images: [], dividers: [], tableFragments: [] }
  let cursor = 0
  for (const block of display.model.blocks) {
    cursor += block.style.spaceBefore
    if (block.type === 'table') {
      const measured = measureDocumentTable(block, bounds.width, markers, measurer, 0, bounds.height)
      const table = tableForBlock(block)!, dx = documentTableHorizontalOffset(table, measured.width, bounds.width)
      cursor += Math.max(0, table.style?.offsetY ?? 0)
      output.lines.push(...measured.lines.map(line => translateTableLine(line, dx, cursor)))
      output.cells.push(...measured.cells.map(cell => ({ ...translateTableRect(cell, dx, cursor), ...(cell.clip ? { clip: translateTableRect(cell.clip, dx, cursor) } : {}) })))
      output.images.push(...measured.images.map(image => ({ ...translateTableRect(image, dx, cursor), ...(image.clip ? { clip: translateTableRect(image.clip, dx, cursor) } : {}) })))
      output.dividers.push(...measured.dividers.map(divider => ({ ...translateTableRect(divider, dx, cursor), ...(divider.clip ? { clip: translateTableRect(divider.clip, dx, cursor) } : {}) })))
      output.tableFragments.push(...measured.tables.map(table => ({ ...translateTableRect(table, dx, cursor),
        ...(table.clip ? { clip: translateTableRect(table.clip, dx, cursor) } : {}), rows: table.rows.map(row => ({ ...row, y: row.y + cursor })),
        page, sourceStart: 0, sourceEnd: table.height, repeatedHeader: false, continuedBefore: false, continuedAfter: false })))
      cursor += measured.height
    } else if (block.type === 'image') {
      const image = imageForBlock(block)
      if (image) {
        const scale = Math.min(1, bounds.width / Math.max(1, image.width), bounds.height / Math.max(1, image.height))
        const width = Math.max(1, image.width) * scale, height = Math.max(1, image.height) * scale
        const dx = block.style.align === 'right' ? bounds.width - width : block.style.align === 'center' ? (bounds.width - width) / 2 : 0
        output.images.push({ blockId: block.id, src: image.src, alt: image.alt, x: dx, y: cursor, width, height, page })
        cursor += height
      }
    } else if (block.type === 'divider') {
      output.dividers.push({ x: 0, y: cursor + 9, width: bounds.width, height: 1, page }); cursor += 18
    } else if (!['pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) {
      const marker = markers.get(block.id), geometry = paragraphGeometry(block, marker, bounds.width, measurer)
      const wrapped = wrapRuns(block.runs as DocumentRun[], geometry.width, measurer, block.style.wordWrap === true)
      for (const [index, raw] of wrapped.entries()) {
        const height = paragraphLineHeight(Math.max(raw.fontHeight, index === 0 ? geometry.markerHeight : 0), block.style)
        const line = positionLine(raw, { blockId: block.id }, index === 0 ? geometry.firstIndent : geometry.indent, cursor,
          geometry.width(index), height, page, block.style.align, index < wrapped.length - 1, index === 0 ? marker?.marker : undefined)
        if (index === 0) appendListMarker(line, block, marker, geometry, 0, output.images)
        output.lines.push(line); cursor += height
      }
    }
    cursor += block.style.spaceAfter
  }
  const top = bottom ? bounds.y + bounds.height - cursor : bounds.y, clip = { ...bounds }
  output.y = top; output.height = cursor
  output.lines = output.lines.map(line => {
    const translated = translateTableLine(sourceLineOffsets(line, display.offsets), bounds.x, top)
    return { ...translated, host: { ...host }, page, sectionId: view.sectionId, clip: translated.clip ? intersectTableRect(translated.clip, clip) : clip }
  })
  const positioned = <T extends LayoutRect & { clip?: LayoutRect }>(item: T): T & { host: DocumentAuxiliaryBlockHost; page: number; sectionId: string } => ({
    ...translateTableRect(item, bounds.x, top), host: { ...host }, page, sectionId: view.sectionId,
    clip: item.clip ? intersectTableRect(translateTableRect(item.clip, bounds.x, top), clip) : clip })
  output.cells = output.cells.map(positioned)
  output.images = output.images.map(positioned)
  output.dividers = output.dividers.map(positioned)
  output.tableFragments = output.tableFragments.map(fragment => ({ ...positioned(fragment), rows: fragment.rows.map(row => ({ ...row, y: row.y + top })) }))
  return output
}

/** A view of the same physical pages for caret, cell hit-testing, and table overlays in one host. */
export function documentLayoutForBlockHost(layout: DocumentLayout, host?: DocumentAuxiliaryBlockHost): DocumentLayout {
  if (!host) return layout
  if (host.kind === 'note' && !layout.auxiliaryContent?.some(content => sameDocumentBlockHost(content.host, host))) return layout
  const content = (layout.auxiliaryContent ?? []).filter(content => sameDocumentBlockHost(content.host, host))
  const lines = content.flatMap(content => content.lines), cells = content.flatMap(content => content.cells), images = content.flatMap(content => content.images)
  const dividers = content.flatMap(content => content.dividers), tableFragments = content.flatMap(content => content.tableFragments)
  return { ...layout, lines, cells, images, dividers, columnSeparators: [], tableFragments, headings: [],
    pageContent: layout.pages.map(page => ({ lines: lines.filter(line => line.page === page.index), cells: cells.filter(cell => cell.page === page.index),
      images: images.filter(image => image.page === page.index), dividers: dividers.filter(divider => divider.page === page.index), columnSeparators: [] })) }
}
