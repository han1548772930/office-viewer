import type { DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { paragraphListMarkers } from '../model/document-lists'
import { documentFieldAutomaticRefreshBlocked, documentTextFieldParagraphs, documentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit } from '../model/document-text-fields'
import { documentModelForBlockHost, resolveDocumentBlockHost } from '../services/document-block-host'
import { documentDefaultStyle } from '../model/document-default-style'
import { remapDocumentStoryFieldRuns } from '../model/document-story'
import { sliceRuns } from '../model/document-model'
import { splitDocumentRunRange } from '../model/document-run-range'
import { intersectTableRect, measureDocumentTable, translateTableLine, translateTableRect } from './document-table-measurement'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutAuxiliaryContent, LayoutLine, LayoutMeasurer, LayoutRect } from './document-layout-types'
import { measureDocumentBlockCollections } from './document-block-measurement'

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

/** Auxiliary hosts share real structure/paragraph/table geometry with table-cell content. */
export function layoutDocumentAuxiliaryBlocks(model: DocumentModel, host: DocumentAuxiliaryBlockHost, bounds: LayoutRect, page: number,
  pageNumber: number, total: number, measurer: LayoutMeasurer, bottom = false, noteDisplay?: DocumentTextLayoutDisplay): LayoutAuxiliaryContent {
  const view = resolveDocumentBlockHost(model, host), scoped = documentModelForBlockHost(model, host)
  const display = host.kind === 'story' ? storyPageDisplay(scoped, pageNumber, total) : { model: scoped, offsets: new Map<string, FieldDisplayMap[]>() }
  const markers = paragraphListMarkers(display.model.blocks)
  const measured = measureDocumentBlockCollections([display.model.blocks], { width: bounds.width, markers, measurer,
    measureTable: measureDocumentTable, maximumObjectHeight: bounds.height, display: noteDisplay })
  const output: LayoutAuxiliaryContent = { ...bounds, host: { ...host }, sourceHost: view.sourceHost, page,
    lines: measured.lines, cells: measured.cells, images: measured.images, dividers: measured.dividers, containers: measured.containers,
    tableFragments: measured.tables.map(table => ({ ...table, page, sourceStart: 0, sourceEnd: table.height,
      repeatedHeader: false, continuedBefore: false, continuedAfter: false })) }
  const top = bottom ? bounds.y + bounds.height - measured.height : bounds.y, clip = { ...bounds }
  output.y = top; output.height = measured.height
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
  output.containers = output.containers?.map(positioned)
  output.tableFragments = output.tableFragments.map(fragment => ({ ...positioned(fragment), rows: fragment.rows.map(row => ({ ...row, y: row.y + top })) }))
  return output
}

/** A view of the same physical pages for caret, cell hit-testing, and table overlays in one host. */
export function documentLayoutForBlockHost(layout: DocumentLayout, host?: DocumentAuxiliaryBlockHost): DocumentLayout {
  if (!host) return layout
  const content = (layout.auxiliaryContent ?? []).filter(content => sameDocumentBlockHost(content.host, host))
  const lines = content.flatMap(content => content.lines), cells = content.flatMap(content => content.cells), images = content.flatMap(content => content.images)
  const dividers = content.flatMap(content => content.dividers), tableFragments = content.flatMap(content => content.tableFragments)
  const containers = content.flatMap(content => content.containers ?? [])
  return { ...layout, lines, cells, images, dividers, columnSeparators: [], tableFragments, containers, headings: [], auxiliaryContent: content, noteDecorations: [],
    pageContent: layout.pages.map(page => ({ lines: lines.filter(line => line.page === page.index), cells: cells.filter(cell => cell.page === page.index),
      images: images.filter(image => image.page === page.index), dividers: dividers.filter(divider => divider.page === page.index),
      containers: containers.filter(container => container.page === page.index), columnSeparators: [] })) }
}
