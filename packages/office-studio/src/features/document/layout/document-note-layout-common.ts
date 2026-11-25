import type { DocumentModel, DocumentStoryContent } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { imageForBlock, tableForBlock, textTargets } from '../model/document-model'
import { migrateDocumentStoryBlocks } from '../model/document-story-blocks'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentModelForBlockHost } from '../services/document-block-host'
import { documentEffectiveSectionSetup } from '../services/document-settings'
import { paragraphListMarkers } from '../model/document-lists'
import { defaultDocumentNoteDecoration, type DocumentNoteDecorationKind } from '../notes/document-note-decorations'
import type { DocumentNoteLabel } from '../notes/document-note-numbering'
import type { DocumentNoteKind } from '../notes/document-note-types'
import { appendDocumentBlocks } from './document-block-layout'
import { DocumentSectionFlow } from './document-section-flow'
import { intersectTableRect, measureDocumentTable, translateTableLine, translateTableRect } from './document-table-measurement'
import { paragraphGeometry } from './document-paragraph-layout'
import { paragraphLineHeight, wrapRuns } from './document-text-layout'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutAuxiliaryContent, LayoutMeasurer, LayoutNoteDecoration, LayoutPage, LayoutRect } from './document-layout-types'

export function createDocumentFlowResult(flow: DocumentSectionFlow): DocumentLayout {
  return { width: 0, height: 0, pages: flow.pages, pageContent: [], lines: [], cells: [], images: [], dividers: [], columnSeparators: [],
    pageHeaderFooters: [], headings: [], auxiliaryContent: [], noteDecorations: [] }
}

export function documentNoteReferenceDisplay(labels: ReadonlyMap<string, DocumentNoteLabel>): DocumentTextLayoutDisplay {
  return { references: new Map([...labels.values()].map(label => [label.reference.run.id, {
    noteKind: label.reference.kind, noteId: label.reference.note.id, label: label.label, textStyle: label.reference.note.referenceTextStyle,
  }])) }
}

export function documentNoteBodyDisplay(model: DocumentModel, label: DocumentNoteLabel, references: DocumentTextLayoutDisplay): DocumentTextLayoutDisplay {
  const host: DocumentAuxiliaryBlockHost = { kind: 'note', noteKind: label.reference.kind, noteId: label.reference.note.id }
  const target = textTargets(documentModelForBlockHost(model, host))[0]
  return { ...references, leadingLabels: label.reference.note.showBodyReference === false || !target ? undefined : new Map([
    [target.cellBlockId ?? target.blockId, { noteKind: label.reference.kind, noteId: label.reference.note.id,
      label: label.label, textStyle: label.reference.note.referenceTextStyle }],
  ]) }
}

export function documentNoteFirstHeight(model: DocumentModel, label: DocumentNoteLabel, width: number, maximum: number, measurer: LayoutMeasurer,
  display: DocumentTextLayoutDisplay): number {
  const scoped = documentModelForBlockHost(model, { kind: 'note', noteKind: label.reference.kind, noteId: label.reference.note.id })
  const block = scoped.blocks[0]
  if (!block) return 1
  const markers = paragraphListMarkers(scoped.blocks), table = tableForBlock(block)
  if (table) {
    const measured = measureDocumentTable(block, width, markers, measurer, 0, maximum, display)
    const first = measured.indivisible.find(interval => interval.kind !== 'row')
    return block.style.spaceBefore + (first?.end ?? Math.min(measured.height, 1)) + Math.max(0, table.style?.offsetY ?? 0)
  }
  const image = imageForBlock(block)
  if (image) return block.style.spaceBefore + Math.min(Math.max(1, image.height), maximum, Math.max(1, image.height) * width / Math.max(1, image.width))
  if (block.type === 'divider') return block.style.spaceBefore + 18
  const geometry = paragraphGeometry(block, markers.get(block.id), width, measurer)
  const line = wrapRuns(block.runs, geometry.width, measurer, block.style.wordWrap === true, undefined, undefined,
    { references: display.references, leading: display.leadingLabels?.get(block.id) })[0]!
  return block.style.spaceBefore + paragraphLineHeight(Math.max(line.fontHeight, geometry.markerHeight), block.style)
}

/** Decorations use the same block kernel, without creating an ordinary note host. */
export class DocumentNoteDecorationLayout {
  private readonly cache = new Map<string, LayoutNoteDecoration | undefined>()
  constructor(private readonly model: DocumentModel, private readonly measurer: LayoutMeasurer) {}

  private content(kind: DocumentNoteKind, decoration: DocumentNoteDecorationKind): DocumentStoryContent | undefined {
    return this.model.noteSettings?.[kind]?.[decoration] ?? defaultDocumentNoteDecoration(this.model, kind, decoration)
  }

  /** Whole stories may exceed one page. The caller chooses a paginated flow instead of discarding or clipping their tail. */
  requiresPagination(kind: DocumentNoteKind, width: number, sectionId: string, available: number): boolean {
    const separator = this.measure(kind, 'separator', width, sectionId)?.height ?? 0
    const continuation = this.measure(kind, 'continuationSeparator', width, sectionId)?.height ?? 0
    const notice = this.measure(kind, 'continuationNotice', width, sectionId)?.height ?? 0
    return Math.max(separator, continuation) + notice >= Math.max(1, available) - 1e-6
  }

  /** All fragments retain the real decoration host and reuse the normal paragraph/table/image kernel. */
  append(flow: DocumentSectionFlow, result: DocumentLayout, kind: DocumentNoteKind, decoration: DocumentNoteDecorationKind, sectionId: string): LayoutNoteDecoration[] {
    const source = this.content(kind, decoration)
    if (!source) return []
    const host: DocumentAuxiliaryBlockHost = { kind: 'noteDecoration', noteKind: kind, decoration }
    const content = migrateDocumentStoryBlocks(cloneModel(source), [], host, documentDefaultStyle(this.model))
    const laidOut = createDocumentFlowResult(flow)
    appendDocumentBlocks(content.blocks!, flow, laidOut, this.measurer)
    const fragments = documentFlowContentFragments(laidOut, flow, host, sectionId).map(fragment => ({
      x: fragment.x, y: fragment.y, width: fragment.width, height: fragment.height, page: fragment.page,
      noteKind: kind, kind: decoration, continuedBefore: fragment.continuedBefore, continuedAfter: fragment.continuedAfter,
      lines: fragment.lines, cells: fragment.cells, images: fragment.images, dividers: fragment.dividers, tableFragments: fragment.tableFragments,
    }))
    result.noteDecorations ??= []
    result.noteDecorations.push(...fragments)
    return fragments
  }

  measure(kind: DocumentNoteKind, decoration: DocumentNoteDecorationKind, width: number, sectionId: string): LayoutNoteDecoration | undefined {
    const key = `${kind}:${decoration}:${width}:${sectionId}`
    if (this.cache.has(key)) return this.cache.get(key)
    const source = this.content(kind, decoration)
    if (!source) { this.cache.set(key, undefined); return undefined }
    const defaults = documentDefaultStyle(this.model)
    const content = migrateDocumentStoryBlocks(cloneModel(source), [], { kind: 'noteDecoration', noteKind: kind, decoration }, defaults)
    const setup = documentEffectiveSectionSetup(this.model, sectionId)
    const scoped: DocumentModel = { ...this.model, blocks: content.blocks!, sectionSetup: { columnCount: 1, columnGap: 0 },
      sectionHeaderFooters: undefined, footnotes: undefined, endnotes: undefined,
      page: { width, height: setup.page.height, marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0 } }
    const flow = new DocumentSectionFlow(scoped, { pageFrame: () => ({ index: 0, x: 0, y: 0, width, height: setup.page.height,
      content: { x: 0, y: 0, width, height: Number.POSITIVE_INFINITY }, sectionId, orientation: setup.orientation,
      pageNumber: 1, pageNumberStart: 1, sectionPageIndex: 0, isParityBlank: false, maximumContentHeight: Number.POSITIVE_INFINITY }) })
    const result = createDocumentFlowResult(flow)
    appendDocumentBlocks(scoped.blocks, flow, result, this.measurer)
    const measured: LayoutNoteDecoration = { x: 0, y: 0, width, height: Math.max(0, flow.cursor), page: 0, noteKind: kind, kind: decoration,
      lines: result.lines, cells: result.cells, images: result.images, dividers: result.dividers, tableFragments: result.tableFragments ?? [] }
    if (!Number.isFinite(measured.height)) throw new RangeError('注释分隔内容产生了无效的排版高度。')
    this.cache.set(key, measured)
    return measured
  }
}

/** Before-page hooks only repeat a separator/notice after actual note content made progress. */
export function documentNoteFlowHasContentOnPage(results: readonly DocumentLayout[], page: number): boolean {
  return results.some(result => [...result.lines, ...result.cells, ...result.images, ...result.dividers, ...(result.tableFragments ?? [])]
    .some(item => item.page === page))
}

export function translateDocumentNoteDecoration(content: LayoutNoteDecoration, dy: number, bounds: LayoutRect): LayoutNoteDecoration {
  const shifted = translateDocumentNoteContent({ ...content, host: { kind: 'noteDecoration', noteKind: content.noteKind, decoration: content.kind },
    sourceHost: { kind: 'noteDecoration', noteKind: content.noteKind, decoration: content.kind } }, dy, bounds)
  return { ...content, y: shifted.y, lines: shifted.lines, cells: shifted.cells, images: shifted.images,
    dividers: shifted.dividers, tableFragments: shifted.tableFragments }
}

export function positionDocumentNoteDecoration(source: LayoutNoteDecoration, page: LayoutPage, x: number, y: number): LayoutNoteDecoration {
  const clip = { ...page.content }
  const host: DocumentAuxiliaryBlockHost = { kind: 'noteDecoration', noteKind: source.noteKind, decoration: source.kind }
  const rect = <T extends LayoutRect & { clip?: LayoutRect }>(item: T): T => ({ ...translateTableRect(item, x, y),
    host: { ...host }, sectionId: page.sectionId,
    clip: item.clip ? intersectTableRect(translateTableRect(item.clip, x, y), clip) : { ...clip } })
  return { ...source, x, y, page: page.index, lines: source.lines.map(line => ({ ...translateTableLine(line, x, y), page: page.index,
    host: { ...host }, sectionId: page.sectionId,
    clip: line.clip ? intersectTableRect(translateTableRect(line.clip, x, y), clip) : { ...clip } })),
    cells: source.cells.map(item => ({ ...rect(item), page: page.index })), images: source.images.map(item => ({ ...rect(item), page: page.index })),
    dividers: source.dividers.map(item => ({ ...rect(item), page: page.index })), tableFragments: source.tableFragments.map(item => ({ ...rect(item),
      page: page.index, rows: item.rows.map(row => ({ ...row, y: row.y + y })) })) }
}

/** One fragment per physical column keeps input and table geometry in its actual block host. */
export function documentFlowContentFragments(result: DocumentLayout, flow: DocumentSectionFlow, host: DocumentAuxiliaryBlockHost,
  sectionId: string, sourceHost = host): LayoutAuxiliaryContent[] {
  const groups = new Map<string, { page: number; column: number; fragment?: number }>()
  for (const item of [...result.lines, ...result.cells, ...result.images, ...result.dividers, ...(result.tableFragments ?? [])]) {
    const column = item.column ?? 0
    groups.set(`${item.page}:${item.sectionFragment}:${column}`, { page: item.page, column, fragment: item.sectionFragment })
  }
  const entries = [...groups.values()].sort((a, b) => a.page - b.page || a.column - b.column)
  return entries.map(({ page: index, column: columnIndex, fragment }, order) => {
    const page = flow.pages.find(page => page.index === index)!
    const column = (page.sections.find(section => section.fragment === fragment)?.columns ?? page.sections[0]!.columns).find(column => column.index === columnIndex)!
    const belongs = (item: { page: number; column?: number; sectionFragment?: number }) => item.page === index && (item.column ?? 0) === columnIndex && item.sectionFragment === fragment
    const lines = result.lines.filter(belongs), cells = result.cells.filter(belongs), images = result.images.filter(belongs), dividers = result.dividers.filter(belongs)
    const tableFragments = (result.tableFragments ?? []).filter(belongs)
    const all = [...lines, ...cells, ...images, ...dividers, ...tableFragments], top = Math.min(...all.map(item => item.y)), bottom = Math.max(...all.map(item => item.y + item.height))
    const clip = { x: column.x, y: column.y, width: column.width, height: column.height }
    const tagged = <T extends LayoutRect & { clip?: LayoutRect }>(item: T) => ({ ...item, host: { ...host }, sectionId,
      clip: item.clip ? intersectTableRect(item.clip, clip) : { ...clip } })
    return { x: column.x, y: top, width: column.width, height: Math.max(1, bottom - top), page: index, host: { ...host }, sourceHost: { ...sourceHost },
      continuedBefore: order > 0, continuedAfter: order < entries.length - 1, columns: [{ ...column }], lines: lines.map(tagged), cells: cells.map(tagged),
      images: images.map(tagged), dividers: dividers.map(tagged), tableFragments: tableFragments.map(tagged) }
  })
}

/** Ordinary notes reuse the same host geometry as explicit independent inspection pages. */
export function documentNoteContentFragments(result: DocumentLayout, flow: DocumentSectionFlow, label: DocumentNoteLabel): LayoutAuxiliaryContent[] {
  return documentFlowContentFragments(result, flow,
    { kind: 'note', noteKind: label.reference.kind, noteId: label.reference.note.id }, label.reference.sectionId)
}

export function translateDocumentNoteContent(content: LayoutAuxiliaryContent, dy: number, bounds?: LayoutRect): LayoutAuxiliaryContent {
  const shifted = <T extends LayoutRect & { clip?: LayoutRect }>(item: T): T => ({ ...translateTableRect(item, 0, dy),
    ...(item.clip || bounds ? { clip: bounds ? item.clip ? intersectTableRect(translateTableRect(item.clip, 0, dy), bounds) : { ...bounds }
      : translateTableRect(item.clip!, 0, dy) } : {}) })
  return { ...content, y: content.y + dy, columns: content.columns?.map(column => shifted(column)),
    lines: content.lines.map(line => { const translated = translateTableLine(line, 0, dy); return { ...translated,
      ...(bounds ? { clip: translated.clip ? intersectTableRect(translated.clip, bounds) : { ...bounds } } : {}) } }), cells: content.cells.map(shifted), images: content.images.map(shifted),
    dividers: content.dividers.map(shifted), tableFragments: content.tableFragments.map(table => ({ ...shifted(table), rows: table.rows.map(row => ({ ...row, y: row.y + dy })) })) }
}

/** Repeated table headers do not allocate a second note or restart its numbering. */
export function documentLayoutReferencePages(layout: DocumentLayout): Map<string, number> {
  const result = new Map<string, number>()
  const lines = [...layout.lines, ...(layout.auxiliaryContent ?? []).flatMap(content => content.lines),
    ...layout.pageHeaderFooters.flatMap(page => [...page.headerLayoutLines, ...page.footerLayoutLines].map(line => ({ ...line, page: page.page })))].sort((a, b) => a.page - b.page)
  for (const line of lines) {
    if ('repeatedHeader' in line && line.repeatedHeader) continue
    for (const glyph of line.glyphs) if (glyph.note?.role === 'reference') {
      const key = `${glyph.note.noteKind}:${glyph.note.noteId}`
      if (!result.has(key)) result.set(key, line.page)
    }
  }
  return result
}
