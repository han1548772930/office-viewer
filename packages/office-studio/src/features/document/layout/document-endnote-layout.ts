import type { DocumentModel } from '../../../domain/model'
import { documentModelForBlockHost } from '../services/document-block-host'
import { documentSectionIds } from '../services/section-header-footer'
import { documentSectionSuppressesEndnotes } from '../notes/document-note-settings'
import type { DocumentNoteLabel } from '../notes/document-note-numbering'
import { appendDocumentBlocks } from './document-block-layout'
import { createDocumentFlowResult, documentNoteBodyDisplay, documentNoteContentFragments, documentNoteFirstHeight,
  positionDocumentNoteDecoration, type DocumentNoteDecorationLayout } from './document-note-layout-common'
import type { DocumentSectionFlow, DocumentAuxiliaryFlowSegment } from './document-section-flow'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutMeasurer, LayoutPage } from './document-layout-types'

/** Endnote stories occupy the same physical flow that later sections and body text use. */
export class DocumentEndnoteLayout {
  readonly noticePages = new Set<number>()
  private readonly placed = new Set<string>()
  private readonly sections: string[]
  constructor(private readonly model: DocumentModel, private readonly measurer: LayoutMeasurer,
    private readonly labels: ReadonlyMap<string, DocumentNoteLabel>, private readonly display: DocumentTextLayoutDisplay,
    private readonly decorations: DocumentNoteDecorationLayout, private readonly reservedNoticePages: ReadonlySet<number>,
    private readonly referencePages: ReadonlyMap<string, number> = new Map()) {
    this.sections = documentSectionIds(model)
  }

  finishSection(flow: DocumentSectionFlow, result: DocumentLayout, sectionId: string, final: boolean): void {
    const sectionIndex = this.sections.indexOf(sectionId)
    const references = [...this.labels.values()].filter(label => label.reference.kind === 'endnote' && !this.placed.has(label.reference.note.id)
      && (!label.reference.host || this.referencePages.has(`endnote:${label.reference.note.id}`))
      && (final || label.settings.position === 'sectEnd' && !documentSectionSuppressesEndnotes(this.model, sectionId)
        && this.sections.indexOf(label.reference.sectionId) <= sectionIndex))
    if (!references.length) return
    const separatorPages = new Set<number>(), groupPages = new Set<number>(), prefix = new Map<number, number>()
    let firstPage: number | undefined
    const measure = (page: LayoutPage, key: 'separator' | 'continuationSeparator' | 'continuationNotice') =>
      this.decorations.measure('endnote', key, page.content.width, sectionId)
    const segment: DocumentAuxiliaryFlowSegment = {
      reservedBottom: page => this.reservedNoticePages.has(page.index) ? measure(page, 'continuationNotice')?.height ?? 0 : 0,
      prefixHeight: page => prefix.get(page.index) ?? 0,
    }
    flow.beginAuxiliarySegment(segment)
    const initial = measure(flow.page, 'separator'), firstDisplay = documentNoteBodyDisplay(this.model, references[0]!, this.display)
    const minimum = documentNoteFirstHeight(this.model, references[0]!, flow.page.content.width,
      Math.max(1, flow.fullPageContentHeight - (initial?.height ?? 0)), this.measurer, firstDisplay)
    flow.ensureRoom((initial?.height ?? 0) + minimum)
    firstPage = flow.page.index
    const appendSeparator = (target: DocumentSectionFlow) => {
      if (separatorPages.has(target.page.index)) return
      const key = target.page.index === firstPage ? 'separator' : 'continuationSeparator', decoration = measure(target.page, key)
      separatorPages.add(target.page.index)
      prefix.set(target.page.index, decoration?.height ?? 0)
      if (decoration) {
        result.noteDecorations!.push(positionDocumentNoteDecoration(decoration, target.page, target.column.x, target.cursor))
        target.consume(decoration.height)
      }
    }
    appendSeparator(flow)
    segment.onPageStart = appendSeparator
    for (const label of references) {
      const scoped = documentModelForBlockHost(this.model, { kind: 'note', noteKind: 'endnote', noteId: label.reference.note.id })
      const body = createDocumentFlowResult(flow)
      appendDocumentBlocks(scoped.blocks, flow, body, this.measurer, { textDisplay: documentNoteBodyDisplay(this.model, label, this.display) })
      const fragments = documentNoteContentFragments(body, flow, label)
      result.auxiliaryContent!.push(...fragments)
      for (const content of fragments) groupPages.add(content.page)
      this.placed.add(label.reference.note.id)
    }
    const last = Math.max(...groupPages)
    for (const pageIndex of groupPages) {
      if (pageIndex >= last) continue
      this.noticePages.add(pageIndex)
      const page = flow.pages[pageIndex]!, notice = measure(page, 'continuationNotice')
      if (!notice) continue
      const contents = result.auxiliaryContent!.filter(content => content.page === pageIndex && content.host.kind === 'note' && content.host.noteKind === 'endnote')
      const bottom = Math.max(page.content.y, ...contents.map(content => content.y + content.height))
      result.noteDecorations!.push(positionDocumentNoteDecoration(notice, page, page.content.x, bottom))
    }
    flow.endAuxiliarySegment()
  }
}
