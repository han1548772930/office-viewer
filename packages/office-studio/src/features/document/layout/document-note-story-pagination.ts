import type { DocumentNoteDecorationKind } from '../notes/document-note-decorations'
import { documentNoteFlowHasContentOnPage } from './document-note-layout-common'
import type { DocumentLayout, LayoutNoteDecoration } from './document-layout-types'
import type { DocumentSectionFlow } from './document-section-flow'

/** Repeat page decorations only after actual note geometry made progress, never during a decoration's own overflow. */
export class DocumentNoteStoryPagination {
  readonly noticePages = new Set<number>()
  readonly separatorKinds = new Map<number, 'separator' | 'continuationSeparator'>()
  private renderingDecoration = false
  private pendingContentPage?: number

  constructor(private readonly notes: readonly DocumentLayout[], private readonly appendDecoration:
    (flow: DocumentSectionFlow, key: DocumentNoteDecorationKind) => readonly LayoutNoteDecoration[]) {}

  appendInitialSeparator(flow: DocumentSectionFlow): void { this.append(flow, 'separator') }

  beforePageOverflow(flow: DocumentSectionFlow): void {
    if (this.renderingDecoration || !documentNoteFlowHasContentOnPage(this.notes, flow.page.index)) return
    const previous = flow.page.index
    this.pendingContentPage = previous
    this.append(flow, 'continuationNotice')
    if (flow.page.index !== previous) {
      this.pendingContentPage = undefined
      this.append(flow, 'continuationSeparator')
    }
  }

  onPageStart(flow: DocumentSectionFlow): void {
    if (this.renderingDecoration) return
    const previous = this.pendingContentPage
    this.pendingContentPage = undefined
    if (previous !== undefined && documentNoteFlowHasContentOnPage(this.notes, previous)) this.append(flow, 'continuationSeparator')
  }

  private append(flow: DocumentSectionFlow, key: DocumentNoteDecorationKind): void {
    this.renderingDecoration = true
    try {
      for (const content of this.appendDecoration(flow, key)) {
        if (key === 'continuationNotice') this.noticePages.add(content.page)
        else if (!this.separatorKinds.has(content.page)) this.separatorKinds.set(content.page, key)
      }
    } finally { this.renderingDecoration = false }
  }
}
