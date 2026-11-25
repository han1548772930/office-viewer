import type { DocumentModel } from '../../../domain/model'
import { documentModelForBlockHost } from '../services/document-block-host'
import { documentEffectiveSectionSetup } from '../services/document-settings'
import type { DocumentNoteLabel } from '../notes/document-note-numbering'
import { appendDocumentBlocks } from './document-block-layout'
import { DocumentSectionFlow, type DocumentFlowPageFrame } from './document-section-flow'
import { DocumentNoteStoryPagination } from './document-note-story-pagination'
import { createDocumentFlowResult, documentNoteBodyDisplay, documentNoteContentFragments, translateDocumentNoteContent,
  translateDocumentNoteDecoration, type DocumentNoteDecorationLayout } from './document-note-layout-common'
import type { DocumentFootnoteLayoutResult } from './document-footnote-layout'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutAuxiliaryContent, LayoutMeasurer, LayoutPage } from './document-layout-types'

interface FootnoteStoryFrame { physical: LayoutPage; frame: DocumentFlowPageFrame; minimumBody: number }
interface FootnoteStoryStream {
  flow: DocumentSectionFlow; frames: Map<number, FootnoteStoryFrame>; first: DocumentNoteLabel
  content: LayoutAuxiliaryContent[]; rawNotes: DocumentLayout[]; decorations: DocumentLayout; pagination: DocumentNoteStoryPagination
}

/** Oversized separator/notice stories consume the actual shared block flow; no measured tail is discarded. */
export function layoutDocumentPaginatedFootnotes(model: DocumentModel, bodyFlow: DocumentSectionFlow, result: DocumentLayout,
  measurer: LayoutMeasurer, labels: ReadonlyMap<string, DocumentNoteLabel>, display: DocumentTextLayoutDisplay,
  ordered: readonly { label: DocumentNoteLabel; page: number }[], minimumBody: ReadonlyMap<number, number>,
  decorations: DocumentNoteDecorationLayout): DocumentFootnoteLayoutResult {
  const output: DocumentFootnoteLayoutResult = { reservations: new Map(), noticePages: new Set(), separatorKinds: new Map() }
  const streams: FootnoteStoryStream[] = []
  let current: FootnoteStoryStream | undefined
  const physicalPage = (index: number): LayoutPage => {
    while (bodyFlow.pages.length <= index) bodyFlow.appendAuxiliaryPage()
    return bodyFlow.pages[index]!
  }
  for (const { label, page: referencePage } of ordered) {
    if (!current || referencePage > current.flow.page.index) {
      const frames = new Map<number, FootnoteStoryFrame>(), indices = new Map<number, number>()
      const section = documentEffectiveSectionSetup(model, label.reference.sectionId)
      const frame = (local: number): DocumentFlowPageFrame => {
        let index = local === 0 ? referencePage : indices.get(local - 1)! + 1
        while (physicalPage(index).isParityBlank) index++
        indices.set(local, index)
        const page = physicalPage(index), settings = documentEffectiveSectionSetup(model, page.sectionId)
        const count = label.settings.columnCount === 0 ? settings.columnCount : label.settings.columnCount
        const gap = count === settings.columnCount ? settings.columnGap : Math.min(48, page.content.width / (Math.max(1, count) * 2))
        const body = Math.min(Math.max(0, page.content.height - 1), minimumBody.get(index) ?? 0)
        const prepared: DocumentFlowPageFrame = { ...page, content: { ...page.content, height: Math.max(0, page.content.height - body) },
          columnCount: Math.max(1, count), columnGap: gap, maximumContentHeight: page.content.height }
        frames.set(index, { physical: page, frame: prepared, minimumBody: body })
        return prepared
      }
      const rawNotes: DocumentLayout[] = []
      const scoped: DocumentModel = { ...model, blocks: [], sectionSetup: { columnCount: section.columnCount, columnGap: section.columnGap },
        sectionHeaderFooters: undefined, footnotes: undefined, endnotes: undefined }
      let pagination: DocumentNoteStoryPagination | undefined
      const flow = new DocumentSectionFlow(scoped, { pageFrame: frame,
        beforePageOverflow: target => pagination?.beforePageOverflow(target), onPageStart: target => pagination?.onPageStart(target) })
      const decorationResults = createDocumentFlowResult(flow)
      pagination = new DocumentNoteStoryPagination(rawNotes, (target, key) => {
        target.beginAuxiliarySegment({})
        try { return decorations.append(target, decorationResults, 'footnote', key, target.page.sectionId) }
        finally { target.endAuxiliarySegment() }
      })
      current = { flow, frames, first: label, rawNotes, content: [], decorations: decorationResults, pagination }
      streams.push(current)
      pagination.appendInitialSeparator(flow)
    }
    const scoped = documentModelForBlockHost(model, { kind: 'note', noteKind: 'footnote', noteId: label.reference.note.id })
    const note = createDocumentFlowResult(current.flow)
    // Register before append so an overflow hook observes the note's already placed table/paragraph fragments.
    current.rawNotes.push(note)
    appendDocumentBlocks(scoped.blocks, current.flow, note, measurer, { textDisplay: documentNoteBodyDisplay(model, label, display) })
    current.content.push(...documentNoteContentFragments(note, current.flow, label))
  }
  for (const stream of streams) {
    const decorationContent = stream.decorations.noteDecorations ?? []
    const pages = [...new Set([...stream.content.map(content => content.page), ...decorationContent.map(content => content.page)])].sort((a, b) => a - b)
    for (const pageIndex of pages) {
      const prepared = stream.frames.get(pageIndex)!, page = prepared.physical
      const content = stream.content.filter(content => content.page === pageIndex), decorationFragments = decorationContent.filter(content => content.page === pageIndex)
      const local = stream.flow.pages.find(page => page.index === pageIndex)!
      const used = Math.max(0, ...local.sections.map(section => section.y + section.usedHeight - prepared.frame.content.y))
      const reservation = Math.max(0, Math.min(page.content.height - prepared.minimumBody, used))
      output.reservations.set(pageIndex, reservation)
      const first = content[0], position = first?.host.kind === 'note'
        ? labels.get(`footnote:${first.host.noteId}`)?.settings.position ?? stream.first.settings.position : stream.first.settings.position
      const bodyBottom = Math.max(page.content.y, ...page.sections.map(section => section.y + section.usedHeight))
      const bottomTop = page.content.y + page.content.height - reservation
      const top = position === 'beneathText' ? Math.min(bodyBottom, bottomTop) : bottomTop
      const dy = top - prepared.frame.content.y, bounds = { x: page.content.x, y: top, width: page.content.width, height: reservation }
      result.auxiliaryContent!.push(...content.map(content => translateDocumentNoteContent(content, dy, bounds)))
      result.noteDecorations!.push(...decorationFragments.map(content => translateDocumentNoteDecoration(content, dy, bounds)))
      if (used > reservation + 1e-6 && [...content, ...decorationFragments].some(content => content.lines.some(line =>
        line.y + line.height > prepared.frame.content.y + prepared.frame.content.height + 1e-6))) {
        result.diagnostics ??= []
        result.diagnostics.push({ code: 'DOCUMENT_NOTE_OVERSIZE', message: '注释故事中有单个不可拆内容超过物理页可用高度；源数据和完整排版对象已保留。', noteKind: 'footnote',
          noteId: first?.host.kind === 'note' ? first.host.noteId : undefined })
      }
    }
    for (const page of stream.pagination.noticePages) output.noticePages.add(page)
    for (const [page, kind] of stream.pagination.separatorKinds) output.separatorKinds.set(page, kind)
  }
  result.diagnostics ??= []
  result.diagnostics.push({ code: 'DOCUMENT_NOTE_DECORATION_PAGINATED', noteKind: 'footnote',
    message: '超页脚注分隔故事按真实块流跨页显示；最终 Office 续页位置尚待文件样本验收。' })
  return output
}
