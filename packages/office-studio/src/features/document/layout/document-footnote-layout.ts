import type { DocumentModel } from '../../../domain/model'
import { documentModelForBlockHost } from '../services/document-block-host'
import { documentEffectiveSectionSetup } from '../services/document-settings'
import type { DocumentNoteLabel } from '../notes/document-note-numbering'
import { appendDocumentBlocks } from './document-block-layout'
import { DocumentSectionFlow, type DocumentFlowPageFrame } from './document-section-flow'
import { createDocumentFlowResult, documentNoteBodyDisplay, documentNoteContentFragments, positionDocumentNoteDecoration,
  translateDocumentNoteContent, type DocumentNoteDecorationLayout } from './document-note-layout-common'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutAuxiliaryContent, LayoutMeasurer, LayoutNoteDecoration, LayoutPage } from './document-layout-types'

export interface DocumentFootnoteLayoutResult {
  reservations: Map<number, number>; noticePages: Set<number>; separatorKinds: Map<number, 'separator' | 'continuationSeparator'>
}
interface FootnoteFrame {
  physical: LayoutPage; frame: DocumentFlowPageFrame; separator?: LayoutNoteDecoration; notice?: LayoutNoteDecoration; minimumBody: number
}
interface FootnoteStream { flow: DocumentSectionFlow; frames: Map<number, FootnoteFrame>; content: LayoutAuxiliaryContent[]; first: DocumentNoteLabel }

/** Per-layout streams reserve physical pages and share paragraph/table/image pagination with body flow. */
export function layoutDocumentFootnotes(model: DocumentModel, bodyFlow: DocumentSectionFlow, result: DocumentLayout, measurer: LayoutMeasurer,
  labels: ReadonlyMap<string, DocumentNoteLabel>, display: DocumentTextLayoutDisplay, referencePages: ReadonlyMap<string, number>,
  decorations: DocumentNoteDecorationLayout, reservedNoticePages: ReadonlySet<number>, reservedSeparatorKinds: ReadonlyMap<number, 'separator' | 'continuationSeparator'>): DocumentFootnoteLayoutResult {
  const output: DocumentFootnoteLayoutResult = { reservations: new Map(), noticePages: new Set(), separatorKinds: new Map() }
  const ordered = [...labels.values()].filter(label => label.reference.kind === 'footnote' && referencePages.has(`footnote:${label.reference.note.id}`))
    .map((label, order) => ({ label, order, page: referencePages.get(`footnote:${label.reference.note.id}`)! }))
    .sort((a, b) => a.page - b.page || a.order - b.order)
  if (!ordered.length) return output
  const minimumBody = new Map<number, number>()
  for (const line of [...result.lines, ...(result.auxiliaryContent ?? []).flatMap(content => content.lines)]) {
    if (line.repeatedHeader || !line.glyphs.some(glyph => glyph.note?.role === 'reference' && glyph.note.noteKind === 'footnote')) continue
    minimumBody.set(line.page, Math.max(minimumBody.get(line.page) ?? 0, line.height))
  }
  const streams: FootnoteStream[] = []
  let current: FootnoteStream | undefined
  const physicalPage = (index: number) => {
    while (bodyFlow.pages.length <= index) {
      if (bodyFlow.pages.length >= 40_000) throw new RangeError('注释排版超过单次物理页预算。')
      bodyFlow.appendAuxiliaryPage()
    }
    return bodyFlow.pages[index]!
  }
  for (const { label, page: referencePage } of ordered) {
    if (!current || referencePage > current.flow.page.index) {
      const frames = new Map<number, FootnoteFrame>(), indices = new Map<number, number>(), section = documentEffectiveSectionSetup(model, label.reference.sectionId)
      const frame = (local: number): DocumentFlowPageFrame => {
        let physical = local === 0 ? referencePage : indices.get(local - 1)! + 1
        while (physicalPage(physical).isParityBlank) physical++
        indices.set(local, physical)
        const page = physicalPage(physical), settings = documentEffectiveSectionSetup(model, page.sectionId)
        const count = label.settings.columnCount === 0 ? settings.columnCount : label.settings.columnCount
        const gap = count === settings.columnCount ? settings.columnGap : Math.min(48, page.content.width / (Math.max(1, count) * 2))
        const separatorKind = reservedSeparatorKinds.get(physical) ?? (local === 0 ? 'separator' : 'continuationSeparator')
        const separator = decorations.measure('footnote', separatorKind, page.content.width, page.sectionId)
        const notice = reservedNoticePages.has(physical) ? decorations.measure('footnote', 'continuationNotice', page.content.width, page.sectionId) : undefined
        const prefix = separator?.height ?? 0, suffix = notice?.height ?? 0
        const maximum = page.content.height - prefix - suffix
        if (maximum <= 0) throw new RangeError('脚注分隔内容和续页提示需要留下正的注释高度。')
        const body = Math.min(Math.max(0, page.content.height - 1), minimumBody.get(physical) ?? 0)
        const content = { x: page.content.x, y: page.content.y + prefix, width: page.content.width, height: Math.max(0, maximum - body) }
        const prepared: DocumentFlowPageFrame = { ...page, content, columnCount: Math.max(1, count), columnGap: gap, maximumContentHeight: maximum }
        frames.set(physical, { physical: page, frame: prepared, separator, notice, minimumBody: body })
        return prepared
      }
      const scoped: DocumentModel = { ...model, blocks: [], sectionSetup: { columnCount: section.columnCount, columnGap: section.columnGap },
        sectionHeaderFooters: undefined, footnotes: undefined, endnotes: undefined }
      current = { flow: new DocumentSectionFlow(scoped, { pageFrame: frame }), frames, content: [], first: label }
      streams.push(current)
    }
    const host = { kind: 'note' as const, noteKind: 'footnote' as const, noteId: label.reference.note.id }
    const scoped = documentModelForBlockHost(model, host), note = createDocumentFlowResult(current.flow)
    appendDocumentBlocks(scoped.blocks, current.flow, note, measurer, { textDisplay: documentNoteBodyDisplay(model, label, display) })
    current.content.push(...documentNoteContentFragments(note, current.flow, label))
  }
  for (const stream of streams) {
    const usedPages = [...new Set(stream.content.map(content => content.page))].sort((a, b) => a - b), last = usedPages.at(-1)
    for (const index of usedPages) {
      const prepared = stream.frames.get(index)!, page = prepared.physical, content = stream.content.filter(content => content.page === index)
      const local = stream.flow.pages.find(page => page.index === index)!
      const used = Math.max(0, ...local.sections.map(section => section.usedHeight))
      const firstContent = content[0]!, kind = firstContent.continuedBefore ? 'continuationSeparator' : 'separator'
      output.separatorKinds.set(index, kind)
      const separator = decorations.measure('footnote', kind, page.content.width, page.sectionId)
      const notice = index === last ? undefined : decorations.measure('footnote', 'continuationNotice', page.content.width, page.sectionId)
      if (index !== last) output.noticePages.add(index)
      const total = (separator?.height ?? 0) + used + (notice?.height ?? 0)
      const reservation = Math.min(page.content.height - prepared.minimumBody, total)
      output.reservations.set(index, Math.max(0, reservation))
      const bodyBottom = Math.max(page.content.y, ...page.sections.map(section => section.y + section.usedHeight))
      const position = labels.get(`footnote:${firstContent.host.kind === 'note' ? firstContent.host.noteId : ''}`)?.settings.position ?? stream.first.settings.position
      const top = (position === 'beneathText' ? bodyBottom : page.content.y + page.content.height - reservation) + (separator?.height ?? 0)
      const dy = top - prepared.frame.content.y, bounds = { x: page.content.x, y: top, width: page.content.width, height: Math.max(0, reservation - (separator?.height ?? 0) - (notice?.height ?? 0)) }
      result.auxiliaryContent!.push(...content.map(content => translateDocumentNoteContent(content, dy, bounds)))
      if (separator) result.noteDecorations!.push(positionDocumentNoteDecoration(separator, page, page.content.x, top - separator.height))
      if (notice) result.noteDecorations!.push(positionDocumentNoteDecoration(notice, page, page.content.x, top + bounds.height))
      if (total > reservation + 1e-6 && content.some(content => content.lines.some(line => line.y + line.height > prepared.frame.content.y + prepared.frame.content.height + 1e-6))) {
        result.diagnostics ??= []
        result.diagnostics.push({ code: 'DOCUMENT_NOTE_OVERSIZE', message: '注释中有超出物理页可用高度的内容；完整源内容已保留。', noteKind: 'footnote',
          noteId: firstContent.host.kind === 'note' ? firstContent.host.noteId : undefined })
      }
    }
  }
  return output
}
