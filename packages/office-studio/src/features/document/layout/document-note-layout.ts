import type { DocumentModel } from '../../../domain/model'
import { resolveDocumentNoteLabels, type DocumentNoteLabel } from '../notes/document-note-numbering'
import { appendDocumentBlocks } from './document-block-layout'
import { DocumentSectionFlow } from './document-section-flow'
import { DocumentEndnoteLayout } from './document-endnote-layout'
import { layoutDocumentFootnotes } from './document-footnote-layout'
import { arrangeDocumentPages } from './document-page-arrangement'
import { layoutDocumentPageStories } from './document-page-stories'
import { createDocumentFlowResult, documentLayoutReferencePages, DocumentNoteDecorationLayout, documentNoteReferenceDisplay } from './document-note-layout-common'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutMeasurer } from './document-layout-types'

const MAX_NOTE_LAYOUT_PASSES = 64
const EPSILON = 1e-6

function sameNumberMap<Key>(left: ReadonlyMap<Key, number>, right: ReadonlyMap<Key, number>): boolean {
  return left.size === right.size && [...left].every(([key, value]) => {
    const candidate = right.get(key)
    return candidate !== undefined && Math.abs(candidate - value) <= EPSILON
  })
}
function sameSet(left: ReadonlySet<number>, right: ReadonlySet<number>): boolean {
  return left.size === right.size && [...left].every(value => right.has(value))
}
function sameSeparatorMap(left: ReadonlyMap<number, string>, right: ReadonlyMap<number, string>): boolean {
  return left.size === right.size && [...left].every(([key, value]) => right.get(key) === value)
}
function sameLabels(left: ReadonlyMap<string, DocumentNoteLabel>, right: ReadonlyMap<string, DocumentNoteLabel>): boolean {
  return left.size === right.size && [...left].every(([key, label]) => {
    const candidate = right.get(key)
    return candidate?.label === label.label && candidate.number === label.number && candidate.page === label.page
      && candidate.recognized === label.recognized
  })
}

/** Refresh linked story projections after continuation pages change NUMPAGES. */
function refreshStories(model: DocumentModel, result: DocumentLayout, flow: DocumentSectionFlow, measurer: LayoutMeasurer,
  display: DocumentTextLayoutDisplay): void {
  result.pageHeaderFooters = layoutDocumentPageStories(model, result.pages, flow.settings, measurer, display)
  const notes = (result.auxiliaryContent ?? []).filter(content => content.host.kind !== 'story')
  result.auxiliaryContent = [...notes, ...result.pageHeaderFooters.flatMap(page => [page.headerBlockLayout, page.footerBlockLayout]
    .filter((content): content is NonNullable<typeof content> => !!content))]
}

/**
 * References, page restarts and bottom reservations form a bounded fixed point.
 * Every pass is a derived view of one source model. No cached reference text,
 * note body, source offset or printed page number is mutated by pagination.
 */
export function layoutDocumentWithNotes(model: DocumentModel, measurer: LayoutMeasurer): DocumentLayout {
  const decorations = new DocumentNoteDecorationLayout(model, measurer)
  let referencePages = new Map<string, number>(), reservations = new Map<number, number>()
  let footnoteNoticePages = new Set<number>(), endnoteNoticePages = new Set<number>()
  let separatorKinds = new Map<number, 'separator' | 'continuationSeparator'>()
  const referenceFloors = new Map<string, number>()
  let output: DocumentLayout | undefined, stable = false
  for (let pass = 0; pass < MAX_NOTE_LAYOUT_PASSES; pass++) {
    const labels = resolveDocumentNoteLabels(model, referencePages), display = documentNoteReferenceDisplay(labels)
    const flow = new DocumentSectionFlow(model, { reservedPageBottom: reservations, referencePageFloors: referenceFloors })
    const result = createDocumentFlowResult(flow)
    const endnotes = new DocumentEndnoteLayout(model, measurer, labels, display, decorations, endnoteNoticePages, referencePages)
    appendDocumentBlocks(model.blocks, flow, result, measurer, { textDisplay: display,
      beforeSectionBreak: () => endnotes.finishSection(flow, result, flow.location.sectionId, false) })
    endnotes.finishSection(flow, result, flow.location.sectionId, true)
    refreshStories(model, result, flow, measurer, display)
    const reached = documentLayoutReferencePages(result)
    // A reference inside a note becomes visible only after its parent was laid
    // out. Carry these addresses across passes, never invent an orphan region.
    for (const [key, label] of labels) if (label.reference.host?.kind === 'note' && !reached.has(key) && referencePages.has(key)) {
      reached.set(key, referencePages.get(key)!)
    }
    const footnotes = layoutDocumentFootnotes(model, flow, result, measurer, labels, display, reached, decorations,
      footnoteNoticePages, separatorKinds)
    refreshStories(model, result, flow, measurer, display)
    const actualPages = documentLayoutReferencePages(result), actualLabels = resolveDocumentNoteLabels(model, actualPages)
    let floorsChanged = false
    for (const [key, label] of labels) {
      const page = actualPages.get(key)
      if (label.reference.host || page === undefined) continue
      const previous = referenceFloors.get(label.reference.run.id) ?? 0
      if (page > previous) { referenceFloors.set(label.reference.run.id, page); floorsChanged = true }
    }
    stable = labels.size === 0 || !floorsChanged && sameLabels(labels, actualLabels) && sameNumberMap(referencePages, actualPages)
      && sameNumberMap(reservations, footnotes.reservations) && sameSet(footnoteNoticePages, footnotes.noticePages)
      && sameSet(endnoteNoticePages, endnotes.noticePages) && sameSeparatorMap(separatorKinds, footnotes.separatorKinds)
    result.noteLabels = labels
    output = result
    if (stable) break
    referencePages = actualPages; reservations = footnotes.reservations
    footnoteNoticePages = footnotes.noticePages; endnoteNoticePages = endnotes.noticePages; separatorKinds = footnotes.separatorKinds
  }
  if (!output) throw new RangeError('注释排版未产生文档页面。')
  if (!stable) {
    output.diagnostics ??= []
    output.diagnostics.push({ code: 'DOCUMENT_NOTE_LAYOUT_NOT_CONVERGED',
      message: `注释分页在 ${MAX_NOTE_LAYOUT_PASSES} 次派生计算内未收敛；当前显示保留完整源内容，需要检查注释高度和引用位置。` })
  }
  for (const label of output.noteLabels?.values() ?? []) if (!label.recognized) {
    output.diagnostics ??= []
    output.diagnostics.push({ code: 'DOCUMENT_NOTE_NUMBER_FORMAT_FALLBACK', noteKind: label.reference.kind, noteId: label.reference.note.id,
      message: `未识别注释编号格式 ${label.settings.numberFormat}，当前显示十进制；原始编号设置保留。` })
  }
  arrangeDocumentPages(output)
  return output
}
