import type { DocumentModel, DocumentParagraphStyle } from '../../domain/model'
import { blockAtPoint } from '../../features/document/commands/document-commands'
import {
  marksAtPoint, orderedSelection, runsForTarget, sameTarget, selectedPlainText,
  selectionCollapsed, textTargets, type DocumentRun, type DocumentSelection, type RunPatch,
} from '../../features/document/model/document-model'

const markKeys = ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'color', 'highlight', 'href'] as const
const paragraphKeys = ['align', 'indent', 'lineHeight', 'spaceBefore', 'spaceAfter'] as const

export interface DocumentToolbarState {
  marks: RunPatch
  mixedMarks: readonly (keyof RunPatch)[]
  paragraph: Partial<DocumentParagraphStyle>
  blockStyle?: string
  hasSelection: boolean
}

/** Resolve the selected runs, rather than echoing the last clicked toolbar value. */
export function documentToolbarState(model: DocumentModel, selection: DocumentSelection, pending: RunPatch = {}): DocumentToolbarState {
  const collapsed = selectionCollapsed(selection)
  const { start, end } = orderedSelection(model, selection)
  const targets = textTargets(model)
  const first = targets.findIndex((target) => sameTarget(target, start))
  const last = targets.findIndex((target) => sameTarget(target, end))
  const selected = targets.slice(Math.max(0, first), last + 1)
  const marks: RunPatch[] = []
  const paragraphs: DocumentParagraphStyle[] = []
  const styles: string[] = []
  for (const [index, target] of selected.entries()) {
    const runs = runsForTarget(model, target)
    const from = index === 0 ? start.offset : 0
    const to = index === selected.length - 1 ? end.offset : Number.POSITIVE_INFINITY
    const included: DocumentRun[] = []
    let offset = 0
    for (const run of runs) {
      if (!collapsed && offset < to && offset + run.text.length > from) included.push(run)
      offset += run.text.length
    }
    // A range ending at the next paragraph's offset zero does not select that paragraph.
    if (!collapsed && index > 0 && to === 0) continue
    marks.push(...included)
    const block = blockAtPoint(model, target)
    if (block) {
      paragraphs.push(block.style)
      styles.push(block.type === 'heading' ? `H${block.level}` : block.type === 'paragraph' ? 'P' : block.type)
    }
  }
  if (collapsed || !marks.length) marks.push({ ...marksAtPoint(model, selection.focus), ...pending })
  const uniform: RunPatch = {}
  const mixedMarks: (keyof RunPatch)[] = []
  for (const key of markKeys) {
    const firstValue = marks[0]?.[key]
    if (marks.every((mark) => mark[key] === firstValue)) Object.assign(uniform, { [key]: firstValue })
    else mixedMarks.push(key)
  }
  const paragraph: Partial<DocumentParagraphStyle> = {}
  for (const key of paragraphKeys) {
    const firstValue = paragraphs[0]?.[key]
    if (paragraphs.every((style) => style[key] === firstValue)) Object.assign(paragraph, { [key]: firstValue })
  }
  return { marks: uniform, mixedMarks, paragraph, blockStyle: styles.every((style) => style === styles[0]) ? styles[0] : undefined, hasSelection: !collapsed }
}

/** Count displayed characters as grapheme clusters, including composed emoji only once. */
export function documentCharacterCount(text: string): number {
  let count = 0
  for (const { segment } of new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)) {
    if (!/^\s+$/u.test(segment)) count++
  }
  return count
}

export function documentSelectionCharacterCount(model: DocumentModel, selection: DocumentSelection): number {
  return selectionCollapsed(selection) ? 0 : documentCharacterCount(selectedPlainText(model, selection))
}

export function clampDocumentZoom(value: number): number {
  return Number.isFinite(value) ? Math.max(10, Math.min(400, Math.round(value))) : 100
}
