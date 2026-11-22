import type { DocumentModel, DocumentParagraphStyle } from '../../domain/model'
import { caretUsesEastAsianFont, hasEastAsianLetters } from '../../features/document/model/document-font'
import { blockListPreset, type ListPreset } from '../../features/document/model/document-lists'
import { blockAtPoint } from '../../features/document/commands/document-commands'
import { documentTableSelection, type DocumentTableSelection } from '../../features/document/commands/table-commands'
import { DOCUMENT_PARAGRAPH_SETTING_FIELDS, documentParagraphSettingsStyle, type DocumentParagraphSettingField } from './document-paragraph-settings-state'
import {
  DOCUMENT_NAMED_PARAGRAPH_STYLES, marksAtPoint, orderedSelection, runsForTarget, sameTarget, selectedPlainText,
  selectionCollapsed, textTargets, type DocumentRun, type DocumentSelection, type RunPatch,
} from '../../features/document/model/document-model'

const markKeys = ['bold', 'italic', 'underline', 'strike', 'verticalAlign', 'fontFamily', 'fontSize', 'color', 'highlight', 'href'] as const

export interface DocumentToolbarState {
  marks: RunPatch
  mixedMarks: readonly (keyof RunPatch)[]
  paragraph: Partial<DocumentParagraphStyle>
  mixedParagraph?: readonly DocumentParagraphSettingField[]
  blockStyle?: string
  listPreset?: ListPreset
  hasSelection: boolean
  table?: DocumentTableSelection
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
  const presets: (ListPreset | undefined)[] = []
  for (const [index, target] of selected.entries()) {
    const runs = runsForTarget(model, target)
    const from = index === 0 ? start.offset : 0
    const to = index === selected.length - 1 ? end.offset : Number.POSITIVE_INFINITY
    const included: DocumentRun[] = []
    let offset = 0
    for (const run of runs) {
      if (!collapsed && offset < to && offset + run.text.length > from) {
        const eastAsian = run.eastAsiaFontFamily && hasEastAsianLetters(run.text.slice(Math.max(0, from - offset), to - offset))
        included.push(eastAsian ? { ...run, fontFamily: run.eastAsiaFontFamily! } : run)
      }
      offset += run.text.length
    }
    // A range ending at the next paragraph's offset zero does not select that paragraph.
    if (!collapsed && index > 0 && to === 0) continue
    marks.push(...included)
    const block = blockAtPoint(model, target)
    if (block) {
      presets.push(blockListPreset(block))
      paragraphs.push(documentParagraphSettingsStyle(block.style))
      const namedStyle = block.metadata.namedStyleType
      if (typeof namedStyle === 'string' && Object.hasOwn(DOCUMENT_NAMED_PARAGRAPH_STYLES, namedStyle)) styles.push(namedStyle.startsWith('heading') ? `H${namedStyle.slice(-1)}` : namedStyle)
      else if (block.type === 'heading' && block.level >= 1 && block.level <= 5) styles.push(`H${block.level}`)
      else if (block.type === 'heading' && block.level === 6) styles.push('H6')
      else styles.push(block.type === 'paragraph' ? 'normal' : block.type)
    }
  }
  if (collapsed || !marks.length) {
    const mark = { ...marksAtPoint(model, selection.focus), ...pending }
    if (mark.eastAsiaFontFamily && caretUsesEastAsianFont(model, selection)) mark.fontFamily = mark.eastAsiaFontFamily
    marks.push(mark)
  }
  const uniform: RunPatch = {}
  const mixedMarks: (keyof RunPatch)[] = []
  for (const key of markKeys) {
    const firstValue = marks[0]?.[key]
    if (marks.every((mark) => mark[key] === firstValue)) Object.assign(uniform, { [key]: firstValue })
    else mixedMarks.push(key)
  }
  const paragraph: Partial<DocumentParagraphStyle> = {}
  const mixedParagraph: DocumentParagraphSettingField[] = []
  for (const key of DOCUMENT_PARAGRAPH_SETTING_FIELDS) {
    const firstValue = paragraphs[0]?.[key]
    if (paragraphs.every((style) => style[key] === firstValue)) Object.assign(paragraph, { [key]: firstValue })
    else mixedParagraph.push(key)
  }
  if (mixedParagraph.includes('spacingRule') && !mixedParagraph.includes('lineHeight')) {
    mixedParagraph.push('lineHeight')
    delete paragraph.lineHeight
  }
  return { marks: uniform, mixedMarks, paragraph, mixedParagraph, blockStyle: styles.every((style) => style === styles[0]) ? styles[0] : undefined, listPreset: presets.every(p => p === presets[0]) ? presets[0] : undefined, hasSelection: !collapsed, table: documentTableSelection(model, selection) }
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
