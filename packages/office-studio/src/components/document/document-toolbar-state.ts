import type { DocumentModel, DocumentParagraphStyle } from '../../domain/model'
import { caretUsesEastAsianFont, hasEastAsianLetters } from '../../features/document/model/document-font'
import { blockListPreset, type ListPreset } from '../../features/document/model/document-lists'
import { blockAtPoint } from '../../features/document/commands/document-commands'
import { documentStorySelectionFormatting } from '../../features/document/services/document-story-formatting'
import type { DocumentStoryRange } from '../../features/document/model/document-story'
import { documentTableSelection, type DocumentTableSelection } from '../../features/document/commands/table-commands'
import { documentTableSelectedTextTargets } from '../../features/document/model/document-table-selection'
import { documentParagraphPagination } from '../../features/document/model/document-paragraph-pagination'
import { DOCUMENT_PARAGRAPH_SETTING_FIELDS, documentParagraphSettingsStyle, type DocumentParagraphSettingField } from './document-paragraph-settings-state'
import {
  DOCUMENT_NAMED_PARAGRAPH_STYLES, applyRunPatch, createRun, marksAtPoint, orderedSelection, runsForTarget, sameTarget, selectedPlainText,
  selectionCollapsed, textTargets, type DocumentNamedParagraphStyle, type DocumentRun, type DocumentSelection, type RunPatch,
} from '../../features/document/model/document-model'

const markKeys = ['bold', 'italic', 'underline', 'strike', 'verticalAlign', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'color', 'highlight', 'href'] as const

export interface DocumentToolbarState {
  marks: RunPatch
  mixedMarks: readonly (keyof RunPatch)[]
  paragraph: Partial<DocumentParagraphStyle>
  mixedParagraph?: readonly DocumentParagraphSettingField[]
  blockStyle?: string
  namedStyle?: DocumentNamedParagraphStyle
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
  const selected = selection.tableSelection ? documentTableSelectedTextTargets(model, selection.tableSelection) : targets.slice(Math.max(0, first), last + 1)
  const marks: RunPatch[] = []
  const paragraphs: DocumentParagraphStyle[] = []
  const styles: string[] = []
  const namedStyles: DocumentNamedParagraphStyle[] = []
  const presets: (ListPreset | undefined)[] = []
  for (const [index, target] of selected.entries()) {
    const runs = runsForTarget(model, target)
    const from = !selection.tableSelection && index === 0 ? start.offset : 0
    const to = !selection.tableSelection && index === selected.length - 1 ? end.offset : Number.POSITIVE_INFINITY
    const included: DocumentRun[] = []
    let offset = 0
    for (const run of runs) {
      if (!collapsed && (offset < to && offset + run.text.length > from || selection.tableSelection && !run.text.length)) {
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
      paragraphs.push(documentParagraphSettingsStyle({ ...block.style, ...documentParagraphPagination(block.style, model) }))
      const namedStyle = block.metadata.namedStyleType
      if (typeof namedStyle === 'string' && Object.hasOwn(DOCUMENT_NAMED_PARAGRAPH_STYLES, namedStyle)) {
        styles.push(block.type === 'checklist' ? 'checklist' : namedStyle.startsWith('heading') ? `H${namedStyle.slice(-1)}` : namedStyle)
        namedStyles.push(namedStyle as DocumentNamedParagraphStyle)
        continue
      }
      if (block.type === 'heading' && block.level >= 1 && block.level <= 5) {
        styles.push(`H${block.level}`)
        namedStyles.push(`heading${block.level}` as DocumentNamedParagraphStyle)
        continue
      }
      styles.push(block.type === 'heading' && block.level === 6 ? 'H6' : block.type === 'paragraph' ? 'normal' : block.type)
      namedStyles.push('normal')
    }
  }
  if (collapsed || !marks.length && (!selection.tableSelection || selected.length)) {
    const point = selection.tableSelection ? { ...selected[0]!, offset: 0 } : selection.focus
    const mark = applyRunPatch(createRun('', marksAtPoint(model, point)), collapsed ? pending : {})
    if (mark.eastAsiaFontFamily && caretUsesEastAsianFont(model, selection)) mark.fontFamily = mark.eastAsiaFontFamily
    marks.push(mark)
  }
  const { marks: uniform, mixedMarks } = documentToolbarMarks(marks)
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
  return { marks: uniform, mixedMarks, paragraph, mixedParagraph, blockStyle: styles.every((style) => style === styles[0]) ? styles[0] : undefined, namedStyle: namedStyles[0], listPreset: presets.every(p => p === presets[0]) ? presets[0] : undefined, hasSelection: !collapsed, table: documentTableSelection(model, selection) }
}

function documentToolbarMarks(styles: readonly RunPatch[]): Pick<DocumentToolbarState, 'marks' | 'mixedMarks'> {
  const marks: RunPatch = {}, mixedMarks: (keyof RunPatch)[] = []
  for (const key of markKeys) {
    const value = styles[0]?.[key]
    if (styles.every(style => style[key] === value)) Object.assign(marks, { [key]: value })
    else mixedMarks.push(key)
  }
  return { marks, mixedMarks }
}

/** Shared character controls use the active header/footer input range and its field caches. */
export function documentStoryToolbarState(model: DocumentModel, range: DocumentStoryRange, pending: RunPatch = {}): DocumentToolbarState {
  const selected = documentStorySelectionFormatting(model, range, pending)
  const paragraph: Partial<DocumentParagraphStyle> = {}, mixedParagraph: DocumentParagraphSettingField[] = []
  const paragraphs = selected.paragraphs.map(style => documentParagraphSettingsStyle({ ...style, ...documentParagraphPagination(style, model) }))
  for (const key of DOCUMENT_PARAGRAPH_SETTING_FIELDS) {
    const value = paragraphs[0]?.[key]
    if (paragraphs.every(style => style[key] === value)) Object.assign(paragraph, { [key]: value })
    else mixedParagraph.push(key)
  }
  if (mixedParagraph.includes('spacingRule') && !mixedParagraph.includes('lineHeight')) {
    mixedParagraph.push('lineHeight'); delete paragraph.lineHeight
  }
  return { ...documentToolbarMarks(selected.styles), paragraph, mixedParagraph, namedStyle: selected.namedStyle, hasSelection: selected.hasSelection }
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
