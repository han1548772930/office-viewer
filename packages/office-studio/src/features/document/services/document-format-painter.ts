import type { DocumentBlock, DocumentModel, DocumentParagraphStyle } from '../../../domain/model'
import { blockAtPoint, defaultParagraphStyle, marksAtPoint, orderedSelection, runText, runsForTarget, sameTarget, selectionCollapsed, type DocumentSelection, type RunPatch } from '../model/document-model'

type CharacterFormat = Required<Pick<RunPatch, 'fontFamily' | 'eastAsiaFontFamily' | 'fontSize' | 'color' | 'bold' | 'italic' | 'underline' | 'strike' | 'verticalAlign' | 'highlight'>>
interface ParagraphFormat {
  type: DocumentBlock['type']; level: number; style: DocumentParagraphStyle; namedStyle?: string; checklistId?: string
}
export interface DocumentPainterFormat { text: CharacterFormat; paragraph?: ParagraphFormat }

function characterFormat(marks: RunPatch): CharacterFormat {
  // Copy only appearance. Hyperlinks, field codes and annotation metadata belong to the target.
  return { fontFamily: marks.fontFamily ?? 'Arial', eastAsiaFontFamily: marks.eastAsiaFontFamily ?? marks.fontFamily ?? 'Arial', fontSize: marks.fontSize ?? 11, color: marks.color ?? '#1f2937',
    bold: marks.bold ?? false, italic: marks.italic ?? false, underline: marks.underline ?? false, strike: marks.strike ?? false,
    highlight: marks.highlight ?? '', verticalAlign: marks.verticalAlign ?? 'baseline' }
}

export function clearDocumentFormat(): DocumentPainterFormat {
  return { text: characterFormat({}), paragraph: { type: 'paragraph', level: 0, style: defaultParagraphStyle() } }
}

export function captureDocumentFormat(model: DocumentModel, selection: DocumentSelection): DocumentPainterFormat | undefined {
  const { start, end } = orderedSelection(model, selection)
  const block = blockAtPoint(model, start)
  if (!block || !['paragraph', 'heading', 'quote', 'numbered', 'bullet', 'checklist'].includes(block.type)) return
  // marksAtPoint reads the preceding character; a non-empty selection samples the first selected one.
  const text = characterFormat(marksAtPoint(model, { ...start, offset: start.offset + (selectionCollapsed(selection) ? 0 : 1) }))
  if (sameTarget(start, end)) return { text }
  return { text, paragraph: { type: block.type, level: block.level, style: { ...block.style },
    namedStyle: typeof block.metadata.namedStyleType === 'string' ? block.metadata.namedStyleType : undefined, checklistId: block.checklist?.listId } }
}

/** Like the upstream word-boundary lookup, try the caret first and the preceding character second. */
export function documentPainterTarget(model: DocumentModel, selection: DocumentSelection): DocumentSelection | undefined {
  if (!selectionCollapsed(selection)) return selection
  const point = selection.focus, text = runText(runsForTarget(model, point))
  const segments = typeof Intl.Segmenter === 'function' ? new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text) : undefined
  for (const at of [point.offset, point.offset - 1]) {
    if (at < 0 || at >= text.length) continue
    const segment = segments?.containing(at)
    const word = segments ? (segment?.isWordLike ? { start: segment.index, length: segment.segment.length } : undefined)
      : Array.from(text.matchAll(/[\p{L}\p{N}\p{M}_]+/gu), match => ({ start: match.index!, length: match[0].length })).find(word => at >= word.start && at < word.start + word.length)
    if (word) return { anchor: { ...point, offset: word.start }, focus: { ...point, offset: word.start + word.length } }
  }
}

export function applyPainterParagraph(block: DocumentBlock, format: ParagraphFormat) {
  block.style = { ...format.style }
  block.type = format.type
  block.level = format.level
  if (format.namedStyle === undefined) delete block.metadata.namedStyleType
  else block.metadata.namedStyleType = format.namedStyle
  if (format.type === 'checklist' && format.checklistId) block.checklist = { listId: format.checklistId, checked: block.checklist?.checked ?? false }
  else delete block.checklist
}
