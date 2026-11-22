import type { DocumentModel } from '../../../domain/model'
import { runText, runsForTarget, selectedPlainText, selectionCollapsed, type DocumentSelection, type RunPatch } from './document-model'

// Only letters select the East Asian font slot; shared punctuation alone does not.
export const hasEastAsianLetters = (text: string) => /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}]/u.test(text)

export function caretUsesEastAsianFont(model: DocumentModel, selection: DocumentSelection) {
  const point = selection.focus
  const text = runText(runsForTarget(model, point))
  const lowSurrogate = point.offset > 0 && /[\uDC00-\uDFFF]/.test(text.charAt(point.offset - 1))
  return hasEastAsianLetters(text.slice(Math.max(0, point.offset - (lowSurrogate ? 2 : 1)), point.offset))
}

export function selectionFontPatch(model: DocumentModel, selection: DocumentSelection, patch: RunPatch): RunPatch {
  if (!patch.fontFamily || patch.eastAsiaFontFamily !== undefined) return patch
  const cjk = selectionCollapsed(selection) ? caretUsesEastAsianFont(model, selection) : hasEastAsianLetters(selectedPlainText(model, selection))
  return cjk ? { ...patch, eastAsiaFontFamily: patch.fontFamily } : patch
}
