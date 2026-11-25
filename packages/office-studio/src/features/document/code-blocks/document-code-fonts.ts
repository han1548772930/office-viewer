import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { applyRunPatch, createRun, runText, textTargets } from '../model/document-model'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { assertDocumentTextSelectionEditable } from '../services/document-text-formatting'
import { validateDocumentRunPatch } from '../model/document-run-range'

export interface DocumentCodeFontFamilyOption { value: string; label?: string; category: 'monospace' }

// Fixed Univer FontService monospace entries, plus the workbench's existing Windows defaults.
const defaults = ['Courier New', 'Courier', 'BIZ UDGothic', 'BIZ UDMincho', 'Andale Mono', 'Menlo', 'Monaco', 'PCMyungjo', 'PT Mono',
  'Consolas', 'Cascadia Code', 'monospace']
const registrations: Array<{ id: symbol; fonts: DocumentCodeFontFamilyOption[] }> = []

/** Host resources explicitly identify custom monospaced fonts; registration does not change document content. */
export function registerDocumentCodeFontFamilies(fonts: readonly DocumentCodeFontFamilyOption[]): () => void {
  const checked = fonts.map(font => {
    if (font.category !== 'monospace' || !font.value.trim() || font.value.length > 256) throw new RangeError('代码字体必须声明为有效的等宽字体。')
    validateDocumentRunPatch({ fontFamily: font.value })
    return { ...font }
  })
  const id = Symbol('document-code-fonts'); registrations.push({ id, fonts: checked })
  return () => { const at = registrations.findIndex(value => value.id === id); if (at >= 0) registrations.splice(at, 1) }
}

export function documentCodeFontFamilyOptions(current?: string): Array<{ value: string; label: string; disabled?: boolean }> {
  const fonts = new Map(defaults.map(value => [value, { value, label: value }]))
  for (const registration of registrations) for (const font of registration.fonts) fonts.set(font.value, { value: font.value, label: font.label ?? font.value })
  const result: Array<{ value: string; label: string; disabled?: boolean }> = [...fonts.values()]
  if (current && !fonts.has(current)) result.push({ value: current, label: current, disabled: true })
  return result
}

export function validateDocumentCodeFontFamily(fontFamily: string): void {
  validateDocumentRunPatch({ fontFamily })
  if (!documentCodeFontFamilyOptions().some(font => font.value.toLowerCase() === fontFamily.toLowerCase())) {
    throw new RangeError('代码块字体需要是已登记的等宽字体。')
  }
}

function paragraphs(block: DocumentBlock): DocumentBlock[] {
  return documentTextFieldParagraphs({ blocks: block.type === 'code' ? block.children ?? [] : [block] }).map(value => value.block)
}

export function documentCodeFontFamily(block: DocumentBlock): string {
  return paragraphs(block)[0]?.runs[0]?.fontFamily ?? 'monospace'
}

export function documentCodeUsesFontFamily(block: DocumentBlock, fontFamily: string): boolean {
  return paragraphs(block).every(paragraph => paragraph.runs.length > 0 && paragraph.runs.every(run => run.fontFamily === fontFamily
    && run.eastAsiaFontFamily === fontFamily) && paragraph.metadata.documentCodeFontFamily === fontFamily)
}

/** Whole-block direct formatting and paragraph defaults share content locks and the outer document candidate. */
export function setDocumentCodeFontFamily(model: DocumentModel, block: DocumentBlock, fontFamily: string): string[] {
  validateDocumentCodeFontFamily(fontFamily)
  const content = paragraphs(block), targets = textTargets(model), patch = { fontFamily, eastAsiaFontFamily: fontFamily }
  for (const paragraph of content) {
    const target = targets.find(value => (value.cellBlockId ?? value.blockId) === paragraph.id)
    if (!target) throw new RangeError('代码字体目标段落已不存在。')
    assertDocumentTextSelectionEditable(model, { anchor: { ...target, offset: 0 }, focus: { ...target, offset: runText(paragraph.runs).length } })
    paragraph.runs = paragraph.runs.length ? paragraph.runs.map(run => applyRunPatch(run, patch)) : [createRun('', patch)]
    paragraph.metadata = { ...paragraph.metadata, documentCodeFontFamily: fontFamily }
  }
  return content.map(paragraph => paragraph.id)
}
