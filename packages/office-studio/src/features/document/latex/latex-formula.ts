import type { DocumentBlock, DocumentModel } from '../../../domain/model.ts'
import type { DocumentPoint, DocumentRun, DocumentSelection, RunPatch, TextTarget } from '../model/document-model.ts'
import { blockAtPoint, createParagraph, createRun, findDocumentBlockContainer, orderedSelection, runText, runsForTarget, sameTarget, setTargetRuns, sliceRuns, textTargets } from '../model/document-model.ts'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFieldRangeIntersects, documentTextFieldRanges, transformDocumentTextFieldsForBlockEdit, transformDocumentTextFieldsForBlockRangeEdit } from '../model/document-text-fields'
import { documentRunStyleAt, validateDocumentRunPatch } from '../model/document-run-range'
import { assertDocumentTextSelectionEditable } from '../services/document-text-formatting'
import { getDocumentContentControls, transformContentControlsForBlockEdit } from '../services/content-controls'
import { transformDocumentCommentAnchorsForBodyRangeEdit } from '../comments/document-comment-anchors'

import type { LatexFormula, LatexFormulaMode } from './latex-types'
import { parseLatexFormula } from './latex-parser'
export type { LatexFormula, LatexFormulaMode, LatexNode, LatexTextStyle, LatexAccent, LatexMetrics, LatexPrimitive, LatexLayout } from './latex-types'
export { assertLatexFormulaSource, LatexFormulaError, parseLatexFormula } from './latex-parser'
export { layoutLatexFormula, latexFormulaDrawCommands, type LatexLayoutOptions } from './latex-layout'

export interface DocumentLatexFormulaTarget {
  blockId: string
  runId: string
  rangeId: string
  formula: LatexFormula
  target: TextTarget
  startOffset: number
  endOffset: number
  run: DocumentRun
}

export const DOCUMENT_LATEX_ATOM = '\uFFFC'
export const DOCUMENT_LATEX_RANGE_KIND = 'univer-docs-latex-formula'
export function createLatexFormulaRun(latex: string, mode: LatexFormulaMode = 'inline', style: Partial<DocumentRun> = {}): DocumentRun {
  validateLatexFormulaMode(mode)
  parseLatexFormula(latex)
  const run = createRun(DOCUMENT_LATEX_ATOM, style)
  run.metadata = { ...run.metadata, latexFormula: { latex, mode, version: 1 } satisfies LatexFormula }
  return run
}

export function getLatexFormula(run: DocumentRun): LatexFormula | undefined {
  const value = run.metadata?.latexFormula
  if (run.text !== DOCUMENT_LATEX_ATOM || !value || typeof value !== 'object') return undefined
  const formula = value as Partial<LatexFormula>
  return typeof formula.latex === 'string' && (formula.mode === 'inline' || formula.mode === 'block') && formula.version === 1
    ? { latex: formula.latex, mode: formula.mode, version: 1 }
    : undefined
}

/** Imported IDs survive edits/reopen. A copied run receives its own range ID. */
export function latexFormulaRangeId(run: DocumentRun): string {
  const source = run.metadata?.univerBodyLatexRangeSource
  if (source && typeof source === 'object' && !Array.isArray(source)
    && Reflect.get(source, 'runId') === run.id && typeof Reflect.get(source, 'rangeId') === 'string') {
    return Reflect.get(source, 'rangeId') as string
  }
  return run.id
}

export function getLatexFormulas(model: DocumentModel): DocumentLatexFormulaTarget[] {
  const result: DocumentLatexFormulaTarget[] = []
  for (const target of textTargets(model)) {
    let offset = 0
    for (const run of runsForTarget(model, target)) {
      const formula = getLatexFormula(run)
      if (formula) result.push({ blockId: target.cellBlockId ?? target.blockId, runId: run.id,
        rangeId: latexFormulaRangeId(run), formula, target, startOffset: offset, endOffset: offset + run.text.length, run })
      offset += run.text.length
    }
  }
  return result
}

/** Only a caret before/inside the atom or its exact selection opens editing. */
export function selectedLatexFormula(model: DocumentModel, selection: DocumentSelection): DocumentLatexFormulaTarget | undefined {
  if (selection.tableSelection) return undefined
  const { start, end } = orderedSelection(model, selection)
  if (!sameTarget(start, end)) return undefined
  return getLatexFormulas(model).find(item => sameTarget(item.target, start)
    && (start.offset === end.offset ? start.offset >= item.startOffset && start.offset < item.endOffset
      : start.offset === item.startOffset && end.offset === item.endOffset))
}

/** Reading an editor seed expands formula atoms to their original source. */
export function documentLatexSourceForSelection(model: DocumentModel, selection: DocumentSelection): string {
  if (selection.tableSelection) throw new RangeError('公式编辑器需要文字选区。')
  const targets = textTargets(model)
  for (const point of [selection.anchor, selection.focus]) {
    if (!Number.isSafeInteger(point.offset) || point.offset < 0 || !targets.some(target => sameTarget(target, point))
      || point.offset > runText(runsForTarget(model, point)).length) throw new RangeError('公式文字范围已越界或目标已不存在。')
  }
  const { start, end } = orderedSelection(model, selection)
  if (!sameTarget(start, end)) throw new RangeError('公式编辑器需要同一段落内的文字范围。')
  const runs = runsForTarget(model, start), text = runText(runs)
  for (const offset of [start.offset, end.offset]) {
    if (offset > 0 && offset < text.length && text.charCodeAt(offset - 1) >= 0xd800 && text.charCodeAt(offset - 1) <= 0xdbff
      && text.charCodeAt(offset) >= 0xdc00 && text.charCodeAt(offset) <= 0xdfff) throw new RangeError('公式文字范围不能切开 Unicode 字符。')
  }
  const source = sliceRuns(runs, start.offset, end.offset).map(run => getLatexFormula(run)?.latex ?? run.text).join('')
  if (source.length > 65_536) throw new RangeError('公式源文本超过编辑器预算。')
  return source
}

export function validateLatexFormulaMode(mode: unknown): asserts mode is LatexFormulaMode {
  if (mode !== 'inline' && mode !== 'block') throw new TypeError('公式显示形式无效。')
}

export function assertLatexFormulaSelectionEditable(model: DocumentModel, selection: DocumentSelection): void {
  assertDocumentTextSelectionEditable(model, selection)
  const { start, end } = orderedSelection(model, selection), paragraphs = documentTextFieldParagraphs(model)
  const from = paragraphs.find(item => item.block.id === (start.cellBlockId ?? start.blockId))
  const to = paragraphs.find(item => item.block.id === (end.cellBlockId ?? end.blockId))
  if (!from || !to) throw new RangeError('公式编辑目标已不存在。')
  const locked = documentTextFieldRanges(model).find(range => documentFieldLocked(range.field)
    && documentTextFieldRangeIntersects(range, from.absoluteStart + start.offset, to.absoluteStart + end.offset))
  if (locked) throw new RangeError(`公式位于锁定字段中：${locked.field.id}`)
}

export function updateLatexFormula(model: DocumentModel, runId: string, latex: string): boolean {
  parseLatexFormula(latex)
  const item = getLatexFormulas(model).find(item => item.runId === runId)
  if (!item) return false
  assertLatexFormulaSelectionEditable(model, { anchor: { ...item.target, offset: item.startOffset }, focus: { ...item.target, offset: item.endOffset } })
  item.run.metadata = { ...item.run.metadata, latexFormula: { ...item.formula, latex } }
  return true
}

export function createLatexFormulaBlock(latex: string, style: Partial<DocumentRun> = {}): DocumentBlock {
  const block = createParagraph()
  block.style = { ...block.style, align: 'center', spaceBefore: 8, spaceAfter: 8 }
  block.runs = [createLatexFormulaRun(latex, 'block', style)]
  block.metadata = { ...block.metadata, latexFormulaBlock: true }
  return block
}

export function insertLatexFormula(model: DocumentModel, selection: DocumentSelection, latex: string, mode: LatexFormulaMode = 'inline', marks: RunPatch = {}): DocumentPoint {
  validateLatexFormulaMode(mode)
  validateDocumentRunPatch(marks)
  parseLatexFormula(latex)
  if (selection.tableSelection) throw new RangeError('公式需要单元格内的文字选区或插入位置')
  const { start, end } = orderedSelection(model, selection)
  if (!sameTarget(start, end)) throw new RangeError('公式插入暂时要求在同一段落内选择内容。')
  const block = blockAtPoint(model, start), container = block && findDocumentBlockContainer(model, block.id)
  if (!block || !container) throw new RangeError('公式插入目标已不存在')
  assertLatexFormulaSelectionEditable(model, { anchor: start, focus: end })
  const sourceParagraphs = documentTextFieldParagraphs(model), paragraph = sourceParagraphs.find(paragraph => paragraph.block.id === block.id)
  if (!paragraph) throw new RangeError('公式插入目标缺少正文映射')
  const before = sliceRuns(runsForTarget(model, start), 0, start.offset)
  const after = sliceRuns(runsForTarget(model, start), end.offset)
  const style = { ...documentRunStyleAt(runsForTarget(model, start), start.offset, end.offset), ...marks }
  const controls: DocumentModel = { ...model, contentControls: model.contentControls && getDocumentContentControls(model) }
  if (mode === 'inline') {
    const formula = createLatexFormulaRun(latex, mode, style)
    transformContentControlsForBlockEdit(controls, block.id, start.offset, end.offset, formula.text.length)
    transformDocumentTextFieldsForBlockEdit(block, start.offset, end.offset, formula.text, undefined, model)
    model.contentControls = controls.contentControls
    setTargetRuns(model, start, [...before, formula, ...after])
    transformDocumentCommentAnchorsForBodyRangeEdit(model, sourceParagraphs, documentTextFieldParagraphs(model),
      paragraph.absoluteStart + start.offset, paragraph.absoluteStart + end.offset, formula.text)
    return { ...start, offset: start.offset + formula.text.length }
  }
  const index = container.indexOf(block)
  const equation = createLatexFormulaBlock(latex, style)
  const trailing = createParagraph()
  trailing.style = { ...block.style }
  trailing.runs = after.length ? after : [createRun()]
  const replacement = `\n${DOCUMENT_LATEX_ATOM}\n`
  transformContentControlsForBlockEdit(controls, block.id, start.offset, end.offset, replacement.length, {
    blockId: trailing.id, firstBreakOffset: start.offset, lastParagraphOffset: start.offset + replacement.length,
  })
  transformDocumentTextFieldsForBlockRangeEdit([block], start.offset, end.offset, replacement, [block, equation, trailing], undefined, model)
  model.contentControls = controls.contentControls
  setTargetRuns(model, start, before.length ? before : [createRun()])
  container.splice(index + 1, 0, equation, trailing)
  transformDocumentCommentAnchorsForBodyRangeEdit(model, sourceParagraphs, documentTextFieldParagraphs(model),
    paragraph.absoluteStart + start.offset, paragraph.absoluteStart + end.offset, replacement)
  return start.row === undefined ? { blockId: equation.id, offset: DOCUMENT_LATEX_ATOM.length }
    : { ...start, cellBlockId: equation.id, offset: DOCUMENT_LATEX_ATOM.length }
}
