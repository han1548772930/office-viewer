import type { DocumentModel } from '../../../domain/model'
import { assertLatexFormulaSelectionEditable, documentLatexSourceForSelection, getLatexFormulas, insertLatexFormula, updateLatexFormula,
  type LatexFormulaMode } from '../latex/latex-formula'
import { blockAtPoint, collapsedSelection, type DocumentSelection, type RunPatch } from '../model/document-model'
import { assertDocumentPermission, DocumentPermissionError, permissionTargetsForSelection, resolveDocumentPermission, type DocumentPermissionContext } from '../permissions/document-permissions'

export type DocumentLatexCommand =
  | { id: 'doc.command.insert-latex-formula'; latex: string; mode?: LatexFormulaMode; marks?: RunPatch }
  | { id: 'doc.command.update-latex-formula'; rangeId: string; latex: string }
  | { id: 'doc.command.remove-latex-formula'; rangeId: string }
  | { id: 'doc.command.replace-latex-formula-with-text'; rangeId: string; text: string }

export function isDocumentLatexCommand(command: { id: string }): command is DocumentLatexCommand {
  return ['doc.command.insert-latex-formula', 'doc.command.update-latex-formula',
    'doc.command.remove-latex-formula', 'doc.command.replace-latex-formula-with-text'].includes(command.id)
}

export function documentLatexSelectionState(model: DocumentModel, selection: DocumentSelection, context: DocumentPermissionContext,
  disabled = false): { available: boolean; reason?: string } {
  if (disabled) return { available: false }
  const permission = resolveDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, selection))
  if (!permission.allowed) return { available: false, reason: new DocumentPermissionError(permission).message }
  try { assertLatexFormulaSelectionEditable(model, selection); documentLatexSourceForSelection(model, selection); return { available: true } }
  catch (error) {
    if (error instanceof Error) return { available: false, reason: error.message }
    throw error
  }
}

/** The ordinary text mutation owns replacement/anchor transforms and history. */
export function executeDocumentLatexCommand(model: DocumentModel, selection: DocumentSelection, command: DocumentLatexCommand,
  context: DocumentPermissionContext, replaceText: (model: DocumentModel, selection: DocumentSelection, text: string) => DocumentSelection): DocumentSelection {
  if (command.id === 'doc.command.insert-latex-formula') {
    assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, selection))
    return collapsedSelection(insertLatexFormula(model, selection, command.latex, command.mode, command.marks))
  }
  if (typeof command.rangeId !== 'string' || !command.rangeId.trim()) throw new TypeError('公式范围标识无效。')
  const item = getLatexFormulas(model).find(item => item.rangeId === command.rangeId)
  if (!item) throw new RangeError('公式已被删除或替换。')
  const target: DocumentSelection = { anchor: { ...item.target, offset: item.startOffset }, focus: { ...item.target, offset: item.endOffset } }
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, target))
  assertLatexFormulaSelectionEditable(model, target)
  if (command.id === 'doc.command.update-latex-formula') {
    updateLatexFormula(model, item.runId, command.latex)
    return target
  }
  if (command.id === 'doc.command.replace-latex-formula-with-text' && typeof command.text !== 'string') throw new TypeError('公式替换文本必须为字符串。')
  const result = replaceText(model, target, command.id === 'doc.command.remove-latex-formula' ? '' : command.text)
  const block = blockAtPoint(model, item.target)
  if (block) delete block.metadata.latexFormulaBlock
  return result
}
