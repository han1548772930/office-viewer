import type { DocumentCodeConfig, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { clampPoint, findDocumentBlock, sameTarget, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentBlockHostForId, documentEditingBlockHost } from '../model/document-block-host'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { documentCodeSourceEdit, normalizeDocumentCodeSource } from '../model/document-code-source'
import { isDocumentTextBlock, validateDocumentCodeConfig } from '../model/document-structural-blocks'
import { reconcileDocumentContentControlBlockEdits } from '../model/content-control-block-edits'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentPointAtClampedBodyOffset } from './document-block-insertion'
import { applyDocumentBlockHostModel, documentModelForBlockHost } from './document-block-host'
import { replaceDocumentRange } from './document-facade'
import { assertDocumentTextSelectionEditable } from './document-text-formatting'
import { resolveDocumentContentControlTextInputId, synchronizeDocumentContentControlsAfterEdit } from './content-control-editing'
import { synchronizeDocumentContentControlBindings } from './content-control-bindings'
import { validateDocumentContentControlRanges } from './content-controls'
import { removeDeletedDocumentRepeatingNoteScopes } from './repeating-content-control-copy'
import { documentCodeUsesFontFamily, setDocumentCodeFontFamily, validateDocumentCodeFontFamily } from '../code-blocks/document-code-fonts'

export interface DocumentStructuredCodeUpdate {
  blockId: string
  source?: string
  patch?: Partial<DocumentCodeConfig>
  fontFamily?: string
  host?: DocumentAuxiliaryBlockHost
}

/** Source, linked controls and auxiliary stories are committed as one full-document candidate. */
export function updateDocumentStructuredCodeBlock(model: DocumentModel, selection: DocumentSelection,
  update: DocumentStructuredCodeUpdate, context: DocumentPermissionContext): { changed: boolean; selection: DocumentSelection } {
  const host = update.host ?? documentBlockHostForId(model, update.blockId)
  const source = host ? documentModelForBlockHost(model, host, true) : model
  const original = findDocumentBlock(source, update.blockId)
  if (original?.type !== 'code' || !original.code || !original.children) throw new RangeError('代码块已不存在。')
  const config = { ...original.code, ...update.patch }
  validateDocumentCodeConfig(config)
  if (update.fontFamily !== undefined) validateDocumentCodeFontFamily(update.fontFamily)
  const fontChanged = update.fontFamily !== undefined && !documentCodeUsesFontFamily(original, update.fontFamily)
  const paragraphs = documentTextFieldParagraphs(source)
  const children = documentTextFieldParagraphs({ blocks: original.children })
  const currentText = children.map(paragraph => paragraph.text).join('\n')
  const nextText = update.source === undefined ? currentText : normalizeDocumentCodeSource(update.source)
  const textChanged = nextText !== currentText
  if (!textChanged && !fontChanged && JSON.stringify(config) === JSON.stringify(original.code)) return { changed: false, selection: cloneModel(selection) }
  if (textChanged && (!original.children.length || original.children.some(block => !isDocumentTextBlock(block)))) {
    throw new RangeError('代码源文本编辑不能移除嵌套的结构内容。')
  }
  let candidate = cloneModel(model)
  let scoped = host ? documentModelForBlockHost(candidate, host, true) : candidate
  let nextSelection = cloneModel(selection)
  let editSelection: DocumentSelection | undefined
  let activeControlId: string | undefined
  let originTextEdits: Array<{ startOffset: number; endOffset: number; text: string }> | undefined
  if (textChanged) {
    const first = paragraphs.find(paragraph => paragraph.block.id === children[0]!.block.id)
    if (!first) throw new RangeError('代码块缺少真实输入段落。')
    const change = documentCodeSourceEdit(currentText, nextText)
    const startOffset = first.absoluteStart + change.startOffset, endOffset = first.absoluteStart + change.endOffset
    editSelection = { anchor: documentPointAtClampedBodyOffset(source, startOffset), focus: documentPointAtClampedBodyOffset(source, endOffset) }
    assertDocumentTextSelectionEditable(source, editSelection)
    activeControlId = resolveDocumentContentControlTextInputId(source, editSelection)
    scoped = replaceDocumentRange(scoped, startOffset, endOffset, change.text)
    const mapPoint = (point: DocumentPoint): DocumentPoint => {
      const paragraph = paragraphs.find(paragraph => paragraph.block.id === (point.cellBlockId ?? point.blockId))
      if (!paragraph) return clampPoint(scoped, point)
      const offset = paragraph.absoluteStart + point.offset
      const mapped = offset <= startOffset ? offset : offset >= endOffset ? offset + change.text.length - (endOffset - startOffset) : startOffset + change.text.length
      const target = documentPointAtClampedBodyOffset(scoped, mapped)
      return sameTarget(target, point) ? { ...target, pageAffinity: point.pageAffinity } : target
    }
    nextSelection = { ...nextSelection, anchor: mapPoint(selection.anchor), focus: mapPoint(selection.focus) }
    originTextEdits = [{ startOffset, endOffset, text: change.text }]
  }
  const code = findDocumentBlock(scoped, update.blockId)
  if (code?.type !== 'code') throw new RangeError('代码源文本修改不能删除其容器。')
  code.code = config
  if (update.fontFamily !== undefined) setDocumentCodeFontFamily(scoped, code, update.fontFamily)
  if (host) applyDocumentBlockHostModel(candidate, host, scoped)
  else candidate = scoped
  removeDeletedDocumentRepeatingNoteScopes(model, candidate)
  reconcileDocumentContentControlBlockEdits(model, candidate, { selection, absorbAdjacentInsertions: false })
  if (editSelection) {
    const editResult = synchronizeDocumentContentControlsAfterEdit(model, candidate, { origin: 'text-input', selection: editSelection, activeControlId })
    const bindings = synchronizeDocumentContentControlBindings(model, candidate, { sourceSelection: editSelection, selection: nextSelection,
      host: host ?? documentEditingBlockHost(model), originTextEdits, activeControlId, editResult, permissionContext: context })
    nextSelection = bindings.selection ?? nextSelection
  }
  validateDocumentContentControlRanges(candidate)
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, candidate))
  Object.assign(model, candidate)
  return { changed: true, selection: nextSelection }
}
