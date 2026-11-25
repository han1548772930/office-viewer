import type {
  DocumentBlock,
  DocumentContentControl,
  DocumentModel,
} from '../../../domain/model'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFieldRangeIntersects, documentTextFieldRanges, documentTextFields, insertDocumentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit, transformDocumentTextFieldsForBlockRangeEdit, type DocumentTextFieldParagraph } from '../model/document-text-fields'
import {
  compactRuns,
  createParagraph,
  createRun,
  documentBlockAncestry,
  documentId,
  runText,
  runsForTarget,
  setTargetRuns,
  sliceRuns,
  tableForBlock,
  textTargets,
  type DocumentRun,
  type DocumentSelection,
  type RunPatch,
  type TextTarget,
} from '../model/document-model'
import type { DocumentImage } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentRunHasNoteReference } from '../model/document-note-reference'
import { assertDocumentNoteReferenceBoundaryRank, documentNoteReferenceOffsets, type DocumentNoteReferenceOffset } from '../model/document-note-boundary'
import { assertDocumentNoteReferenceOffsets, compareDocumentNoteBoundaries, documentNoteReferenceOffsetsAfterTextEdit, remapDocumentNoteBoundaryParagraph, type DocumentNoteBoundaryParagraph, type DocumentNoteEditRanks } from '../model/document-note-boundary-edit'
import { assertDocumentNoteReferenceParagraphIds, documentNoteEmptyReplacementBoundary, documentNoteInsertionBoundary,
  documentNoteReferenceEmptyRemovalRange, documentNoteReferenceEmptyReplacementAffinity, documentNoteReferenceInsertionAffinity } from '../model/document-note-insertion-affinity'
import { assignDocumentContentControlNoteBoundaries, documentContentControlNoteBoundaryRanks, documentContentControlNoteRangeContains,
  documentContentControlNoteRangesOverlap, documentContentControlPointBeforeInsertion, partitionDocumentContentControlNoteRuns } from '../model/content-control-note-boundary'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { documentBlockHostForId, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { applyDocumentBlockHostModel, documentModelForBlockHost } from './document-block-host'
import { cloneModel } from '../../../domain/serialization'
import { assertDocumentStoryLocksPreserved } from './document-story-locks'
import { resolveSectionHeaderFooter } from './section-header-footer'
import { contentControlCheckboxGlyph, formatContentControlDate, validContentControlIsoDate } from './content-control-presentation'
import { validateContentControlPropertyPatch } from './content-control-properties'
import { synchronizeDocumentContentControlsAfterEdit, unwrapDocumentContentControlShells } from './content-control-editing'
import { documentRunStyleAt, splitDocumentRunRange } from '../model/document-run-range'
import { documentContentControlSourceOrders } from '../model/content-control-source-order'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/document-comment-anchors'
import { documentContentControlContains, documentContentControlContainsBlockRange, documentContentControlContainsSelection,
  documentContentControlOwnedRowBlockIds, documentContentControlTableRowBlockIds, reanchorDocumentContentControlTableRows,
  reconcileDocumentContentControlTableRows, resolveDocumentContentControlTableRows } from '../model/content-control-table-rows'
import { assertDocumentRepeatingContentControlRowsEditable, prepareDocumentRepeatingContentControlRowsMutation } from './repeating-content-control-rows'
import { documentContentControlBlockTextTargets, documentContentControlOwnedBlockIds, reanchorDocumentBlockContentControl,
  resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
import { assertDocumentRepeatingContentControlBlocksEditable, prepareDocumentRepeatingContentControlBlocksMutation } from './repeating-content-control-blocks'
import { assertDocumentRepeatingNoteContentRemovable, copyDocumentRepeatingContentControlComments, copyDocumentRepeatingContentControlNoteControls, copyDocumentRepeatingContentControlNotes,
  createDocumentRepeatingContentControlCopyIds, documentRepeatingContentControlInlineCopySource, indexDocumentRepeatingContentControlCopies,
  remapDocumentRepeatingContentControlMetadata, removeDeletedDocumentRepeatingNoteScopes, visitDocumentRepeatingContentControlBlocks,
  type DocumentRepeatingContentControlCopyIds } from './repeating-content-control-copy'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentContentControlPermissionTargets } from '../permissions/content-control-permissions'
import { copyDocumentBlocksForPaste } from './document-block-paste'
import { reconcileDocumentContentControlBlockEdits } from '../model/content-control-block-edits'
import { reconcileDocumentTableAnchors } from './document-table-anchors'
import { replaceDocumentBlockContentControlImage, replaceDocumentBlockContentControlText } from './content-control-block-value'
import { finalizeDocumentContentControlHostResources } from './content-control-host-resources'
import { documentContentControlGroupEditAllowed } from '../model/content-control-group-lock'

export type ContentControlInput = Omit<DocumentContentControl, 'id'> & { id?: string }

export interface ContentControlRangeEditOptions {
  /** A checked input context explicitly addresses the contents at start/end boundaries. */
  activeControlId?: string
  /** Captured future geometry also carries note atom changes with no display-text delta. */
  projectedNoteReferences?: readonly DocumentNoteReferenceOffset[]
  /** Derived renumbering may only replace this surviving reference's complete label. */
  noteReferenceLabelRunId?: string
  /** Explicit selections may leave empty note atoms before/after equal display offsets. */
  sourceNoteRanks?: DocumentNoteEditRanks
  /** Body structural contexts require the actual detached model captured before editing. */
  sourceModel?: DocumentModel
}

function insideControlEditChain(model: DocumentModel, blockId: string, startOffset: number, endOffset: number,
  activeControlId: string | undefined, sourceParagraph?: Pick<DocumentTextFieldParagraph, 'text' | 'noteReferences'>,
  startNoteReferenceCount?: number, endNoteReferenceCount?: number): Set<string> {
  const result = new Set<string>()
  if (!activeControlId) return result
  const index = new Map((model.contentControls ?? []).map(control => [control.id, control]))
  const contains = (control: DocumentContentControl) => sourceParagraph && !control.blockRange && !control.tableRowRange
    ? control.blockId === blockId && isDocumentTextBoundary(sourceParagraph.text, startOffset) && isDocumentTextBoundary(sourceParagraph.text, endOffset)
      && compareDocumentNoteBoundaries(control.startOffset, documentContentControlNoteBoundaryRanks(control, sourceParagraph.noteReferences).start,
        startOffset, assertDocumentNoteReferenceBoundaryRank(sourceParagraph.noteReferences, startOffset, startNoteReferenceCount)) <= 0
      && compareDocumentNoteBoundaries(endOffset, assertDocumentNoteReferenceBoundaryRank(sourceParagraph.noteReferences, endOffset, endNoteReferenceCount),
        control.endOffset, documentContentControlNoteBoundaryRanks(control, sourceParagraph.noteReferences).end) <= 0
    : documentContentControlContainsBlockRange(model, control, blockId, startOffset, endOffset, startNoteReferenceCount, endNoteReferenceCount)
  let current = index.get(activeControlId)
  if (!current || current.metadata?.officeStoryLockSource
    || !Number.isSafeInteger(startOffset) || !Number.isSafeInteger(endOffset)
    || !contains(current)) {
    throw new RangeError('内容控件内部编辑上下文已失效：' + activeControlId)
  }
  while (current) {
    if (result.has(current.id)) throw new RangeError('内容控件内部编辑遇到循环父级：' + current.id)
    if (!contains(current)) {
      throw new RangeError('内容控件内部编辑不属于其父级范围：' + current.id)
    }
    result.add(current.id)
    const parentId: string | undefined = current.parentId
    if (parentId && !index.has(parentId)) throw new RangeError('内容控件内部编辑的父级不存在：' + parentId)
    current = parentId ? index.get(parentId) : undefined
  }
  return result
}

/** A structural body selection can span paragraphs. Its ownership must be
 * checked against the old tree, rather than applying its global end to the
 * first paragraph or resolving deleted rows in the destination tree. */
function insideBodyControlEditChain(model: DocumentModel, sourceModel: DocumentModel | undefined,
  paragraphs: readonly DocumentTextFieldParagraph[], first: DocumentTextFieldParagraph, last: DocumentTextFieldParagraph,
  start: number, startRank: number, end: number, endRank: number, activeControlId: string | undefined): Set<string> {
  const result = new Set<string>()
  if (!activeControlId) return result
  const context = sourceModel ?? model, index = new Map((context.contentControls ?? []).map(control => [control.id, control]))
  const byId = new Map(paragraphs.map(paragraph => [paragraph.block.id, paragraph]))
  const sourceTargets = sourceModel ? textTargets(sourceModel) : []
  const selection = (): DocumentSelection => {
    if (!sourceModel) throw new RangeError('内容控件结构编辑需要真实的编辑前模型。')
    const anchor = sourceTargets.find(target => (target.cellBlockId ?? target.blockId) === first.block.id)
    const focus = sourceTargets.find(target => (target.cellBlockId ?? target.blockId) === last.block.id)
    if (!anchor || !focus) throw new RangeError('内容控件结构编辑的源选择不可寻址。')
    return { anchor: { ...anchor, offset: start }, focus: { ...focus, offset: end } }
  }
  const contains = (control: DocumentContentControl): boolean => {
    if (control.blockRange || control.tableRowRange) return documentContentControlContainsSelection(context, control, selection())
    const paragraph = byId.get(control.blockId)
    if (!paragraph || paragraph !== first || paragraph !== last) return false
    const ranks = documentContentControlNoteBoundaryRanks(control, paragraph.noteReferences)
    return compareDocumentNoteBoundaries(control.startOffset, ranks.start, start, startRank) <= 0
      && compareDocumentNoteBoundaries(end, endRank, control.endOffset, ranks.end) <= 0
  }
  let current = index.get(activeControlId)
  if (!current || current.metadata?.officeStoryLockSource || !contains(current)) {
    throw new RangeError('内容控件内部编辑上下文已失效：' + activeControlId)
  }
  while (current) {
    if (result.has(current.id)) throw new RangeError('内容控件内部编辑遇到循环父级：' + current.id)
    if (!contains(current)) throw new RangeError('内容控件内部编辑不属于其父级范围：' + current.id)
    if ((current.blockRange || current.tableRowRange) && isContentLocked(current)
      && !(current.type === 'group' && (context.contentControls ?? []).some(child => child.parentId === current!.id && contains(child)))) {
      throw new Error(`Content control content is locked: ${current.id}`)
    }
    result.add(current.id)
    const parentId: string | undefined = current.parentId
    if (parentId && !index.has(parentId)) throw new RangeError('内容控件内部编辑的父级不存在：' + parentId)
    current = parentId ? index.get(parentId) : undefined
  }
  return result
}

function assertBodyEditSourceModel(model: DocumentModel, sourceModel: DocumentModel,
  paragraphs: readonly DocumentTextFieldParagraph[]): void {
  if (sourceModel.id !== model.id) throw new RangeError('内容控件正文编辑的源模型属于其他文档。')
  const source = documentTextFieldParagraphs(sourceModel)
  if (source.length !== paragraphs.length || source.some((paragraph, index) => {
    const captured = paragraphs[index]!
    return paragraph.block.id !== captured.block.id || paragraph.absoluteStart !== captured.absoluteStart
      || paragraph.absoluteEnd !== captured.absoluteEnd || paragraph.text !== captured.text
      || JSON.stringify(paragraph.noteReferences) !== JSON.stringify(captured.noteReferences)
  })) throw new RangeError('内容控件正文编辑的源模型与捕获段落不一致。')
}

export type ContentControlPatch = Partial<Omit<DocumentContentControl, 'id' | 'blockId' | 'startOffset' | 'endOffset' | 'startNoteReferenceCount' | 'endNoteReferenceCount' | 'value' | 'checked' | 'dateValue'>>

const lockedContent = new Set<DocumentContentControl['lock']>(['contentLocked', 'sdtContentLocked'])
const lockedControl = new Set<DocumentContentControl['lock']>(['controlLocked', 'sdtLocked', 'sdtContentLocked'])

function assertControlAcceptsParagraphBreak(control: DocumentContentControl): void {
  if (!['richText', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj'].includes(control.type)
    && !(control.type === 'text' && control.multiLine === true)) throw new RangeError(`此类型内容控件不接受段落换行：${control.id}`)
}

function editHostedControl<T>(model: DocumentModel, blockId: string, edit: (scoped: DocumentModel) => T, requestedHost?: DocumentAuxiliaryBlockHost,
  complete?: (before: DocumentModel, candidate: DocumentModel) => void): { value: T } | undefined {
  const host = requestedHost ?? documentBlockHostForId(model, blockId)
  if (!host) return undefined
  const candidate = cloneModel(model), scoped = documentModelForBlockHost(candidate, host, true), value = edit(scoped)
  applyDocumentBlockHostModel(candidate, host, scoped)
  if (complete) complete(model, candidate)
  else if (JSON.stringify(documentModelForBlockHost(model, host).blocks) !== JSON.stringify(scoped.blocks)) {
    finalizeDocumentContentControlHostResources(model, candidate)
  }
  validateDocumentContentControlRanges(candidate)
  if (host.kind === 'story') assertDocumentStoryLocksPreserved(resolveSectionHeaderFooter(model, host.sectionId), resolveSectionHeaderFooter(candidate, host.sectionId), host.story, model)
  Object.assign(model, candidate)
  return { value }
}

function hostedControlBlock(model: DocumentModel, id: string): string | undefined {
  const control = model.contentControls?.find(control => control.id === id)
  return control?.blockRange?.blockIds[0] ?? control?.tableRowRange?.tableId ?? control?.blockId
}

/** A paragraph split maps the suffix of a linear replacement into its new paragraph. */
export interface ContentControlParagraphSplit {
  blockId: string
  firstBreakOffset: number
  lastParagraphOffset: number
}

export function getDocumentContentControls(model: DocumentModel): DocumentContentControl[] {
  return (model.contentControls ?? []).map((control) => cloneControl(control))
}

export function getDocumentContentControl(model: DocumentModel, id: string): DocumentContentControl | undefined {
  const control = model.contentControls?.find((candidate) => candidate.id === id)
  return control ? cloneControl(control) : undefined
}

export function selectedDocumentContentControl(model: DocumentModel, blockId: string, startOffset: number, endOffset = startOffset): DocumentContentControl | undefined {
  const target = textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === blockId)
  return target && selectedDocumentContentControlForSelection(model, { anchor: { ...target, offset: startOffset }, focus: { ...target, offset: endOffset } })
}

export function selectedDocumentContentControlForSelection(model: DocumentModel, selection: DocumentSelection, activeId?: string): DocumentContentControl | undefined {
  const controls = model.contentControls ?? []
  const targetIds = new Set(textTargets(model).map(target => target.cellBlockId ?? target.blockId))
  const sourceOrders = documentContentControlSourceOrders(controls.filter(control => targetIds.has(control.blockId)))
  const explicit = controls.find(control => control.id === activeId && !control.metadata?.officeStoryLockSource
    && documentContentControlContainsSelection(model, control, selection))
  return explicit ?? controls.filter(control => !control.metadata?.officeStoryLockSource && documentContentControlContainsSelection(model, control, selection))
    .sort((left, right) => {
      if (hasAncestor(controls, left, right.id)) return -1
      if (hasAncestor(controls, right, left.id)) return 1
      const leftContains = documentContentControlContains(model, left, right), rightContains = documentContentControlContains(model, right, left)
      if (leftContains !== rightContains) return leftContains ? 1 : -1
      if (leftContains && rightContains) {
        const leftOrder = sourceOrders.get(left.id), rightOrder = sourceOrders.get(right.id)
        if (leftOrder !== undefined && rightOrder !== undefined) return leftOrder - rightOrder
      }
      return (left.blockRange?.blockIds.length ?? 0) - (right.blockRange?.blockIds.length ?? 0)
        || (left.tableRowRange?.rowIds.length ?? 0) - (right.tableRowRange?.rowIds.length ?? 0)
        || (left.endOffset - left.startOffset) - (right.endOffset - right.startOffset)
    })[0]
}

export function createDocumentContentControl(
  model: DocumentModel,
  input: ContentControlInput,
  id = input.id ?? `content-control-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
  host?: DocumentAuxiliaryBlockHost,
): DocumentContentControl {
  if (typeof id !== 'string' || !id.trim()) throw new TypeError('Content control id must not be empty')
  if (model.contentControls?.some((control) => control.id === id)) throw new Error(`Content control already exists: ${id}`)
  const hosted = editHostedControl(model, input.blockRange?.blockIds[0] ?? input.tableRowRange?.tableId ?? input.blockId, scoped => createDocumentContentControl(scoped, input, id), host)
  if (hosted) return hosted.value
  const { id: _inputId, blockId: _blockId, startOffset: _start, endOffset: _end, startNoteReferenceCount: _startRank,
    endNoteReferenceCount: _endRank, value: _value, checked: _checked, dateValue: _dateValue, ...properties } = input
  validateContentControlPropertyPatch(properties)
  if (input.checked !== undefined && typeof input.checked !== 'boolean') throw new TypeError('复选框状态必须是布尔值。')
  if (input.value !== undefined && typeof input.value !== 'string') throw new TypeError('内容控件值必须是文本。')
  const candidate = cloneModel(model)
  const control: DocumentContentControl = cloneControl({
    ...input,
    id,
    lock: input.lock ?? 'unlocked',
    showingPlaceholder: input.showingPlaceholder ?? false,
  })
  const controls = candidate.contentControls ?? (candidate.contentControls = [])
  if ((control.tableRowRange || control.blockRange) && ['group', 'repeatingSection', 'repeatingSectionItem'].includes(control.type)) {
    if (!control.parentId) control.parentId = controls.filter(parent => ['group', 'repeatingSection', 'repeatingSectionItem'].includes(parent.type)
      && findBlock(candidate, parent.blockRange?.blockIds[0] ?? parent.tableRowRange?.tableId ?? parent.blockId) && documentContentControlContains(candidate, parent, control))
      .sort((left, right) => hasAncestor(controls, left, right.id) ? -1 : hasAncestor(controls, right, left.id) ? 1
        : (left.blockRange?.blockIds.length ?? 0) - (right.blockRange?.blockIds.length ?? 0)
          || (left.tableRowRange?.rowIds.length ?? 0) - (right.tableRowRange?.rowIds.length ?? 0))[0]?.id
    for (const child of controls) if (child.id !== control.parentId && child.parentId === control.parentId
      && findBlock(candidate, child.blockRange?.blockIds[0] ?? child.tableRowRange?.tableId ?? child.blockId) && documentContentControlContains(candidate, control, child)) child.parentId = control.id
  }
  validateControl(candidate, control)
  if (control.blockRange) reanchorDocumentBlockContentControl(candidate, control)
  assertNoPartialOverlap(candidate, controls, control)
  controls.push(control)
  controls.sort((left, right) => left.blockId.localeCompare(right.blockId) || left.startOffset - right.startOffset || left.id.localeCompare(right.id))
  if (control.type === 'checkbox') {
    assertAncestorContentEditable(candidate, control)
    control.checked = input.checked ?? input.checkbox?.checked ?? false
    control.checkbox = { ...control.checkbox, checked: control.checked }
    const glyph = contentControlCheckboxGlyph(control)
    replaceControlText(candidate, control, glyph.text, { fontFamily: glyph.fontFamily })
    control.value = control.checked ? 'true' : 'false'
    control.showingPlaceholder = false
  } else if (control.showingPlaceholder && control.placeholder && ['text', 'richText', 'date', 'dropDownList', 'comboBox', 'picture'].includes(control.type)) {
    assertAncestorContentEditable(candidate, control)
    replaceControlText(candidate, control, control.placeholder)
    control.value = control.placeholder
  }
  if (control.type === 'checkbox' || control.showingPlaceholder && control.placeholder
    && ['text', 'richText', 'date', 'dropDownList', 'comboBox', 'picture'].includes(control.type)) {
    finalizeDocumentContentControlHostResources(model, candidate)
  }
  Object.assign(model, candidate)
  return requireControl(model, id)
}

export function updateDocumentContentControl(model: DocumentModel, id: string, patch: ContentControlPatch): void {
  const blockId = hostedControlBlock(model, id)
  if (blockId && editHostedControl(model, blockId, scoped => updateDocumentContentControl(scoped, id, patch))) return
  const control = requireControl(model, id), candidate = cloneModel(model)
  validateContentControlPropertyPatch(patch, control)
  const next: DocumentContentControl = cloneControl({
    ...control,
    ...patch,
    id: control.id,
  })
  validateControl(candidate, next)
  const controls = candidate.contentControls!
  assertNoPartialOverlap(candidate, controls.filter((candidate) => candidate.id !== id), next)
  controls[controls.findIndex(control => control.id === id)] = next
  let display: { text: string; marks?: RunPatch } | undefined
  const changed = (key: keyof ContentControlPatch) => Object.hasOwn(patch, key) && JSON.stringify(control[key]) !== JSON.stringify(next[key])
  if (next.type === 'checkbox' && (changed('checkbox') || changed('type'))) {
    next.checked = patch.checkbox?.checked ?? next.checked ?? next.checkbox?.checked ?? false
    next.checkbox = { ...next.checkbox, checked: next.checked }
    const glyph = contentControlCheckboxGlyph(next)
    display = { text: glyph.text, marks: { fontFamily: glyph.fontFamily } }
    next.value = next.checked ? 'true' : 'false'
    next.showingPlaceholder = false
  } else if (next.showingPlaceholder && ['text', 'richText', 'date', 'dropDownList', 'comboBox', 'picture'].includes(next.type)
    && (changed('placeholder') || changed('showingPlaceholder'))) {
    display = { text: next.placeholder ?? '' }
    next.value = next.placeholder ?? ''
  } else if (next.type === 'date' && !next.showingPlaceholder && next.dateValue
    && (changed('dateFormat') || changed('dateLocale') || changed('dateCalendar') || changed('type'))) {
    display = { text: formatContentControlDate(next.dateValue, next.dateFormat, next.dateLocale, next.dateCalendar) }
  } else if (['dropDownList', 'comboBox'].includes(next.type) && !next.showingPlaceholder && next.value !== undefined && changed('items')) {
    display = { text: next.items?.find(item => item.value === next.value)?.displayText ?? next.value }
  }
  if (display) {
    if (isContentLocked(control) && isContentLocked(next)) throw new Error(`Content control content is locked: ${id}`)
    assertAncestorContentEditable(candidate, next)
    replaceControlText(candidate, next, display.text, display.marks)
    finalizeDocumentContentControlHostResources(model, candidate)
  }
  // A wrapper lock protects removal, not its properties. Publish after all text/field checks succeed.
  Object.assign(model, candidate)
}

export function setDocumentContentControlValue(model: DocumentModel, id: string, value: string): void {
  const blockId = hostedControlBlock(model, id)
  if (blockId && editHostedControl(model, blockId, scoped => setDocumentContentControlValue(scoped, id, value))) return
  const control = requireControl(model, id)
  if (typeof value !== 'string') throw new TypeError('内容控件值必须是文本。')
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
  if (!['text', 'richText', 'date', 'dropDownList', 'comboBox'].includes(control.type)) {
    throw new TypeError(`Content control ${id} does not accept text values`)
  }
  if (control.type === 'dropDownList' && !control.items?.some((item) => item.value === value)) {
    throw new RangeError(`Value is not present in dropdown content control: ${value}`)
  }
  if (control.type === 'text' && control.multiLine !== true && /[\r\n]/.test(value)) throw new RangeError('此纯文本控件未启用多行输入。')
  if (control.type === 'date' && value !== '' && !validContentControlIsoDate(value)) throw new RangeError('Date content control value must be YYYY-MM-DD')
  const displayValue = control.type === 'date'
    ? formatContentControlDate(value, control.dateFormat, control.dateLocale, control.dateCalendar)
    : ['dropDownList', 'comboBox'].includes(control.type)
      ? control.items?.find((item) => item.value === value)?.displayText ?? value
      : value.replace(/\r\n?/g, '\n')
  if (control.type === 'comboBox' && /[\r\n]/.test(displayValue)) throw new RangeError('组合框内容不能包含段落换行。')
  const storedValue = control.type === 'text' || control.type === 'richText' ? displayValue : value
  const previousText = contentControlText(model, control)
  const replacesContents = previousText !== displayValue || contentControlHasNoteReferences(model, control)
  if (!replacesContents && !control.showingPlaceholder && control.value === storedValue
    && (control.type !== 'date' || control.dateValue === value)) return
  const candidate = cloneModel(model), next = requireControl(candidate, id)
  if (replacesContents) replaceControlText(candidate, next, displayValue)
  next.value = storedValue
  if (next.type === 'date') next.dateValue = value
  next.showingPlaceholder = false
  if (replacesContents) finalizeDocumentContentControlHostResources(model, candidate)
  synchronizeDocumentContentControlsAfterEdit(model, candidate, { origin: 'control-value', controlId: id })
  Object.assign(model, candidate)
}

export function setDocumentContentControlChecked(model: DocumentModel, id: string, checked: boolean): void {
  const blockId = hostedControlBlock(model, id)
  if (blockId && editHostedControl(model, blockId, scoped => setDocumentContentControlChecked(scoped, id, checked))) return
  const control = requireControl(model, id)
  if (typeof checked !== 'boolean') throw new TypeError('复选框状态必须是布尔值。')
  if (control.type !== 'checkbox') throw new TypeError(`Content control ${id} is not a checkbox`)
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
  const glyph = contentControlCheckboxGlyph(control, checked), previousText = contentControlText(model, control)
  const storedValue = checked ? 'true' : 'false'
  const replacesContents = previousText !== glyph.text || contentControlHasNoteReferences(model, control)
  if (!replacesContents && !control.showingPlaceholder && control.value === storedValue
    && control.checked === checked && control.checkbox?.checked === checked) return
  const candidate = cloneModel(model), next = requireControl(candidate, id)
  next.checked = checked
  next.value = storedValue
  next.checkbox = { ...next.checkbox, checked }
  if (replacesContents) replaceControlText(candidate, next, glyph.text, { fontFamily: glyph.fontFamily })
  next.showingPlaceholder = false
  if (replacesContents) finalizeDocumentContentControlHostResources(model, candidate)
  synchronizeDocumentContentControlsAfterEdit(model, candidate, { origin: 'control-value', controlId: id })
  Object.assign(model, candidate)
}

export function resetDocumentContentControlPlaceholder(model: DocumentModel, id: string, useDefaultPlaceholder = false): void {
  const blockId = hostedControlBlock(model, id)
  if (blockId && editHostedControl(model, blockId, scoped => resetDocumentContentControlPlaceholder(scoped, id, useDefaultPlaceholder))) return
  const control = requireControl(model, id)
  const placeholder = useDefaultPlaceholder ? documentContentControlPlaceholderText(control) : control.placeholder
  if (placeholder === undefined || !useDefaultPlaceholder && !placeholder) throw new Error(`Content control has no placeholder: ${id}`)
  if (control.showingPlaceholder && !contentControlHasNoteReferences(model, control)) return
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
  const candidate = cloneModel(model), next = requireControl(candidate, id)
  resetContentControlPlaceholderCandidate(candidate, next, placeholder)
  finalizeDocumentContentControlHostResources(model, candidate)
  synchronizeDocumentContentControlsAfterEdit(model, candidate, { origin: 'control-value', controlId: id, placeholderText: placeholder })
  Object.assign(model, candidate)
}

/** Picture/field restoration can fail after structural work; callers publish this candidate only. */
function resetContentControlPlaceholderCandidate(model: DocumentModel, control: DocumentContentControl, placeholder: string): void {
  const id = control.id
  if (control.type === 'picture') {
    if (control.blockRange) {
      // Replace the real owned image tree before clearing its pointers; no stale paragraph merge.
      replaceControlText(model, control, placeholder)
      const { pictureBlockId: _pictureBlockId, pictureAnchorBlockId: _anchorBlockId, pictureTailBlockId: _tailBlockId, pictureImage: _pictureImage, ...metadata } = control.metadata ?? {}
      control.metadata = metadata
      control.value = placeholder; control.showingPlaceholder = true
      return
    }
    const pictureId = typeof control.metadata?.pictureBlockId === 'string' ? control.metadata.pictureBlockId : undefined
    const picture = pictureId ? findBlock(model, pictureId) : undefined
    const anchorId = typeof control.metadata?.pictureAnchorBlockId === 'string' ? control.metadata.pictureAnchorBlockId : undefined
    const tailId = typeof control.metadata?.pictureTailBlockId === 'string' ? control.metadata.pictureTailBlockId : undefined
    const anchor = anchorId ? findBlock(model, anchorId) : undefined
    const container = pictureId ? blockContainerContaining(model, pictureId) : undefined
    if (picture && anchor && tailId && container) {
      const tailIndex = container.findIndex(block => block.id === tailId)
      const tail = tailIndex >= 0 ? container[tailIndex] : undefined
      if (tail) {
        const prefixLength = runText(anchor.runs as DocumentRun[]).length
        transformDocumentTextFieldsForBlockRangeEdit([anchor, tail], prefixLength, 0, '', [anchor], undefined, model)
        anchor.runs = compactRuns([...(anchor.runs as DocumentRun[]), ...(tail.runs as DocumentRun[])])
        for (const candidate of model.contentControls ?? []) {
          if (candidate.blockId !== tail.id) continue
          candidate.blockId = anchor.id
          candidate.startOffset += prefixLength
          candidate.endOffset += prefixLength
        }
        const pictureIndex = container.findIndex(block => block.id === picture.id)
        for (const index of [pictureIndex, tailIndex].filter(index => index >= 0).sort((left, right) => right - left)) container.splice(index, 1)
      } else {
        const pictureIndex = container.findIndex(block => block.id === picture.id)
        if (pictureIndex >= 0) container.splice(pictureIndex, 1)
      }
    } else if (picture && container) {
      const pictureIndex = container.findIndex(block => block.id === picture.id)
      if (pictureIndex >= 0) container.splice(pictureIndex, 1)
    }
    const { pictureBlockId: _pictureBlockId, pictureAnchorBlockId: _anchorBlockId, pictureTailBlockId: _tailBlockId, pictureImage: _pictureImage, ...metadata } = control.metadata ?? {}
    control.metadata = metadata
    replaceControlText(model, control, placeholder)
    control.value = placeholder
    control.showingPlaceholder = true
    return
  }
  if (!['text', 'richText', 'date', 'dropDownList', 'comboBox'].includes(control.type)) {
    throw new TypeError(`Content control ${id} does not accept text placeholders`)
  }
  replaceControlText(model, control, placeholder)
  control.value = placeholder
  control.showingPlaceholder = true
}

/** Public SDT fallback order; a document-part name is not the displayed prompt text. */
export function documentContentControlPlaceholderText(control: DocumentContentControl): string {
  return (control.placeholder === control.placeholderDocPart?.name ? undefined : control.placeholder)
    ?? control.placeholderDocPart?.text ?? '\u200b'
}

export function replacePictureContentControlImage(model: DocumentModel, id: string, image: DocumentImage): string {
  const blockId = hostedControlBlock(model, id), hosted = blockId && editHostedControl(model, blockId, scoped => replacePictureContentControlImage(scoped, id, image))
  if (hosted) return hosted.value
  const candidate = cloneModel(model), pictureId = replacePictureContentControlImageInModel(candidate, id, image)
  finalizeDocumentContentControlHostResources(model, candidate)
  synchronizeDocumentContentControlsAfterEdit(model, candidate, { origin: 'control-value', controlId: id })
  if (!candidate.contentControls?.some(control => control.id === id)) {
    const picture = findBlock(candidate, pictureId)
    if (picture?.metadata.contentControlId === id) delete picture.metadata.contentControlId
  }
  validateDocumentContentControlRanges(candidate)
  Object.assign(model, candidate)
  return pictureId
}

function replacePictureContentControlImageInModel(model: DocumentModel, id: string, image: DocumentImage): string {
  const control = requireControl(model, id)
  if (control.type !== 'picture') throw new TypeError(`Content control ${id} is not a picture control`)
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
  if (control.blockRange) {
    const pictureId = replaceDocumentBlockContentControlImage(model, control, image)
    control.value = undefined; control.showingPlaceholder = false
    return pictureId
  }
  const pointer = control.metadata?.pictureBlockId
  if (pointer !== undefined && (typeof pointer !== 'string' || !pointer)) throw new TypeError('图片内容控件的对象身份无效：' + id)
  const existingPictureId = typeof control.metadata?.pictureBlockId === 'string' ? control.metadata.pictureBlockId : undefined
  const existingPicture = existingPictureId ? findBlock(model, existingPictureId) : undefined
  if (existingPictureId && (!existingPicture || existingPicture.type !== 'image')) throw new RangeError('图片内容控件的对象不是实际图片：' + id)
  if (existingPicture?.type === 'image') {
    for (const other of model.contentControls ?? []) if (other.id !== id && other.type === 'picture' && other.metadata?.pictureBlockId === existingPicture.id) {
      if (isContentLocked(other)) throw new RangeError('图片替换不能改变锁定子控件所引用的图片：' + other.id)
      throw new RangeError('图片对象不能被多个图片内容控件同时拥有：' + existingPicture.id)
    }
    existingPicture.metadata = { ...existingPicture.metadata, image: { ...image }, pictureInputAnchor: true }
    control.metadata = { ...control.metadata, pictureImage: { ...image } }
    control.value = undefined; control.showingPlaceholder = false
    return existingPicture.id
  }
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control paragraph not found: ${control.blockId}`)
  const block = findBlock(model, control.blockId)
  const container = block && blockContainerForTarget(model, target)
  const index = container?.findIndex(candidate => candidate.id === block?.id) ?? -1
  if (!block || !container || index < 0) throw new Error(`Content control paragraph not found: ${control.blockId}`)
  const runs = runsForTarget(model, target)
  const { startOffset, endOffset } = control
  const controls = model.contentControls ?? []
  const nested = controls.filter(candidate => candidate.id !== control.id
    && candidate.blockId === control.blockId && hasAncestor(controls, candidate, control.id))
  if (nested.some(candidate => isControlLocked(candidate) || isContentLocked(candidate))) throw new Error('A nested content control is locked')
  const removedIds = new Set(nested.map(candidate => candidate.id))
  for (const candidate of controls) {
    if (candidate.id === control.id || removedIds.has(candidate.id) || candidate.blockId !== control.blockId) continue
    if (candidate.startOffset <= startOffset && candidate.endOffset >= endOffset
      && (candidate.startOffset < startOffset || candidate.endOffset > endOffset)) {
      throw new RangeError(`Picture replacement cannot split an enclosing content control: ${candidate.id}`)
    }
    if (candidate.startOffset < endOffset && startOffset < candidate.endOffset
      && !hasAncestor(controls, control, candidate.id)) throw new RangeError(`Picture replacement overlaps another content control: ${candidate.id}`)
  }
  const tail = createParagraph()
  tail.style = { ...block.style }
  tail.type = block.type
  tail.level = block.level
  tail.metadata = { ...block.metadata }
  delete tail.metadata.html
  delete tail.metadata.documentTextFields
  tail.runs = sliceRuns(runs, endOffset)
  const fieldSource = { ...block, metadata: { ...block.metadata } }
  setDocumentTextFields(fieldSource, documentTextFields(block).filter(field => field.endBlockId !== undefined && field.endBlockId !== block.id || field.startOffset < startOffset || field.endOffset > endOffset))
  transformDocumentTextFieldsForBlockRangeEdit([fieldSource], startOffset, endOffset, '\n', [fieldSource, tail], undefined, model)
  block.metadata = fieldSource.metadata
  model.contentControls = controls.filter(candidate => !removedIds.has(candidate.id))
  setTargetRuns(model, target, sliceRuns(runs, 0, startOffset))
  const picture = createParagraph()
  picture.type = 'image'
  picture.runs = []
  picture.metadata = { image: { ...image }, contentControlId: control.id, pictureInputAnchor: true, anchorBlockId: block.id, anchorOffset: startOffset, tailBlockId: tail.id }
  const pictureId = picture.id
  container.splice(index + 1, 0, picture, tail)
  for (const candidate of model.contentControls ?? []) {
    if (candidate.id === control.id || candidate.blockId !== control.blockId) continue
    if (hasAncestor(model.contentControls ?? [], control, candidate.id)) candidate.endOffset = startOffset
    else if (candidate.startOffset >= endOffset) {
      candidate.blockId = tail.id
      candidate.startOffset -= endOffset
      candidate.endOffset -= endOffset
    }
  }
  control.endOffset = control.startOffset
  control.value = undefined
  control.showingPlaceholder = false
  control.metadata = {
    ...control.metadata,
    pictureBlockId: pictureId,
    pictureAnchorBlockId: block.id,
    pictureTailBlockId: tail.id,
    pictureImage: { ...image },
  }
  return pictureId
}

/** The toolbar and mutations share the same placement, range and nested-lock rules. */
export function assertDocumentRepeatingSectionItemEditable(model: DocumentModel, id: string, action: 'duplicate' | 'delete'): void {
  const item = requireControl(model, id), host = documentBlockHostForId(model, item.blockRange?.blockIds[0] ?? item.tableRowRange?.tableId ?? item.blockId)
  if (host) { assertDocumentRepeatingSectionItemEditable(documentModelForBlockHost(model, host, true), id, action); return }
  if (item.placement === 'row') { assertDocumentRepeatingContentControlRowsEditable(model, id, action); return }
  if (item.blockRange) { assertDocumentRepeatingContentControlBlocksEditable(model, id, action); return }
  if (item.type !== 'repeatingSectionItem') throw new TypeError(`Content control ${id} is not a repeating-section item`)
  if ((action === 'duplicate' ? ['row', 'cell', 'table'] : ['cell', 'table']).includes(item.placement ?? '')) {
    throw new TypeError(`This repeating-section placement cannot be ${action === 'duplicate' ? 'duplicated' : 'deleted'}`)
  }
  const controls = model.contentControls ?? [], section = repeatingSectionAncestor(controls, item)
  if (!section) throw new Error(`Repeating-section parent is missing: ${id}`)
  if (section.repeatingSection?.doNotAllowInsertDelete) throw new Error(`Repeating-section items cannot be changed: ${section.id}`)
  const target = targetForControl(model, item)
  if (!target) throw new Error(`Content control paragraph not found: ${item.blockId}`)
  const runs = runsForTarget(model, target), text = runText(runs), references = documentNoteReferenceOffsets(runs)
  const itemRange = documentContentControlNoteBoundaryRanks(item, references)
  const itemControls = repeatingSectionItemControls(controls, item, references)
  const locked = action === 'duplicate' ? isControlLocked(item) || isContentLocked(item) || isContentLocked(section)
    : itemControls.some(control => isControlLocked(control) || isContentLocked(control))
  if (locked) throw new Error(`Repeating-section item is locked: ${id}`)
  assertAncestorContentEditable(model, item)
  if (item.startOffset > item.endOffset || !isDocumentTextBoundary(text, item.startOffset) || !isDocumentTextBoundary(text, item.endOffset)) {
    throw new RangeError(`Repeating-section item range is invalid: ${id}`)
  }
  if (action === 'delete') for (const control of controls) {
    if (control.id === item.id || control.blockId !== item.blockId || control.blockRange || control.tableRowRange
      || !documentContentControlNoteRangesOverlap(control, item, references)) continue
    if (hasAncestor(controls, item, control.id) || documentContentControlNoteRangeContains(item, control, references)) continue
    throw new RangeError(`Repeating-section deletion cannot split another content control: ${control.id}`)
  }
  const paragraph = documentTextFieldParagraphs(model).find(paragraph => paragraph.block.id === item.blockId)!
  const start = paragraph.absoluteStart + item.startOffset, end = paragraph.absoluteStart + item.endOffset
  const source = documentRepeatingContentControlInlineCopySource(paragraph.block, item.startOffset, item.endOffset, itemRange.start, itemRange.end)
  const selectedFields = new Set(documentTextFields(source).map(field => field.id))
  for (const field of documentTextFieldRanges(model)) {
    const point = compareDocumentNoteBoundaries(field.absoluteStart, field.startNoteReferenceCount, field.absoluteEnd, field.endNoteReferenceCount) === 0
    const touches = selectedFields.has(field.field.id) || (point
      ? compareDocumentNoteBoundaries(field.absoluteStart, field.startNoteReferenceCount, start, itemRange.start) > 0
        && compareDocumentNoteBoundaries(field.absoluteEnd, field.endNoteReferenceCount, end, itemRange.end) < 0
      : compareDocumentNoteBoundaries(field.absoluteStart, field.startNoteReferenceCount, end, itemRange.end) < 0
        && compareDocumentNoteBoundaries(start, itemRange.start, field.absoluteEnd, field.endNoteReferenceCount) < 0)
    if (!touches) continue
    if (action === 'duplicate' && (compareDocumentNoteBoundaries(field.absoluteStart, field.startNoteReferenceCount, start, itemRange.start) < 0
      || compareDocumentNoteBoundaries(field.absoluteEnd, field.endNoteReferenceCount, end, itemRange.end) > 0)) throw new RangeError('复制重复项目不能切开跨项目字段：' + field.field.id)
    if (action === 'delete' && documentFieldLocked(field.field)) throw new RangeError('删除重复项目不能改变锁定字段：' + field.field.id)
  }
  if (action === 'duplicate' && text.length + item.endOffset - item.startOffset > 16 * 1024 * 1024) throw new RangeError('重复项目复制超过段落文字预算。')
  const runIds = new Set(source.runs.map(run => run.id))
  if (action === 'delete') for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(model[kind] ?? {})) {
    if (note.referenceBlockId === item.blockId && runIds.has(note.referenceRunId)) assertDocumentRepeatingNoteContentRemovable(model, note)
  }
  if (action === 'duplicate' && itemControls.some(control => control.type === 'picture'
    && typeof control.metadata?.pictureBlockId === 'string' && findBlock(model, control.metadata.pictureBlockId))) {
    throw new RangeError('此旧行内重复项目的图片位于文字范围之外，需要先归一化为完整块范围。')
  }
}

function repeatingSectionItemControls(controls: readonly DocumentContentControl[], item: DocumentContentControl,
  references: readonly DocumentNoteReferenceOffset[]): DocumentContentControl[] {
  return controls.filter(control => control.blockId === item.blockId && documentContentControlNoteRangeContains(item, control, references)
    && (control.id === item.id || hasAncestor(controls, control, item.id)))
}

export function duplicateRepeatingSectionItem(model: DocumentModel, id: string, permission?: DocumentPermissionContext,
  identityMap?: DocumentRepeatingContentControlCopyIds): DocumentContentControl {
  const blockId = hostedControlBlock(model, id), host = blockId && documentBlockHostForId(model, blockId)
  if (host) {
    const item = requireControl(model, id), ids = createDocumentRepeatingContentControlCopyIds(), scoped = documentModelForBlockHost(model, host, true)
    if (permission) {
      const targets = documentContentControlPermissionTargets(scoped, item, model)
      assertDocumentPermission(model.permissionRules ?? [], permission, 'edit', targets)
      assertDocumentPermission(model.permissionRules ?? [], permission, 'copy', targets)
    }
    const inline = !item.blockRange && !item.tableRowRange ? { blockId: item.blockId, startOffset: item.startOffset, endOffset: item.endOffset,
      startNoteReferenceCount: item.startNoteReferenceCount, endNoteReferenceCount: item.endNoteReferenceCount,
      insertionOffset: item.endOffset, insertionNoteReferenceCount: item.endNoteReferenceCount } : undefined
    const hosted = editHostedControl(model, blockId!, view => duplicateRepeatingSectionItem(view, id, permission, ids), host, (before, candidate) => {
      finalizeDocumentContentControlHostResources(before, candidate, ids, inline)
      if (permission) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', permissionTargetsForChangedBlocks(before, candidate))
    })!
    const result = cloneControl(requireControl(model, hosted.value.id))
    if (identityMap) for (const key of Object.keys(identityMap) as Array<keyof DocumentRepeatingContentControlCopyIds>) {
      for (const [oldId, newId] of ids[key]) identityMap[key].set(oldId, newId)
    }
    return result
  }
  const item = requireControl(model, id)
  if (item.placement === 'row') {
    const plan = prepareDocumentRepeatingContentControlRowsMutation(model, id, 'duplicate', permission)
    const copy = plan.model.contentControls?.find(control => control.id === plan.controlId)
    if (!copy) throw new RangeError('复制的重复行控件不存在。')
    Object.assign(model, plan.model)
    if (identityMap) for (const key of Object.keys(identityMap) as Array<keyof DocumentRepeatingContentControlCopyIds>) {
      for (const [oldId, newId] of plan.identityMap[key]) identityMap[key].set(oldId, newId)
    }
    return cloneControl(copy)
  }
  if (item.blockRange) {
    const plan = prepareDocumentRepeatingContentControlBlocksMutation(model, id, 'duplicate', permission)
    const copy = plan.model.contentControls?.find(control => control.id === plan.controlId)
    if (!copy) throw new RangeError('复制的重复块控件不存在。')
    Object.assign(model, plan.model)
    if (identityMap) for (const key of Object.keys(identityMap) as Array<keyof DocumentRepeatingContentControlCopyIds>) {
      for (const [oldId, newId] of plan.identityMap[key]) identityMap[key].set(oldId, newId)
    }
    return cloneControl(copy)
  }
  assertDocumentRepeatingSectionItemEditable(model, id, 'duplicate')
  const destination = model, before = model, ids = createDocumentRepeatingContentControlCopyIds()
  if (permission) {
    const targets = documentContentControlPermissionTargets(model, item)
    assertDocumentPermission(model.permissionRules ?? [], permission, 'edit', targets)
    assertDocumentPermission(model.permissionRules ?? [], permission, 'copy', targets)
  }
  model = cloneModel(model)
  const target = targetForControl(model, item)
  if (!target) throw new Error(`Content control paragraph not found: ${item.blockId}`)
  const runs = runsForTarget(model, target)
  const textLength = item.endOffset - item.startOffset
  const insertionOffset = item.endOffset
  const paragraph = findBlock(model, item.blockId)!
  const sourceNoteReferences = documentNoteReferenceOffsets(runs), itemRange = documentContentControlNoteBoundaryRanks(item, sourceNoteReferences)
  const source = documentRepeatingContentControlInlineCopySource(paragraph, item.startOffset, item.endOffset, itemRange.start, itemRange.end)
  const copied = copyDocumentBlocksForPaste([source])[0]!
  indexDocumentRepeatingContentControlCopies([source], [copied], ids)
  // The paragraph shell is borrowed, never inserted. The new runs belong to the existing real host.
  ids.block.set(paragraph.id, paragraph.id)
  const { before: prefix, after: suffix } = partitionDocumentContentControlNoteRuns(runs,
    { startOffset: insertionOffset, endOffset: insertionOffset, startNoteReferenceCount: itemRange.end, endNoteReferenceCount: itemRange.end })
  const futureRuns = [...prefix, ...copied.runs, ...suffix], futureReferences = documentNoteReferenceOffsets(futureRuns)
  const section = repeatingSectionAncestor(model.contentControls ?? [], item)!
  transformContentControlsForBlockEdit(model, paragraph.id, insertionOffset, insertionOffset, textLength, undefined, true,
    { activeControlId: section.id, projectedNoteReferences: futureReferences, sourceNoteRanks: { start: itemRange.end, end: itemRange.end } })
  transformDocumentTextFieldsForBlockEdit(paragraph, insertionOffset, insertionOffset, runText(copied.runs), undefined, model, futureReferences,
    { start: itemRange.end, end: itemRange.end })
  paragraph.runs = futureRuns
  refreshContentControlParagraphText(model, paragraph)
  insertDocumentTextFields(paragraph, documentTextFields(copied), insertionOffset, documentNoteReferenceOffsets(copied.runs))
  transformDocumentCommentAnchorsForTextEdit(model, paragraph.id, insertionOffset, insertionOffset, runText(copied.runs), Date.now(), sourceNoteReferences)
  copyDocumentRepeatingContentControlNotes(before, model, ids)
  const controls = model.contentControls ?? []
  const itemControls = repeatingSectionItemControls(before.contentControls ?? [], item, sourceNoteReferences)
  const descendants = itemControls.filter(control => !control.metadata?.officeStoryLockSource)
  const idMap = new Map(descendants.map((control) => [control.id, newContentControlId()]))
  for (const [oldId, newId] of idMap) ids.control.set(oldId, newId)
  const prefixReferenceCount = documentNoteReferenceOffsets(prefix).length
  const copies = descendants.map((control) => {
    const ranks = documentContentControlNoteBoundaryRanks(control, sourceNoteReferences)
    return { ...cloneControl(control),
      id: idMap.get(control.id)!,
      blockId: ids.block.get(control.blockId)!,
      startOffset: control.blockId === item.blockId ? insertionOffset + control.startOffset - item.startOffset : control.startOffset,
      endOffset: control.blockId === item.blockId ? insertionOffset + control.endOffset - item.startOffset : control.endOffset,
      startNoteReferenceCount: prefixReferenceCount + ranks.start - itemRange.start,
      endNoteReferenceCount: prefixReferenceCount + ranks.end - itemRange.start,
      parentId: control.id === item.id ? item.parentId : idMap.get(control.parentId ?? '') ?? control.parentId,
      ...(control.blockRange ? { blockRange: { blockIds: control.blockRange.blockIds.map(id => ids.block.get(id)!) } } : {}),
      ...(control.tableRowRange ? { tableRowRange: { tableId: ids.block.get(control.tableRowRange.tableId)!, rowIds: control.tableRowRange.rowIds.map(id => ids.row.get(id)!) } } : {}),
      metadata: remapDocumentRepeatingContentControlMetadata(control.metadata, ids),
    }
  })
  controls.push(...copies)
  copyDocumentRepeatingContentControlNoteControls(before, model, ids)
  model.contentControls!.sort((left, right) => left.blockId.localeCompare(right.blockId) || left.startOffset - right.startOffset || left.id.localeCompare(right.id))
  const newBlockIds = new Set([...ids.block.values()].filter(blockId => blockId !== paragraph.id))
  for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(model[kind] ?? {})) if (note.content?.blocks) {
    visitDocumentRepeatingContentControlBlocks(note.content.blocks, block => {
      if (newBlockIds.has(block.id)) block.metadata = remapDocumentRepeatingContentControlMetadata(block.metadata, ids)!
    })
  }
  copyDocumentRepeatingContentControlComments(before, model, ids, { blockId: paragraph.id, startOffset: item.startOffset, endOffset: item.endOffset,
    startNoteReferenceCount: itemRange.start, endNoteReferenceCount: itemRange.end, insertionOffset, insertionNoteReferenceCount: itemRange.end })
  reconcileDocumentContentControlBlockEdits(before, model, { selection: { anchor: { ...target, offset: item.startOffset }, focus: { ...target, offset: item.endOffset } }, activeControlId: id, absorbAdjacentInsertions: false })
  reconcileDocumentContentControlTableRows(before, model)
  const live = requireControl(model, idMap.get(item.id)!), selection = { anchor: { ...target, offset: live.startOffset }, focus: { ...target, offset: live.endOffset } }
  reconcileDocumentTableAnchors(before, model, selection, new Set(documentTextFieldRanges(before).map(range => range.field.id)))
  assertRepeatingItemGroupLocksPreserved(before, model, item)
  if (permission) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', permissionTargetsForChangedBlocks(before, model))
  const result = cloneControl(requireControl(model, idMap.get(item.id)!))
  Object.assign(destination, model)
  if (identityMap) for (const key of Object.keys(identityMap) as Array<keyof DocumentRepeatingContentControlCopyIds>) {
    for (const [oldId, newId] of ids[key]) identityMap[key].set(oldId, newId)
  }
  return result
}

export function deleteRepeatingSectionItem(model: DocumentModel, id: string, permission?: DocumentPermissionContext): void {
  const blockId = hostedControlBlock(model, id)
  const host = blockId && documentBlockHostForId(model, blockId)
  if (host) {
    if (permission) assertDocumentPermission(model.permissionRules ?? [], permission, 'edit',
      documentContentControlPermissionTargets(documentModelForBlockHost(model, host, true), requireControl(model, id), model))
    editHostedControl(model, blockId!, scoped => deleteRepeatingSectionItem(scoped, id, permission), host, (before, candidate) => {
      finalizeDocumentContentControlHostResources(before, candidate)
      if (permission) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', permissionTargetsForChangedBlocks(before, candidate))
    })
    return
  }
  const item = requireControl(model, id)
  if (item.placement === 'row') {
    const plan = prepareDocumentRepeatingContentControlRowsMutation(model, id, 'delete', permission)
    Object.assign(model, plan.model)
    return
  }
  if (item.blockRange) {
    const plan = prepareDocumentRepeatingContentControlBlocksMutation(model, id, 'delete', permission)
    Object.assign(model, plan.model)
    return
  }
  assertDocumentRepeatingSectionItemEditable(model, id, 'delete')
  const destination = model, before = model
  if (permission) assertDocumentPermission(model.permissionRules ?? [], permission, 'edit', documentContentControlPermissionTargets(model, item))
  model = cloneModel(model)
  const target = targetForControl(model, item)
  if (!target) throw new Error(`Content control paragraph not found: ${item.blockId}`)
  const runs = runsForTarget(model, target)
  const startOffset = item.startOffset
  const endOffset = item.endOffset
  const paragraph = findBlock(model, item.blockId)!
  const sourceNoteReferences = documentNoteReferenceOffsets(runs), itemRange = documentContentControlNoteBoundaryRanks(item, sourceNoteReferences)
  const removed = repeatingSectionItemControls(model.contentControls ?? [], item, sourceNoteReferences)
  const removedFields = new Set(documentTextFields(documentRepeatingContentControlInlineCopySource(paragraph, startOffset, endOffset,
    itemRange.start, itemRange.end)).map(field => field.id))
  const { before: prefix, after: suffix } = partitionDocumentContentControlNoteRuns(runs, item)
  const futureRuns = [...prefix, ...suffix], futureReferences = documentNoteReferenceOffsets(futureRuns)
  const removedIds = new Set(removed.map(control => control.id))
  model.contentControls = (model.contentControls ?? []).filter(control => !removedIds.has(control.id))
  setDocumentTextFields(paragraph, documentTextFields(paragraph).filter(field => !removedFields.has(field.id)))
  transformContentControlsForBlockEdit(model, paragraph.id, startOffset, endOffset, 0, undefined, true,
    { projectedNoteReferences: futureReferences, sourceNoteRanks: itemRange })
  transformDocumentTextFieldsForBlockEdit(paragraph, startOffset, endOffset, '', undefined, model, futureReferences, itemRange)
  paragraph.runs = futureRuns
  if (!paragraph.runs.length) paragraph.runs = [createRun()]
  refreshContentControlParagraphText(model, paragraph)
  transformDocumentCommentAnchorsForTextEdit(model, paragraph.id, startOffset, endOffset, '', Date.now(), sourceNoteReferences)
  removeDeletedDocumentRepeatingNoteScopes(before, model)
  reconcileDocumentContentControlBlockEdits(before, model, { selection: { anchor: { ...target, offset: startOffset }, focus: { ...target, offset: endOffset } }, activeControlId: id, absorbAdjacentInsertions: false })
  reconcileDocumentContentControlTableRows(before, model)
  reconcileDocumentTableAnchors(before, model, { anchor: { ...target, offset: startOffset }, focus: { ...target, offset: startOffset } }, new Set(documentTextFieldRanges(before).map(range => range.field.id)))
  assertRepeatingItemGroupLocksPreserved(before, model, item)
  if (permission) assertDocumentPermission(before.permissionRules ?? [], permission, 'edit', permissionTargetsForChangedBlocks(before, model))
  Object.assign(destination, model)
}

function assertRepeatingItemGroupLocksPreserved(before: DocumentModel, after: DocumentModel, item: DocumentContentControl): void {
  for (const group of before.contentControls ?? []) if (group.type === 'group' && isContentLocked(group)
    && hasAncestor(before.contentControls ?? [], item, group.id) && !documentContentControlGroupEditAllowed(before, after, group, { activeControlId: item.id })) {
    throw new RangeError('重复项目编辑不能改变锁定分组的范围外内容。')
  }
}

export function selectDocumentContentControlItem(model: DocumentModel, id: string, value: string): void {
  const blockId = hostedControlBlock(model, id)
  if (blockId && editHostedControl(model, blockId, scoped => selectDocumentContentControlItem(scoped, id, value))) return
  const control = requireControl(model, id)
  if (control.type !== 'dropDownList' && control.type !== 'comboBox') throw new TypeError(`Content control ${id} is not a list control`)
  const item = control.items?.find((candidate) => candidate.value === value)
  if (!item && control.type === 'dropDownList') throw new RangeError(`Value is not present in dropdown content control: ${value}`)
  setDocumentContentControlValue(model, id, value)
}

export function removeDocumentContentControl(model: DocumentModel, id: string): void {
  const blockId = hostedControlBlock(model, id)
  if (blockId && editHostedControl(model, blockId, scoped => removeDocumentContentControl(scoped, id))) return
  const control = requireControl(model, id)
  if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${id}`)
  const removing = new Set([id])
  let changed = true
  while (changed) {
    changed = false
    for (const candidate of model.contentControls ?? []) {
      if (candidate.parentId && removing.has(candidate.parentId) && !removing.has(candidate.id)) {
        if (isControlLocked(candidate)) throw new Error(`Nested content control wrapper is locked: ${candidate.id}`)
        removing.add(candidate.id)
        changed = true
      }
    }
  }
  model.contentControls = (model.contentControls ?? []).filter((candidate) => !removing.has(candidate.id))
}

export function contentControlText(model: DocumentModel, control: DocumentContentControl): string {
  const host = documentBlockHostForId(model, control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)
  if (host) return contentControlText(documentModelForBlockHost(model, host), control)
  if (control.blockRange) return documentContentControlBlockTextTargets(model, control.blockRange)
    .map(target => runText(runsForTarget(model, target))).join('\n')
  if (control.tableRowRange) {
    const owned = documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(model, control.tableRowRange))
    return textTargets(model).filter(target => owned.has(target.cellBlockId ?? target.blockId)).map(target => runText(runsForTarget(model, target))).join('\n')
  }
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control paragraph not found: ${control.blockId}`)
  documentContentControlNoteBoundaryRanks(control, documentNoteReferenceOffsets(runsForTarget(model, target)))
  return runText(sliceRuns(runsForTarget(model, target), control.startOffset, control.endOffset))
}

/** A scalar value assignment replaces selected note entities even when their
 * labels contribute no change to the displayed text. Plain equal text stays a no-op. */
function contentControlHasNoteReferences(model: DocumentModel, control: DocumentContentControl): boolean {
  if (control.blockRange || control.tableRowRange) {
    const ownedRows = control.tableRowRange ? documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(model, control.tableRowRange)) : undefined
    const targets = control.blockRange ? documentContentControlBlockTextTargets(model, control.blockRange)
      : textTargets(model).filter(target => ownedRows?.has(target.cellBlockId ?? target.blockId))
    return targets.some(target => runsForTarget(model, target).some(documentRunHasNoteReference))
  }
  const target = targetForControl(model, control)
  if (!target) throw new RangeError('内容控件的注释引用宿主不存在：' + control.id)
  const references = documentNoteReferenceOffsets(runsForTarget(model, target)), range = documentContentControlNoteBoundaryRanks(control, references)
  return references.some((reference, index) => compareDocumentNoteBoundaries(reference.startOffset, index, control.endOffset, range.end) < 0
    && compareDocumentNoteBoundaries(control.startOffset, range.start, reference.endOffset, index + 1) < 0)
}

export function contentControlRange(control: DocumentContentControl): Pick<DocumentContentControl,
  'blockId' | 'startOffset' | 'endOffset' | 'startNoteReferenceCount' | 'endNoteReferenceCount'> {
  return { blockId: control.blockId, startOffset: control.startOffset, endOffset: control.endOffset,
    ...(control.startNoteReferenceCount === undefined ? {} : { startNoteReferenceCount: control.startNoteReferenceCount }),
    ...(control.endNoteReferenceCount === undefined ? {} : { endNoteReferenceCount: control.endNoteReferenceCount }) }
}

/** Shared capability check for Canvas controls and the actual value commands. */
export function assertDocumentContentControlValueEditable(model: DocumentModel, id: string): void {
  const control = requireControl(model, id), host = documentBlockHostForId(model, control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)
  if (host) { assertDocumentContentControlValueEditable(documentModelForBlockHost(model, host, true), id); return }
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
}

export function assertTextRangeEditable(model: DocumentModel, startOffset: number, endOffset: number): void {
  const body = flattenBodySegments(model)
  for (const control of model.contentControls ?? []) {
    if (!isContentLocked(control)) continue
    if (control.tableRowRange || control.blockRange) {
      if (!findBlock(model, control.blockRange?.blockIds[0] ?? control.tableRowRange!.tableId)) continue
      const owned = control.blockRange ? documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange))
        : documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(model, control.tableRowRange!))
      for (const segment of body) {
        const blockId = segment.cellBlockId ?? segment.blockId
        if (!owned.has(blockId)) continue
        const overlaps = startOffset === endOffset ? startOffset >= segment.startOffset && startOffset <= segment.endOffset
          : startOffset < segment.endOffset && segment.startOffset < endOffset
        if (overlaps && isStructuralContentControlLockedAtBlockEdit(model, control, blockId,
          Math.max(0, startOffset - segment.startOffset), Math.min(segment.endOffset - segment.startOffset, endOffset - segment.startOffset))) {
          throw new Error(`Content control row content is locked: ${control.id}`)
        }
      }
      continue
    }
    const segment = body.find((candidate) => (candidate.cellBlockId ?? candidate.blockId) === control.blockId)
    if (!segment) continue
    const controlStart = segment.startOffset + control.startOffset
    const controlEnd = segment.startOffset + control.endOffset
    const blocked = isTextEditBlockedByControl(model, control, startOffset - segment.startOffset, endOffset - segment.startOffset, startOffset === endOffset, controlStart, controlEnd, startOffset, endOffset)
    if (blocked) throw new Error(`Content control content is locked: ${control.id}`)
  }
}

/** Keeps paragraph-local SDT anchors aligned with edits made through Docs text commands. */
export function transformContentControlsForTextEdit(
  model: DocumentModel,
  segments: readonly { blockId: string; cellBlockId?: string; startOffset: number; endOffset: number }[],
  startOffset: number,
  endOffset: number,
  replacementLength: number,
): void {
  if (!model.contentControls?.length || (startOffset === endOffset && replacementLength === 0)) return
  const startSegment = segmentAt(segments, startOffset)
  const endSegment = segmentAt(segments, endOffset)
  if (!startSegment || !endSegment) return
  if (startSegment === endSegment) {
    transformContentControlsForBlockEdit(model, startSegment.cellBlockId ?? startSegment.blockId,
      startOffset - startSegment.startOffset, endOffset - startSegment.startOffset, replacementLength)
    return
  }
  const first = segments.indexOf(startSegment)
  const last = segments.indexOf(endSegment)
  const affectedIds = new Set(segments.slice(first, last + 1).map((segment) => segment.cellBlockId ?? segment.blockId))
  const control = model.contentControls.find((candidate) => !candidate.tableRowRange && !candidate.blockRange && affectedIds.has(candidate.blockId)
    && !((startSegment.cellBlockId ?? startSegment.blockId) === candidate.blockId
      && candidate.endOffset <= startOffset - startSegment.startOffset))
  if (control) throw new RangeError(`A text edit cannot cross a content control paragraph: ${control.id}`)
}

/** Local-offset variant for the editor's selection command path. */
export function transformContentControlsForBlockEdit(
  model: DocumentModel,
  blockId: string,
  startOffset: number,
  endOffset: number,
  replacementLength: number,
  split?: ContentControlParagraphSplit,
  preserveControlRanges = false,
  options: ContentControlRangeEditOptions = {},
): void {
  if (!model.contentControls?.length || (startOffset === endOffset && replacementLength === 0 && !split
    && options.projectedNoteReferences === undefined && options.sourceNoteRanks === undefined)) return
  const block = findBlock(model, blockId)
  if (!block) throw new RangeError('内容控件编辑的源段落不存在：' + blockId)
  const sourceText = runText(block.runs), sourceReferences = documentNoteReferenceOffsets(block.runs)
  if (!isDocumentTextBoundary(sourceText, startOffset) || !isDocumentTextBoundary(sourceText, endOffset) || endOffset < startOffset
    || !Number.isSafeInteger(replacementLength) || replacementLength < 0) throw new RangeError('内容控件编辑的文字位置或替换长度无效。')
  const projectedText = sourceText.slice(0, startOffset) + ' '.repeat(replacementLength) + sourceText.slice(endOffset)
  const projectedReferences = options.projectedNoteReferences ?? documentNoteReferenceOffsetsAfterTextEdit(sourceReferences, startOffset, endOffset, replacementLength, options.sourceNoteRanks)
  assertDocumentNoteReferenceOffsets(projectedReferences, projectedText)
  let editStartRank = assertDocumentNoteReferenceBoundaryRank(sourceReferences, startOffset, options.sourceNoteRanks?.start)
  let editEndRank = assertDocumentNoteReferenceBoundaryRank(sourceReferences, endOffset, options.sourceNoteRanks?.end)
  const derivedLabel = options.noteReferenceLabelRunId !== undefined
  const emptyRemoval = !derivedLabel && startOffset === endOffset
    ? documentNoteReferenceEmptyRemovalRange(sourceReferences, new Set(projectedReferences.map(reference => reference.runId)), startOffset) : undefined
  if (derivedLabel) {
    const index = sourceReferences.findIndex(reference => reference.runId === options.noteReferenceLabelRunId)
    const reference = sourceReferences[index], delta = replacementLength - (endOffset - startOffset)
    if (!reference || split || reference.startOffset !== startOffset || reference.endOffset !== endOffset
      || projectedReferences.length !== sourceReferences.length || sourceReferences.some((entry, position) => {
        const output = projectedReferences[position]
        return !output || output.runId !== entry.runId
          || output.startOffset !== (position <= index ? entry.startOffset : entry.startOffset + delta)
          || output.endOffset !== (position < index ? entry.endOffset : entry.endOffset + delta)
      })) throw new RangeError('派生注释标签编辑必须保留同一引用 run 及完整的段落引用顺序。')
    editStartRank = index; editEndRank = index + 1
  } else if (emptyRemoval) { editStartRank = emptyRemoval.startRank; editEndRank = emptyRemoval.endRank }
  const destinations: DocumentNoteBoundaryParagraph[] = split ? [
    { id: blockId, text: projectedText.slice(0, split.firstBreakOffset), noteReferences: projectedReferences.filter(reference => reference.endOffset <= split.firstBreakOffset) },
    { id: split.blockId, text: projectedText.slice(split.lastParagraphOffset), noteReferences: projectedReferences.filter(reference => reference.startOffset >= split.lastParagraphOffset)
      .map(reference => ({ ...reference, startOffset: reference.startOffset - split.lastParagraphOffset, endOffset: reference.endOffset - split.lastParagraphOffset })) },
  ] : [{ id: blockId, text: projectedText, noteReferences: projectedReferences }]
  if (split && (!Number.isSafeInteger(split.firstBreakOffset) || !Number.isSafeInteger(split.lastParagraphOffset)
    || split.firstBreakOffset < 0 || split.lastParagraphOffset < split.firstBreakOffset || split.lastParagraphOffset > projectedText.length)) throw new RangeError('内容控件编辑的拆分段落映射无效。')
  for (const paragraph of destinations) assertDocumentNoteReferenceOffsets(paragraph.noteReferences, paragraph.text)
  assertDocumentNoteReferenceParagraphIds(destinations)
  const completeSplit = !split || split.lastParagraphOffset === split.firstBreakOffset + 1
  const insertionAffinity = !derivedLabel && !emptyRemoval && startOffset === endOffset && completeSplit
    ? documentNoteReferenceInsertionAffinity(sourceReferences, destinations, startOffset, replacementLength) : undefined
  const replacementAffinity = emptyRemoval && completeSplit
    ? documentNoteReferenceEmptyReplacementAffinity(sourceReferences, destinations, startOffset, replacementLength, emptyRemoval) : undefined
  if (options.sourceNoteRanks && (emptyRemoval && (options.sourceNoteRanks.start !== emptyRemoval.startRank || options.sourceNoteRanks.end !== emptyRemoval.endRank)
    || insertionAffinity && (options.sourceNoteRanks.start !== insertionAffinity.sourceRank || options.sourceNoteRanks.end !== insertionAffinity.sourceRank)
    || derivedLabel && (options.sourceNoteRanks.start !== editStartRank || options.sourceNoteRanks.end !== editEndRank))) {
    throw new RangeError('内容控件编辑的源注释引用端点与真实目的内容不一致。')
  }
  if (insertionAffinity) editStartRank = editEndRank = insertionAffinity.sourceRank
  if (compareDocumentNoteBoundaries(startOffset, editStartRank, endOffset, editEndRank) > 0) throw new RangeError('内容控件编辑的源注释引用端点倒置。')
  const insertion = compareDocumentNoteBoundaries(startOffset, editStartRank, endOffset, editEndRank) === 0
  // Validate every affected range on a detached plan before changing any live anchors.
  const next: DocumentModel = { ...model, contentControls: getDocumentContentControls(model) }
  const original = new Map(next.contentControls!.map(control => [control.id, cloneControl(control)]))
  const emptyInsertionControl = !preserveControlRanges && insertion ? (model.contentControls ?? []).filter(control => {
    if (control.blockRange || control.tableRowRange || control.blockId !== blockId || control.metadata?.officeStoryLockSource) return false
    const range = documentContentControlNoteBoundaryRanks(control, sourceReferences)
    return control.startOffset === startOffset && control.endOffset === endOffset && range.start === editStartRank && range.end === editEndRank
  }).sort((left, right) => hasAncestor(model.contentControls ?? [], left, right.id) ? -1 : hasAncestor(model.contentControls ?? [], right, left.id) ? 1 : 0)[0] : undefined
  const inside = derivedLabel ? new Set<string>() : insideControlEditChain(model, blockId, startOffset, endOffset, options.activeControlId ?? emptyInsertionControl?.id,
    { text: sourceText, noteReferences: sourceReferences }, editStartRank, editEndRank)
  for (const control of next.contentControls!) {
    if (!next.contentControls!.some((candidate) => candidate.id === control.id)) continue
    if (control.tableRowRange || control.blockRange) {
      if (!documentContentControlContainsBlockRange(model, control, blockId, startOffset, endOffset, editStartRank, editEndRank)) continue
      if (!derivedLabel && isStructuralContentControlLockedAtBlockEdit(model, control, blockId, startOffset, endOffset)) throw new Error(`Content control content is locked: ${control.id}`)
      if (control.blockId === blockId) {
        const target = targetForControl(model, control), length = target ? runText(runsForTarget(model, target)).length : 0
        control.startOffset = 0; control.endOffset = split?.firstBreakOffset ?? Math.max(0, length + replacementLength - (endOffset - startOffset))
        control.startNoteReferenceCount = 0; control.endNoteReferenceCount = destinations[0]!.noteReferences.length
      }
      continue
    }
    if (control.blockId !== blockId) continue
    const ranks = documentContentControlNoteBoundaryRanks(control, sourceReferences)
    const overlaps = insertion
      ? compareDocumentNoteBoundaries(startOffset, editStartRank, control.startOffset, ranks.start) > 0
        && compareDocumentNoteBoundaries(startOffset, editStartRank, control.endOffset, ranks.end) < 0
      : compareDocumentNoteBoundaries(startOffset, editStartRank, control.endOffset, ranks.end) < 0
        && compareDocumentNoteBoundaries(control.startOffset, ranks.start, endOffset, editEndRank) < 0
    const removesWholeControl = !derivedLabel && !preserveControlRanges && !inside.has(control.id) && overlaps
      && compareDocumentNoteBoundaries(startOffset, editStartRank, control.startOffset, ranks.start) <= 0
      && compareDocumentNoteBoundaries(endOffset, editEndRank, control.endOffset, ranks.end) >= 0 && replacementLength === 0
    if (removesWholeControl) {
      if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
      removeControlTreeFromModel(next, control)
      continue
    }
    if (!derivedLabel && overlaps && isContentControlLockedAtEdit(model, control, startOffset, endOffset, insertion,
      sourceReferences, editStartRank, editEndRank)) throw new Error(`Content control content is locked: ${control.id}`)
    if (insertion && inside.has(control.id)) {
      if (isContentLocked(control) && (control.type !== 'group' || isContentControlLockedAtEdit(model, control, startOffset, endOffset, true,
        sourceReferences, editStartRank, editEndRank))) throw new Error(`Content control content is locked: ${control.id}`)
      assertAncestorContentEditable(model, control)
      control.endOffset += replacementLength
      continue
    }
    if (!overlaps) {
      if (!insertion && compareDocumentNoteBoundaries(endOffset, editEndRank, control.startOffset, ranks.start) <= 0) {
        const delta = replacementLength - (endOffset - startOffset)
        control.startOffset += delta
        control.endOffset += delta
      } else if (insertion && compareDocumentNoteBoundaries(startOffset, editStartRank, control.startOffset, ranks.start) <= 0) {
        control.startOffset += replacementLength
        control.endOffset += replacementLength
      }
      continue
    }
    if (!derivedLabel && isContentControlLockedAtEdit(model, control, startOffset, endOffset, insertion,
      sourceReferences, editStartRank, editEndRank)) throw new Error(`Content control content is locked: ${control.id}`)
    if (insertion) {
      control.endOffset += replacementLength
      continue
    }
    const delta = replacementLength - (endOffset - startOffset)
    if (compareDocumentNoteBoundaries(startOffset, editStartRank, control.startOffset, ranks.start) <= 0
      && compareDocumentNoteBoundaries(endOffset, editEndRank, control.endOffset, ranks.end) >= 0) {
      control.startOffset = startOffset
      control.endOffset = startOffset + replacementLength
    } else {
      const oldStart = control.startOffset
      const oldEnd = control.endOffset
      control.startOffset = compareDocumentNoteBoundaries(oldStart, ranks.start, endOffset, editEndRank) >= 0 ? oldStart + delta
        : compareDocumentNoteBoundaries(oldStart, ranks.start, startOffset, editStartRank) >= 0 ? startOffset : oldStart
      control.endOffset = compareDocumentNoteBoundaries(oldEnd, ranks.end, endOffset, editEndRank) >= 0 ? oldEnd + delta
        : compareDocumentNoteBoundaries(oldEnd, ranks.end, startOffset, editStartRank) > 0 ? startOffset + replacementLength : oldEnd
      if (control.endOffset < control.startOffset) control.endOffset = control.startOffset
    }
  }
  if (split) {
    const unwrapping = new Set<string>()
    if (!preserveControlRanges) {
      for (const control of next.contentControls!) {
        if (control.tableRowRange || control.blockRange || control.blockId !== blockId || control.endOffset <= split.firstBreakOffset
          || control.startOffset >= split.lastParagraphOffset || !control.temporary || isControlLocked(control)
          || control.metadata?.officeStoryLockSource) continue
        const previous = model.contentControls!.find(candidate => candidate.id === control.id)!
        if (isContentLocked(previous)) throw new Error(`Content control content is locked: ${control.id}`)
        assertAncestorContentEditable(model, previous)
        assertControlAcceptsParagraphBreak(control)
        unwrapping.add(control.id)
      }
    }
    next.contentControls = unwrapDocumentContentControlShells(next.contentControls!, unwrapping)
    const controls = next.contentControls
    const movesToSuffix = (control: DocumentContentControl, visited = new Set<string>()): boolean => {
      if (visited.has(control.id)) throw new RangeError(`Content control parent relationships must not contain a cycle: ${control.id}`)
      visited.add(control.id)
      if (control.endOffset <= split.firstBreakOffset) {
        // Empty children at the split boundary follow their enclosing wrapper.
        const parent = control.startOffset === control.endOffset && control.parentId
          ? controls.find((candidate) => candidate.id === control.parentId && candidate.startOffset < candidate.endOffset)
          : undefined
        return parent ? movesToSuffix(parent, visited) : false
      }
      if (control.startOffset < split.lastParagraphOffset) {
        throw new RangeError(`A paragraph break cannot split a content control range: ${control.id}`)
      }
      return true
    }
    const moving = controls.filter((control) => !control.tableRowRange && !control.blockRange && control.blockId === blockId && movesToSuffix(control))
    for (const control of moving) {
      control.blockId = split.blockId
      control.startOffset -= split.lastParagraphOffset
      control.endOffset -= split.lastParagraphOffset
    }
  }
  for (const control of next.contentControls!) {
    const previous = original.get(control.id)
    if (!previous || previous.blockId !== blockId || control.blockRange || control.tableRowRange) continue
    const fallback = destinations.find(paragraph => paragraph.id === control.blockId)
    if (!fallback) throw new RangeError('内容控件编辑后的段落宿主不存在：' + control.id)
    let start = remapDocumentNoteBoundaryParagraph(sourceReferences, previous.startOffset, previous.startNoteReferenceCount, destinations,
      { paragraph: fallback, offset: control.startOffset })
    let end = remapDocumentNoteBoundaryParagraph(sourceReferences, previous.endOffset, previous.endNoteReferenceCount, destinations,
      { paragraph: fallback, offset: control.endOffset })
    const ranks = documentContentControlNoteBoundaryRanks(previous, sourceReferences)
    const point = compareDocumentNoteBoundaries(previous.startOffset, ranks.start, previous.endOffset, ranks.end) === 0
    const pointBefore = insertionAffinity && documentContentControlPointBeforeInsertion(previous, original, sourceReferences,
      insertionAffinity.sourceOffset, insertionAffinity.sourceRank, inside)
    start = documentNoteEmptyReplacementBoundary(replacementAffinity, previous.startOffset, ranks.start, 'start')
      ?? (pointBefore ? insertionAffinity?.before : documentNoteInsertionBoundary(insertionAffinity, previous.startOffset, ranks.start, 'start', point, inside.has(control.id))) ?? start
    end = documentNoteEmptyReplacementBoundary(replacementAffinity, previous.endOffset, ranks.end, 'end')
      ?? (pointBefore ? insertionAffinity?.before : documentNoteInsertionBoundary(insertionAffinity, previous.endOffset, ranks.end, 'end', point, inside.has(control.id))) ?? end
    if (start.paragraph.id !== end.paragraph.id) throw new RangeError('内容控件的引用端点不能跨拆分后的段落：' + control.id)
    control.blockId = start.paragraph.id
    assignDocumentContentControlNoteBoundaries(control, start, end)
    if (!isDocumentTextBoundary(start.paragraph.text, control.startOffset) || !isDocumentTextBoundary(start.paragraph.text, control.endOffset)) throw new RangeError('内容控件编辑后的 UTF-16 端点无效。')
    documentContentControlNoteBoundaryRanks(control, start.paragraph.noteReferences)
  }
  const geometry = new Map(destinations.map(paragraph => [paragraph.id, paragraph.noteReferences]))
  for (const control of next.contentControls!) if (geometry.has(control.blockId) && !control.blockRange && !control.tableRowRange) {
    const parent = control.parentId ? next.contentControls!.find(candidate => candidate.id === control.parentId) : undefined
    if (parent && !parent.blockRange && !parent.tableRowRange && !documentContentControlNoteRangeContains(parent, control, geometry.get(control.blockId)!)) throw new RangeError('内容控件编辑后的父子引用范围不一致。')
    assertNoPartialOverlap(next, next.contentControls!, control, geometry)
  }
  model.contentControls = next.contentControls
}

export function assertContentControlBlockRemovable(model: DocumentModel, blockId: string): void {
  for (const control of model.contentControls ?? []) if (control.tableRowRange && control.blockId === blockId
    && findBlock(model, control.tableRowRange.tableId)) reanchorDocumentContentControlTableRows(model, control)
  for (const control of model.contentControls ?? []) if (control.blockRange && control.blockId === blockId
    && findBlock(model, control.blockRange.blockIds[0]!)) reanchorDocumentBlockContentControl(model, control)
  const controls = model.contentControls?.filter((candidate) => !candidate.blockRange && !candidate.tableRowRange && candidate.blockId === blockId) ?? []
  for (const control of controls) if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
  for (const control of controls) removeControlTreeFromModel(model, control)
}

/** Validate mapped ranges and the wrapper tree before publishing a structural edit. */
export function validateDocumentContentControlRanges(model: DocumentModel): void {
  const paragraphs = new Map(documentTextFieldParagraphs(model).map(paragraph => [paragraph.block.id, paragraph]))
  for (const control of model.contentControls ?? []) {
    const host = documentBlockHostForId(model, control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId), scoped = host ? documentModelForBlockHost(model, host) : model
    validateControl(scoped, cloneControl(control))
    const paragraph = host ? documentTextFieldParagraphs(scoped).find(candidate => candidate.block.id === control.blockId) : paragraphs.get(control.blockId)
    if (paragraph && (!isDocumentTextBoundary(paragraph.text, control.startOffset) || !isDocumentTextBoundary(paragraph.text, control.endOffset))) {
      throw new RangeError(`Content control UTF-16 range splits a character: ${control.id}`)
    }
    assertNoPartialOverlap(scoped, scoped.contentControls ?? [], control)
  }
}

/** Map paragraph-local SDT wrappers after a body edit; field refreshes retain the wrappers. */
export function transformContentControlsForBodyRangeEdit(
  model: DocumentModel,
  sourceParagraphs: readonly DocumentTextFieldParagraph[],
  startOffset: number,
  endOffset: number,
  replacementLength: number,
  preserveControlRanges = true,
  options: ContentControlRangeEditOptions = {},
): void {
  if (!model.contentControls?.length) return
  if (options.noteReferenceLabelRunId !== undefined || options.projectedNoteReferences !== undefined) throw new TypeError('派生标签与显式引用投影必须使用段落编辑接口。')
  if (options.sourceModel) assertBodyEditSourceModel(model, options.sourceModel, sourceParagraphs)
  const sourceContext = options.sourceModel ?? model
  const source = new Map(sourceParagraphs.map(paragraph => [paragraph.block.id, paragraph]))
  const destination = documentTextFieldParagraphs(model)
  const targets = destination.map(paragraph => ({ ...paragraph, id: paragraph.block.id }))
  assertDocumentNoteReferenceParagraphIds(sourceParagraphs)
  assertDocumentNoteReferenceParagraphIds(targets)
  const next: DocumentModel = { ...model, contentControls: getDocumentContentControls(model) }
  const original = new Map((sourceContext.contentControls ?? []).map(control => [control.id, cloneControl(control)]))
  const delta = replacementLength - (endOffset - startOffset)
  const startParagraph = sourceParagraphs.find(paragraph => startOffset >= paragraph.absoluteStart && startOffset <= paragraph.absoluteEnd)
  const endParagraph = sourceParagraphs.find(paragraph => endOffset >= paragraph.absoluteStart && endOffset <= paragraph.absoluteEnd)
  if (!startParagraph || !endParagraph || !Number.isSafeInteger(replacementLength) || replacementLength < 0 || endOffset < startOffset
    || !isDocumentTextBoundary(startParagraph.text, startOffset - startParagraph.absoluteStart)
    || !isDocumentTextBoundary(endParagraph.text, endOffset - endParagraph.absoluteStart)) throw new RangeError('内容控件正文编辑的源端点无效。')
  let editStartRank = assertDocumentNoteReferenceBoundaryRank(startParagraph.noteReferences, startOffset - startParagraph.absoluteStart, options.sourceNoteRanks?.start)
  let editEndRank = assertDocumentNoteReferenceBoundaryRank(endParagraph.noteReferences, endOffset - endParagraph.absoluteStart, options.sourceNoteRanks?.end)
  const localStart = startOffset - startParagraph.absoluteStart
  const emptyRemoval = startOffset === endOffset ? documentNoteReferenceEmptyRemovalRange(startParagraph.noteReferences,
    new Set(targets.flatMap(paragraph => paragraph.noteReferences.map(reference => reference.runId))), localStart) : undefined
  const editedOutputs = startOffset === endOffset ? targets.filter(paragraph => paragraph.absoluteStart >= startParagraph.absoluteStart
    && paragraph.absoluteEnd <= startParagraph.absoluteEnd + replacementLength) : []
  if (startOffset === endOffset && (!editedOutputs.length || editedOutputs[0]!.absoluteStart !== startParagraph.absoluteStart
    || editedOutputs.at(-1)!.absoluteEnd !== startParagraph.absoluteEnd + replacementLength)) throw new RangeError('正文插入的完整目的段落列表与源范围不一致。')
  const insertionAffinity = startOffset === endOffset && !emptyRemoval
    ? documentNoteReferenceInsertionAffinity(startParagraph.noteReferences, editedOutputs, localStart, replacementLength) : undefined
  const replacementAffinity = emptyRemoval
    ? documentNoteReferenceEmptyReplacementAffinity(startParagraph.noteReferences, editedOutputs, localStart, replacementLength, emptyRemoval) : undefined
  if (options.sourceNoteRanks && (emptyRemoval && (editStartRank !== emptyRemoval.startRank || editEndRank !== emptyRemoval.endRank)
    || insertionAffinity && (editStartRank !== insertionAffinity.sourceRank || editEndRank !== insertionAffinity.sourceRank))) {
    throw new RangeError('正文内容控件编辑的源注释引用端点与真实目的内容不一致。')
  }
  if (emptyRemoval) { editStartRank = emptyRemoval.startRank; editEndRank = emptyRemoval.endRank }
  else if (insertionAffinity) editStartRank = editEndRank = insertionAffinity.sourceRank
  if (compareDocumentNoteBoundaries(startOffset, editStartRank, endOffset, editEndRank) > 0) throw new RangeError('正文内容控件编辑的源注释引用端点倒置。')
  const insertion = compareDocumentNoteBoundaries(startOffset, editStartRank, endOffset, editEndRank) === 0
  const emptyInputControl = !preserveControlRanges && insertion ? (model.contentControls ?? [])
    .filter(control => {
      const paragraph = source.get(control.blockId)
      if (!paragraph || control.blockRange || control.tableRowRange || control.metadata?.officeStoryLockSource
        || paragraph.absoluteStart + control.startOffset !== startOffset || control.endOffset !== control.startOffset) return false
      const range = documentContentControlNoteBoundaryRanks(control, paragraph.noteReferences)
      return range.start === editStartRank && range.end === editEndRank
    })
    .sort((left, right) => hasAncestor(model.contentControls ?? [], left, right.id) ? -1 : hasAncestor(model.contentControls ?? [], right, left.id) ? 1 : 0)[0] : undefined
  const inputControlId = options.activeControlId ?? emptyInputControl?.id
  const inside = insideBodyControlEditChain(model, options.sourceModel, sourceParagraphs, startParagraph, endParagraph,
    localStart, editStartRank, endOffset - endParagraph.absoluteStart, editEndRank, inputControlId)
  const unwrapping = new Set<string>()
  for (const control of [...next.contentControls!]) {
    if (!next.contentControls!.some(candidate => candidate.id === control.id)) continue
    // The complete source block tree is reconciled by the outer structural planner after mutation.
    if (control.blockRange) continue
    if (control.tableRowRange) {
      reanchorDocumentContentControlTableRows(model, control)
      const target = targetForControl(model, control)
      if (!target) throw new RangeError(`Row content control has no surviving input target: ${control.id}`)
      control.startOffset = 0; control.endOffset = runText(runsForTarget(model, target)).length
      continue
    }
    const previous = original.get(control.id), paragraph = previous && source.get(previous.blockId)
    if (!previous || !paragraph) continue
    const ranks = documentContentControlNoteBoundaryRanks(previous, paragraph.noteReferences)
    const from = paragraph.absoluteStart + previous.startOffset, to = paragraph.absoluteStart + previous.endOffset
    const overlaps = insertion ? compareDocumentNoteBoundaries(startOffset, editStartRank, from, ranks.start) > 0
      && compareDocumentNoteBoundaries(startOffset, editStartRank, to, ranks.end) < 0
      : compareDocumentNoteBoundaries(startOffset, editStartRank, to, ranks.end) < 0 && compareDocumentNoteBoundaries(from, ranks.start, endOffset, editEndRank) < 0
    if (!preserveControlRanges && !inside.has(control.id) && overlaps
      && compareDocumentNoteBoundaries(startOffset, editStartRank, from, ranks.start) <= 0
      && compareDocumentNoteBoundaries(endOffset, editEndRank, to, ranks.end) >= 0 && replacementLength === 0) {
      if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
      removeControlTreeFromModel(next, control)
      continue
    }
    if (isContentControlLockedAtEdit(sourceContext, previous, startOffset - paragraph.absoluteStart, endOffset - paragraph.absoluteStart, insertion,
      paragraph.noteReferences, paragraph === startParagraph ? editStartRank : undefined, paragraph === endParagraph ? editEndRank : undefined)) {
      throw new Error(`Content control content is locked: ${control.id}`)
    }
    let start = from, end = to
    if (insertion && inside.has(control.id)) {
      if (isContentLocked(previous) && (previous.type !== 'group' || isContentControlLockedAtEdit(sourceContext, previous, startOffset - paragraph.absoluteStart,
        endOffset - paragraph.absoluteStart, true, paragraph.noteReferences, paragraph === startParagraph ? editStartRank : undefined,
        paragraph === endParagraph ? editEndRank : undefined))) throw new Error(`Content control content is locked: ${control.id}`)
      assertAncestorContentEditable(sourceContext, previous)
      end += replacementLength
    } else if (insertion) {
      if (compareDocumentNoteBoundaries(startOffset, editStartRank, from, ranks.start) <= 0) { start += replacementLength; end += replacementLength }
      else if (compareDocumentNoteBoundaries(startOffset, editStartRank, to, ranks.end) < 0) end += replacementLength
    } else if (compareDocumentNoteBoundaries(endOffset, editEndRank, from, ranks.start) <= 0) { start += delta; end += delta }
    else if (overlaps) {
      start = compareDocumentNoteBoundaries(from, ranks.start, startOffset, editStartRank) >= 0 ? startOffset : from
      end = compareDocumentNoteBoundaries(to, ranks.end, endOffset, editEndRank) >= 0 ? to + delta : startOffset + replacementLength
      if (end < start) end = start
    }
    const first = destination.find(candidate => start >= candidate.absoluteStart && start <= candidate.absoluteEnd)
    const last = destination.find(candidate => end >= candidate.absoluteStart && end <= candidate.absoluteEnd)
    if (!first || !last) throw new RangeError(`正文编辑无法把内容控件 ${control.id} 映射到同一段落。`)
    let mappedStart = remapDocumentNoteBoundaryParagraph(paragraph.noteReferences, previous.startOffset, previous.startNoteReferenceCount, targets,
      { paragraph: targets.find(candidate => candidate.id === first.block.id)!, offset: start - first.absoluteStart })
    let mappedEnd = remapDocumentNoteBoundaryParagraph(paragraph.noteReferences, previous.endOffset, previous.endNoteReferenceCount, targets,
      { paragraph: targets.find(candidate => candidate.id === last.block.id)!, offset: end - last.absoluteStart })
    if (paragraph.block.id === startParagraph.block.id) {
      const point = compareDocumentNoteBoundaries(previous.startOffset, ranks.start, previous.endOffset, ranks.end) === 0
      const pointBefore = insertionAffinity && documentContentControlPointBeforeInsertion(previous, original, paragraph.noteReferences,
        insertionAffinity.sourceOffset, insertionAffinity.sourceRank, inside)
      mappedStart = documentNoteEmptyReplacementBoundary(replacementAffinity, previous.startOffset, ranks.start, 'start')
        ?? (pointBefore ? insertionAffinity?.before : documentNoteInsertionBoundary(insertionAffinity, previous.startOffset, ranks.start, 'start', point, inside.has(control.id))) ?? mappedStart
      mappedEnd = documentNoteEmptyReplacementBoundary(replacementAffinity, previous.endOffset, ranks.end, 'end')
        ?? (pointBefore ? insertionAffinity?.before : documentNoteInsertionBoundary(insertionAffinity, previous.endOffset, ranks.end, 'end', point, inside.has(control.id))) ?? mappedEnd
    }
    if (mappedStart.paragraph.id !== mappedEnd.paragraph.id) {
      if (preserveControlRanges || !control.temporary || isControlLocked(control) || control.metadata?.officeStoryLockSource) {
        throw new RangeError(`正文编辑无法把内容控件 ${control.id} 映射到同一段落。`)
      }
      if (isContentLocked(previous)) throw new Error(`Content control content is locked: ${control.id}`)
      assertAncestorContentEditable(sourceContext, previous)
      assertControlAcceptsParagraphBreak(control)
      unwrapping.add(control.id)
      continue
    }
    control.blockId = mappedStart.paragraph.id
    assignDocumentContentControlNoteBoundaries(control, mappedStart, mappedEnd)
    documentContentControlNoteBoundaryRanks(control, mappedStart.paragraph.noteReferences)
  }
  next.contentControls = unwrapDocumentContentControlShells(next.contentControls!, unwrapping)
  const inline = next.contentControls.filter(control => !control.blockRange && !control.tableRowRange)
  for (const control of inline) {
    if (!destination.some(paragraph => paragraph.block.id === control.blockId)) continue
    const parent = control.parentId ? inline.find(candidate => candidate.id === control.parentId) : undefined
    if (parent && !documentContentControlNoteRangeContains(parent, control, destination.find(paragraph => paragraph.block.id === control.blockId)!.noteReferences)) throw new RangeError('正文编辑后的内容控件父子引用范围不一致。')
    assertNoPartialOverlap(next, inline, control)
  }
  model.contentControls = next.contentControls
}

function replaceControlText(model: DocumentModel, control: DocumentContentControl, value: string, marks: RunPatch = {}): void {
  if (control.blockRange) { replaceDocumentBlockContentControlText(model, control, value, marks); return }
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control paragraph not found: ${control.blockId}`)
  const oldLength = control.endOffset - control.startOffset
  const editStart = control.startOffset, editEnd = control.endOffset
  const block = findBlock(model, control.blockId)!, sourceNoteReferences = documentNoteReferenceOffsets(block.runs)
  const sourceRange = documentContentControlNoteBoundaryRanks(control, sourceNoteReferences)
  const inside = insideControlEditChain(model, block.id, editStart, editEnd, control.id,
    { text: runText(block.runs), noteReferences: sourceNoteReferences }, sourceRange.start, sourceRange.end)
  const original = new Map((model.contentControls ?? []).filter(candidate => candidate.blockId === control.blockId).map(candidate => [candidate.id, cloneControl(candidate)]))
  for (const candidate of model.contentControls ?? []) {
    if (candidate.id === control.id || candidate.blockId !== control.blockId) continue
    if (candidate.blockRange || candidate.tableRowRange) continue
    const overlaps = documentContentControlNoteRangesOverlap(candidate, control, sourceNoteReferences)
    if (hasAncestor(model.contentControls ?? [], candidate, control.id) || overlaps && documentContentControlNoteRangeContains(control, candidate, sourceNoteReferences)) {
      throw new RangeError(`Cannot replace content control text while a nested control is present: ${candidate.id}`)
    }
    if (overlaps && !documentContentControlNoteRangeContains(candidate, control, sourceNoteReferences)) {
      throw new RangeError(`Cannot replace content control text across a partially overlapping control: ${candidate.id}`)
    }
  }
  const runs = runsForTarget(model, target)
  const { before: prefix, after: suffix } = partitionDocumentContentControlNoteRuns(runs, control)
  const inherited = documentRunStyleAt(runs, control.startOffset, control.endOffset) ?? {}
  const inserted: DocumentRun[] = value ? [createRun(value, { ...inherited, ...marks })] : []
  const nextRuns = [...prefix, ...inserted, ...suffix]
  const nextReferences = documentNoteReferenceOffsets(nextRuns)
  const paragraph = documentTextFieldParagraphs(model).find(paragraph => paragraph.block.id === block.id)!
  const absoluteStart = paragraph.absoluteStart + editStart, absoluteEnd = paragraph.absoluteStart + editEnd
  const insertion = compareDocumentNoteBoundaries(editStart, sourceRange.start, editEnd, sourceRange.end) === 0
  for (const field of documentTextFieldRanges(model)) {
    const touches = insertion ? compareDocumentNoteBoundaries(absoluteStart, sourceRange.start, field.absoluteStart, field.startNoteReferenceCount) > 0
      && compareDocumentNoteBoundaries(absoluteStart, sourceRange.start, field.absoluteEnd, field.endNoteReferenceCount) < 0
      : documentTextFieldRangeIntersects(field, absoluteStart, absoluteEnd, sourceRange.start, sourceRange.end)
    if (documentFieldLocked(field.field) && touches) throw new RangeError('内容控件设值不能改变锁定字段：' + field.field.id)
  }
  transformDocumentTextFieldsForBlockEdit(block, control.startOffset, control.endOffset, value, undefined, model, nextReferences, sourceRange)
  // The editor's shared range splitter keeps note/field entities at the boundaries exactly once.
  block.runs = nextRuns.length ? nextRuns : [createRun('', inherited)]
  const code = block.metadata.codeBlock
  if (code && typeof code === 'object' && !Array.isArray(code) && typeof Reflect.get(code, 'source') === 'string') {
    block.metadata = { ...block.metadata, codeBlock: { ...code, source: runText(block.runs) } }
  }
  transformDocumentCommentAnchorsForTextEdit(model, block.id, editStart, editEnd, value, Date.now(), sourceNoteReferences)
  for (const ancestor of documentBlockAncestry(model, block.id) ?? []) delete ancestor.metadata.html
  const delta = value.length - oldLength
  const destination: DocumentNoteBoundaryParagraph = { id: block.id, text: runText(block.runs), noteReferences: nextReferences }
  const emptyRemoval = editStart === editEnd ? documentNoteReferenceEmptyRemovalRange(sourceNoteReferences,
    new Set(nextReferences.map(reference => reference.runId)), editStart) : undefined
  const insertionAffinity = editStart === editEnd && !emptyRemoval
    ? documentNoteReferenceInsertionAffinity(sourceNoteReferences, [destination], editStart, value.length) : undefined
  const replacementAffinity = emptyRemoval
    ? documentNoteReferenceEmptyReplacementAffinity(sourceNoteReferences, [destination], editStart, value.length, emptyRemoval) : undefined
  for (const candidate of model.contentControls ?? []) {
    if (candidate.id === control.id || candidate.blockId !== control.blockId || candidate.blockRange || candidate.tableRowRange) continue
    const previous = original.get(candidate.id)!, rank = documentContentControlNoteBoundaryRanks(previous, sourceNoteReferences)
    if (inside.has(candidate.id)) {
      candidate.endOffset += delta
    } else if (compareDocumentNoteBoundaries(previous.startOffset, rank.start, editEnd, sourceRange.end) >= 0) {
      candidate.startOffset += delta
      candidate.endOffset += delta
    }
    let from = remapDocumentNoteBoundaryParagraph(sourceNoteReferences, previous.startOffset, previous.startNoteReferenceCount, [destination], { paragraph: destination, offset: candidate.startOffset })
    let to = remapDocumentNoteBoundaryParagraph(sourceNoteReferences, previous.endOffset, previous.endNoteReferenceCount, [destination], { paragraph: destination, offset: candidate.endOffset })
    const point = compareDocumentNoteBoundaries(previous.startOffset, rank.start, previous.endOffset, rank.end) === 0
    const pointBefore = insertionAffinity && documentContentControlPointBeforeInsertion(previous, original, sourceNoteReferences,
      insertionAffinity.sourceOffset, insertionAffinity.sourceRank, inside)
    from = documentNoteEmptyReplacementBoundary(replacementAffinity, previous.startOffset, rank.start, 'start')
      ?? (pointBefore ? insertionAffinity?.before : documentNoteInsertionBoundary(insertionAffinity, previous.startOffset, rank.start, 'start', point, inside.has(candidate.id))) ?? from
    to = documentNoteEmptyReplacementBoundary(replacementAffinity, previous.endOffset, rank.end, 'end')
      ?? (pointBefore ? insertionAffinity?.before : documentNoteInsertionBoundary(insertionAffinity, previous.endOffset, rank.end, 'end', point, inside.has(candidate.id))) ?? to
    assignDocumentContentControlNoteBoundaries(candidate, from, to)
    documentContentControlNoteBoundaryRanks(candidate, nextReferences)
  }
  const rank = documentNoteReferenceOffsets(prefix).length
  assignDocumentContentControlNoteBoundaries(control, { offset: editStart, rank }, { offset: editStart + value.length, rank })
  documentContentControlNoteBoundaryRanks(control, nextReferences)
  for (const candidate of model.contentControls ?? []) if (candidate.blockId === block.id && (candidate.blockRange || candidate.tableRowRange)) {
    candidate.startOffset = 0; candidate.endOffset = destination.text.length
    candidate.startNoteReferenceCount = 0; candidate.endNoteReferenceCount = nextReferences.length
  }
  for (const candidate of model.contentControls ?? []) if (candidate.blockId === block.id) assertNoPartialOverlap(model, model.contentControls ?? [], candidate)
}

function targetForControl(model: DocumentModel, control: DocumentContentControl): TextTarget | undefined {
  return textTargets(model).find((target) => (target.cellBlockId ?? target.blockId) === control.blockId)
}

function refreshContentControlParagraphText(model: DocumentModel, block: DocumentBlock): void {
  const code = block.metadata.codeBlock
  if (code && typeof code === 'object' && !Array.isArray(code)) block.metadata.codeBlock = { ...code, source: runText(block.runs) }
  for (const ancestor of documentBlockAncestry(model, block.id) ?? []) delete ancestor.metadata.html
}

function flattenBodySegments(model: DocumentModel): Array<{ blockId: string; cellBlockId?: string; startOffset: number; endOffset: number }> {
  const segments: Array<{ blockId: string; cellBlockId?: string; startOffset: number; endOffset: number }> = []
  let offset = 0
  for (const target of textTargets(model)) {
    const text = runText(runsForTarget(model, target))
    segments.push({ blockId: target.blockId, ...(target.cellBlockId ? { cellBlockId: target.cellBlockId } : {}), startOffset: offset, endOffset: offset + text.length })
    offset += text.length + 1
  }
  return segments
}

function segmentAt<T extends { startOffset: number; endOffset: number }>(segments: readonly T[], offset: number): T | undefined {
  return segments.find((segment) => offset >= segment.startOffset && offset <= segment.endOffset)
}

function findBlock(model: DocumentModel, id: string): DocumentBlock | undefined {
  const visit = (blocks: DocumentBlock[]): DocumentBlock | undefined => {
    for (const block of blocks) {
      if (block.id === id) return block
      const table = tableForBlock(block)
      for (const row of table?.rows ?? []) {
        for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) {
          const nested = visit(content)
          if (nested) return nested
        }
      }
    }
    return undefined
  }
  return visit(model.blocks)
}

function blockContainerForTarget(model: DocumentModel, target: TextTarget): DocumentBlock[] | undefined {
  if (!target.cellBlockId) return model.blocks
  const visit = (blocks: DocumentBlock[]): DocumentBlock[] | undefined => {
    for (const block of blocks) {
      const table = tableForBlock(block)
      for (const row of table?.rows ?? []) {
        for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) {
          if (content.some(candidate => candidate.id === target.cellBlockId)) return content
          const nested = visit(content)
          if (nested) return nested
        }
      }
    }
    return undefined
  }
  return visit(model.blocks)
}

function blockContainerContaining(model: DocumentModel, blockId: string): DocumentBlock[] | undefined {
  const visit = (blocks: DocumentBlock[]): DocumentBlock[] | undefined => {
    if (blocks.some(block => block.id === blockId)) return blocks
    for (const block of blocks) {
      for (const row of tableForBlock(block)?.rows ?? []) {
        for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) {
          const nested = visit(content)
          if (nested) return nested
        }
      }
    }
    return undefined
  }
  return visit(model.blocks)
}

function validateControl(model: DocumentModel, control: DocumentContentControl): void {
  const block = findBlock(model, control.blockId)
  if (!block || ['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) throw new Error(`Content control target is not a text paragraph: ${control.blockId}`)
  if (!Number.isSafeInteger(control.startOffset) || !Number.isSafeInteger(control.endOffset) || control.startOffset < 0 || control.endOffset < control.startOffset) throw new RangeError('Content control range is invalid')
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control target is not addressable: ${control.blockId}`)
  const text = runText(runsForTarget(model, target)), length = text.length
  if (control.endOffset > length) throw new RangeError(`Content control range exceeds paragraph length: ${length}`)
  if (!isDocumentTextBoundary(text, control.startOffset) || !isDocumentTextBoundary(text, control.endOffset)) throw new RangeError(`Content control UTF-16 range splits a character: ${control.id}`)
  documentContentControlNoteBoundaryRanks(control, documentNoteReferenceOffsets(runsForTarget(model, target)))
  if (!['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox', 'picture', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj'].includes(control.type)) throw new TypeError('Unsupported content control type')
  if (control.lock === undefined) control.lock = 'unlocked'
  if (control.type === 'date' && control.dateValue !== undefined && control.dateValue !== '' && !validContentControlIsoDate(control.dateValue)) throw new RangeError('Date content control value must be YYYY-MM-DD')
  if ((control.type === 'dropDownList' || control.type === 'comboBox') && (!control.items?.length || control.items.some((item) => !item.value))) throw new RangeError('List content controls require non-empty values')
  if (control.type === 'dropDownList' && !control.showingPlaceholder && control.value !== undefined && !control.items?.some((item) => item.value === control.value)) throw new RangeError(`Dropdown value is not present in its item list: ${control.value}`)
  if (control.type === 'checkbox' && control.checked === undefined) control.checked = false
  if (control.placement === 'row') {
    if (!control.tableRowRange) throw new RangeError(`Row content control requires a persisted table row range: ${control.id}`)
    const rows = resolveDocumentContentControlTableRows(model, control.tableRowRange)
    if (!documentContentControlTableRowBlockIds(rows).has(control.blockId)) throw new RangeError(`Row content control input fallback is outside its table rows: ${control.id}`)
  } else if (control.tableRowRange) throw new RangeError(`Only row content controls can carry a table row range: ${control.id}`)
  if (control.blockRange) {
    if (control.tableRowRange || control.placement !== 'block') throw new RangeError(`Complete block content controls require block placement: ${control.id}`)
    if (!documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange)).has(control.blockId)) {
      throw new RangeError(`Block content control input fallback is outside its owned blocks: ${control.id}`)
    }
  }
  if (control.parentId) {
    if (control.parentId === control.id || hasAncestor(model.contentControls ?? [], control, control.id)) throw new RangeError('Content control parent relationships must not contain a cycle')
    const parent = model.contentControls?.find((candidate) => candidate.id === control.parentId)
    if (!parent || !['group', 'repeatingSection', 'repeatingSectionItem'].includes(parent.type)) throw new Error(`Invalid content control parent: ${control.parentId}`)
    if (!documentContentControlContains(model, parent, control)) throw new RangeError(`Child content control must be within its parent: ${control.parentId}`)
  }
}

function assertNoPartialOverlap(model: DocumentModel, controls: DocumentContentControl[], next: DocumentContentControl,
  geometry?: ReadonlyMap<string, readonly DocumentNoteReferenceOffset[]>): void {
  const references = (blockId: string) => geometry?.get(blockId) ?? documentNoteReferenceOffsets(findBlock(model, blockId)?.runs ?? [])
  const contains = (parent: DocumentContentControl, child: DocumentContentControl) => parent.blockRange || child.blockRange || parent.tableRowRange || child.tableRowRange
    ? documentContentControlContains(model, parent, child) : documentContentControlNoteRangeContains(parent, child, references(parent.blockId))
  for (const current of controls) {
    if (current.id === next.id || current.metadata?.officeStoryLockSource || next.metadata?.officeStoryLockSource) continue
    if (!findBlock(model, current.blockRange?.blockIds[0] ?? current.tableRowRange?.tableId ?? current.blockId) && !geometry?.has(current.blockId)) continue
    let overlaps: boolean
    if (current.blockRange || next.blockRange) {
      const ids = (control: DocumentContentControl) => control.blockRange
        ? documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange))
        : control.tableRowRange ? documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(model, control.tableRowRange)) : new Set([control.blockId])
      const left = ids(current), right = ids(next)
      overlaps = [...left].some(id => right.has(id))
    } else if (current.tableRowRange && next.tableRowRange && current.tableRowRange.tableId === next.tableRowRange.tableId) {
      const left = resolveDocumentContentControlTableRows(model, current.tableRowRange), right = resolveDocumentContentControlTableRows(model, next.tableRowRange)
      overlaps = left.startRow <= right.endRow && right.startRow <= left.endRow
    } else if (current.tableRowRange || next.tableRowRange) {
      overlaps = documentContentControlContains(model, current, next) || documentContentControlContains(model, next, current)
    } else overlaps = current.blockId === next.blockId && documentContentControlNoteRangesOverlap(current, next, references(current.blockId))
    if (!overlaps) continue
    const nextContains = contains(next, current)
    const currentContains = contains(current, next)
    const parentTypes = ['group', 'repeatingSection', 'repeatingSectionItem']
    const validParent = (currentContains && hasAncestor(controls, next, current.id) && parentTypes.includes(current.type))
      || (nextContains && hasAncestor(controls, current, next.id) && parentTypes.includes(next.type))
    if (!validParent) throw new RangeError(`Content control ranges overlap without a valid parent relationship: ${current.id} and ${next.id}`)
  }
}

function requireControl(model: DocumentModel, id: string): DocumentContentControl {
  const control = model.contentControls?.find((candidate) => candidate.id === id)
  if (!control) throw new Error(`Content control not found: ${id}`)
  return control
}

function isContentLocked(control: DocumentContentControl): boolean { return lockedContent.has(control.lock) }
function isControlLocked(control: DocumentContentControl): boolean { return lockedControl.has(control.lock) }
function assertAncestorContentEditable(model: DocumentModel, control: DocumentContentControl): void {
  for (const candidate of model.contentControls ?? []) {
    if (candidate.id === control.id) continue
    if (candidate.tableRowRange || candidate.blockRange) {
      if (findBlock(model, candidate.blockRange?.blockIds[0] ?? candidate.tableRowRange!.tableId) && documentContentControlContains(model, candidate, control)
        && isContentLocked(candidate) && !(candidate.type === 'group' && (model.contentControls ?? []).some(child => child.parentId === candidate.id
          && (child.id === control.id || documentContentControlContains(model, child, control))))) {
        throw new Error(`Content control content is locked: ${candidate.id}`)
      }
      continue
    }
    if (control.blockRange || control.tableRowRange) continue
    if (candidate.blockId !== control.blockId) continue
    const references = documentNoteReferenceOffsets(findBlock(model, control.blockId)?.runs ?? [])
    const range = documentContentControlNoteBoundaryRanks(control, references)
    if (isContentControlLockedAtEdit(model, candidate, control.startOffset, control.endOffset,
      compareDocumentNoteBoundaries(control.startOffset, range.start, control.endOffset, range.end) === 0, references, range.start, range.end)) {
      throw new Error(`Content control content is locked: ${candidate.id}`)
    }
  }
}
function cloneControl(control: DocumentContentControl): DocumentContentControl {
  return cloneModel(control)
}

function hasAncestor(controls: readonly DocumentContentControl[], control: DocumentContentControl, ancestorId: string): boolean {
  let parentId = control.parentId
  const visited = new Set<string>()
  while (parentId && !visited.has(parentId)) {
    if (parentId === ancestorId) return true
    visited.add(parentId)
    parentId = controls.find((candidate) => candidate.id === parentId)?.parentId
  }
  return false
}

function repeatingSectionAncestor(controls: readonly DocumentContentControl[], control: DocumentContentControl): DocumentContentControl | undefined {
  let parentId = control.parentId
  const visited = new Set<string>()
  while (parentId && !visited.has(parentId)) {
    visited.add(parentId)
    const parent = controls.find(candidate => candidate.id === parentId)
    if (!parent) return undefined
    if (parent.type === 'repeatingSection') return parent
    parentId = parent.parentId
  }
  return undefined
}

function newContentControlId(): string {
  return documentId('content-control')
}

function removeControlTreeFromModel(model: DocumentModel, control: DocumentContentControl): void {
  const removing = new Set([control.id])
  let changed = true
  while (changed) {
    changed = false
    for (const candidate of model.contentControls ?? []) {
      if (candidate.parentId && removing.has(candidate.parentId) && !removing.has(candidate.id)) {
        if (isControlLocked(candidate)) throw new Error(`Nested content control wrapper is locked: ${candidate.id}`)
        removing.add(candidate.id)
        changed = true
      }
    }
  }
  model.contentControls = (model.contentControls ?? []).filter((candidate) => !removing.has(candidate.id))
}

function isTextEditBlockedByControl(
  model: DocumentModel,
  control: DocumentContentControl,
  localStart: number,
  localEnd: number,
  insertion: boolean,
  controlStart = control.startOffset,
  controlEnd = control.endOffset,
  globalStart = localStart,
  globalEnd = localEnd,
): boolean {
  if (control.type === 'group') {
    const children = (model.contentControls ?? []).filter((candidate) => candidate.parentId === control.id
      && !candidate.blockRange && !candidate.tableRowRange && candidate.blockId === control.blockId)
    const withinChild = children.some((child) => insertion
      ? localStart > child.startOffset && localStart < child.endOffset
      : localStart >= child.startOffset && localEnd <= child.endOffset)
    if (withinChild) return false
  }
  return insertion
    ? globalStart > controlStart && globalStart < controlEnd
    : globalStart < controlEnd && controlStart < globalEnd
}

function isContentControlLockedAtEdit(model: DocumentModel, control: DocumentContentControl, start: number, end: number, insertion: boolean,
  references?: readonly DocumentNoteReferenceOffset[], startRank?: number, endRank?: number): boolean {
  if (!isContentLocked(control)) return false
  if (references) {
    const range = documentContentControlNoteBoundaryRanks(control, references)
    const from = assertDocumentNoteReferenceBoundaryRank(references, start, startRank), to = assertDocumentNoteReferenceBoundaryRank(references, end, endRank)
    if (control.type === 'group' && (model.contentControls ?? []).some(child => child.parentId === control.id
      && !child.blockRange && !child.tableRowRange && child.blockId === control.blockId
      && documentContentControlContainsNoteEdit(child, references, start, from, end, to))) return false
    return insertion ? compareDocumentNoteBoundaries(start, from, control.startOffset, range.start) > 0
      && compareDocumentNoteBoundaries(start, from, control.endOffset, range.end) < 0
      : compareDocumentNoteBoundaries(start, from, control.endOffset, range.end) < 0 && compareDocumentNoteBoundaries(control.startOffset, range.start, end, to) < 0
  }
  if (control.type === 'group') {
    const children = (model.contentControls ?? []).filter((candidate) => candidate.parentId === control.id
      && !candidate.blockRange && !candidate.tableRowRange && candidate.blockId === control.blockId)
    if (children.some((child) => insertion
      ? start > child.startOffset && start < child.endOffset
      : start >= child.startOffset && end <= child.endOffset)) return false
  }
  return insertion
    ? start > control.startOffset && start < control.endOffset
    : start < control.endOffset && control.startOffset < end
}

function documentContentControlContainsNoteEdit(control: DocumentContentControl, references: readonly DocumentNoteReferenceOffset[],
  start: number, startRank: number, end: number, endRank: number): boolean {
  const range = documentContentControlNoteBoundaryRanks(control, references)
  return compareDocumentNoteBoundaries(control.startOffset, range.start, start, startRank) <= 0
    && compareDocumentNoteBoundaries(end, endRank, control.endOffset, range.end) <= 0
}

function isStructuralContentControlLockedAtBlockEdit(model: DocumentModel, control: DocumentContentControl, blockId: string, start: number, end: number): boolean {
  if (!isContentLocked(control)) return false
  if (control.type === 'group' && (model.contentControls ?? []).some(child => child.parentId === control.id
    && documentContentControlContainsBlockRange(model, child, blockId, start, end))) return false
  return true
}

/** A whole-block operation cannot treat a partial inline child as ownership of its paragraph. */
function fullChildBlockIds(model: DocumentModel, child: DocumentContentControl): Set<string> {
  if (child.blockRange) return documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, child.blockRange))
  if (child.tableRowRange) return documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(model, child.tableRowRange))
  const paragraph = findBlock(model, child.blockId)
  if (!paragraph) return new Set()
  const references = documentNoteReferenceOffsets(paragraph.runs), range = documentContentControlNoteBoundaryRanks(child, references)
  return child.startOffset === 0 && range.start === 0 && child.endOffset === runText(paragraph.runs).length
    && range.end === references.length ? new Set([child.blockId]) : new Set()
}

function groupOwnsWholeBlockEdit(model: DocumentModel, group: DocumentContentControl, touched: readonly string[]): boolean {
  return group.type === 'group' && (model.contentControls ?? []).some(child => {
    if (child.parentId !== group.id || !documentContentControlContains(model, group, child)) return false
    const owned = fullChildBlockIds(model, child)
    return touched.every(id => owned.has(id))
  })
}

/** Replacement may re-anchor a row fallback; removal of its table still respects the wrapper lock. */
export function assertDocumentContentControlBlocksEditable(model: DocumentModel, blockIds: ReadonlySet<string>): void {
  for (const control of model.contentControls ?? []) {
    if (control.blockRange && findBlock(model, control.blockRange.blockIds[0]!)) {
      const range = resolveDocumentContentControlBlocks(model, control.blockRange), owned = documentContentControlOwnedBlockIds(range)
      const touched = [...owned].filter(id => blockIds.has(id))
      if (!touched.length) continue
      if (range.blockIds.every(id => blockIds.has(id)) && (isControlLocked(control) || isContentLocked(control))) {
        throw new RangeError(`Cannot remove a locked complete block content control: ${control.id}`)
      }
      if (isContentLocked(control)) {
        if (!groupOwnsWholeBlockEdit(model, control, touched)) throw new RangeError(`Complete block content control content is locked: ${control.id}`)
      }
      continue
    }
    if (!control.tableRowRange || !findBlock(model, control.tableRowRange.tableId)) continue
    if (blockIds.has(control.tableRowRange.tableId) && (isControlLocked(control) || isContentLocked(control))) {
      throw new RangeError(`Cannot remove a locked content control table row wrapper: ${control.id}`)
    }
    const touched = [...documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(model, control.tableRowRange))].filter(id => blockIds.has(id))
    if (isContentLocked(control) && touched.length && !groupOwnsWholeBlockEdit(model, control, touched)) throw new RangeError(`Content control row content is locked: ${control.id}`)
  }
}
