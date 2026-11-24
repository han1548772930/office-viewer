import type {
  DocumentBlock,
  DocumentContentControl,
  DocumentModel,
} from '../../../domain/model'
import { documentTextFieldParagraphs, documentTextFields, insertDocumentTextFields, setDocumentTextFields, transformDocumentTextFieldsForBlockEdit, transformDocumentTextFieldsForBlockRangeEdit, type DocumentTextFieldParagraph } from '../model/document-text-fields'
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
  type RunPatch,
  type TextTarget,
} from '../model/document-model'
import type { DocumentImage } from '../model/document-model'
import { isDocumentTextBoundary } from '../model/document-story'
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
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/document-comment-anchors'

export type ContentControlInput = Omit<DocumentContentControl, 'id'> & { id?: string }

export interface ContentControlRangeEditOptions {
  /** A checked input context explicitly addresses the contents at start/end boundaries. */
  activeControlId?: string
}

function insideControlEditChain(model: DocumentModel, blockId: string, startOffset: number, endOffset: number,
  activeControlId: string | undefined): Set<string> {
  const result = new Set<string>()
  if (!activeControlId) return result
  const index = new Map((model.contentControls ?? []).map(control => [control.id, control]))
  let current = index.get(activeControlId)
  if (!current || current.metadata?.officeStoryLockSource || current.blockId !== blockId
    || !Number.isSafeInteger(startOffset) || !Number.isSafeInteger(endOffset)
    || startOffset < current.startOffset || endOffset > current.endOffset || startOffset > endOffset) {
    throw new RangeError('内容控件内部编辑上下文已失效：' + activeControlId)
  }
  while (current) {
    if (result.has(current.id)) throw new RangeError('内容控件内部编辑遇到循环父级：' + current.id)
    if (current.blockId !== blockId || startOffset < current.startOffset || endOffset > current.endOffset) {
      throw new RangeError('内容控件内部编辑不属于其父级范围：' + current.id)
    }
    result.add(current.id)
    const parentId: string | undefined = current.parentId
    if (parentId && !index.has(parentId)) throw new RangeError('内容控件内部编辑的父级不存在：' + parentId)
    current = parentId ? index.get(parentId) : undefined
  }
  return result
}

export type ContentControlPatch = Partial<Omit<DocumentContentControl, 'id' | 'blockId' | 'startOffset' | 'endOffset' | 'value' | 'checked' | 'dateValue'>>

const lockedContent = new Set<DocumentContentControl['lock']>(['contentLocked', 'sdtContentLocked'])
const lockedControl = new Set<DocumentContentControl['lock']>(['controlLocked', 'sdtLocked', 'sdtContentLocked'])

function assertControlAcceptsParagraphBreak(control: DocumentContentControl): void {
  if (!['richText', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj'].includes(control.type)
    && !(control.type === 'text' && control.multiLine === true)) throw new RangeError(`此类型内容控件不接受段落换行：${control.id}`)
}

function editHostedControl<T>(model: DocumentModel, blockId: string, edit: (scoped: DocumentModel) => T, requestedHost?: DocumentAuxiliaryBlockHost): { value: T } | undefined {
  const host = requestedHost ?? documentBlockHostForId(model, blockId)
  if (!host) return undefined
  const candidate = cloneModel(model), scoped = documentModelForBlockHost(candidate, host, true), value = edit(scoped)
  applyDocumentBlockHostModel(candidate, host, scoped)
  if (host.kind === 'story') assertDocumentStoryLocksPreserved(resolveSectionHeaderFooter(model, host.sectionId), resolveSectionHeaderFooter(candidate, host.sectionId), host.story, model)
  Object.assign(model, candidate)
  return { value }
}

function hostedControlBlock(model: DocumentModel, id: string): string | undefined {
  return model.contentControls?.find(control => control.id === id)?.blockId
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
  const controls = model.contentControls ?? []
  return controls.filter(control => control.blockId === blockId
    && startOffset >= control.startOffset && endOffset <= control.endOffset)
    .sort((left, right) => {
      const lengthDifference = (left.endOffset - left.startOffset) - (right.endOffset - right.startOffset)
      if (lengthDifference) return lengthDifference
      if (hasAncestor(controls, left, right.id)) return -1
      if (hasAncestor(controls, right, left.id)) return 1
      return 0
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
  const hosted = editHostedControl(model, input.blockId, scoped => createDocumentContentControl(scoped, input, id), host)
  if (hosted) return hosted.value
  const { id: _inputId, blockId: _blockId, startOffset: _start, endOffset: _end, value: _value, checked: _checked, dateValue: _dateValue, ...properties } = input
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
  validateControl(candidate, control)
  const controls = candidate.contentControls ?? (candidate.contentControls = [])
  assertNoPartialOverlap(controls, control)
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
  Object.assign(model, candidate)
  return control
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
  assertNoPartialOverlap(controls.filter((candidate) => candidate.id !== id), next)
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
  if (!control.showingPlaceholder && previousText === displayValue && control.value === storedValue
    && (control.type !== 'date' || control.dateValue === value)) return
  const candidate = cloneModel(model), next = requireControl(candidate, id)
  if (previousText !== displayValue) replaceControlText(candidate, next, displayValue)
  next.value = storedValue
  if (next.type === 'date') next.dateValue = value
  next.showingPlaceholder = false
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
  if (!control.showingPlaceholder && previousText === glyph.text && control.value === storedValue
    && control.checked === checked && control.checkbox?.checked === checked) return
  const candidate = cloneModel(model), next = requireControl(candidate, id)
  next.checked = checked
  next.value = storedValue
  next.checkbox = { ...next.checkbox, checked }
  if (previousText !== glyph.text) replaceControlText(candidate, next, glyph.text, { fontFamily: glyph.fontFamily })
  next.showingPlaceholder = false
  synchronizeDocumentContentControlsAfterEdit(model, candidate, { origin: 'control-value', controlId: id })
  Object.assign(model, candidate)
}

export function resetDocumentContentControlPlaceholder(model: DocumentModel, id: string, useDefaultPlaceholder = false): void {
  const blockId = hostedControlBlock(model, id)
  if (blockId && editHostedControl(model, blockId, scoped => resetDocumentContentControlPlaceholder(scoped, id, useDefaultPlaceholder))) return
  const control = requireControl(model, id)
  const placeholder = useDefaultPlaceholder ? documentContentControlPlaceholderText(control) : control.placeholder
  if (!placeholder) throw new Error(`Content control has no placeholder: ${id}`)
  if (control.showingPlaceholder) return
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
  if (control.type === 'picture') {
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
  const control = requireControl(model, id)
  if (control.type !== 'picture') throw new TypeError(`Content control ${id} is not a picture control`)
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
  const existingPictureId = typeof control.metadata?.pictureBlockId === 'string' ? control.metadata.pictureBlockId : undefined
  const existingPicture = existingPictureId ? findBlock(model, existingPictureId) : undefined
  if (existingPicture?.type === 'image') {
    existingPicture.metadata = { ...existingPicture.metadata, image: { ...image } }
    control.metadata = { ...control.metadata, pictureImage: { ...image } }
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
  picture.metadata = { image: { ...image }, contentControlId: control.id, anchorBlockId: block.id, anchorOffset: startOffset, tailBlockId: tail.id }
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

export function duplicateRepeatingSectionItem(model: DocumentModel, id: string): DocumentContentControl {
  const blockId = hostedControlBlock(model, id), hosted = blockId && editHostedControl(model, blockId, scoped => duplicateRepeatingSectionItem(scoped, id))
  if (hosted) return hosted.value
  const item = requireControl(model, id)
  if (item.type !== 'repeatingSectionItem') throw new TypeError(`Content control ${id} is not a repeating-section item`)
  if (['row', 'cell', 'table'].includes(item.placement ?? '')) throw new TypeError('This repeating-section placement cannot be duplicated')
  const section = repeatingSectionAncestor(model.contentControls ?? [], item)
  if (!section) throw new Error(`Repeating-section parent is missing: ${id}`)
  if (section.repeatingSection?.doNotAllowInsertDelete) throw new Error(`Repeating-section items cannot be changed: ${section.id}`)
  if (isControlLocked(item) || isContentLocked(item) || isContentLocked(section)) throw new Error(`Repeating-section item is locked: ${id}`)
  assertAncestorContentEditable(model, item)

  const target = targetForControl(model, item)
  if (!target) throw new Error(`Content control paragraph not found: ${item.blockId}`)
  const runs = runsForTarget(model, target)
  const text = sliceRuns(runs, item.startOffset, item.endOffset)
  const textLength = item.endOffset - item.startOffset
  const insertionOffset = item.endOffset
  const paragraph = findBlock(model, item.blockId)!
  const copiedFields = documentTextFields(paragraph).filter(field => (field.endBlockId === undefined || field.endBlockId === paragraph.id) && field.startOffset >= item.startOffset && field.endOffset <= item.endOffset)
    .map(field => ({ ...field, id: documentId('field'), startOffset: field.startOffset - item.startOffset, endOffset: field.endOffset - item.startOffset }))
  transformDocumentTextFieldsForBlockEdit(paragraph, insertionOffset, insertionOffset, runText(text), undefined, model)
  setTargetRuns(model, target, [
    ...sliceRuns(runs, 0, insertionOffset),
    ...text,
    ...sliceRuns(runs, insertionOffset),
  ])
  insertDocumentTextFields(paragraph, copiedFields, insertionOffset)

  const controls = model.contentControls ?? []
  const descendants = controls.filter((control) => control.blockId === item.blockId
    && control.startOffset >= item.startOffset && control.endOffset <= item.endOffset
    && (control.id === item.id || hasAncestor(controls, control, item.id)))
  const idMap = new Map(descendants.map((control) => [control.id, newContentControlId()]))
  for (const control of controls) {
    if (control.id === item.id || hasAncestor(controls, control, item.id)) continue
    if (control.blockId !== item.blockId) continue
    if (control.startOffset >= insertionOffset) {
      control.startOffset += textLength
      control.endOffset += textLength
    } else if (control.startOffset <= item.startOffset && control.endOffset >= insertionOffset) {
      control.endOffset += textLength
    }
  }
  const copies = descendants.map((control) => ({
    ...cloneControl(control),
    id: idMap.get(control.id)!,
    startOffset: insertionOffset + control.startOffset - item.startOffset,
    endOffset: insertionOffset + control.endOffset - item.startOffset,
    parentId: control.id === item.id ? item.parentId : idMap.get(control.parentId ?? '') ?? control.parentId,
  }))
  controls.push(...copies)
  controls.sort((left, right) => left.blockId.localeCompare(right.blockId) || left.startOffset - right.startOffset || left.id.localeCompare(right.id))
  return cloneControl(copies.find((control) => control.type === 'repeatingSectionItem')!)
}

export function deleteRepeatingSectionItem(model: DocumentModel, id: string): void {
  const blockId = hostedControlBlock(model, id)
  if (blockId && editHostedControl(model, blockId, scoped => deleteRepeatingSectionItem(scoped, id))) return
  const item = requireControl(model, id)
  if (item.type !== 'repeatingSectionItem') throw new TypeError(`Content control ${id} is not a repeating-section item`)
  if (['cell', 'table'].includes(item.placement ?? '')) throw new TypeError('This repeating-section placement cannot be deleted')
  const section = repeatingSectionAncestor(model.contentControls ?? [], item)
  if (!section) throw new Error(`Repeating-section parent is missing: ${id}`)
  if (section.repeatingSection?.doNotAllowInsertDelete) throw new Error(`Repeating-section items cannot be changed: ${section.id}`)
  const removed = (model.contentControls ?? []).filter((control) => control.blockId === item.blockId
    && control.startOffset >= item.startOffset && control.endOffset <= item.endOffset
    && (control.id === item.id || hasAncestor(model.contentControls ?? [], control, item.id)))
  if (removed.some((control) => isControlLocked(control) || isContentLocked(control))) throw new Error(`Repeating-section item is locked: ${id}`)
  assertAncestorContentEditable(model, item)
  const target = targetForControl(model, item)
  if (!target) throw new Error(`Content control paragraph not found: ${item.blockId}`)
  const runs = runsForTarget(model, target)
  const startOffset = item.startOffset
  const endOffset = item.endOffset
  const textLength = item.endOffset - item.startOffset
  for (const control of model.contentControls ?? []) {
    if (control.id === item.id || control.blockId !== item.blockId || control.endOffset <= startOffset || control.startOffset >= endOffset) continue
    if (hasAncestor(model.contentControls ?? [], item, control.id)) continue
    if (control.startOffset >= startOffset && control.endOffset <= endOffset) continue
    throw new RangeError(`Repeating-section deletion cannot split another content control: ${control.id}`)
  }
  const paragraph = findBlock(model, item.blockId)!
  const removedFields = new Set(documentTextFields(paragraph).filter(field => (field.endBlockId === undefined || field.endBlockId === paragraph.id) && field.startOffset >= startOffset && field.endOffset <= endOffset).map(field => field.id))
  transformDocumentTextFieldsForBlockEdit(paragraph, startOffset, endOffset, '', undefined, model)
  setTargetRuns(model, target, [
    ...sliceRuns(runs, 0, startOffset),
    ...sliceRuns(runs, endOffset),
  ])
  setDocumentTextFields(paragraph, documentTextFields(paragraph).filter(field => !removedFields.has(field.id)))
  const removedIds = new Set(removed.map((control) => control.id))
  model.contentControls = (model.contentControls ?? []).filter((control) => !removedIds.has(control.id))
  for (const control of model.contentControls) {
    if (control.blockId !== item.blockId) continue
    if (control.startOffset >= endOffset) {
      control.startOffset -= textLength
      control.endOffset -= textLength
    } else if (control.startOffset <= startOffset && control.endOffset >= endOffset) {
      control.endOffset -= textLength
    }
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
  const host = documentBlockHostForId(model, control.blockId)
  if (host) return contentControlText(documentModelForBlockHost(model, host), control)
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control paragraph not found: ${control.blockId}`)
  return runText(sliceRuns(runsForTarget(model, target), control.startOffset, control.endOffset))
}

export function contentControlRange(control: DocumentContentControl): { blockId: string; startOffset: number; endOffset: number } {
  return { blockId: control.blockId, startOffset: control.startOffset, endOffset: control.endOffset }
}

/** Shared capability check for Canvas controls and the actual value commands. */
export function assertDocumentContentControlValueEditable(model: DocumentModel, id: string): void {
  const control = requireControl(model, id), host = documentBlockHostForId(model, control.blockId)
  if (host) { assertDocumentContentControlValueEditable(documentModelForBlockHost(model, host, true), id); return }
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
}

export function assertTextRangeEditable(model: DocumentModel, startOffset: number, endOffset: number): void {
  const body = flattenBodySegments(model)
  for (const control of model.contentControls ?? []) {
    if (!isContentLocked(control)) continue
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
  const control = model.contentControls.find((candidate) => affectedIds.has(candidate.blockId)
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
  if (!model.contentControls?.length || (startOffset === endOffset && replacementLength === 0 && !split)) return
  // Validate every affected range on a detached plan before changing any live anchors.
  const next: DocumentModel = { ...model, contentControls: getDocumentContentControls(model) }
  const selectedAtEmptyInsertion = !preserveControlRanges && startOffset === endOffset
    ? selectedDocumentContentControl(model, blockId, startOffset) : undefined
  const emptyInsertionControl = selectedAtEmptyInsertion?.startOffset === selectedAtEmptyInsertion?.endOffset
    ? selectedAtEmptyInsertion : undefined
  const inside = insideControlEditChain(model, blockId, startOffset, endOffset, options.activeControlId ?? emptyInsertionControl?.id)
  for (const control of next.contentControls!) {
    if (!next.contentControls!.some((candidate) => candidate.id === control.id)) continue
    if (control.blockId !== blockId) continue
    const insertion = startOffset === endOffset
    const overlaps = insertion
      ? startOffset > control.startOffset && startOffset < control.endOffset
      : startOffset < control.endOffset && control.startOffset < endOffset
    const removesWholeControl = !preserveControlRanges && !inside.has(control.id) && overlaps
      && startOffset <= control.startOffset && endOffset >= control.endOffset && replacementLength === 0
    if (removesWholeControl) {
      if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
      removeControlTreeFromModel(next, control)
      continue
    }
    if (overlaps && isContentControlLockedAtEdit(model, control, startOffset, endOffset, insertion)) throw new Error(`Content control content is locked: ${control.id}`)
    if (insertion && inside.has(control.id)) {
      if (isContentLocked(control)) throw new Error(`Content control content is locked: ${control.id}`)
      assertAncestorContentEditable(model, control)
      control.endOffset += replacementLength
      continue
    }
    if (!overlaps) {
      if (!insertion && endOffset <= control.startOffset) {
        const delta = replacementLength - (endOffset - startOffset)
        control.startOffset += delta
        control.endOffset += delta
      } else if (insertion && startOffset <= control.startOffset) {
        control.startOffset += replacementLength
        control.endOffset += replacementLength
      }
      continue
    }
    if (isContentControlLockedAtEdit(model, control, startOffset, endOffset, insertion)) throw new Error(`Content control content is locked: ${control.id}`)
    if (insertion) {
      control.endOffset += replacementLength
      continue
    }
    const delta = replacementLength - (endOffset - startOffset)
    if (startOffset <= control.startOffset && endOffset >= control.endOffset) {
      control.startOffset = startOffset
      control.endOffset = startOffset + replacementLength
    } else {
      const oldStart = control.startOffset
      const oldEnd = control.endOffset
      control.startOffset = oldStart >= endOffset ? oldStart + delta : oldStart >= startOffset ? startOffset : oldStart
      control.endOffset = oldEnd >= endOffset ? oldEnd + delta : oldEnd > startOffset ? startOffset + replacementLength : oldEnd
      if (control.endOffset < control.startOffset) control.endOffset = control.startOffset
    }
  }
  if (split) {
    const unwrapping = new Set<string>()
    if (!preserveControlRanges) {
      for (const control of next.contentControls!) {
        if (control.blockId !== blockId || control.endOffset <= split.firstBreakOffset
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
    const moving = controls.filter((control) => control.blockId === blockId && movesToSuffix(control))
    for (const control of moving) {
      control.blockId = split.blockId
      control.startOffset -= split.lastParagraphOffset
      control.endOffset -= split.lastParagraphOffset
    }
  }
  model.contentControls = next.contentControls
}

export function assertContentControlBlockRemovable(model: DocumentModel, blockId: string): void {
  const controls = model.contentControls?.filter((candidate) => candidate.blockId === blockId) ?? []
  for (const control of controls) if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
  for (const control of controls) removeControlTreeFromModel(model, control)
}

/** Validate mapped ranges and the wrapper tree before publishing a structural edit. */
export function validateDocumentContentControlRanges(model: DocumentModel): void {
  const paragraphs = new Map(documentTextFieldParagraphs(model).map(paragraph => [paragraph.block.id, paragraph]))
  for (const control of model.contentControls ?? []) {
    const host = documentBlockHostForId(model, control.blockId), scoped = host ? documentModelForBlockHost(model, host) : model
    validateControl(scoped, cloneControl(control))
    const paragraph = host ? documentTextFieldParagraphs(scoped).find(candidate => candidate.block.id === control.blockId) : paragraphs.get(control.blockId)
    if (paragraph && (!isDocumentTextBoundary(paragraph.text, control.startOffset) || !isDocumentTextBoundary(paragraph.text, control.endOffset))) {
      throw new RangeError(`Content control UTF-16 range splits a character: ${control.id}`)
    }
    assertNoPartialOverlap(model.contentControls ?? [], control)
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
  const source = new Map(sourceParagraphs.map(paragraph => [paragraph.block.id, paragraph]))
  const destination = documentTextFieldParagraphs(model)
  const next: DocumentModel = { ...model, contentControls: getDocumentContentControls(model) }
  const delta = replacementLength - (endOffset - startOffset)
  const insertion = startOffset === endOffset
  const emptyInputControl = !preserveControlRanges && insertion ? (model.contentControls ?? [])
    .filter(control => control.startOffset === control.endOffset && source.get(control.blockId)?.absoluteStart !== undefined
      && source.get(control.blockId)!.absoluteStart + control.startOffset === startOffset)
    .sort((left, right) => hasAncestor(model.contentControls ?? [], left, right.id) ? -1 : hasAncestor(model.contentControls ?? [], right, left.id) ? 1 : 0)[0] : undefined
  const inputControlId = options.activeControlId ?? emptyInputControl?.id
  const inputControl = inputControlId ? model.contentControls?.find(control => control.id === inputControlId) : undefined
  const inputParagraph = inputControl && source.get(inputControl.blockId)
  if (inputControlId && (!inputControl || !inputParagraph)) throw new RangeError('内容控件内部编辑的源段落不存在：' + inputControlId)
  const inside = inputControl && inputParagraph
    ? insideControlEditChain(model, inputControl.blockId, startOffset - inputParagraph.absoluteStart, endOffset - inputParagraph.absoluteStart, inputControlId)
    : new Set<string>()
  const unwrapping = new Set<string>()
  for (const control of [...next.contentControls!]) {
    if (!next.contentControls!.some(candidate => candidate.id === control.id)) continue
    const paragraph = source.get(control.blockId)
    if (!paragraph) continue
    const from = paragraph.absoluteStart + control.startOffset, to = paragraph.absoluteStart + control.endOffset
    const overlaps = insertion ? startOffset > from && startOffset < to : startOffset < to && from < endOffset
    if (!preserveControlRanges && !inside.has(control.id) && overlaps && startOffset <= from && endOffset >= to && replacementLength === 0) {
      if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
      removeControlTreeFromModel(next, control)
      continue
    }
    if (isContentControlLockedAtEdit(model, control, startOffset - paragraph.absoluteStart, endOffset - paragraph.absoluteStart, insertion)) {
      throw new Error(`Content control content is locked: ${control.id}`)
    }
    let start = from, end = to
    if (insertion && inside.has(control.id)) {
      if (isContentLocked(control)) throw new Error(`Content control content is locked: ${control.id}`)
      assertAncestorContentEditable(model, control)
      end += replacementLength
    } else if (insertion) {
      if (startOffset <= from) { start += replacementLength; end += replacementLength }
      else if (startOffset < to) end += replacementLength
    } else if (endOffset <= from) { start += delta; end += delta }
    else if (startOffset < to && from < endOffset) {
      start = from >= startOffset ? startOffset : from
      end = to >= endOffset ? to + delta : startOffset + replacementLength
      if (end < start) end = start
    }
    const first = destination.find(candidate => start >= candidate.absoluteStart && start <= candidate.absoluteEnd)
    const last = destination.find(candidate => end >= candidate.absoluteStart && end <= candidate.absoluteEnd)
    if (!first || !last) throw new RangeError(`正文编辑无法把内容控件 ${control.id} 映射到同一段落。`)
    if (first.block.id !== last.block.id) {
      if (preserveControlRanges || !control.temporary || isControlLocked(control) || control.metadata?.officeStoryLockSource) {
        throw new RangeError(`正文编辑无法把内容控件 ${control.id} 映射到同一段落。`)
      }
      if (isContentLocked(control)) throw new Error(`Content control content is locked: ${control.id}`)
      assertAncestorContentEditable(model, control)
      assertControlAcceptsParagraphBreak(control)
      unwrapping.add(control.id)
      continue
    }
    control.blockId = first.block.id
    control.startOffset = start - first.absoluteStart
    control.endOffset = end - last.absoluteStart
  }
  model.contentControls = unwrapDocumentContentControlShells(next.contentControls!, unwrapping)
}

function replaceControlText(model: DocumentModel, control: DocumentContentControl, value: string, marks: RunPatch = {}): void {
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control paragraph not found: ${control.blockId}`)
  const oldLength = control.endOffset - control.startOffset
  const editStart = control.startOffset, editEnd = control.endOffset
  for (const candidate of model.contentControls ?? []) {
    if (candidate.id === control.id || candidate.blockId !== control.blockId) continue
    if (candidate.parentId === control.id || (candidate.startOffset >= control.startOffset && candidate.endOffset <= control.endOffset
      && (candidate.startOffset > control.startOffset || candidate.endOffset < control.endOffset))) {
      throw new RangeError(`Cannot replace content control text while a nested control is present: ${candidate.id}`)
    }
    if (candidate.startOffset < control.endOffset && control.startOffset < candidate.endOffset
      && !(candidate.startOffset <= control.startOffset && candidate.endOffset >= control.endOffset)) {
      throw new RangeError(`Cannot replace content control text across a partially overlapping control: ${candidate.id}`)
    }
  }
  const runs = runsForTarget(model, target)
  const { before: prefix, after: suffix } = splitDocumentRunRange(runs, control.startOffset, control.endOffset)
  const inherited = documentRunStyleAt(runs, control.startOffset, control.endOffset) ?? {}
  const inserted: DocumentRun[] = value ? [createRun(value, { ...inherited, ...marks })] : []
  const block = findBlock(model, control.blockId)!
  transformDocumentTextFieldsForBlockEdit(block, control.startOffset, control.endOffset, value, undefined, model)
  const nextRuns = [...prefix, ...inserted, ...suffix]
  // The editor's shared range splitter keeps note/field entities at the boundaries exactly once.
  block.runs = nextRuns.length ? nextRuns : [createRun('', inherited)]
  const code = block.metadata.codeBlock
  if (code && typeof code === 'object' && !Array.isArray(code) && typeof Reflect.get(code, 'source') === 'string') {
    block.metadata = { ...block.metadata, codeBlock: { ...code, source: runText(block.runs) } }
  }
  transformDocumentCommentAnchorsForTextEdit(model, block.id, editStart, editEnd, value)
  for (const ancestor of documentBlockAncestry(model, block.id) ?? []) delete ancestor.metadata.html
  const delta = value.length - oldLength
  if (delta === 0) return
  for (const candidate of model.contentControls ?? []) {
    if (candidate.id === control.id || candidate.blockId !== control.blockId) continue
    if (candidate.startOffset <= control.startOffset && candidate.endOffset >= control.endOffset
      && hasAncestor(model.contentControls ?? [], control, candidate.id)) {
      candidate.endOffset += delta
    } else if (candidate.startOffset >= control.endOffset) {
      candidate.startOffset += delta
      candidate.endOffset += delta
    } else if (candidate.startOffset <= control.startOffset && candidate.endOffset >= control.endOffset) {
      candidate.endOffset += delta
    } else if (candidate.startOffset >= control.startOffset && candidate.endOffset <= control.endOffset) {
      throw new RangeError(`Cannot replace content control text while a nested control is present: ${candidate.id}`)
    } else if (candidate.startOffset < control.endOffset && control.startOffset < candidate.endOffset) {
      throw new RangeError(`Cannot replace content control text across a partially overlapping control: ${candidate.id}`)
    }
  }
  control.endOffset += delta
}

function targetForControl(model: DocumentModel, control: DocumentContentControl): TextTarget | undefined {
  return textTargets(model).find((target) => (target.cellBlockId ?? target.blockId) === control.blockId)
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
  if (!['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox', 'picture', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj'].includes(control.type)) throw new TypeError('Unsupported content control type')
  if (control.lock === undefined) control.lock = 'unlocked'
  if (control.type === 'date' && control.dateValue !== undefined && control.dateValue !== '' && !validContentControlIsoDate(control.dateValue)) throw new RangeError('Date content control value must be YYYY-MM-DD')
  if ((control.type === 'dropDownList' || control.type === 'comboBox') && (!control.items?.length || control.items.some((item) => !item.value))) throw new RangeError('List content controls require non-empty values')
  if (control.type === 'dropDownList' && !control.showingPlaceholder && control.value !== undefined && !control.items?.some((item) => item.value === control.value)) throw new RangeError(`Dropdown value is not present in its item list: ${control.value}`)
  if (control.type === 'checkbox' && control.checked === undefined) control.checked = false
  if (control.parentId) {
    if (control.parentId === control.id || hasAncestor(model.contentControls ?? [], control, control.id)) throw new RangeError('Content control parent relationships must not contain a cycle')
    const parent = model.contentControls?.find((candidate) => candidate.id === control.parentId)
    if (!parent || !['group', 'repeatingSection', 'repeatingSectionItem'].includes(parent.type)) throw new Error(`Invalid content control parent: ${control.parentId}`)
    if (parent.blockId !== control.blockId || parent.startOffset > control.startOffset || parent.endOffset < control.endOffset) throw new RangeError(`Child content control must be within its parent: ${control.parentId}`)
  }
}

function assertNoPartialOverlap(controls: DocumentContentControl[], next: DocumentContentControl): void {
  for (const current of controls) {
    if (current.blockId !== next.blockId || current.id === next.id) continue
    const overlaps = next.startOffset < current.endOffset && current.startOffset < next.endOffset
    if (!overlaps) continue
    const nextContains = next.startOffset <= current.startOffset && next.endOffset >= current.endOffset
    const currentContains = current.startOffset <= next.startOffset && current.endOffset >= next.endOffset
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
    if (candidate.id === control.id || candidate.blockId !== control.blockId) continue
    if (isContentControlLockedAtEdit(model, candidate, control.startOffset, control.endOffset, control.startOffset === control.endOffset)) {
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
    const children = (model.contentControls ?? []).filter((candidate) => candidate.parentId === control.id)
    const withinChild = children.some((child) => insertion
      ? localStart > child.startOffset && localStart < child.endOffset
      : localStart >= child.startOffset && localEnd <= child.endOffset)
    if (withinChild) return false
  }
  return insertion
    ? globalStart > controlStart && globalStart < controlEnd
    : globalStart < controlEnd && controlStart < globalEnd
}

function isContentControlLockedAtEdit(model: DocumentModel, control: DocumentContentControl, start: number, end: number, insertion: boolean): boolean {
  if (!isContentLocked(control)) return false
  if (control.type === 'group') {
    const children = (model.contentControls ?? []).filter((candidate) => candidate.parentId === control.id)
    if (children.some((child) => insertion
      ? start > child.startOffset && start < child.endOffset
      : start >= child.startOffset && end <= child.endOffset)) return false
  }
  return insertion
    ? start > control.startOffset && start < control.endOffset
    : start < control.endOffset && control.startOffset < end
}
