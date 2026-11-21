import type {
  DocumentBlock,
  DocumentContentControl,
  DocumentModel,
} from '../../../domain/model'
import {
  compactRuns,
  createParagraph,
  createRun,
  documentId,
  runText,
  runsForTarget,
  setTargetRuns,
  sliceRuns,
  tableForBlock,
  textTargets,
  type DocumentRun,
  type TextTarget,
} from '../model/document-model'
import type { DocumentImage } from '../model/document-model'

export type ContentControlInput = Omit<DocumentContentControl, 'id'> & { id?: string }

export type ContentControlPatch = Partial<Omit<DocumentContentControl, 'id' | 'blockId' | 'startOffset' | 'endOffset' | 'value' | 'checked' | 'dateValue'>>

const lockedContent = new Set<DocumentContentControl['lock']>(['contentLocked', 'sdtContentLocked'])
const lockedControl = new Set<DocumentContentControl['lock']>(['controlLocked', 'sdtLocked', 'sdtContentLocked'])

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
): DocumentContentControl {
  if (!id.trim()) throw new TypeError('Content control id must not be empty')
  if (model.contentControls?.some((control) => control.id === id)) throw new Error(`Content control already exists: ${id}`)
  const control: DocumentContentControl = {
    ...input,
    id,
    lock: input.lock ?? 'unlocked',
    showingPlaceholder: input.showingPlaceholder ?? false,
    items: input.items?.map((item) => ({ ...item })),
    metadata: input.metadata ? { ...input.metadata } : undefined,
  }
  validateControl(model, control)
  const controls = model.contentControls ?? (model.contentControls = [])
  assertNoPartialOverlap(controls, control)
  controls.push(control)
  controls.sort((left, right) => left.blockId.localeCompare(right.blockId) || left.startOffset - right.startOffset || left.id.localeCompare(right.id))
  if (control.showingPlaceholder && control.placeholder && ['text', 'richText', 'date', 'dropDownList', 'comboBox'].includes(control.type)) {
    replaceControlText(model, control, control.placeholder)
    control.value = control.placeholder
  }
  return control
}

export function updateDocumentContentControl(model: DocumentModel, id: string, patch: ContentControlPatch): void {
  const control = requireControl(model, id)
  if (hasControlStructurePatch(control, patch) && isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${id}`)
  const next: DocumentContentControl = {
    ...control,
    ...patch,
    id: control.id,
    items: patch.items?.map((item) => ({ ...item })) ?? control.items?.map((item) => ({ ...item })),
    metadata: patch.metadata ? { ...patch.metadata } : control.metadata ? { ...control.metadata } : undefined,
  }
  validateControl(model, next)
  const controls = model.contentControls!
  assertNoPartialOverlap(controls.filter((candidate) => candidate.id !== id), next)
  Object.assign(control, next)
  if (control.type === 'date' && !control.showingPlaceholder && control.dateValue) {
    replaceControlText(model, control, formatContentControlDate(control.dateValue, control.dateFormat, control.dateLocale))
  }
}

export function setDocumentContentControlValue(model: DocumentModel, id: string, value: string): void {
  const control = requireControl(model, id)
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
  if (!['text', 'richText', 'date', 'dropDownList', 'comboBox'].includes(control.type)) {
    throw new TypeError(`Content control ${id} does not accept text values`)
  }
  if (control.type === 'dropDownList' && !control.items?.some((item) => item.value === value)) {
    throw new RangeError(`Value is not present in dropdown content control: ${value}`)
  }
  if (control.type === 'date' && value !== '' && !validIsoDate(value)) throw new RangeError('Date content control value must be YYYY-MM-DD')
  const displayValue = control.type === 'date'
    ? formatContentControlDate(value, control.dateFormat, control.dateLocale)
    : ['dropDownList', 'comboBox'].includes(control.type)
      ? control.items?.find((item) => item.value === value)?.displayText ?? value
      : value
  replaceControlText(model, control, displayValue)
  control.value = value
  if (control.type === 'date') control.dateValue = value
  control.showingPlaceholder = false
}

export function setDocumentContentControlChecked(model: DocumentModel, id: string, checked: boolean): void {
  const control = requireControl(model, id)
  if (control.type !== 'checkbox') throw new TypeError(`Content control ${id} is not a checkbox`)
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  assertAncestorContentEditable(model, control)
  control.checked = checked
  control.value = checked ? 'true' : 'false'
  control.checkbox = { ...control.checkbox, checked }
  control.showingPlaceholder = false
}

export function resetDocumentContentControlPlaceholder(model: DocumentModel, id: string): void {
  const control = requireControl(model, id)
  if (!control.placeholder) throw new Error(`Content control has no placeholder: ${id}`)
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
    replaceControlText(model, control, control.placeholder)
    control.value = control.placeholder
    control.showingPlaceholder = true
    return
  }
  if (!['text', 'richText', 'date', 'dropDownList', 'comboBox'].includes(control.type)) {
    throw new TypeError(`Content control ${id} does not accept text placeholders`)
  }
  replaceControlText(model, control, control.placeholder)
  control.value = control.placeholder
  control.showingPlaceholder = true
}

export function replacePictureContentControlImage(model: DocumentModel, id: string, image: DocumentImage): string {
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
  model.contentControls = controls.filter(candidate => !removedIds.has(candidate.id))
  setTargetRuns(model, target, sliceRuns(runs, 0, startOffset))
  const tail = createParagraph()
  tail.style = { ...block.style }
  tail.type = block.type
  tail.level = block.level
  tail.metadata = { ...block.metadata }
  tail.runs = sliceRuns(runs, endOffset)
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
  setTargetRuns(model, target, [
    ...sliceRuns(runs, 0, insertionOffset),
    ...text,
    ...sliceRuns(runs, insertionOffset),
  ])

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
  setTargetRuns(model, target, [
    ...sliceRuns(runs, 0, startOffset),
    ...sliceRuns(runs, endOffset),
  ])
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
  const control = requireControl(model, id)
  if (control.type !== 'dropDownList' && control.type !== 'comboBox') throw new TypeError(`Content control ${id} is not a list control`)
  const item = control.items?.find((candidate) => candidate.value === value)
  if (!item && control.type === 'dropDownList') throw new RangeError(`Value is not present in dropdown content control: ${value}`)
  setDocumentContentControlValue(model, id, value)
}

export function removeDocumentContentControl(model: DocumentModel, id: string): void {
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
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control paragraph not found: ${control.blockId}`)
  return runText(sliceRuns(runsForTarget(model, target), control.startOffset, control.endOffset))
}

export function contentControlRange(control: DocumentContentControl): { blockId: string; startOffset: number; endOffset: number } {
  return { blockId: control.blockId, startOffset: control.startOffset, endOffset: control.endOffset }
}

function formatContentControlDate(value: string, format?: string, locale?: string): string {
  const date = new Date(`${value}T00:00:00.000Z`)
  if (Number.isNaN(date.valueOf())) return value
  if (!format || format === 'yyyy-MM-dd') return value
  const monthNames = (() => {
    try {
      const formatter = new Intl.DateTimeFormat(locale || 'en-US', { month: 'long', timeZone: 'UTC' })
      return Array.from({ length: 12 }, (_, month) => formatter.format(new Date(Date.UTC(2020, month, 1))))
    } catch {
      return []
    }
  })()
  const year = date.getUTCFullYear()
  const month = date.getUTCMonth()
  const day = date.getUTCDate()
  return format.replace(/yyyy|yy|MMMM|MMM|MM|M|dd|d/g, (token) => {
    switch (token) {
      case 'yyyy': return String(year).padStart(4, '0')
      case 'yy': return String(year).slice(-2)
      case 'MMMM': return monthNames[month] ?? String(month + 1).padStart(2, '0')
      case 'MMM': return monthNames[month]?.slice(0, 3) ?? String(month + 1).padStart(2, '0')
      case 'MM': return String(month + 1).padStart(2, '0')
      case 'M': return String(month + 1)
      case 'dd': return String(day).padStart(2, '0')
      default: return String(day)
    }
  })
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
): void {
  if (!model.contentControls?.length || (startOffset === endOffset && replacementLength === 0 && !split)) return
  // Validate every affected range on a detached plan before changing any live anchors.
  const next: DocumentModel = { ...model, contentControls: getDocumentContentControls(model) }
  for (const control of next.contentControls!) {
    if (!next.contentControls!.some((candidate) => candidate.id === control.id)) continue
    if (control.blockId !== blockId) continue
    const insertion = startOffset === endOffset
    const overlaps = insertion
      ? startOffset > control.startOffset && startOffset < control.endOffset
      : startOffset < control.endOffset && control.startOffset < endOffset
    const removesWholeControl = overlaps && startOffset <= control.startOffset && endOffset >= control.endOffset && replacementLength === 0
    if (removesWholeControl) {
      if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
      removeControlTreeFromModel(next, control)
      continue
    }
    if (overlaps && isContentControlLockedAtEdit(model, control, startOffset, endOffset, insertion)) throw new Error(`Content control content is locked: ${control.id}`)
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
    const controls = next.contentControls!
    const movesToSuffix = (control: DocumentContentControl): boolean => {
      if (control.endOffset <= split.firstBreakOffset) {
        // Empty children at the split boundary follow their enclosing wrapper.
        const parent = control.startOffset === control.endOffset && control.parentId
          ? controls.find((candidate) => candidate.id === control.parentId && candidate.startOffset < candidate.endOffset)
          : undefined
        return parent ? movesToSuffix(parent) : false
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

function replaceControlText(model: DocumentModel, control: DocumentContentControl, value: string): void {
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control paragraph not found: ${control.blockId}`)
  const oldLength = control.endOffset - control.startOffset
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
  const prefix = sliceRuns(runs, 0, control.startOffset)
  const suffix = sliceRuns(runs, control.endOffset)
  const source = runs.find((run) => run.text.length > 0) ?? createRun()
  const { id: _id, text: _text, ...inherited } = source
  const inserted: DocumentRun[] = value ? [createRun(value, inherited)] : []
  setTargetRuns(model, target, compactRuns([...prefix, ...inserted, ...suffix]))
  const delta = value.length - oldLength
  if (delta === 0) return
  for (const candidate of model.contentControls ?? []) {
    if (candidate.id === control.id || candidate.blockId !== control.blockId) continue
    if (candidate.startOffset <= control.startOffset && candidate.endOffset >= control.endOffset
      && candidate.id === control.parentId) {
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
        for (const cell of row.cells) {
          const nested = visit(cell.blocks)
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
        for (const cell of row.cells) {
          if (cell.blocks.some(candidate => candidate.id === target.cellBlockId)) return cell.blocks
          const nested = visit(cell.blocks)
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
        for (const cell of row.cells) {
          const nested = visit(cell.blocks)
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
  if (!Number.isInteger(control.startOffset) || !Number.isInteger(control.endOffset) || control.startOffset < 0 || control.endOffset < control.startOffset) throw new RangeError('Content control range is invalid')
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control target is not addressable: ${control.blockId}`)
  const length = runText(runsForTarget(model, target)).length
  if (control.endOffset > length) throw new RangeError(`Content control range exceeds paragraph length: ${length}`)
  if (!['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox', 'picture', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj'].includes(control.type)) throw new TypeError('Unsupported content control type')
  if (control.lock === undefined) control.lock = 'unlocked'
  if (control.type === 'date' && control.dateValue !== undefined && control.dateValue !== '' && !validIsoDate(control.dateValue)) throw new RangeError('Date content control value must be YYYY-MM-DD')
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
function hasControlStructurePatch(control: DocumentContentControl, patch: ContentControlPatch): boolean {
  return Object.entries(patch).some(([key, value]) => key !== 'lock' && JSON.stringify(Reflect.get(control, key)) !== JSON.stringify(value))
}

function assertAncestorContentEditable(model: DocumentModel, control: DocumentContentControl): void {
  for (const candidate of model.contentControls ?? []) {
    if (candidate.id === control.id || candidate.blockId !== control.blockId) continue
    if (isContentControlLockedAtEdit(model, candidate, control.startOffset, control.endOffset, control.startOffset === control.endOffset)) {
      throw new Error(`Content control content is locked: ${candidate.id}`)
    }
  }
}
function validIsoDate(value: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false
  const date = new Date(`${value}T00:00:00Z`)
  return !Number.isNaN(date.valueOf()) && date.toISOString().slice(0, 10) === value
}
function cloneControl(control: DocumentContentControl): DocumentContentControl {
  return JSON.parse(JSON.stringify(control)) as DocumentContentControl
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
