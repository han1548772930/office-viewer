import type {
  DocumentBlock,
  DocumentContentControl,
  DocumentModel,
} from '../../../domain/model'
import {
  compactRuns,
  createRun,
  runText,
  runsForTarget,
  setTargetRuns,
  sliceRuns,
  tableForBlock,
  textTargets,
  type DocumentRun,
  type TextTarget,
} from '../model/document-model'

export type ContentControlInput = Omit<DocumentContentControl, 'id'> & { id?: string }

export type ContentControlPatch = Partial<Omit<DocumentContentControl, 'id' | 'blockId' | 'startOffset' | 'endOffset' | 'value' | 'checked' | 'dateValue'>>

const lockedContent = new Set<DocumentContentControl['lock']>(['contentLocked', 'sdtContentLocked'])
const lockedControl = new Set<DocumentContentControl['lock']>(['sdtLocked', 'sdtContentLocked'])

export function getDocumentContentControls(model: DocumentModel): DocumentContentControl[] {
  return (model.contentControls ?? []).map((control) => cloneControl(control))
}

export function getDocumentContentControl(model: DocumentModel, id: string): DocumentContentControl | undefined {
  const control = model.contentControls?.find((candidate) => candidate.id === id)
  return control ? cloneControl(control) : undefined
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
  return control
}

export function updateDocumentContentControl(model: DocumentModel, id: string, patch: ContentControlPatch): void {
  const control = requireControl(model, id)
  if (hasControlStructurePatch(patch) && isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${id}`)
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
}

export function setDocumentContentControlValue(model: DocumentModel, id: string, value: string): void {
  const control = requireControl(model, id)
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  if (!['text', 'richText', 'date', 'dropDownList', 'comboBox'].includes(control.type)) {
    throw new TypeError(`Content control ${id} does not accept text values`)
  }
  if (control.type === 'dropDownList' && !control.items?.some((item) => item.value === value)) {
    throw new RangeError(`Value is not present in dropdown content control: ${value}`)
  }
  if (control.type === 'date' && value !== '' && !validIsoDate(value)) throw new RangeError('Date content control value must be YYYY-MM-DD')
  replaceControlText(model, control, value)
  control.value = value
  if (control.type === 'date') control.dateValue = value
  control.showingPlaceholder = false
}

export function setDocumentContentControlChecked(model: DocumentModel, id: string, checked: boolean): void {
  const control = requireControl(model, id)
  if (control.type !== 'checkbox') throw new TypeError(`Content control ${id} is not a checkbox`)
  if (isContentLocked(control)) throw new Error(`Content control content is locked: ${id}`)
  control.checked = checked
  control.value = checked ? 'true' : 'false'
  control.showingPlaceholder = false
}

export function selectDocumentContentControlItem(model: DocumentModel, id: string, value: string): void {
  const control = requireControl(model, id)
  if (control.type !== 'dropDownList' && control.type !== 'comboBox') throw new TypeError(`Content control ${id} is not a list control`)
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
  for (const control of model.contentControls) {
    const segment = segments.find((candidate) => (candidate.cellBlockId ?? candidate.blockId) === control.blockId)
    if (!segment) continue
    const globalStart = segment.startOffset + control.startOffset
    const globalEnd = segment.startOffset + control.endOffset
    const hitsControl = startOffset === endOffset
      ? startOffset > globalStart && startOffset < globalEnd
      : startOffset < globalEnd && globalStart < endOffset
    const editsControlBlock = startSegment && endSegment && startSegment !== endSegment
      && (startSegment.cellBlockId ?? startSegment.blockId) !== control.blockId
      && (endSegment.cellBlockId ?? endSegment.blockId) !== control.blockId
      ? false
      : segment === startSegment || segment === endSegment || (startOffset <= globalStart && endOffset >= globalEnd)
    if (!editsControlBlock) continue
    if (startOffset !== endOffset && startSegment !== endSegment && (hitsControl || (startOffset <= globalStart && endOffset >= globalEnd))) {
      throw new RangeError(`A text edit cannot cross a content control paragraph: ${control.id}`)
    }
    if (startOffset === endOffset) {
      if (startOffset <= globalStart) {
        const delta = replacementLength
        control.startOffset += delta
        control.endOffset += delta
      } else if (startOffset < globalEnd) {
        if (isContentLocked(control)) throw new Error(`Content control content is locked: ${control.id}`)
        control.endOffset += replacementLength
      }
      continue
    }
    if (globalEnd <= startOffset) continue
    if (globalStart >= endOffset) {
      const delta = replacementLength - (endOffset - startOffset)
      control.startOffset += delta
      control.endOffset += delta
      continue
    }
    const localStart = startOffset - segment.startOffset
    const localEnd = endOffset - segment.startOffset
    const removesWholeControl = localStart <= control.startOffset && localEnd >= control.endOffset && replacementLength === 0
    if (removesWholeControl) {
      if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
      removeControlTreeFromModel(model, control)
      continue
    }
    if (isContentControlLockedAtEdit(model, control, localStart, localEnd, startOffset === endOffset)) throw new Error(`Content control content is locked: ${control.id}`)
    const delta = replacementLength - (localEnd - localStart)
    if (localStart <= control.startOffset && localEnd >= control.endOffset) {
      control.startOffset = Math.max(0, localStart)
      control.endOffset = control.startOffset + replacementLength
    } else {
      const oldStart = control.startOffset
      const oldEnd = control.endOffset
      control.startOffset = oldStart >= localEnd ? oldStart + delta : oldStart >= localStart ? localStart : oldStart
      control.endOffset = oldEnd >= localEnd ? oldEnd + delta : oldEnd > localStart ? localStart + replacementLength : oldEnd
      if (control.endOffset < control.startOffset) control.endOffset = control.startOffset
    }
  }
}

/** Local-offset variant for the editor's selection command path. */
export function transformContentControlsForBlockEdit(
  model: DocumentModel,
  blockId: string,
  startOffset: number,
  endOffset: number,
  replacementLength: number,
): void {
  if (!model.contentControls?.length || (startOffset === endOffset && replacementLength === 0)) return
  for (const control of model.contentControls) {
    if (control.blockId !== blockId) continue
    const insertion = startOffset === endOffset
    const overlaps = insertion
      ? startOffset > control.startOffset && startOffset < control.endOffset
      : startOffset < control.endOffset && control.startOffset < endOffset
    const removesWholeControl = overlaps && startOffset <= control.startOffset && endOffset >= control.endOffset && replacementLength === 0
    if (removesWholeControl) {
      if (isControlLocked(control)) throw new Error(`Content control wrapper is locked: ${control.id}`)
      removeControlTreeFromModel(model, control)
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
    if (candidate.startOffset >= control.endOffset) {
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

function validateControl(model: DocumentModel, control: DocumentContentControl): void {
  const block = findBlock(model, control.blockId)
  if (!block || ['table', 'image', 'divider', 'pageBreak'].includes(block.type)) throw new Error(`Content control target is not a text paragraph: ${control.blockId}`)
  if (!Number.isInteger(control.startOffset) || !Number.isInteger(control.endOffset) || control.startOffset < 0 || control.endOffset < control.startOffset) throw new RangeError('Content control range is invalid')
  const target = targetForControl(model, control)
  if (!target) throw new Error(`Content control target is not addressable: ${control.blockId}`)
  const length = runText(runsForTarget(model, target)).length
  if (control.endOffset > length) throw new RangeError(`Content control range exceeds paragraph length: ${length}`)
  if (!['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox', 'picture', 'group', 'repeatingSection', 'repeatingSectionItem', 'docPartObj'].includes(control.type)) throw new TypeError('Unsupported content control type')
  if (control.lock === undefined) control.lock = 'unlocked'
  if (control.type === 'date' && control.dateValue !== undefined && control.dateValue !== '' && !validIsoDate(control.dateValue)) throw new RangeError('Date content control value must be YYYY-MM-DD')
  if ((control.type === 'dropDownList' || control.type === 'comboBox') && (!control.items?.length || control.items.some((item) => !item.value))) throw new RangeError('List content controls require non-empty values')
  if (control.type === 'dropDownList' && control.value !== undefined && !control.items?.some((item) => item.value === control.value)) throw new RangeError(`Dropdown value is not present in its item list: ${control.value}`)
  if (control.type === 'checkbox' && control.checked === undefined) control.checked = false
  if (control.parentId) {
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
    const validParent = (currentContains && next.parentId === current.id && parentTypes.includes(current.type))
      || (nextContains && current.parentId === next.id && parentTypes.includes(next.type))
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
function hasControlStructurePatch(patch: ContentControlPatch): boolean {
  return Object.keys(patch).some((key) => key !== 'lock')
}
function validIsoDate(value: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false
  const date = new Date(`${value}T00:00:00Z`)
  return !Number.isNaN(date.valueOf()) && date.toISOString().slice(0, 10) === value
}
function cloneControl(control: DocumentContentControl): DocumentContentControl {
  return JSON.parse(JSON.stringify(control)) as DocumentContentControl
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
