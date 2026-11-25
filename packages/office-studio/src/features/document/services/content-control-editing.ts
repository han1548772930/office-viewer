import type { DocumentBlock, DocumentContentControl, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentBlockHostId, documentEditingBlockHost, documentStoredBlockHosts, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentNoteDecorationContent, setDocumentNoteDecorationContent } from '../model/document-note-decoration-host'
import { applyRunPatch, documentBlockAncestry, runText, runsForTarget, sameTarget, sliceRuns, textTargets, type DocumentRun, type DocumentSelection } from '../model/document-model'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { isDocumentTextBoundary } from '../model/document-story'
import { projectDocumentStoryBlocks, type DocumentStoryBlockProjection } from '../model/document-story-blocks'
import { contentControlCheckboxGlyph, formatContentControlDate, parseContentControlDateText, validContentControlIsoDate } from './content-control-presentation'
import { documentContentControlContainsSelection, documentContentControlTableRowContentSignature } from '../model/content-control-table-rows'
import { documentContentControlBlockContentSignature, documentContentControlBlocksSelection, documentContentControlBlockTextTargets } from '../model/content-control-block-range'
import { documentContentControlGroupEditAllowed } from '../model/content-control-group-lock'
import type { DocumentNoteReferenceOffset } from '../model/document-note-boundary'
import { compareDocumentNoteBoundaries, remapDocumentNoteBoundaryParagraph } from '../model/document-note-boundary-edit'
import { documentNoteInsertionBoundary, documentNoteReferenceInsertionAffinity } from '../model/document-note-insertion-affinity'
import { assignDocumentContentControlNoteBoundaries, documentContentControlNoteBoundaryRanks } from '../model/content-control-note-boundary'

/** Derived caches, property changes and formatting never use the ordinary input finalizer. */
export const DOCUMENT_CONTENT_CONTROL_TEXT_INPUT_COMMAND_IDS = [
  'doc.command.insert-text', 'doc.command.delete', 'doc.command.paste-blocks',
  'doc.command.table-set-cell-text', 'doc.command.table-set-range-text',
  'doc.command.replace-range',
  'doc.command.insert-latex-formula', 'doc.command.remove-latex-formula', 'doc.command.replace-latex-formula-with-text',
] as const

export function documentContentControlCommandIsTextInput(command: string | { id: string }): boolean {
  const id = typeof command === 'string' ? command : command.id
  return (DOCUMENT_CONTENT_CONTROL_TEXT_INPUT_COMMAND_IDS as readonly string[]).includes(id)
}

/** Whole-cell targets use their table planner, independently of a text control at the UI caret. */
export function documentContentControlCommandUsesWholeCellInput(command: string | { id: string }, selection: DocumentSelection): boolean {
  const id = typeof command === 'string' ? command : command.id
  return id === 'doc.command.table-set-cell-text' || id === 'doc.command.table-set-range-text'
    || !!selection.tableSelection && (id === 'doc.command.insert-text' || id === 'doc.command.delete')
}

export const CONTENT_CONTROL_LAST_KNOWN_DATE_METADATA = 'contentControlLastKnownDate'

export interface DocumentContentControlEditOptions {
  origin: 'text-input' | 'control-value'
  /** The explicit value command's target; its stored list value remains authoritative. */
  controlId?: string
  /** Explicit reset prompt, including a valid empty fallback; it is not a saved date/list value. */
  placeholderText?: string
  /** Prepared source selection, before text/range mutation. Required for empty-range insertion. */
  selection?: DocumentSelection
  /** Explicitly identifies insertion inside a nonempty control at its start/end boundary. */
  activeControlId?: string
}

export interface DocumentContentControlEditResult {
  changedControlIds: string[]
  /** Also includes temporary wrappers removed by the ordinary structural range transformer. */
  removedTemporaryControlIds: string[]
  retainedLockedTemporaryControlIds: string[]
  /** Final scalar/range state, retained only by the caller's transaction while binding peers are prepared. */
  removedTemporaryControls?: DocumentContentControl[]
}

interface ControlParagraph {
  block: DocumentBlock
  text: string
  scopeId: string
  blocks: DocumentBlock[]
  noteReferences: readonly DocumentNoteReferenceOffset[]
}

function paragraphs(model: DocumentModel): Map<string, ControlParagraph> {
  const result = new Map<string, ControlParagraph>()
  const editingHost = documentEditingBlockHost(model)
  const scopes = [{ blocks: model.blocks, scopeId: editingHost ? documentBlockHostId(editingHost) : 'body' },
    ...documentStoredBlockHosts(model).map(entry => ({ blocks: entry.blocks, scopeId: documentBlockHostId(entry.host) }))]
  for (const scope of scopes) {
    for (const paragraph of documentTextFieldParagraphs({ blocks: scope.blocks })) {
      if (result.has(paragraph.block.id)) throw new RangeError('内容控件段落 ID 在不同内容区重复：' + paragraph.block.id)
      result.set(paragraph.block.id, { block: paragraph.block, text: paragraph.text, scopeId: scope.scopeId, blocks: scope.blocks, noteReferences: paragraph.noteReferences })
    }
  }
  return result
}

function rangeText(paragraph: ControlParagraph | undefined, control: DocumentContentControl): string {
  if (!paragraph) throw new RangeError('内容控件段落不存在：' + control.id)
  const { startOffset: start, endOffset: end } = control
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > paragraph.text.length
    || !isDocumentTextBoundary(paragraph.text, start) || !isDocumentTextBoundary(paragraph.text, end)) {
    throw new RangeError('内容控件输入范围不是有效的 UTF-16 区间：' + control.id)
  }
  documentContentControlNoteBoundaryRanks(control, paragraph.noteReferences)
  return paragraph.text.slice(start, end)
}

/** Empty reference entities are content even when the displayed string is unchanged. */
function rangeNoteSignature(paragraph: ControlParagraph, control: DocumentContentControl): string {
  const range = documentContentControlNoteBoundaryRanks(control, paragraph.noteReferences)
  const runs = new Map(paragraph.block.runs.map(run => [run.id, run]))
  return JSON.stringify(paragraph.noteReferences.slice(range.start, range.end).map(reference => {
    const run = runs.get(reference.runId)
    if (!run) throw new RangeError('内容控件引用身份不属于真实段落。')
    return { runId: run.id, noteReference: run.metadata?.noteReference, note: run.metadata?.note }
  }))
}

function wrapperLocked(control: DocumentContentControl): boolean {
  return ['controlLocked', 'sdtLocked', 'sdtContentLocked'].includes(control.lock ?? '')
}

function assertContentEditable(control: DocumentContentControl, controls: Map<string, DocumentContentControl>, before: DocumentModel, draft: DocumentModel,
  source: Map<string, ControlParagraph>, destination: Map<string, ControlParagraph>, options: DocumentContentControlEditOptions): void {
  const seen = new Set<string>()
  let current: DocumentContentControl | undefined = control
  while (current) {
    if (seen.has(current.id)) throw new RangeError('内容控件父级关系循环：' + current.id)
    seen.add(current.id)
    if (current.lock === 'contentLocked' || current.lock === 'sdtContentLocked') {
      const from = source.get(current.blockId), mapped = draft.contentControls?.find(control => control.id === current!.id)
      const to = destination.get(mapped?.blockId ?? current.blockId)
      if (!from || !to || from.scopeId !== to.scopeId || !documentContentControlGroupEditAllowed(
        { ...before, blocks: from.blocks }, { ...draft, blocks: to.blocks }, current,
        { selection: options.selection, activeControlId: options.activeControlId ?? options.controlId })) {
        throw new Error('内容控件内容已锁定：' + current.id)
      }
    }
    current = current.parentId ? controls.get(current.parentId) : undefined
  }
}

function depth(control: DocumentContentControl, controls: Map<string, DocumentContentControl>): number {
  let value = 0, parent = control.parentId
  const seen = new Set([control.id])
  while (parent && !seen.has(parent)) {
    seen.add(parent); value++
    parent = controls.get(parent)?.parentId
  }
  return value
}

/** An explicitly selected control supplies boundary affinity; a newly placed caret does not. */
export function resolveDocumentContentControlTextInputId(
  model: DocumentModel, selection: DocumentSelection, activeControlId?: string,
): string | undefined {
  if (activeControlId !== undefined && (typeof activeControlId !== 'string' || !activeControlId.trim())) throw new TypeError('内容控件输入标识无效。')
  const controls = model.contentControls ?? [], index = new Map(controls.map(control => [control.id, control]))
  const sameParagraph = !selection.tableSelection && sameTarget(selection.anchor, selection.focus)
  const start = Math.min(selection.anchor.offset, selection.focus.offset), end = Math.max(selection.anchor.offset, selection.focus.offset)
  if (sameParagraph) {
    const target = textTargets(model).find(target => sameTarget(target, selection.focus)), text = target && runText(runsForTarget(model, target))
    if (!target || text === undefined || !isDocumentTextBoundary(text, start) || !isDocumentTextBoundary(text, end)) {
      if (activeControlId !== undefined) throw new RangeError('内容控件输入身份与当前文字选区不一致。')
      return undefined
    }
  }
  const contains = (control: DocumentContentControl) => !control.metadata?.officeStoryLockSource
    && documentContentControlContainsSelection(model, control, selection)
  if (activeControlId !== undefined) {
    const active = index.get(activeControlId)
    if (!active || !contains(active)) throw new RangeError('内容控件输入身份与当前文字选区不一致。')
    return active.id
  }
  return controls.filter(control => contains(control) && (control.blockRange || control.tableRowRange || !sameParagraph || start !== end || control.showingPlaceholder
    || control.startOffset === control.endOffset || start > control.startOffset && start < control.endOffset))
    .sort((left, right) => depth(right, index) - depth(left, index)
      || (left.blockRange || left.tableRowRange ? 1 : 0) - (right.blockRange || right.tableRowRange ? 1 : 0)
      || left.endOffset - left.startOffset - (right.endOffset - right.startOffset))[0]?.id
}

/** A placeholder edit replaces its prompt once; this only prepares a logical source selection. */
export function prepareDocumentContentControlTextInputSelection(
  model: DocumentModel, selection: DocumentSelection, options: { controlId?: string } = {},
): DocumentSelection {
  const result = cloneModel(selection)
  if (selection.tableSelection) return result
  const structural = (model.contentControls ?? []).find(control => control.blockRange && control.showingPlaceholder
    && (!options.controlId || control.id === options.controlId) && documentContentControlContainsSelection(model, control, selection))
  if (structural?.blockRange) return documentContentControlBlocksSelection(model, structural.blockRange)
  if (!sameTarget(selection.anchor, selection.focus)) return result
  const blockId = selection.focus.cellBlockId ?? selection.focus.blockId
  const from = Math.min(selection.anchor.offset, selection.focus.offset), to = Math.max(selection.anchor.offset, selection.focus.offset)
  const controls = model.contentControls ?? [], index = new Map(controls.map(control => [control.id, control]))
  const control = controls.filter(control => control.blockId === blockId && control.showingPlaceholder
    && (!options.controlId || control.id === options.controlId) && from >= control.startOffset && to <= control.endOffset
    && !control.metadata?.officeStoryLockSource)
    .sort((left, right) => left.endOffset - left.startOffset - (right.endOffset - right.startOffset) || depth(right, index) - depth(left, index))[0]
  if (!control) return result
  const reverse = selection.anchor.offset > selection.focus.offset
  result.anchor.offset = reverse ? control.endOffset : control.startOffset
  result.focus.offset = reverse ? control.startOffset : control.endOffset
  return result
}

/**
 * Half-open source anchors treat ordinary boundary insertion as outside. An empty selected
 * range, or an explicit inside-control identity, supplies the missing input affinity.
 */
function repairInsideInsertion(
  before: DocumentModel, controls: DocumentContentControl[], source: Map<string, ControlParagraph>, destination: Map<string, ControlParagraph>,
  options: DocumentContentControlEditOptions,
): void {
  const selection = options.selection
  if (options.origin !== 'text-input' || !selection || selection.tableSelection
    || !sameTarget(selection.anchor, selection.focus) || selection.anchor.offset !== selection.focus.offset) return
  const blockId = selection.focus.cellBlockId ?? selection.focus.blockId, offset = selection.focus.offset
  const original = before.contentControls ?? [], index = new Map(original.map(control => [control.id, control]))
  const from = source.get(blockId), to = destination.get(blockId)
  if (!from || !to || from.scopeId !== to.scopeId) return
  const added = to.text.length - from.text.length
  if (added < 0 || to.text.slice(0, offset) !== from.text.slice(0, offset)
    || to.text.slice(offset + added) !== from.text.slice(offset)) return
  const target = { id: blockId, text: to.text, noteReferences: to.noteReferences }
  const affinity = documentNoteReferenceInsertionAffinity(from.noteReferences, [target], offset, added)
  if (!affinity) return
  const selected = original.filter(control => {
    if (control.blockRange || control.tableRowRange || control.blockId !== blockId || control.metadata?.officeStoryLockSource) return false
    const range = documentContentControlNoteBoundaryRanks(control, from.noteReferences)
    return options.activeControlId ? control.id === options.activeControlId
      && compareDocumentNoteBoundaries(control.startOffset, range.start, offset, affinity.sourceRank) <= 0
      && compareDocumentNoteBoundaries(offset, affinity.sourceRank, control.endOffset, range.end) <= 0
      : control.startOffset === offset && control.endOffset === offset && range.start === affinity.sourceRank && range.end === affinity.sourceRank
  }).sort((left, right) => depth(right, index) - depth(left, index))[0]
  if (!selected) return
  let current: DocumentContentControl | undefined = selected
  const seen = new Set<string>()
  while (current && !seen.has(current.id)) {
    seen.add(current.id)
    const mapped = controls.find(control => control.id === current!.id)
    if (!mapped) break
    if (!mapped.blockRange && !mapped.tableRowRange) {
      const range = documentContentControlNoteBoundaryRanks(current, from.noteReferences)
      if (mapped.blockId !== blockId || compareDocumentNoteBoundaries(current.startOffset, range.start, offset, affinity.sourceRank) > 0
        || compareDocumentNoteBoundaries(offset, affinity.sourceRank, current.endOffset, range.end) > 0) break
      let start = remapDocumentNoteBoundaryParagraph(from.noteReferences, current.startOffset, current.startNoteReferenceCount, [target],
        { paragraph: target, offset: current.startOffset })
      let end = remapDocumentNoteBoundaryParagraph(from.noteReferences, current.endOffset, current.endNoteReferenceCount, [target],
        { paragraph: target, offset: current.endOffset + added })
      const point = compareDocumentNoteBoundaries(current.startOffset, range.start, current.endOffset, range.end) === 0
      start = documentNoteInsertionBoundary(affinity, current.startOffset, range.start, 'start', point, true) ?? start
      end = documentNoteInsertionBoundary(affinity, current.endOffset, range.end, 'end', point, true) ?? end
      assignDocumentContentControlNoteBoundaries(mapped, start, end)
      documentContentControlNoteBoundaryRanks(mapped, to.noteReferences)
    }
    current = current.parentId ? index.get(current.parentId) : undefined
  }
}

/** Remove only these wrappers. Content, child identities and unknown child metadata remain intact. */
export function unwrapDocumentContentControlShells(controls: readonly DocumentContentControl[], ids: ReadonlySet<string>): DocumentContentControl[] {
  const index = new Map(controls.map(control => [control.id, control]))
  return controls.filter(control => !ids.has(control.id)).map(control => {
    const next = cloneModel(control)
    let parent = next.parentId
    const seen = new Set([next.id])
    while (parent && ids.has(parent)) {
      if (seen.has(parent)) throw new RangeError('临时控件卸壳遇到循环父级：' + parent)
      seen.add(parent); parent = index.get(parent)?.parentId
    }
    if (parent) next.parentId = parent
    else delete next.parentId
    return next
  })
}

function explicitValueChanged(before: DocumentContentControl, after: DocumentContentControl, options: DocumentContentControlEditOptions): boolean {
  if (options.origin !== 'control-value' || options.controlId !== after.id) return false
  if (before.showingPlaceholder && !after.showingPlaceholder) return true
  if (options.placeholderText !== undefined && !!before.showingPlaceholder !== !!after.showingPlaceholder) return true
  return before.value !== after.value || before.dateValue !== after.dateValue
    || (before.checked ?? before.checkbox?.checked) !== (after.checked ?? after.checkbox?.checked)
}

function checkboxFontRuns(runs: DocumentRun[], start: number, end: number, fontFamily: string): DocumentRun[] {
  let offset = 0
  return runs.flatMap(run => {
    const from = Math.max(start, offset), to = Math.min(end, offset + run.text.length), runStart = offset
    offset += run.text.length
    if (from >= to || run.fontFamily === fontFamily) return [cloneModel(run)]
    return [
      ...sliceRuns([run], 0, from - runStart),
      ...sliceRuns([run], from - runStart, to - runStart).map(part => applyRunPatch(part, { fontFamily })),
      ...sliceRuns([run], to - runStart),
    ]
  })
}

interface ControlHostFontPlan {
  host: DocumentAuxiliaryBlockHost
  blocks: DocumentBlock[]
  projection: DocumentStoryBlockProjection
}

/** Prepare compatibility projections before any font/control plans are published to the draft. */
function hostFontPlans(model: DocumentModel, destinations: Map<string, ControlParagraph>, plans: Map<DocumentBlock, DocumentRun[]>): ControlHostFontPlan[] {
  if (!plans.size) return []
  const byId = new Map<string, DocumentRun[]>([...plans].map(([block, runs]) => [block.id, runs] as const))
  const scopes = new Set([...byId.keys()].map(id => destinations.get(id)!.scopeId))
  return documentStoredBlockHosts(model).filter(entry => scopes.has(documentBlockHostId(entry.host))).map(entry => {
    const host = entry.host
    const content = host.kind === 'story' ? model.sectionHeaderFooters?.[host.sectionId]?.storyContent?.[host.story]
      : host.kind === 'note' ? (host.noteKind === 'footnote' ? model.footnotes : model.endnotes)?.[host.noteId]?.content
        : documentNoteDecorationContent(model, host)
    if (!content?.blocks) throw new RangeError('内容控件字体的内容区块树不存在：' + documentBlockHostId(host))
    const staged = cloneModel(content)
    for (const paragraph of documentTextFieldParagraphs({ blocks: staged.blocks! })) {
      const runs = byId.get(paragraph.block.id)
      if (!runs) continue
      paragraph.block.runs = cloneModel(runs)
      for (const ancestor of documentBlockAncestry({ blocks: staged.blocks! }, paragraph.block.id) ?? []) delete ancestor.metadata.html
    }
    return { host, blocks: content.blocks, projection: projectDocumentStoryBlocks(staged, host.kind === 'story' ? 'literal' : 'display') }
  })
}

function publishHostFontPlans(model: DocumentModel, plans: readonly ControlHostFontPlan[]): void {
  for (const { host, blocks, projection } of plans) {
    const content = { ...projection.content, blocks }
    if (host.kind === 'story') {
      const section = model.sectionHeaderFooters![host.sectionId]!
      section[host.story] = projection.text
      section.storyContent = { ...section.storyContent, [host.story]: content }
      section.pageFields = [...(section.pageFields ?? []).filter(field => field.story !== host.story),
        ...projection.pageFields.map(field => ({ ...field, story: host.story }))]
    } else if (host.kind === 'note') {
      const note = (host.noteKind === 'footnote' ? model.footnotes : model.endnotes)![host.noteId]!
      note.content = content; note.text = projection.text
    } else setDocumentNoteDecorationContent(model, host, content)
  }
}

/**
 * Finalize only a caller-owned transaction candidate. All changed controls are checked before
 * controls or run/font plans are published; rejection never mutates the source model.
 */
export function synchronizeDocumentContentControlsAfterEdit(
  before: DocumentModel, draft: DocumentModel, options: DocumentContentControlEditOptions,
): DocumentContentControlEditResult {
  if (options.placeholderText !== undefined && (typeof options.placeholderText !== 'string' || options.origin !== 'control-value' || !options.controlId)) {
    throw new TypeError('占位恢复必须提供显式控件 ID 和文本候选。')
  }
  const result: DocumentContentControlEditResult = { changedControlIds: [], removedTemporaryControlIds: [], retainedLockedTemporaryControlIds: [] }
  if (!before.contentControls?.length) return result
  const source = paragraphs(before), destination = paragraphs(draft)
  const retained = new Set((draft.contentControls ?? []).map(control => control.id))
  const original = new Map(before.contentControls.map(control => [control.id, control]))
  for (const control of before.contentControls) {
    if (!control.temporary || retained.has(control.id) || control.metadata?.officeStoryLockSource) continue
    if (wrapperLocked(control)) throw new Error('正文输入不能删除已锁定的临时控件外壳：' + control.id)
    assertContentEditable(control, original, before, draft, source, destination, options)
    result.changedControlIds.push(control.id)
    result.removedTemporaryControlIds.push(control.id)
  }
  if (!draft.contentControls?.length) return result
  const planned = cloneModel(draft.contentControls)
  repairInsideInsertion(before, planned, source, destination, options)
  const remove = new Set<string>(), fontPlans = new Map<DocumentBlock, DocumentRun[]>()
  for (const control of planned) {
    const previous = original.get(control.id)
    if (!previous || control.metadata?.officeStoryLockSource) continue
    if (previous.type !== control.type) throw new RangeError('正文输入不能改变内容控件类型：' + control.id)
    const from = source.get(previous.blockId), to = destination.get(control.blockId)
    if (!from && !to && previous.blockId === control.blockId && previous.startOffset === control.startOffset && previous.endOffset === control.endOffset
      && previous.startNoteReferenceCount === control.startNoteReferenceCount && previous.endNoteReferenceCount === control.endNoteReferenceCount) continue
    if (!from || !to || from.scopeId !== to.scopeId) throw new RangeError('正文输入不能把内容控件移到其他内容区：' + control.id)
    if (previous.tableRowRange || control.tableRowRange) {
      if (!previous.tableRowRange || !control.tableRowRange) throw new RangeError('正文输入不能改变内容控件的持久行范围类型：' + control.id)
      rangeText(from, previous); rangeText(to, control)
      const oldContent = documentContentControlTableRowContentSignature({ ...before, blocks: from.blocks }, previous.tableRowRange)
      const newContent = documentContentControlTableRowContentSignature({ ...draft, blocks: to.blocks }, control.tableRowRange)
      if (oldContent === newContent) continue
      assertContentEditable(previous, original, before, draft, source, destination, options)
      control.showingPlaceholder = false
      if (control.dataBinding) control.dataBinding.dirty = true
      result.changedControlIds.push(control.id)
      if (control.temporary) {
        if (wrapperLocked(control)) result.retainedLockedTemporaryControlIds.push(control.id)
        else { remove.add(control.id); result.removedTemporaryControlIds.push(control.id) }
      }
      continue
    }
    let oldText: string, text: string, structureChanged = false, noteEntitiesChanged = false
    if (previous.blockRange || control.blockRange) {
      if (!previous.blockRange || !control.blockRange) throw new RangeError('正文输入不能改变完整块内容控件的范围类型：' + control.id)
      rangeText(from, previous); rangeText(to, control)
      const fromModel = { ...before, blocks: from.blocks }, toModel = { ...draft, blocks: to.blocks }
      structureChanged = documentContentControlBlockContentSignature(fromModel, previous.blockRange) !== documentContentControlBlockContentSignature(toModel, control.blockRange)
      oldText = documentContentControlBlockTextTargets(fromModel, previous.blockRange).map(target => runText(runsForTarget(fromModel, target))).join('\n')
      text = documentContentControlBlockTextTargets(toModel, control.blockRange).map(target => runText(runsForTarget(toModel, target))).join('\n')
    } else {
      oldText = rangeText(from, previous); text = rangeText(to, control)
      noteEntitiesChanged = rangeNoteSignature(from, previous) !== rangeNoteSignature(to, control)
    }
    if (!structureChanged && !noteEntitiesChanged && oldText === text && !explicitValueChanged(previous, control, options)) continue
    assertContentEditable(previous, original, before, draft, source, destination, options)
    const explicit = options.origin === 'control-value' && options.controlId === control.id
    const placeholder = explicit && options.placeholderText !== undefined
    if (placeholder) {
      if (control.showingPlaceholder !== true || control.value !== options.placeholderText || text !== options.placeholderText) {
        throw new RangeError('内容控件占位标记、提示值与正文不一致：' + control.id)
      }
      if (!['text', 'richText', 'date', 'dropDownList', 'comboBox', 'picture'].includes(control.type)) {
        throw new TypeError('此内容控件不支持占位恢复：' + control.id)
      }
      if (control.type === 'date') {
        if (previous.dateValue && validContentControlIsoDate(previous.dateValue)) control.metadata = { ...control.metadata, [CONTENT_CONTROL_LAST_KNOWN_DATE_METADATA]: previous.dateValue }
        delete control.dateValue
      }
    } else if (control.type === 'dropDownList') {
      if (!explicit) throw new Error('下拉列表内容不能直接输入，请选择列表项：' + control.id)
      const item = control.items?.find(item => item.value === control.value)
      if (!item || text !== (item.displayText ?? item.value)) throw new RangeError('下拉列表保存值与显示内容不一致：' + control.id)
    } else if (control.type === 'checkbox') {
      const checked = contentControlCheckboxGlyph(control, true), unchecked = contentControlCheckboxGlyph(control, false)
      let state: boolean
      if (explicit) {
        state = control.checked ?? control.checkbox?.checked ?? false
        if (text !== (state ? checked.text : unchecked.text)) throw new RangeError('复选框正文与选中状态不一致：' + control.id)
      } else if (text === checked.text && text === unchecked.text) state = previous.checked ?? previous.checkbox?.checked ?? false
      else if (text === checked.text) state = true
      else if (text === unchecked.text) state = false
      else throw new RangeError('复选框只能包含配置的一个选中或未选中符号：' + control.id)
      control.checked = state; control.value = state ? 'true' : 'false'
      control.checkbox = { ...control.checkbox, checked: state }
      const glyph = state ? checked : unchecked
      const currentRuns = fontPlans.get(to.block) ?? to.block.runs
      const nextRuns = checkboxFontRuns(currentRuns, control.startOffset, control.endOffset, glyph.fontFamily)
      if (JSON.stringify(nextRuns) !== JSON.stringify(currentRuns)) fontPlans.set(to.block, nextRuns)
    } else if (control.type === 'date') {
      const parsed = explicit ? control.dateValue : parseContentControlDateText(text, control.dateFormat, control.dateLocale, control.dateCalendar)
      if (explicit && parsed === undefined) throw new RangeError('日期值命令必须提供明确的 ISO 日期或空值：' + control.id)
      if (parsed !== undefined) {
        if (parsed !== '' && !validContentControlIsoDate(parsed)) throw new RangeError('日期内容控件保存值不是严格 ISO 日期：' + control.id)
        if (explicit && text !== formatContentControlDate(parsed, control.dateFormat, control.dateLocale, control.dateCalendar)) {
          throw new RangeError('日期控件的保存值与显示内容不一致：' + control.id)
        }
        control.dateValue = parsed; control.value = parsed
        const knownDate = parsed || (previous.dateValue && validContentControlIsoDate(previous.dateValue) ? previous.dateValue : undefined)
        if (knownDate) control.metadata = { ...control.metadata, [CONTENT_CONTROL_LAST_KNOWN_DATE_METADATA]: knownDate }
      } else {
        control.value = text
        if (previous.dateValue && validContentControlIsoDate(previous.dateValue)) control.metadata = { ...control.metadata, [CONTENT_CONTROL_LAST_KNOWN_DATE_METADATA]: previous.dateValue }
        delete control.dateValue
      }
    } else if (control.type === 'text' || control.type === 'richText') {
      if (control.type === 'text' && control.multiLine !== true && /[\r\n]/.test(text)) throw new RangeError('此纯文本控件未启用多行输入：' + control.id)
      control.value = text.replace(/\r\n?/g, '\n')
    } else if (control.type === 'comboBox') {
      if (/[\r\n]/.test(text)) throw new RangeError('组合框内容不能包含段落换行：' + control.id)
      if (!explicit) control.value = text
      else {
        const display = control.items?.find(item => item.value === control.value)?.displayText ?? control.value
        if (display !== text) throw new RangeError('组合框保存值与显示内容不一致：' + control.id)
      }
    } else if (control.type === 'picture' && !explicit) throw new Error('图片内容控件请使用图片替换命令：' + control.id)
    control.showingPlaceholder = placeholder
    if (control.dataBinding) control.dataBinding.dirty = true
    result.changedControlIds.push(control.id)
    if (control.temporary) {
      if (wrapperLocked(control)) result.retainedLockedTemporaryControlIds.push(control.id)
      else { remove.add(control.id); result.removedTemporaryControlIds.push(control.id) }
    }
  }
  if (!result.changedControlIds.length) return result
  if (remove.size) result.removedTemporaryControls = planned.filter(control => remove.has(control.id)).map(control => cloneModel(control))
  const next = unwrapDocumentContentControlShells(planned, remove)
  const projections = hostFontPlans(draft, destination, fontPlans)
  draft.contentControls = next
  for (const [block, runs] of fontPlans) {
    block.runs = runs
    const paragraph = destination.get(block.id)!
    for (const ancestor of documentBlockAncestry({ blocks: paragraph.blocks }, block.id) ?? []) delete ancestor.metadata.html
  }
  publishHostFontPlans(draft, projections)
  return result
}
