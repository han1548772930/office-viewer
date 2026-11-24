import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentBlockHostForId, documentEditingBlockHost, sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { runText, runsForTarget, sameTarget, textTargets, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import { cloneDocumentSelection } from '../model/document-table-selection'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { assertDocumentPermission, permissionTargetsForBlock, type DocumentPermissionContext } from '../permissions/document-permissions'
import type { DocumentContentControlEditResult } from './content-control-editing'
import { contentControlCheckboxGlyph, formatContentControlDate, parseContentControlDateText } from './content-control-presentation'
import { assertDocumentContentControlValueEditable, contentControlText, documentContentControlPlaceholderText,
  resetDocumentContentControlPlaceholder, setDocumentContentControlChecked, setDocumentContentControlValue } from './content-controls'
import { documentModelForBlockHost } from './document-block-host'

const scalarTypes = new Set<DocumentContentControl['type']>(['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox'])

export interface DocumentContentControlBindingTextEdit {
  controlId: string
  blockId: string
  startOffset: number
  endOffset: number
  text: string
  host?: DocumentAuxiliaryBlockHost
}

/** Exact UTF-16 ranges in the flattened source host, captured by the originating text planner. */
export interface DocumentContentControlBindingOriginTextEdit {
  startOffset: number
  endOffset: number
  text: string
}

export interface DocumentContentControlBindingOptions {
  /** The originating mutation's result coordinates, before linked controls are updated. */
  selection?: DocumentSelection
  host?: DocumentAuxiliaryBlockHost
  /** Prepared source coordinates allow a cross-paragraph temporary shell to retain its scalar. */
  sourceSelection?: DocumentSelection
  /** Whole-cell and replace-all commands address actual ranges independently of the UI caret. */
  originTextEdits?: readonly DocumentContentControlBindingOriginTextEdit[]
  activeControlId?: string
  editResult?: DocumentContentControlEditResult
  /** Explicit value/reset commands can supply the scalar after their temporary source has disappeared. */
  sourceControlId?: string
  sourceValue?: string
  permissionContext?: DocumentPermissionContext
}

export interface DocumentContentControlBindingResult {
  changedControlIds: string[]
  textEdits: DocumentContentControlBindingTextEdit[]
  selection?: DocumentSelection
}

/** Fixed Univer 1.6.0 key: only the custom-XML store ID is case insensitive. */
export function documentContentControlBindingKey(control: DocumentContentControl): string | undefined {
  const binding = control.dataBinding
  if (!binding?.storeItemID || !binding.xpath) return undefined
  return JSON.stringify([binding.storeItemID.toLowerCase(), binding.xpath, binding.prefixMappings ?? ''])
}

/** Values are independent of localized labels, date masks, and checkbox glyphs. */
export function documentContentControlBindingValue(model: DocumentModel, control: DocumentContentControl): string | undefined {
  if (!scalarTypes.has(control.type) || control.metadata?.officeStoryLockSource) return undefined
  if (control.showingPlaceholder) return ''
  if (control.type === 'checkbox') return (control.checked ?? control.checkbox?.checked) ? 'true' : 'false'
  const text = contentControlText(model, control)
  if (control.type === 'date') return control.dateValue ?? text
  if (control.type === 'dropDownList' || control.type === 'comboBox') return control.value ?? text
  return text
}

function scopedModel(model: DocumentModel, host?: DocumentAuxiliaryBlockHost): DocumentModel {
  const editingHost = documentEditingBlockHost(model)
  return !host || sameDocumentBlockHost(editingHost, host) ? model : documentModelForBlockHost(model, host)
}

function hostForControl(model: DocumentModel, control: DocumentContentControl): DocumentAuxiliaryBlockHost | undefined {
  return documentBlockHostForId(model, control.blockId) ?? documentEditingBlockHost(model)
}

function absolutePoint(model: DocumentModel, point: DocumentPoint): number | undefined {
  const paragraph = documentTextFieldParagraphs(model).find(paragraph => paragraph.block.id === (point.cellBlockId ?? point.blockId))
  return paragraph && isDocumentTextBoundary(paragraph.text, point.offset) ? paragraph.absoluteStart + point.offset : undefined
}

interface OriginTextEdit { start: number; end: number; text: string; source: string; destination: string }

function plannedOriginTextEdits(before: DocumentModel, after: DocumentModel,
  options: DocumentContentControlBindingOptions): OriginTextEdit[] | undefined {
  if (!options.originTextEdits) {
    const edit = originTextEdit(before, after, options)
    return edit ? [edit] : undefined
  }
  const source = documentTextFieldParagraphs(scopedModel(before, options.host)).map(paragraph => paragraph.text).join('\n')
  const destination = documentTextFieldParagraphs(scopedModel(after, options.host)).map(paragraph => paragraph.text).join('\n')
  const edits = options.originTextEdits.map(edit => {
    if (!Number.isSafeInteger(edit.startOffset) || !Number.isSafeInteger(edit.endOffset) || edit.startOffset < 0
      || edit.endOffset < edit.startOffset || typeof edit.text !== 'string'
      || !isDocumentTextBoundary(source, edit.startOffset) || !isDocumentTextBoundary(source, edit.endOffset)) {
      throw new RangeError('绑定源编辑范围不是有效的 UTF-16 区间。')
    }
    return { start: edit.startOffset, end: edit.endOffset, text: edit.text.replace(/\r\n?/g, '\n'), source, destination }
  }).sort((left, right) => right.start - left.start || right.end - left.end)
  let expected = source, followingStart = source.length + 1
  for (const edit of edits) {
    if (edit.end > followingStart || edit.start === followingStart) throw new RangeError('绑定源编辑范围重叠。')
    expected = expected.slice(0, edit.start) + edit.text + expected.slice(edit.end)
    followingStart = edit.start
  }
  if (expected !== destination) throw new RangeError('绑定源编辑计划与实际候选文字不一致。')
  return edits
}

/** Use actual source/caret positions, avoiding an ambiguous longest-prefix diff for repeated text. */
function originTextEdit(before: DocumentModel, after: DocumentModel, options: DocumentContentControlBindingOptions): OriginTextEdit | undefined {
  const sourceSelection = options.sourceSelection, selection = options.selection
  if (!sourceSelection || !selection || sourceSelection.tableSelection || selection.tableSelection
    || !sameTarget(selection.anchor, selection.focus) || selection.anchor.offset !== selection.focus.offset) return undefined
  const source = scopedModel(before, options.host), destination = scopedModel(after, options.host)
  const a = absolutePoint(source, sourceSelection.anchor), b = absolutePoint(source, sourceSelection.focus)
  const caret = absolutePoint(destination, selection.focus)
  if (a === undefined || b === undefined || caret === undefined) return undefined
  const fromText = documentTextFieldParagraphs(source).map(paragraph => paragraph.text).join('\n')
  const toText = documentTextFieldParagraphs(destination).map(paragraph => paragraph.text).join('\n')
  let start = Math.min(a, b), end = Math.max(a, b)
  if (start === end && toText.length < fromText.length) {
    if (caret < start) { end = start; start = caret }
    else end = start + fromText.length - toText.length
  }
  if (caret < start || fromText.slice(0, start) !== toText.slice(0, start) || fromText.slice(end) !== toText.slice(caret)) return undefined
  return { start, end, text: toText.slice(start, caret), source: fromText, destination: toText }
}

function inActiveChain(before: DocumentModel, control: DocumentContentControl, activeId?: string): boolean {
  const controls = new Map((before.contentControls ?? []).map(control => [control.id, control]))
  let active = activeId ? controls.get(activeId) : undefined
  const seen = new Set<string>()
  while (active && !seen.has(active.id)) {
    if (active.id === control.id) return true
    seen.add(active.id); active = active.parentId ? controls.get(active.parentId) : undefined
  }
  return false
}

function removedSourceValue(before: DocumentModel, after: DocumentModel, control: DocumentContentControl,
  options: DocumentContentControlBindingOptions, originEdits: () => OriginTextEdit[] | undefined): string | undefined {
  const captured = options.editResult?.removedTemporaryControls?.find(candidate => candidate.id === control.id)
  if (captured) return documentContentControlBindingValue(after, captured)
  if (options.sourceControlId === control.id && options.sourceValue !== undefined) return control.type === 'text' || control.type === 'richText'
    ? options.sourceValue.replace(/\r\n?/g, '\n') : options.sourceValue
  const active = inActiveChain(before, control, options.activeControlId)
  if (!active && !options.originTextEdits) return undefined
  const edits = originEdits()
  const paragraph = documentTextFieldParagraphs(scopedModel(before, options.host)).find(paragraph => paragraph.block.id === control.blockId)
  if (!paragraph) {
    if (active) throw new RangeError('已卸壳绑定控件的源内容区不存在：' + control.id)
    return undefined
  }
  if (!edits) throw new RangeError('无法保留已卸壳绑定控件的实际输入值：' + control.id)
  let start = paragraph.absoluteStart + control.startOffset, end = paragraph.absoluteStart + control.endOffset, changed = false
  for (const edit of edits) {
    const insertion = edit.start === edit.end, delta = edit.text.length - (edit.end - edit.start)
    const insideInsertion = insertion && (edit.start > start && edit.start < end
      || edit.start === start && start === end || active && edit.start >= start && edit.start <= end)
    const overlaps = insertion ? insideInsertion : edit.start < end && start < edit.end
    if (overlaps && edit.source.slice(edit.start, edit.end) !== edit.text) changed = true
    if (insideInsertion) end += edit.text.length
    else if (insertion) {
      if (edit.start <= start) { start += delta; end += delta }
      else if (edit.start < end) end += delta
    } else if (edit.end <= start) { start += delta; end += delta }
    else if (edit.start < end && start < edit.end) {
      start = start >= edit.start ? edit.start : start
      end = end >= edit.end ? end + delta : edit.start + edit.text.length
      if (end < start) end = start
    }
  }
  if (!changed) return undefined
  const destination = edits[0]?.destination ?? ''
  if (!isDocumentTextBoundary(destination, start) || !isDocumentTextBoundary(destination, end)) {
    throw new RangeError('已卸壳绑定控件无法映射到候选文字：' + control.id)
  }
  const text = destination.slice(start, end)
  if (control.type === 'text' && control.multiLine !== true && /[\r\n]/.test(text)) throw new RangeError('此纯文本控件未启用多行输入：' + control.id)
  if (control.type === 'comboBox' && /[\r\n]/.test(text)) throw new RangeError('组合框内容不能包含段落换行：' + control.id)
  if (control.type === 'checkbox') {
    if (!text) return ''
    const checked = contentControlCheckboxGlyph(control, true).text, unchecked = contentControlCheckboxGlyph(control, false).text
    if (text === checked && text === unchecked) return (control.checked ?? control.checkbox?.checked) ? 'true' : 'false'
    if (text === checked) return 'true'
    if (text === unchecked) return 'false'
    throw new RangeError('复选框只能包含配置的选中或未选中符号：' + control.id)
  }
  if (control.type === 'dropDownList' && text) throw new Error('下拉列表内容不能直接输入，请选择列表项：' + control.id)
  if (control.type === 'date') return parseContentControlDateText(text, control.dateFormat, control.dateLocale, control.dateCalendar) ?? text
  return text
}

interface BindingSource { control: DocumentContentControl; key: string; value: string }
interface BindingGroup { value: string; sourceIds: Set<string> }

function bindingSources(before: DocumentModel, after: DocumentModel, options: DocumentContentControlBindingOptions): BindingSource[] {
  const next = new Map((after.contentControls ?? []).map(control => [control.id, control]))
  const result: BindingSource[] = []
  let preparedEdits: OriginTextEdit[] | undefined, originPrepared = false
  const originEdits = () => {
    if (!originPrepared) { preparedEdits = plannedOriginTextEdits(before, after, options); originPrepared = true }
    return preparedEdits
  }
  for (const control of before.contentControls ?? []) {
    const key = documentContentControlBindingKey(control)
    if (!key || !scalarTypes.has(control.type) || control.metadata?.officeStoryLockSource) continue
    const current = next.get(control.id)
    if (current && (documentContentControlBindingKey(current) !== key || current.type !== control.type)) continue
    const oldValue = documentContentControlBindingValue(before, control)
    const value = current ? documentContentControlBindingValue(after, current)
      : control.temporary ? removedSourceValue(before, after, control, options, originEdits) : undefined
    if (value === undefined || current && oldValue === value && control.showingPlaceholder === current.showingPlaceholder) continue
    result.push({ control: current ?? control, key, value })
  }
  return result
}

function checkedScalar(value: string): boolean {
  if (value === 'true' || value === '1') return true
  if (value === 'false' || value === '0' || value === '') return false
  throw new RangeError('绑定值不是复选框可接受的布尔标量。')
}

function equivalentScalar(control: DocumentContentControl, actual: string | undefined, value: string): boolean {
  return actual === (control.type === 'checkbox' ? checkedScalar(value) ? 'true' : 'false' : value)
}

function peerNeedsUpdate(model: DocumentModel, control: DocumentContentControl, value: string): boolean {
  return !equivalentScalar(control, documentContentControlBindingValue(model, control), value)
    || control.type !== 'checkbox' && (value === '' ? control.showingPlaceholder !== true : control.showingPlaceholder === true)
}

function setBindingScalar(model: DocumentModel, control: DocumentContentControl, value: string): void {
  if (control.type === 'checkbox') setDocumentContentControlChecked(model, control.id, checkedScalar(value))
  else if (value === '') resetDocumentContentControlPlaceholder(model, control.id, true)
  else setDocumentContentControlValue(model, control.id, value)
}

/** Map one paragraph replacement through the originating caret/range in the same host. */
export function mapDocumentContentControlBindingSelection(selection: DocumentSelection, edit: DocumentContentControlBindingTextEdit,
  host?: DocumentAuxiliaryBlockHost): DocumentSelection {
  const result = cloneDocumentSelection(selection)
  if (!sameDocumentBlockHost(host, edit.host)) return result
  const map = (point: DocumentPoint): DocumentPoint => {
    if ((point.cellBlockId ?? point.blockId) !== edit.blockId) return point
    const offset = point.offset, delta = edit.text.length - (edit.endOffset - edit.startOffset)
    if (edit.startOffset === edit.endOffset) return offset < edit.startOffset ? point : { ...point, offset: offset + delta }
    if (offset <= edit.startOffset) return point
    return { ...point, offset: offset >= edit.endOffset ? offset + delta : edit.startOffset + edit.text.length }
  }
  return { ...result, anchor: map(result.anchor), focus: map(result.focus) }
}

function scalarTextEdit(before: DocumentModel, after: DocumentModel, control: DocumentContentControl): DocumentContentControlBindingTextEdit | undefined {
  const host = hostForControl(before, control), source = scopedModel(before, host), destination = scopedModel(after, host)
  const from = documentTextFieldParagraphs(source).find(paragraph => paragraph.block.id === control.blockId)
  const to = documentTextFieldParagraphs(destination).find(paragraph => paragraph.block.id === control.blockId)
  if (!from || !to) throw new RangeError('绑定控件修改后的文字宿主不存在：' + control.id)
  const end = control.endOffset + to.text.length - from.text.length
  if (from.text.slice(0, control.startOffset) !== to.text.slice(0, control.startOffset) || from.text.slice(control.endOffset) !== to.text.slice(end)
    || !isDocumentTextBoundary(to.text, end)) throw new RangeError('绑定控件修改超出其文字范围：' + control.id)
  const text = to.text.slice(control.startOffset, end)
  return text === from.text.slice(control.startOffset, control.endOffset) ? undefined
    : { controlId: control.id, blockId: control.blockId, startOffset: control.startOffset, endOffset: control.endOffset, text,
      ...(host ? { host: { ...host } } : {}) }
}

function assertControlPermission(model: DocumentModel, control: DocumentContentControl, context?: DocumentPermissionContext): void {
  assertDocumentContentControlValueEditable(model, control.id)
  if (context) assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForBlock(model, control.blockId))
}

/**
 * Compose scalar peers into a caller-owned originating candidate. XML/XPath evaluation and
 * custom-XML part writes are separate IO responsibilities; dirty/cachedResult are not an XML write.
 */
export function synchronizeDocumentContentControlBindings(before: DocumentModel, draft: DocumentModel,
  options: DocumentContentControlBindingOptions = {}): DocumentContentControlBindingResult {
  const result: DocumentContentControlBindingResult = { changedControlIds: [], textEdits: [],
    ...(options.selection ? { selection: cloneDocumentSelection(options.selection) } : {}) }
  const sources = bindingSources(before, draft, options)
  if (!sources.length) return result
  const candidate = cloneModel(draft), groups = new Map<string, BindingGroup>()
  const addSources = (changes: BindingSource[], initial = false) => {
    for (const source of changes) {
      const group = groups.get(source.key)
      if (group && group.value !== source.value && (initial || !equivalentScalar(source.control, source.value, group.value))) {
        throw new RangeError('同一编辑事务产生了冲突的绑定标量：' + source.control.id)
      }
      if (group) group.sourceIds.add(source.control.id)
      else groups.set(source.key, { value: source.value, sourceIds: new Set([source.control.id]) })
    }
  }
  addSources(sources, true)
  const updated = new Set<string>()
  for (const [key, group] of groups) {
    const ids = (candidate.contentControls ?? []).filter(control => documentContentControlBindingKey(control) === key && scalarTypes.has(control.type)
      && !control.metadata?.officeStoryLockSource).map(control => control.id)
    for (const id of ids) {
      const control = candidate.contentControls?.find(control => control.id === id)
      if (!control || group.sourceIds.has(id) || !peerNeedsUpdate(candidate, control, group.value)) continue
      if (updated.has(id)) throw new RangeError('绑定同步重复改变同一个内容控件：' + id)
      assertControlPermission(candidate, control, options.permissionContext)
      const stepBefore = cloneModel(candidate), previous = cloneModel(control)
      setBindingScalar(candidate, control, group.value)
      const edit = scalarTextEdit(stepBefore, candidate, previous)
      if (edit) {
        result.textEdits.push(edit)
        if (result.selection) result.selection = mapDocumentContentControlBindingSelection(result.selection, edit, options.host ?? documentEditingBlockHost(draft))
      }
      updated.add(id); result.changedControlIds.push(id)
      const target = candidate.contentControls?.find(control => control.id === id)
      if (target?.dataBinding) target.dataBinding = { ...target.dataBinding, cachedResult: group.value, dirty: true }
      const host = hostForControl(stepBefore, previous), source = scopedModel(stepBefore, host)
      const targetPoint = textTargets(source).find(point => (point.cellBlockId ?? point.blockId) === previous.blockId)
      addSources(bindingSources(stepBefore, candidate, { sourceControlId: id, sourceValue: group.value, activeControlId: id, host,
        ...(targetPoint ? { sourceSelection: { anchor: { ...targetPoint, offset: previous.startOffset }, focus: { ...targetPoint, offset: previous.endOffset } },
          selection: { anchor: { ...targetPoint, offset: previous.startOffset + (edit?.text.length ?? previous.endOffset - previous.startOffset) },
            focus: { ...targetPoint, offset: previous.startOffset + (edit?.text.length ?? previous.endOffset - previous.startOffset) } } } : {}) }))
    }
  }
  for (const source of sources) {
    if (options.permissionContext) assertDocumentPermission(before.permissionRules ?? [], options.permissionContext, 'edit', permissionTargetsForBlock(before, source.control.blockId))
  }
  for (const group of groups.values()) for (const id of group.sourceIds) {
    const control = candidate.contentControls?.find(control => control.id === id)
    if (control?.dataBinding) control.dataBinding = { ...control.dataBinding, cachedResult: group.value, dirty: true }
  }
  if (result.selection) {
    const scoped = scopedModel(candidate, options.host ?? documentEditingBlockHost(candidate))
    for (const point of [result.selection.anchor, result.selection.focus]) {
      const target = textTargets(scoped).find(target => sameTarget(target, point))
      if (!target || !isDocumentTextBoundary(runText(runsForTarget(scoped, target)), point.offset)) throw new RangeError('绑定同步后的选区不再是有效的文字位置。')
    }
  }
  candidate.wordCount = textTargets(candidate).reduce((count, target) => count + runText(runsForTarget(candidate, target)).replace(/\s/g, '').length, 0)
  Object.assign(draft, candidate)
  return result
}

function scalarDisplay(control: DocumentContentControl, value: string, peer: boolean): string {
  if (control.type === 'checkbox') return contentControlCheckboxGlyph(control, checkedScalar(value)).text
  if (peer && value === '') return documentContentControlPlaceholderText(control)
  if (control.type === 'date') return formatContentControlDate(value, control.dateFormat, control.dateLocale, control.dateCalendar)
  if (control.type === 'comboBox' || control.type === 'dropDownList') return control.items?.find(item => item.value === value)?.displayText ?? value
  return value.replace(/\r\n?/g, '\n')
}

/** Verify and map an explicit temporary value edit, including disjoint same-paragraph peers. */
export function documentContentControlTemporaryBindingCaret(before: DocumentModel, after: DocumentModel,
  controlId: string, value: string, host?: DocumentAuxiliaryBlockHost): DocumentSelection | undefined {
  const source = scopedModel(before, host), destination = scopedModel(after, host)
  const control = source.contentControls?.find(control => control.id === controlId), key = control && documentContentControlBindingKey(control)
  if (!control || !key || !control.temporary || destination.contentControls?.some(candidate => candidate.id === controlId)) return undefined
  const from = documentTextFieldParagraphs(source).find(paragraph => paragraph.block.id === control.blockId)
  const to = documentTextFieldParagraphs(destination).find(paragraph => paragraph.block.id === control.blockId)
  const target = textTargets(destination).find(target => (target.cellBlockId ?? target.blockId) === control.blockId)
  if (!from || !to || !target) return undefined
  const edits: DocumentContentControlBindingTextEdit[] = []
  for (const peer of source.contentControls ?? []) {
    if (peer.blockId !== control.blockId || documentContentControlBindingKey(peer) !== key || !scalarTypes.has(peer.type)) continue
    if (peer.id !== controlId && !peerNeedsUpdate(source, peer, value)) continue
    const text = scalarDisplay(peer, value, peer.id !== controlId)
    if (!isDocumentTextBoundary(from.text, peer.startOffset) || !isDocumentTextBoundary(from.text, peer.endOffset)) return undefined
    const duplicate = edits.find(edit => edit.startOffset === peer.startOffset && edit.endOffset === peer.endOffset)
    if (duplicate) { if (duplicate.text !== text) return undefined; continue }
    if (edits.some(edit => edit.startOffset < peer.endOffset && peer.startOffset < edit.endOffset)) return undefined
    edits.push({ controlId: peer.id, blockId: control.blockId, startOffset: peer.startOffset, endOffset: peer.endOffset, text })
  }
  const primary = edits.find(edit => edit.controlId === controlId) ?? edits.find(edit => edit.startOffset === control.startOffset && edit.endOffset === control.endOffset)
  if (!primary) return undefined
  let expected = from.text, offset = control.startOffset + primary.text.length
  for (const edit of [...edits].sort((left, right) => right.startOffset - left.startOffset)) {
    expected = expected.slice(0, edit.startOffset) + edit.text + expected.slice(edit.endOffset)
    if (edit !== primary && edit.endOffset <= control.startOffset) offset += edit.text.length - (edit.endOffset - edit.startOffset)
  }
  if (expected !== to.text || !isDocumentTextBoundary(to.text, offset)) return undefined
  return { anchor: { ...target, offset }, focus: { ...target, offset } }
}
