import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentBlockHostForId, documentEditingBlockHost, sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { findDocumentBlock, runText, runsForTarget, sameTarget, tableForBlock, textTargets, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import { documentTableGrid } from '../model/document-table-grid'
import { cloneDocumentSelection } from '../model/document-table-selection'
import { documentTableCellBlocks } from '../model/document-table-continuations'
import { documentTableTextOriginCells } from '../model/document-table-text-origin'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { assertDocumentPermission, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentContentControlPermissionTargets } from '../permissions/content-control-permissions'
import { documentContentControlBlocksSelection, documentContentControlBlockTextTargets, documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
import { documentContentControlBlocksValueMutationRange } from '../model/content-control-block-mutations'
import type { DocumentContentControlEditResult } from './content-control-editing'
import { contentControlCheckboxGlyph, formatContentControlDate, parseContentControlDateText } from './content-control-presentation'
import { assertDocumentContentControlValueEditable, contentControlText, documentContentControlPlaceholderText,
  resetDocumentContentControlPlaceholder, setDocumentContentControlChecked, setDocumentContentControlValue } from './content-controls'
import { documentModelForBlockHost } from './document-block-host'
import { createDocumentCustomXmlBindingSession, type DocumentCustomXmlControlBinding } from './document-custom-xml-bindings'

const scalarTypes = new Set<DocumentContentControl['type']>(['text', 'richText', 'checkbox', 'date', 'dropDownList', 'comboBox'])

export interface DocumentContentControlBindingTextEdit {
  controlId: string
  blockId: string
  startOffset: number
  endOffset: number
  text: string
  host?: DocumentAuxiliaryBlockHost
  /** A real block splice maps replaced paragraph/table IDs to the resulting logical range. */
  blockRangeEdit?: { sourceBlockIds: string[]; sourceSelection: DocumentSelection; selection: DocumentSelection }
}

/** Exact UTF-16 ranges in the flattened source host, captured by the originating text planner. */
export interface DocumentContentControlBindingOriginTextEdit {
  startOffset: number
  endOffset: number
  text: string
  /** Exact cell identity and semantic value keep structural delimiters out of temporary scalars. */
  tableCell?: {
    tableId: string
    cellId: string
    row: number
    column: number
    order: number
    value: string
  }
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
  /** A custom-XML command replaces the data source; existing UI text must not overwrite it. */
  refreshCustomXml?: boolean
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
  return documentBlockHostForId(model, control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId) ?? documentEditingBlockHost(model)
}

function absolutePoint(model: DocumentModel, point: DocumentPoint): number | undefined {
  const paragraph = documentTextFieldParagraphs(model).find(paragraph => paragraph.block.id === (point.cellBlockId ?? point.blockId))
  return paragraph && isDocumentTextBoundary(paragraph.text, point.offset) ? paragraph.absoluteStart + point.offset : undefined
}

interface OriginTextEdit {
  start: number; end: number; text: string; source: string; destination: string
  cell?: { tableId: string; cellId: string; order: number; value: string; blockIds: Set<string> }
}

function plannedOriginTextEdits(before: DocumentModel, after: DocumentModel,
  options: DocumentContentControlBindingOptions): OriginTextEdit[] | undefined {
  if (!options.originTextEdits) {
    const edit = originTextEdit(before, after, options)
    return edit ? [edit] : undefined
  }
  const sourceModel = scopedModel(before, options.host), destinationModel = scopedModel(after, options.host)
  const paragraphs = documentTextFieldParagraphs(sourceModel), source = paragraphs.map(paragraph => paragraph.text).join('\n')
  const destination = documentTextFieldParagraphs(destinationModel).map(paragraph => paragraph.text).join('\n')
  const origins = options.originTextEdits.some(edit => edit.tableCell) ? documentTableTextOriginCells(sourceModel, paragraphs) : []
  const lastEmptyOrder = options.originTextEdits.reduce((maximum, edit) => Math.max(maximum, edit.tableCell?.order ?? -1), -1)
  const edits = options.originTextEdits.map(edit => {
    if (!Number.isSafeInteger(edit.startOffset) || !Number.isSafeInteger(edit.endOffset) || edit.startOffset < 0
      || edit.endOffset < edit.startOffset || typeof edit.text !== 'string'
      || !isDocumentTextBoundary(source, edit.startOffset) || !isDocumentTextBoundary(source, edit.endOffset)) {
      throw new RangeError('绑定源编辑范围不是有效的 UTF-16 区间。')
    }
    const normalized = edit.text.replace(/\r\n?/g, '\n'), recorded = edit.tableCell
    let cell: OriginTextEdit['cell']
    if (recorded) {
      if (typeof recorded.tableId !== 'string' || !recorded.tableId || typeof recorded.cellId !== 'string' || !recorded.cellId
        || !Number.isSafeInteger(recorded.row) || recorded.row < 0 || !Number.isSafeInteger(recorded.column) || recorded.column < 0
        || !Number.isSafeInteger(recorded.order) || recorded.order < 0 || typeof recorded.value !== 'string') throw new RangeError('绑定源单元格身份或实际值无效。')
      const sourceBlock = findDocumentBlock(sourceModel, recorded.tableId), targetBlock = findDocumentBlock(destinationModel, recorded.tableId)
      const sourceTable = sourceBlock && tableForBlock(sourceBlock), targetTable = targetBlock && tableForBlock(targetBlock)
      if (!sourceTable || !targetTable) throw new RangeError('绑定源单元格的表格不存在。')
      const owner = documentTableGrid(sourceTable).cellAt(recorded.row, recorded.column), target = documentTableGrid(targetTable).cellAt(recorded.row, recorded.column)
      const origin = owner && origins.find(origin => origin.tableId === recorded.tableId && origin.cell === owner.cell)
      if (!owner || owner.cell.id !== recorded.cellId || owner.row !== recorded.row || owner.column !== recorded.column
        || !target || target.cell.id !== recorded.cellId || target.row !== recorded.row || target.column !== recorded.column
        || !origin || origin.order !== recorded.order) throw new RangeError('绑定源单元格的原始身份或结构顺序不一致。')
      const inputIds = new Set(documentTextFieldParagraphs({ blocks: documentTableCellBlocks(owner.cell) }).map(paragraph => paragraph.block.id))
      const inputs = paragraphs.filter(paragraph => inputIds.has(paragraph.block.id)), first = inputs[0], last = inputs.at(-1)
      const value = recorded.value.replace(/\r\n?/g, '\n'), outputs = documentTextFieldParagraphs({ blocks: documentTableCellBlocks(target.cell) })
      if (outputs.map(paragraph => paragraph.text).join('\n') !== value || outputs.length !== value.split('\n').length) throw new RangeError('绑定源单元格的实际值与候选内容不一致。')
      const following = paragraphs[origin.paragraphIndex], previous = paragraphs.at(-1), start = first?.absoluteStart ?? following?.absoluteStart ?? previous?.absoluteEnd ?? 0
      const end = last?.absoluteEnd ?? start, text = first && last ? value : following ? value + '\n' : previous ? '\n' + value
        : value + (recorded.order === lastEmptyOrder ? '' : '\n')
      if (edit.startOffset !== start || edit.endOffset !== end || normalized !== text) throw new RangeError('绑定源单元格的文字坐标或分隔符计划不一致。')
      cell = { tableId: recorded.tableId, cellId: recorded.cellId, order: recorded.order, value, blockIds: inputIds }
    }
    return { start: edit.startOffset, end: edit.endOffset, text: normalized, source, destination, ...(cell ? { cell } : {}) }
  }).sort((left, right) => right.start - left.start || right.end - left.end || (right.cell?.order ?? -1) - (left.cell?.order ?? -1))
  let expected = source, following: OriginTextEdit | undefined
  for (const edit of edits) {
    if (following && (edit.end > following.start || edit.start === following.start && (edit.start !== edit.end
      || !edit.cell || !following.cell || edit.cell.order >= following.cell.order
      || edit.cell.tableId === following.cell.tableId && edit.cell.cellId === following.cell.cellId))) throw new RangeError('绑定源编辑范围重叠。')
    expected = expected.slice(0, edit.start) + edit.text + expected.slice(edit.end)
    following = edit
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
  const cell = edits?.find(edit => edit.cell?.blockIds.has(control.blockId))?.cell
  if (cell) {
    const value = removedInputScalar(control, cell.value)
    return !control.showingPlaceholder && documentContentControlBindingValue(before, control) === value ? undefined : value
  }
  const sourceModel = scopedModel(before, options.host), paragraphs = documentTextFieldParagraphs(sourceModel)
  const paragraph = paragraphs.find(paragraph => paragraph.block.id === control.blockId)
  if (!paragraph) {
    if (active) throw new RangeError('已卸壳绑定控件的源内容区不存在：' + control.id)
    return undefined
  }
  if (!edits) throw new RangeError('无法保留已卸壳绑定控件的实际输入值：' + control.id)
  const blockTargets = control.blockRange ? documentContentControlBlockTextTargets(sourceModel, control.blockRange) : []
  const first = blockTargets.length ? paragraphs.find(paragraph => paragraph.block.id === (blockTargets[0]!.cellBlockId ?? blockTargets[0]!.blockId)) : undefined
  const last = blockTargets.length ? paragraphs.find(paragraph => paragraph.block.id === (blockTargets.at(-1)!.cellBlockId ?? blockTargets.at(-1)!.blockId)) : undefined
  let start = first?.absoluteStart ?? paragraph.absoluteStart + control.startOffset
  let end = last?.absoluteEnd ?? paragraph.absoluteStart + control.endOffset, changed = false
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
  return removedInputScalar(control, destination.slice(start, end))
}

/** The originating planner supplies actual input, independent of structural paragraph delimiters. */
function removedInputScalar(control: DocumentContentControl, text: string): string {
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

/** Detached source snapshots for the originating candidate; collecting never publishes changes. */
export interface DocumentContentControlBindingSource {
  readonly control: DocumentContentControl
  readonly key: string
  readonly value: string
}
interface BindingGroup { value: string; sourceIds: Set<string> }

export function documentContentControlBindingSources(before: DocumentModel, after: DocumentModel,
  options: DocumentContentControlBindingOptions = {}): readonly DocumentContentControlBindingSource[] {
  const next = new Map((after.contentControls ?? []).map(control => [control.id, control]))
  const result: DocumentContentControlBindingSource[] = []
  let preparedEdits: OriginTextEdit[] | undefined, originPrepared = false
  const originEdits = () => {
    if (!originPrepared) { preparedEdits = plannedOriginTextEdits(before, after, options); originPrepared = true }
    return preparedEdits
  }
  if (options.originTextEdits) originEdits()
  for (const control of before.contentControls ?? []) {
    const key = documentContentControlBindingKey(control)
    if (!key || !scalarTypes.has(control.type) || control.metadata?.officeStoryLockSource) continue
    const current = next.get(control.id)
    if (current && (documentContentControlBindingKey(current) !== key || current.type !== control.type)) continue
    const oldValue = documentContentControlBindingValue(before, control)
    const value = current ? documentContentControlBindingValue(after, current)
      : control.temporary ? removedSourceValue(before, after, control, options, originEdits) : undefined
    if (value === undefined || current && oldValue === value && control.showingPlaceholder === current.showingPlaceholder) continue
    result.push({ control: cloneModel(current ?? control), key, value })
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
  if (edit.blockRangeEdit) {
    const structure = edit.blockRangeEdit, owned = new Set(structure.sourceBlockIds)
    const map = (point: DocumentPoint): DocumentPoint => {
      if (!owned.has(point.cellBlockId ?? point.blockId)) return point
      const start = structure.sourceSelection.anchor
      const destination = sameTarget(point, start) && point.offset <= start.offset ? structure.selection.anchor : structure.selection.focus
      return { ...destination }
    }
    if (result.tableSelection && owned.has(result.tableSelection.tableId)) delete result.tableSelection
    return { ...result, anchor: map(result.anchor), focus: map(result.focus) }
  }
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
  if (control.blockRange) {
    const replacement = documentContentControlBlocksValueMutationRange(source, destination, control.id)
    if (!replacement) throw new RangeError('绑定控件完整块替换后的逻辑范围不存在：' + control.id)
    const text = documentContentControlBlockTextTargets(destination, { blockIds: replacement.blockIds }).map(target => runText(runsForTarget(destination, target))).join('\n')
    if (text === contentControlText(source, control) && JSON.stringify(replacement.blockIds) === JSON.stringify(control.blockRange.blockIds)) return undefined
    return { controlId: control.id, blockId: control.blockId, startOffset: control.startOffset, endOffset: control.endOffset, text,
      blockRangeEdit: { sourceBlockIds: [...documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(source, control.blockRange))],
        sourceSelection: documentContentControlBlocksSelection(source, control.blockRange), selection: replacement.selection }, ...(host ? { host: { ...host } } : {}) }
  }
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
  if (context) assertDocumentPermission(model.permissionRules ?? [], context, 'edit', documentContentControlPermissionTargets(scopedModel(model, hostForControl(model, control)), control, model))
}

/**
 * Compose scalar peers and custom-XML writes into a caller-owned originating candidate.
 * The platform session owns detached XML handles; legacy snapshots retain raw-key peer matching.
 */
export function synchronizeDocumentContentControlBindings(before: DocumentModel, draft: DocumentModel,
  options: DocumentContentControlBindingOptions = {}): DocumentContentControlBindingResult {
  const result: DocumentContentControlBindingResult = { changedControlIds: [], textEdits: [],
    ...(options.selection ? { selection: cloneDocumentSelection(options.selection) } : {}) }
  const sources = options.refreshCustomXml ? [] : documentContentControlBindingSources(before, draft, options)
  if (!sources.length && !options.refreshCustomXml) return result
  const candidate = cloneModel(draft), groups = new Map<string, BindingGroup>()
  const session = createDocumentCustomXmlBindingSession(candidate)
  const pending: string[] = [], queued = new Set<string>()
  let xmlBindings = new Map<string, DocumentCustomXmlControlBinding>()
  const enqueue = (key: string) => { if (!queued.has(key)) { queued.add(key); pending.push(key) } }
  const addSources = (changes: readonly DocumentContentControlBindingSource[], initial = false) => {
    for (const source of changes) {
      const group = groups.get(source.key)
      if (!session && group && group.value !== source.value && (initial || !equivalentScalar(source.control, source.value, group.value))) {
        throw new RangeError('同一编辑事务产生了冲突的绑定标量：' + source.control.id)
      }
      if (group) { group.sourceIds.add(source.control.id); if (session) group.value = source.value }
      else groups.set(source.key, { value: source.value, sourceIds: new Set([source.control.id]) })
      enqueue(source.key)
    }
  }
  const addXmlBindings = (values: readonly DocumentCustomXmlControlBinding[]) => {
    xmlBindings = new Map(values.map(binding => [binding.controlId, binding]))
    for (const binding of values) {
      const group = groups.get(binding.key)
      if (group) group.value = binding.value
      else groups.set(binding.key, { value: binding.value, sourceIds: new Set() })
      enqueue(binding.key)
    }
  }
  try {
    if (options.refreshCustomXml) { if (session) addXmlBindings(session.refresh()) }
    else {
      addSources(session ? session.applySources(sources) : sources, true)
      if (session) addXmlBindings(session.bindings())
    }
    const updated = new Set<string>()
    for (let cursor = 0; cursor < pending.length; cursor++) {
      const key = pending[cursor]!, group = groups.get(key)!
      queued.delete(key)
      const ids = session ? [...xmlBindings.values()].filter(binding => binding.key === key).map(binding => binding.controlId)
        : (candidate.contentControls ?? []).filter(control => documentContentControlBindingKey(control) === key && scalarTypes.has(control.type)
          && !control.metadata?.officeStoryLockSource).map(control => control.id)
      for (const id of ids) {
        const control = candidate.contentControls?.find(control => control.id === id)
        if (!control || (session ? xmlBindings.get(id)?.key !== key : group.sourceIds.has(id)) || !peerNeedsUpdate(candidate, control, group.value)) continue
        if (updated.has(id)) throw new RangeError('绑定同步重复改变同一个内容控件：' + id)
        assertControlPermission(candidate, control, options.permissionContext)
        const stepBefore = cloneModel(candidate), previous = cloneModel(control), value = group.value
        setBindingScalar(candidate, control, value)
        const edit = scalarTextEdit(stepBefore, candidate, previous)
        if (edit) {
          result.textEdits.push(edit)
          if (result.selection) result.selection = mapDocumentContentControlBindingSelection(result.selection, edit, options.host ?? documentEditingBlockHost(draft))
        }
        updated.add(id); result.changedControlIds.push(id)
        const target = candidate.contentControls?.find(control => control.id === id)
        if (!session && target?.dataBinding) target.dataBinding = { ...target.dataBinding, cachedResult: value, dirty: true }
        const host = hostForControl(stepBefore, previous), source = scopedModel(stepBefore, host)
        const targetPoint = textTargets(source).find(point => (point.cellBlockId ?? point.blockId) === previous.blockId)
        const sourceSelection = previous.blockRange ? documentContentControlBlocksSelection(source, previous.blockRange)
          : targetPoint ? { anchor: { ...targetPoint, offset: previous.startOffset }, focus: { ...targetPoint, offset: previous.endOffset } } : undefined
        const destinationPoint = edit?.blockRangeEdit?.selection.focus
        const selection = destinationPoint ? { anchor: { ...destinationPoint }, focus: { ...destinationPoint } }
          : targetPoint ? { anchor: { ...targetPoint, offset: previous.startOffset + (edit?.text.length ?? previous.endOffset - previous.startOffset) },
            focus: { ...targetPoint, offset: previous.startOffset + (edit?.text.length ?? previous.endOffset - previous.startOffset) } } : undefined
        let cascades = documentContentControlBindingSources(stepBefore, candidate, { sourceControlId: id, sourceValue: value, activeControlId: id, host,
          ...(sourceSelection && selection ? { sourceSelection, selection } : {}) })
        // This peer already reflects its XML node. Only other changed controls may write a cascade.
        // In particular, a source value predicate can now select another node without rewriting it.
        if (session) cascades = cascades.filter(source => source.control.id !== id)
        addSources(session ? session.applySources(cascades) : cascades)
        if (session) addXmlBindings(session.bindings())
      }
    }
    for (const source of sources) {
      const original = before.contentControls?.find(control => control.id === source.control.id) ?? source.control
      if (options.permissionContext) assertDocumentPermission(before.permissionRules ?? [], options.permissionContext, 'edit', documentContentControlPermissionTargets(scopedModel(before, hostForControl(before, original)), original, before))
    }
    if (session) {
      // Session publication also updates caches of unchanged XML peers; those metadata writes
      // must obey the same document/paragraph/story permissions as visible scalar updates.
      for (const binding of session.bindings()) {
        const control = candidate.contentControls?.find(control => control.id === binding.controlId)
        if (control?.dataBinding && (control.dataBinding.cachedResult !== binding.value || control.dataBinding.dirty !== false)
          && options.permissionContext) assertDocumentPermission(before.permissionRules ?? [], options.permissionContext, 'edit', documentContentControlPermissionTargets(scopedModel(candidate, hostForControl(candidate, control)), control, candidate))
      }
    } else for (const group of groups.values()) for (const id of group.sourceIds) {
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
    session?.commit()
    Object.assign(draft, candidate)
    return result
  } finally {
    session?.dispose()
  }
}

function scalarDisplay(control: DocumentContentControl, value: string, peer: boolean): string {
  if (control.type === 'checkbox') return contentControlCheckboxGlyph(control, checkedScalar(value)).text
  if (peer && value === '') return documentContentControlPlaceholderText(control)
  if (control.type === 'date') return formatContentControlDate(value, control.dateFormat, control.dateLocale, control.dateCalendar)
  if (control.type === 'comboBox' || control.type === 'dropDownList') return control.items?.find(item => item.value === value)?.displayText ?? value
  return value.replace(/\r\n?/g, '\n')
}

/** Verify an explicit temporary value edit using actual surviving peer ranges, including XML aliases. */
export function documentContentControlTemporaryBindingCaret(before: DocumentModel, after: DocumentModel,
  controlId: string, value: string, host?: DocumentAuxiliaryBlockHost, sourceText?: string): DocumentSelection | undefined {
  const source = scopedModel(before, host), destination = scopedModel(after, host)
  const control = source.contentControls?.find(control => control.id === controlId), key = control && documentContentControlBindingKey(control)
  if (!control || !key || !control.temporary || destination.contentControls?.some(candidate => candidate.id === controlId)) return undefined
  const from = documentTextFieldParagraphs(source).find(paragraph => paragraph.block.id === control.blockId)
  const to = documentTextFieldParagraphs(destination).find(paragraph => paragraph.block.id === control.blockId)
  const target = textTargets(destination).find(target => (target.cellBlockId ?? target.blockId) === control.blockId)
  if (!from || !to || !target) return undefined
  const retained = new Map((destination.contentControls ?? []).map(peer => [peer.id, peer]))
  const edits: DocumentContentControlBindingTextEdit[] = []
  const observed: { edit: DocumentContentControlBindingTextEdit; control: DocumentContentControl }[] = []
  for (const peer of source.contentControls ?? []) {
    const peerKey = documentContentControlBindingKey(peer), current = retained.get(peer.id)
    if (peer.blockId !== control.blockId || !peerKey || !scalarTypes.has(peer.type) || peer.metadata?.officeStoryLockSource) continue
    if (peer.startOffset > peer.endOffset || !isDocumentTextBoundary(from.text, peer.startOffset)
      || !isDocumentTextBoundary(from.text, peer.endOffset)) return undefined
    let text: string
    if (peer.id === controlId) text = sourceText ?? scalarDisplay(peer, value, false)
    else if (current) {
      if (current.type !== peer.type || current.blockId !== peer.blockId || documentContentControlBindingKey(current) !== peerKey
        || current.startOffset > current.endOffset || !isDocumentTextBoundary(to.text, current.startOffset)
        || !isDocumentTextBoundary(to.text, current.endOffset)) return undefined
      // Use the committed display, not the source scalar: a predicate may now select
      // another XML node, or an independent binding in the touched store may refresh.
      text = to.text.slice(current.startOffset, current.endOffset)
      if (text === from.text.slice(peer.startOffset, peer.endOffset)) continue
    } else {
      // A removed peer has no final anchor. Only the existing raw-key contract
      // supplies its scalar; other removed aliases still require an exact edit trace.
      if (!peer.temporary || peerKey !== key || !peerNeedsUpdate(source, peer, value)) continue
      text = scalarDisplay(peer, value, true)
    }
    const duplicate = edits.find(edit => edit.startOffset === peer.startOffset && edit.endOffset === peer.endOffset)
    if (duplicate) {
      if (duplicate.text !== text) return undefined
      if (current) observed.push({ edit: duplicate, control: current })
      continue
    }
    if (edits.some(edit => edit.startOffset < peer.endOffset && peer.startOffset < edit.endOffset)) return undefined
    const edit = { controlId: peer.id, blockId: control.blockId, startOffset: peer.startOffset, endOffset: peer.endOffset, text }
    edits.push(edit)
    if (current) observed.push({ edit, control: current })
  }
  const primary = edits.find(edit => edit.controlId === controlId) ?? edits.find(edit => edit.startOffset === control.startOffset && edit.endOffset === control.endOffset)
  if (!primary) return undefined
  let expected = from.text, offset = control.startOffset + primary.text.length
  for (const edit of [...edits].sort((left, right) => right.startOffset - left.startOffset || right.endOffset - left.endOffset)) {
    expected = expected.slice(0, edit.startOffset) + edit.text + expected.slice(edit.endOffset)
    if (edit !== primary && edit.endOffset <= control.startOffset) offset += edit.text.length - (edit.endOffset - edit.startOffset)
  }
  if (expected !== to.text || !isDocumentTextBoundary(to.text, offset)) return undefined
  for (const { edit, control: current } of observed) {
    const start = edits.reduce((at, other) => other !== edit && other.endOffset <= edit.startOffset
      ? at + other.text.length - (other.endOffset - other.startOffset) : at, edit.startOffset)
    if (current.startOffset !== start || current.endOffset !== start + edit.text.length) return undefined
  }
  return { anchor: { ...target, offset }, focus: { ...target, offset } }
}
