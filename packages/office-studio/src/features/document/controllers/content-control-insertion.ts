import type { DocumentContentControlType, DocumentModel } from '../../../domain/model'
import { documentId, findDocumentBlock, findDocumentBlockContainer, runText, runsForTarget, sameTarget, tableForBlock, textTargets, type DocumentSelection } from '../model/document-model'
import { createDocumentTableViewSelection, resolveDocumentTableViewSelection } from '../model/document-table-selection'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentContentControlTableRowBlockIds, resolveDocumentContentControlTableRows } from '../model/content-control-table-rows'
import { documentContentControlBlocksSelection, documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import type { ContentControlInput } from '../services/content-controls'
import { CONTENT_CONTROL_CHECKBOX_STYLES, CONTENT_CONTROL_DATE_FORMATS } from '../services/content-control-presentation'
import { documentContentControlSelectionForId } from './content-control-interaction'

/** Turn a live text/whole-row selection into the same persisted input used by the public add command. */
export function documentContentControlInputForSelection(model: DocumentModel, type: DocumentContentControlType,
  selection: DocumentSelection, id = documentId('content-control')): ContentControlInput & { id: string } {
  if (type === 'repeatingSectionItem') throw new RangeError('重复节项由重复节控件生成，不能单独插入。')
  if (selection.tableSelection) {
    if (type !== 'repeatingSection') throw new RangeError('请在单元格正文中选择此控件的插入范围。')
    const view = resolveDocumentTableViewSelection(model, selection.tableSelection)
    const canonical = createDocumentTableViewSelection(model, view.tableId, view.range, view.kind, view.pageAffinity, selection)
    if (!['row', 'table', 'range'].includes(view.kind) || view.range.startColumn !== 0
      || !sameTarget(canonical.anchor, selection.anchor) || !sameTarget(canonical.focus, selection.focus)
      || canonical.anchor.offset !== selection.anchor.offset || canonical.focus.offset !== selection.focus.offset) {
      throw new RangeError('重复行节必须使用当前完整行的真实结构选区。')
    }
    const block = resolveDocumentContentControlTableRows(model, {
      tableId: view.tableId, rowIds: tableRowIds(model, view.tableId, view.range.startRow, view.range.endRow),
    })
    if (view.range.endColumn !== block.grid.columnCount - 1
      || !documentContentControlTableRowBlockIds(block).has(canonical.anchor.cellBlockId ?? canonical.anchor.blockId)) {
      throw new RangeError('重复行节必须覆盖完整行并保留合法的输入段落。')
    }
    return { id, type, blockId: canonical.anchor.cellBlockId ?? canonical.anchor.blockId,
      startOffset: 0, endOffset: runText(runsForTarget(model, canonical.anchor)).length, placement: 'row',
      tableRowRange: { tableId: view.tableId, rowIds: [...block.rowIds] }, repeatingSection: { doNotAllowInsertDelete: false } }
  }
  const defaults: Partial<Record<DocumentContentControlType, Partial<ContentControlInput>>> = {
    richText: { placeholder: '在此处输入', showingPlaceholder: true },
    text: { placeholder: '在此处输入', showingPlaceholder: true, multiLine: false },
    checkbox: { checked: false, checkbox: { checked: false, ...CONTENT_CONTROL_CHECKBOX_STYLES.Cross }, metadata: { checkboxStyle: 'Cross' } },
    date: { placeholder: '选择日期', showingPlaceholder: true, dateFormat: CONTENT_CONTROL_DATE_FORMATS[0], dateCalendar: 'gregorian', dateStorage: 'dateTime' },
    dropDownList: { placeholder: '选择一项', showingPlaceholder: true, items: [{ displayText: '选择一项', value: '选择一项' }] },
    comboBox: { placeholder: '选择或输入一项', showingPlaceholder: true, items: [{ displayText: '选择一项', value: '选择一项' }] },
    picture: { placeholder: '选择图片', showingPlaceholder: true },
    repeatingSection: { repeatingSection: { doNotAllowInsertDelete: false } },
    docPartObj: { docPart: { gallery: 'Quick Parts', category: 'General', unique: false } },
  }
  const targets = textTargets(model), anchor = targets.find(target => sameTarget(target, selection.anchor)), focus = targets.find(target => sameTarget(target, selection.focus))
  if (!anchor || !focus || !isDocumentTextBoundary(runText(runsForTarget(model, anchor)), selection.anchor.offset)
    || !isDocumentTextBoundary(runText(runsForTarget(model, focus)), selection.focus.offset)) throw new RangeError('内容控件插入范围已失效。')
  if (!sameTarget(anchor, focus) || type === 'repeatingSection' || type === 'group') {
    if (!['text', 'richText', 'group', 'repeatingSection', 'docPartObj'].includes(type)) throw new RangeError('此控件类型需要单段文字范围。')
    const container = findDocumentBlockContainer(model, anchor.cellBlockId ?? anchor.blockId)
    if (!container || container !== findDocumentBlockContainer(model, focus.cellBlockId ?? focus.blockId)) {
      throw new RangeError('完整块控件不能跨越单元格、续格或内容区容器。')
    }
    const first = container.findIndex(block => block.id === (anchor.cellBlockId ?? anchor.blockId))
    const last = container.findIndex(block => block.id === (focus.cellBlockId ?? focus.blockId))
    const blockRange = { blockIds: container.slice(Math.min(first, last), Math.max(first, last) + 1).map(block => block.id) }
    const resolved = resolveDocumentContentControlBlocks(model, blockRange), full = documentContentControlBlocksSelection(model, blockRange)
    const collapsed = sameTarget(anchor, focus) && selection.anchor.offset === selection.focus.offset
    const forward = sameTarget(full.anchor, selection.anchor) && full.anchor.offset === selection.anchor.offset
      && sameTarget(full.focus, selection.focus) && full.focus.offset === selection.focus.offset
    const reverse = sameTarget(full.focus, selection.anchor) && full.focus.offset === selection.anchor.offset
      && sameTarget(full.anchor, selection.focus) && full.anchor.offset === selection.focus.offset
    if (!collapsed && !forward && !reverse) throw new RangeError('完整块控件必须选中首尾段落的完整内容。')
    const owned = documentContentControlOwnedBlockIds(resolved)
    const value = targets.filter(target => owned.has(target.cellBlockId ?? target.blockId))
      .map(target => runText(runsForTarget(model, target))).join('\n')
    return { ...defaults[type], id, type, blockId: full.anchor.cellBlockId ?? full.anchor.blockId, startOffset: 0,
      endOffset: runText(runsForTarget(model, full.anchor)).length, placement: 'block', blockRange,
      showingPlaceholder: false, ...(['text', 'richText'].includes(type) ? { value, ...(type === 'text' ? { multiLine: true } : {}) } : {}) }
  }
  const target = focus, start = Math.min(selection.anchor.offset, selection.focus.offset), end = Math.max(selection.anchor.offset, selection.focus.offset)
  return { ...defaults[type], id, type, blockId: target.cellBlockId ?? target.blockId,
    startOffset: start, endOffset: end, placement: target.cellBlockId ? 'cell' : 'inline',
    ...(start !== end ? { showingPlaceholder: false, ...(['text', 'richText'].includes(type) ? { value: runText(runsForTarget(model, target)).slice(start, end) } : {}) } : {}) }
}

function tableRowIds(model: DocumentModel, tableId: string, start: number, end: number): string[] {
  // Resolve the table itself rather than deriving a row identity from a paragraph/storage column.
  const table = findDocumentBlock(model, tableId)
  const value = table && tableForBlock(table)
  if (!value || end >= value.rows.length) throw new RangeError('内容控件所属表格已不存在或行范围已失效。')
  return value.rows.slice(start, end + 1).map(row => row.id)
}

/** Cancellation and ambiguous automatic children never produce an active ID from a facade placeholder. */
export function documentContentControlSelectionAfterInsertion(before: DocumentModel, after: DocumentModel,
  input: ContentControlInput & { id: string }, host?: DocumentAuxiliaryBlockHost): { controlId: string; selection: DocumentSelection } | undefined {
  if (before.id !== after.id || before.contentControls?.some(control => control.id === input.id)) return undefined
  const control = after.contentControls?.find(control => control.id === input.id && control.type === input.type
    && control.blockId === input.blockId && control.placement === input.placement
    && JSON.stringify(control.tableRowRange) === JSON.stringify(input.tableRowRange)
    && JSON.stringify(control.blockRange) === JSON.stringify(input.blockRange))
  if (!control) return undefined
  const oldIds = new Set(before.contentControls?.map(control => control.id))
  const items = control.type === 'repeatingSection' ? after.contentControls?.filter(item => !oldIds.has(item.id)
    && item.type === 'repeatingSectionItem' && item.parentId === control.id && item.placement === control.placement
    && item.blockId === control.blockId && item.startOffset === control.startOffset && item.endOffset === control.endOffset
    && JSON.stringify(item.tableRowRange) === JSON.stringify(control.tableRowRange)
    && JSON.stringify(item.blockRange) === JSON.stringify(control.blockRange)) ?? [] : [control]
  if (items.length !== 1) return undefined
  const selected = items[0]!, selection = documentContentControlSelectionForId(after, selected.id, host)
  return selection ? { controlId: selected.id, selection } : undefined
}
