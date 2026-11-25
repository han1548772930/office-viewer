import type { DocumentBlock, DocumentCalloutConfig, DocumentCodeConfig, DocumentColumnGroup, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { clampPoint, collapsedSelection, createParagraph, documentBlockAncestry, documentId, findDocumentBlock,
  findDocumentBlockContainer, orderedSelection, runText, runsForTarget, textTargets, applyRunPatch, documentRunFormatPatch, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import { defaultDocumentCalloutConfig, defaultDocumentCodeConfig, documentBlockChildContainers, isDocumentStructuredBlock,
  isDocumentTextBlock, validateDocumentCalloutConfig, validateDocumentCodeConfig, validateDocumentColumnGroup,
  visitDocumentBlocks, DOCUMENT_COLUMN_GROUP_MIN_COLUMNS, DOCUMENT_COLUMN_GROUP_MAX_COLUMNS, type DocumentDecoratedBlockType } from '../model/document-structural-blocks'
import { documentDefaultStyle } from '../model/document-default-style'
import { validateDocumentStoryBlocks } from '../model/document-story-blocks'
import { reconcileDocumentContentControlBlockEdits } from '../model/content-control-block-edits'
import { resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
import { validateDocumentContentControlRanges } from '../services/content-controls'
import { assertDocumentTextSelectionEditable } from '../services/document-text-formatting'
import { removeDeletedDocumentRepeatingNoteScopes } from '../services/repeating-content-control-copy'
import { assertDocumentPermission, permissionTargetsForChangedBlocks, type DocumentPermissionContext } from '../permissions/document-permissions'
import { updateDocumentStructuredCodeBlock } from '../services/document-code-editing'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFields } from '../model/document-text-fields'
import { rebaseDocumentTableViewSelection } from '../model/document-table-selection'
import { reconcileDocumentNotesAfterStructure } from '../notes/document-notes'
import { codeBlockData, indentCodeBlockWithSelection } from '../code-blocks/code-blocks'
import { validateDocumentRunPatch } from '../model/document-run-range'
import { clearDocumentQuoteParagraphBorders, type DocumentQuoteStyle } from '../model/document-structured-appearance'
import { planDocumentTableFieldAnchors } from '../services/document-table-anchors'
import { setDocumentTextFields, documentTextFieldRanges } from '../model/document-text-fields'
import { reconcileDocumentTocFieldParents } from '../toc/document-toc-edit'
import { insertDocumentBlocksAtPoint } from '../services/document-block-insertion'

export interface DocumentColumnGroupCreateOptions {
  columnGroupId?: string
  columnIds?: string[]
  columnCount?: number
  gap?: number
  horizontalPadding?: number
  minHeight?: number
  widthRatios?: number[]
  minWidths?: Array<number | null>
  layout?: DocumentColumnGroup['layout']
  responsive?: DocumentColumnGroup['responsive']
}

export type DocumentStructureCommand =
  | { id: 'doc.command.wrap-structured-block'; blockType: DocumentDecoratedBlockType; blockId?: string; blockIds?: string[]; callout?: Partial<DocumentCalloutConfig>; code?: Partial<DocumentCodeConfig>; quote?: Partial<DocumentQuoteStyle>; textColor?: string | null }
  | { id: 'doc.command.insert-structured-block'; blockType: DocumentDecoratedBlockType; blockId?: string; afterBlockId?: string; at?: DocumentPoint; source?: string; callout?: Partial<DocumentCalloutConfig>; code?: Partial<DocumentCodeConfig>; quote?: Partial<DocumentQuoteStyle>; textColor?: string | null }
  | { id: 'doc.command.unwrap-structured-block'; blockId: string }
  | { id: 'doc.command.remove-structured-block'; blockId: string }
  | { id: 'doc.command.exit-structured-block'; blockId: string }
  | { id: 'doc.command.update-callout'; blockId: string; patch: Partial<DocumentCalloutConfig>; textColor?: string | null }
  | { id: 'doc.command.update-quote-style'; blockId: string; patch: Partial<DocumentQuoteStyle> }
  | { id: 'doc.command.update-code-settings'; blockId: string; patch: Partial<DocumentCodeConfig> }
  | { id: 'doc.command.update-code-block'; blockId: string; source?: string; patch?: Partial<DocumentCodeConfig>; host?: DocumentAuxiliaryBlockHost }
  | { id: 'doc.command.indent-code-block'; blockId: string; direction: -1 | 1 }
  | ({ id: 'doc.command.insert-column-group'; afterBlockId?: string; at?: DocumentPoint } & DocumentColumnGroupCreateOptions)
  | ({ id: 'doc.command.wrap-column-group'; blockIds?: string[] } & DocumentColumnGroupCreateOptions)
  | { id: 'doc.command.update-column-group'; blockId: string; patch?: Partial<Pick<DocumentColumnGroup, 'gap' | 'horizontalPadding' | 'minHeight' | 'layout' | 'responsive'>>; widthRatios?: number[]; minWidths?: Array<number | null> }
  | { id: 'doc.command.add-column'; blockId: string; columnId?: string; newColumnId?: string; position?: 'before' | 'after' }
  | { id: 'doc.command.remove-column'; blockId: string; columnId: string; keepContents?: boolean }
  | { id: 'doc.command.move-column'; blockId: string; columnId: string; index: number }
  | { id: 'doc.command.set-column-text'; blockId: string; columnId: string; text: string }
  | { id: 'doc.command.insert-column-paragraph'; blockId: string; columnId: string; index: number; text?: string; paragraphId?: string }

const ids = new Set<DocumentStructureCommand['id']>(['doc.command.wrap-structured-block', 'doc.command.insert-structured-block',
  'doc.command.unwrap-structured-block', 'doc.command.remove-structured-block', 'doc.command.exit-structured-block', 'doc.command.update-callout', 'doc.command.update-quote-style', 'doc.command.update-code-settings', 'doc.command.update-code-block', 'doc.command.indent-code-block',
  'doc.command.insert-column-group', 'doc.command.wrap-column-group', 'doc.command.update-column-group', 'doc.command.add-column',
  'doc.command.remove-column', 'doc.command.move-column', 'doc.command.set-column-text', 'doc.command.insert-column-paragraph'])
export function isDocumentStructureCommand(command: { id: string }): command is DocumentStructureCommand {
  return ids.has(command.id as DocumentStructureCommand['id'])
}

export function documentStructuredBlockAtSelection(model: DocumentModel, selection: DocumentSelection, type?: DocumentBlock['type']): DocumentBlock | undefined {
  const id = selection.tableSelection?.tableId ?? selection.focus.cellBlockId ?? selection.focus.blockId
  return documentBlockAncestry(model, id)?.reverse().find(block => isDocumentStructuredBlock(block) && (type === undefined || block.type === type))
}

interface BlockRange { container: DocumentBlock[]; blocks: DocumentBlock[]; start: number; end: number }
function blocksForSelection(model: DocumentModel, selection: DocumentSelection, blockIds?: readonly string[]): BlockRange {
  if (blockIds && (!blockIds.length || new Set(blockIds).size !== blockIds.length)) throw new RangeError('高级块范围必须包含唯一的块身份。')
  let firstId: string, lastId: string
  if (blockIds) { firstId = blockIds[0]!; lastId = blockIds.at(-1)! }
  else if (selection.tableSelection) {
    if (selection.tableSelection.kind !== 'table') throw new RangeError('请选择完整表格后再包裹高级块。')
    firstId = lastId = selection.tableSelection.tableId
  } else {
    const { start, end } = orderedSelection(model, selection)
    firstId = start.cellBlockId ?? start.blockId; lastId = end.cellBlockId ?? end.blockId
    if (firstId !== lastId && end.offset === 0) {
      const targets = textTargets(model), position = targets.findIndex(target => (target.cellBlockId ?? target.blockId) === lastId)
      if (position > 0) lastId = targets[position - 1]!.cellBlockId ?? targets[position - 1]!.blockId
    }
  }
  const container = findDocumentBlockContainer(model, firstId), start = container?.findIndex(block => block.id === firstId) ?? -1
  const end = container?.findIndex(block => block.id === lastId) ?? -1
  if (!container || start < 0 || end < start || findDocumentBlockContainer(model, lastId) !== container) throw new RangeError('高级块不能切开不同内容容器。')
  const blocks = container.slice(start, end + 1)
  if (blockIds && (blocks.length !== blockIds.length || blocks.some((block, index) => block.id !== blockIds[index]))) throw new RangeError('高级块范围必须在一个内容容器中连续排列。')
  if (blocks.some(block => block.type === 'sectionBreak')) throw new RangeError('高级块不能包含独立分节。')
  return { container, blocks, start, end }
}

function freshParagraph(model: DocumentModel, text = '', code = false): DocumentBlock {
  const defaults = documentDefaultStyle(model), paragraph = createParagraph(text, { ...defaults.text, ...(code ? { fontFamily: 'Consolas' } : {}) })
  paragraph.style = { ...defaults.paragraph, ...(code ? { lineHeight: 1.35, spaceBefore: 0, spaceAfter: 0 } : {}) }
  return paragraph
}
function structureId(model: DocumentModel, prefix: string, requested?: string): string {
  if (requested === undefined) return documentId(prefix)
  if (typeof requested !== 'string' || !requested.trim() || requested.length > 4096 || findDocumentBlock(model, requested)) throw new RangeError('高级块身份无效或已经存在。')
  let columnExists = false
  visitDocumentBlocks(model.blocks, block => { if (block.columns?.columns.some(column => column.id === requested)) columnExists = true })
  if (columnExists) throw new RangeError('栏身份已经存在。')
  return requested
}
function createContainer(model: DocumentModel, type: DocumentDecoratedBlockType, children: DocumentBlock[], callout?: Partial<DocumentCalloutConfig>, code?: Partial<DocumentCodeConfig>, blockId?: string): DocumentBlock {
  const block = freshParagraph(model); block.id = structureId(model, type, blockId); block.type = type; block.runs = []; block.children = children
  block.style = { ...block.style, spaceBefore: 8, spaceAfter: 8 }
  if (type === 'callout') { block.callout = { ...defaultDocumentCalloutConfig(), ...callout }; validateDocumentCalloutConfig(block.callout) }
  if (type === 'code') { block.code = { ...defaultDocumentCodeConfig(), ...code }; validateDocumentCodeConfig(block.code) }
  return block
}

function createColumns(model: DocumentModel, command: DocumentColumnGroupCreateOptions, firstBlocks?: DocumentBlock[]): DocumentBlock {
  const count = command.columnCount ?? command.widthRatios?.length ?? 2
  if (!Number.isSafeInteger(count) || count < DOCUMENT_COLUMN_GROUP_MIN_COLUMNS || count > DOCUMENT_COLUMN_GROUP_MAX_COLUMNS
    || command.widthRatios && command.widthRatios.length !== count || command.minWidths && command.minWidths.length !== count
    || command.columnIds && (command.columnIds.length !== count || new Set(command.columnIds).size !== count)) throw new RangeError('多栏组必须包含 2 至 5 栏及对应的比例。')
  const block = freshParagraph(model); block.id = structureId(model, 'columns', command.columnGroupId); block.type = 'columns'; block.runs = []
  block.columns = { version: 1, gap: command.gap ?? 16, layout: command.layout ?? 'fixed', responsive: command.responsive ?? 'shrink', columns: Array.from({ length: count }, (_, index) => ({
    id: structureId(model, 'column', command.columnIds?.[index]), widthRatio: command.widthRatios?.[index] ?? 1,
    ...(command.minWidths?.[index] == null ? {} : { minWidth: command.minWidths[index]! }), blocks: index === 0 && firstBlocks ? firstBlocks : [freshParagraph(model)] })) }
  if (command.horizontalPadding !== undefined) block.columns.horizontalPadding = command.horizontalPadding
  if (command.minHeight !== undefined) block.columns.minHeight = command.minHeight
  validateDocumentColumnGroup(block.columns)
  if (block.columns.columns.some(column => column.id === block.id)) throw new RangeError('多栏组与栏不能共享身份。')
  return block
}

function setStructuredTextColor(model: DocumentModel, wrapper: DocumentBlock, value: string | null | undefined): void {
  if (value === undefined) return
  const patch = value === null ? documentRunFormatPatch({}, ['color']) : { color: value }
  validateDocumentRunPatch(patch)
  const owned = new Set(documentTextFieldParagraphs({ blocks: wrapper.children ?? wrapper.columns?.columns.flatMap(column => column.blocks) ?? [wrapper] }).map(paragraph => paragraph.block.id))
  for (const target of textTargets(model)) {
    const id = target.cellBlockId ?? target.blockId
    if (!owned.has(id)) continue
    const block = findDocumentBlock(model, id)!
    assertDocumentTextSelectionEditable(model, { anchor: { ...target, offset: 0 }, focus: { ...target, offset: runText(block.runs).length } })
    block.runs = block.runs.map(run => applyRunPatch(run, patch))
  }
}
function setQuoteStyle(model: DocumentModel, block: DocumentBlock, patch: Partial<DocumentQuoteStyle>): void {
  if (block.type !== 'quote') throw new TypeError('当前块不是引用块。')
  if (patch.lineColor !== undefined) {
    validateDocumentRunPatch({ color: patch.lineColor })
    block.metadata.quoteStyle = { version: 1, lineColor: patch.lineColor }
  }
  setStructuredTextColor(model, block, patch.textColor)
}

function remapControlsForWrap(model: DocumentModel, range: BlockRange, wrapper: DocumentBlock): void {
  const selected = new Set(range.blocks.map(block => block.id))
  for (const control of model.contentControls ?? []) {
    if (!control.blockRange) continue
    const owned = resolveDocumentContentControlBlocks(model, control.blockRange)
    if (owned.container !== range.container) continue
    const overlaps = owned.blockIds.some(id => selected.has(id))
    if (!overlaps) continue
    if (owned.startIndex <= range.start && owned.endIndex >= range.end) {
      const ids = [...owned.blockIds], at = ids.indexOf(range.blocks[0]!.id)
      ids.splice(at, range.blocks.length, wrapper.id); control.blockRange = { blockIds: ids }
    } else if (owned.startIndex < range.start || owned.endIndex > range.end) throw new RangeError('高级块不能切开内容控件的完整块范围。')
  }
}
function remapControlsForUnwrap(model: DocumentModel, wrapper: DocumentBlock, children: DocumentBlock[]): void {
  for (const control of model.contentControls ?? []) if (control.blockRange?.blockIds.includes(wrapper.id)) {
    const ids = [...control.blockRange.blockIds], at = ids.indexOf(wrapper.id)
    ids.splice(at, 1, ...children.map(block => block.id)); control.blockRange = { blockIds: ids }
  }
}
function insertionContainer(model: DocumentModel, selection: DocumentSelection, afterBlockId?: string): { container: DocumentBlock[]; index: number } {
  const id = afterBlockId ?? selection.tableSelection?.tableId ?? selection.focus.cellBlockId ?? selection.focus.blockId
  const container = findDocumentBlockContainer(model, id), index = container?.findIndex(block => block.id === id) ?? -1
  if (!container || index < 0) throw new RangeError('高级块插入位置已不存在。')
  return { container, index: index + 1 }
}
function selectionInBlock(model: DocumentModel, block: DocumentBlock): DocumentSelection {
  const owned = new Set<string>(), visit = (blocks: DocumentBlock[]) => { for (const block of blocks) { owned.add(block.id); for (const child of documentBlockChildContainers(block)) visit(child) } }
  visit([block])
  const target = textTargets(model).find(target => owned.has(target.cellBlockId ?? target.blockId))
  if (!target) throw new RangeError('高级块缺少可编辑的输入段落。')
  return collapsedSelection({ ...target, offset: 0 })
}

/** Complete wrappers are reconciled against the original model after the whole deletion. */
function removeBlockResources(model: DocumentModel, blocks: readonly DocumentBlock[]): Set<string> {
  const removed = new Set<string>()
  const visit = (content: readonly DocumentBlock[]) => {
    for (const block of content) {
      removed.add(block.id)
      if (documentTextFields(block).some(documentFieldLocked)) throw new RangeError('删除高级块不能移除锁定的字段。')
      for (const child of documentBlockChildContainers(block)) visit(child)
    }
  }
  visit(blocks)
  for (const control of model.contentControls ?? []) {
    if (control.blockRange || !removed.has(control.tableRowRange?.tableId ?? control.blockId)) continue
    if (control.lock && control.lock !== 'unlocked') throw new RangeError('删除高级块不能移除锁定的内容控件。')
  }
  model.contentControls = model.contentControls?.filter(control => !!control.blockRange || !removed.has(control.tableRowRange?.tableId ?? control.blockId))
  for (const [id, comment] of Object.entries(model.comments ?? {})) {
    comment.anchor = comment.anchor.filter(anchor => !removed.has(anchor.blockId))
    if (!comment.anchor.length) delete model.comments![id]
    else comment.referenceText = comment.anchor.map(anchor => anchor.text).join('\n')
  }
  return removed
}

function selectionAfterRemoval(model: DocumentModel, beforeTargets: ReturnType<typeof textTargets>, removed: ReadonlySet<string>,
  selection: DocumentSelection, replacement?: DocumentBlock): DocumentSelection {
  if (!removed.has(selection.anchor.cellBlockId ?? selection.anchor.blockId)
    && !removed.has(selection.focus.cellBlockId ?? selection.focus.blockId)
    && !removed.has(selection.tableSelection?.tableId ?? '')) return selection
  if (replacement) return selectionInBlock(model, replacement)
  const first = beforeTargets.findIndex(target => removed.has(target.cellBlockId ?? target.blockId))
  const following = beforeTargets.slice(Math.max(0, first)).find(target => !removed.has(target.cellBlockId ?? target.blockId))
  if (following) return collapsedSelection(clampPoint(model, { ...following, offset: 0 }))
  const preceding = beforeTargets.slice(0, Math.max(0, first)).at(-1)
  if (preceding) return collapsedSelection(clampPoint(model, { ...preceding, offset: Number.MAX_SAFE_INTEGER }))
  const target = textTargets(model)[0]
  if (!target) throw new RangeError('删除高级块后缺少输入段落。')
  return collapsedSelection({ ...target, offset: 0 })
}

/** One candidate covers content, controls, fields, permissions and the resulting input point. */
export function executeDocumentStructureCommand(model: DocumentModel, selection: DocumentSelection, command: DocumentStructureCommand,
  context: DocumentPermissionContext): DocumentSelection {
  if (command.id === 'doc.command.update-code-block') return updateDocumentStructuredCodeBlock(model, selection, command, context).selection
  if (command.id === 'doc.command.indent-code-block') return indentCodeBlockWithSelection(model, command.blockId, selection, command.direction, context)
  const candidate = cloneModel(model), currentSelection = cloneModel(selection)
  const removedIds = new Set<string>()
  const affectedIds = new Set<string>()
  let nextSelection = currentSelection
  if (command.id === 'doc.command.wrap-structured-block' || command.id === 'doc.command.wrap-column-group') {
    const range = blocksForSelection(candidate, currentSelection, command.blockIds)
    if (!currentSelection.tableSelection) assertDocumentTextSelectionEditable(model, currentSelection)
    if (command.id === 'doc.command.wrap-structured-block' && command.blockType === 'code' && range.blocks.some(block => !isDocumentTextBlock(block))) throw new RangeError('代码块需要完整的文本段落。')
    const wrapper = command.id === 'doc.command.wrap-column-group' ? createColumns(candidate, command, range.blocks)
      : createContainer(candidate, command.blockType, range.blocks, command.callout, command.code, command.blockId)
    remapControlsForWrap(candidate, range, wrapper)
    range.container.splice(range.start, range.blocks.length, wrapper)
    for (const block of range.blocks) affectedIds.add(block.id)
    affectedIds.add(wrapper.id)
    if (command.id === 'doc.command.wrap-structured-block') {
      if (command.quote) setQuoteStyle(candidate, wrapper, command.quote)
      setStructuredTextColor(candidate, wrapper, command.textColor)
    }
  } else if (command.id === 'doc.command.insert-structured-block' || command.id === 'doc.command.insert-column-group') {
    const block = command.id === 'doc.command.insert-column-group' ? createColumns(candidate, command)
      : createContainer(candidate, command.blockType, (command.source ?? '').replace(/\r\n?/g, '\n').split('\n').map(text => freshParagraph(candidate, text, command.blockType === 'code')), command.callout, command.code, command.blockId)
    if (command.at) {
      assertDocumentTextSelectionEditable(candidate, collapsedSelection(command.at))
      insertDocumentBlocksAtPoint(candidate, command.at, [block]); affectedIds.add(command.at.cellBlockId ?? command.at.blockId)
    } else {
      const at = insertionContainer(candidate, currentSelection, command.afterBlockId)
      affectedIds.add(at.container[at.index - 1]!.id); at.container.splice(at.index, 0, block)
    }
    nextSelection = selectionInBlock(candidate, block); affectedIds.add(block.id)
    if (command.id === 'doc.command.insert-structured-block') {
      if (command.quote) setQuoteStyle(candidate, block, command.quote)
      setStructuredTextColor(candidate, block, command.textColor)
    }
  } else {
    const block = findDocumentBlock(candidate, command.blockId)
    const legacy = block && !isDocumentStructuredBlock(block) && (block.type === 'quote' || !!codeBlockData(block))
    if (!block || !isDocumentStructuredBlock(block) && !legacy) throw new RangeError('高级块已不存在。')
    affectedIds.add(block.id)
    if (command.id === 'doc.command.unwrap-structured-block') {
      if (block.type === 'quote') clearDocumentQuoteParagraphBorders(block)
      if (legacy) { block.type = 'paragraph'; delete block.metadata.codeBlock; delete block.metadata.quoteStyle }
      else {
        const container = findDocumentBlockContainer(candidate, block.id)!, children = block.children ?? block.columns!.columns.flatMap(column => column.blocks)
        if (!children.length) children.push(freshParagraph(candidate))
        remapControlsForUnwrap(candidate, block, children)
        container.splice(container.indexOf(block), 1, ...children)
        for (const child of children) affectedIds.add(child.id)
      }
    } else if (command.id === 'doc.command.remove-structured-block') {
      const targets = textTargets(candidate), container = findDocumentBlockContainer(candidate, block.id)!, at = container.indexOf(block)
      const removed = removeBlockResources(candidate, [block])
      for (const id of removed) removedIds.add(id)
      container.splice(at, 1)
      let replacement: DocumentBlock | undefined
      if (!documentTextFieldParagraphs({ blocks: container }).length) {
        replacement = freshParagraph(candidate); container.splice(Math.min(at, container.length), 0, replacement)
      }
      nextSelection = selectionAfterRemoval(candidate, targets, removed, currentSelection, replacement)
    } else if (command.id === 'doc.command.exit-structured-block') {
      const container = findDocumentBlockContainer(candidate, block.id)!, paragraph = freshParagraph(candidate)
      container.splice(container.indexOf(block) + 1, 0, paragraph); nextSelection = selectionInBlock(candidate, paragraph)
    } else if (command.id === 'doc.command.update-callout') {
      if (block.type !== 'callout') throw new TypeError('当前块不是 Callout。')
      block.callout = { ...block.callout!, ...command.patch }; validateDocumentCalloutConfig(block.callout)
      setStructuredTextColor(candidate, block, command.textColor)
    } else if (command.id === 'doc.command.update-quote-style') {
      setQuoteStyle(candidate, block, command.patch)
    } else if (command.id === 'doc.command.update-code-settings') {
      if (block.type !== 'code') throw new TypeError('当前块不是代码块。')
      block.code = { ...block.code!, ...command.patch }; validateDocumentCodeConfig(block.code)
    } else {
      if (block.type !== 'columns' || !block.columns) throw new TypeError('当前块不是多栏组。')
      const group = block.columns
      if (command.id === 'doc.command.update-column-group') {
        Object.assign(group, command.patch ?? {})
        if (command.widthRatios && command.widthRatios.length !== group.columns.length || command.minWidths && command.minWidths.length !== group.columns.length) throw new RangeError('栏宽比例和最小宽度需要与当前栏数一致。')
        for (const [index, column] of group.columns.entries()) {
          if (command.widthRatios) column.widthRatio = command.widthRatios[index]!
          if (command.minWidths) { if (command.minWidths[index] === null) delete column.minWidth; else column.minWidth = command.minWidths[index]! }
        }
      } else if (command.id === 'doc.command.add-column') {
        if (group.columns.length >= DOCUMENT_COLUMN_GROUP_MAX_COLUMNS) throw new RangeError('多栏组最多包含 5 栏。')
        const at = command.columnId === undefined ? group.columns.length - 1 : group.columns.findIndex(column => column.id === command.columnId)
        if (at < 0) throw new RangeError('目标栏已不存在。')
        group.columns.splice(at + (command.position === 'before' ? 0 : 1), 0, { id: structureId(candidate, 'column', command.newColumnId), widthRatio: 1, blocks: [freshParagraph(candidate)] })
      } else if (command.id === 'doc.command.remove-column') {
        const at = group.columns.findIndex(column => column.id === command.columnId)
        if (at < 0 || group.columns.length <= DOCUMENT_COLUMN_GROUP_MIN_COLUMNS) throw new RangeError('多栏组必须保留至少两栏；可以移除整个多栏容器。')
        const source = group.columns[at]!
        if (command.keepContents !== false) {
          const destination = group.columns[at > 0 ? at - 1 : 1]!
          if (at > 0) destination.blocks.push(...source.blocks)
          else destination.blocks.unshift(...source.blocks)
        } else {
          const removed = removeBlockResources(candidate, source.blocks), targets = textTargets(candidate)
          for (const id of removed) removedIds.add(id)
          group.columns.splice(at, 1)
          nextSelection = selectionAfterRemoval(candidate, targets, removed, currentSelection)
        }
        if (command.keepContents !== false) group.columns.splice(at, 1)
      } else if (command.id === 'doc.command.move-column') {
        const at = group.columns.findIndex(column => column.id === command.columnId)
        if (at < 0 || !Number.isSafeInteger(command.index) || command.index < 0 || command.index >= group.columns.length) throw new RangeError('栏移动位置无效。')
        const column = group.columns.splice(at, 1)[0]!; group.columns.splice(command.index, 0, column)
      } else if (command.id === 'doc.command.set-column-text') {
        if (typeof command.text !== 'string') throw new TypeError('栏内容必须为字符串。')
        const column = group.columns.find(column => column.id === command.columnId)
        if (!column) throw new RangeError('目标栏已不存在。')
        const text = command.text.replace(/\r\n?/g, '\n')
        if (column.blocks.every(isDocumentTextBlock) && column.blocks.map(block => runText(block.runs)).join('\n') === text) return currentSelection
        const targets = textTargets(candidate), removed = removeBlockResources(candidate, column.blocks)
        for (const id of removed) removedIds.add(id)
        column.blocks = text.split('\n').map(value => freshParagraph(candidate, value))
        nextSelection = selectionAfterRemoval(candidate, targets, removed, currentSelection, column.blocks[0])
      } else if (command.id === 'doc.command.insert-column-paragraph') {
        const column = group.columns.find(column => column.id === command.columnId)
        if (!column) throw new RangeError('目标栏已不存在。')
        const paragraphs = documentTextFieldParagraphs({ blocks: column.blocks })
        if (!Number.isSafeInteger(command.index) || command.index < 0 || command.index > paragraphs.length
          || command.text !== undefined && typeof command.text !== 'string') throw new RangeError('栏内段落位置无效。')
        const inserted = (command.text ?? '').replace(/\r\n?/g, '\n').split('\n').map(text => freshParagraph(candidate, text))
        if (command.paragraphId !== undefined) inserted[0]!.id = structureId(candidate, 'paragraph', command.paragraphId)
        const before = paragraphs[command.index]?.block
        if (before) {
          const container = findDocumentBlockContainer(candidate, before.id)!
          affectedIds.add(before.id); container.splice(container.indexOf(before), 0, ...inserted)
        } else column.blocks.push(...inserted)
        for (const paragraph of inserted) affectedIds.add(paragraph.id)
      }
      validateDocumentColumnGroup(group)
    }
  }
  removeDeletedDocumentRepeatingNoteScopes(model, candidate)
  if (removedIds.size) reconcileDocumentNotesAfterStructure(candidate, removedIds)
  const paragraphs = documentTextFieldParagraphs(candidate)
  const originalFields = new Map(documentTextFieldParagraphs(model).flatMap(paragraph => documentTextFields(paragraph.block).map(field => [field.id, { owner: paragraph.block.id, field }] as const)))
  const transformedFields = new Set(paragraphs.flatMap(paragraph => documentTextFields(paragraph.block).flatMap(field => {
    const original = originalFields.get(field.id)
    return original && (original.owner !== paragraph.block.id || JSON.stringify(original.field) !== JSON.stringify(field)) ? [field.id] : []
  })))
  const fields = planDocumentTableFieldAnchors(model, paragraphs, transformedFields)
  for (const paragraph of paragraphs) if (paragraph.block.metadata.documentTextFields !== undefined || fields.has(paragraph.block.id)) {
    setDocumentTextFields(paragraph.block, fields.get(paragraph.block.id) ?? [])
  }
  reconcileDocumentTocFieldParents(model, candidate); documentTextFieldRanges(candidate)
  for (const id of affectedIds) for (const ancestor of [...(documentBlockAncestry(model, id) ?? []), ...(documentBlockAncestry(candidate, id) ?? [])]) {
    const current = findDocumentBlock(candidate, ancestor.id)
    if (current) delete current.metadata.html
  }
  reconcileDocumentContentControlBlockEdits(model, candidate, { selection: currentSelection, absorbAdjacentInsertions: false })
  validateDocumentContentControlRanges(candidate); validateDocumentStoryBlocks(candidate.blocks)
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, candidate))
  candidate.wordCount = textTargets(candidate).reduce((count, target) => count + runText(runsForTarget(candidate, target)).replace(/\s/g, '').length, 0)
  const tableSelection = rebaseDocumentTableViewSelection(model, candidate, nextSelection.tableSelection, removedIds.size > 0)
  Object.assign(model, candidate)
  return { anchor: clampPoint(candidate, nextSelection.anchor), focus: clampPoint(candidate, nextSelection.focus),
    ...(tableSelection ? { tableSelection } : {}) }
}
