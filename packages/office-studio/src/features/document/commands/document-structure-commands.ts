import type { DocumentBlock, DocumentCalloutConfig, DocumentCodeConfig, DocumentColumnGroup, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { clampPoint, collapsedSelection, createParagraph, documentBlockAncestry, documentId, findDocumentBlock,
  findDocumentBlockContainer, orderedSelection, runText, runsForTarget, textTargets, type DocumentSelection } from '../model/document-model'
import { defaultDocumentCalloutConfig, defaultDocumentCodeConfig, documentBlockChildContainers, isDocumentStructuredBlock,
  isDocumentTextBlock, validateDocumentCalloutConfig, validateDocumentCodeConfig, validateDocumentColumnGroup, type DocumentDecoratedBlockType } from '../model/document-structural-blocks'
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

export type DocumentStructureCommand =
  | { id: 'doc.command.wrap-structured-block'; blockType: DocumentDecoratedBlockType; blockIds?: string[]; callout?: Partial<DocumentCalloutConfig>; code?: Partial<DocumentCodeConfig> }
  | { id: 'doc.command.insert-structured-block'; blockType: DocumentDecoratedBlockType; afterBlockId?: string; source?: string; callout?: Partial<DocumentCalloutConfig>; code?: Partial<DocumentCodeConfig> }
  | { id: 'doc.command.unwrap-structured-block'; blockId: string }
  | { id: 'doc.command.exit-structured-block'; blockId: string }
  | { id: 'doc.command.update-callout'; blockId: string; patch: Partial<DocumentCalloutConfig> }
  | { id: 'doc.command.update-code-settings'; blockId: string; patch: Partial<DocumentCodeConfig> }
  | { id: 'doc.command.update-code-block'; blockId: string; source?: string; patch?: Partial<DocumentCodeConfig>; host?: DocumentAuxiliaryBlockHost }
  | { id: 'doc.command.insert-column-group'; columnCount?: number; afterBlockId?: string; gap?: number; widthRatios?: number[] }
  | { id: 'doc.command.wrap-column-group'; columnCount?: number; blockIds?: string[]; gap?: number; widthRatios?: number[] }
  | { id: 'doc.command.update-column-group'; blockId: string; patch?: Partial<Pick<DocumentColumnGroup, 'gap' | 'layout' | 'responsive'>>; widthRatios?: number[]; minWidths?: Array<number | null> }
  | { id: 'doc.command.add-column'; blockId: string; columnId?: string; position?: 'before' | 'after' }
  | { id: 'doc.command.remove-column'; blockId: string; columnId: string; keepContents?: boolean }
  | { id: 'doc.command.move-column'; blockId: string; columnId: string; index: number }

const ids = new Set<DocumentStructureCommand['id']>(['doc.command.wrap-structured-block', 'doc.command.insert-structured-block',
  'doc.command.unwrap-structured-block', 'doc.command.exit-structured-block', 'doc.command.update-callout', 'doc.command.update-code-settings', 'doc.command.update-code-block',
  'doc.command.insert-column-group', 'doc.command.wrap-column-group', 'doc.command.update-column-group', 'doc.command.add-column',
  'doc.command.remove-column', 'doc.command.move-column'])
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
function createContainer(model: DocumentModel, type: DocumentDecoratedBlockType, children: DocumentBlock[], callout?: Partial<DocumentCalloutConfig>, code?: Partial<DocumentCodeConfig>): DocumentBlock {
  const block = freshParagraph(model); block.id = documentId(type); block.type = type; block.runs = []; block.children = children
  block.style = { ...block.style, spaceBefore: 8, spaceAfter: 8 }
  if (type === 'callout') { block.callout = { ...defaultDocumentCalloutConfig(), ...callout }; validateDocumentCalloutConfig(block.callout) }
  if (type === 'code') { block.code = { ...defaultDocumentCodeConfig(), ...code }; validateDocumentCodeConfig(block.code) }
  return block
}

function createColumns(model: DocumentModel, command: { columnCount?: number; gap?: number; widthRatios?: number[] }, firstBlocks?: DocumentBlock[]): DocumentBlock {
  const count = command.columnCount ?? command.widthRatios?.length ?? 2
  if (!Number.isSafeInteger(count) || count < 1 || count > 12 || command.widthRatios && command.widthRatios.length !== count) throw new RangeError('多栏组必须包含 1 至 12 栏及对应的比例。')
  const block = freshParagraph(model); block.id = documentId('columns'); block.type = 'columns'; block.runs = []
  block.columns = { version: 1, gap: command.gap ?? 16, layout: 'fixed', responsive: 'shrink', columns: Array.from({ length: count }, (_, index) => ({
    id: documentId('column'), widthRatio: command.widthRatios?.[index] ?? 1, blocks: index === 0 && firstBlocks ? firstBlocks : [freshParagraph(model)] })) }
  validateDocumentColumnGroup(block.columns)
  return block
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

/** One candidate covers content, controls, fields, permissions and the resulting input point. */
export function executeDocumentStructureCommand(model: DocumentModel, selection: DocumentSelection, command: DocumentStructureCommand,
  context: DocumentPermissionContext): DocumentSelection {
  if (command.id === 'doc.command.update-code-block') return updateDocumentStructuredCodeBlock(model, selection, command, context).selection
  const candidate = cloneModel(model), currentSelection = cloneModel(selection)
  let nextSelection = currentSelection
  if (command.id === 'doc.command.wrap-structured-block' || command.id === 'doc.command.wrap-column-group') {
    const range = blocksForSelection(candidate, currentSelection, command.blockIds)
    if (!currentSelection.tableSelection) assertDocumentTextSelectionEditable(model, currentSelection)
    if (command.id === 'doc.command.wrap-structured-block' && command.blockType === 'code' && range.blocks.some(block => !isDocumentTextBlock(block))) throw new RangeError('代码块需要完整的文本段落。')
    const wrapper = command.id === 'doc.command.wrap-column-group' ? createColumns(candidate, command, range.blocks)
      : createContainer(candidate, command.blockType, range.blocks, command.callout, command.code)
    remapControlsForWrap(candidate, range, wrapper)
    range.container.splice(range.start, range.blocks.length, wrapper)
  } else if (command.id === 'doc.command.insert-structured-block' || command.id === 'doc.command.insert-column-group') {
    const at = insertionContainer(candidate, currentSelection, command.afterBlockId)
    const block = command.id === 'doc.command.insert-column-group' ? createColumns(candidate, command)
      : createContainer(candidate, command.blockType, (command.source ?? '').replace(/\r\n?/g, '\n').split('\n').map(text => freshParagraph(candidate, text, command.blockType === 'code')), command.callout, command.code)
    at.container.splice(at.index, 0, block); nextSelection = selectionInBlock(candidate, block)
  } else {
    const block = findDocumentBlock(candidate, command.blockId)
    if (!block || !isDocumentStructuredBlock(block)) throw new RangeError('高级块已不存在。')
    if (command.id === 'doc.command.unwrap-structured-block') {
      const container = findDocumentBlockContainer(candidate, block.id)!, children = block.children ?? block.columns!.columns.flatMap(column => column.blocks)
      if (!children.length) children.push(freshParagraph(candidate))
      remapControlsForUnwrap(candidate, block, children)
      container.splice(container.indexOf(block), 1, ...children)
    } else if (command.id === 'doc.command.exit-structured-block') {
      const container = findDocumentBlockContainer(candidate, block.id)!, paragraph = freshParagraph(candidate)
      container.splice(container.indexOf(block) + 1, 0, paragraph); nextSelection = selectionInBlock(candidate, paragraph)
    } else if (command.id === 'doc.command.update-callout') {
      if (block.type !== 'callout') throw new TypeError('当前块不是 Callout。')
      block.callout = { ...block.callout!, ...command.patch }; validateDocumentCalloutConfig(block.callout)
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
        if (group.columns.length >= 12) throw new RangeError('多栏组最多包含 12 栏。')
        const at = command.columnId === undefined ? group.columns.length - 1 : group.columns.findIndex(column => column.id === command.columnId)
        if (at < 0) throw new RangeError('目标栏已不存在。')
        group.columns.splice(at + (command.position === 'before' ? 0 : 1), 0, { id: documentId('column'), widthRatio: group.columns[at]!.widthRatio, blocks: [freshParagraph(candidate)] })
      } else if (command.id === 'doc.command.remove-column') {
        const at = group.columns.findIndex(column => column.id === command.columnId)
        if (at < 0 || group.columns.length <= 1) throw new RangeError('多栏组必须保留至少一栏。')
        const source = group.columns[at]!
        if (command.keepContents !== false) {
          const destination = group.columns[at > 0 ? at - 1 : 1]!
          if (at > 0) destination.blocks.push(...source.blocks)
          else destination.blocks.unshift(...source.blocks)
        } else {
          const removed = new Set<string>(), visit = (blocks: DocumentBlock[]) => { for (const block of blocks) { removed.add(block.id); for (const child of documentBlockChildContainers(block)) visit(child) } }
          visit(source.blocks)
          for (const control of candidate.contentControls ?? []) if (removed.has(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)
            && control.lock && control.lock !== 'unlocked') throw new RangeError('删除栏不能移除锁定的内容控件。')
          candidate.contentControls = candidate.contentControls?.filter(control => !removed.has(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId))
          for (const [id, comment] of Object.entries(candidate.comments ?? {})) {
            comment.anchor = comment.anchor.filter(anchor => !removed.has(anchor.blockId))
            if (!comment.anchor.length) delete candidate.comments![id]
            else comment.referenceText = comment.anchor.map(anchor => anchor.text).join('\n')
          }
        }
        group.columns.splice(at, 1)
      } else if (command.id === 'doc.command.move-column') {
        const at = group.columns.findIndex(column => column.id === command.columnId)
        if (at < 0 || !Number.isSafeInteger(command.index) || command.index < 0 || command.index >= group.columns.length) throw new RangeError('栏移动位置无效。')
        const column = group.columns.splice(at, 1)[0]!; group.columns.splice(command.index, 0, column)
      }
      validateDocumentColumnGroup(group)
    }
  }
  removeDeletedDocumentRepeatingNoteScopes(model, candidate)
  reconcileDocumentContentControlBlockEdits(model, candidate, { selection: currentSelection, absorbAdjacentInsertions: false })
  validateDocumentContentControlRanges(candidate); validateDocumentStoryBlocks(candidate.blocks)
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, candidate))
  candidate.wordCount = textTargets(candidate).reduce((count, target) => count + runText(runsForTarget(candidate, target)).replace(/\s/g, '').length, 0)
  Object.assign(model, candidate)
  return { anchor: clampPoint(candidate, nextSelection.anchor), focus: clampPoint(candidate, nextSelection.focus),
    ...(nextSelection.tableSelection ? { tableSelection: cloneModel(nextSelection.tableSelection) } : {}) }
}
