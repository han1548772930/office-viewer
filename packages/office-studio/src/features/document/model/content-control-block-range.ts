import type { DocumentBlock, DocumentContentControl, DocumentModel } from '../../../domain/model'
import { runText, runsForTarget, textTargets, type DocumentSelection, type TextTarget } from './document-model'
import { documentBlockChildContainers } from './document-structural-blocks'
import { reanchorDocumentContentControlNoteParagraph } from './content-control-note-boundary'
import { documentNoteReferenceOffsets } from './document-note-boundary'

export type DocumentContentControlBlockRange = NonNullable<DocumentContentControl['blockRange']>
export interface DocumentContentControlBlocks {
  container: DocumentBlock[]
  blocks: DocumentBlock[]
  startIndex: number
  endIndex: number
  blockIds: string[]
}

export function assertDocumentContentControlBlockRange(value: unknown): asserts value is DocumentContentControlBlockRange {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.keys(value).some(key => key !== 'blockIds')) {
    throw new TypeError('内容控件块范围必须包含完整块身份序列。')
  }
  const ids = Reflect.get(value, 'blockIds')
  if (!Array.isArray(ids) || !ids.length || ids.length > 100_000 || ids.some(id => typeof id !== 'string' || !id.trim() || id.length > 4096)
    || new Set(ids).size !== ids.length) throw new TypeError('内容控件块范围身份无效。')
}

/** The caller supplies one editing host. A range never crosses a cell or story boundary. */
export function resolveDocumentContentControlBlocks(model: Pick<DocumentModel, 'blocks'>, range: DocumentContentControlBlockRange): DocumentContentControlBlocks {
  assertDocumentContentControlBlockRange(range)
  const containers: DocumentBlock[][] = []
  const visit = (blocks: DocumentBlock[]) => {
    if (blocks.some(block => block.id === range.blockIds[0])) containers.push(blocks)
    for (const block of blocks) for (const content of documentBlockChildContainers(block)) visit(content)
  }
  visit(model.blocks)
  if (containers.length !== 1) throw new RangeError('内容控件块范围不属于唯一的当前内容容器。')
  const container = containers[0]!, startIndex = container.findIndex(block => block.id === range.blockIds[0])
  const endIndex = startIndex + range.blockIds.length - 1
  if (endIndex >= container.length || range.blockIds.some((id, offset) => container[startIndex + offset]?.id !== id)) {
    throw new RangeError('内容控件块身份不连续、顺序已改变或块已删除。')
  }
  return { container, blocks: container.slice(startIndex, endIndex + 1), startIndex, endIndex, blockIds: [...range.blockIds] }
}

export function documentContentControlOwnedBlockIds(range: DocumentContentControlBlocks): Set<string> {
  const ids = new Set<string>()
  const visit = (blocks: DocumentBlock[]) => {
    for (const block of blocks) {
      ids.add(block.id)
      for (const content of documentBlockChildContainers(block)) visit(content)
    }
  }
  visit(range.blocks)
  return ids
}

export function documentBlockContentControlContains(model: DocumentModel, parent: DocumentContentControl, child: DocumentContentControl): boolean {
  if (!parent.blockRange || parent.id === child.id) return false
  const owned = documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, parent.blockRange))
  if (child.blockRange) return resolveDocumentContentControlBlocks(model, child.blockRange).blockIds.every(id => owned.has(id))
  return owned.has(child.tableRowRange?.tableId ?? child.blockId)
}

/** Full rich structure, including nested tables/continuations, participates in locks and temporary SDTs. */
export function documentContentControlBlockContentSignature(model: DocumentModel, range: DocumentContentControlBlockRange): string {
  const resolved = resolveDocumentContentControlBlocks(model, range)
  return JSON.stringify({ blockIds: resolved.blockIds, blocks: resolved.blocks })
}

export function documentContentControlBlockTextTargets(model: DocumentModel, range: DocumentContentControlBlockRange): TextTarget[] {
  const owned = documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, range))
  return textTargets(model).filter(target => owned.has(target.cellBlockId ?? target.blockId))
}

export function documentContentControlBlocksSelection(model: DocumentModel, range: DocumentContentControlBlockRange): DocumentSelection {
  const targets = documentContentControlBlockTextTargets(model, range), first = targets[0], last = targets.at(-1)
  if (!first || !last) throw new RangeError('内容控件完整块范围没有可用的文字输入锚点。')
  return { anchor: { ...first, offset: 0 }, focus: { ...last, offset: runText(runsForTarget(model, last)).length } }
}

/** Paragraph offsets are a legal input fallback; ownership always uses the complete block range. */
export function reanchorDocumentBlockContentControl(model: DocumentModel, control: DocumentContentControl): void {
  if (!control.blockRange) return
  const target = documentContentControlBlockTextTargets(model, control.blockRange)[0]
  if (!target) throw new RangeError('内容控件块范围没有可用的文字输入宿主。')
  reanchorDocumentContentControlNoteParagraph(control, target.cellBlockId ?? target.blockId, runsForTarget(model, target))
}

/** Structural planners update block IDs first; this validates and restores every resulting fallback atomically. */
export function reconcileDocumentBlockContentControlRanges(model: DocumentModel): void {
  const controls = (model.contentControls ?? []).filter(control => control.blockRange)
  const fallbacks = controls.map(control => {
    if (control.tableRowRange || control.placement !== 'block') throw new TypeError('完整块范围不能同时属于物理表格行。')
    const target = documentContentControlBlockTextTargets(model, control.blockRange!)[0]
    if (!target) throw new RangeError('内容控件块范围没有可用的文字输入宿主。')
    const runs = runsForTarget(model, target)
    return { control, blockId: target.cellBlockId ?? target.blockId, endOffset: runText(runs).length, endRank: documentNoteReferenceOffsets(runs).length }
  })
  for (const { control, blockId, endOffset, endRank } of fallbacks) {
    control.blockId = blockId; control.startOffset = 0; control.endOffset = endOffset
    control.startNoteReferenceCount = 0; control.endNoteReferenceCount = endRank
  }
}
