import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { collapsedSelection, runText, runsForTarget, tableForBlock, textTargets, type DocumentSelection } from './document-model'
import { documentTableCellContentContainers } from './document-table-continuations'
import { documentTableDescendantIds } from './document-table-selection'
import { documentContentControlBlocksSelection, documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from './content-control-block-range'

interface BlockContainerStep { tableId: string; cellId: string; contentIndex: number }

function containerPath(blocks: DocumentBlock[], target: DocumentBlock[], steps: BlockContainerStep[] = []): BlockContainerStep[] | undefined {
  if (blocks === target) return steps
  for (const block of blocks) for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) {
    for (const [contentIndex, content] of documentTableCellContentContainers(cell).entries()) {
      const found = containerPath(content, target, [...steps, { tableId: block.id, cellId: cell.id, contentIndex }])
      if (found) return found
    }
  }
  return undefined
}

function containerAtPath(blocks: DocumentBlock[], steps: readonly BlockContainerStep[]): DocumentBlock[] | undefined {
  let result = blocks
  for (const step of steps) {
    const block = result.find(block => block.id === step.tableId), table = block && tableForBlock(block)
    const cells = table?.rows.flatMap(row => row.cells).filter(cell => cell.id === step.cellId) ?? []
    if (cells.length !== 1) return undefined
    const container = documentTableCellContentContainers(cells[0]!)[step.contentIndex]
    if (!container) return undefined
    result = container
  }
  return result
}

/** Resolve an actual scalar replacement after a temporary block wrapper has disappeared. */
export function documentContentControlBlocksValueMutationRange(before: DocumentModel, after: DocumentModel, id: string):
  { blockIds: string[]; selection: DocumentSelection } | undefined {
  if (before.id !== after.id) return undefined
  try {
    const sourceControl = before.contentControls?.find(control => control.id === id)
    if (!sourceControl?.blockRange) return undefined
    const source = resolveDocumentContentControlBlocks(before, sourceControl.blockRange), path = containerPath(before.blocks, source.container)
    const container = path && containerAtPath(after.blocks, path)
    if (!container) return undefined
    const current = after.contentControls?.find(control => control.id === id)
    const originalIds = documentTableDescendantIds(before.blocks)
    if (current) {
      if (!current.blockRange || current.tableRowRange || current.type !== sourceControl.type || current.placement !== sourceControl.placement
        || current.parentId !== sourceControl.parentId) return undefined
      const range = resolveDocumentContentControlBlocks(after, current.blockRange)
      if (range.container !== container || range.blockIds.some(blockId => originalIds.has(blockId) && !source.blockIds.includes(blockId))) return undefined
      // Other binding peers may replace adjacent blocks in the same transaction. The surviving
      // wrapper is the authoritative live boundary; their new IDs must not invalidate this one.
      return { blockIds: [...range.blockIds], selection: documentContentControlBlocksSelection(after, current.blockRange) }
    }
    const prefix = source.container.slice(0, source.startIndex).map(block => block.id), suffix = source.container.slice(source.endIndex + 1).map(block => block.id)
    const ids = container.slice(prefix.length, container.length - suffix.length).map(block => block.id)
    if (!ids.length || JSON.stringify(container.map(block => block.id)) !== JSON.stringify([...prefix, ...ids, ...suffix])) return undefined
    if (ids.some(id => originalIds.has(id) && !source.blockIds.includes(id))) return undefined
    const blockRange = { blockIds: [...ids] }, range = resolveDocumentContentControlBlocks(after, blockRange)
    if (range.container !== container || range.startIndex !== source.startIndex) return undefined
    if (!sourceControl.temporary || ['controlLocked', 'sdtLocked', 'sdtContentLocked'].includes(sourceControl.lock ?? '')) return undefined
    return { blockIds: [...ids], selection: documentContentControlBlocksSelection(after, blockRange) }
  } catch (error) {
    if (error instanceof TypeError || error instanceof RangeError) return undefined
    throw error
  }
}

/** Strictly resolve the committed block sequence and newly allocated item; no stale paragraph guess. */
export function documentContentControlBlocksMutationSelection(before: DocumentModel, after: DocumentModel, id: string,
  action: 'duplicate' | 'delete'): { controlId?: string; selection: DocumentSelection } | undefined {
  if (before.id !== after.id) return undefined
  try {
    const item = before.contentControls?.find(control => control.id === id)
    if (item?.type !== 'repeatingSectionItem' || item.placement !== 'block' || !item.blockRange || item.tableRowRange) return undefined
    const source = resolveDocumentContentControlBlocks(before, item.blockRange), path = containerPath(before.blocks, source.container)
    const container = path && containerAtPath(after.blocks, path)
    if (!container) return undefined
    const previousIds = source.container.map(block => block.id), oldBlockIds = documentTableDescendantIds(before.blocks)
    if (action === 'duplicate') {
      const original = after.contentControls?.find(control => control.id === id)
      if (!original || original.type !== item.type || original.placement !== item.placement || original.parentId !== item.parentId
        || JSON.stringify(original.blockRange) !== JSON.stringify(item.blockRange)) return undefined
      const oldIds = new Set(before.contentControls?.map(control => control.id))
      const copies = (after.contentControls ?? []).filter(control => !oldIds.has(control.id) && control.type === item.type
        && control.parentId === item.parentId && control.placement === item.placement && !control.tableRowRange
        && control.blockRange?.blockIds.length === source.blockIds.length && control.blockRange.blockIds.every(id => !oldBlockIds.has(id)))
        .filter(control => {
          const range = resolveDocumentContentControlBlocks(after, control.blockRange!)
          return range.container === container && range.startIndex === source.endIndex + 1
            && JSON.stringify(container.map(block => block.id)) === JSON.stringify([
              ...previousIds.slice(0, source.endIndex + 1), ...range.blockIds, ...previousIds.slice(source.endIndex + 1),
            ])
        })
      if (copies.length !== 1) return undefined
      return { controlId: copies[0]!.id, selection: documentContentControlBlocksSelection(after, copies[0]!.blockRange!) }
    }
    if (after.contentControls?.some(control => control.id === id)) return undefined
    const removed = new Set(source.blockIds), expected = previousIds.filter(id => !removed.has(id))
    const actual = container.map(block => block.id), newIds = actual.filter(id => !oldBlockIds.has(id))
    if (newIds.length > 1 || JSON.stringify(actual.filter(id => oldBlockIds.has(id))) !== JSON.stringify(expected)) return undefined
    if (newIds.length && (actual.indexOf(newIds[0]!) !== Math.min(source.startIndex, expected.length)
      || !textTargets({ ...after, blocks: container.filter(block => newIds.includes(block.id)) }).length)) return undefined
    const targets = textTargets(after), within = documentTableDescendantIds(container)
    const boundary = Math.min(source.startIndex, container.length)
    const following = documentTableDescendantIds(container.slice(boundary)), preceding = documentTableDescendantIds(container.slice(0, boundary))
    const next = targets.find(target => following.has(target.cellBlockId ?? target.blockId))
    if (next) return { selection: collapsedSelection({ ...next, offset: 0 }) }
    const previous = targets.filter(target => preceding.has(target.cellBlockId ?? target.blockId)).at(-1)
    if (previous) return { selection: collapsedSelection({ ...previous, offset: runText(runsForTarget(after, previous)).length }) }
    const owned = documentContentControlOwnedBlockIds(source), oldTargets = textTargets(before)
    const first = oldTargets.findIndex(target => owned.has(target.cellBlockId ?? target.blockId))
    const last = oldTargets.reduce((result, target, index) => owned.has(target.cellBlockId ?? target.blockId) ? index : result, -1)
    for (const target of [...oldTargets.slice(last + 1), ...oldTargets.slice(0, first).reverse()]) {
      const current = targets.find(current => (current.cellBlockId ?? current.blockId) === (target.cellBlockId ?? target.blockId))
      if (current && !within.has(current.cellBlockId ?? current.blockId)) return { selection: collapsedSelection({ ...current,
        offset: oldTargets.indexOf(target) > last ? 0 : runText(runsForTarget(after, current)).length }) }
    }
    const fallback = targets[0]
    return fallback ? { selection: collapsedSelection({ ...fallback, offset: 0 }) } : undefined
  } catch (error) {
    if (error instanceof TypeError || error instanceof RangeError) return undefined
    throw error
  }
}
