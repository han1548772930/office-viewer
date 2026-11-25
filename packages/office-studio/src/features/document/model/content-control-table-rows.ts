import type { DocumentBlock, DocumentContentControl, DocumentModel } from '../../../domain/model'
import { documentContentControlGroupEditAllowed } from './content-control-group-lock'
import { collapsedSelection, findDocumentBlock, runText, runsForTarget, sameTarget, tableForBlock, textTargets, type DocumentSelection, type DocumentTable } from './document-model'
import { documentTableGrid, type DocumentTableGrid } from './document-table-grid'
import { documentTableCellBlocks, documentTableContinuationEntries } from './document-table-continuations'
import { createDocumentTableViewSelection, documentTableDescendantIds, resolveDocumentTableViewSelection } from './document-table-selection'
import { isDocumentTextBoundary } from './document-story'
import { documentBlockContentControlContains, documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from './content-control-block-range'
import { documentContentControlNoteBoundaryRanks, reanchorDocumentContentControlNoteParagraph } from './content-control-note-boundary'
import { assertDocumentNoteReferenceBoundaryRank, documentNoteReferenceOffsets } from './document-note-boundary'
import { compareDocumentNoteBoundaries } from './document-note-boundary-edit'

export type DocumentContentControlTableRowRange = NonNullable<DocumentContentControl['tableRowRange']>
export interface DocumentContentControlTableRows {
  block: DocumentBlock
  table: DocumentTable
  grid: DocumentTableGrid
  startRow: number
  endRow: number
  rowIds: string[]
}

/** Row IDs are the persisted scope. Paragraph offsets only provide an input fallback. */
export function assertDocumentContentControlTableRowRange(value: unknown): asserts value is DocumentContentControlTableRowRange {
  if (!value || typeof value !== 'object' || Array.isArray(value)
    || Object.keys(value).some(key => key !== 'tableId' && key !== 'rowIds')) throw new TypeError('内容控件行范围必须包含表格和连续行身份。')
  const tableId = Reflect.get(value, 'tableId'), rowIds = Reflect.get(value, 'rowIds')
  if (typeof tableId !== 'string' || !tableId.trim() || !Array.isArray(rowIds) || !rowIds.length || rowIds.length > 10_000
    || rowIds.some(id => typeof id !== 'string' || !id.trim()) || new Set(rowIds).size !== rowIds.length) {
    throw new TypeError('内容控件行范围身份无效。')
  }
}

export function resolveDocumentContentControlTableRows(model: DocumentModel, range: DocumentContentControlTableRowRange): DocumentContentControlTableRows {
  assertDocumentContentControlTableRowRange(range)
  const block = findDocumentBlock(model, range.tableId), table = block && tableForBlock(block)
  if (!block || !table) throw new RangeError('重复项目所属表格已不存在或已切换内容区。')
  const grid = documentTableGrid(table), startRow = table.rows.findIndex(row => row.id === range.rowIds[0]), endRow = startRow + range.rowIds.length - 1
  if (startRow < 0 || endRow >= grid.rowCount || range.rowIds.some((id, offset) => table.rows[startRow + offset]!.id !== id)) {
    throw new RangeError('内容控件行身份不连续、顺序已改变或行已删除。')
  }
  return { block, table, grid, startRow, endRow, rowIds: [...range.rowIds] }
}

/** Includes complete merged owners and nested/continuation content affected by a row operation. */
export function documentContentControlTableRowBlockIds(rows: DocumentContentControlTableRows): Set<string> {
  const blocks = rows.grid.anchors.filter(owner => owner.row <= rows.endRow && owner.row + owner.rowSpan > rows.startRow)
    .flatMap(owner => documentTableCellBlocks(owner.cell))
  return documentTableDescendantIds(blocks)
}

/** Physical content belongs to a row wrapper; merged content outside its rows is not a copied child. */
export function documentContentControlOwnedRowBlockIds(rows: DocumentContentControlTableRows): Set<string> {
  const selected = new Set(rows.rowIds), blocks: DocumentBlock[] = []
  for (const owner of rows.grid.anchors) {
    if (selected.has(rows.table.rows[owner.row]!.id)) blocks.push(...owner.cell.blocks)
    for (const entry of documentTableContinuationEntries(owner.cell)) if (selected.has(entry.content.rowId)) blocks.push(...entry.content.blocks)
  }
  return documentTableDescendantIds(blocks)
}

export function documentContentControlContains(model: DocumentModel, parent: DocumentContentControl, child: DocumentContentControl): boolean {
  if (parent.id === child.id) return false
  if (parent.blockRange) {
    if (parent.tableRowRange || !findDocumentBlock(model, parent.blockRange.blockIds[0]!)) return false
    if (child.blockRange && !findDocumentBlock(model, child.blockRange.blockIds[0]!)) return false
    return documentBlockContentControlContains(model, parent, child)
  }
  if (!parent.tableRowRange) {
    if (child.tableRowRange || child.blockRange || parent.blockId !== child.blockId) return false
    const block = findDocumentBlock(model, parent.blockId)
    if (!block) return false
    const references = documentNoteReferenceOffsets(block.runs), outer = documentContentControlNoteBoundaryRanks(parent, references), inner = documentContentControlNoteBoundaryRanks(child, references)
    return compareDocumentNoteBoundaries(parent.startOffset, outer.start, child.startOffset, inner.start) <= 0
      && compareDocumentNoteBoundaries(parent.endOffset, outer.end, child.endOffset, inner.end) >= 0
  }
  if (!findDocumentBlock(model, parent.tableRowRange.tableId)) return false
  const rows = resolveDocumentContentControlTableRows(model, parent.tableRowRange)
  if (child.tableRowRange?.tableId === parent.tableRowRange.tableId) {
    const nested = resolveDocumentContentControlTableRows(model, child.tableRowRange)
    return nested.startRow >= rows.startRow && nested.endRow <= rows.endRow
  }
  const owned = documentContentControlOwnedRowBlockIds(rows)
  if (child.blockRange) return !!findDocumentBlock(model, child.blockRange.blockIds[0]!)
    && resolveDocumentContentControlBlocks(model, child.blockRange).blockIds.every(id => owned.has(id))
  return child.tableRowRange ? owned.has(child.tableRowRange.tableId) : owned.has(child.blockId)
}

/** The paragraph range is local input; a row wrapper owns every physical content container in its rows. */
export function documentContentControlContainsBlockRange(model: DocumentModel, control: DocumentContentControl, blockId: string,
  startOffset: number, endOffset: number, startRank?: number, endRank?: number): boolean {
  if (!Number.isSafeInteger(startOffset) || !Number.isSafeInteger(endOffset) || startOffset < 0 || endOffset < startOffset) return false
  const block = findDocumentBlock(model, blockId)
  if (!block || !isDocumentTextBoundary(runText(block.runs), startOffset) || !isDocumentTextBoundary(runText(block.runs), endOffset)) return false
  if (control.blockRange) {
    if (control.tableRowRange || !findDocumentBlock(model, control.blockRange.blockIds[0]!)) return false
    return documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange)).has(blockId)
  }
  if (!control.tableRowRange) {
    if (control.blockId !== blockId) return false
    const references = documentNoteReferenceOffsets(block.runs), outer = documentContentControlNoteBoundaryRanks(control, references)
    const from = assertDocumentNoteReferenceBoundaryRank(references, startOffset, startRank), to = assertDocumentNoteReferenceBoundaryRank(references, endOffset, endRank)
    return compareDocumentNoteBoundaries(control.startOffset, outer.start, startOffset, from) <= 0
      && compareDocumentNoteBoundaries(endOffset, to, control.endOffset, outer.end) <= 0
      && compareDocumentNoteBoundaries(startOffset, from, endOffset, to) <= 0
  }
  // Stored controls from another story/note are outside the current input model.
  if (!findDocumentBlock(model, control.tableRowRange.tableId)) return false
  return documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(model, control.tableRowRange)).has(blockId)
}

/** Used for lock preservation and temporary wrappers; no first-paragraph scalar substitutes for row content. */
export function documentContentControlTableRowContentSignature(model: DocumentModel, range: DocumentContentControlTableRowRange): string {
  const rows = resolveDocumentContentControlTableRows(model, range), selected = new Set(rows.rowIds)
  const content = rows.grid.anchors.flatMap(owner => {
    const { nativeContinuationContent: _contents, nativeContinuations, ...style } = owner.cell.style
    return [
      ...(selected.has(rows.table.rows[owner.row]!.id) ? [{ rowId: rows.table.rows[owner.row]!.id, cellId: owner.cell.id,
        column: owner.column, colSpan: owner.colSpan, rowSpan: owner.rowSpan, style, blocks: owner.cell.blocks }] : []),
      ...documentTableContinuationEntries(owner.cell).filter(entry => selected.has(entry.content.rowId))
        .map(entry => ({ rowId: entry.content.rowId, cellId: owner.cell.id, column: owner.column, colSpan: entry.content.colSpan,
          source: entry.content.source, properties: nativeContinuations && typeof nativeContinuations === 'object'
            ? Reflect.get(nativeContinuations, String(entry.offset)) : undefined, blocks: entry.content.blocks })),
    ]
  })
  const properties = rows.table.rows.slice(rows.startRow, rows.endRow + 1).map((row, offset) => ({
    id: row.id, height: row.height, heightRule: row.heightRule, allowSplit: row.allowSplit, style: row.style,
    legacy: rows.table.rowStyles?.[rows.startRow + offset], header: rows.startRow + offset < (rows.table.headerRowCount ?? 0),
  }))
  return JSON.stringify({ tableId: range.tableId, rowIds: rows.rowIds, properties, content })
}

export function documentContentControlTableRowSelection(model: DocumentModel, range: DocumentContentControlTableRowRange): DocumentSelection {
  const rows = resolveDocumentContentControlTableRows(model, range)
  return createDocumentTableViewSelection(model, range.tableId,
    { startRow: rows.startRow, endRow: rows.endRow, startColumn: 0, endColumn: rows.grid.columnCount - 1 }, 'row')
}

/** Capability checks accept real text inside the physical row scope or its canonical structural selection. */
export function documentContentControlContainsSelection(model: DocumentModel, control: DocumentContentControl, selection: DocumentSelection): boolean {
  try {
    const targets = textTargets(model), points = [selection.anchor, selection.focus]
    if (!points.every(point => {
      const target = targets.find(target => sameTarget(target, point))
      return !!target && isDocumentTextBoundary(runText(runsForTarget(model, target)), point.offset)
    })) return false
    if (control.blockRange) {
      if (control.tableRowRange || !findDocumentBlock(model, control.blockRange.blockIds[0]!)) return false
      const owned = documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(model, control.blockRange))
      if (selection.tableSelection) {
        const view = resolveDocumentTableViewSelection(model, selection.tableSelection)
        const canonical = createDocumentTableViewSelection(model, view.tableId, view.range, view.kind, view.pageAffinity, selection)
        if (!owned.has(view.tableId) || !sameTarget(canonical.anchor, selection.anchor) || canonical.anchor.offset !== selection.anchor.offset
          || !sameTarget(canonical.focus, selection.focus) || canonical.focus.offset !== selection.focus.offset) return false
      }
      const first = targets.findIndex(target => sameTarget(target, selection.anchor)), last = targets.findIndex(target => sameTarget(target, selection.focus))
      return first >= 0 && last >= 0 && targets.slice(Math.min(first, last), Math.max(first, last) + 1)
        .every(target => owned.has(target.cellBlockId ?? target.blockId))
    }
    if (!control.tableRowRange) return !selection.tableSelection && sameTarget(selection.anchor, selection.focus)
      && documentContentControlContainsBlockRange(model, control, selection.focus.cellBlockId ?? selection.focus.blockId,
        Math.min(selection.anchor.offset, selection.focus.offset), Math.max(selection.anchor.offset, selection.focus.offset))
    if (selection.tableSelection) {
      const expected = documentContentControlTableRowSelection(model, control.tableRowRange), view = resolveDocumentTableViewSelection(model, selection.tableSelection)
      const canonical = createDocumentTableViewSelection(model, view.tableId, view.range, view.kind, view.pageAffinity, selection)
      return ['row', 'table', 'range'].includes(view.kind) && view.tableId === expected.tableSelection!.tableId
        && view.range.startRow === expected.tableSelection!.range.startRow && view.range.endRow === expected.tableSelection!.range.endRow
        && view.range.startColumn === expected.tableSelection!.range.startColumn && view.range.endColumn === expected.tableSelection!.range.endColumn
        && sameTarget(canonical.anchor, selection.anchor) && canonical.anchor.offset === selection.anchor.offset
        && sameTarget(canonical.focus, selection.focus) && canonical.focus.offset === selection.focus.offset
    }
    const owned = documentContentControlOwnedRowBlockIds(resolveDocumentContentControlTableRows(model, control.tableRowRange))
    const first = targets.findIndex(target => sameTarget(target, selection.anchor)), last = targets.findIndex(target => sameTarget(target, selection.focus))
    return first >= 0 && last >= 0 && targets.slice(Math.min(first, last), Math.max(first, last) + 1)
      .every(target => owned.has(target.cellBlockId ?? target.blockId))
  } catch (error) {
    if (error instanceof TypeError || error instanceof RangeError) return false
    throw error
  }
}

/** A submitted repeating-row edit is resolved from live row identities, including shifted table coordinates. */
export function documentContentControlTableRowsMutationSelection(before: DocumentModel, after: DocumentModel, id: string,
  action: 'duplicate' | 'delete'): { controlId?: string; selection: DocumentSelection } | undefined {
  if (before.id !== after.id) return undefined
  try {
    const item = before.contentControls?.find(control => control.id === id)
    if (item?.type !== 'repeatingSectionItem' || item.placement !== 'row' || !item.tableRowRange) return undefined
    const source = resolveDocumentContentControlTableRows(before, item.tableRowRange)
    const block = findDocumentBlock(after, source.block.id), table = block && tableForBlock(block)
    const previousRows = source.table.rows.map(row => row.id), removed = new Set(source.rowIds)
    if (action === 'duplicate') {
      const original = after.contentControls?.find(control => control.id === id)
      if (!table || !original || original.type !== item.type || original.placement !== item.placement || original.parentId !== item.parentId
        || JSON.stringify(original.tableRowRange) !== JSON.stringify(item.tableRowRange)) return undefined
      const oldIds = new Set(before.contentControls?.map(control => control.id)), oldRows = new Set(previousRows)
      const copies = (after.contentControls ?? []).filter(control => !oldIds.has(control.id) && control.type === item.type
        && control.placement === 'row' && control.parentId === item.parentId && control.tableRowRange?.tableId === source.block.id
        && control.tableRowRange.rowIds.length === source.rowIds.length && control.tableRowRange.rowIds.every(rowId => !oldRows.has(rowId)))
        .filter(control => {
          const copied = resolveDocumentContentControlTableRows(after, control.tableRowRange!)
          return copied.startRow === source.endRow + 1
            && JSON.stringify(table.rows.map(row => row.id)) === JSON.stringify([
              ...previousRows.slice(0, source.endRow + 1), ...copied.rowIds, ...previousRows.slice(source.endRow + 1),
            ])
        })
      if (copies.length !== 1) return undefined
      return { controlId: copies[0]!.id, selection: documentContentControlTableRowSelection(after, copies[0]!.tableRowRange!) }
    }
    if (after.contentControls?.some(control => control.id === id)) return undefined
    if (table) {
      if (JSON.stringify(table.rows.map(row => row.id)) !== JSON.stringify(previousRows.filter(rowId => !removed.has(rowId)))) return undefined
      const grid = documentTableGrid(table), targets = textTargets(after)
      const nextRow = Math.min(source.startRow, grid.rowCount - 1)
      // An object-only or merged row can have no owned paragraph; use the nearest live row's real input target.
      for (const row of table.rows.map((_row, index) => index).sort((left, right) => Math.abs(left - nextRow) - Math.abs(right - nextRow) || left - right)) {
        const rows = resolveDocumentContentControlTableRows(after, { tableId: source.block.id, rowIds: [table.rows[row]!.id] })
        const owned = documentContentControlTableRowBlockIds(rows), target = targets.find(target => owned.has(target.cellBlockId ?? target.blockId))
        if (target) return { selection: collapsedSelection({ ...target, offset: 0 }) }
      }
    } else if (source.startRow !== 0 || source.endRow !== source.grid.rowCount - 1) return undefined
    const previousTargets = textTargets(before), liveTargets = textTargets(after)
    const affected = documentContentControlTableRowBlockIds(source)
    const first = previousTargets.findIndex(target => affected.has(target.cellBlockId ?? target.blockId))
    const last = previousTargets.reduce((index, target, current) => affected.has(target.cellBlockId ?? target.blockId) ? current : index, -1)
    const findLive = (blockId: string) => liveTargets.find(target => (target.cellBlockId ?? target.blockId) === blockId)
    if (first >= 0) {
      for (const target of previousTargets.slice(last + 1)) {
        const current = findLive(target.cellBlockId ?? target.blockId)
        if (current) return { selection: collapsedSelection({ ...current, offset: 0 }) }
      }
      for (const target of previousTargets.slice(0, first).reverse()) {
        const current = findLive(target.cellBlockId ?? target.blockId)
        if (current) return { selection: collapsedSelection({ ...current, offset: runText(runsForTarget(after, current)).length }) }
      }
    }
    const target = liveTargets[0]
    return target ? { selection: collapsedSelection({ ...target, offset: 0 }) } : undefined
  } catch (error) {
    if (error instanceof TypeError || error instanceof RangeError) return undefined
    throw error
  }
}

/** Re-anchor a surviving row wrapper after deleting its former first paragraph. */
export function reanchorDocumentContentControlTableRows(model: DocumentModel, control: DocumentContentControl): void {
  if (!control.tableRowRange) return
  const rows = resolveDocumentContentControlTableRows(model, control.tableRowRange), owned = documentContentControlTableRowBlockIds(rows)
  const targets = textTargets(model), current = targets.find(target => (target.cellBlockId ?? target.blockId) === control.blockId)
  if (current && owned.has(control.blockId)) {
    reanchorDocumentContentControlNoteParagraph(control, control.blockId, runsForTarget(model, current)); return
  }
  const physical = documentContentControlOwnedRowBlockIds(rows)
  const target = targets.find(target => physical.has(target.cellBlockId ?? target.blockId)) ?? targets.find(target => owned.has(target.cellBlockId ?? target.blockId))
  if (!target) throw new RangeError('重复行必须保留合法的输入段落。')
  reanchorDocumentContentControlNoteParagraph(control, target.cellBlockId ?? target.blockId, runsForTarget(model, target))
}

/** Regular table mutations update row wrappers by surviving identities before paragraph anchors are reconciled. */
export function reconcileDocumentContentControlTableRows(before: DocumentModel, after: DocumentModel): void {
  const previous = new Map(before.contentControls?.filter(control => control.tableRowRange).map(control => [control.id, control]))
  const removing = new Set<string>()
  for (const control of after.contentControls ?? []) {
    if (!control.tableRowRange) continue
    const source = previous.get(control.id), block = findDocumentBlock(after, control.tableRowRange.tableId), table = block && tableForBlock(block)
    // A body candidate retains controls from stored stories/notes; their own tables are outside this scope.
    if (!table && !findDocumentBlock(before, source?.tableRowRange?.tableId ?? control.tableRowRange.tableId)) continue
    if (!table) { removing.add(control.id); continue }
    const liveIds = new Set(table.rows.map(row => row.id)), rowIds = control.tableRowRange.rowIds.filter(id => liveIds.has(id))
    if (!rowIds.length) { removing.add(control.id); continue }
    if (source?.tableRowRange && JSON.stringify(source.tableRowRange) === JSON.stringify(control.tableRowRange)) {
      const sourceRows = resolveDocumentContentControlTableRows(before, source.tableRowRange)
      const from = table.rows.findIndex(row => row.id === rowIds[0]), to = table.rows.findIndex(row => row.id === rowIds.at(-1))
      if (from < 0 || to < from) throw new RangeError('表格编辑反转了内容控件行范围。')
      // New rows inserted inside a wrapper belong to it; appending after its final row stays outside.
      const sourceIds = new Set(sourceRows.table.rows.map(row => row.id))
      const expanded = table.rows.slice(from, to + 1).map(row => row.id)
      if (expanded.some(id => !rowIds.includes(id) && sourceIds.has(id))) throw new RangeError('表格移动不能把其他已有行移入重复项目。')
      control.tableRowRange = { tableId: control.tableRowRange.tableId, rowIds: expanded }
    } else control.tableRowRange = { tableId: control.tableRowRange.tableId, rowIds }
    resolveDocumentContentControlTableRows(after, control.tableRowRange)
    reanchorDocumentContentControlTableRows(after, control)
    if (source?.tableRowRange && ['contentLocked', 'sdtContentLocked'].includes(source.lock ?? '')
      && documentContentControlTableRowContentSignature(before, source.tableRowRange) !== documentContentControlTableRowContentSignature(after, control.tableRowRange)) {
      if (!documentContentControlGroupEditAllowed(before, after, source)) throw new RangeError('表格编辑不能改变锁定的内容控件行内容。')
    }
    if (source?.type === 'repeatingSection' && source.repeatingSection?.doNotAllowInsertDelete
      && JSON.stringify(source.tableRowRange?.rowIds) !== JSON.stringify(control.tableRowRange.rowIds)) throw new RangeError('此重复节禁止通过表格命令增加或删除项目行。')
  }
  let changed = true
  while (changed) {
    changed = false
    for (const control of after.contentControls ?? []) if (control.parentId && removing.has(control.parentId) && !removing.has(control.id)) {
      removing.add(control.id); changed = true
    }
  }
  for (const id of removing) {
    const control = before.contentControls?.find(control => control.id === id)
    if (control && ['controlLocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')) throw new RangeError('表格编辑不能移除锁定的内容控件行范围。')
  }
  if (removing.size) after.contentControls = after.contentControls?.filter(control => !removing.has(control.id))
}
