import type { DocumentBlock, DocumentContentControl } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { runText, tableForBlock } from '../features/document/model/document-model'
import { documentBlockChildContainers } from '../features/document/model/document-structural-blocks'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'
import { documentBodyOfficeSource, documentBodySourceMetadata } from './document-body-source'
import { DOCUMENT_BODY_TOKENS as T, documentBodyIdentity, type DocumentBodyBlockPosition, type DocumentBodyParagraphPosition,
  type DocumentBodyRowPosition, type DocumentBodyStructureIndex } from './document-body-types'

export interface DocumentBodyControlBounds { start: number; end: number }
const SOURCE_KEYS = ['univerBodySdtSource', 'univerBodyFieldSource']
const PLACEMENTS = ['inline', 'block', 'table', 'row', 'cell', 'header', 'footer']

function identities(value: unknown, label: string): string[] {
  if (!Array.isArray(value) || !value.length) throw new TypeError(`${label} 必须包含实际结构身份。`)
  const ids = value.map(value => documentBodyIdentity(value, label))
  if (new Set(ids).size !== ids.length) throw new TypeError(`${label} 身份重复。`)
  return ids
}
function blockBounds(ids: readonly string[], index: DocumentBodyStructureIndex): { bounds: DocumentBodyControlBounds; entries: DocumentBodyBlockPosition[] } {
  const entries = ids.map(id => {
    const entry = index.blocks.find(entry => entry.block.id === id)
    if (!entry) throw new RangeError('控件完整块已不存在：' + id)
    return entry
  })
  const container = entries[0]!.container, start = container.findIndex(block => block.id === ids[0])
  if (entries.some((entry, position) => entry.container !== container || container[start + position]?.id !== entry.block.id)) throw new RangeError('控件块必须在同一内容区连续排列。')
  return { bounds: { start: entries[0]!.start, end: entries.at(-1)!.end }, entries }
}
function rowBounds(tableId: string, ids: readonly string[], index: DocumentBodyStructureIndex): { bounds: DocumentBodyControlBounds; entries: DocumentBodyRowPosition[] } {
  const entries = ids.map(id => {
    const entry = index.rows.find(entry => entry.tableBlock.id === tableId && entry.rowId === id)
    if (!entry) throw new RangeError('控件完整表格行已不存在：' + id)
    return entry
  })
  if (entries.some((entry, position) => entry.rowIndex !== entries[0]!.rowIndex + position)) throw new RangeError('控件表格行必须连续排列。')
  return { bounds: { start: entries[0]!.start, end: entries.at(-1)!.end }, entries }
}

/** Current structural ownership is authoritative over obsolete source indices. */
export function documentBodyStructuralControlBounds(control: DocumentContentControl, index: DocumentBodyStructureIndex): DocumentBodyControlBounds | undefined {
  if (control.tableRowRange) return rowBounds(control.tableRowRange.tableId, control.tableRowRange.rowIds, index).bounds
  if (!control.blockRange) return undefined
  const { bounds, entries } = blockBounds(control.blockRange.blockIds, index)
  const publicPlacement = bodyRecord(bodyRecord(control.metadata?.univerBodySdtSource).properties).placement
  if (publicPlacement === 'cell') {
    const cell = index.cells.find(cell => cell.blocks === entries[0]!.container && cell.blocks.length === entries.length)
    if (cell) return { start: cell.start, end: cell.end }
  }
  return bounds
}
export function documentBodyPublicControlPlacement(control: DocumentContentControl, index: DocumentBodyStructureIndex): string {
  const original = bodyRecord(bodyRecord(control.metadata?.univerBodySdtSource).properties).placement
  if (control.blockRange && original === 'table' && control.blockRange.blockIds.length === 1
    && index.blocks.some(entry => entry.block.id === control.blockRange!.blockIds[0] && tableForBlock(entry.block))) return 'table'
  if (control.blockRange && original === 'cell' && index.cells.some(cell => cell.blocks.length === control.blockRange!.blockIds.length
    && cell.blocks.every((block, position) => block.id === control.blockRange!.blockIds[position]))) return 'cell'
  return control.placement ?? 'inline'
}
export function documentBodyControlSource(control: DocumentContentControl): UniverBodyRecord {
  return { version: 1, blockId: control.blockId, placement: control.placement ?? 'inline',
    ...(control.parentId === undefined ? {} : { parentId: control.parentId }),
    ...(control.blockRange ? { blockRange: cloneModel(control.blockRange) } : {}),
    ...(control.tableRowRange ? { tableRowRange: cloneModel(control.tableRowRange) } : {}),
    metadata: documentBodySourceMetadata(control.metadata, '内容控件', SOURCE_KEYS) }
}

function ownedBlockIds(blocks: readonly DocumentBlock[]): Set<string> {
  const ids = new Set<string>()
  const visit = (blocks: readonly DocumentBlock[]) => {
    for (const block of blocks) {
      ids.add(block.id)
      for (const children of documentBlockChildContainers(block)) visit(children)
    }
  }
  visit(blocks)
  return ids
}

/** Interpret true table/row/cell boundaries and restore the existing domain's
 * complete-block/row contract with one addressable input fallback. */
export function applyDocumentBodyControlOwnership(control: DocumentContentControl, record: UniverBodyRecord, from: number, to: number,
  positions: readonly DocumentBodyParagraphPosition[], index: DocumentBodyStructureIndex, stream: string): void {
  const source = documentBodyOfficeSource(record, '内容控件')
  if (source.placement !== undefined && (typeof source.placement !== 'string' || !PLACEMENTS.includes(source.placement))) throw new TypeError('内容控件来源位置无效。')
  if (source.parentId !== undefined) control.parentId = documentBodyIdentity(source.parentId, '父控件')
  control.metadata = { ...documentBodySourceMetadata(source.metadata, '内容控件', SOURCE_KEYS), ...control.metadata }
  const canonicalStart = (offset: number) => { while (offset < stream.length && (stream[offset] === T.rangeStart || stream[offset] === T.rangeEnd)) offset++; return offset }
  const canonicalEnd = (offset: number) => {
    while (offset > 0 && (stream[offset - 1] === T.rangeStart || stream[offset - 1] === T.rangeEnd)) offset--
    if (stream[offset - 1] === T.paragraph) offset--
    while (offset > 0 && (stream[offset - 1] === T.rangeStart || stream[offset - 1] === T.rangeEnd)) offset--
    return offset
  }
  const begins = canonicalStart(from), ends = canonicalEnd(to)
  const matches = (bounds: DocumentBodyControlBounds) => canonicalStart(bounds.start) === begins && canonicalEnd(bounds.end) === ends
  let bounds: DocumentBodyControlBounds | undefined, owned: Set<string> | undefined
  if (source.blockRange !== undefined) {
    const ids = identities(bodyRecord(source.blockRange).blockIds, '控件完整块'), result = blockBounds(ids, index)
    control.blockRange = { blockIds: ids }; control.placement = 'block'; bounds = result.bounds; owned = ownedBlockIds(result.entries.map(entry => entry.block))
    if (!matches(bounds) && bodyRecord(record.properties).placement === 'cell') {
      const cell = index.cells.find(cell => cell.blocks === result.entries[0]!.container && cell.blocks.length === ids.length)
      if (cell) bounds = { start: cell.start, end: cell.end }
    }
  } else if (source.tableRowRange !== undefined || control.placement === 'row') {
    const raw = bodyRecord(source.tableRowRange)
    const rows = source.tableRowRange === undefined ? index.rows.filter(row => row.start >= begins && row.end <= to) : []
    const tableId = source.tableRowRange === undefined ? rows[0]?.tableBlock.id : documentBodyIdentity(raw.tableId, '控件表格')
    if (!tableId || rows.some(row => row.tableBlock.id !== tableId)) throw new RangeError('行控件必须拥有同一表格的完整行。')
    const ids = source.tableRowRange === undefined ? rows.map(row => row.rowId) : identities(raw.rowIds, '控件表格行')
    if (!ids.length) throw new RangeError('行控件缺少完整表格行。')
    const result = rowBounds(tableId, ids, index)
    control.tableRowRange = { tableId, rowIds: ids }; control.placement = 'row'; bounds = result.bounds
    owned = new Set(positions.filter(position => position.start >= bounds!.start && position.end <= bounds!.end).map(position => position.block.id))
  } else if (control.placement === 'table') {
    const entry = index.blocks.find(entry => tableForBlock(entry.block) && matches(entry))
    if (!entry) throw new RangeError('表格控件范围必须覆盖完整表格。')
    control.blockRange = { blockIds: [entry.block.id] }; control.placement = 'block'; bounds = entry; owned = ownedBlockIds([entry.block])
  } else if (control.placement === 'cell') {
    const cell = index.cells.find(cell => matches(cell))
    if (cell) { control.blockRange = { blockIds: cell.blocks.map(block => block.id) }; control.placement = 'block'; bounds = cell; owned = ownedBlockIds(cell.blocks) }
  } else if (control.placement === 'block') {
    const candidates = index.blocks.filter(entry => canonicalStart(entry.start) >= begins && canonicalEnd(entry.end) <= ends)
    const groups = [...new Set(candidates.map(entry => entry.container))].map(container => candidates.filter(entry => entry.container === container)
      .sort((a, b) => container.indexOf(a.block) - container.indexOf(b.block)))
    const group = groups.filter(entries => entries.length && matches({ start: entries[0]!.start, end: entries.at(-1)!.end }))
      .sort((a, b) => a[0]!.start - b[0]!.start || b.at(-1)!.end - a.at(-1)!.end)[0]
    if (group) {
      const ids = group.map(entry => entry.block.id), result = blockBounds(ids, index)
      control.blockRange = { blockIds: ids }; bounds = result.bounds; owned = ownedBlockIds(group.map(entry => entry.block))
    }
  }
  if (source.blockRange !== undefined && source.tableRowRange !== undefined) throw new TypeError('控件不能同时拥有完整块与表格行。')
  if (bounds && (!matches(bounds) || !owned)) throw new RangeError('内容控件来源身份与公开结构范围不匹配。')
  if (owned) {
    for (const entry of index.blocks) {
      const anchor = bodyRecord(entry.block.metadata.univerBodyStructuralInputAnchor)
      if (anchor.version !== 1 || typeof anchor.imageId !== 'string' || !owned.has(anchor.imageId)) continue
      owned.add(entry.block.id)
      const at = control.blockRange?.blockIds.indexOf(anchor.imageId) ?? -1
      if (at >= 0 && !control.blockRange!.blockIds.includes(entry.block.id)) control.blockRange!.blockIds.splice(at, 0, entry.block.id)
    }
    const fallbackId = source.blockId === undefined ? undefined : documentBodyIdentity(source.blockId, '控件输入段落')
    const fallback = positions.find(position => owned!.has(position.block.id) && (fallbackId === undefined || position.block.id === fallbackId))
    if (!fallback) throw new RangeError('结构内容控件缺少其范围内的输入段落。')
    control.blockId = fallback.block.id; control.startOffset = 0; control.endOffset = runText(fallback.block.runs).length
  } else if (source.blockId !== undefined && documentBodyIdentity(source.blockId, '控件输入段落') !== control.blockId) throw new RangeError('内容控件输入身份与文字范围不匹配。')
}
