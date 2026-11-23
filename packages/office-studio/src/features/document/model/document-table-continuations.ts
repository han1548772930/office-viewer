import type { DocumentBlock } from '../../../domain/model'
import type { DocumentCell, DocumentTable, DocumentTableContinuationContent } from './document-model'

const record = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value)

function readEntries(cell: DocumentCell, ownerCellId = cell.id, checkGeometry = true): Array<{ offset: number; content: DocumentTableContinuationContent }> {
  const mapping = cell.style.nativeContinuationContent
  if (mapping === undefined) return []
  if (!record(mapping)) throw new TypeError('表格续格正文映射必须包含稳定身份与正文容器')
  const rowIds = new Set<string>()
  const entries = Object.entries(mapping).map(([key, value]) => {
    const offset = Number(key)
    if (!Number.isSafeInteger(offset) || offset < 1 || String(offset) !== key || !record(value)
      || value.version !== 1 || value.ownerCellId !== ownerCellId || typeof value.rowId !== 'string' || !value.rowId
      || rowIds.has(value.rowId) || !Number.isSafeInteger(value.colSpan) || value.colSpan < 1 || !Array.isArray(value.blocks)
      || checkGeometry && (offset >= cell.rowSpan || value.colSpan !== cell.colSpan)) throw new RangeError('表格续格正文身份、跨度或映射无效，原内容已保留')
    rowIds.add(value.rowId)
    if (value.blocks.some(block => !record(block) || typeof block.id !== 'string' || !block.id || !Array.isArray(block.runs)
      || !record(block.style) || !record(block.metadata))) throw new TypeError('表格续格正文必须使用规范化的 DocumentBlock')
    return { offset, content: value as unknown as DocumentTableContinuationContent }
  })
  return entries.sort((left, right) => left.offset - right.offset)
}

/** Each array is its own authority. The flattened view never copies or owns its blocks. */
export function documentTableCellContentContainers(cell: DocumentCell): DocumentBlock[][] {
  return [cell.blocks, ...readEntries(cell).map(entry => entry.content.blocks)]
}

export function documentTableCellBlocks(cell: DocumentCell): DocumentBlock[] {
  return documentTableCellContentContainers(cell).flat()
}

export function documentTableContinuationEntries(cell: DocumentCell): Array<{ offset: number; content: DocumentTableContinuationContent }> {
  return readEntries(cell)
}

export function assertDocumentTableContinuationRows(table: DocumentTable, cell: DocumentCell, row: number): void {
  for (const entry of readEntries(cell)) if (table.rows[row + entry.offset]?.id !== entry.content.rowId) {
    throw new RangeError('表格续格正文与稳定行身份不一致，原内容已保留')
  }
  const native = cell.style.nativeContinuations
  if (native === undefined) return
  if (!record(native)) throw new TypeError('表格续格属性来源映射无效')
  const rowIds = new Set<string>()
  for (const [key, value] of Object.entries(native)) {
    const offset = Number(key)
    if (!Number.isSafeInteger(offset) || offset < 1 || String(offset) !== key || offset >= cell.rowSpan || !record(value)
      || value.ownerCellId !== cell.id || typeof value.rowId !== 'string' || !value.rowId || rowIds.has(value.rowId)
      || value.colSpan !== cell.colSpan || table.rows[row + offset]?.id !== value.rowId) throw new RangeError('表格续格属性来源与稳定行身份不一致，原内容已保留')
    rowIds.add(value.rowId)
  }
}

/** Structural mutations rekey by row identity; deleted physical rows deliberately remove their content. */
export function rebindDocumentTableContinuations(table: DocumentTable, cell: DocumentCell, row: number, previousOwnerCellId = cell.id,
  copiedRowIds?: ReadonlyMap<string, string>): void {
  const byId = new Map(table.rows.map((source, index) => [source.id, index]))
  const next: Record<string, DocumentTableContinuationContent> = {}
  for (const entry of readEntries(cell, previousOwnerCellId, false)) {
    const rowId = copiedRowIds?.get(entry.content.rowId) ?? entry.content.rowId, destination = byId.get(rowId)
    if (destination === undefined) continue
    if (destination < row || destination >= row + cell.rowSpan) throw new RangeError('表格结构变换把续格正文移出了所属合并单元格')
    if (destination === row) { cell.blocks.push(...entry.content.blocks); continue }
    next[destination - row] = { ...entry.content, ownerCellId: cell.id, rowId, colSpan: cell.colSpan }
  }
  if (Object.keys(next).length) cell.style.nativeContinuationContent = next
  else delete cell.style.nativeContinuationContent
  const native = cell.style.nativeContinuations
  if (native === undefined) return
  if (!record(native)) throw new TypeError('表格续格属性来源映射无效')
  const properties: Record<string, unknown> = {}
  for (const value of Object.values(native)) {
    if (!record(value) || value.ownerCellId !== previousOwnerCellId || typeof value.rowId !== 'string' || !value.rowId) throw new RangeError('表格续格属性来源缺少稳定身份')
    const rowId = copiedRowIds?.get(value.rowId) ?? value.rowId, destination = byId.get(rowId)
    if (destination === undefined || destination === row) continue
    if (destination < row || destination >= row + cell.rowSpan) throw new RangeError('表格续格属性来源超出所属合并单元格')
    properties[destination - row] = { ...value, ownerCellId: cell.id, rowId, colSpan: cell.colSpan }
  }
  if (Object.keys(properties).length) cell.style.nativeContinuations = properties
  else delete cell.style.nativeContinuations
}
