import type { DocumentContentControl, DocumentStoryContent } from '../domain/model'
import { documentId } from '../features/document/model/document-model'
import { projectDocumentStoryBlocks } from '../features/document/model/document-story-blocks'
import { copyDocumentBlocksForPaste } from '../features/document/services/document-block-paste'
import { createDocumentRepeatingContentControlCopyIds, indexDocumentRepeatingContentControlCopies,
  remapDocumentRepeatingContentControlMetadata, visitDocumentRepeatingContentControlBlocks } from '../features/document/services/repeating-content-control-copy'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'
import type { DocumentBodyDecodeResult, UniverDocumentBodyResources } from './document-body-types'
import { documentBodyContentIdentities } from './document-body-identities'

export function publicDocumentRecord(value: unknown, label: string): UniverBodyRecord {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label} 必须为对象。`)
  return bodyRecord(value)
}

export function publicDocumentResources(value: UniverBodyRecord, fallback: UniverDocumentBodyResources = {}): UniverDocumentBodyResources {
  const result: UniverDocumentBodyResources = {}
  for (const key of ['tableSource', 'drawings', 'lists'] as const) {
    if (value[key] === undefined && fallback[key] === undefined) continue
    const own = value[key] === undefined ? {} : publicDocumentRecord(value[key], key)
    const inherited = fallback[key] === undefined ? {} : publicDocumentRecord(fallback[key], key)
    const entries = { ...inherited, ...own }
    for (const [id, entry] of Object.entries(entries)) publicDocumentRecord(entry, `${key}.${id}`)
    Object.assign(result, { [key]: entries })
  }
  if (value.drawingsOrder !== undefined) {
    if (!Array.isArray(value.drawingsOrder) || value.drawingsOrder.some(id => typeof id !== 'string')) throw new TypeError('drawingsOrder 必须为标识数组。')
    result.drawingsOrder = [...value.drawingsOrder]
  } else if (fallback.drawingsOrder) result.drawingsOrder = [...fallback.drawingsOrder]
  return result
}

export function reservePublicDocumentContent(content: DocumentStoryContent, controls: readonly DocumentContentControl[], claimed: Set<string>): void {
  const own = new Set<string>()
  for (const id of documentBodyContentIdentities(content.blocks ?? [], controls)) {
    if (own.has(id) || claimed.has(id)) throw new TypeError(`文档内容区身份重复：${id}`)
    own.add(id)
  }
  for (const id of own) claimed.add(id)
}

/** Shared source parts keep their public IDs at first use; copied hosts obtain
 * one real ID map for blocks, fields, table owners and SDT wrappers. */
export function claimPublicDocumentBody(decoded: DocumentBodyDecodeResult, claimed: Set<string>): DocumentBodyDecodeResult {
  const sourceBlocks = decoded.content.blocks ?? []
  if (!documentBodyContentIdentities(sourceBlocks, decoded.contentControls).some(id => claimed.has(id))) {
    reservePublicDocumentContent(decoded.content, decoded.contentControls, claimed)
    return decoded
  }
  const blocks = copyDocumentBlocksForPaste(sourceBlocks), ids = createDocumentRepeatingContentControlCopyIds()
  indexDocumentRepeatingContentControlCopies(sourceBlocks, blocks, ids)
  for (const control of decoded.contentControls) ids.control.set(control.id, documentId('content-control'))
  const mapped = (mapping: ReadonlyMap<string, string>, id: string, label: string) => {
    const result = mapping.get(id)
    if (!result) throw new RangeError(`${label} 未指向同一来源内容区：${id}`)
    return result
  }
  const controls = decoded.contentControls.map(control => ({ ...control,
    id: mapped(ids.control, control.id, '控件'), blockId: mapped(ids.block, control.blockId, '控件输入锚点'),
    ...(control.parentId === undefined ? {} : { parentId: mapped(ids.control, control.parentId, '父控件') }),
    ...(control.blockRange ? { blockRange: { blockIds: control.blockRange.blockIds.map(id => mapped(ids.block, id, '块范围')) } } : {}),
    ...(control.tableRowRange ? { tableRowRange: { tableId: mapped(ids.block, control.tableRowRange.tableId, '表格'),
      rowIds: control.tableRowRange.rowIds.map(id => mapped(ids.row, id, '表格行')) } } : {}),
    metadata: remapDocumentRepeatingContentControlMetadata(control.metadata, ids),
  }))
  visitDocumentRepeatingContentControlBlocks(blocks, block => {
    block.metadata = remapDocumentRepeatingContentControlMetadata(block.metadata, ids) ?? {}
    for (const run of block.runs) if (run.metadata) run.metadata = remapDocumentRepeatingContentControlMetadata(run.metadata, ids)
  })
  const content = projectDocumentStoryBlocks({ ...decoded.content, blocks,
    metadata: { ...decoded.content.metadata, univerBodyCopiedSource: true } }, 'display').content
  reservePublicDocumentContent(content, controls, claimed)
  return { ...decoded, content, contentControls: controls }
}
