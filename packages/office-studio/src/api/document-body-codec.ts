import type { DocumentContentControl, DocumentDefaultStyle, DocumentStoryContent } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentDefaultStyle } from '../features/document/model/document-default-style'
import { tableForBlock } from '../features/document/model/document-model'
import { projectDocumentStoryBlocks } from '../features/document/model/document-story-blocks'
import { documentTableCellContentContainers } from '../features/document/model/document-table-continuations'
import { bodyRecord } from './document-body-style'
import { readDocumentBody } from './document-body-read'
import { applyDocumentBodyRanges } from './document-body-ranges'
import { writeDocumentBody } from './document-body-write'
import type { DocumentBodyDecodeOptions, DocumentBodyDecodeResult, DocumentBodyEncodeResult, UniverDocumentBodyData, UniverDocumentBodyResources } from './document-body-types'

export type { UniverDocumentBody, UniverDocumentBodyData, UniverDocumentBodyResources, DocumentBodyDecodeResult, DocumentBodyEncodeResult } from './document-body-types'

/** The source payload is JSON control data, with explicit depth and size bounds. */
function assertSerializable(value: unknown): void {
  const ancestors = new Set<object>(); let count = 0
  const visit = (value: unknown, depth: number) => {
    if (++count > 2_000_000 || depth > 64) throw new RangeError('文档公开数据超过单次结构预算。')
    if (value === undefined || value === null || typeof value === 'string' || typeof value === 'boolean') return
    if (typeof value === 'number' && Number.isFinite(value)) return
    if (!value || typeof value !== 'object' || ancestors.has(value)) throw new TypeError('文档公开数据必须为可保存的 JSON 数据。')
    const prototype = Object.getPrototypeOf(value)
    if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) throw new TypeError('文档公开数据不能包含浏览器或类实例。')
    ancestors.add(value); for (const entry of Object.values(value)) visit(entry, depth + 1); ancestors.delete(value)
  }
  visit(value, 0)
}

export function decodeDocumentBody(body: unknown, resources: UniverDocumentBodyResources = {}, options: Partial<DocumentBodyDecodeOptions> = {}): DocumentBodyDecodeResult {
  assertSerializable({ body, resources, options })
  for (const key of ['tableSource', 'drawings', 'lists'] as const) if (resources[key] !== undefined) {
    const source = resources[key]
    if (!source || typeof source !== 'object' || Array.isArray(source) || Object.values(source).some(value => !value || typeof value !== 'object' || Array.isArray(value))) {
      throw new TypeError(`文档 ${key} 必须按稳定标识提供对象。`)
    }
  }
  if (resources.drawingsOrder !== undefined && (!Array.isArray(resources.drawingsOrder) || resources.drawingsOrder.some(id => typeof id !== 'string'))) throw new TypeError('文档绘图顺序必须为标识数组。')
  const config: DocumentBodyDecodeOptions = { ...options, defaults: options.defaults ?? documentDefaultStyle() }
  const result = readDocumentBody(body, resources, config), controls = applyDocumentBodyRanges(body, result.positions, result.diagnostics)
  const content: DocumentStoryContent = { blocks: result.blocks, runs: [], paragraphs: [], metadata: { univerBodySource: {
    version: 1, body: cloneModel(bodyRecord(body)), resources: cloneModel(resources), ...(result.diagnostics.length ? { diagnostics: [...result.diagnostics] } : {}) } } }
  return { content: projectDocumentStoryBlocks(content, 'display').content, contentControls: controls, diagnostics: result.diagnostics }
}

export function encodeDocumentBody(content: DocumentStoryContent, defaults: DocumentDefaultStyle = documentDefaultStyle(),
  contentControls: readonly DocumentContentControl[] = []): UniverDocumentBodyData {
  return writeDocumentBody(content, defaults, contentControls)
}

/** Ephemeral stream-to-model coordinates are returned separately from the serializable public body. */
export function encodeDocumentBodyWithPositions(content: DocumentStoryContent, defaults: DocumentDefaultStyle = documentDefaultStyle(),
  contentControls: readonly DocumentContentControl[] = []): DocumentBodyEncodeResult {
  const positions: DocumentBodyEncodeResult['positions'] = []
  return { data: writeDocumentBody(content, defaults, contentControls, positions), positions }
}

/** Note and header/footer exports include only controls actually owned by their block tree. */
export function documentBodyContentControls(content: DocumentStoryContent, controls: readonly DocumentContentControl[] = []): DocumentContentControl[] {
  const ids = new Set<string>()
  const visit = (blocks: NonNullable<DocumentStoryContent['blocks']>) => {
    for (const block of blocks) {
      ids.add(block.id)
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const child of documentTableCellContentContainers(cell)) visit(child)
    }
  }
  visit(content.blocks ?? [])
  return controls.filter(control => ids.has(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)).map(cloneModel)
}
