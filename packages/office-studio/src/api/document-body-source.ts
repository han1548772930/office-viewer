import type { DocumentBlock, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { runText } from '../features/document/model/document-model'
import { isListPreset, listKind } from '../features/document/model/document-lists'
import { isDocumentStructuredBlock } from '../features/document/model/document-structural-blocks'
import { bodyRecord, documentRunStyleFromBody, documentRunStyleToBody, type UniverBodyRecord } from './document-body-style'
import { documentBodyIdentity } from './document-body-types'

const BLOCK_TYPES: DocumentBlock['type'][] = ['paragraph', 'heading', 'bullet', 'numbered', 'checklist', 'quote', 'divider',
  'callout', 'code', 'columns', 'table', 'image', 'pageBreak', 'columnBreak', 'sectionBreak']
const STRUCTURAL_BLOCK_TYPES = new Set<DocumentBlock['type']>(['callout', 'code', 'columns', 'table', 'image', 'pageBreak', 'columnBreak', 'sectionBreak'])
function isBlockType(value: unknown): value is DocumentBlock['type'] {
  return typeof value === 'string' && BLOCK_TYPES.some(type => type === value)
}
const BLOCK_DERIVED = ['table', 'html', 'documentTextFields', 'univerBodyParagraphSource', 'univerBodyTableSource',
  'univerBodyCustomBlock', 'univerBodyDrawingSource', 'univerBodySectionSource', 'univerBodyBreakSource', 'univerNoteSeparatorRangeSource',
  'univerBodyBlockRangeSource', 'univerBodyColumnGroupSource', 'univerBodyColumnGroupBaseline', 'univerBodyCalloutSource', 'univerBodyCodeSource']
const RUN_DERIVED = ['univerTextStyle', 'univerTextStyleResolution', 'univerBodyRunSource', 'univerBodyNoteRangeSource',
  'univerBodyHyperlinkSource', 'univerBodyCustomBlock', 'univerBodyDrawingSource', 'univerBodyRawCustomBlock']

/** The extension accompanies the actual public paragraph/run/table record. It
 * stores identity and provenance, never a parallel block tree or body string. */
export function documentBodyOfficeSource(record: UniverBodyRecord, label: string): UniverBodyRecord {
  if (record.officeStudio === undefined) return {}
  const source = record.officeStudio
  if (!source || typeof source !== 'object' || Array.isArray(source) || Reflect.get(source, 'version') !== 1) {
    throw new TypeError(`${label} 的 Office 来源扩展版本或结构无效。`)
  }
  return source as UniverBodyRecord
}
export function documentBodySourceMetadata(value: unknown, label: string, excluded: readonly string[] = []): UniverBodyRecord {
  if (value === undefined) return {}
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label} 来源必须为对象。`)
  const result = cloneModel(bodyRecord(value))
  for (const key of excluded) delete result[key]
  return result
}
export function documentBodyBlockSource(block: DocumentBlock): UniverBodyRecord {
  const metadata = documentBodySourceMetadata(block.metadata, '块', BLOCK_DERIVED)
  if (metadata.codeBlock !== undefined) {
    const code = documentBodySourceMetadata(metadata.codeBlock, '代码块')
    delete code.source
    metadata.codeBlock = code
  }
  return { version: 1, blockType: block.type, level: block.level,
    ...(block.list === undefined ? {} : { list: cloneModel(block.list) }),
    ...(block.checklist === undefined ? {} : { checklist: cloneModel(block.checklist) }),
    ...(Object.keys(metadata).length ? { metadata } : {}),
    ...(block.runs.length === 1 && !block.runs[0]!.text ? { emptyRun: { runId: block.runs[0]!.id,
      ts: documentRunStyleToBody(block.runs[0]!), metadata: documentBodySourceMetadata(block.runs[0]!.metadata, '空运行', RUN_DERIVED) } } : {}) }
}
export function applyDocumentBodyBlockSource(block: DocumentBlock, record: UniverBodyRecord): void {
  const source = documentBodyOfficeSource(record, '块')
  if (!Object.keys(source).length) return
  if (source.blockType !== undefined) {
    if (!isBlockType(source.blockType)) throw new TypeError('公开块类型无效。')
    if (source.blockType !== block.type && (isDocumentStructuredBlock(block) || STRUCTURAL_BLOCK_TYPES.has(block.type) || STRUCTURAL_BLOCK_TYPES.has(source.blockType))) {
      throw new TypeError('公开结构块类型与内容流标记不匹配。')
    }
    block.type = source.blockType
  }
  if (source.level !== undefined) {
    if (typeof source.level !== 'number' || !Number.isSafeInteger(source.level) || source.level < 0 || source.level > 9) throw new RangeError('公开块层级无效。')
    block.level = source.level
  }
  const metadata = documentBodySourceMetadata(source.metadata, '块', BLOCK_DERIVED)
  if (metadata.codeBlock !== undefined) metadata.codeBlock = { ...documentBodySourceMetadata(metadata.codeBlock, '代码块'), source: runText(block.runs) }
  block.metadata = { ...metadata, ...block.metadata }
  if (source.list !== undefined) {
    const list = bodyRecord(source.list)
    if (!isListPreset(list.preset) || (block.type === 'bullet' || block.type === 'numbered') && listKind(list.preset) !== block.type) throw new TypeError('公开列表预设无效。')
    block.list = { id: documentBodyIdentity(list.id, '列表'), preset: list.preset }
  }
  if (source.checklist !== undefined) {
    const checklist = bodyRecord(source.checklist)
    if (block.type !== 'checklist' || typeof checklist.checked !== 'boolean') throw new TypeError('公开检查列表状态无效。')
    block.checklist = { listId: documentBodyIdentity(checklist.listId, '检查列表'), checked: checklist.checked }
  }
  if (source.emptyRun !== undefined) {
    const empty = bodyRecord(source.emptyRun)
    if (block.runs.some(run => run.text.length)) throw new RangeError('非空段落不能用空运行来源覆盖文字。')
    const run = block.runs[0]
    if (!run) throw new RangeError('空运行来源缺少真实段落。')
    run.id = documentBodyIdentity(empty.runId, '空运行')
    Object.assign(run, documentRunStyleFromBody(empty.ts))
    run.metadata = { ...documentBodySourceMetadata(empty.metadata, '空运行', RUN_DERIVED), ...run.metadata }
  }
}
export function documentBodyRunSource(run: InlineRun): UniverBodyRecord {
  const metadata = documentBodySourceMetadata(run.metadata, '运行', RUN_DERIVED)
  return { version: 1, runId: run.id, ...(Object.keys(metadata).length ? { metadata } : {}) }
}
export function documentBodyRunMetadata(record?: UniverBodyRecord): UniverBodyRecord {
  return record ? documentBodySourceMetadata(documentBodyOfficeSource(record, '运行').metadata, '运行', RUN_DERIVED) : {}
}
