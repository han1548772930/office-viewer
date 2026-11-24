import type { DocumentContentControl, DocumentModel, DocumentStoryContent } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentDefaultStyle } from '../model/document-default-style'
import { createParagraph, createRun, documentId, runText, tableForBlock } from '../model/document-model'
import { documentRunHasNoteReference } from '../model/document-note-reference'
import { migrateDocumentStoryBlocks, projectDocumentStoryBlocks, validateDocumentStoryBlocks } from '../model/document-story-blocks'
import { documentStoryParagraphs, validateDocumentStoryContent, validateDocumentStoryLocks } from '../model/document-story'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { copyDocumentBlocksForPaste } from '../services/document-block-paste'
import { validateContentControlPropertyPatch } from '../services/content-control-properties'
import { validateDocumentContentControlRanges, type ContentControlPatch } from '../services/content-controls'
import { createDocumentRepeatingContentControlCopyIds, indexDocumentRepeatingContentControlCopies, remapDocumentRepeatingContentControlMetadata,
  visitDocumentRepeatingContentControlBlocks, type DocumentRepeatingContentControlCopyIds } from '../services/repeating-content-control-copy'
import type { DocumentNoteKind } from './document-note-types'

/** Runtime requests and snapshots contain JSON data, never functions or browser handles. */
export function assertDocumentNoteSerializable(value: unknown): void {
  const ancestors = new Set<object>()
  let entries = 0
  const visit = (item: unknown, depth: number) => {
    if (++entries > 2_000_000 || depth > 64) throw new RangeError('注释数据超过单次内容预算。')
    if (item === null || item === undefined || typeof item === 'string' || typeof item === 'boolean') return
    if (typeof item === 'number' && Number.isFinite(item)) return
    if (typeof item !== 'object' || ancestors.has(item)) throw new TypeError('注释数据必须可保存且不能包含循环引用。')
    const prototype = Object.getPrototypeOf(item)
    if (!Array.isArray(item) && prototype !== Object.prototype && prototype !== null) throw new TypeError('注释数据不能包含非 JSON 对象。')
    ancestors.add(item)
    for (const child of Object.values(item)) visit(child, depth + 1)
    ancestors.delete(item)
  }
  visit(value, 0)
}

export function validateDocumentNoteContent(content: DocumentStoryContent, kind: DocumentNoteKind = 'footnote', id = 'validation'): void {
  assertDocumentNoteSerializable(content)
  if (!content || typeof content !== 'object' || Array.isArray(content) || !Array.isArray(content.runs) || !Array.isArray(content.paragraphs)) {
    throw new TypeError('注释内容需要真实段落与字符模型。')
  }
  const defaults = documentDefaultStyle()
  if (content.blocks === undefined) {
    validateDocumentStoryContent(content, runText(content.runs))
    const runIds = new Set<string>()
    for (const run of content.runs) {
      if (!run.id || runIds.has(run.id)) throw new TypeError('注释来源字符标识重复。')
      runIds.add(run.id)
    }
  }
  const prepared = migrateDocumentStoryBlocks(cloneModel(content), [], { kind: 'note', noteKind: kind, noteId: id }, defaults)
  validateDocumentStoryBlocks(prepared.blocks!)
  validateDocumentStoryLocks(prepared.locks, runText(prepared.runs))
  let count = 0, bytes = 0
  const visit = (blocks: NonNullable<DocumentStoryContent['blocks']>) => {
    for (const block of blocks) {
      if (++count > 250_000) throw new RangeError('注释内容超过块预算。')
      if (['sectionBreak', 'pageBreak', 'columnBreak'].includes(block.type)) throw new RangeError('注释内容不包含独立的分节或页面设置。')
      for (const run of block.runs) {
        bytes += run.text.length * 2
        if (bytes > 32 * 1024 * 1024) throw new RangeError('注释内容超过文本预算。')
        if (documentRunHasNoteReference(run)) throw new RangeError('脚注和尾注内容不能嵌套新的注释引用。')
      }
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const container of documentTableCellContentContainers(cell)) visit(container)
    }
  }
  visit(prepared.blocks!)
}

export interface PreparedDocumentNoteContent {
  content: DocumentStoryContent
  contentControls: DocumentContentControl[]
}

function noteControlModel(model: DocumentModel, blocks: NonNullable<DocumentStoryContent['blocks']>, contentControls: DocumentContentControl[]): DocumentModel {
  return { kind: 'document', name: model.name, blocks, page: model.page, modified: false, wordCount: 0,
    defaultStyle: documentDefaultStyle(model), customXmlParts: model.customXmlParts, contentControls }
}

function validateNoteContentControls(model: DocumentModel, blocks: NonNullable<DocumentStoryContent['blocks']>, controls: readonly DocumentContentControl[]): DocumentContentControl[] {
  assertDocumentNoteSerializable(controls)
  if (!Array.isArray(controls) || controls.length > 100_000) throw new TypeError('注释内容控件需要在预算内按数组提供。')
  const owned = cloneModel([...controls]), ids = new Set<string>()
  for (const control of owned) {
    if (!control || typeof control !== 'object' || Array.isArray(control) || typeof control.id !== 'string' || !control.id.trim()
      || control.id.length > 4096 || ids.has(control.id) || typeof control.blockId !== 'string' || !control.blockId.trim()) {
      throw new TypeError('注释内容控件需要唯一标识和真实的块锚点。')
    }
    ids.add(control.id)
    if (control.metadata?.officeStoryLockSource !== undefined) throw new TypeError('临时内容锁不能作为原生内容控件复制，请保留故事的 locks 数据。')
    const properties: ContentControlPatch = { ...control }
    for (const key of ['id', 'blockId', 'startOffset', 'endOffset', 'value', 'checked', 'dateValue']) Reflect.deleteProperty(properties, key)
    validateContentControlPropertyPatch(properties, control)
    if (control.value !== undefined && typeof control.value !== 'string' || control.dateValue !== undefined && typeof control.dateValue !== 'string'
      || control.checked !== undefined && typeof control.checked !== 'boolean') throw new TypeError('注释内容控件的值类型无效。')
  }
  validateDocumentContentControlRanges(noteControlModel(model, blocks, owned))
  return owned
}

function mappedNoteIdentity(ids: ReadonlyMap<string, string>, source: string, label: string): string {
  const destination = ids.get(source)
  if (!destination) throw new RangeError(`注释内容控件的${label}不属于同一来源内容区：${source}`)
  return destination
}

/** Only live structural addresses are remapped; nested source XML/codec provenance stays unchanged. */
function validateNoteCopyMetadata(metadata: Record<string, unknown> | undefined, ids: DocumentRepeatingContentControlCopyIds): void {
  if (!metadata) return
  for (const key of ['anchorBlockId', 'tailBlockId', 'pictureBlockId', 'pictureAnchorBlockId', 'pictureTailBlockId']) {
    const value = metadata[key]
    if (value === undefined) continue
    if (typeof value !== 'string') throw new TypeError('注释内容的图片或块锚点类型无效：' + key)
    mappedNoteIdentity(ids.block, value, key)
  }
  if (metadata.contentControlId !== undefined) {
    if (typeof metadata.contentControlId !== 'string') throw new TypeError('注释内容的控件锚点类型无效。')
    mappedNoteIdentity(ids.control, metadata.contentControlId, '控件锚点')
  }
}

/** Blocks and SDT wrappers use the same actual copy map and are published by one story transaction. */
export function prepareDocumentNoteContentWithControls(model: DocumentModel, kind: DocumentNoteKind, id: string,
  input?: DocumentStoryContent, text = '', copy = true, contentControls: readonly DocumentContentControl[] = []): PreparedDocumentNoteContent {
  const defaults = documentDefaultStyle(model)
  const source = input ?? { runs: [createRun(text, defaults.text)], paragraphs: documentStoryParagraphs(text, defaults.paragraph) }
  validateDocumentNoteContent(source, kind, id)
  const migrated = migrateDocumentStoryBlocks(cloneModel(source), [], { kind: 'note', noteKind: kind, noteId: id }, defaults)
  if (!migrated.blocks!.length) migrated.blocks = [createParagraph('', defaults.text)]
  const sourceControls = validateNoteContentControls(model, migrated.blocks!, contentControls)
  const blocks = copy ? copyDocumentBlocksForPaste(migrated.blocks!) : migrated.blocks!
  const ids = createDocumentRepeatingContentControlCopyIds()
  indexDocumentRepeatingContentControlCopies(migrated.blocks!, blocks, ids)
  for (const control of sourceControls) ids.control.set(control.id, copy ? documentId('content-control') : control.id)
  const controls = sourceControls.map(control => {
    validateNoteCopyMetadata(control.metadata, ids)
    const owned = { ...control, id: ids.control.get(control.id)!, blockId: mappedNoteIdentity(ids.block, control.blockId, '输入锚点') }
    if (control.parentId !== undefined) owned.parentId = mappedNoteIdentity(ids.control, control.parentId, '父控件')
    if (control.blockRange) owned.blockRange = { blockIds: control.blockRange.blockIds.map(id => mappedNoteIdentity(ids.block, id, '块范围')) }
    if (control.tableRowRange) owned.tableRowRange = { tableId: mappedNoteIdentity(ids.block, control.tableRowRange.tableId, '表格'),
      rowIds: control.tableRowRange.rowIds.map(id => mappedNoteIdentity(ids.row, id, '表格行')) }
    owned.metadata = remapDocumentRepeatingContentControlMetadata(control.metadata, ids)
    return owned
  })
  visitDocumentRepeatingContentControlBlocks(blocks, block => {
    validateNoteCopyMetadata(block.metadata, ids)
    block.metadata = remapDocumentRepeatingContentControlMetadata(block.metadata, ids)!
    for (const run of block.runs) {
      validateNoteCopyMetadata(run.metadata, ids)
      if (run.metadata) run.metadata = remapDocumentRepeatingContentControlMetadata(run.metadata, ids)
    }
  })
  validateDocumentContentControlRanges(noteControlModel(model, blocks, controls))
  return { content: projectDocumentStoryBlocks({ ...migrated, blocks }, 'display').content, contentControls: controls }
}

/** Compatibility wrapper for callers whose story contains no separate SDT ownership payload. */
export function prepareDocumentNoteContent(model: DocumentModel, kind: DocumentNoteKind, id: string,
  input?: DocumentStoryContent, text = '', copy = true): DocumentStoryContent {
  return prepareDocumentNoteContentWithControls(model, kind, id, input, text, copy).content
}

/** Caller owns a detached document candidate; foreign control identities cannot be overwritten. */
export function appendDocumentNoteContentControls(model: DocumentModel, controls: readonly DocumentContentControl[]): void {
  if (!controls.length) return
  const ids = new Set((model.contentControls ?? []).map(control => control.id))
  for (const control of controls) {
    if (ids.has(control.id)) throw new RangeError('注释内容控件的目标标识已存在：' + control.id)
    ids.add(control.id)
  }
  model.contentControls = [...(model.contentControls ?? []), ...controls]
}

export function documentNoteContentText(content: DocumentStoryContent): string {
  return content.blocks !== undefined ? projectDocumentStoryBlocks(content, 'display').text : runText(content.runs)
}
