import type { DocumentBlock, DocumentContentControl, DocumentModel, DocumentNote } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { createRun, documentId, runText, tableForBlock, type DocumentRun } from '../model/document-model'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentNoteReferenceHosts } from '../notes/document-note-references'
import { documentNoteReferenceRunMatches } from '../model/document-note-reference'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { documentTableDescendantIds } from '../model/document-table-selection'
import { documentFieldLocked, documentTextFieldRanges, documentTextFields, setDocumentTextFields } from '../model/document-text-fields'
import { migrateDocumentStoryBlocks, projectDocumentStoryBlocks } from '../model/document-story-blocks'
import { documentStoryParagraphs, isDocumentTextBoundary } from '../model/document-story'
import { copyDocumentBlocksForPaste } from './document-block-paste'
import { documentContentControlNoteBoundaryRanks, partitionDocumentContentControlNoteRuns } from '../model/content-control-note-boundary'
import { assertDocumentNoteReferenceBoundaryRank, documentNoteReferenceOffsets, remapDocumentNoteReferenceBoundary } from '../model/document-note-boundary'
import { compareDocumentNoteBoundaries } from '../model/document-note-boundary-edit'

export interface DocumentRepeatingContentControlCopyIds {
  block: Map<string, string>
  run: Map<string, string>
  row: Map<string, string>
  cell: Map<string, string>
  field: Map<string, string>
  control: Map<string, string>
  note: Map<string, string>
  comment: Map<string, string>
}

export interface DocumentRepeatingContentControlInlineCopyRange {
  blockId: string
  startOffset: number
  endOffset: number
  startNoteReferenceCount?: number
  endNoteReferenceCount?: number
  insertionOffset: number
  insertionNoteReferenceCount?: number
}

/** Extract a borrowed inline source; the common clipboard copier still allocates every new ID. */
export function documentRepeatingContentControlInlineCopySource(block: DocumentBlock, start: number, end: number,
  startNoteReferenceCount?: number, endNoteReferenceCount?: number): DocumentBlock {
  const text = runText(block.runs)
  if (!isDocumentTextBoundary(text, start) || !isDocumentTextBoundary(text, end) || end < start) throw new RangeError('重复项目复制范围不是合法的 UTF-16 区间。')
  const range = { startOffset: start, endOffset: end, startNoteReferenceCount, endNoteReferenceCount }
  const references = documentNoteReferenceOffsets(block.runs), boundary = documentContentControlNoteBoundaryRanks(range, references)
  const runs = cloneModel(partitionDocumentContentControlNoteRuns(block.runs, range).selected)
  const source = { ...cloneModel(block), runs }
  delete source.metadata.html
  const local = documentTextFields(block).filter(field => field.endBlockId === undefined || field.endBlockId === block.id)
  const selected = new Set<string>()
  const contained = local.map(field => ({ field,
    start: assertDocumentNoteReferenceBoundaryRank(references, field.startOffset, field.startNoteReferenceCount),
    end: assertDocumentNoteReferenceBoundaryRank(references, field.endOffset, field.endNoteReferenceCount) })).filter(entry =>
    compareDocumentNoteBoundaries(start, boundary.start, entry.field.startOffset, entry.start) <= 0
    && compareDocumentNoteBoundaries(entry.field.endOffset, entry.end, end, boundary.end) <= 0)
  for (const { field, start: first, end: last } of contained) if (compareDocumentNoteBoundaries(field.startOffset, first, field.endOffset, last) < 0
    || compareDocumentNoteBoundaries(start, boundary.start, field.startOffset, first) < 0
      && compareDocumentNoteBoundaries(field.endOffset, last, end, boundary.end) < 0) selected.add(field.id)
  let changed = true
  while (changed) {
    changed = false
    for (const { field } of contained) {
      const source = field.metadata?.source
      if (source && typeof source === 'object' && !Array.isArray(source) && typeof Reflect.get(source, 'parentFieldId') === 'string'
        && selected.has(Reflect.get(source, 'parentFieldId')) && !selected.has(field.id)) { selected.add(field.id); changed = true }
    }
  }
  const targetReferences = documentNoteReferenceOffsets(runs)
  const fields = contained.filter(({ field }) => selected.has(field.id)).map(({ field }) => {
    const { startNoteReferenceCount: startRank, endNoteReferenceCount: endRank, ...definition } = cloneModel(field)
    const from = remapDocumentNoteReferenceBoundary(references, targetReferences, field.startOffset, field.startOffset - start, startRank)
    const to = remapDocumentNoteReferenceBoundary(references, targetReferences, field.endOffset, field.endOffset - start, endRank)
    return { ...definition, startOffset: from.offset, endOffset: to.offset,
      ...(from.rank === undefined ? {} : { startNoteReferenceCount: from.rank }), ...(to.rank === undefined ? {} : { endNoteReferenceCount: to.rank }) }
  })
  setDocumentTextFields(source, fields)
  return source
}

export function createDocumentRepeatingContentControlCopyIds(): DocumentRepeatingContentControlCopyIds {
  return { block: new Map(), run: new Map(), row: new Map(), cell: new Map(), field: new Map(),
    control: new Map(), note: new Map(), comment: new Map() }
}

/** The regular clipboard copier owns ID allocation; this records its actual structural correspondence. */
export function indexDocumentRepeatingContentControlCopies(source: readonly DocumentBlock[], copies: readonly DocumentBlock[], ids: DocumentRepeatingContentControlCopyIds): void {
  if (source.length !== copies.length) throw new RangeError('复制内容的结构身份不一致。')
  for (const [index, before] of source.entries()) {
    const after = copies[index]!
    ids.block.set(before.id, after.id)
    if (before.runs.length !== after.runs.length) throw new RangeError('复制内容的文字身份不一致。')
    for (const [run, value] of before.runs.entries()) ids.run.set(value.id, after.runs[run]!.id)
    const oldFields = documentTextFields(before), newFields = documentTextFields(after)
    if (oldFields.length !== newFields.length) throw new RangeError('复制内容的字段身份不一致。')
    for (const [field, value] of oldFields.entries()) ids.field.set(value.id, newFields[field]!.id)
    const oldTable = tableForBlock(before), newTable = tableForBlock(after)
    if (!oldTable && !newTable) continue
    if (!oldTable || !newTable || oldTable.rows.length !== newTable.rows.length) throw new RangeError('复制表格的行身份不一致。')
    for (const [row, oldRow] of oldTable.rows.entries()) {
      const newRow = newTable.rows[row]!
      ids.row.set(oldRow.id, newRow.id)
      if (oldRow.cells.length !== newRow.cells.length) throw new RangeError('复制表格的单元格身份不一致。')
      for (const [column, oldCell] of oldRow.cells.entries()) {
        const newCell = newRow.cells[column]!
        ids.cell.set(oldCell.id, newCell.id)
        const oldContainers = documentTableCellContentContainers(oldCell), newContainers = documentTableCellContentContainers(newCell)
        if (oldContainers.length !== newContainers.length) throw new RangeError('复制表格的续格内容不一致。')
        for (const [offset, blocks] of oldContainers.entries()) indexDocumentRepeatingContentControlCopies(blocks, newContainers[offset]!, ids)
      }
    }
  }
}

export function visitDocumentRepeatingContentControlBlocks(blocks: readonly DocumentBlock[], visit: (block: DocumentBlock) => void): void {
  for (const block of blocks) {
    visit(block)
    for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) {
      for (const container of documentTableCellContentContainers(cell)) visitDocumentRepeatingContentControlBlocks(container, visit)
    }
  }
}

export function assertDocumentRepeatingNoteContentRemovable(model: DocumentModel, note: DocumentNote): void {
  const notes = (['footnotes', 'endnotes'] as const).flatMap(kind => Object.values(model[kind] ?? {}).map(note => ({ kind, note })))
  const pending = [notes.find(entry => entry.note === note) ?? { kind: 'footnotes' as const, note }], seen = new Set<string>(), defaults = documentDefaultStyle(model)
  while (pending.length) {
    const current = pending.pop()!
    const key = `${current.kind}:${current.note.id}`
    if (seen.has(key)) continue
    seen.add(key)
    const content = current.note.content && migrateDocumentStoryBlocks(current.note.content, [],
      { kind: 'note', noteKind: current.kind === 'footnotes' ? 'footnote' : 'endnote', noteId: current.note.id }, defaults)
    const ids = documentTableDescendantIds(content?.blocks ?? [])
    if (content?.locks?.some(lock => lock.lock !== 'unlocked') || model.contentControls?.some(control => ids.has(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)
      && ['controlLocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'].includes(control.lock ?? ''))) {
      throw new RangeError('删除重复项目不能移除脚注或尾注中的锁定内容。')
    }
    if (content?.blocks && documentTextFieldRanges({ blocks: content.blocks }).some(range => documentFieldLocked(range.field))) {
      throw new RangeError('删除重复项目不能移除脚注或尾注中的锁定字段。')
    }
    const blocks = new Map<string, DocumentBlock>()
    visitDocumentRepeatingContentControlBlocks(content?.blocks ?? [], block => { blocks.set(block.id, block) })
    for (const nested of notes) if (!seen.has(`${nested.kind}:${nested.note.id}`) && hasNoteReference(blocks, nested.note, nested.kind)) pending.push(nested)
  }
}

function noteBlocksById(model: DocumentModel): Map<string, DocumentBlock> {
  const blocks = new Map<string, DocumentBlock>(), runs = new Set<string>()
  for (const owner of documentNoteReferenceHosts(model)) {
    visitDocumentRepeatingContentControlBlocks(owner.blocks, block => {
      if (blocks.has(block.id)) throw new RangeError('注释引用块不能被多个内容区共享：' + block.id)
      blocks.set(block.id, block)
      for (const run of block.runs) {
        if (runs.has(run.id)) throw new RangeError('注释引用字符不能被多个内容区共享：' + run.id)
        runs.add(run.id)
      }
    })
  }
  return blocks
}

function hasNoteReference(blocks: ReadonlyMap<string, DocumentBlock>, note: DocumentNote, kind: 'footnotes' | 'endnotes'): boolean {
  const matches = (block: DocumentBlock) => block.runs.some(run => documentNoteReferenceRunMatches(run, note, kind === 'footnotes' ? 'footnote' : 'endnote'))
  const owners = [...blocks.values()].filter(matches)
  if (owners.length > 1 || (owners[0]?.runs.filter(run => documentNoteReferenceRunMatches(run, note, kind === 'footnotes' ? 'footnote' : 'endnote')).length ?? 0) > 1) {
    throw new RangeError('注释引用没有唯一的真实字符或宿主身份：' + note.id)
  }
  return owners.length === 1
}

export function copyDocumentRepeatingContentControlNotes(before: DocumentModel, after: DocumentModel, ids: DocumentRepeatingContentControlCopyIds): void {
  const pending = (['footnotes', 'endnotes'] as const).flatMap(kind => Object.values(before[kind] ?? {}).map(note => ({ kind, note })))
  const sourceBlocks = noteBlocksById(before), blocks = noteBlocksById(after), copiedKeys = new Set<string>(), copiedNotes: DocumentNote[] = [], defaults = documentDefaultStyle(before)
  const sourceRunIds = new Set([...sourceBlocks.values()].flatMap(block => block.runs.map(run => run.id)))
  const sourceCounts = new Map<string, number>()
  for (const { note } of pending) sourceCounts.set(note.id, (sourceCounts.get(note.id) ?? 0) + 1)
  let changed = true, copiedBlockCount = ids.block.size, copiedTextLength = 0
  while (changed) {
    changed = false
    for (const { kind, note } of pending) {
      const key = `${kind}:${note.id}`
      if (copiedKeys.has(key)) continue
      const noteKind = kind === 'footnotes' ? 'footnote' as const : 'endnote' as const
      const originals = [...sourceBlocks.values()].flatMap(block => block.runs
        .filter(run => documentNoteReferenceRunMatches(run, note, noteKind)).map(run => ({ block, run })))
      const mapped = originals.filter(entry => ids.run.has(entry.run.id))
      if (mapped.length && originals.length !== 1) throw new RangeError('复制注释没有唯一的来源字符身份：' + note.id)
      const original = mapped[0]
      if (!original) continue
      const runId = ids.run.get(original.run.id)
      let blockId = ids.block.get(original.block.id)
      if (!blockId && runId) {
        const owners = [...new Set(blocks.values())].filter(block => block.runs.some(run => run.id === runId))
        if (owners.length !== 1) throw new RangeError('复制旧注释缺少唯一的真实引用宿主：' + note.id)
        blockId = owners[0]!.id
        ids.block.set(original.block.id, blockId)
      }
      if (!blockId || !runId) continue
      if (note.referenceRunId && (note.referenceRunId === original.run.id || !sourceRunIds.has(note.referenceRunId))) {
        const existing = ids.run.get(note.referenceRunId)
        if (existing && existing !== runId) throw new RangeError('复制注释不能拆分引用字符的来源身份。')
        ids.run.set(note.referenceRunId, runId)
      }
      if (note.referenceBlockId && !sourceBlocks.has(note.referenceBlockId)) {
        const existing = ids.block.get(note.referenceBlockId)
        if (existing && existing !== blockId) throw new RangeError('复制注释的旧宿主别名与实际结构身份冲突。')
        ids.block.set(note.referenceBlockId, blockId)
      }
      const reference = blocks.get(blockId)?.runs.find(run => run.id === runId)
      if (!reference) throw new RangeError('复制注释缺少真实的新引用字符：' + note.id)
      if (!documentNoteReferenceRunMatches(reference, { ...note, referenceBlockId: blockId, referenceRunId: runId }, noteKind)) {
        throw new RangeError('复制注释的引用类型或身份不一致：' + note.id)
      }
      const id = documentId(kind === 'footnotes' ? 'footnote' : 'endnote'), copied = cloneModel(note)
      // Office may reuse a numeric ID across the two note kinds; keep both scoped identities.
      ids.note.set(key, id)
      if (sourceCounts.get(note.id) === 1) ids.note.set(note.id, id)
      copied.id = id; copied.referenceBlockId = blockId; copied.referenceRunId = runId
      const legacy = copied.content ?? { runs: [createRun(copied.text, defaults.text)], paragraphs: documentStoryParagraphs(copied.text, defaults.paragraph) }
      const source = migrateDocumentStoryBlocks(legacy, [], { kind: 'note', noteKind: kind === 'footnotes' ? 'footnote' : 'endnote', noteId: note.id }, defaults)
      copiedBlockCount += documentTableDescendantIds(source.blocks!).size
      copiedTextLength += runText(source.runs).length
      if (copiedBlockCount > 250_000 || copiedTextLength > 16 * 1024 * 1024) throw new RangeError('重复项目注释复制超过内容预算。')
      const copies = copyDocumentBlocksForPaste(source.blocks!)
      indexDocumentRepeatingContentControlCopies(source.blocks!, copies, ids)
      copied.content = projectDocumentStoryBlocks({ ...source, blocks: copies }, 'display').content
      copied.text = runText(copied.content.runs)
      const collection = after[kind] ?? (after[kind] = {})
      collection[id] = copied
      visitDocumentRepeatingContentControlBlocks(copies, block => { blocks.set(block.id, block) })
      const metadata = { ...reference.metadata }
      for (const key of ['note', 'noteReference']) {
        const marker = metadata[key]
        if (marker === undefined || marker === null) continue
        const next = marker && typeof marker === 'object' && !Array.isArray(marker) ? { ...marker } : {}
        for (const identity of ['noteId', 'id', 'footnoteId', 'endnoteId']) if (Reflect.get(next, identity) !== undefined) Reflect.set(next, identity, id)
        for (const identity of ['type', 'kind', 'noteType']) if (Reflect.get(next, identity) !== undefined) Reflect.set(next, identity, noteKind)
        metadata[key] = { ...next, type: noteKind, noteId: id }
      }
      reference.metadata = { ...metadata, note: { ...(metadata.note && typeof metadata.note === 'object' && !Array.isArray(metadata.note) ? metadata.note : {}), type: noteKind, noteId: id } }
      copiedKeys.add(key); copiedNotes.push(copied); changed = true
    }
  }
  // Nested references are remapped after their parent body was projected. Rebuild that derived bridge once.
  for (const note of copiedNotes) {
    const projection = projectDocumentStoryBlocks(note.content!, 'display')
    note.content = projection.content; note.text = projection.text
  }
}

export function remapDocumentRepeatingContentControlMetadata(metadata: Record<string, unknown> | undefined, ids: DocumentRepeatingContentControlCopyIds): Record<string, unknown> | undefined {
  if (!metadata) return metadata
  const next = { ...metadata }
  for (const key of ['anchorBlockId', 'tailBlockId', 'pictureBlockId', 'pictureAnchorBlockId', 'pictureTailBlockId'] as const) {
    const value = next[key]
    if (typeof value === 'string' && ids.block.has(value)) next[key] = ids.block.get(value)!
  }
  if (typeof next.contentControlId === 'string' && ids.control.has(next.contentControlId)) next.contentControlId = ids.control.get(next.contentControlId)!
  return next
}

/** A scoped text planner has no foreign note controls; add their complete, copied ownership once. */
export function copyDocumentRepeatingContentControlNoteControls(before: DocumentModel, after: DocumentModel, ids: DocumentRepeatingContentControlCopyIds): void {
  const sourceIds = new Set<string>(), defaults = documentDefaultStyle(before)
  for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(before[kind] ?? {})) {
    if (!note.content) continue
    const source = migrateDocumentStoryBlocks(note.content, [], { kind: 'note', noteKind: kind === 'footnotes' ? 'footnote' : 'endnote', noteId: note.id }, defaults)
    for (const id of documentTableDescendantIds(source.blocks!)) sourceIds.add(id)
  }
  const selected = (before.contentControls ?? []).filter(control => !control.metadata?.officeStoryLockSource && !ids.control.has(control.id)
    && sourceIds.has(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)
    && ids.block.get(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)
      !== (control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId)
    && (control.blockRange ? control.blockRange.blockIds.every(id => ids.block.has(id))
      : control.tableRowRange ? ids.block.has(control.tableRowRange.tableId) && control.tableRowRange.rowIds.every(id => ids.row.has(id)) : ids.block.has(control.blockId)))
  for (const control of selected) ids.control.set(control.id, documentId('content-control'))
  const copies: DocumentContentControl[] = selected.map(control => {
    const next = cloneModel(control), blockId = ids.block.get(control.blockId)
    if (!blockId) throw new RangeError('复制注释控件缺少真实的输入锚点：' + control.id)
    next.id = ids.control.get(control.id)!; next.blockId = blockId
    if (control.parentId) {
      const parent = ids.control.get(control.parentId)
      if (!parent) throw new RangeError('复制注释控件缺少同一内容区的父控件：' + control.id)
      next.parentId = parent
    }
    if (next.blockRange) next.blockRange = { blockIds: next.blockRange.blockIds.map(id => ids.block.get(id)!) }
    if (next.tableRowRange) next.tableRowRange = { tableId: ids.block.get(next.tableRowRange.tableId)!, rowIds: next.tableRowRange.rowIds.map(id => ids.row.get(id)!) }
    next.metadata = remapDocumentRepeatingContentControlMetadata(next.metadata, ids)
    return next
  })
  after.contentControls = [...(after.contentControls ?? []), ...copies]
  const newIds = new Set([...ids.block].filter(([source, destination]) => source !== destination).map(([, destination]) => destination))
  for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(after[kind] ?? {})) {
    if (!note.content?.blocks || !note.content.blocks.some(block => newIds.has(block.id))) continue
    visitDocumentRepeatingContentControlBlocks(note.content.blocks, block => {
      if (newIds.has(block.id)) block.metadata = remapDocumentRepeatingContentControlMetadata(block.metadata, ids)!
    })
    const projection = projectDocumentStoryBlocks(note.content, 'display')
    note.content = projection.content; note.text = projection.text
  }
}

export function copyDocumentRepeatingContentControlComments(before: DocumentModel, after: DocumentModel, ids: DocumentRepeatingContentControlCopyIds,
  inline?: DocumentRepeatingContentControlInlineCopyRange): void {
  const paragraphs = noteBlocksById(before)
  for (const comment of Object.values(before.comments ?? {})) {
    const anchor = comment.anchor.flatMap(segment => {
      const blockId = ids.block.get(segment.blockId)
      if (!blockId) return []
      const source = paragraphs.get(segment.blockId), sourceText = source && runText(source.runs)
      if (sourceText === undefined || !isDocumentTextBoundary(sourceText, segment.startOffset) || !isDocumentTextBoundary(sourceText, segment.endOffset)
        || segment.endOffset < segment.startOffset) throw new RangeError('复制批注的原始锚点不是合法的 UTF-16 区间：' + comment.id)
      if (!inline || segment.blockId !== inline.blockId) return [{ ...segment, blockId }]
      const start = Math.max(segment.startOffset, inline.startOffset), end = Math.min(segment.endOffset, inline.endOffset)
      const empty = segment.startOffset === segment.endOffset && start === end && start > inline.startOffset && start < inline.endOffset
      if (start >= end && !empty) return []
      const text = sourceText.slice(start, end)
      return [{ ...segment, blockId, startOffset: inline.insertionOffset + start - inline.startOffset,
        endOffset: inline.insertionOffset + end - inline.startOffset, text }]
    })
    if (!anchor.length) continue
    const id = documentId('thread-comment')
    ids.comment.set(comment.id, id)
    const comments = after.comments ?? (after.comments = {})
    comments[id] = { ...cloneModel(comment), id, anchor, referenceText: anchor.map(segment => segment.text).join('\n'),
      messages: comment.messages.map(message => ({ ...cloneModel(message), id: documentId('comment-message') })) }
  }
}

/** A removed reference removes its note scope and anchors in the same candidate. */
export function removeDeletedDocumentRepeatingNoteScopes(before: DocumentModel, model: DocumentModel): void {
  const removed = new Set<string>(), source = noteBlocksById(before), checked = new Set<string>(), defaults = documentDefaultStyle(before)
  let changed = true
  while (changed) {
    changed = false
    const current = noteBlocksById(model)
    for (const kind of ['footnotes', 'endnotes'] as const) for (const note of Object.values(before[kind] ?? {})) {
      const key = `${kind}:${note.id}`
      if (checked.has(key) || !hasNoteReference(source, note, kind) || hasNoteReference(current, note, kind)) continue
      assertDocumentRepeatingNoteContentRemovable(before, note)
      const content = note.content && migrateDocumentStoryBlocks(note.content, [],
        { kind: 'note', noteKind: kind === 'footnotes' ? 'footnote' : 'endnote', noteId: note.id }, defaults)
      for (const id of documentTableDescendantIds(content?.blocks ?? [])) removed.add(id)
      if (model[kind]) delete model[kind]![note.id]
      checked.add(key); changed = true
    }
  }
  if (!removed.size) return
  model.contentControls = model.contentControls?.filter(control => !removed.has(control.blockRange?.blockIds[0] ?? control.tableRowRange?.tableId ?? control.blockId))
  for (const [id, comment] of Object.entries(model.comments ?? {})) {
    const anchor = comment.anchor.filter(segment => !removed.has(segment.blockId))
    if (anchor.length) model.comments![id] = { ...comment, anchor, referenceText: anchor.map(segment => segment.text).join('\n') }
    else delete model.comments![id]
  }
}
