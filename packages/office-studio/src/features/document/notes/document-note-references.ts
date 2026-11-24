import type { DocumentBlock, DocumentModel, DocumentNote } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentBlockHostId, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentBlockAncestry, runText, tableForBlock, textTargets, type DocumentPoint, type DocumentRun } from '../model/document-model'
import { documentNoteReferenceRunMatches } from '../model/document-note-reference'
import { DOCUMENT_STORY_KEYS, documentStoryDefaultStyle, isDocumentTextBoundary } from '../model/document-story'
import { migrateDocumentStoryBlocks } from '../model/document-story-blocks'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { DEFAULT_DOCUMENT_SECTION_ID, sectionIdForBlock, sectionStoryLinkedToPrevious } from '../services/section-header-footer'
import type { DocumentNoteKind } from './document-note-types'

export interface DocumentNoteReferenceHost {
  blocks: DocumentBlock[]
  host?: DocumentAuxiliaryBlockHost
  /** Linked sections may retain inactive local source blocks; these have ownership but no editable view. */
  linked?: boolean
}

export interface DocumentNoteReferenceRecord {
  kind: DocumentNoteKind
  note: DocumentNote
  block: DocumentBlock
  run: DocumentRun
  runIndex: number
  offset: number
  host?: DocumentAuxiliaryBlockHost
  sectionId: string
  point?: DocumentPoint
  linked?: boolean
}

/** Enumerate actual stored ownership, never duplicate a resolved linked story. */
export function documentNoteReferenceHosts(model: DocumentModel): DocumentNoteReferenceHost[] {
  const result: DocumentNoteReferenceHost[] = [{ blocks: model.blocks }]
  for (const [sectionId, section] of Object.entries(model.sectionHeaderFooters ?? {})) for (const story of DOCUMENT_STORY_KEYS) {
    const content = section.storyContent?.[story]
    if (!content) continue
    const host: DocumentAuxiliaryBlockHost = { kind: 'story', sectionId, story }, linked = sectionStoryLinkedToPrevious(section, story)
    if (linked && content.blocks === undefined) continue
    const blocks = content.blocks ?? migrateDocumentStoryBlocks(cloneModel(content), (section.pageFields ?? []).filter(field => field.story === story), host, documentStoryDefaultStyle(model)).blocks!
    result.push({ blocks, host, linked })
  }
  for (const kind of ['footnote', 'endnote'] as const) for (const note of Object.values(kind === 'footnote' ? model.footnotes ?? {} : model.endnotes ?? {})) {
    if (!note.content) continue
    const host: DocumentAuxiliaryBlockHost = { kind: 'note', noteKind: kind, noteId: note.id }
    result.push({ host, blocks: note.content.blocks ?? migrateDocumentStoryBlocks(cloneModel(note.content), [], host, documentDefaultStyle(model)).blocks! })
  }
  return result
}

function paragraphs(hosts: readonly DocumentNoteReferenceHost[]): Array<{ owner: DocumentNoteReferenceHost; block: DocumentBlock }> {
  const result: Array<{ owner: DocumentNoteReferenceHost; block: DocumentBlock }> = []
  const identities = new Map<string, string>(), runIds = new Set<string>()
  const visit = (owner: DocumentNoteReferenceHost, blocks: DocumentBlock[], depth: number) => {
    if (depth > 64) throw new RangeError('注释引用超过内容区嵌套预算。')
    for (const block of blocks) {
      const hostId = owner.host ? documentBlockHostId(owner.host) : 'body'
      if (identities.has(block.id)) throw new RangeError('注释引用块被多个位置或内容区共享：' + block.id)
      identities.set(block.id, hostId)
      for (const run of block.runs) {
        if (runIds.has(run.id)) throw new RangeError('注释引用字符被多个位置或内容区共享：' + run.id)
        runIds.add(run.id)
      }
      result.push({ owner, block })
      if (result.length > 250_000) throw new RangeError('注释引用超过单次块预算。')
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) visit(owner, content, depth + 1)
    }
  }
  for (const owner of hosts) visit(owner, owner.blocks, 0)
  return result
}

/** Stable note/run identity repairs stale pointers; a reused ID or multiple owners is rejected. */
export function documentNoteReferenceRecords(model: DocumentModel, onlyKind?: DocumentNoteKind): DocumentNoteReferenceRecord[] {
  const entries = paragraphs(documentNoteReferenceHosts(model)), result: DocumentNoteReferenceRecord[] = []
  const pending = (['footnote', 'endnote'] as const).filter(kind => onlyKind === undefined || kind === onlyKind)
    .flatMap(kind => Object.entries(kind === 'footnote' ? model.footnotes ?? {} : model.endnotes ?? {}).map(([key, note]) => ({ kind, key, note })))
  for (const { kind, key, note } of pending) {
    if (note.id !== key) throw new RangeError('注释集合标识与内容身份不一致：' + key)
    const matches = entries.flatMap(({ owner, block }) => block.runs.flatMap((run, runIndex) => documentNoteReferenceRunMatches(run, note, kind) ? [{ owner, block, run, runIndex }] : []))
    if (matches.length > 1) throw new RangeError('注释引用没有唯一的真实字符或内容宿主：' + note.id)
    const match = matches[0]
    if (!match) continue
    const offset = match.block.runs.slice(0, match.runIndex).reduce((offset, run) => offset + run.text.length, 0)
    if (!isDocumentTextBoundary(runText(match.block.runs), offset)) throw new RangeError('注释引用起点不是合法的 UTF-16 边界。')
    const targets = textTargets({ blocks: match.owner.blocks })
    const target = targets.find(target => (target.cellBlockId ?? target.blockId) === match.block.id)
    const rootId = documentBlockAncestry({ blocks: match.owner.blocks }, match.block.id)?.[0]?.id ?? match.block.id
    const sectionId = match.owner.host?.kind === 'story' ? match.owner.host.sectionId
      : match.owner.host?.kind === 'note' ? DEFAULT_DOCUMENT_SECTION_ID : sectionIdForBlock(model, rootId)
    result.push({ kind, note, block: match.block, run: match.run as DocumentRun, runIndex: match.runIndex, offset,
      sectionId, ...(match.owner.host ? { host: { ...match.owner.host } } : {}), ...(match.owner.linked ? { linked: true } : {}),
      ...(target ? { point: { ...target, offset } } : {}) })
  }
  const parentSection = (record: DocumentNoteReferenceRecord, seen = new Set<string>()): string => {
    const host = record.host
    if (host?.kind !== 'note') return record.sectionId
    const key = `${host.noteKind}:${host.noteId}`
    if (seen.has(key)) throw new RangeError('注释引用形成循环内容区。')
    seen.add(key)
    const owner = result.find(parent => parent.kind === host.noteKind && parent.note.id === host.noteId)
    return owner ? parentSection(owner, seen) : record.sectionId
  }
  // Include both kinds when resolving a legacy note nested inside another kind.
  if (onlyKind === undefined) for (const record of result) record.sectionId = parentSection(record)
  return result
}

export function documentNoteReferenceLocation(model: DocumentModel, kind: DocumentNoteKind, id: string): { point: DocumentPoint; host?: DocumentAuxiliaryBlockHost; sectionId: string } | undefined {
  const reference = documentNoteReferenceRecords(model).find(reference => reference.kind === kind && reference.note.id === id)
  if (!reference?.point || reference.linked) return undefined
  return { point: { ...reference.point }, ...(reference.host ? { host: { ...reference.host } } : {}), sectionId: reference.sectionId }
}

/** Public order follows main-body references; detached data cannot mutate the document. */
export function getDocumentNotesInReferenceOrder(model: DocumentModel, kind: DocumentNoteKind): DocumentNote[] {
  const order = paragraphs([{ blocks: model.blocks }]).map(entry => entry.block.id)
  return documentNoteReferenceRecords(model).filter(reference => reference.kind === kind && !reference.host)
    .sort((left, right) => order.indexOf(left.block.id) - order.indexOf(right.block.id) || left.offset - right.offset)
    .map(reference => cloneModel(reference.note))
}
