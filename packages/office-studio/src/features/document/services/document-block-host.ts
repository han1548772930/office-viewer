import type { DocumentContentControl, DocumentModel, DocumentStoryContent, DocumentStoryLock, DocumentThreadComment } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentDefaultStyle } from '../model/document-default-style'
import { documentBlockHostId, documentStoredBlockHosts, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentNoteDecorationContent, setDocumentNoteDecorationContent, DOCUMENT_NOTE_DECORATION_HOST_KEYS } from '../model/document-note-decoration-host'
import { documentBlockAncestry, runText, textTargets, type DocumentPoint, type DocumentSelection } from '../model/document-model'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { DOCUMENT_STORY_KEYS, documentStoryContent, documentStoryDefaultStyle, documentStoryParagraphs, type DocumentStoryKey } from '../model/document-story'
import { documentStoryDisplayToInputOffset, documentStoryInputToDisplayOffset, migrateDocumentStoryBlocks, projectDocumentStoryBlocks, validateDocumentStoryBlocks } from '../model/document-story-blocks'
import { DEFAULT_DOCUMENT_SECTION_ID, documentSectionIds, resolveSectionHeaderFooter, sectionIdForBlock, sectionStoryLinkedToPrevious } from './section-header-footer'
import { createRun } from '../model/document-model'
import { defaultDocumentNoteDecoration } from '../notes/document-note-decorations'

export interface DocumentBlockHostView {
  host: DocumentAuxiliaryBlockHost
  sourceHost: DocumentAuxiliaryBlockHost
  content: DocumentStoryContent
  sectionId: string
  linked: boolean
}

export function validateDocumentBlockHost(host: DocumentAuxiliaryBlockHost): void {
  if (!host || typeof host !== 'object' || Array.isArray(host)) throw new TypeError('内容区地址无效。')
  if (host.kind === 'story') {
    if (typeof host.sectionId !== 'string' || !host.sectionId.trim() || !DOCUMENT_STORY_KEYS.includes(host.story)) throw new TypeError('页眉页脚块内容区地址无效。')
  } else if (host.kind === 'noteDecoration') {
    if (!['footnote', 'endnote'].includes(host.noteKind) || !DOCUMENT_NOTE_DECORATION_HOST_KEYS.includes(host.decoration)
      || host.dormantIndex !== undefined && (!Number.isSafeInteger(host.dormantIndex) || host.dormantIndex < 0)) throw new TypeError('注释分隔内容区地址无效。')
  } else if (host.kind !== 'note' || !['footnote', 'endnote'].includes(host.noteKind) || typeof host.noteId !== 'string' || !host.noteId.trim()) throw new TypeError('脚注或尾注块内容区地址无效。')
}

export function resolveDocumentBlockHost(model: DocumentModel, host: DocumentAuxiliaryBlockHost): DocumentBlockHostView {
  validateDocumentBlockHost(host)
  if (host.kind === 'story') {
    const sections = documentSectionIds(model), index = sections.indexOf(host.sectionId)
    if (index < 0) throw new RangeError('页眉页脚块内容区不存在。')
    const section = resolveSectionHeaderFooter(model, host.sectionId), linked = sectionStoryLinkedToPrevious(section, host.story)
    let sourceIndex = index
    while (sourceIndex > 0 && sectionStoryLinkedToPrevious(resolveSectionHeaderFooter(model, sections[sourceIndex]!), host.story)) sourceIndex--
    const sourceHost: DocumentAuxiliaryBlockHost = { ...host, sectionId: sections[sourceIndex]! }
    const content = migrateDocumentStoryBlocks(documentStoryContent(section, host.story, model),
      (section.pageFields ?? []).filter(field => field.story === host.story), sourceHost, documentStoryDefaultStyle(model))
    return { host: { ...host }, sourceHost, content, sectionId: host.sectionId, linked }
  }
  if (host.kind === 'noteDecoration') {
    const source = documentNoteDecorationContent(model, host) ?? (host.dormantIndex === undefined
      ? defaultDocumentNoteDecoration(model, host.noteKind, host.decoration, true) : undefined)
    if (!source) throw new RangeError('注释分隔内容区不存在。')
    const content = migrateDocumentStoryBlocks(source, [], host, documentDefaultStyle(model))
    return { host: { ...host }, sourceHost: { ...host }, content, sectionId: DEFAULT_DOCUMENT_SECTION_ID, linked: false }
  }
  const note = (host.noteKind === 'footnote' ? model.footnotes : model.endnotes)?.[host.noteId]
  if (!note) throw new RangeError('脚注或尾注块内容区不存在。')
  const defaults = documentDefaultStyle(model)
  const legacy = note.content ?? { runs: [createRun(note.text, defaults.text)], paragraphs: documentStoryParagraphs(note.text, defaults.paragraph) }
  const content = migrateDocumentStoryBlocks(legacy, [], host, defaults)
  const ancestry = documentBlockAncestry(model, note.referenceBlockId)
  return { host: { ...host }, sourceHost: { ...host }, content, sectionId: sectionIdForBlock(model, ancestry?.[0]?.id ?? note.referenceBlockId), linked: false }
}

function blockIds(blocks: NonNullable<DocumentStoryContent['blocks']>): Set<string> {
  const ids = new Set(documentTextFieldParagraphs({ blocks }).map(entry => entry.block.id))
  const visit = (content: NonNullable<DocumentStoryContent['blocks']>) => {
    for (const block of content) {
      ids.add(block.id)
      const table = block.metadata.table as import('../model/document-model').DocumentTable | undefined
      for (const row of table?.rows ?? []) for (const cell of row.cells) {
        visit(cell.blocks)
        for (const continuation of Object.values(cell.style?.nativeContinuationContent ?? {})) visit(continuation.blocks)
      }
    }
  }
  visit(blocks)
  return ids
}

function legacyStoryLockControls(view: DocumentBlockHostView): DocumentContentControl[] {
  const content = view.content, ids = blockIds(content.blocks!), projection = projectDocumentStoryBlocks(content, view.host.kind === 'story' ? 'literal' : 'display')
  const paragraphs = documentTextFieldParagraphs({ blocks: content.blocks! }), result: DocumentContentControl[] = []
  for (const lock of content.locks ?? []) {
    const start = documentStoryInputToDisplayOffset(projection, lock.startOffset), end = documentStoryInputToDisplayOffset(projection, lock.endOffset, 'end')
    const selected = paragraphs.filter(paragraph => start === end ? start >= paragraph.absoluteStart && start <= paragraph.absoluteEnd
      : start <= paragraph.absoluteEnd && end >= paragraph.absoluteStart)
    for (const paragraph of selected) {
      if (!ids.has(paragraph.block.id)) continue
      result.push({ id: `story-lock:${encodeURIComponent(documentBlockHostId(view.sourceHost))}:${encodeURIComponent(lock.id)}:${paragraph.block.id}`,
        type: 'richText', blockId: paragraph.block.id, startOffset: Math.max(0, start - paragraph.absoluteStart),
        endOffset: Math.max(0, Math.min(paragraph.text.length, end - paragraph.absoluteStart)), lock: lock.lock,
        placement: 'inline', metadata: { officeStoryLockSource: { version: 1, id: lock.id, hostId: documentBlockHostId(view.sourceHost) } } })
    }
  }
  return result
}

/** The same command/model/table engine works on a scoped view, with document-level resources kept outside it. */
export function documentModelForBlockHost(model: DocumentModel, host: DocumentAuxiliaryBlockHost, forEdit = false): DocumentModel {
  const view = resolveDocumentBlockHost(model, host)
  if (forEdit && view.linked) throw new Error('请先取消“链接到前一节”再编辑页眉页脚块内容。')
  const blocks = view.content.blocks!, ids = blockIds(blocks), sourceId = documentBlockHostId(view.sourceHost)
  const controls = (model.contentControls ?? []).filter(control => ids.has(control.blockId))
  for (const control of legacyStoryLockControls(view)) if (!controls.some(candidate => candidate.id === control.id)) controls.push(control)
  const comments: Record<string, DocumentThreadComment> = {}
  for (const [id, comment] of Object.entries(model.comments ?? {})) {
    const anchor = comment.anchor.filter(segment => ids.has(segment.blockId))
    if (anchor.length) comments[id] = { ...comment, anchor }
  }
  return { ...model, blocks, defaultStyle: host.kind === 'story' ? documentStoryDefaultStyle(model) : documentDefaultStyle(model),
    sectionSetup: undefined, sectionHeaderFooters: undefined, footnotes: undefined, endnotes: undefined,
    contentControls: controls, comments,
    metadata: { ...model.metadata, documentEditingBlockHost: { ...host }, documentEditingBlockSourceId: sourceId, documentEditingSectionId: view.sectionId,
      ...(host.kind === 'noteDecoration' ? { documentEditingSectionIds: documentSectionIds(model) } : {}) } }
}

/** Flattened offsets are compatibility queries; actual editing targets remain block/cell coordinates. */
export function documentPointAtBlockHostOffset(model: DocumentModel, host: DocumentAuxiliaryBlockHost, offset: number): DocumentPoint {
  const scoped = documentModelForBlockHost(model, host), paragraphs = documentTextFieldParagraphs(scoped)
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > (paragraphs.at(-1)?.absoluteEnd ?? 0)) throw new RangeError('内容区文字偏移无效。')
  const paragraph = paragraphs.find(candidate => offset <= candidate.absoluteEnd) ?? paragraphs.at(-1)
  const target = paragraph && textTargets(scoped).find(candidate => (candidate.cellBlockId ?? candidate.blockId) === paragraph.block.id)
  if (!target) throw new RangeError('内容区没有可输入段落。')
  return { ...target, offset: Math.max(0, offset - paragraph!.absoluteStart) }
}

export function documentBlockHostSelectionFromLegacyRange(model: DocumentModel, sectionId: string, story: DocumentStoryKey, start: number, end = start): DocumentSelection {
  const host: DocumentAuxiliaryBlockHost = { kind: 'story', sectionId, story }, view = resolveDocumentBlockHost(model, host)
  const projection = projectDocumentStoryBlocks(view.content)
  return { anchor: documentPointAtBlockHostOffset(model, host, documentStoryInputToDisplayOffset(projection, Math.min(start, end))),
    focus: documentPointAtBlockHostOffset(model, host, documentStoryInputToDisplayOffset(projection, Math.max(start, end), start === end ? 'start' : 'end')) }
}

function projectedStoryLocks(content: DocumentStoryContent, scoped: DocumentModel, sourceHost: DocumentAuxiliaryBlockHost): DocumentStoryLock[] | undefined {
  if (!content.locks?.length) return content.locks
  const projection = projectDocumentStoryBlocks(content, sourceHost.kind === 'story' ? 'literal' : 'display')
  const paragraphs = new Map(documentTextFieldParagraphs(scoped).map(paragraph => [paragraph.block.id, paragraph]))
  return content.locks.map(lock => {
    const anchors = (scoped.contentControls ?? []).flatMap(control => {
      const raw = control.metadata?.officeStoryLockSource, paragraph = paragraphs.get(control.blockId)
      if (!paragraph || !raw || typeof raw !== 'object' || Reflect.get(raw, 'id') !== lock.id || Reflect.get(raw, 'hostId') !== documentBlockHostId(sourceHost)) return []
      return [{ start: paragraph.absoluteStart + control.startOffset, end: paragraph.absoluteStart + control.endOffset }]
    })
    if (!anchors.length) throw new Error(`内容区编辑不能删除内容锁：${lock.id}`)
    return { ...lock, startOffset: documentStoryDisplayToInputOffset(projection, Math.min(...anchors.map(anchor => anchor.start))),
      endOffset: documentStoryDisplayToInputOffset(projection, Math.max(...anchors.map(anchor => anchor.end)), 'end') }
  })
}

/** Commit only this host and its anchors into the caller's detached full-document candidate. */
export function applyDocumentBlockHostModel(model: DocumentModel, host: DocumentAuxiliaryBlockHost, scoped: DocumentModel): void {
  const previous = resolveDocumentBlockHost(model, host)
  if (previous.linked) throw new Error('请先取消“链接到前一节”再编辑页眉页脚块内容。')
  validateDocumentStoryBlocks(scoped.blocks)
  const oldIds = blockIds(previous.content.blocks!), nextIds = blockIds(scoped.blocks), allIds = new Set([...oldIds, ...nextIds])
  // Stable IDs have one owner. Copying between hosts must use the regular paste ID-remapping planner.
  const foreignIds = new Set<string>()
  for (const id of blockIds(model.blocks)) if (!oldIds.has(id)) foreignIds.add(id)
  for (const entry of documentStoredBlockHosts(model)) if (documentBlockHostId(entry.host) !== documentBlockHostId(previous.sourceHost)) for (const id of blockIds(entry.blocks)) foreignIds.add(id)
  if ([...nextIds].some(id => foreignIds.has(id))) throw new TypeError('块标识不能被多个内容区共享。')
  let content: DocumentStoryContent = { ...previous.content, blocks: scoped.blocks }
  content = projectDocumentStoryBlocks(content, host.kind === 'story' ? 'literal' : 'display').content
  content.locks = projectedStoryLocks(content, scoped, previous.sourceHost)
  if (host.kind === 'story') {
    const section = resolveSectionHeaderFooter(model, host.sectionId), projection = projectDocumentStoryBlocks(content)
    section[host.story] = projection.text
    section.storyContent = { ...section.storyContent, [host.story]: content }
    section.pageFields = [...(section.pageFields ?? []).filter(field => field.story !== host.story), ...projection.pageFields.map(field => ({ ...field, story: host.story }))]
    model.sectionHeaderFooters = { ...model.sectionHeaderFooters, [host.sectionId]: section }
  } else if (host.kind === 'note') {
    const property = host.noteKind === 'footnote' ? 'footnotes' : 'endnotes', note = model[property]![host.noteId]!
    model[property] = { ...model[property], [host.noteId]: { ...note, text: runText(content.runs), content } }
  } else setDocumentNoteDecorationContent(model, host, content)
  model.contentControls = [...(model.contentControls ?? []).filter(control => !allIds.has(control.blockId)),
    ...(scoped.contentControls ?? []).filter(control => !control.metadata?.officeStoryLockSource)]
  const comments: Record<string, DocumentThreadComment> = { ...model.comments }
  for (const [id, original] of Object.entries(model.comments ?? {})) {
    if (!original.anchor.some(anchor => oldIds.has(anchor.blockId))) continue
    const outside = original.anchor.filter(anchor => !oldIds.has(anchor.blockId)), edited = scoped.comments?.[id]
    if (!edited && !outside.length) { delete comments[id]; continue }
    comments[id] = { ...(edited ?? original), anchor: [...outside, ...(edited?.anchor ?? [])] }
  }
  for (const [id, comment] of Object.entries(scoped.comments ?? {})) if (!Object.hasOwn(comments, id)) comments[id] = comment
  if (model.comments !== undefined || Object.keys(comments).length) model.comments = comments
}

export function cloneDocumentBlockHostModel(model: DocumentModel, host: DocumentAuxiliaryBlockHost): DocumentModel {
  return documentModelForBlockHost(cloneModel(model), host, true)
}
