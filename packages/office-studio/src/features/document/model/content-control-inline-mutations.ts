import type { DocumentContentControl, DocumentModel, DocumentNote } from '../../../domain/model'
import { runText, runsForTarget, sameTarget, tableForBlock, textTargets, type DocumentRun, type DocumentSelection } from './document-model'
import { documentNoteReferenceRunMatches, documentRunHasNoteReference, type DocumentNoteReferenceKind } from './document-note-reference'
import { isDocumentTextBoundary } from './document-story'
import { documentTableCellContentContainers } from './document-table-continuations'
import { documentNoteReferenceHosts } from '../notes/document-note-references'
import { resolveDocumentNoteLabels } from '../notes/document-note-numbering'

interface InlineNote { kind: DocumentNoteReferenceKind; note: DocumentNote }
interface InlineToken { text: string; start: number; end: number; properties: string; runIds: string[]; note?: InlineNote }

function runProperties(run: DocumentRun, note: boolean): string {
  const { id: _id, text: _text, ...properties } = run
  if (note && properties.metadata) {
    const metadata = { ...properties.metadata }
    for (const key of ['note', 'noteReference']) {
      const marker = metadata[key]
      if (!marker || typeof marker !== 'object' || Array.isArray(marker)) { delete metadata[key]; continue }
      const retained = { ...marker }
      for (const identity of ['id', 'noteId', 'footnoteId', 'endnoteId', 'type', 'kind', 'noteType']) Reflect.deleteProperty(retained, identity)
      if (Object.keys(retained).length) metadata[key] = retained
      else delete metadata[key]
    }
    delete metadata.officeStoryRunSourceId
    if (Object.keys(metadata).length) properties.metadata = metadata
    else delete properties.metadata
  }
  return JSON.stringify(properties)
}

/** An invalid note trace must not fall back to matching plain text alone. */
export function documentContentControlInlineMutationNeedsNoteTrace(model: DocumentModel, resources: DocumentModel, blockId: string): boolean {
  const target = textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === blockId)
  if (!target) return false
  const runs = runsForTarget(model, target)
  return runs.some(documentRunHasNoteReference) || (['footnote', 'endnote'] as const).some(kind =>
    Object.values(kind === 'footnote' ? resources.footnotes ?? {} : resources.endnotes ?? {}).some(note =>
      note.referenceBlockId === blockId || runs.some(run => documentNoteReferenceRunMatches(run, note, kind))))
}

function inlineTokens(model: DocumentModel, resources: DocumentModel, blockId: string, after: boolean): InlineToken[] | undefined {
  const target = textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === blockId)
  if (!target) return undefined
  const notes = (['footnote', 'endnote'] as const).flatMap(kind => Object.values(kind === 'footnote' ? resources.footnotes ?? {} : resources.endnotes ?? {})
    .map(note => ({ kind, note })))
  const result: InlineToken[] = []
  let labels: ReturnType<typeof resolveDocumentNoteLabels> | undefined
  if (after) {
    try { labels = resolveDocumentNoteLabels(resources) }
    catch { return undefined }
  }
  let offset = 0
  for (const run of runsForTarget(model, target)) {
    const owners = notes.filter(entry => documentNoteReferenceRunMatches(run, entry.note, entry.kind))
    if (owners.length > 1 || !owners.length && documentRunHasNoteReference(run)) return undefined
    const note = owners[0], end = offset + run.text.length
    if (note && after) {
      const expected = labels?.get(`${note.kind}:${note.note.id}`)
      if (!expected || run.text !== expected.label || expected.reference.block.id !== blockId || expected.reference.run.id !== run.id
        || note.note.referenceBlockId !== blockId || note.note.referenceRunId !== run.id && run.metadata?.officeStoryRunSourceId !== note.note.referenceRunId) return undefined
    }
    result.push({ text: run.text, start: offset, end, properties: runProperties(run, !!note), runIds: [run.id], ...(note ? { note } : {}) })
    offset = end
  }
  return result
}

function tokenSlice(tokens: readonly InlineToken[], start: number, end: number, empties: 'inside' | 'prefix' | 'suffix'): InlineToken[] | undefined {
  const result: InlineToken[] = []
  for (const token of tokens) {
    if (token.start === token.end) {
      const included = empties === 'prefix' ? token.start <= end : empties === 'suffix' ? token.start >= start : token.start > start && token.start < end
      if (included) result.push(token)
      continue
    }
    const from = Math.max(start, token.start), to = Math.min(end, token.end)
    if (from >= to) continue
    if (token.note && (from !== token.start || to !== token.end)) return undefined
    result.push({ ...token, text: token.text.slice(from - token.start, to - token.start), start: from, end: to })
  }
  return result
}

function normalized(tokens: readonly InlineToken[]): InlineToken[] {
  const result: InlineToken[] = []
  for (const token of tokens) {
    const previous = result.at(-1)
    if (previous && !previous.note && !token.note && previous.text && token.text && previous.properties === token.properties) {
      previous.text += token.text; previous.end = token.end; previous.runIds.push(...token.runIds)
    } else result.push({ ...token, runIds: [...token.runIds] })
  }
  return result
}

function sameTokens(expected: readonly InlineToken[], actual: readonly InlineToken[], copied: boolean, before: DocumentModel): boolean {
  const left = normalized(expected), right = normalized(actual)
  if (left.length !== right.length) return false
  const noteIds = new Set<string>()
  for (const [index, source] of left.entries()) {
    const current = right[index]!
    if (source.properties !== current.properties || !!source.note !== !!current.note) return false
    if (!source.note) { if (source.text !== current.text) return false; continue }
    if (source.note.kind !== current.note!.kind) return false
    const key = `${current.note!.kind}:${current.note!.note.id}`
    if (copied) {
      const notes = current.note!.kind === 'footnote' ? before.footnotes : before.endnotes
      if (notes?.[current.note!.note.id] || noteIds.has(key) || current.note!.note.id === source.note.note.id) return false
      noteIds.add(key)
    } else if (source.note.note.id !== current.note!.note.id || source.runIds[0] !== current.runIds[0]) return false
  }
  return true
}

/** Map a verified prefix through variable-width derived note labels, including a boundary inside a plain run. */
function prefixOffset(prefix: readonly InlineToken[], actual: readonly InlineToken[]): number | undefined {
  const left = normalized(prefix), right = normalized(actual)
  let index = 0, used = 0, offset = 0
  for (const source of left) {
    const current = right[index]
    if (!current || source.properties !== current.properties || !!source.note !== !!current.note) return undefined
    if (source.note) {
      if (used || source.note.kind !== current.note!.kind || source.note.note.id !== current.note!.note.id || source.runIds[0] !== current.runIds[0]) return undefined
      offset = current.end; index++; continue
    }
    if (!source.text && current.text) return undefined
    if (current.text.slice(used, used + source.text.length) !== source.text) return undefined
    used += source.text.length; offset = current.start + used
    if (used === current.text.length) { index++; used = 0 }
  }
  return offset
}

/** A note-aware inline trace checks surrounding content, old identities and the fresh copy before choosing live offsets. */
export function documentContentControlInlineMutationSelection(before: DocumentModel, after: DocumentModel, id: string, action: 'duplicate' | 'delete',
  beforeResources: DocumentModel = before, afterResources: DocumentModel = after): { controlId?: string; selection: DocumentSelection } | undefined {
  const item = before.contentControls?.find(control => control.id === id)
  if (item?.type !== 'repeatingSectionItem' || item.blockRange || item.tableRowRange) return undefined
  const target = textTargets(before).find(target => (target.cellBlockId ?? target.blockId) === item.blockId)
  const currentTarget = target && textTargets(after).find(current => sameTarget(current, target))
  const source = inlineTokens(before, beforeResources, item.blockId, false), current = inlineTokens(after, afterResources, item.blockId, true)
  if (!target || !currentTarget || !source || !current || !source.some(token => token.note)) return undefined
  const oldText = runText(runsForTarget(before, target)), text = runText(runsForTarget(after, currentTarget))
  if (!isDocumentTextBoundary(oldText, item.startOffset) || !isDocumentTextBoundary(oldText, item.endOffset) || item.startOffset > item.endOffset) return undefined
  const point = (offset: number) => ({ ...currentTarget, offset })
  if (action === 'delete') {
    if (after.contentControls?.some(control => control.id === id)) return undefined
    const prefix = tokenSlice(source, 0, item.startOffset, 'prefix'), suffix = tokenSlice(source, item.endOffset, oldText.length, 'suffix')
    if (!prefix || !suffix) return undefined
    if (item.startOffset === item.endOffset) while (suffix[0]?.start === item.endOffset && !suffix[0].text) suffix.shift()
    if (!sameTokens([...prefix, ...suffix], current, false, beforeResources)) return undefined
    const offset = prefixOffset(prefix, current)
    return offset !== undefined && isDocumentTextBoundary(text, offset) ? { selection: { anchor: point(offset), focus: point(offset) } } : undefined
  }
  const original = after.contentControls?.find(control => control.id === id)
  if (!original || original.type !== item.type || original.parentId !== item.parentId || original.blockId !== item.blockId || original.blockRange || original.tableRowRange) return undefined
  const oldIds = new Set(before.contentControls?.map(control => control.id))
  const copies = (after.contentControls ?? []).filter(control => !oldIds.has(control.id) && control.type === item.type && control.parentId === item.parentId
    && control.blockId === item.blockId && !control.blockRange && !control.tableRowRange && control.startOffset === original.endOffset)
  if (copies.length !== 1) return undefined
  const copy: DocumentContentControl = copies[0]!
  if (original.startOffset > original.endOffset || copy.startOffset > copy.endOffset
    || ![original.startOffset, original.endOffset, copy.startOffset, copy.endOffset].every(offset => isDocumentTextBoundary(text, offset))) return undefined
  const oldPrefix = tokenSlice(source, 0, item.endOffset, 'prefix'), prefix = tokenSlice(current, 0, copy.startOffset, 'prefix')
  const oldContent = tokenSlice(source, item.startOffset, item.endOffset, 'inside'), content = tokenSlice(current, copy.startOffset, copy.endOffset, 'inside')
  const oldSuffix = tokenSlice(source, item.endOffset, oldText.length, 'suffix'), suffix = tokenSlice(current, copy.endOffset, text.length, 'suffix')
  const oldStart = tokenSlice(source, 0, item.startOffset, 'prefix'), currentStart = tokenSlice(current, 0, original.startOffset, 'prefix')
  if (!oldPrefix || !prefix || !oldContent || !content || !oldSuffix || !suffix || !oldStart || !currentStart) return undefined
  // Insertion assigns a zero-width boundary entity to its prefix once.
  while (oldSuffix[0]?.start === item.endOffset && !oldSuffix[0]?.text) oldSuffix.shift()
  while (suffix[0]?.start === copy.endOffset && !suffix[0]?.text) suffix.shift()
  const oldRunIds = new Set<string>()
  const collect = (blocks: DocumentModel['blocks']) => {
    for (const block of blocks) {
      for (const run of block.runs) oldRunIds.add(run.id)
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const container of documentTableCellContentContainers(cell)) collect(container)
    }
  }
  for (const owner of documentNoteReferenceHosts(beforeResources)) collect(owner.blocks)
  if (content.some(token => token.runIds.some(runId => oldRunIds.has(runId))) || !sameTokens(oldPrefix, prefix, false, beforeResources)
    || !sameTokens(oldContent, content, true, beforeResources) || !sameTokens(oldSuffix, suffix, false, beforeResources)
    || !sameTokens(oldStart, currentStart, false, beforeResources)) return undefined
  return { controlId: copy.id, selection: { anchor: point(copy.startOffset), focus: point(copy.endOffset) } }
}
