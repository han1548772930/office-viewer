import type { DocumentBlock, DocumentEndnoteSettings, DocumentFootnoteSettings, DocumentModel, DocumentStoryContent } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { tableForBlock } from '../features/document/model/document-model'
import { documentTableCellContentContainers } from '../features/document/model/document-table-continuations'
import { normalizeDocumentNotes, renumberDocumentNotes } from '../features/document/notes/document-notes'
import { appendDocumentNoteContentControls, prepareDocumentNoteContentWithControls } from '../features/document/notes/document-note-content'
import { DOCUMENT_NOTE_DECORATION_KEYS, validateDocumentNoteSettings, type DocumentNoteSettingsValue } from '../features/document/notes/document-note-settings'
import { decodeDocumentBody } from './document-body-codec'
import { bodyRecord, type UniverBodyRecord } from './document-body-style'
import { documentBodyIdentity, type DocumentBodyDecodeOptions, type DocumentBodyDecodeResult, type UniverDocumentBodyResources } from './document-body-types'
import { documentNoteInputFromPublic, documentNoteSettingsFromPublic, type DocumentNoteContentInput, type DocumentNoteSettingsInput } from './document-note-body'
import { claimPublicDocumentBody, publicDocumentRecord, publicDocumentResources } from './document-input-common'

type NoteKind = 'footnote' | 'endnote'
type DecorationKind = typeof DOCUMENT_NOTE_DECORATION_KEYS[number]
interface SpecialNoteInput { id: string; kind: NoteKind; type: DecorationKind; value: UniverBodyRecord }

function referenceBlocks(model: DocumentModel): DocumentBlock[] {
  return [...model.blocks, ...Object.values(model.sectionHeaderFooters ?? {}).flatMap(section => Object.values(section.storyContent ?? {}).flatMap(content => content?.blocks ?? []))]
}

function prepareBody(model: DocumentModel, kind: NoteKind, id: string, decoded: DocumentBodyDecodeResult, claimed: Set<string>): DocumentStoryContent {
  const owned = claimPublicDocumentBody(decoded, claimed)
  const prepared = prepareDocumentNoteContentWithControls(model, kind, id, owned.content, '', false, owned.contentControls)
  appendDocumentNoteContentControls(model, prepared.contentControls)
  return prepared.content
}

function samePublicBody(first: unknown, second: unknown): boolean {
  if (!first || !second || typeof first !== 'object' || typeof second !== 'object') return false
  const a = cloneModel(bodyRecord(first)), b = cloneModel(bodyRecord(second))
  for (const key of ['tableSource', 'drawings', 'drawingsOrder', 'lists']) { delete a[key]; delete b[key] }
  return JSON.stringify(a) === JSON.stringify(b)
}

/** Ordinary definitions and native special definitions stay separate. Numbering
 * is applied to the real candidate only after all body/story anchors exist. */
export function applyPublicDocumentNotes(model: DocumentModel, input: UniverBodyRecord, resources: UniverDocumentBodyResources,
  options: DocumentBodyDecodeOptions, claimed: Set<string>): void {
  const normal: Record<NoteKind, Map<string, UniverBodyRecord>> = { footnote: new Map(), endnote: new Map() }, special: SpecialNoteInput[] = []
  for (const kind of ['footnote', 'endnote'] as const) {
    const property = kind === 'footnote' ? 'footnotes' : 'endnotes'
    const entries = input[property] === undefined ? {} : publicDocumentRecord(input[property], property)
    for (const [key, value] of Object.entries(entries)) {
      const raw = publicDocumentRecord(value, `${property}.${key}`), id = documentBodyIdentity(raw.noteId ?? key, '注释')
      if (id !== key) throw new TypeError(`注释资源身份不匹配：${key}`)
      publicDocumentRecord(raw.body, `${property}.${key}.body`)
      if (typeof raw.type === 'string' && DOCUMENT_NOTE_DECORATION_KEYS.includes(raw.type as DecorationKind)) {
        special.push({ id, kind, type: raw.type as DecorationKind, value: raw })
      } else if (raw.type === undefined || raw.type === 'normal' || raw.type === kind) normal[kind].set(id, raw)
      else throw new TypeError(`注释资源类型无效：${String(raw.type)}`)
    }
  }
  const notes = input.notes === undefined ? {} : publicDocumentRecord(input.notes, 'notes')
  for (const [key, value] of Object.entries(notes)) {
    const raw = publicDocumentRecord(value, `notes.${key}`), id = documentBodyIdentity(raw.noteId ?? key, '注释')
    if (id !== key || raw.type !== 'footnote' && raw.type !== 'endnote') throw new TypeError(`可编辑注释身份或类型无效：${key}`)
    publicDocumentRecord(raw.body, `notes.${key}.body`)
    normal[raw.type].set(id, { ...normal[raw.type].get(id), ...raw })
  }
  const normalizedInput: UniverBodyRecord = {}
  for (const kind of ['footnote', 'endnote'] as const) {
    const values: Record<string, UniverBodyRecord> = {}
    for (const [id, raw] of normal[kind]) {
      const properties = documentNoteInputFromPublic(model, { ...raw, ...publicDocumentResources(raw, resources) } as DocumentNoteContentInput, { kind, id }, options)
      if (!properties?.content) throw new TypeError(`注释 ${id} 缺少真实内容。`)
      const content = prepareBody(model, kind, id, { content: properties.content, contentControls: properties.contentControls ?? [], diagnostics: [] }, claimed)
      values[id] = { ...properties, noteId: id, type: kind, content }
    }
    normalizedInput[kind === 'footnote' ? 'footnotes' : 'endnotes'] = values
  }
  Object.assign(model, normalizeDocumentNotes(normalizedInput, cloneModel(referenceBlocks(model))))

  const rawSettings = input.noteSettings === undefined ? {} : publicDocumentRecord(input.noteSettings, 'noteSettings')
  if (Object.keys(rawSettings).some(key => key !== 'footnote' && key !== 'endnote')) throw new TypeError('文档注释设置包含未知类型。')
  for (const kind of ['footnote', 'endnote'] as const) {
    const raw = rawSettings[kind] === undefined ? {} : publicDocumentRecord(rawSettings[kind], `noteSettings.${kind}`)
    const adapted = documentNoteSettingsFromPublic(model, raw as DocumentNoteSettingsInput, options), settings: DocumentNoteSettingsValue = adapted.settings ?? {}
    const explicit = new Set<DecorationKind>()
    for (const type of DOCUMENT_NOTE_DECORATION_KEYS) {
      const content = settings[type]
      if (!content) continue
      explicit.add(type)
      settings[type] = prepareBody(model, kind, `${kind}:${type}`, { content, contentControls: adapted.decorationContentControls?.[type] ?? [], diagnostics: [] }, claimed)
    }
    const dormant: DocumentStoryContent[] = []
    for (const entry of special.filter(entry => entry.kind === kind)) {
      const source = { version: 1, kind, type: entry.type, noteId: entry.id }
      if (explicit.has(entry.type) && samePublicBody(entry.value.body, raw[entry.type])) {
        settings[entry.type]!.metadata = { ...settings[entry.type]!.metadata, univerNoteDecorationSource: source }
        continue
      }
      const decoded = decodeDocumentBody(entry.value.body, publicDocumentResources(entry.value, resources), { ...options, allowSections: false })
      decoded.content.metadata = { ...decoded.content.metadata, univerNoteDecorationSource: source }
      const content = prepareBody(model, kind, `${kind}:${entry.type}:${entry.id}`, decoded, claimed)
      if (settings[entry.type] === undefined) settings[entry.type] = content
      else dormant.push(content)
    }
    if (dormant.length) settings.metadata = { ...settings.metadata, officeDormantNoteDecorations: { version: 1, contents: dormant } }
    validateDocumentNoteSettings(settings, kind)
    if (Object.keys(settings).length) {
      model.noteSettings ??= {}
      // The kind-specific validator above rejects the other kind's position,
      // restart and column properties before narrowing the stored payload.
      if (kind === 'footnote') model.noteSettings.footnote = settings as DocumentFootnoteSettings
      else model.noteSettings.endnote = settings as DocumentEndnoteSettings
    }
  }
  const visit = (blocks: readonly DocumentBlock[]) => {
    for (const block of blocks) {
      for (const run of block.runs) {
        const marker = bodyRecord(run.metadata?.noteReference ?? run.metadata?.note), kind = marker.kind ?? marker.type, id = marker.noteId
        if ((kind === 'footnote' || kind === 'endnote') && typeof id === 'string' && !(kind === 'footnote' ? model.footnotes : model.endnotes)?.[id]) {
          throw new RangeError(`注释引用 ${kind}:${id} 缺少对应定义。`)
        }
      }
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const children of documentTableCellContentContainers(cell)) visit(children)
    }
  }
  visit(referenceBlocks(model))
  renumberDocumentNotes(model)
}
