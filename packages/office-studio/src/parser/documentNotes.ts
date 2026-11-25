import type { DocumentBlock, DocumentDefaultStyle, DocumentEndnoteSettings, DocumentFootnoteSettings, DocumentModel,
  DocumentNote, DocumentSectionSetup, DocumentStoryContent } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { normalizeDocumentNotes } from '../features/document/notes/document-notes'
import { DOCUMENT_NOTE_DECORATION_KEYS, validateDocumentNoteSettings, type DocumentNoteSettingsValue } from '../features/document/notes/document-note-settings'
import { tableForBlock } from '../features/document/model/document-model'
import { DOCUMENT_STORY_KEYS } from '../features/document/model/document-story'
import { projectDocumentStoryBlocks } from '../features/document/model/document-story-blocks'
import { documentTableCellContentContainers } from '../features/document/model/document-table-continuations'
import { claimParsedDocumentStoryBlocks, normalizeParsedDocumentBlockGroup, type ParsedDocumentBlockNormalizer } from './documentAuxiliaryContent'
import { projectParsedDocumentFieldRanges } from './documentTextProjection'
import { array, record } from './validation'
import { documentNoteDecorationSourceIdentity } from '../features/document/model/document-note-decoration-host'
import { removeParsedDocumentNoteBodyMarkerText } from './documentNoteBodyMarkers'

type NoteKind = 'footnote' | 'endnote'
type ReferenceStyleNormalizer = (value: unknown) => DocumentDefaultStyle['text']

function decorationContent(value: unknown, kind: NoteKind, contents: Map<string, DocumentStoryContent>, type?: string): DocumentStoryContent {
  const raw = record(value), source = record(raw.decorationSource)
  if (source.version !== 1 || source.kind !== kind || !DOCUMENT_NOTE_DECORATION_KEYS.includes(source.type as typeof DOCUMENT_NOTE_DECORATION_KEYS[number])
    || type !== undefined && source.type !== type || typeof source.noteId !== 'number' || !Number.isSafeInteger(source.noteId)
    || typeof source.partName !== 'string' || !source.partName || typeof source.sourceXml !== 'string') throw new TypeError('导入的注释分隔内容来源无效。')
  const id = String(source.noteId), content = contents.get(id)
  if (!content) throw new TypeError(`注释分隔内容 ${kind}:${id} 缺少对应的块树。`)
  contents.delete(id)
  const next: DocumentStoryContent = { ...content, metadata: { ...content.metadata, decorationSource: cloneModel(source),
    ...(raw.univerNoteDecorationSource === undefined ? {} : { univerNoteDecorationSource: cloneModel(raw.univerNoteDecorationSource) }) } }
  documentNoteDecorationSourceIdentity(next, kind, source.type as typeof DOCUMENT_NOTE_DECORATION_KEYS[number])
  return projectDocumentStoryBlocks(next, 'display').content
}

function settingsProperties(value: unknown, kind: NoteKind, sectionId?: string): DocumentNoteSettingsValue {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('导入的注释设置不是对象。')
  const raw = record(value), next: DocumentNoteSettingsValue = {}
  for (const key of ['position', 'restart', 'startNumber', 'numberFormat', 'columnCount', 'metadata'] as const) {
    if (raw[key] !== undefined) Reflect.set(next, key, cloneModel(raw[key]))
  }
  validateDocumentNoteSettings(next, kind, sectionId)
  return next
}

export function parsedDocumentSectionNoteSetup(section: Record<string, unknown>, sectionId: string): Pick<DocumentSectionSetup, 'noteProperties' | 'suppressEndnotes'> {
  const next: Pick<DocumentSectionSetup, 'noteProperties' | 'suppressEndnotes'> = {}
  if (section.noteProperties !== undefined && section.noteProperties !== null) {
    if (typeof section.noteProperties !== 'object' || Array.isArray(section.noteProperties)) throw new TypeError('导入的节注释设置无效。')
    const raw = record(section.noteProperties)
    if (Object.keys(raw).some(key => key !== 'footnote' && key !== 'endnote')) throw new TypeError('导入的节注释类型无效。')
    next.noteProperties = {
      ...(raw.footnote !== undefined ? { footnote: settingsProperties(raw.footnote, 'footnote', sectionId) as DocumentFootnoteSettings } : {}),
      ...(raw.endnote !== undefined ? { endnote: settingsProperties(raw.endnote, 'endnote', sectionId) as DocumentEndnoteSettings } : {}),
    }
  }
  if (section.suppressEndnotes !== undefined && section.suppressEndnotes !== null) {
    if (typeof section.suppressEndnotes !== 'boolean') throw new TypeError('导入的节尾注延后标志无效。')
    next.suppressEndnotes = section.suppressEndnotes
  }
  return next
}

function normalizeSourceReferenceStyles(blocks: readonly DocumentBlock[], normalize: ReferenceStyleNormalizer): void {
  for (const block of blocks) {
    for (const run of block.runs) {
      const raw = record(run.metadata?.noteReference)
      const style = raw.referenceTextStyle ?? raw.textStyle
      if (style !== undefined) run.metadata = { ...run.metadata, noteReference: { ...raw, referenceTextStyle: normalize(style) } }
    }
    for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) normalizeSourceReferenceStyles(content, normalize)
  }
}

function noteSourceData(value: unknown, normalize: ReferenceStyleNormalizer): unknown {
  const normalizeValue = (value: unknown) => {
    const raw = record(value)
    const source = record(record(raw.metadata).officeNoteSource)
    const baseline = source.referenceTextStyleBaseline
    if (baseline !== undefined && (!baseline || typeof baseline !== 'object' || Array.isArray(baseline))) {
      throw new TypeError('原生注释引用样式来源基线无效。')
    }
    return { ...raw, ...(raw.referenceTextStyle !== undefined ? { referenceTextStyle: normalize(raw.referenceTextStyle) } : {}),
      ...(baseline === undefined ? {} : { metadata: { ...record(raw.metadata), officeNoteSource: {
        ...source, referenceTextStyleBaseline: normalize(baseline),
      } } }) }
  }
  if (Array.isArray(value)) return value.map(normalizeValue)
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(record(value)).map(([id, value]) => [id, normalizeValue(value)]))
  return value
}

/** Claim ordinary and special native stories before indexing references. Rich
 * note bodies are attached once, so indexing never overwrites them with text. */
export function normalizeParsedDocumentNotes(model: DocumentModel, data: Record<string, unknown>, stories: Record<string, unknown[]>,
  normalizeBlock: ParsedDocumentBlockNormalizer, claimed: Set<string>, normalizeStyle: ReferenceStyleNormalizer): void {
  const groups = new Map<NoteKind, Map<string, DocumentStoryContent>>()
  for (const kind of ['footnote', 'endnote'] as const) {
    const rawGroups = new Map<string, unknown[]>()
    for (const raw of stories[kind] ?? []) {
      const noteId = record(raw).noteId
      if (typeof noteId !== 'number' || !Number.isSafeInteger(noteId)) throw new TypeError('注释内容区缺少有效的来源标识。')
      const id = String(noteId), group = rawGroups.get(id) ?? []
      group.push(raw); rawGroups.set(id, group)
    }
    const contents = new Map<string, DocumentStoryContent>()
    for (const [id, raw] of rawGroups) {
      const parsed = normalizeParsedDocumentBlockGroup(raw, `${kind}-${id}`, normalizeBlock)
      removeParsedDocumentNoteBodyMarkerText(parsed, kind)
      projectParsedDocumentFieldRanges(parsed)
      const owned = claimParsedDocumentStoryBlocks(parsed, claimed)
      contents.set(id, { blocks: owned.blocks, runs: [], paragraphs: [], metadata: {
        officeStoryBlockSource: { version: 1, story: kind, noteId: id, sourceRefs: raw.flatMap(value => array(record(value).sourceRefs)),
          ...(owned.copied ? { copiedSharedPart: true } : {}) },
      } })
    }
    groups.set(kind, contents)
  }
  if (data.noteSettings !== undefined && data.noteSettings !== null) {
    if (typeof data.noteSettings !== 'object' || Array.isArray(data.noteSettings)) throw new TypeError('导入的文档注释设置不是对象。')
    const rawSettings = record(data.noteSettings)
    if (Object.keys(rawSettings).some(key => key !== 'footnote' && key !== 'endnote')) throw new TypeError('导入的文档注释设置类型无效。')
    model.noteSettings = {}
    for (const kind of ['footnote', 'endnote'] as const) {
      if (rawSettings[kind] === undefined) continue
      const raw = record(rawSettings[kind]), settings = settingsProperties(rawSettings[kind], kind)
      for (const key of DOCUMENT_NOTE_DECORATION_KEYS) {
        if (raw[key] === undefined) continue
        settings[key] = decorationContent(raw[key], kind, groups.get(kind)!, key)
      }
      if (raw.dormantDecorations !== undefined) {
        if (!Array.isArray(raw.dormantDecorations)) throw new TypeError('导入的休眠注释分隔内容必须为数组。')
        settings.metadata = { ...settings.metadata, officeDormantNoteDecorations: { version: 1,
          contents: raw.dormantDecorations.map(value => decorationContent(value, kind, groups.get(kind)!)) } }
      }
      validateDocumentNoteSettings(settings, kind)
      if (kind === 'footnote') model.noteSettings.footnote = settings as DocumentFootnoteSettings
      else model.noteSettings.endnote = settings as DocumentEndnoteSettings
    }
  }
  const referenceBlocks = [...model.blocks,
    ...Object.values(model.sectionHeaderFooters ?? {}).flatMap(section => Object.values(section.storyContent ?? {}).flatMap(content => content?.blocks ?? [])),
    ...Array.from(groups.values()).flatMap(contents => Array.from(contents.values()).flatMap(content => content.blocks ?? [])),
  ]
  normalizeSourceReferenceStyles(referenceBlocks, normalizeStyle)
  const normalized = normalizeDocumentNotes({ ...data, footnotes: noteSourceData(data.footnotes, normalizeStyle),
    endnotes: noteSourceData(data.endnotes, normalizeStyle), notes: noteSourceData(data.notes, normalizeStyle) }, referenceBlocks)
  Object.assign(model, normalized)
  for (const kind of ['footnote', 'endnote'] as const) {
    const notes = kind === 'footnote' ? model.footnotes : model.endnotes
    for (const [id, content] of groups.get(kind)!) {
      const note = notes?.[id]
      if (!note) throw new TypeError(`注释内容区 ${kind}:${id} 缺少对应的注释记录。`)
      projectParsedDocumentFieldRanges(content.blocks!)
      note.content = projectDocumentStoryBlocks({ ...content, metadata: { ...note.content?.metadata, ...content.metadata } }, 'display').content
      note.text = note.content.runs.map(run => run.text).join('')
    }
  }
  // Body/header references were changed while the complete host index was built.
  projectParsedDocumentFieldRanges(model.blocks)
  projectParsedDocumentNoteStories(model)
}

/** Rebuild compatibility projections after final numbering, including page
 * fields; the real paragraph/table blocks remain the authoritative content. */
export function projectParsedDocumentNoteStories(model: DocumentModel): void {
  for (const section of Object.values(model.sectionHeaderFooters ?? {})) for (const story of DOCUMENT_STORY_KEYS) {
    const content = section.storyContent?.[story]
    if (content?.blocks === undefined) continue
    projectParsedDocumentFieldRanges(content.blocks)
    const projection = projectDocumentStoryBlocks(content)
    section[story] = projection.text
    section.storyContent![story] = projection.content
    section.pageFields = [...(section.pageFields ?? []).filter(field => field.story !== story), ...projection.pageFields.map(field => ({ ...field, story }))]
  }
  const projectNote = (note: DocumentNote) => {
    if (note.content?.blocks === undefined) return
    projectParsedDocumentFieldRanges(note.content.blocks)
    const projection = projectDocumentStoryBlocks(note.content, 'display')
    note.content = projection.content; note.text = projection.text
  }
  for (const note of Object.values(model.footnotes ?? {})) projectNote(note)
  for (const note of Object.values(model.endnotes ?? {})) projectNote(note)
}
