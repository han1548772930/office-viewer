import type { DocumentContentControl, DocumentEndnoteSettings, DocumentFootnoteSettings, DocumentModel, DocumentNoteDecorations, DocumentSectionSetup } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { DEFAULT_DOCUMENT_SECTION_ID, documentSectionIds } from '../services/section-header-footer'
import { appendDocumentNoteContentControls, assertDocumentNoteSerializable, prepareDocumentNoteContentWithControls, validateDocumentNoteContent } from './document-note-content'
import type { DocumentNoteDecorationControls, DocumentNoteKind } from './document-note-types'
import { documentDormantNoteDecorationHost, documentDormantNoteDecorations, documentNoteDecorationSourceIdentity,
  DOCUMENT_NOTE_DECORATION_HOST_KEYS } from '../model/document-note-decoration-host'
import { assertDocumentNoteDecorationRemovable, preserveDocumentNoteDecorationSource, removeDocumentNoteDecorationAnchors } from './document-note-decoration-editing'
import { documentModelForBlockHost } from '../services/document-block-host'
import { validateDocumentContentControlRanges } from '../services/content-controls'

export type DocumentNoteSettingsValue = DocumentFootnoteSettings | DocumentEndnoteSettings
export const DOCUMENT_NOTE_DECORATION_KEYS = DOCUMENT_NOTE_DECORATION_HOST_KEYS

export interface EffectiveDocumentNoteSettings extends DocumentNoteDecorations {
  position: 'pageBottom' | 'beneathText' | 'docEnd' | 'sectEnd'
  restart: 'continuous' | 'eachSect' | 'eachPage'
  numberFormat: string
  startNumber: number
  columnCount: number
}

function explicitSectionSetup(model: DocumentModel, id: string): DocumentSectionSetup | undefined {
  if (!documentSectionIds(model).includes(id)) throw new RangeError('注释设置所属节不存在。')
  return id === DEFAULT_DOCUMENT_SECTION_ID ? model.sectionSetup : model.blocks.find(block => block.id === id && block.type === 'sectionBreak')?.sectionBreak?.setup
}

/** Reads explicit settings; inheritance is an independent query and never changes stored data. */
export function getDocumentNoteSettings(model: DocumentModel, kind: DocumentNoteKind, sectionId?: string): DocumentNoteSettingsValue | undefined {
  const value = sectionId === undefined ? model.noteSettings?.[kind] : explicitSectionSetup(model, sectionId)?.noteProperties?.[kind]
  return value === undefined ? undefined : cloneModel(value)
}

export function effectiveDocumentNoteSettings(model: DocumentModel, kind: DocumentNoteKind, sectionId: string): EffectiveDocumentNoteSettings {
  const document = getDocumentNoteSettings(model, kind), section = getDocumentNoteSettings(model, kind, sectionId)
  const merged = { ...document, ...section }
  return { ...merged, position: merged.position ?? (kind === 'footnote' ? 'pageBottom' : 'docEnd'),
    restart: merged.restart ?? 'continuous', numberFormat: merged.numberFormat ?? (kind === 'footnote' ? 'decimal' : 'lowerRoman'),
    startNumber: merged.startNumber ?? 1,
    columnCount: kind === 'footnote' && 'columnCount' in merged ? merged.columnCount ?? 0 : 1 }
}

export function validateDocumentNoteSettings(settings: DocumentNoteSettingsValue, kind: DocumentNoteKind, sectionId?: string): void {
  assertDocumentNoteSerializable(settings)
  if (!settings || typeof settings !== 'object' || Array.isArray(settings)) throw new TypeError('注释设置必须是可保存的对象。')
  const keys = new Set<string>(['position', 'restart', 'startNumber', 'numberFormat', 'metadata', ...(kind === 'footnote' ? ['columnCount'] : []),
    ...(sectionId === undefined ? DOCUMENT_NOTE_DECORATION_KEYS : [])])
  for (const key of Object.keys(settings)) if (!keys.has(key)) throw new TypeError('此注释设置不适用于当前类型或作用域：' + key)
  if (settings.position !== undefined && !(kind === 'footnote' ? ['pageBottom', 'beneathText'] : ['docEnd', 'sectEnd']).includes(settings.position)) throw new TypeError('注释位置无效。')
  if (settings.restart !== undefined && !(kind === 'footnote' ? ['continuous', 'eachSect', 'eachPage'] : ['continuous', 'eachSect']).includes(settings.restart)) throw new TypeError('注释编号方式无效。')
  if (settings.startNumber !== undefined && (!Number.isSafeInteger(settings.startNumber) || settings.startNumber < 1)) throw new RangeError('注释起始编号需要正的安全整数。')
  if (settings.numberFormat !== undefined && (typeof settings.numberFormat !== 'string' || !settings.numberFormat || settings.numberFormat.length > 4096)) throw new TypeError('注释编号格式无效。')
  if ('columnCount' in settings && settings.columnCount !== undefined && (!Number.isInteger(settings.columnCount) || settings.columnCount < 0 || settings.columnCount > 4)) throw new RangeError('脚注列数需要在 0 到 4 之间。')
  if (settings.metadata !== undefined && (!settings.metadata || typeof settings.metadata !== 'object' || Array.isArray(settings.metadata))) throw new TypeError('注释来源元数据无效。')
  for (const key of DOCUMENT_NOTE_DECORATION_KEYS) if (settings[key] !== undefined) {
    validateDocumentNoteContent(settings[key]!, kind, `${kind}:${key}`)
    documentNoteDecorationSourceIdentity(settings[key]!, kind, key)
  }
  const noteSettings = kind === 'footnote' ? { footnote: settings as DocumentFootnoteSettings } : { endnote: settings as DocumentEndnoteSettings }
  for (const [index, content] of documentDormantNoteDecorations({ noteSettings }, kind).entries()) {
    documentDormantNoteDecorationHost({ noteSettings }, kind, index)
    validateDocumentNoteContent(content, kind, `${kind}:dormant:${index}`)
  }
}

function validateDecorationControlsPayload(settings: DocumentNoteSettingsValue | null, payload: DocumentNoteDecorationControls | undefined, sectionId?: string): void {
  if (payload === undefined) return
  assertDocumentNoteSerializable(payload)
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new TypeError('注释分隔内容控件需要按分隔类型提供。')
  for (const key of Object.keys(payload)) {
    if (!DOCUMENT_NOTE_DECORATION_KEYS.some(decoration => decoration === key)) throw new TypeError('未知的注释分隔控件类型：' + key)
  }
  for (const key of DOCUMENT_NOTE_DECORATION_KEYS) if (payload[key] !== undefined) {
    if (sectionId !== undefined || settings === null || settings[key] === undefined || !Array.isArray(payload[key])) {
      throw new TypeError('分隔内容控件需要与全文作用域的对应分隔故事一起提供。')
    }
  }
}

function storedDecorationControls(model: DocumentModel, kind: DocumentNoteKind, decoration: typeof DOCUMENT_NOTE_DECORATION_KEYS[number]): DocumentContentControl[] {
  return (documentModelForBlockHost(model, { kind: 'noteDecoration', noteKind: kind, decoration }).contentControls ?? [])
    .filter(control => !control.metadata?.officeStoryLockSource)
}

/** Replaces settings, stories and their SDTs atomically; null removes explicit values while retaining source provenance. */
export function setDocumentNoteSettings(model: DocumentModel, kind: DocumentNoteKind, settings: DocumentNoteSettingsValue | null,
  sectionId?: string, suppressEndnotes?: boolean, decorationContentControls?: DocumentNoteDecorationControls): boolean {
  if (settings !== null) validateDocumentNoteSettings(settings, kind, sectionId)
  validateDecorationControlsPayload(settings, decorationContentControls, sectionId)
  if (suppressEndnotes !== undefined && (kind !== 'endnote' || sectionId === undefined || typeof suppressEndnotes !== 'boolean')) throw new TypeError('延后尾注只能应用于明确的文档节。')
  const candidate = cloneModel(model), before = getDocumentNoteSettings(candidate, kind, sectionId)
  const setup = sectionId === undefined ? undefined : explicitSectionSetup(candidate, sectionId)
  if (sectionId !== undefined && JSON.stringify(before) === JSON.stringify(settings ?? undefined)
    && (suppressEndnotes === undefined || setup?.suppressEndnotes === suppressEndnotes)) return false
  if (sectionId === undefined) {
    const next = { ...candidate.noteSettings }, addedControls: DocumentContentControl[] = []
    const owned: DocumentNoteSettingsValue = cloneModel(settings ?? {}), removed = new Set<string>()
    // Native dormant stories are edited by their explicit host, not by replacing numbering metadata.
    const dormant = before?.metadata?.officeDormantNoteDecorations
    if (owned.metadata?.officeDormantNoteDecorations !== undefined && JSON.stringify(owned.metadata.officeDormantNoteDecorations) !== JSON.stringify(dormant)) {
      throw new RangeError('休眠注释分隔内容需要通过其独立内容区编辑。')
    }
    if (before?.metadata || owned.metadata) owned.metadata = { ...cloneModel(before?.metadata ?? {}), ...owned.metadata }
    for (const key of DOCUMENT_NOTE_DECORATION_KEYS) {
      const controls = decorationContentControls?.[key]
      const changed = JSON.stringify(owned[key]) !== JSON.stringify(before?.[key]) || controls !== undefined
        && JSON.stringify(controls) !== JSON.stringify(storedDecorationControls(candidate, kind, key))
      if (!changed) continue
      for (const id of assertDocumentNoteDecorationRemovable(candidate, { kind: 'noteDecoration', noteKind: kind, decoration: key })) removed.add(id)
      if (owned[key] !== undefined) {
        const prepared = prepareDocumentNoteContentWithControls(candidate, kind, `${kind}:${key}`, owned[key], '', true, controls)
        owned[key] = preserveDocumentNoteDecorationSource(before?.[key], prepared.content)
        addedControls.push(...prepared.contentControls)
      }
    }
    if (!Object.keys(owned).length) delete next[kind]
    else if (kind === 'footnote') next.footnote = owned as DocumentFootnoteSettings
    else next.endnote = owned as DocumentEndnoteSettings
    if (Object.keys(next).length) candidate.noteSettings = next
    else delete candidate.noteSettings
    removeDocumentNoteDecorationAnchors(candidate, removed)
    appendDocumentNoteContentControls(candidate, addedControls)
  } else {
    const next: DocumentSectionSetup = { ...setup, noteProperties: { ...setup?.noteProperties } }
    if (settings === null) delete next.noteProperties![kind]
    else if (kind === 'footnote') next.noteProperties!.footnote = cloneModel(settings as DocumentFootnoteSettings)
    else next.noteProperties!.endnote = cloneModel(settings as DocumentEndnoteSettings)
    if (!Object.keys(next.noteProperties!).length) delete next.noteProperties
    if (suppressEndnotes !== undefined) next.suppressEndnotes = suppressEndnotes
    if (sectionId === DEFAULT_DOCUMENT_SECTION_ID) candidate.sectionSetup = next
    else {
      const block = candidate.blocks.find(block => block.id === sectionId && block.type === 'sectionBreak')!
      block.sectionBreak = { ...block.sectionBreak!, setup: next }
    }
  }
  if (JSON.stringify(model) === JSON.stringify(candidate)) return false
  validateDocumentContentControlRanges(candidate)
  for (const key of Object.keys(model)) if (!Object.hasOwn(candidate, key)) Reflect.deleteProperty(model, key)
  Object.assign(model, candidate)
  return true
}

export function documentSectionSuppressesEndnotes(model: DocumentModel, sectionId: string): boolean {
  return explicitSectionSetup(model, sectionId)?.suppressEndnotes === true
}
