import type { DocumentModel } from '../../../domain/model'
import type { DocumentNoteDecorationHost } from '../model/document-block-host'
import { documentDormantNoteDecorationHost, documentDormantNoteDecorations } from '../model/document-note-decoration-host'
import { DOCUMENT_NOTE_DECORATION_KEYS } from '../notes/document-note-settings'
import { assertDocumentNoteDecorationRemovable } from '../notes/document-note-decoration-editing'
import type { DocumentNoteKind } from '../notes/document-note-types'
import { documentNoteDecorationPermissionTargets } from '../permissions/document-note-permissions'
import { resolveDocumentPermission, type DocumentPermissionContext } from '../permissions/document-permissions'

export const DOCUMENT_NOTE_DECORATION_NAMES = { separator: '分隔符', continuationSeparator: '续页分隔符', continuationNotice: '续页提示' } as const
export interface DocumentNoteDecorationEntry {
  key: string; host: DocumentNoteDecorationHost; label: string; custom: boolean; dormant: boolean; nativeId?: string | number
}
export interface DocumentNoteDecorationUiContext { permission: DocumentPermissionContext; disabled?: boolean; pendingSettings?: boolean; pendingMark?: boolean }

/** Active/default stories and every imported dormant story remain distinct addresses. */
export function documentNoteDecorationEntries(model: DocumentModel, kind: DocumentNoteKind): DocumentNoteDecorationEntry[] {
  const active: DocumentNoteDecorationEntry[] = DOCUMENT_NOTE_DECORATION_KEYS.map(decoration => ({ key: decoration,
    host: { kind: 'noteDecoration', noteKind: kind, decoration }, label: DOCUMENT_NOTE_DECORATION_NAMES[decoration],
    custom: model.noteSettings?.[kind]?.[decoration] !== undefined, dormant: false }))
  return [...active, ...documentDormantNoteDecorations(model, kind).map((content, index) => {
    const host = documentDormantNoteDecorationHost(model, kind, index), raw = content.metadata?.decorationSource
    const nativeId = raw && typeof raw === 'object' ? Reflect.get(raw, 'noteId') : undefined
    return { key: 'dormant:' + index, host, label: '休眠' + DOCUMENT_NOTE_DECORATION_NAMES[host.decoration], custom: true, dormant: true,
      ...(typeof nativeId === 'string' || typeof nativeId === 'number' ? { nativeId } : {}) }
  })]
}

export function documentNoteDecorationInspectState(model: DocumentModel, host: DocumentNoteDecorationHost,
  context: DocumentNoteDecorationUiContext): { available: boolean; reason?: string } {
  if (context.disabled) return { available: false, reason: '文档或输入状态已变化，请重新打开设置。' }
  if (context.pendingSettings) return { available: false, reason: '请先应用当前设置，再进入内容检查。' }
  if (context.pendingMark) return { available: false, reason: '自定义标记草稿用于插入注释，请先插入或清空草稿再进入内容检查。' }
  const decision = resolveDocumentPermission(model.permissionRules ?? [], context.permission, 'view', documentNoteDecorationPermissionTargets(model, host))
  return decision.allowed ? { available: true } : { available: false, reason: '此分隔内容区没有查看权限。' }
}

export function documentNoteDecorationResetState(model: DocumentModel, host: DocumentNoteDecorationHost,
  context: DocumentNoteDecorationUiContext): { available: boolean; reason?: string } {
  if (context.disabled) return { available: false, reason: '文档或输入状态已变化，请重新打开设置。' }
  const decision = resolveDocumentPermission(model.permissionRules ?? [], context.permission, 'edit', documentNoteDecorationPermissionTargets(model, host))
  if (!decision.allowed) return { available: false, reason: '此分隔内容区没有编辑权限。' }
  try { assertDocumentNoteDecorationRemovable(model, host); return { available: true } }
  catch (error) { if (error instanceof Error) return { available: false, reason: error.message }; throw error }
}
