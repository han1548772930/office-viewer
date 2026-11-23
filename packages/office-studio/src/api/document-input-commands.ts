import type { DocumentModel, DocumentPermissionAction } from '../domain/model.ts'
import type { DocumentEditorCommandRequest } from './document-editor-commands.ts'
import { resolveDocumentEditorSelection } from './document-editor-commands.ts'
import { documentRunFormatPatch, runText, runsForTarget, type DocumentSelection, type RunPatch } from '../features/document/model/document-model.ts'
import { DOCUMENT_STORY_KEYS, documentStoryInputEdit, isDocumentTextBoundary, type DocumentStoryRange } from '../features/document/model/document-story.ts'
import { documentInputLocation, documentInputLocationCollapsed, sameDocumentInputLocation, sanitizeDocumentInputSelection, type DocumentInputLocation,
  type DocumentInputSession, type DocumentStoryInputCommit } from '../features/document/model/document-input-session.ts'
import { selectionFontPatch } from '../features/document/model/document-font.ts'
import { validateDocumentRunPatch } from '../features/document/model/document-run-range.ts'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../features/document/permissions/document-permissions.ts'
import { assertDocumentTextSelectionEditable } from '../features/document/services/document-text-formatting.ts'
import { prepareDocumentStoryTextFormatting } from '../features/document/services/document-story-formatting.ts'
import { documentSectionIds, resolveSectionHeaderFooter } from '../features/document/services/section-header-footer.ts'
import { storyBuffer } from '../features/document/services/story-editor.ts'
import type { DocumentAuxiliaryBlockHost } from '../features/document/model/document-block-host.ts'
import { documentModelForBlockHost, validateDocumentBlockHost } from '../features/document/services/document-block-host.ts'

export interface DocumentInputSelectionRequest { unitId: string; selection: DocumentSelection; story?: DocumentStoryRange; host?: DocumentAuxiliaryBlockHost }

export function parseDocumentInputSelectionRequest(params: DocumentInputSelectionRequest): DocumentInputSelectionRequest {
  if (!params || typeof params !== 'object' || Array.isArray(params) || typeof params.unitId !== 'string' || !params.unitId.trim()
    || !params.selection || typeof params.selection !== 'object') throw new TypeError('文档输入选区参数无效。')
  if (params.host) validateDocumentBlockHost(params.host)
  if (params.host && params.story) throw new RangeError('输入选区不能同时指定块内容区和旧页眉页脚范围。')
  return params
}

export function resolveDocumentInputLocation(model: DocumentModel, selection: DocumentSelection, story: DocumentStoryRange | undefined,
  context: DocumentPermissionContext, action?: DocumentPermissionAction, host?: DocumentAuxiliaryBlockHost): { selection: DocumentSelection; location: DocumentInputLocation } {
  if (host && story) throw new RangeError('输入选区不能同时指定块内容区和旧页眉页脚范围。')
  const scoped = host ? documentModelForBlockHost(model, host, action === 'edit') : model
  const bodySelection = story ? sanitizeDocumentInputSelection(model, { anchor: selection.anchor, focus: selection.focus }) : resolveDocumentEditorSelection(scoped, selection)
  if (story) {
    if (!DOCUMENT_STORY_KEYS.includes(story.story) || !documentSectionIds(model).includes(story.sectionId)) throw new RangeError('页眉页脚片段不存在。')
    const section = resolveSectionHeaderFooter(model, story.sectionId), buffer = storyBuffer(section, story.story)
    if (!isDocumentTextBoundary(buffer.value, story.start) || !isDocumentTextBoundary(buffer.value, story.end)) throw new RangeError('页眉页脚选区的 UTF-16 位置无效。')
    if (action) assertDocumentPermission(model.permissionRules ?? [], context, action, [{ scope: 'section', id: story.sectionId },
      { scope: 'entity', id: `${story.sectionId}:${story.story}` }, { scope: 'entity', id: `${story.sectionId}:${story.story.toLowerCase().includes('header') ? 'header' : 'footer'}` }])
    return { selection: bodySelection, location: documentInputLocation(bodySelection, story) }
  }
  for (const point of [bodySelection.anchor, bodySelection.focus]) {
    if (!isDocumentTextBoundary(runText(runsForTarget(scoped, point)), point.offset)) throw new RangeError('文档选区的 UTF-16 位置无效。')
  }
  if (action) assertDocumentPermission(model.permissionRules ?? [], context, action, permissionTargetsForSelection(scoped, bodySelection))
  return { selection: bodySelection, location: documentInputLocation(bodySelection, undefined, host) }
}

/** Caret formatting validates the same permission, lock, Unicode and patch boundary before touching session state. */
export function prepareCollapsedDocumentInputFormat(model: DocumentModel, request: DocumentEditorCommandRequest, context: DocumentPermissionContext):
  { selection: DocumentSelection; location: DocumentInputLocation; patch: RunPatch } | undefined {
  if (request.command.id !== 'doc.command.format-text') return undefined
  const command = request.command, resolved = resolveDocumentInputLocation(model, request.selection, command.story, context, 'edit', command.host)
  if (!documentInputLocationCollapsed(resolved.location)) return undefined
  const patch = documentRunFormatPatch(command.patch, command.clearFields)
  validateDocumentRunPatch(patch)
  if (resolved.location.kind === 'story') {
    const plan = prepareDocumentStoryTextFormatting(model, resolved.location.range, patch, context)
    return { ...resolved, patch: plan.typingPatch }
  }
  const scoped = command.host ? documentModelForBlockHost(model, command.host, true) : model
  assertDocumentTextSelectionEditable(scoped, resolved.selection)
  return { ...resolved, patch: selectionFontPatch(scoped, resolved.selection, patch) }
}

/** Anchor the splice to the actual source selection, including repeated text and native deletion. */
export function assertDocumentStoryInputSplice(before: string, after: string, range: Pick<DocumentStoryRange, 'start' | 'end'>, caret: number): void {
  documentStoryInputEdit(before, after, range, caret)
  if (before === after && range.start === range.end) throw new RangeError('页眉页脚输入没有改变文字。')
}

export function prepareDocumentStoryInputAdvance(before: DocumentModel, after: DocumentModel, selection: DocumentSelection,
  sectionId: string, input: DocumentStoryInputCommit, session: DocumentInputSession | undefined, modelRevision: number,
  context: DocumentPermissionContext): DocumentInputLocation {
  if (!input || typeof input !== 'object' || !input.range || typeof input.range !== 'object'
    || !Number.isSafeInteger(input.caret) || !Number.isSafeInteger(input.modelRevision) || input.modelRevision < 0
    || !Number.isSafeInteger(input.generation) || input.generation < 1) throw new TypeError('页眉页脚输入会话参数无效。')
  const { location } = resolveDocumentInputLocation(before, selection, input.range, context, 'edit')
  if (!session?.valid || session.modelRevision !== modelRevision || input.modelRevision !== modelRevision
    || input.generation !== session.generation || !sameDocumentInputLocation(session.location, location)) {
    throw new Error('页眉页脚输入会话已更改，请重新输入。')
  }
  if (location.kind !== 'story' || location.range.sectionId !== sectionId) throw new RangeError('页眉页脚输入目标与命令不一致。')
  const previous = storyBuffer(resolveSectionHeaderFooter(before, sectionId), location.range.story)
  const next = storyBuffer(resolveSectionHeaderFooter(after, sectionId), location.range.story)
  assertDocumentStoryInputSplice(previous.value, next.value, location.range, input.caret)
  return { kind: 'story', range: { ...location.range, start: input.caret, end: input.caret } }
}
