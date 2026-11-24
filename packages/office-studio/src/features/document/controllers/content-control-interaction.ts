import type { DocumentContentControl, DocumentModel } from '../../../domain/model'
import type { DocumentContentControlRepeatingItemCommand, DocumentContentControlSelectionRequest, DocumentContentControlValueCommand } from '../commands/content-control-command-types'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { cloneDocumentInputSession, documentInputLocation, sameDocumentInputLocation, type DocumentInputSession } from '../model/document-input-session'
import { runText, runsForTarget, sameTarget, textTargets, type DocumentSelection } from '../model/document-model'
import { cloneDocumentSelection, createDocumentTableViewSelection, resolveDocumentTableViewSelection } from '../model/document-table-selection'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentModelForBlockHost } from '../services/document-block-host'
import { documentContentControlTemporaryBindingCaret } from '../services/content-control-bindings'
import { documentContentControlContains, documentContentControlContainsSelection, documentContentControlTableRowSelection,
  documentContentControlTableRowsMutationSelection, resolveDocumentContentControlTableRows } from '../model/content-control-table-rows'
import { documentContentControlBlocksSelection, documentContentControlOwnedBlockIds, resolveDocumentContentControlBlocks } from '../model/content-control-block-range'
import { documentContentControlBlocksMutationSelection, documentContentControlBlocksValueMutationRange } from '../model/content-control-block-mutations'
import { documentContentControlInlineMutationNeedsNoteTrace, documentContentControlInlineMutationSelection } from '../model/content-control-inline-mutations'

export interface DocumentContentControlInteractionContext {
  source: DocumentModel
  host?: DocumentAuxiliaryBlockHost
  session?: DocumentInputSession
  requireSession?: boolean
  disabled?: boolean
}

export interface DocumentContentControlFocusToken {
  request: DocumentContentControlSelectionRequest
  session: DocumentInputSession
}

export interface DocumentContentControlFocusContext extends DocumentContentControlInteractionContext {
  selection: DocumentSelection
}

/** Resolve a control's logical text range in the requested host, including its real table coordinates. */
export function documentContentControlSelectionForId(model: DocumentModel, controlId: string,
  host?: DocumentAuxiliaryBlockHost): DocumentSelection | undefined {
  const scoped = contentControlScopedModel(model, host)
  return scoped && contentControlSelectionInModel(scoped, controlId)
}

/** Explicit parents are authoritative; older ranges without parent IDs can resolve strict containment. */
export function documentContentControlParentSelection(model: DocumentModel, controlId: string,
  host?: DocumentAuxiliaryBlockHost): { controlId: string; selection: DocumentSelection } | undefined {
  const scoped = contentControlScopedModel(model, host)
  const controls = scoped?.contentControls?.filter(control => !control.metadata?.officeStoryLockSource)
  const child = controls?.find(control => control.id === controlId)
  if (!scoped || !child || !contentControlSelectionInModel(scoped, child.id)) return undefined
  const contains = (parent: DocumentContentControl) => documentContentControlContains(scoped, parent, child)
  const strict = (parent: DocumentContentControl) => parent.blockRange
    ? !child.blockRange || JSON.stringify(parent.blockRange) !== JSON.stringify(child.blockRange)
    : parent.tableRowRange
    ? !child.tableRowRange || JSON.stringify(parent.tableRowRange) !== JSON.stringify(child.tableRowRange)
    : parent.startOffset < child.startOffset || parent.endOffset > child.endOffset
  const span = (control: DocumentContentControl) => control.blockRange
    ? documentContentControlOwnedBlockIds(resolveDocumentContentControlBlocks(scoped, control.blockRange)).size * 1_000_000
    : control.tableRowRange
    ? resolveDocumentContentControlTableRows(scoped, control.tableRowRange).rowIds.length * 1_000_000
    : control.endOffset - control.startOffset
  try {
    const parent = child.parentId ? controls?.find(control => control.id === child.parentId && contains(control))
      : controls?.filter(control => contains(control) && strict(control)).sort((left, right) => span(left) - span(right))[0]
    const selection = parent && contentControlSelectionInModel(scoped, parent.id)
    return parent && selection ? { controlId: parent.id, selection } : undefined
  } catch (error) {
    if (error instanceof TypeError || error instanceof RangeError) return undefined
    throw error
  }
}

function contentControlScopedModel(model: DocumentModel, host?: DocumentAuxiliaryBlockHost): DocumentModel | undefined {
  if (!host) return model
  try { return documentModelForBlockHost(model, host) }
  catch (error) { if (error instanceof RangeError || error instanceof TypeError) return undefined; throw error }
}

function contentControlSelectionInModel(model: DocumentModel, controlId: string): DocumentSelection | undefined {
  const control = model.contentControls?.find(control => control.id === controlId)
  const target = control && textTargets(model).find(target => (target.cellBlockId ?? target.blockId) === control.blockId)
  if (!control || control.metadata?.officeStoryLockSource || !target) return undefined
  const text = runText(runsForTarget(model, target))
  if (control.startOffset > control.endOffset || !isDocumentTextBoundary(text, control.startOffset)
    || !isDocumentTextBoundary(text, control.endOffset)) return undefined
  if (control.tableRowRange) {
    try { return documentContentControlTableRowSelection(model, control.tableRowRange) }
    catch (error) { if (error instanceof TypeError || error instanceof RangeError) return undefined; throw error }
  }
  if (control.blockRange) {
    try { return documentContentControlBlocksSelection(model, control.blockRange) }
    catch (error) { if (error instanceof TypeError || error instanceof RangeError) return undefined; throw error }
  }
  return { anchor: { ...target, offset: control.startOffset }, focus: { ...target, offset: control.endOffset } }
}

/** Resolve the exact new item ID or deletion caret from the committed paragraph, never an old block-only point. */
export function documentContentControlSelectionAfterRepeatingItemEdit(
  request: DocumentContentControlSelectionRequest & { command: DocumentContentControlRepeatingItemCommand },
  current: DocumentModel,
): { controlId?: string; selection: DocumentSelection } | undefined {
  if (request.source.id !== current.id || request.command.params.controlId !== request.controlId) return undefined
  const before = contentControlScopedModel(request.source, request.host), after = contentControlScopedModel(current, request.host)
  const previous = before?.contentControls?.find(control => control.id === request.controlId)
  if (!before || !after || previous?.type !== 'repeatingSectionItem') return undefined
  if (previous.placement === 'row') return documentContentControlTableRowsMutationSelection(before, after, previous.id,
    request.command.id === 'doc.command.duplicate-repeating-section-item' ? 'duplicate' : 'delete')
  if (previous.blockRange) return documentContentControlBlocksMutationSelection(before, after, previous.id,
    request.command.id === 'doc.command.duplicate-repeating-section-item' ? 'duplicate' : 'delete')
  if (documentContentControlInlineMutationNeedsNoteTrace(before, request.source, previous.blockId)
    || documentContentControlInlineMutationNeedsNoteTrace(after, current, previous.blockId)) {
    return documentContentControlInlineMutationSelection(before, after, previous.id,
      request.command.id === 'doc.command.duplicate-repeating-section-item' ? 'duplicate' : 'delete', request.source, current)
  }
  const previousRange = contentControlSelectionInModel(before, previous.id)
  const target = previousRange && textTargets(after).find(target => sameTarget(target, previousRange.anchor))
  if (!previousRange || !target) return undefined
  const oldText = runText(runsForTarget(before, previousRange.anchor)), newText = runText(runsForTarget(after, target))
  if (request.command.id === 'doc.command.delete-repeating-section-item') {
    if (after.contentControls?.some(control => control.id === previous.id)
      || newText !== oldText.slice(0, previous.startOffset) + oldText.slice(previous.endOffset)
      || !isDocumentTextBoundary(newText, previous.startOffset)) return undefined
    return { selection: { anchor: { ...target, offset: previous.startOffset }, focus: { ...target, offset: previous.startOffset } } }
  }
  const original = after.contentControls?.find(control => control.id === previous.id)
  if (!original || original.type !== previous.type || original.parentId !== previous.parentId || original.blockId !== previous.blockId
    || original.startOffset !== previous.startOffset || original.endOffset !== previous.endOffset
    || newText !== oldText.slice(0, previous.endOffset) + oldText.slice(previous.startOffset, previous.endOffset) + oldText.slice(previous.endOffset)) return undefined
  const oldIds = new Set(before.contentControls?.map(control => control.id))
  const copies = after.contentControls?.filter(control => !oldIds.has(control.id) && control.type === 'repeatingSectionItem'
    && control.parentId === previous.parentId && control.blockId === previous.blockId && control.startOffset === previous.endOffset
    && control.endOffset === previous.endOffset + previous.endOffset - previous.startOffset) ?? []
  if (copies.length !== 1) return undefined
  const selection = contentControlSelectionInModel(after, copies[0]!.id)
  return selection ? { controlId: copies[0]!.id, selection } : undefined
}

/** Value commands replace text within one paragraph; a temporary SDT leaves a collapsed end caret. */
export function documentContentControlSelectionAfterValueEdit(request: DocumentContentControlSelectionRequest & { command?: DocumentContentControlValueCommand },
  current: DocumentModel): DocumentSelection | undefined {
  if (request.source.id !== current.id) return undefined
  let before = request.source, after = current
  if (request.host) {
    try { before = documentModelForBlockHost(before, request.host); after = documentModelForBlockHost(after, request.host) }
    catch (error) { if (error instanceof RangeError || error instanceof TypeError) return undefined; throw error }
  }
  const previous = before.contentControls?.find(control => control.id === request.controlId)
  if (!previous) return undefined
  const control = after.contentControls?.find(control => control.id === request.controlId)
  if (control) return contentControlSelectionInModel(after, control.id)
  if (previous.blockRange) {
    const replacement = documentContentControlBlocksValueMutationRange(before, after, previous.id)
    const point = replacement?.selection.focus
    return point ? { anchor: { ...point }, focus: { ...point } } : undefined
  }
  const target = textTargets(after).find(target => (target.cellBlockId ?? target.blockId) === previous.blockId)
  if (!target) return undefined
  const text = runText(runsForTarget(after, target))
  if (!previous.temporary || ['controlLocked', 'sdtLocked', 'sdtContentLocked'].includes(previous.lock ?? '')) return undefined
  if (request.command) {
    const command = request.command
    const value = command.id === 'doc.command.reset-content-control-placeholder' ? ''
      : command.id === 'doc.command.set-content-control-checked' ? command.params.checked ? 'true' : 'false' : command.params.value
    const mapped = documentContentControlTemporaryBindingCaret(request.source, current, request.controlId, value, request.host,
      command.id === 'doc.command.reset-content-control-placeholder' ? previous.placeholder : undefined)
    if (mapped) return mapped
  }
  const sourceTarget = textTargets(before).find(target => (target.cellBlockId ?? target.blockId) === previous.blockId)
  if (!sourceTarget) return undefined
  const sourceText = runText(runsForTarget(before, sourceTarget))
  if (!isDocumentTextBoundary(sourceText, previous.startOffset) || !isDocumentTextBoundary(sourceText, previous.endOffset)
    || previous.startOffset > previous.endOffset) return undefined
  const offset = previous.endOffset + text.length - sourceText.length
  if (!isDocumentTextBoundary(text, offset) || text.slice(0, previous.startOffset) !== sourceText.slice(0, previous.startOffset)
    || text.slice(offset) !== sourceText.slice(previous.endOffset)) return undefined
  return { anchor: { ...target, offset }, focus: { ...target, offset } }
}

/** Capture the checked result of synchronous command dispatch, including no-op/cancelled commands. */
export function captureDocumentContentControlFocusToken(request: DocumentContentControlSelectionRequest,
  current: DocumentInputSession | undefined): DocumentContentControlFocusToken | undefined {
  const source = request.inputSession
  if (!source?.valid || !current?.valid || source.location.kind !== 'body' || current.location.kind !== 'body'
    || !sameDocumentBlockHost(source.location.host, request.host) || !sameDocumentBlockHost(current.location.host, request.host)
    || !Number.isSafeInteger(source.modelRevision) || !Number.isSafeInteger(current.modelRevision)
    || !Number.isSafeInteger(source.generation) || !Number.isSafeInteger(current.generation)
    || source.modelRevision < 0 || source.generation < 0
    || current.generation < source.generation
    || current.modelRevision !== source.modelRevision && current.modelRevision !== source.modelRevision + 1) return undefined
  return { request: { ...request, selection: cloneDocumentSelection(request.selection),
    ...(request.host ? { host: { ...request.host } } : {}), inputSession: cloneDocumentInputSession(source) },
    session: cloneDocumentInputSession(current) }
}

export function documentContentControlFocusTokenMatchesRequest(token: DocumentContentControlFocusToken,
  request: DocumentContentControlSelectionRequest): boolean {
  const source = token.request, captured = source.inputSession, current = request.inputSession
  return source.source === request.source && source.controlId === request.controlId
    && sameDocumentBlockHost(source.host, request.host)
    && sameDocumentInputLocation(documentInputLocation(source.selection, undefined, source.host), documentInputLocation(request.selection, undefined, request.host))
    && !!captured && !!current && captured.modelRevision === current.modelRevision && captured.generation === current.generation
    && sameDocumentInputLocation(captured.location, current.location)
}

/** Focus the current logical selection; the original SDT may have removed its temporary shell. */
export function documentContentControlFocusCurrent(token: DocumentContentControlFocusToken,
  context: DocumentContentControlFocusContext): boolean {
  const current = context.session, captured = token.session
  if (context.disabled || token.request.source.id !== context.source.id || !sameDocumentBlockHost(token.request.host, context.host)
    || !captured.valid || !current?.valid || current.location.kind !== 'body'
    || captured.modelRevision !== current.modelRevision || captured.generation !== current.generation
    || !sameDocumentInputLocation(captured.location, current.location)
    || !sameDocumentInputLocation(current.location, documentInputLocation(context.selection, undefined, context.host))) return false
  let scoped = context.source
  if (context.host) {
    try { scoped = documentModelForBlockHost(context.source, context.host) }
    catch (error) { if (error instanceof RangeError || error instanceof TypeError) return false; throw error }
  }
  const targets = textTargets(scoped)
  if (context.selection.tableSelection) {
    try {
      const view = resolveDocumentTableViewSelection(scoped, context.selection.tableSelection)
      const canonical = createDocumentTableViewSelection(scoped, view.tableId, view.range, view.kind, view.pageAffinity, context.selection)
      if (!sameTarget(canonical.anchor, context.selection.anchor) || canonical.anchor.offset !== context.selection.anchor.offset
        || !sameTarget(canonical.focus, context.selection.focus) || canonical.focus.offset !== context.selection.focus.offset) return false
    } catch (error) { if (error instanceof TypeError || error instanceof RangeError) return false; throw error }
  }
  return [context.selection.anchor, context.selection.focus].every(point => {
    const target = targets.find(target => sameTarget(target, point))
    return !!target && isDocumentTextBoundary(runText(runsForTarget(scoped, target)), point.offset)
  })
}

/** A delayed popover must still address the same root model, host and input session. */
export function documentContentControlInteractionCurrent(request: DocumentContentControlSelectionRequest,
  context: DocumentContentControlInteractionContext): boolean {
  if (context.disabled || request.source !== context.source || !sameDocumentBlockHost(request.host, context.host)) return false
  const { selection } = request
  let scoped = context.source
  if (context.host) {
    try { scoped = documentModelForBlockHost(context.source, context.host) }
    catch (error) {
      // A delayed overlay can outlive a deleted section, note or invalidated tree.
      if (error instanceof RangeError || error instanceof TypeError) return false
      throw error
    }
  }
  const control = scoped.contentControls?.find(control => control.id === request.controlId)
  if (!control || !documentContentControlContainsSelection(scoped, control, selection)) return false
  const session = request.inputSession, current = context.session
  if (context.requireSession && !session) return false
  if (session && (!session.valid || !current?.valid || session.modelRevision !== current.modelRevision
    || session.generation !== current.generation || !sameDocumentInputLocation(session.location, current.location))) return false
  // A root body request must never use a legacy header/footer text session.
  if (session && (session.location.kind !== 'body' || !sameDocumentBlockHost(session.location.host, request.host))) return false
  return true
}
