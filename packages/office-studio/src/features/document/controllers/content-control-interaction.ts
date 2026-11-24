import type { DocumentModel } from '../../../domain/model'
import type { DocumentContentControlSelectionRequest } from '../commands/content-control-command-types'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { cloneDocumentInputSession, documentInputLocation, sameDocumentInputLocation, type DocumentInputSession } from '../model/document-input-session'
import { findDocumentBlock, runText, runsForTarget, sameTarget, textTargets, type DocumentSelection } from '../model/document-model'
import { cloneDocumentSelection } from '../model/document-table-selection'
import { isDocumentTextBoundary } from '../model/document-story'
import { documentModelForBlockHost } from '../services/document-block-host'

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

/** Value commands replace text within one paragraph; a temporary SDT leaves a collapsed end caret. */
export function documentContentControlSelectionAfterValueEdit(request: DocumentContentControlSelectionRequest,
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
  const target = textTargets(after).find(target => (target.cellBlockId ?? target.blockId) === (control?.blockId ?? previous.blockId))
  if (!target) return undefined
  const text = runText(runsForTarget(after, target))
  if (control) {
    if (!isDocumentTextBoundary(text, control.startOffset) || !isDocumentTextBoundary(text, control.endOffset)
      || control.startOffset > control.endOffset) return undefined
    return { anchor: { ...target, offset: control.startOffset }, focus: { ...target, offset: control.endOffset } }
  }
  if (!previous.temporary || ['controlLocked', 'sdtLocked', 'sdtContentLocked'].includes(previous.lock ?? '')) return undefined
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
    || !sameDocumentInputLocation(current.location, documentInputLocation(context.selection, undefined, context.host))
    || context.selection.tableSelection) return false
  let scoped = context.source
  if (context.host) {
    try { scoped = documentModelForBlockHost(context.source, context.host) }
    catch (error) { if (error instanceof RangeError || error instanceof TypeError) return false; throw error }
  }
  const targets = textTargets(scoped)
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
  if (selection.tableSelection || !sameTarget(selection.anchor, selection.focus)) return false
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
  if (!control || (selection.focus.cellBlockId ?? selection.focus.blockId) !== control.blockId) return false
  const start = Math.min(selection.anchor.offset, selection.focus.offset), end = Math.max(selection.anchor.offset, selection.focus.offset)
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < control.startOffset || end > control.endOffset) return false
  const session = request.inputSession, current = context.session
  if (context.requireSession && !session) return false
  if (session && (!session.valid || !current?.valid || session.modelRevision !== current.modelRevision
    || session.generation !== current.generation || !sameDocumentInputLocation(session.location, current.location))) return false
  // A root body request must never use a legacy header/footer text session.
  if (session && (session.location.kind !== 'body' || !sameDocumentBlockHost(session.location.host, request.host))) return false
  const block = scoped.blocks.length ? findTargetText(scoped, control.blockId) : undefined
  return block !== undefined && isDocumentTextBoundary(block, start) && isDocumentTextBoundary(block, end)
}

function findTargetText(model: DocumentModel, id: string): string | undefined {
  const block = findDocumentBlock(model, id)
  return block ? runText(block.runs) : undefined
}
