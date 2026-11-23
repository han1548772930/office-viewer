import { cloneModel } from '../../../domain/serialization.ts'
import { sameTarget, selectionCollapsed, type DocumentSelection, type RunPatch } from './document-model.ts'
import type { DocumentStoryRange } from './document-story.ts'
import { validateDocumentRunPatch } from './document-run-range.ts'

export const SET_DOCUMENT_INPUT_SELECTION_OPERATION = 'doc.operation.set-input-selection'
export const CLEAR_DOCUMENT_INPUT_MARKS_OPERATION = 'doc.operation.clear-input-marks'
export const ADVANCE_DOCUMENT_INPUT_MARKS_OPERATION = 'doc.operation.advance-input-marks'

export type DocumentInputLocation = { kind: 'body'; selection: DocumentSelection } | { kind: 'story'; range: DocumentStoryRange }
export interface DocumentInputSession {
  location: DocumentInputLocation
  marks: RunPatch
  modelRevision: number
  generation: number
  valid: boolean
}

/** A story input names the source session and its resulting native caret. */
export interface DocumentStoryInputCommit {
  range: DocumentStoryRange
  caret: number
  modelRevision: number
  generation: number
}

export function documentInputLocation(selection: DocumentSelection, story?: DocumentStoryRange): DocumentInputLocation {
  return story ? { kind: 'story', range: { ...story, start: Math.min(story.start, story.end), end: Math.max(story.start, story.end) } }
    : { kind: 'body', selection: { anchor: { ...selection.anchor }, focus: { ...selection.focus } } }
}

export function cloneDocumentInputLocation(location: DocumentInputLocation): DocumentInputLocation {
  return location.kind === 'story' ? { kind: 'story', range: { ...location.range } }
    : { kind: 'body', selection: { anchor: { ...location.selection.anchor }, focus: { ...location.selection.focus } } }
}

export function documentInputLocationCollapsed(location: DocumentInputLocation): boolean {
  return location.kind === 'story' ? location.range.start === location.range.end : selectionCollapsed(location.selection)
}

export function sameDocumentInputLocation(left: DocumentInputLocation, right: DocumentInputLocation): boolean {
  if (left.kind === 'story') return right.kind === 'story' && left.range.sectionId === right.range.sectionId && left.range.story === right.range.story
    && left.range.start === right.range.start && left.range.end === right.range.end
  return right.kind === 'body' && sameTarget(left.selection.anchor, right.selection.anchor) && sameTarget(left.selection.focus, right.selection.focus)
    && left.selection.anchor.offset === right.selection.anchor.offset && left.selection.focus.offset === right.selection.focus.offset
}

/** Undefined color/highlight is a deliberate clear; JSON-cloning the patch would lose that intent. */
export function cloneDocumentInputMarks(patch: RunPatch): RunPatch {
  return { ...patch, ...(patch.metadata ? { metadata: cloneModel(patch.metadata) } : {}) }
}

export function mergeDocumentInputMarks(previous: RunPatch, patch: RunPatch): RunPatch {
  validateDocumentRunPatch(patch)
  return cloneDocumentInputMarks({ ...previous, ...patch })
}

export function sameDocumentInputMarks(left: RunPatch, right: RunPatch): boolean {
  return Object.keys(left).sort().join('\0') === Object.keys(right).sort().join('\0') && JSON.stringify(left) === JSON.stringify(right)
}

export function cloneDocumentInputSession(session: DocumentInputSession): DocumentInputSession {
  return { ...session, location: cloneDocumentInputLocation(session.location), marks: cloneDocumentInputMarks(session.marks) }
}
