import { defineStore } from 'pinia'
import { shallowRef } from 'vue'
import { sameTarget, type RunPatch } from '../model/document-model.ts'
import { cloneDocumentInputLocation, cloneDocumentInputMarks, cloneDocumentInputSession, documentInputLocationCollapsed,
  mergeDocumentInputMarks, sameDocumentInputLocation, sameDocumentInputMarks, type DocumentInputLocation, type DocumentInputSession } from '../model/document-input-session.ts'

/** Typing preferences belong to a document session and never enter its content snapshot or undo stack. */
export const useDocumentInputStore = defineStore('document-input', () => {
  const units = shallowRef<Record<string, DocumentInputSession>>({})
  const advances = new Map<string, { location: DocumentInputLocation; modelRevision: number; generation: number; sourceRevision: number; sourceValid: boolean }>()
  const read = (unitId: string) => Object.hasOwn(units.value, unitId) ? units.value[unitId] : undefined
  function get(unitId: string): DocumentInputSession | undefined {
    const state = read(unitId)
    return state ? cloneDocumentInputSession(state) : undefined
  }
  function getMarks(unitId: string, modelRevision: number, location?: DocumentInputLocation): RunPatch {
    const state = read(unitId)
    return state?.valid && state.modelRevision === modelRevision && (!location || sameDocumentInputLocation(state.location, location))
      ? cloneDocumentInputMarks(state.marks) : {}
  }
  function select(unitId: string, location: DocumentInputLocation, modelRevision: number): boolean {
    const previous = read(unitId)
    advances.delete(unitId)
    units.value = { ...units.value, [unitId]: { location: cloneDocumentInputLocation(location), marks: {}, modelRevision,
      generation: (previous?.generation ?? 0) + 1, valid: true } }
    return !previous || !previous.valid || !sameDocumentInputLocation(previous.location, location) || Object.keys(previous.marks).length > 0
  }
  function format(unitId: string, location: DocumentInputLocation, patch: RunPatch, modelRevision: number): boolean {
    if (!documentInputLocationCollapsed(location)) throw new RangeError('待输入格式需要折叠的文字选区。')
    const previous = read(unitId), marks = mergeDocumentInputMarks(getMarks(unitId, modelRevision, location), patch)
    if (previous?.valid && previous.modelRevision === modelRevision && sameDocumentInputLocation(previous.location, location)
      && sameDocumentInputMarks(previous.marks, marks)) return false
    advances.delete(unitId)
    units.value = { ...units.value, [unitId]: { location: cloneDocumentInputLocation(location), marks, modelRevision,
      generation: (previous?.generation ?? 0) + 1, valid: true } }
    return true
  }
  /** Retain a stale draft only for immediate successful-input advancement; readers see no marks. */
  function invalidate(unitId: string): boolean {
    advances.delete(unitId)
    const previous = read(unitId)
    if (!previous?.valid) return false
    units.value = { ...units.value, [unitId]: { ...previous, valid: false } }
    return true
  }
  /** Only a checked, committed input may authorize recovery of the stale draft. */
  function allowAdvance(unitId: string, location: DocumentInputLocation, modelRevision: number, generation: number): boolean {
    const previous = read(unitId)
    if (!previous || previous.generation !== generation
      || (previous.valid ? modelRevision !== previous.modelRevision || documentInputLocationCollapsed(previous.location) : modelRevision !== previous.modelRevision + 1)
      || !documentInputLocationCollapsed(location) || previous.location.kind !== location.kind
      || location.kind === 'story' && (previous.location.kind !== 'story' || previous.location.range.sectionId !== location.range.sectionId
        || previous.location.range.story !== location.range.story)) return false
    advances.set(unitId, { location: cloneDocumentInputLocation(location), modelRevision, generation,
      sourceRevision: previous.modelRevision, sourceValid: previous.valid })
    return true
  }
  function advance(unitId: string, location: DocumentInputLocation, modelRevision: number): boolean {
    const previous = read(unitId), ticket = advances.get(unitId)
    if (!previous || !ticket || ticket.modelRevision !== modelRevision || ticket.generation !== previous.generation
      || ticket.sourceRevision !== previous.modelRevision || ticket.sourceValid !== previous.valid
      || modelRevision !== previous.modelRevision + (previous.valid ? 0 : 1) || !sameDocumentInputLocation(ticket.location, location)) return false
    advances.delete(unitId)
    units.value = { ...units.value, [unitId]: { ...previous, location: cloneDocumentInputLocation(location), modelRevision, valid: true,
      marks: cloneDocumentInputMarks(previous.marks), generation: previous.generation + 1 } }
    return true
  }
  /** Body input publishes its new caret and overrides in one reactive assignment. */
  function completeInput(unitId: string, location: DocumentInputLocation, marks: RunPatch, modelRevision: number, generation: number | undefined): boolean {
    const previous = read(unitId)
    if (previous?.generation !== generation || previous && (previous.valid || previous.modelRevision >= modelRevision)
      || !documentInputLocationCollapsed(location)) return false
    const preparedMarks = mergeDocumentInputMarks({}, marks)
    advances.delete(unitId)
    units.value = { ...units.value, [unitId]: { location: cloneDocumentInputLocation(location), marks: preparedMarks,
      modelRevision, generation: (generation ?? 0) + 1, valid: true } }
    return true
  }
  /** Restore a checked refresh session, mapping body offsets while retaining typing marks. */
  function completeUnchangedSelection(unitId: string, source: DocumentInputSession, modelRevision: number, location: DocumentInputLocation = source.location): boolean {
    const previous = read(unitId)
    if (!source.valid || !previous || previous.valid || modelRevision !== source.modelRevision + 1
      || previous.modelRevision !== source.modelRevision || previous.generation !== source.generation
      || location.kind !== source.location.kind
      || !sameDocumentInputLocation(previous.location, source.location) || !sameDocumentInputMarks(previous.marks, source.marks)) return false
    if (location.kind === 'story' && source.location.kind === 'story' && !sameDocumentInputLocation(location, source.location)) return false
    if (location.kind === 'body' && source.location.kind === 'body'
      && (!sameTarget(location.selection.anchor, source.location.selection.anchor) || !sameTarget(location.selection.focus, source.location.selection.focus))) return false
    advances.delete(unitId)
    units.value = { ...units.value, [unitId]: { ...previous, location: cloneDocumentInputLocation(location),
      marks: cloneDocumentInputMarks(source.marks), modelRevision, valid: true, generation: previous.generation + 1 } }
    return true
  }
  function clearMarks(unitId: string): boolean {
    advances.delete(unitId)
    const previous = read(unitId)
    if (!previous) return false
    const changed = Object.keys(previous.marks).length > 0
    units.value = { ...units.value, [unitId]: { ...previous, marks: {}, valid: false, generation: previous.generation + 1 } }
    return changed
  }
  function removeUnit(unitId: string): void {
    advances.delete(unitId)
    if (!Object.hasOwn(units.value, unitId)) return
    const next = { ...units.value }; delete next[unitId]; units.value = next
  }
  function clear(): void { advances.clear(); units.value = {} }
  return { units, get, getMarks, select, format, invalidate, allowAdvance, advance, completeInput, completeUnchangedSelection, clearMarks, removeUnit, clear }
})
