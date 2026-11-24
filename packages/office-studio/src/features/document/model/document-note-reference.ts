import type { DocumentNote } from '../../../domain/model'
import type { DocumentRun } from './document-model'

export type DocumentNoteReferenceKind = 'footnote' | 'endnote'

export function documentRunHasNoteReference(run: Pick<DocumentRun, 'metadata'>): boolean {
  return run.metadata?.note !== undefined && run.metadata.note !== null
    || run.metadata?.noteReference !== undefined && run.metadata.noteReference !== null
}

/** A migrated run's recorded source ID is an alias, while explicit note identity stays authoritative. */
export function documentNoteReferenceRunMatches(run: Pick<DocumentRun, 'id' | 'metadata'>, note: DocumentNote,
  kind?: DocumentNoteReferenceKind): boolean {
  const identity = !!note.referenceRunId && (run.id === note.referenceRunId || run.metadata?.officeStoryRunSourceId === note.referenceRunId)
  let explicit = false, explicitKind: DocumentNoteReferenceKind | undefined
  for (const marker of [run.metadata?.note, run.metadata?.noteReference]) {
    if (marker === undefined || marker === null) continue
    if (typeof marker === 'string' || typeof marker === 'number') {
      if (typeof marker === 'number' && !Number.isSafeInteger(marker) || String(marker) !== note.id) return false
      explicit = true
      continue
    }
    if (typeof marker !== 'object' || Array.isArray(marker)) return false
    for (const key of ['noteId', 'id', 'footnoteId', 'endnoteId']) {
      const referenceId = Reflect.get(marker, key)
      if (referenceId === undefined || referenceId === null) continue
      if (typeof referenceId !== 'string' && typeof referenceId !== 'number'
        || typeof referenceId === 'number' && !Number.isSafeInteger(referenceId) || String(referenceId) !== note.id) return false
      explicit = true
    }
    const kinds = [Reflect.get(marker, 'type'), Reflect.get(marker, 'kind'), Reflect.get(marker, 'noteType'),
      Reflect.get(marker, 'footnoteId') != null ? 'footnote' : undefined, Reflect.get(marker, 'endnoteId') != null ? 'endnote' : undefined]
    for (const referenceKind of kinds) {
      if (referenceKind === undefined || referenceKind === null) continue
      if (referenceKind !== 'footnote' && referenceKind !== 'endnote'
        || kind !== undefined && referenceKind !== kind || explicitKind !== undefined && referenceKind !== explicitKind) return false
      explicitKind = referenceKind
    }
  }
  // A kindless numeric ID can be reused in footnotes and endnotes. An existing
  // reference run disambiguates it; a typed marker can repair a migrated ID.
  return identity || explicit && (kind === undefined || explicitKind !== undefined || !note.referenceRunId)
}
