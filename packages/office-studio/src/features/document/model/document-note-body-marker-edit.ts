import { cloneModel } from '../../../domain/serialization'
import { compareDocumentNoteBoundaries, remapDocumentNoteBoundaryParagraph, type DocumentNoteBoundaryParagraph } from './document-note-boundary-edit'
import { type DocumentNoteBodyMarkerCollection, validateDocumentNoteBodyMarkerCollection } from './document-note-body-marker'
import type { DocumentNoteEmptyReplacementAffinity, DocumentNoteInsertionAffinity } from './document-note-insertion-affinity'

interface BodyMarkerParagraph extends DocumentNoteBoundaryParagraph {
  absoluteStart: number
  absoluteEnd: number
  collection?: DocumentNoteBodyMarkerCollection
}

interface BodyMarkerTarget extends DocumentNoteBoundaryParagraph {
  absoluteStart: number
  absoluteEnd: number
}

interface BodyMarkerEdit {
  firstId: string
  start: number
  end: number
  startRank: number
  endRank: number
  insertedLength: number
  outputIds: ReadonlySet<string>
}

/** Prepare every point before publishing field or paragraph metadata. Inserting
 * at the symbol's exact native boundary enters body text after the symbol. */
export function planDocumentNoteBodyMarkersForRangeEdit<T extends BodyMarkerTarget>(
  source: readonly BodyMarkerParagraph[], targets: readonly T[], edit: BodyMarkerEdit,
  insertion?: DocumentNoteInsertionAffinity<{ paragraph: T; offset: number; rank: number }>,
  replacement?: DocumentNoteEmptyReplacementAffinity<{ paragraph: T; offset: number; rank: number }>,
  incoming: ReadonlyMap<string, DocumentNoteBodyMarkerCollection> = new Map(),
): Map<string, DocumentNoteBodyMarkerCollection | undefined> {
  const byId = new Map(source.map(paragraph => [paragraph.id, paragraph]))
  const affected = source.find(paragraph => paragraph.collection !== undefined
    && paragraph.absoluteEnd >= edit.start && paragraph.absoluteStart <= edit.end)?.collection
  const plans = new Map<string, DocumentNoteBodyMarkerCollection | undefined>(targets.map(paragraph => {
    const collection = byId.get(paragraph.id)?.collection ?? (edit.outputIds.has(paragraph.id) ? affected : undefined)
    return [paragraph.id, collection === undefined ? undefined : { ...cloneModel(collection), markers: [] }]
  }))
  const append = (id: string, collection: DocumentNoteBodyMarkerCollection) => {
    let plan = plans.get(id)
    if (!plan) { plan = { ...cloneModel(collection), markers: [] }; plans.set(id, plan) }
    if (plan.kind !== collection.kind) throw new TypeError('段落编辑不能混合不同类型的注释正文标记。')
    return plan
  }
  const collapsed = compareDocumentNoteBoundaries(edit.start, edit.startRank, edit.end, edit.endRank) === 0
  const delta = edit.insertedLength - (edit.end - edit.start)
  for (const paragraph of source) {
    const collection = paragraph.collection
    if (!collection) continue
    for (const marker of collection.markers) {
      const absolute = paragraph.absoluteStart + marker.offset
      const from = compareDocumentNoteBoundaries(absolute, marker.noteReferenceCount, edit.start, edit.startRank)
      const to = compareDocumentNoteBoundaries(absolute, marker.noteReferenceCount, edit.end, edit.endRank)
      if (!collapsed && from > 0 && to < 0) continue
      const mappedAbsolute = absolute + (collapsed ? from > 0 ? edit.insertedLength : 0 : to >= 0 && from > 0 ? delta : 0)
      const fallback = targets.find(target => mappedAbsolute >= target.absoluteStart && mappedAbsolute <= target.absoluteEnd)
      if (!fallback) throw new RangeError(`注释正文标记 ${marker.id} 的编辑后位置无法映射。`)
      let mapped = remapDocumentNoteBoundaryParagraph(paragraph.noteReferences, marker.offset, marker.noteReferenceCount, targets,
        { paragraph: fallback, offset: mappedAbsolute - fallback.absoluteStart })
      if (paragraph.id === edit.firstId && insertion && marker.offset === insertion.sourceOffset && marker.noteReferenceCount === insertion.sourceRank) {
        mapped = insertion.before
      } else if (paragraph.id === edit.firstId && replacement && marker.offset === replacement.sourceOffset) {
        if (marker.noteReferenceCount === replacement.startRank) mapped = replacement.before
        else if (marker.noteReferenceCount === replacement.endRank) mapped = replacement.after
      }
      const destination = append(mapped.paragraph.id, collection)
      destination.markers.push({ ...cloneModel(marker), offset: mapped.offset,
        noteReferenceCount: mapped.rank ?? marker.noteReferenceCount })
    }
  }
  for (const [id, collection] of incoming) {
    const destination = append(id, collection)
    destination.markers.push(...cloneModel(collection.markers))
    // Incoming points and surviving source points have independent provenance.
    // Stable ordering at the same native boundary keeps their source sequence.
    destination.markers.sort((a, b) => compareDocumentNoteBoundaries(a.offset, a.noteReferenceCount, b.offset, b.noteReferenceCount))
  }
  const ids = new Set<string>()
  let count = 0
  for (const paragraph of targets) {
    const collection = plans.get(paragraph.id)
    if (!collection) continue
    validateDocumentNoteBodyMarkerCollection(collection, paragraph.text, paragraph.noteReferences)
    for (const marker of collection.markers) {
      if (ids.has(marker.id)) throw new TypeError('注释正文标记在编辑后的多个段落重复。')
      ids.add(marker.id)
      if (++count > 4096) throw new RangeError('注释正文标记超过单个内容区的数量预算。')
    }
  }
  return plans
}
