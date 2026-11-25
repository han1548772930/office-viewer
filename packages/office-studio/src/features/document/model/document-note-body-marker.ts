import type { DocumentBlock } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentId, runText } from './document-model'
import { assertDocumentNoteReferenceBoundaryRank, documentNoteReferenceOffsets, type DocumentNoteReferenceOffset } from './document-note-boundary'
import { assertDocumentNoteReferenceOffsets, compareDocumentNoteBoundaries } from './document-note-boundary-edit'
import { isDocumentTextBoundary } from './document-story'

/** A body marker is generated decoration at a real paragraph boundary. Its
 * source runs are provenance; they are never editable duplicate label text. */
export interface DocumentNoteBodyMarker {
  id: string
  offset: number
  noteReferenceCount: number
  source: Record<string, unknown>
}

export interface DocumentNoteBodyMarkerCollection {
  version: 1
  kind: 'footnote' | 'endnote'
  markers: DocumentNoteBodyMarker[]
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function sourceFormat(value: unknown): void {
  if (!isRecord(value) || value.version !== 1 || typeof value.propertiesXml !== 'string'
    || !isRecord(value.namespaces) || Object.values(value.namespaces).some(namespace => typeof namespace !== 'string')
    || value.paragraphStyleId !== undefined && typeof value.paragraphStyleId !== 'string') {
    throw new TypeError('原生注释正文标记的文字格式来源无效。')
  }
}

export function validateDocumentNoteBodyMarkerSource(value: unknown): asserts value is Record<string, unknown> {
  if (!isRecord(value) || value.version !== 1) throw new TypeError('原生注释正文标记的来源版本或结构无效。')
  if (value.runs !== undefined) {
    if (typeof value.text !== 'string' || value.text.length > 4096 || !Array.isArray(value.runs)
      || !value.runs.length || value.runs.length > 4096) throw new TypeError('原生注释正文自定义标记的来源无效。')
    for (const run of value.runs) {
      if (!isRecord(run) || typeof run.sourceXml !== 'string') throw new TypeError('原生注释正文自定义标记缺少真实 run 来源。')
      sourceFormat(run.sourceFormat)
    }
  } else {
    if (typeof value.sourceXml !== 'string') throw new TypeError('原生注释正文自动标记缺少真实 XML 来源。')
    sourceFormat(value.sourceFormat)
  }
}

export function validateDocumentNoteBodyMarkerCollection(value: unknown, text: string,
  references: readonly DocumentNoteReferenceOffset[]): asserts value is DocumentNoteBodyMarkerCollection {
  assertDocumentNoteReferenceOffsets(references, text)
  if (!isRecord(value) || value.version !== 1 || !['footnote', 'endnote'].includes(String(value.kind))
    || !Array.isArray(value.markers) || value.markers.length > 4096) throw new TypeError('原生注释正文标记集合版本或结构无效。')
  const ids = new Set<string>()
  let previous: DocumentNoteBodyMarker | undefined
  for (const marker of value.markers) {
    if (!isRecord(marker) || typeof marker.id !== 'string' || !marker.id.trim() || marker.id.length > 4096
      || /[\r\n\0]/u.test(marker.id) || ids.has(marker.id)) throw new TypeError('原生注释正文标记的身份无效或重复。')
    if (typeof marker.offset !== 'number' || !isDocumentTextBoundary(text, marker.offset)
      || typeof marker.noteReferenceCount !== 'number' || !Number.isSafeInteger(marker.noteReferenceCount) || marker.noteReferenceCount < 0) {
      throw new RangeError('原生注释正文标记的 UTF-16 位置或累计引用序号无效。')
    }
    assertDocumentNoteReferenceBoundaryRank(references, marker.offset, marker.noteReferenceCount)
    validateDocumentNoteBodyMarkerSource(marker.source)
    if (previous && compareDocumentNoteBoundaries(previous.offset, previous.noteReferenceCount, marker.offset, marker.noteReferenceCount) > 0) {
      throw new RangeError('原生注释正文标记的来源顺序与段落位置不一致。')
    }
    ids.add(marker.id)
    previous = { id: marker.id, offset: marker.offset, noteReferenceCount: marker.noteReferenceCount, source: marker.source }
  }
}

export function documentNoteBodyMarkerCollection(block: DocumentBlock, text = runText(block.runs),
  references: readonly DocumentNoteReferenceOffset[] = documentNoteReferenceOffsets(block.runs)): DocumentNoteBodyMarkerCollection | undefined {
  const collection = block.metadata.officeNoteBodyMarkers
  if (collection === undefined) return undefined
  validateDocumentNoteBodyMarkerCollection(collection, text, references)
  return collection
}

export function documentNoteBodyMarkers(block: DocumentBlock): readonly DocumentNoteBodyMarker[] {
  return documentNoteBodyMarkerCollection(block)?.markers ?? []
}

export function setDocumentNoteBodyMarkers(block: DocumentBlock, collection: DocumentNoteBodyMarkerCollection,
  text = runText(block.runs), references: readonly DocumentNoteReferenceOffset[] = documentNoteReferenceOffsets(block.runs)): void {
  validateDocumentNoteBodyMarkerCollection(collection, text, references)
  block.metadata = { ...block.metadata, officeNoteBodyMarkers: cloneModel(collection) }
}

/** A copied paragraph keeps immutable native provenance but owns new marker
 * identities, including copies made inside a repeating section. */
export function remapDocumentNoteBodyMarkerIds(block: DocumentBlock): void {
  const collection = documentNoteBodyMarkerCollection(block)
  if (collection) setDocumentNoteBodyMarkers(block, { ...collection,
    markers: collection.markers.map(marker => ({ ...marker, id: documentId('note-body-marker') })) })
}
