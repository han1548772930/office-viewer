import type { DocumentBlock, DocumentDefaultStyle } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { createParagraph, runText } from '../features/document/model/document-model'
import { bodyRecord } from './document-body-style'
import { documentBodyOfficeSource } from './document-body-source'
import { DOCUMENT_BODY_TOKENS as T, documentBodyArray, documentBodyIdentity,
  type DocumentBodyParagraphPosition, type DocumentBodyStructureIndex } from './document-body-types'

/** A picture-only SDT still needs an actual caret host. Its empty paragraph is
 * projected onto the picture's existing public atom and physical Canvas line. */
export function prepareDocumentBodyStructuralInputs(value: unknown, positions: DocumentBodyParagraphPosition[],
  index: DocumentBodyStructureIndex, defaults: DocumentDefaultStyle): void {
  const body = bodyRecord(value), stream = typeof body.dataStream === 'string' ? body.dataStream : ''
  for (const raw of documentBodyArray(body.customRanges, 'customRanges')) {
    if (raw.rangeType !== 2 || typeof raw.startIndex !== 'number' || typeof raw.endIndex !== 'number') continue
    const from = stream[raw.startIndex] === T.rangeStart ? raw.startIndex + 1 : raw.startIndex
    const to = stream[raw.endIndex] === T.rangeEnd ? raw.endIndex : raw.endIndex + 1
    if (positions.some(position => position.start < to && position.end > from)) continue
    const images = index.blocks.filter(entry => entry.block.type === 'image' && entry.start >= from && entry.end <= to)
    const image = images[0]
    if (!image || images.some(entry => entry.container !== image.container)) continue
    const container = image.container as DocumentBlock[], at = container.indexOf(image.block)
    const existing = container[at - 1]
    if (existing && existing.id === image.block.metadata.anchorBlockId && runText(existing.runs) === '') {
      if (!positions.some(position => position.block.id === existing.id)) positions.push({ block: existing, start: image.start, end: image.end,
        spans: [{ sourceStart: image.start, sourceEnd: image.end, targetStart: 0, targetEnd: 0 }] })
      continue
    }
    const source = bodyRecord(image.block.metadata.pictureInputAnchorSource)
    if (source.version !== undefined && source.version !== 1) throw new TypeError('图片输入锚点来源版本无效。')
    const anchor = createParagraph('', defaults.text), expected = documentBodyOfficeSource(raw, '内容控件').blockId
    const anchorId = image.block.metadata.anchorBlockId ?? source.anchorId ?? (expected !== image.block.id ? expected : undefined)
    if (anchorId !== undefined) anchor.id = documentBodyIdentity(anchorId, '图片输入锚点')
    if (index.blocks.some(entry => entry.block.id === anchor.id)) throw new TypeError('图片输入锚点身份重复：' + anchor.id)
    if (source.style !== undefined) anchor.style = { ...defaults.paragraph, ...cloneModel(bodyRecord(source.style)) }
    const emptyRun = bodyRecord(source.emptyRun)
    if (emptyRun.runId !== undefined) anchor.runs[0]!.id = documentBodyIdentity(emptyRun.runId, '图片空运行')
    if (emptyRun.marks !== undefined) Object.assign(anchor.runs[0]!, cloneModel(bodyRecord(emptyRun.marks)), { text: '' })
    anchor.metadata.univerBodyStructuralInputAnchor = { version: 1, imageId: image.block.id }
    image.block.metadata = { ...image.block.metadata, pictureInputAnchor: true, anchorBlockId: anchor.id, anchorOffset: 0 }
    container.splice(at, 0, anchor)
    positions.push({ block: anchor, start: image.start, end: image.end,
      spans: [{ sourceStart: image.start, sourceEnd: image.end, targetStart: 0, targetEnd: 0 }] })
    index.blocks.push({ block: anchor, container, start: image.start, end: image.start })
  }
  positions.sort((a, b) => a.start - b.start || a.end - b.end)
}

export function documentBodyProjectedPictureInput(block: DocumentBlock, blocks: readonly DocumentBlock[]): DocumentBlock | undefined {
  const source = bodyRecord(block.metadata.univerBodyStructuralInputAnchor), image = blocks[blocks.indexOf(block) + 1]
  if (source.version !== 1 || !image || image.type !== 'image' || image.id !== source.imageId || image.metadata.anchorBlockId !== block.id
    || image.metadata.anchorOffset !== 0 || runText(block.runs) !== '') return undefined
  return image
}

/** Keep only empty-host identity and marks; the picture atom remains the content source. */
export function documentBodyPictureInputSource(anchor: DocumentBlock): Record<string, unknown> {
  const run = anchor.runs[0], marks = run ? { ...run } : {}
  for (const key of ['id', 'text']) Reflect.deleteProperty(marks, key)
  return { version: 1, anchorId: anchor.id, style: cloneModel(anchor.style),
    ...(run ? { emptyRun: { runId: run.id, marks: cloneModel(marks) } } : {}) }
}
