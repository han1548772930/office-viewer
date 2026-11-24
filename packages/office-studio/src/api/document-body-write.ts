import type { DocumentBlock, DocumentContentControl, DocumentDefaultStyle, DocumentStoryContent, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentId, imageForBlock, runText, tableForBlock } from '../features/document/model/document-model'
import { documentTableGrid } from '../features/document/model/document-table-grid'
import { documentTableContinuationEntries } from '../features/document/model/document-table-continuations'
import { documentTextFields } from '../features/document/model/document-text-fields'
import { migrateDocumentStoryBlocks } from '../features/document/model/document-story-blocks'
import { documentParagraphToBody, documentBodyListForBlock } from './document-body-paragraph'
import { bodyRecord, documentRunStyleToBody, type UniverBodyRecord } from './document-body-style'
import { documentTableCellToBody, documentTableRowToBody, documentTableToBody } from './document-body-table'
import { DOCUMENT_BODY_TOKENS as T, documentBodyLocalOffset, type DocumentBodyParagraphPosition, type DocumentBodyTextSpan,
  type UniverDocumentBody, type UniverDocumentBodyData, type UniverDocumentCustomBlock, type UniverDocumentCustomRange,
  type UniverDocumentCustomTable, type UniverDocumentParagraph, type UniverDocumentTextRun } from './document-body-types'

interface ParagraphMap { block: DocumentBlock; start: number; end: number; spans: DocumentBodyTextSpan[] }
interface PendingRange { id: string; type: number; startBlockId: string; endBlockId: string; start: number; end: number; properties: UniverBodyRecord; source?: UniverBodyRecord }
interface RangeMarker { offset: number; end: number; text: string; order: number; sourceIndex?: number }

/** Rebase an imported, unmodeled annotation through the changed paragraph text. */
function rebasedOffset(before: string, after: string, offset: number, affinity: 'start' | 'end'): number {
  if (before === after) return Math.min(after.length, offset)
  let prefix = 0, suffix = 0
  while (prefix < before.length && prefix < after.length && before[prefix] === after[prefix]) prefix++
  while (suffix < before.length - prefix && suffix < after.length - prefix && before[before.length - suffix - 1] === after[after.length - suffix - 1]) suffix++
  if (offset <= prefix) return offset
  if (offset >= before.length - suffix) return Math.max(prefix, after.length - (before.length - offset))
  return affinity === 'start' ? prefix : after.length - suffix
}
function noteIdentity(run: InlineRun): { id: string; kind: 'footnote' | 'endnote' } | undefined {
  const raw = bodyRecord(run.metadata?.note ?? run.metadata?.noteReference), id = raw.noteId ?? raw.id ?? raw.footnoteId ?? raw.endnoteId
  const kind = raw.kind ?? raw.type ?? raw.noteType ?? (raw.footnoteId !== undefined ? 'footnote' : raw.endnoteId !== undefined ? 'endnote' : undefined)
  return (typeof id === 'string' || typeof id === 'number') && (kind === 'footnote' || kind === 'endnote') ? { id: String(id), kind } : undefined
}
function controlProperties(control: DocumentContentControl): UniverBodyRecord {
  const source = bodyRecord(control.metadata?.univerBodySdtSource), raw = cloneModel(bodyRecord(source.properties))
  raw.kind = control.type; raw.placement = control.placement ?? 'inline'
  for (const key of ['alias', 'tag', 'appearance', 'color', 'placeholder', 'placeholderDocPart', 'showingPlaceholder', 'temporary', 'lock', 'repeatingSection', 'docPart', 'dataBinding', 'rawSdtPr', 'rawSdtEndPr'] as const) {
    if (control[key] !== undefined) raw[key] = cloneModel(control[key]); else delete raw[key]
  }
  if (control.checkbox || control.checked !== undefined) raw.checkbox = { ...control.checkbox, checked: control.checked ?? control.checkbox?.checked ?? false }
  else delete raw.checkbox
  if (control.multiLine !== undefined) raw.text = { ...bodyRecord(raw.text), multiLine: control.multiLine }
  if (control.dateValue !== undefined || control.dateFormat !== undefined) raw.date = { ...bodyRecord(raw.date), fullDate: control.dateValue, format: control.dateFormat,
    locale: control.dateLocale, calendar: control.dateCalendar, storage: control.dateStorage }
  if (control.items) raw.listItems = cloneModel(control.items); else delete raw.listItems
  return raw
}

/** The stream is rebuilt from current blocks; source arrays never override authored text. */
export function writeDocumentBody(input: DocumentStoryContent, defaults: DocumentDefaultStyle, contentControls: readonly DocumentContentControl[] = [],
  positionOutput?: DocumentBodyParagraphPosition[]): UniverDocumentBodyData {
  const content = input.blocks === undefined ? migrateDocumentStoryBlocks(cloneModel(input), [], { kind: 'note', noteKind: 'footnote', noteId: 'public-body' }, defaults) : input
  const provenance = bodyRecord(content.metadata?.univerBodySource), sourceBody = cloneModel(bodyRecord(provenance.body)), sourceResources = bodyRecord(provenance.resources)
  const paragraphs: UniverDocumentParagraph[] = [], textRuns: UniverDocumentTextRun[] = [], tables: UniverDocumentCustomTable[] = [], customBlocks: UniverDocumentCustomBlock[] = [], rawCustomBlocks: UniverDocumentCustomBlock[] = []
  const tableSource: Record<string, UniverBodyRecord> = {}, drawings = cloneModel(bodyRecord(sourceResources.drawings)) as Record<string, UniverBodyRecord>, lists: Record<string, UniverBodyRecord> = {}
  const writtenDrawingIds = new Set<string>(), drawingAliases = new Map<string, string>()
  const sourceDrawingIds = new Set([...documentArray(sourceBody.customBlocks), ...documentArray(sourceBody.docxRawCustomBlocks)]
    .map(raw => raw.blockId).filter((id): id is string => typeof id === 'string' && drawings[id] !== undefined))
  const maps = new Map<string, ParagraphMap>(), ranges: PendingRange[] = [], sections: UniverBodyRecord[] = [], chunks: string[] = []
  const sourceCustomTokens = new Map<UniverDocumentCustomBlock, number>()
  const renderedPageBreaks: number[] = []
  let offset = 0
  const append = (text: string) => { chunks.push(text); offset += text.length }
  const registerDrawing = (id: string, drawing: UniverBodyRecord, sourceId?: unknown) => {
    drawings[id] = { ...drawing, drawingId: id }; writtenDrawingIds.add(id)
    if (typeof sourceId === 'string' && sourceId !== id) drawingAliases.set(sourceId, id)
  }
  const writeBlocks = (blocks: readonly DocumentBlock[]) => {
    for (const block of blocks) {
      const table = tableForBlock(block)
      if (table) {
        const start = offset, source = bodyRecord(block.metadata.univerBodyTableSource), grid = documentTableGrid(table), rows: UniverBodyRecord[] = []
        append(T.tableStart)
        for (const [rowIndex, row] of table.rows.entries()) {
          append(T.rowStart); const cells: UniverBodyRecord[] = []
          for (let column = 0; column < grid.columnCount; column++) {
            const owner = grid.cellAt(rowIndex, column)!
            if (owner.cell.style.univerBodyOmittedCell) continue
            append(T.cellStart)
            if (owner.row === rowIndex && owner.column === column) {
              writeBlocks(owner.cell.blocks); cells.push(documentTableCellToBody(owner.cell))
            } else {
              const continuation = documentTableContinuationEntries(owner.cell).find(entry => entry.offset === rowIndex - owner.row)
              if (owner.column === column && continuation) writeBlocks(continuation.content.blocks)
              else append(T.paragraph)
              cells.push({ ...cloneModel(bodyRecord(continuation?.content.source?.univerBodyCellSource)), rowSpan: 0, columnSpan: 0 })
            }
            append(T.cellEnd)
          }
          rows.push(documentTableRowToBody(table, rowIndex, cells)); append(T.rowEnd)
        }
        append(T.tableEnd)
        tables.push({ ...source, startIndex: start, endIndex: offset, tableId: block.id }); tableSource[block.id] = documentTableToBody(table, block.id, rows)
        continue
      }
      const image = imageForBlock(block)
      if (image) {
        const source = bodyRecord(block.metadata.univerBodyCustomBlock), drawing = cloneModel(bodyRecord(block.metadata.univerBodyDrawingSource)), start = offset
        append(T.customBlock); customBlocks.push({ ...source, startIndex: start, blockId: block.id })
        drawing.drawingId = block.id; drawing.drawingType = 0; drawing.source = image.src; drawing.description = image.alt
        drawing.docTransform = { ...bodyRecord(drawing.docTransform), size: { ...bodyRecord(bodyRecord(drawing.docTransform).size), width: image.width, height: image.height } }
        registerDrawing(block.id, drawing, source.blockId); continue
      }
      if (block.type === 'sectionBreak') {
        const section = cloneModel(bodyRecord(block.metadata.univerBodySectionSource)), types = ['unspecified', 'continuous', 'nextColumn', 'nextPage', 'evenPage', 'oddPage']
        section.startIndex = offset; section.sectionId = block.id; section.sectionType = types.indexOf(block.sectionBreak?.type ?? 'unspecified')
        sections.push(section); append(T.section); continue
      }
      if (block.type === 'pageBreak' || block.type === 'columnBreak') {
        if (block.type === 'pageBreak' && bodyRecord(block.metadata.univerBodyBreakSource).soft === true) renderedPageBreaks.push(offset)
        append(block.type === 'pageBreak' ? T.pageBreak : T.columnBreak); continue
      }
      if (block.type === 'divider') { append(T.paragraph); paragraphs.push({ startIndex: offset - 1, paragraphId: block.id, paragraphStyle: { borderBottom: { width: 0.75, color: { rgb: 'auto' }, dashStyle: 1, padding: 0 } } }); continue }
      const start = offset, spans: DocumentBodyTextSpan[] = [], source = bodyRecord(block.metadata.univerBodyParagraphSource)
      let local = 0
      for (const run of block.runs) {
        const from = offset, note = noteIdentity(run), sourceCustom = bodyRecord(run.metadata?.univerBodyCustomBlock), runSource = bodyRecord(run.metadata?.univerBodyRunSource), publicRun = bodyRecord(runSource.textRun)
        if (note || Object.keys(sourceCustom).length || run.metadata?.univerBodyDrawingSource) {
          const copied = typeof runSource.runId === 'string' && runSource.runId !== run.id
          const blockId = !copied && typeof sourceCustom.blockId === 'string' ? sourceCustom.blockId : run.id
          append(T.customBlock); (run.metadata?.univerBodyRawCustomBlock === true ? rawCustomBlocks : customBlocks).push({ ...sourceCustom, startIndex: from, blockId,
            ...(note ? { noteType: note.kind, noteId: note.id } : {}) })
          if (run.metadata?.univerBodyDrawingSource) registerDrawing(blockId, cloneModel(bodyRecord(run.metadata.univerBodyDrawingSource)), sourceCustom.blockId)
        } else append(run.text)
        spans.push({ sourceStart: local, sourceEnd: local + run.text.length, targetStart: from, targetEnd: offset })
        if (offset > from) textRuns.push({ ...publicRun, st: from, ed: offset, ts: documentRunStyleToBody(run) })
        const href = Reflect.get(run, 'href') ?? bodyRecord(run.metadata?.link).target
        if (typeof href === 'string' && href) {
          const raw = bodyRecord(run.metadata?.univerBodyHyperlinkSource), id = typeof raw.rangeId === 'string' ? raw.rangeId : `hyperlink:${run.id}`
          const previous = ranges.find(range => range.type === 0 && range.id === id)
          if (previous && previous.properties.url === href && (previous.endBlockId !== block.id || previous.end === local)) {
            previous.endBlockId = block.id; previous.end = local + run.text.length
          } else ranges.push({ id: previous ? `${id}:${run.id}` : id, type: 0, startBlockId: block.id, endBlockId: block.id, start: local, end: local + run.text.length, properties: { ...bodyRecord(raw.properties), url: href }, source: raw })
        }
        local += run.text.length
      }
      const end = offset; maps.set(block.id, { block, start, end, spans })
      if (source.terminated !== false) { paragraphs.push({ ...documentParagraphToBody(block, offset), startIndex: offset, paragraphId: block.id }); append(T.paragraph) }
      const list = documentBodyListForBlock(block); if (list) lists[list.id] = list.value
      for (const field of documentTextFields(block)) {
        const raw = bodyRecord(field.metadata?.univerBodyFieldSource)
        ranges.push({ id: field.id, type: 1, startBlockId: block.id, endBlockId: field.endBlockId ?? block.id, start: field.startOffset, end: field.endOffset,
          properties: { ...bodyRecord(raw.properties), ...bodyRecord(bodyRecord(field.metadata?.source).properties), instruction: field.instruction ?? field.field,
            fieldType: field.field, cachedResult: field.cachedResult, locked: field.locked }, source: raw })
      }
    }
  }
  writeBlocks(content.blocks!)
  if (chunks.at(-1) !== T.section) {
    if (!chunks.length) { append(T.paragraph); paragraphs.push({ startIndex: 0, paragraphId: documentId('paragraph') }) }
    append(T.section)
  }
  const baseStream = chunks.join('')
  const sourceStream = typeof sourceBody.dataStream === 'string' ? sourceBody.dataStream : ''
  const finalSection = (Array.isArray(sourceBody.sectionBreaks) ? sourceBody.sectionBreaks : []).map(bodyRecord).find(raw =>
    raw.startIndex === sourceStream.length - 1 && sourceStream.endsWith(T.section))
  if (finalSection) sections.push({ ...cloneModel(finalSection), startIndex: baseStream.length - 1 })
  const atLocal = (id: string, local: number, affinity: 'start' | 'end' = 'start') => {
    const map = maps.get(id)
    if (!map) throw new RangeError(`文档导出范围的块不存在：${id}`)
    return documentBodyLocalOffset({ block: map.block, start: 0, end: runText(map.block.runs).length, spans: map.spans }, local, affinity)
  }
  const atSource = (old: number, affinity: 'start' | 'end') => {
    for (const map of maps.values()) {
      const source = bodyRecord(map.block.metadata.univerBodyParagraphSource)
      if (typeof source.start !== 'number' || typeof source.end !== 'number' || old < source.start || old > source.end) continue
      const sourcePosition: DocumentBodyParagraphPosition = { block: map.block, start: source.start, end: source.end, spans: Array.isArray(source.spans) ? source.spans as DocumentBodyTextSpan[] : [] }
      const local = documentBodyLocalOffset(sourcePosition, old, affinity), before = typeof source.baselineText === 'string' ? source.baselineText : runText(map.block.runs)
      return atLocal(map.block.id, rebasedOffset(before, runText(map.block.runs), local, affinity), affinity)
    }
    const after = [...maps.values()].find(map => typeof bodyRecord(map.block.metadata.univerBodyParagraphSource).start === 'number' && Number(bodyRecord(map.block.metadata.univerBodyParagraphSource).start) > old)
    return after?.start ?? Math.max(0, baseStream.length - 1)
  }
  for (const control of contentControls) {
    const firstId = control.blockRange?.blockIds[0] ?? control.blockId, lastId = control.blockRange?.blockIds.at(-1) ?? control.blockId
    if (!maps.has(firstId) || !maps.has(lastId)) continue
    ranges.push({ id: control.id, type: 2, startBlockId: firstId, endBlockId: lastId, start: control.blockRange ? 0 : control.startOffset,
      end: control.blockRange ? runText(maps.get(lastId)!.block.runs).length : control.endOffset, properties: controlProperties(control), source: bodyRecord(control.metadata?.univerBodySdtSource) })
  }
  for (const range of ranges) if (range.type === 1) {
    const paragraphs = [...maps.values()], first = paragraphs.findIndex(map => map.block.id === range.startBlockId), last = paragraphs.findIndex(map => map.block.id === range.endBlockId)
    if (first < 0 || last < first) throw new RangeError('文档字段结果范围的段落不存在。')
    range.properties.cachedResult = paragraphs.slice(first, last + 1).map((map, index, values) => runText(map.block.runs)
      .slice(index === 0 ? range.start : 0, index === values.length - 1 ? range.end : undefined)).join(T.paragraph)
  }
  const markers: RangeMarker[] = [], customRanges: UniverDocumentCustomRange[] = [], generatedIds = new Set(ranges.map(range => range.id))
  for (const map of maps.values()) {
    const source = bodyRecord(map.block.metadata.univerBodyParagraphSource), before = typeof source.baselineText === 'string' ? source.baselineText : runText(map.block.runs)
    for (const value of Array.isArray(source.boundaryTokens) ? source.boundaryTokens : []) {
      const token = bodyRecord(value)
      if (typeof token.offset !== 'number' || typeof token.text !== 'string' || token.text.length !== 1
        || !Object.values(T).includes(token.text as typeof T[keyof typeof T]) || typeof token.sourceIndex !== 'number') throw new TypeError('文档保留的结构标记无效。')
      const position = atLocal(map.block.id, rebasedOffset(before, runText(map.block.runs), token.offset, 'start'))
      markers.push({ offset: position, end: position, text: token.text, order: -1, sourceIndex: token.sourceIndex })
      if (token.text === T.customBlock) {
        const raw = [...(Array.isArray(sourceBody.customBlocks) ? sourceBody.customBlocks : []),
          ...(Array.isArray(sourceBody.docxRawCustomBlocks) ? sourceBody.docxRawCustomBlocks : [])].map(bodyRecord).find(raw => raw.startIndex === token.sourceIndex)
        if (!raw || typeof raw.blockId !== 'string') throw new RangeError('文档保留的 customBlock 缺少来源定义。')
        const descriptor = { ...raw, startIndex: position, blockId: raw.blockId }; customBlocks.push(descriptor); sourceCustomTokens.set(descriptor, token.sourceIndex)
      }
    }
  }
  for (const raw of Array.isArray(sourceBody.customRanges) ? sourceBody.customRanges : []) {
    const range = bodyRecord(raw)
    if (typeof range.rangeId !== 'string' || generatedIds.has(range.rangeId)) continue
    if (typeof range.startIndex !== 'number' || typeof range.endIndex !== 'number' || typeof range.rangeType !== 'number') continue
    const start = atSource(range.startIndex, 'start'), end = atSource(range.endIndex + 1, 'end')
    const index = customRanges.length
    customRanges.push({ ...range, rangeId: range.rangeId, rangeType: range.rangeType, startIndex: start,
      endIndex: Math.max(start, streamRangeEnd(sourceStream, range.endIndex, end)) })
    if (sourceStream[range.startIndex] === T.rangeStart) markers.push({ offset: start, end, text: T.rangeStart, order: index })
    if (sourceStream[range.endIndex] === T.rangeEnd) markers.push({ offset: end, end: start, text: T.rangeEnd, order: index })
  }
  for (const range of ranges) {
    const start = atLocal(range.startBlockId, range.start), end = atLocal(range.endBlockId, range.end, 'end')
    if (end < start) throw new RangeError('文档导出自定义范围方向无效。')
    const index = customRanges.length
    customRanges.push({ ...range.source, startIndex: start, endIndex: end, rangeId: range.id, rangeType: range.type, properties: range.properties })
    markers.push({ offset: start, end, text: T.rangeStart, order: index }, { offset: end, end: start, text: T.rangeEnd, order: index })
  }
  const markerPriority = (marker: RangeMarker) => marker.sourceIndex !== undefined ? 1 : marker.offset === marker.end ? 2 : marker.text === T.rangeEnd ? 0 : 3
  markers.sort((a, b) => a.offset - b.offset || markerPriority(a) - markerPriority(b)
    || (a.sourceIndex !== undefined && b.sourceIndex !== undefined ? a.sourceIndex - b.sourceIndex
      : a.offset === a.end && b.offset === b.end ? a.order - b.order || (a.text === b.text ? 0 : a.text === T.rangeStart ? -1 : 1)
        : a.text === T.rangeStart ? b.end - a.end || a.order - b.order : b.end - a.end || b.order - a.order))
  const markerCount = (position: number, inclusive: boolean) => {
    let low = 0, high = markers.length
    while (low < high) {
      const middle = (low + high) >>> 1, offset = markers[middle]!.offset
      if (offset < position || inclusive && offset === position) low = middle + 1
      else high = middle
    }
    return low
  }
  const beforeMarker = (position: number) => position + markerCount(position, false)
  const afterMarker = (position: number) => position + markerCount(position, true)
  const sourceTokenPosition = (sourceIndex: number) => {
    const index = markers.findIndex(marker => marker.sourceIndex === sourceIndex)
    return index < 0 ? undefined : markers[index]!.offset + index
  }
  const inserted: string[] = []; let consumed = 0
  for (const marker of markers) { inserted.push(baseStream.slice(consumed, marker.offset), marker.text); consumed = marker.offset }
  inserted.push(baseStream.slice(consumed))
  for (const item of textRuns) { item.st = afterMarker(item.st); item.ed = beforeMarker(item.ed) }
  for (const item of paragraphs) item.startIndex = afterMarker(item.startIndex)
  for (const item of [...customBlocks, ...rawCustomBlocks]) item.startIndex = sourceCustomTokens.has(item) ? sourceTokenPosition(sourceCustomTokens.get(item)!)! : afterMarker(item.startIndex)
  for (const item of tables) { item.startIndex = afterMarker(item.startIndex); item.endIndex = beforeMarker(item.endIndex) }
  for (const section of sections) section.startIndex = afterMarker(Number(section.startIndex))
  for (const [index, position] of renderedPageBreaks.entries()) renderedPageBreaks[index] = afterMarker(position)
  for (const [index, range] of customRanges.entries()) {
    const beginMarker = markers.findIndex(marker => marker.order === index && marker.text === T.rangeStart), endMarker = markers.findIndex(marker => marker.order === index && marker.text === T.rangeEnd)
    range.startIndex = beginMarker >= 0 ? range.startIndex + beginMarker : afterMarker(range.startIndex)
    range.endIndex = endMarker >= 0 ? range.endIndex + endMarker : beforeMarker(range.endIndex + 1) - 1
  }
  if (positionOutput) for (const map of maps.values()) {
    const spans: DocumentBodyTextSpan[] = [], localAt = (span: DocumentBodyTextSpan, value: number, affinity: 'start' | 'end') => {
      if (value <= span.targetStart) return span.sourceStart
      if (value >= span.targetEnd) return span.sourceEnd
      return span.targetEnd - span.targetStart === span.sourceEnd - span.sourceStart ? span.sourceStart + value - span.targetStart
        : affinity === 'start' ? span.sourceStart : span.sourceEnd
    }
    for (const span of map.spans) {
      const cuts = [span.targetStart, ...markers.slice(markerCount(span.targetStart, true), markerCount(span.targetEnd, false))
        .map(marker => marker.offset), span.targetEnd]
      for (let index = 0; index < cuts.length - 1; index++) {
        const start = cuts[index]!, end = cuts[index + 1]!
        if (end <= start) continue
        spans.push({ sourceStart: afterMarker(start), sourceEnd: beforeMarker(end),
          targetStart: localAt(span, start, 'start'), targetEnd: localAt(span, end, 'end') })
      }
    }
    for (let index = markerCount(map.start, false); index < markerCount(map.end, true); index++) {
      const marker = markers[index]!, span = map.spans.find(span => marker.offset >= span.targetStart && marker.offset <= span.targetEnd)
      const local = span ? localAt(span, marker.offset, 'start') : 0
      spans.push({ sourceStart: marker.offset + index, sourceEnd: marker.offset + index + 1, targetStart: local, targetEnd: local })
    }
    spans.sort((a, b) => a.sourceStart - b.sourceStart || a.sourceEnd - b.sourceEnd)
    positionOutput.push({ block: map.block, start: beforeMarker(map.start), end: afterMarker(map.end), spans })
  }
  const body: UniverDocumentBody = { ...sourceBody, dataStream: inserted.join(''), paragraphs, textRuns,
    ...(tables.length ? { tables } : { tables: undefined }), ...(customBlocks.length ? { customBlocks } : { customBlocks: undefined }),
    ...(rawCustomBlocks.length ? { docxRawCustomBlocks: rawCustomBlocks } : { docxRawCustomBlocks: undefined }),
    ...(renderedPageBreaks.length ? { renderedPageBreaks } : { renderedPageBreaks: undefined }),
    ...(customRanges.length ? { customRanges } : { customRanges: undefined }), ...(sections.length ? { sectionBreaks: sections } : { sectionBreaks: undefined }) }
  for (const key of ['customDecorations', 'docxRawBlocks', 'docxExportExcludedRanges', 'columnGroups', 'blockRanges'] as const) if (Array.isArray(sourceBody[key])) body[key] = sourceBody[key].map(value => {
    const raw = cloneModel(bodyRecord(value))
    if (typeof raw.startIndex === 'number') raw.startIndex = sourceTokenPosition(raw.startIndex) ?? afterMarker(atSource(raw.startIndex, 'start'))
    if (typeof raw.endIndex === 'number') raw.endIndex = sourceTokenPosition(raw.endIndex) ?? beforeMarker(atSource(raw.endIndex + 1, 'end')) - 1
    return raw
  })
  // Remove old custom-block resources after deletion/copy, while retaining source
  // resources that never belonged to a modeled custom block (e.g. group parts).
  for (const id of sourceDrawingIds) if (!writtenDrawingIds.has(id)) delete drawings[id]
  const drawingOrder = Array.isArray(sourceResources.drawingsOrder) ? sourceResources.drawingsOrder
    .filter((id): id is string => typeof id === 'string').map(id => drawings[id] ? id : drawingAliases.get(id) ?? id).filter(id => drawings[id]) : []
  const drawingsOrder = [...new Set([...drawingOrder, ...Object.keys(drawings)])]
  return { body, ...(Object.keys(tableSource).length ? { tableSource } : {}), ...(Object.keys(drawings).length ? { drawings, drawingsOrder } : {}),
    ...(Object.keys(lists).length ? { lists } : {}) }
}

function streamRangeEnd(stream: string, sourceEnd: number, offset: number): number {
  return stream[sourceEnd] === T.rangeEnd ? offset : offset - 1
}

function documentArray(value: unknown): UniverBodyRecord[] {
  return Array.isArray(value) ? value.map(bodyRecord) : []
}
