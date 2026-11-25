import type { DocumentBlock, DocumentContentControl, DocumentDefaultStyle, DocumentStoryContent, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentId, imageForBlock, runText, tableForBlock } from '../features/document/model/document-model'
import { documentTableGrid } from '../features/document/model/document-table-grid'
import { documentTableContinuationEntries } from '../features/document/model/document-table-continuations'
import { documentTextFields } from '../features/document/model/document-text-fields'
import { isDocumentStructuredBlock, validateDocumentStructuredBlock } from '../features/document/model/document-structural-blocks'
import { documentQuoteLineColor, documentQuoteParagraphBorderSource } from '../features/document/model/document-structured-appearance'
import { migrateDocumentStoryBlocks } from '../features/document/model/document-story-blocks'
import { documentParagraphToBody, documentBodyListForBlock } from './document-body-paragraph'
import { DocumentBodyListResources, documentBodyListInstanceIds } from './document-body-lists'
import { documentBodyNoteStreamOffset } from './document-body-note-boundaries'
import { DOCUMENT_BODY_SECTION_TYPES, documentSectionSetupToBody } from './document-body-section'
import { bodyRecord, documentRunStyleToBody, type UniverBodyRecord } from './document-body-style'
import { documentTableCellToBody, documentTableRowToBody, documentTableToBody } from './document-body-table'
import { documentBodyBlockSource, documentBodyRunSource, documentBodySourceMetadata } from './document-body-source'
import { documentBodyControlSource, documentBodyPublicControlPlacement, documentBodyStructuralControlBounds } from './document-body-control-ranges'
import { documentBodyPictureInputSource, documentBodyProjectedPictureInput } from './document-body-structural-input'
import { documentCalloutToBody, documentColumnGroupToBody } from './document-body-structures'
import { getLatexFormula, latexFormulaRangeId, DOCUMENT_LATEX_RANGE_KIND } from '../features/document/latex/latex-formula'
import { documentBodyLatexConfig, isDocumentBodyLatexRange } from './document-body-latex'
import { documentImageToDrawing } from './document-image-body'
import { documentImageAtom } from '../features/document/images/document-image-atom'
import { DOCUMENT_BODY_TOKENS as T, documentBodyLocalOffset, type DocumentBodyEncodeOptions, type DocumentBodyParagraphPosition, type DocumentBodyTextSpan,
  type UniverDocumentBody, type UniverDocumentBodyData, type UniverDocumentCustomBlock, type UniverDocumentCustomRange,
  type DocumentBodyStructureIndex, type UniverDocumentCustomTable, type UniverDocumentParagraph, type UniverDocumentTextRun,
  type UniverDocumentBlockRange, type UniverDocumentColumnGroup } from './document-body-types'

interface ParagraphMap { block: DocumentBlock; start: number; end: number; spans: DocumentBodyTextSpan[] }
interface PendingRange { id: string; type: number; startBlockId: string; endBlockId: string; start: number; end: number; startNoteReferenceCount?: number; endNoteReferenceCount?: number; properties: UniverBodyRecord; source?: UniverBodyRecord; atomic?: boolean; absolute?: { start: number; end: number } }
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
  for (const key of ['alias', 'title', 'tag', 'appearance', 'color', 'placeholder', 'placeholderDocPart', 'showingPlaceholder', 'temporary', 'lock', 'value', 'repeatingSection', 'docPart', 'dataBinding', 'rawSdtPr', 'rawSdtEndPr'] as const) {
    if (control[key] !== undefined) raw[key] = cloneModel(control[key]); else delete raw[key]
  }
  if (control.checkbox || control.checked !== undefined) raw.checkbox = { ...control.checkbox, checked: control.checked ?? control.checkbox?.checked ?? false }
  else delete raw.checkbox
  if (control.multiLine !== undefined) raw.text = { ...bodyRecord(raw.text), multiLine: control.multiLine }
  else if (raw.text !== undefined) { const text = { ...bodyRecord(raw.text) }; delete text.multiLine; raw.text = text }
  if ([control.dateValue, control.dateFormat, control.dateLocale, control.dateCalendar, control.dateStorage].some(value => value !== undefined)) raw.date = { ...bodyRecord(raw.date), fullDate: control.dateValue, format: control.dateFormat,
    locale: control.dateLocale, calendar: control.dateCalendar, storage: control.dateStorage }
  else delete raw.date
  if (control.items) raw.listItems = cloneModel(control.items); else delete raw.listItems
  return raw
}

/** The stream is rebuilt from current blocks; source arrays never override authored text. */
export function writeDocumentBody(input: DocumentStoryContent, defaults: DocumentDefaultStyle, contentControls: readonly DocumentContentControl[] = [],
  positionOutput?: DocumentBodyParagraphPosition[], options: DocumentBodyEncodeOptions = {}): UniverDocumentBodyData {
  const content = input.blocks === undefined ? migrateDocumentStoryBlocks(cloneModel(input), [], { kind: 'note', noteKind: 'footnote', noteId: 'public-body' }, defaults) : input
  const provenance = bodyRecord(content.metadata?.univerBodySource), sourceBody = cloneModel(bodyRecord(provenance.body)), sourceResources = bodyRecord(provenance.resources)
  const paragraphs: UniverDocumentParagraph[] = [], textRuns: UniverDocumentTextRun[] = [], tables: UniverDocumentCustomTable[] = [], customBlocks: UniverDocumentCustomBlock[] = [], rawCustomBlocks: UniverDocumentCustomBlock[] = []
  const blockRanges: UniverDocumentBlockRange[] = [], columnGroups: UniverDocumentColumnGroup[] = []
  const tableSource = cloneModel(bodyRecord(sourceResources.tableSource)) as Record<string, UniverBodyRecord>,
    drawings = cloneModel(bodyRecord(sourceResources.drawings)) as Record<string, UniverBodyRecord>, lists = cloneModel(bodyRecord(sourceResources.lists)) as Record<string, UniverBodyRecord>
  const callouts = cloneModel(bodyRecord(sourceResources.callouts)) as Record<string, UniverBodyRecord>, codes = cloneModel(bodyRecord(sourceResources.codes)) as Record<string, UniverBodyRecord>
  const latexFormulas = cloneModel(bodyRecord(sourceResources.latexFormulas)) as Record<string, UniverBodyRecord>, writtenLatex = new Set<string>()
  const writtenCallouts = new Set<string>(), writtenCodes = new Set<string>()
  const sourceTableIds = new Set(documentArray(sourceBody.tables).map(raw => raw.tableId).filter((id): id is string => typeof id === 'string'))
  const writtenTableIds = new Set<string>()
  const writtenLists = new DocumentBodyListResources(), sourceListIds = new Set(documentArray(sourceBody.paragraphs)
    .map(paragraph => bodyRecord(paragraph.bullet).listType).filter((id): id is string => typeof id === 'string'))
  const writtenDrawingIds = new Set<string>(), drawingAliases = new Map<string, string>()
  const sourceDrawingIds = new Set([...documentArray(sourceBody.customBlocks), ...documentArray(sourceBody.docxRawCustomBlocks)]
    .map(raw => raw.blockId).filter((id): id is string => typeof id === 'string' && drawings[id] !== undefined))
  const maps = new Map<string, ParagraphMap>(), ranges: PendingRange[] = [], sections: UniverBodyRecord[] = [], chunks: string[] = []
  const structures: DocumentBodyStructureIndex = { blocks: [], rows: [], cells: [] }
  const listInstanceIds = documentBodyListInstanceIds(content.blocks!)
  const sourceModeledRangeIds = new Set(Array.isArray(provenance.modeledRangeIds) ? provenance.modeledRangeIds.filter((id): id is string => typeof id === 'string') : [])
  const retainedSourceRangeIds = new Set<string>()
  const sourceCustomTokens = new Map<UniverDocumentCustomBlock, number>()
  const sourceRangeTokens = new Map<UniverDocumentCustomRange, number>()
  const renderedPageBreaks: number[] = []
  let offset = 0
  let currentSection = documentSectionSetupToBody(options.initialSection?.setup, content.metadata?.univerBodyFirstSectionSource)
  if (options.initialSection) currentSection.sectionId = options.initialSection.id
  const append = (text: string) => { chunks.push(text); offset += text.length }
  const appendSection = (required = false) => {
    if (required || Object.keys(currentSection).length) sections.push({ ...cloneModel(currentSection),
      sectionId: typeof currentSection.sectionId === 'string' ? currentSection.sectionId : 'first', startIndex: offset })
    append(T.section)
  }
  const registerDrawing = (id: string, drawing: UniverBodyRecord, sourceId?: unknown) => {
    if (writtenDrawingIds.has(id)) throw new TypeError(`文档绘图标识重复：${id}`)
    drawings[id] = { ...drawing, drawingId: id }; writtenDrawingIds.add(id)
    if (typeof sourceId === 'string' && sourceId !== id) drawingAliases.set(sourceId, id)
  }
  const writeBlocks = (blocks: readonly DocumentBlock[], depth = 0, quote?: { id: string; lineColor: string }) => {
    if (depth > 64) throw new RangeError('文档导出结构嵌套超过预算。')
    for (const block of blocks) {
      if (isDocumentStructuredBlock(block)) {
        validateDocumentStructuredBlock(block)
        const start = offset
        if (block.type === 'columns') {
          append(T.groupStart)
          for (const column of block.columns!.columns) { append(T.columnStart); writeBlocks(column.blocks, depth + 1, quote); append(T.columnEnd) }
          append(T.groupEnd)
          columnGroups.push({ ...documentColumnGroupToBody(block), startIndex: start, endIndex: offset - 1, columnGroupId: block.id, officeStudio: documentBodyBlockSource(block) })
        } else {
          append(T.blockStart); writeBlocks(block.children!, depth + 1, block.type === 'quote' ? { id: block.id, lineColor: documentQuoteLineColor(block) } : quote); append(T.blockEnd)
          blockRanges.push({ ...bodyRecord(block.metadata.univerBodyBlockRangeSource), startIndex: start, endIndex: offset - 1,
            blockId: block.id, blockType: block.type as 'quote' | 'callout' | 'code', officeStudio: documentBodyBlockSource(block) })
          if (block.type === 'callout') { callouts[block.id] = documentCalloutToBody(block.callout!, block.metadata.univerBodyCalloutSource); writtenCallouts.add(block.id) }
          if (block.type === 'code') { codes[block.id] = { ...bodyRecord(block.metadata.univerBodyCodeSource), ...block.code! }; writtenCodes.add(block.id) }
        }
        structures.blocks.push({ block, container: blocks, start, end: offset }); continue
      }
      const projectedImage = documentBodyProjectedPictureInput(block, blocks)
      if (projectedImage) {
        maps.set(block.id, { block, start: offset, end: offset + 1, spans: [{ sourceStart: 0, sourceEnd: 0, targetStart: offset, targetEnd: offset + 1 }] })
        structures.blocks.push({ block, container: blocks, start: offset, end: offset })
        continue
      }
      const table = tableForBlock(block)
      if (table) {
        const start = offset, source = bodyRecord(block.metadata.univerBodyTableSource), grid = documentTableGrid(table), rows: UniverBodyRecord[] = []
        append(T.tableStart)
        for (const [rowIndex, row] of table.rows.entries()) {
          const rowStart = offset; append(T.rowStart); const cells: UniverBodyRecord[] = []
          for (let column = 0; column < grid.columnCount; column++) {
            const owner = grid.cellAt(rowIndex, column)!
            if (owner.cell.style.univerBodyOmittedCell) continue
            const cellStart = offset; append(T.cellStart)
            if (owner.row === rowIndex && owner.column === column) {
              writeBlocks(owner.cell.blocks, depth + 1, quote); cells.push(documentTableCellToBody(owner.cell))
            } else {
              const continuation = documentTableContinuationEntries(owner.cell).find(entry => entry.offset === rowIndex - owner.row)
              if (owner.column === column && continuation) writeBlocks(continuation.content.blocks, depth + 1, quote)
              else append(T.paragraph)
              cells.push({ ...cloneModel(bodyRecord(continuation?.content.source?.univerBodyCellSource)), rowSpan: 0, columnSpan: 0,
                officeStudio: { version: 1, ...(continuation ? { continuationSource: documentBodySourceMetadata(continuation.content.source, '续格', ['univerBodyCellSource']) } : {}) } })
            }
            append(T.cellEnd)
            if (owner.row === rowIndex && owner.column === column) structures.cells.push({ tableBlock: block, rowId: row.id, cellId: owner.cell.id,
              blocks: owner.cell.blocks, start: cellStart, end: offset })
          }
          rows.push(documentTableRowToBody(table, rowIndex, cells)); append(T.rowEnd)
          structures.rows.push({ tableBlock: block, rowId: row.id, rowIndex, start: rowStart, end: offset })
        }
        append(T.tableEnd)
        tables.push({ ...source, startIndex: start, endIndex: offset, tableId: block.id, officeStudio: documentBodyBlockSource(block) }); tableSource[block.id] = documentTableToBody(table, block.id, rows)
        writtenTableIds.add(block.id)
        structures.blocks.push({ block, container: blocks, start, end: offset })
        continue
      }
      const image = imageForBlock(block)
      if (image) {
        const source = bodyRecord(block.metadata.univerBodyCustomBlock), drawing = cloneModel(bodyRecord(block.metadata.univerBodyDrawingSource)), start = offset
        const anchor = blocks[blocks.indexOf(block) - 1], extension = documentBodyBlockSource(block)
        if (anchor && documentBodyProjectedPictureInput(anchor, blocks) === block) extension.metadata = {
          ...bodyRecord(extension.metadata), pictureInputAnchorSource: documentBodyPictureInputSource(anchor) }
        append(T.customBlock); customBlocks.push({ ...source, startIndex: start, blockId: block.id, officeStudio: extension })
        drawing.drawingId = block.id; drawing.drawingType = 0
        registerDrawing(block.id, documentImageToDrawing(image, drawing), source.blockId); structures.blocks.push({ block, container: blocks, start, end: offset }); continue
      }
      if (block.type === 'sectionBreak') {
        if (depth > 0) throw new RangeError('单元格中不能导出独立分节。')
        const start = offset; appendSection(true); structures.blocks.push({ block, container: blocks, start, end: offset })
        currentSection = documentSectionSetupToBody(block.sectionBreak?.setup, block.metadata.univerBodySectionSource)
        currentSection.sectionId = block.id
        currentSection.sectionType = DOCUMENT_BODY_SECTION_TYPES.indexOf(block.sectionBreak?.type ?? 'unspecified')
        continue
      }
      if (block.type === 'pageBreak' || block.type === 'columnBreak') {
        const start = offset
        if (block.type === 'pageBreak' && bodyRecord(block.metadata.univerBodyBreakSource).soft === true) renderedPageBreaks.push(offset)
        append(block.type === 'pageBreak' ? T.pageBreak : T.columnBreak); structures.blocks.push({ block, container: blocks, start, end: offset }); continue
      }
      if (block.type === 'divider') { const start = offset; append(T.paragraph); paragraphs.push({ startIndex: offset - 1, paragraphId: block.id,
        officeStudio: documentBodyBlockSource(block), paragraphStyle: { borderBottom: { width: 0.75, color: { rgb: 'auto' }, dashStyle: 1, padding: 0 } } })
        structures.blocks.push({ block, container: blocks, start, end: offset }); continue }
      const start = offset, spans: DocumentBodyTextSpan[] = [], source = bodyRecord(block.metadata.univerBodyParagraphSource)
      let local = 0
      for (const run of block.runs) {
        const from = offset, note = noteIdentity(run), sourceCustom = bodyRecord(run.metadata?.univerBodyCustomBlock), runSource = bodyRecord(run.metadata?.univerBodyRunSource), publicRun = bodyRecord(runSource.textRun)
        const formula = getLatexFormula(run), imageAtom = documentImageAtom(run)
        if (formula) {
          const id = latexFormulaRangeId(run), raw = bodyRecord(run.metadata?.univerBodyLatexRangeSource)
          if (writtenLatex.has(id)) throw new TypeError('公式范围标识重复：' + id)
          writtenLatex.add(id)
          append(formula.latex)
          ranges.push({ id, type: 5, startBlockId: block.id, endBlockId: block.id, start: local, end: local + run.text.length,
            properties: { ...bodyRecord(raw.properties), kind: DOCUMENT_LATEX_RANGE_KIND },
            source: { ...raw, officeStudio: { ...bodyRecord(raw.officeStudio), version: 1, runId: run.id, latexMode: formula.mode } },
            atomic: true, absolute: { start: from, end: offset } })
          latexFormulas[id] = documentBodyLatexConfig(formula.latex, run.metadata?.univerBodyLatexConfigSource ?? latexFormulas[id])
        } else if (imageAtom) {
          append(T.customBlock)
          customBlocks.push({ ...sourceCustom, startIndex: from, blockId: imageAtom.imageId,
            officeStudio: { ...bodyRecord(sourceCustom.officeStudio), version: 1, imageAtom: true } })
          const drawing = documentImageToDrawing(imageAtom.image, cloneModel(bodyRecord(run.metadata?.univerBodyDrawingSource)))
          registerDrawing(imageAtom.imageId, { ...drawing, drawingId: imageAtom.imageId, drawingType: 0 }, sourceCustom.blockId)
        } else if (note) {
          const raw = bodyRecord(run.metadata?.univerBodyNoteRangeSource)
          append('\uFFFC')
          ranges.push({ id: typeof raw.rangeId === 'string' ? raw.rangeId : `note:${run.id}`, type: note.kind === 'footnote' ? 8 : 9,
            startBlockId: block.id, endBlockId: block.id, start: local, end: local + run.text.length,
            properties: { ...bodyRecord(raw.properties), noteId: note.id }, source: raw, atomic: true, absolute: { start: from, end: offset } })
        } else if (Object.keys(sourceCustom).length || run.metadata?.univerBodyDrawingSource) {
          const copied = typeof runSource.runId === 'string' && runSource.runId !== run.id
          const blockId = !copied && typeof sourceCustom.blockId === 'string' ? sourceCustom.blockId : run.id
          append(T.customBlock); (run.metadata?.univerBodyRawCustomBlock === true ? rawCustomBlocks : customBlocks).push({ ...sourceCustom, startIndex: from, blockId })
          const drawing = run.metadata?.univerBodyDrawingSource ?? (typeof sourceCustom.blockId === 'string' ? drawings[sourceCustom.blockId] : undefined)
          if (drawing) registerDrawing(blockId, cloneModel(bodyRecord(drawing)), sourceCustom.blockId)
        } else append(run.text)
        spans.push({ sourceStart: local, sourceEnd: local + run.text.length, targetStart: from, targetEnd: offset, ...(note ? { noteReference: true } : {}) })
        if (offset > from) textRuns.push({ ...publicRun, st: from, ed: offset, ts: documentRunStyleToBody(run), officeStudio: documentBodyRunSource(run) })
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
      if (source.terminated !== false) {
        const paragraph = documentParagraphToBody(block, offset, listInstanceIds.get(block.id))
        if (quote) {
          const style = bodyRecord(paragraph.paragraphStyle), border = bodyRecord(style.borderLeft)
          const generated = { ...border, color: { rgb: quote.lineColor }, width: border.width ?? 2, dashStyle: border.dashStyle ?? 1, padding: border.padding ?? 0 }
          paragraph.paragraphStyle = { ...style, borderLeft: generated }
          const extension = bodyRecord(paragraph.officeStudio)
          paragraph.officeStudio = { ...extension, metadata: { ...bodyRecord(extension.metadata), documentQuoteBorderSource: documentQuoteParagraphBorderSource(block, quote.id, generated) } }
        }
        const mark = block.metadata.univerBodyParagraphMarkTextStyle ?? bodyRecord(paragraph.paragraphStyle).paragraphMarkTextStyle
        if (mark !== undefined) textRuns.push({ st: offset, ed: offset + 1, ts: cloneModel(bodyRecord(mark)) })
        paragraphs.push({ ...paragraph, startIndex: offset, paragraphId: block.id }); append(T.paragraph)
      }
      structures.blocks.push({ block, container: blocks, start, end: offset })
      const list = documentBodyListForBlock(block, listInstanceIds.get(block.id)); if (list) writtenLists.add(list)
      for (const field of documentTextFields(block)) {
        const raw = bodyRecord(field.metadata?.univerBodyFieldSource)
        ranges.push({ id: field.id, type: 1, startBlockId: block.id, endBlockId: field.endBlockId ?? block.id, start: field.startOffset, end: field.endOffset,
          startNoteReferenceCount: field.startNoteReferenceCount, endNoteReferenceCount: field.endNoteReferenceCount,
          properties: { ...bodyRecord(raw.properties), ...bodyRecord(bodyRecord(field.metadata?.source).properties), instruction: field.instruction ?? field.field,
            fieldType: field.field, cachedResult: field.cachedResult, locked: field.locked },
          source: { ...raw, officeStudio: { version: 1, metadata: documentBodySourceMetadata(field.metadata, '字段', ['univerBodyFieldSource']) } } })
      }
    }
  }
  writeBlocks(content.blocks!)
  if (!chunks.length || chunks.at(-1) === T.section) {
    append(T.paragraph); paragraphs.push({ startIndex: offset - 1, paragraphId: documentId('paragraph') })
  }
  appendSection(sections.length > 0)
  const baseStream = chunks.join('')
  const sourceStream = typeof sourceBody.dataStream === 'string' ? sourceBody.dataStream : ''
  const atLocal = (id: string, local: number, affinity: 'start' | 'end' = 'start', rank?: number) => {
    const map = maps.get(id)
    if (!map) throw new RangeError(`文档导出范围的块不存在：${id}`)
    const fallback = documentBodyLocalOffset({ block: map.block, start: 0, end: runText(map.block.runs).length, spans: map.spans }, local, affinity)
    return documentBodyNoteStreamOffset(map.block, map.spans, local, fallback, rank)
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
    const absolute = documentBodyStructuralControlBounds(control, structures)
    if (!absolute && (!maps.has(firstId) || !maps.has(lastId))) throw new RangeError('公开内容控件缺少真实文字或结构范围：' + control.id)
    const properties = controlProperties(control); properties.placement = documentBodyPublicControlPlacement(control, structures)
    ranges.push({ id: control.id, type: 2, startBlockId: firstId, endBlockId: lastId, start: control.blockRange ? 0 : control.startOffset,
      startNoteReferenceCount: control.startNoteReferenceCount, endNoteReferenceCount: control.endNoteReferenceCount,
      end: absolute ? 0 : control.blockRange ? runText(maps.get(lastId)!.block.runs).length : control.endOffset, properties,
      ...(absolute ? { absolute } : {}), source: { ...bodyRecord(control.metadata?.univerBodySdtSource), officeStudio: documentBodyControlSource(control) } })
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
    let hasSeparator = false
    for (const value of Array.isArray(source.boundaryTokens) ? source.boundaryTokens : []) {
      const token = bodyRecord(value)
      if (typeof token.offset !== 'number' || typeof token.text !== 'string' || token.text.length !== 1
        || token.text !== '\uFFFC' && !Object.values(T).includes(token.text as typeof T[keyof typeof T]) || typeof token.sourceIndex !== 'number') throw new TypeError('文档保留的结构标记无效。')
      if (typeof token.sourceRangeId === 'string' && sourceModeledRangeIds.has(token.sourceRangeId)) {
        if (bodyRecord(map.block.metadata.noteSeparator).version !== 1) continue
        retainedSourceRangeIds.add(token.sourceRangeId)
      }
      const position = atLocal(map.block.id, rebasedOffset(before, runText(map.block.runs), token.offset, 'start'))
      markers.push({ offset: position, end: position, text: token.text, order: -1, sourceIndex: token.sourceIndex })
      if (token.text === '\uFFFC') {
        const raw = documentArray(sourceBody.customRanges).find(raw => raw.rangeId === token.sourceRangeId && raw.startIndex === token.sourceIndex)
        if (!raw) throw new RangeError('文档生成标记缺少来源范围。')
        const field = bodyRecord(raw.properties).fieldType
        hasSeparator ||= field === 'FOOTNOTE_SEPARATOR' || field === 'FOOTNOTE_CONTINUATION_SEPARATOR'
      }
      if (token.text === T.customBlock) {
        const raw = [...(Array.isArray(sourceBody.customBlocks) ? sourceBody.customBlocks : []),
          ...(Array.isArray(sourceBody.docxRawCustomBlocks) ? sourceBody.docxRawCustomBlocks : [])].map(bodyRecord).find(raw => raw.startIndex === token.sourceIndex)
        if (!raw || typeof raw.blockId !== 'string') throw new RangeError('文档保留的 customBlock 缺少来源定义。')
        hasSeparator ||= raw.noteType === 'separator' || raw.noteType === 'continuationSeparator'
        const descriptor = { ...raw, startIndex: position, blockId: raw.blockId,
          ...(options.note && raw.noteType === options.note.kind ? { noteId: options.note.id } : {}) }
        customBlocks.push(descriptor); sourceCustomTokens.set(descriptor, token.sourceIndex)
        if (drawings[raw.blockId]) writtenDrawingIds.add(raw.blockId)
      }
    }
    const separator = bodyRecord(map.block.metadata.noteSeparator)
    if (separator.version === 1 && !hasSeparator) {
      const raw = bodyRecord(map.block.metadata.univerNoteSeparatorRangeSource), id = typeof raw.rangeId === 'string' ? raw.rangeId : `${map.block.id}:note-separator`
      if (generatedIds.has(id)) throw new TypeError('注释分隔范围身份重复：' + id)
      generatedIds.add(id)
      const tokenId = -1 - markers.length, at = map.start
      markers.push({ offset: at, end: at, text: '\uFFFC', order: -1, sourceIndex: tokenId })
      const range: UniverDocumentCustomRange = { ...raw, rangeId: id, rangeType: 1, startIndex: at, endIndex: at, wholeEntity: true,
        properties: { ...bodyRecord(raw.properties), fieldType: separator.widthFraction === 1 ? 'FOOTNOTE_CONTINUATION_SEPARATOR' : 'FOOTNOTE_SEPARATOR' } }
      customRanges.push(range); sourceRangeTokens.set(range, tokenId)
    }
  }
  for (const raw of Array.isArray(sourceBody.customRanges) ? sourceBody.customRanges : []) {
    const range = bodyRecord(raw)
    if (typeof range.rangeId !== 'string' || generatedIds.has(range.rangeId)) continue
    if (sourceModeledRangeIds.has(range.rangeId) && !retainedSourceRangeIds.has(range.rangeId)) continue
    if (typeof range.startIndex !== 'number' || typeof range.endIndex !== 'number' || typeof range.rangeType !== 'number') continue
    const sourceToken = markers.find(marker => marker.text === '\uFFFC' && marker.sourceIndex === range.startIndex && range.endIndex === range.startIndex)
    if (sourceToken) {
      const publicRange: UniverDocumentCustomRange = { ...range, rangeId: range.rangeId, rangeType: range.rangeType,
        startIndex: sourceToken.offset, endIndex: sourceToken.offset,
        ...(options.note && (range.rangeType === 8 || range.rangeType === 9)
          ? { properties: { ...bodyRecord(range.properties), noteId: options.note.id } } : {}) }
      customRanges.push(publicRange); sourceRangeTokens.set(publicRange, range.startIndex); continue
    }
    const start = atSource(range.startIndex, 'start'), end = atSource(range.endIndex + 1, 'end')
    const index = customRanges.length
    customRanges.push({ ...range, rangeId: range.rangeId, rangeType: range.rangeType, startIndex: start,
      endIndex: Math.max(start, streamRangeEnd(sourceStream, range.endIndex, end)) })
    if (sourceStream[range.startIndex] === T.rangeStart) markers.push({ offset: start, end, text: T.rangeStart, order: index })
    if (sourceStream[range.endIndex] === T.rangeEnd) markers.push({ offset: end, end: start, text: T.rangeEnd, order: index })
  }
  for (const range of ranges) {
    const start = range.absolute?.start ?? atLocal(range.startBlockId, range.start, 'start', range.startNoteReferenceCount), end = range.absolute?.end ?? atLocal(range.endBlockId, range.end, 'end', range.endNoteReferenceCount)
    if (end < start) throw new RangeError('文档导出自定义范围方向无效。')
    const index = customRanges.length
    customRanges.push({ ...range.source, startIndex: start, endIndex: range.atomic ? Math.max(start, end - 1) : end,
      rangeId: range.id, rangeType: range.type, properties: range.properties, ...(range.atomic ? { wholeEntity: true } : {}) })
    if (!range.atomic) markers.push({ offset: start, end, text: T.rangeStart, order: index }, { offset: end, end: start, text: T.rangeEnd, order: index })
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
  for (const item of [...blockRanges, ...columnGroups]) { item.startIndex = afterMarker(item.startIndex); item.endIndex = beforeMarker(item.endIndex + 1) - 1 }
  for (const section of sections) section.startIndex = afterMarker(Number(section.startIndex))
  for (const [index, position] of renderedPageBreaks.entries()) renderedPageBreaks[index] = afterMarker(position)
  for (const [index, range] of customRanges.entries()) {
    const token = sourceRangeTokens.get(range)
    if (token !== undefined) { range.startIndex = sourceTokenPosition(token)!; range.endIndex = range.startIndex; continue }
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
          targetStart: localAt(span, start, 'start'), targetEnd: localAt(span, end, 'end'), ...(span.noteReference ? { noteReference: true } : {}) })
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
  body.blockRanges = blockRanges.length ? blockRanges : undefined
  body.columnGroups = columnGroups.length ? columnGroups : undefined
  const metadata = options.includeMetadata === false ? {} : documentBodySourceMetadata(content.metadata, '正文',
    ['univerBodySource', 'univerBodyFirstSectionSource', 'univerDocumentSource', 'documentEditingBlockHost', 'documentEditingBlockSourceId', 'documentEditingSectionId', 'documentEditingSectionIds'])
  body.officeStudio = Object.keys(metadata).length ? { version: 1, metadata } : undefined
  for (const key of ['customDecorations', 'docxRawBlocks'] as const) if (Array.isArray(sourceBody[key])) body[key] = sourceBody[key].map(value => {
    const raw = cloneModel(bodyRecord(value))
    if (typeof raw.startIndex === 'number') raw.startIndex = sourceTokenPosition(raw.startIndex) ?? afterMarker(atSource(raw.startIndex, 'start'))
    if (typeof raw.endIndex === 'number') raw.endIndex = sourceTokenPosition(raw.endIndex) ?? beforeMarker(atSource(raw.endIndex + 1, 'end')) - 1
    return raw
  })
  if (Array.isArray(sourceBody.docxExportExcludedRanges)) body.docxExportExcludedRanges = sourceBody.docxExportExcludedRanges.map(value => {
    const raw = cloneModel(bodyRecord(value))
    if (typeof raw.start === 'number') raw.start = sourceTokenPosition(raw.start) ?? afterMarker(atSource(raw.start, 'start'))
    if (typeof raw.end === 'number') raw.end = sourceTokenPosition(raw.end) ?? beforeMarker(atSource(raw.end, 'end'))
    return raw
  })
  // Remove old custom-block resources after deletion/copy, while retaining source
  // resources that never belonged to a modeled custom block (e.g. group parts).
  for (const id of sourceTableIds) if (!writtenTableIds.has(id)) delete tableSource[id]
  for (const id of sourceDrawingIds) if (!writtenDrawingIds.has(id)) delete drawings[id]
  for (const id of sourceListIds) delete lists[id]
  for (const raw of documentArray(sourceBody.blockRanges)) {
    if (raw.blockType === 'callout' && typeof raw.blockId === 'string' && !writtenCallouts.has(raw.blockId)) delete callouts[raw.blockId]
    if (raw.blockType === 'code' && typeof raw.blockId === 'string' && !writtenCodes.has(raw.blockId)) delete codes[raw.blockId]
  }
  for (const raw of documentArray(sourceBody.customRanges)) if (isDocumentBodyLatexRange(raw)
    && typeof raw.rangeId === 'string' && !writtenLatex.has(raw.rangeId)) delete latexFormulas[raw.rangeId]
  Object.assign(lists, writtenLists.finish())
  const drawingOrder = Array.isArray(sourceResources.drawingsOrder) ? sourceResources.drawingsOrder
    .filter((id): id is string => typeof id === 'string').map(id => drawings[id] ? id : drawingAliases.get(id) ?? id).filter(id => drawings[id]) : []
  const drawingsOrder = [...new Set([...drawingOrder, ...Object.keys(drawings)])]
  const imageOrder = drawingsOrder.filter(id => writtenDrawingIds.has(id) && drawings[id]?.drawingType === 0)
    .sort((a, b) => {
      const left = bodyRecord(drawings[a]?.officeStudioImage).layer, right = bodyRecord(drawings[b]?.officeStudioImage).layer
      return (typeof left === 'number' ? left : drawingOrder.indexOf(a)) - (typeof right === 'number' ? right : drawingOrder.indexOf(b))
    })
  const imageIds = new Set(imageOrder)
  let imageIndex = 0
  for (const [index, id] of drawingsOrder.entries()) if (imageIds.has(id)) drawingsOrder[index] = imageOrder[imageIndex++]!
  return { body, ...(Object.keys(tableSource).length ? { tableSource } : {}), ...(Object.keys(drawings).length ? { drawings, drawingsOrder } : {}),
    ...(Object.keys(lists).length ? { lists } : {}), ...(Object.keys(callouts).length ? { callouts } : {}), ...(Object.keys(codes).length ? { codes } : {}),
    ...(Object.keys(latexFormulas).length ? { latexFormulas } : {}) }
}

function streamRangeEnd(stream: string, sourceEnd: number, offset: number): number {
  return stream[sourceEnd] === T.rangeEnd ? offset : offset - 1
}

function documentArray(value: unknown): UniverBodyRecord[] {
  return Array.isArray(value) ? value.map(bodyRecord) : []
}
