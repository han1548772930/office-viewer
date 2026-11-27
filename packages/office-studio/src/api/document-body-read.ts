import type { DocumentBlock, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { createParagraph, createRun, documentId, runText, tableForBlock } from '../features/document/model/document-model'
import { isDocumentTextBoundary } from '../features/document/model/document-story'
import { bodyNumber, bodyRecord, documentRunStyleFromBody, documentRunStyleSource, type UniverBodyRecord } from './document-body-style'
import { documentBodyNamedStyle, documentBodyParagraphMarks, documentParagraphFromBody } from './document-body-paragraph'
import { documentSectionSetupFromBody, documentSectionTypeFromBody } from './document-body-section'
import { documentTableFromBody } from './document-body-table'
import { documentCalloutFromBody, documentCodeFromBody, documentColumnGroupFromBody } from './document-body-structures'
import { applyDocumentBodyBlockSource, documentBodyOfficeSource, documentBodyRunMetadata } from './document-body-source'
import { readDocumentBodyLatexRanges } from './document-body-latex'
import { DOCUMENT_LATEX_ATOM } from '../features/document/latex/latex-formula'
import { documentImageFromDrawing } from './document-image-body'
import { DOCUMENT_IMAGE_ATOM, documentImageAtom, setDocumentImageAtom } from '../features/document/images/document-image-atom'
import { DOCUMENT_BODY_TOKENS as T, documentBodyArray, documentBodyIdentity, documentBodyOffset,
  type DocumentBodyDecodeOptions, type DocumentBodyParagraphPosition, type DocumentBodyStructureIndex, type DocumentBodyTextSpan, type UniverDocumentBodyResources } from './document-body-types'

interface TextPart { text: string; start: number; end: number; source?: UniverBodyRecord; metadata?: UniverBodyRecord }
interface ReadBodyResult { blocks: DocumentBlock[]; positions: DocumentBodyParagraphPosition[]; structures: DocumentBodyStructureIndex; diagnostics: string[]; initialSection?: UniverBodyRecord }
const containerEnds = new Set<string>([T.cellEnd, T.rowEnd, T.tableEnd, T.blockEnd, T.columnEnd, T.groupEnd])
const nonprinting = new Set<string>([T.rangeStart, T.rangeEnd])
const flowTokens = new Set<string>([T.paragraph, T.section, T.tableStart, T.pageBreak, T.columnBreak, T.customBlock, T.blockStart, T.groupStart, T.columnStart])

/** Stream parsing belongs at the DTO boundary; the result is the ordinary editor block tree. */
export function readDocumentBody(value: unknown, resources: UniverDocumentBodyResources, options: DocumentBodyDecodeOptions): ReadBodyResult {
  const body = bodyRecord(value)
  if (typeof body.dataStream !== 'string') throw new TypeError('文档 body.dataStream 必须为字符串。')
  const stream = body.dataStream
  const latexRanges = readDocumentBodyLatexRanges(body, resources)
  const latexRangesOrdered = [...latexRanges.values()]
  const latexAt = (offset: number) => {
    let low = 0, high = latexRangesOrdered.length
    while (low < high) { const middle = (low + high) >>> 1; if (latexRangesOrdered[middle]!.start <= offset) low = middle + 1; else high = middle }
    const range = latexRangesOrdered[low - 1]
    return range && range.end > offset ? range : undefined
  }
  if (stream.length * 2 > 32 * 1024 * 1024) throw new RangeError('文档公开内容超过单次文本预算。')
  const paragraphs = new Map<number, UniverBodyRecord>(), tables = new Map<number, UniverBodyRecord>(), customBlocks = new Map<number, UniverBodyRecord>()
  const blockRanges = new Map<number, UniverBodyRecord>(), columnGroups = new Map<number, UniverBodyRecord>()
  const rawCustomOffsets = new Set<number>(), sections = new Map<number, UniverBodyRecord>(), softBreaks = new Set<number>()
  const generatedRanges = new Map<number, { rangeId: string; separator?: 'separator' | 'continuationSeparator' }>()
  const noteRanges = new Map<number, { kind: 'footnote' | 'endnote'; id: string; source: UniverBodyRecord }>()
  const checkPosition = (value: unknown, label: string) => {
    const offset = documentBodyOffset(value, stream.length, label)
    if (!isDocumentTextBoundary(stream, offset)) throw new RangeError(`${label} 不能截断 UTF-16 字符。`)
    return offset
  }
  const mapEntries = (input: unknown, map: Map<number, UniverBodyRecord>, token: string, label: string) => {
    for (const raw of documentBodyArray(input, label)) {
      const offset = checkPosition(raw.startIndex, label)
      if (latexAt(offset)) throw new RangeError(`${label} 不能位于公式实体内部。`)
      if (stream[offset] !== token || map.has(offset)) throw new RangeError(`${label} 与内容流标记不匹配或位置重复。`)
      map.set(offset, raw)
    }
  }
  mapEntries(body.paragraphs, paragraphs, T.paragraph, 'paragraphs')
  mapEntries(body.tables, tables, T.tableStart, 'tables')
  mapEntries(body.customBlocks, customBlocks, T.customBlock, 'customBlocks')
  mapEntries(body.docxRawCustomBlocks, customBlocks, T.customBlock, 'docxRawCustomBlocks')
  for (const raw of documentBodyArray(body.docxRawCustomBlocks, 'docxRawCustomBlocks')) rawCustomOffsets.add(Number(raw.startIndex))
  mapEntries(body.sectionBreaks, sections, T.section, 'sectionBreaks')
  mapEntries(body.blockRanges, blockRanges, T.blockStart, 'blockRanges')
  mapEntries(body.columnGroups, columnGroups, T.groupStart, 'columnGroups')
  for (const [start, raw] of [...blockRanges, ...columnGroups]) {
    const end = checkPosition(raw.endIndex, '高级块结束边界'), closing = blockRanges.get(start) === raw ? T.blockEnd : T.groupEnd
    if (end <= start || stream[end] !== closing) throw new RangeError('高级块的 inclusive 结束位置与内容流标记不一致。')
  }
  // The fixed upstream default separator is a FIELD over U+FFFC, rather
  // than a customBlock. Its marker belongs to layout, never editable text.
  for (const raw of documentBodyArray(body.customRanges, 'customRanges')) {
    const properties = bodyRecord(raw.properties), field = properties.fieldType
    const separator = raw.rangeType === 1 && field === 'FOOTNOTE_SEPARATOR' ? 'separator'
      : raw.rangeType === 1 && field === 'FOOTNOTE_CONTINUATION_SEPARATOR' ? 'continuationSeparator' : undefined
    const kind = raw.rangeType === 8 ? 'footnote' : raw.rangeType === 9 ? 'endnote' : undefined
    const ownReference = kind !== undefined && options.note?.kind === kind
      && (options.note.id === undefined || options.note.id === properties.noteId)
    if (!separator && !ownReference) {
      if (kind !== undefined) {
        const at = checkPosition(raw.startIndex, '注释引用')
        if (raw.endIndex === at && stream[at] === '\uFFFC') {
          if (noteRanges.has(at)) throw new RangeError('同一公开原子不能对应多个注释引用。')
          noteRanges.set(at, { kind, id: documentBodyIdentity(properties.noteId, '注释'), source: raw })
        }
      }
      continue
    }
    const at = checkPosition(raw.startIndex, '生成标记')
    if (raw.endIndex !== at || stream[at] !== '\uFFFC') throw new RangeError('注释生成标记必须对应一个 U+FFFC。')
    if (generatedRanges.has(at)) throw new RangeError('注释生成标记位置重复。')
    generatedRanges.set(at, { rangeId: documentBodyIdentity(raw.rangeId, '生成标记'), ...(separator ? { separator } : {}) })
  }
  const sectionSources: UniverBodyRecord[] = [], sectionIds = new Set<string>()
  for (let at = stream.indexOf(T.section); at >= 0; at = stream.indexOf(T.section, at + 1)) {
    if (latexAt(at)) continue
    const raw = sections.get(at) ?? {}, sectionId = documentBodyIdentity(raw.sectionId ?? documentId('section'), '分节')
    if (sectionIds.has(sectionId)) throw new TypeError(`分节标识重复：${sectionId}`)
    sectionIds.add(sectionId); documentSectionTypeFromBody(raw.sectionType)
    sectionSources.push({ ...cloneModel(raw), startIndex: at, sectionId })
  }
  const nextSections = new Map(sectionSources.slice(0, -1).map((section, index) => [Number(section.startIndex), sectionSources[index + 1]!]))
  if (body.renderedPageBreaks !== undefined) {
    if (!Array.isArray(body.renderedPageBreaks)) throw new TypeError('软分页位置必须为数组。')
    for (const value of body.renderedPageBreaks) {
      const offset = checkPosition(value, 'renderedPageBreaks')
      if (stream[offset] !== T.pageBreak || softBreaks.has(offset)) throw new RangeError('软分页位置与内容流不匹配或位置重复。')
      softBreaks.add(offset)
    }
  }
  const textRuns = documentBodyArray(body.textRuns, 'textRuns').map(raw => {
    const start = checkPosition(raw.st, 'textRuns.st'), end = checkPosition(raw.ed, 'textRuns.ed')
    if (end < start) throw new RangeError('文字运行结束位置早于开始位置。')
    return { raw, start, end }
  })
  const boundaries = [...new Set(textRuns.flatMap(run => [run.start, run.end]))].sort((a, b) => a - b)
  const positions: DocumentBodyParagraphPosition[] = [], diagnostics: string[] = [], identities = new Set<string>()
  const structures: DocumentBodyStructureIndex = { blocks: [], rows: [], cells: [] }
  const runSourceIds = new Map<string, string>(), drawingIds = new Set<string>()
  let cursor = 0, count = 0
  const reserve = (block: DocumentBlock) => {
    if (++count > 250_000) throw new RangeError('文档公开内容超过块预算。')
    if (identities.has(block.id)) throw new TypeError(`文档块标识重复：${block.id}`)
    identities.add(block.id); return block
  }
  const paragraphAt = (position: number) => {
    let terminal = stream.indexOf(T.paragraph, position)
    while (terminal >= 0 && latexAt(terminal)) terminal = stream.indexOf(T.paragraph, latexAt(terminal)!.end)
    return terminal < 0 ? {} : paragraphs.get(terminal) ?? {}
  }
  const readBlocks = (until?: string, depth = 0): DocumentBlock[] => {
    if (depth > 64) throw new RangeError('文档结构嵌套超过预算。')
    const blocks: DocumentBlock[] = []
    let start = cursor, parts: TextPart[] = [], spans: DocumentBodyTextSpan[] = [], length = 0
    let pendingTokens: Array<{ offset: number; text: string; sourceIndex: number; sourceRangeId?: string }> = []
    let pendingSeparator: 'separator' | 'continuationSeparator' | undefined
    const flush = (end: number, terminal: boolean, force = false) => {
      if (!force && !parts.length && !pendingTokens.length) { start = cursor; return }
      const raw = terminal ? paragraphs.get(end) ?? {} : paragraphAt(start), marks = documentBodyParagraphMarks(raw, options)
      const runs: InlineRun[] = parts.map(part => {
        const character = documentBodyNamedStyle(part.source?.sId, options.styles)
        const inherited = { ...marks, ...bodyRecord(character.textStyle) }, explicit = bodyRecord(part.source?.ts)
        const style = documentRunStyleFromBody({ ...inherited, ...explicit })
        const run = createRun(part.text, { ...options.defaults.text, ...style })
        const identity = part.source && documentBodyOfficeSource(part.source, '运行').runId
        if (identity !== undefined) {
          const id = documentBodyIdentity(identity, '运行'), key = `${part.source!.st}:${part.source!.ed}`
          const previous = runSourceIds.get(id)
          if (previous !== undefined && previous !== key) throw new TypeError('公开运行身份被不同文字范围重复使用：' + id)
          if (previous === undefined) run.id = id
          runSourceIds.set(id, key)
        }
        run.metadata = { ...documentBodyRunMetadata(part.source), ...run.metadata, ...documentRunStyleSource(run, explicit, inherited),
          ...part.metadata, univerBodyRunSource: { version: 1, runId: run.id, start: part.start, end: part.end,
          ...(part.source ? { textRun: cloneModel(part.source) } : {}) } }
        if (part.metadata?.latexFormula) {
          const range = bodyRecord(run.metadata.univerBodyLatexRangeSource), extension = bodyRecord(range.officeStudio)
          if (extension.runId !== undefined) run.id = documentBodyIdentity(extension.runId, '公式运行')
          run.metadata.univerBodyLatexRangeSource = { ...range, runId: run.id }
          run.metadata.univerBodyRunSource = { ...bodyRecord(run.metadata.univerBodyRunSource), runId: run.id }
        }
        const imageAtom = documentImageAtom(run)
        if (imageAtom) setDocumentImageAtom(run, typeof bodyRecord(part.metadata?.imageAtom).imageId === 'string'
          ? String(bodyRecord(part.metadata?.imageAtom).imageId) : imageAtom.imageId, imageAtom.image)
        return run
      })
      const paragraph = documentParagraphFromBody(terminal ? raw : { ...raw, paragraphId: undefined }, runs, options, resources.lists)
      if (terminal) {
        const markRuns = textRuns.filter(run => run.start <= end && run.end > end)
        if (markRuns.length) paragraph.metadata.univerBodyParagraphMarkTextStyle = Object.assign({}, ...markRuns.map(run => bodyRecord(run.raw.ts)))
      }
      if (pendingSeparator) paragraph.metadata.noteSeparator = { version: 1, widthFraction: pendingSeparator === 'separator' ? 1 / 3 : 1,
        fieldType: pendingSeparator === 'separator' ? 'FOOTNOTE_SEPARATOR' : 'FOOTNOTE_CONTINUATION_SEPARATOR' }
      const source = bodyRecord(paragraph.metadata.univerBodyParagraphSource)
      paragraph.metadata.univerBodyParagraphSource = { ...source, start, end, spans: cloneModel(spans), terminated: terminal, baselineText: runText(paragraph.runs),
        ...(pendingTokens.length ? { boundaryTokens: pendingTokens } : {}) }
      reserve(paragraph); blocks.push(paragraph); positions.push({ block: paragraph, start, end, spans })
      structures.blocks.push({ block: paragraph, container: blocks, start, end: terminal ? end + 1 : end })
      parts = []; spans = []; length = 0; pendingTokens = []; pendingSeparator = undefined; start = cursor
    }
    const append = (text: string, from: number, to: number, metadata?: UniverBodyRecord) => {
      if (metadata) {
        const active = textRuns.filter(run => run.start <= from && run.end >= to)
        const source = active.length ? { ...active.at(-1)!.raw, ts: Object.assign({}, ...active.map(run => bodyRecord(run.raw.ts))) } : undefined
        parts.push({ text, start: from, end: to, metadata, source }); spans.push({ sourceStart: from, sourceEnd: to, targetStart: length, targetEnd: length + text.length,
          ...(metadata.noteReference || metadata.note ? { noteReference: true } : {}) }); length += text.length; return
      }
      const cuts = [from, ...boundaries.filter(boundary => boundary > from && boundary < to), to]
      for (let index = 0; index < cuts.length - 1; index++) {
        const a = cuts[index]!, b = cuts[index + 1]!, active = textRuns.filter(run => run.start <= a && run.end >= b)
        const source = active.length ? { ...active.at(-1)!.raw, ts: Object.assign({}, ...active.map(run => bodyRecord(run.raw.ts))) } : undefined
        const slice = text.slice(a - from, b - from)
        parts.push({ text: slice, start: a, end: b, source }); spans.push({ sourceStart: a, sourceEnd: b, targetStart: length, targetEnd: length + slice.length }); length += slice.length
      }
    }
    while (cursor < stream.length) {
      const token = stream[cursor]!
      const formula = latexRanges.get(cursor)
      if (formula) {
        append(DOCUMENT_LATEX_ATOM, cursor, formula.end, { latexFormula: formula.formula,
          univerBodyLatexRangeSource: formula.source, univerBodyLatexConfigSource: formula.config })
        cursor = formula.end; continue
      }
      if (token === until) {
        flush(cursor, false)
        if (!blocks.length) {
          const block = reserve(createParagraph('', options.defaults.text)); block.style = { ...options.defaults.paragraph }
          blocks.push(block); positions.push({ block, start: cursor, end: cursor, spans: [] })
          structures.blocks.push({ block, container: blocks, start: cursor, end: cursor })
        }
        cursor++; return blocks
      }
      const generated = generatedRanges.get(cursor)
      if (generated) {
        spans.push({ sourceStart: cursor, sourceEnd: cursor + 1, targetStart: length, targetEnd: length })
        pendingTokens.push({ offset: length, text: token, sourceIndex: cursor, sourceRangeId: generated.rangeId })
        pendingSeparator ??= generated.separator
        cursor++; continue
      }
      const reference = noteRanges.get(cursor)
      if (reference) {
        const from = cursor++
        append('\uFFFC', from, cursor, { noteReference: { kind: reference.kind, noteId: reference.id }, univerBodyNoteRangeSource: cloneModel(reference.source) })
        continue
      }
      if (containerEnds.has(token)) throw new RangeError('文档结构标记未正确闭合。')
      if (token === T.paragraph) { const end = cursor++; flush(end, true, true); start = cursor; continue }
      if (token === T.section) {
        flush(cursor, false); const offset = cursor++
        if (cursor < stream.length) {
          if (!options.allowSections || depth > 0) throw new RangeError('注释或单元格内容不能包含独立分节。')
          const section = nextSections.get(offset) ?? { sectionId: documentId('section') }
          const block = createParagraph(); block.runs = []; block.type = 'sectionBreak'
          block.sectionBreak = { type: documentSectionTypeFromBody(section.sectionType), setup: documentSectionSetupFromBody(section) }
          block.id = documentBodyIdentity(section.sectionId, '分节')
          block.metadata.univerBodySectionSource = cloneModel(section); blocks.push(reserve(block))
          structures.blocks.push({ block, container: blocks, start: offset, end: cursor })
        }
        start = cursor; continue
      }
      if (token === T.tableStart) {
        flush(cursor, false); const offset = cursor++, descriptor = tables.get(offset)
        if (!descriptor) throw new RangeError('内容流中的表格缺少 tables 描述。')
        const id = documentBodyIdentity(descriptor.tableId, '表格'), source = resources.tableSource?.[id]
        if (!source) throw new RangeError(`表格 ${id} 缺少 tableSource。`)
        const rows: DocumentBlock[][][] = [], bounds: Array<{ start: number; end: number; cells: Array<{ start: number; end: number }> }> = []
        // Row and cell SDTs may place custom-range sentinels between table
        // tokens. Those delimiters are coordinates, not additional rows.
        const skipRangeMarkers = () => { while (stream[cursor] === T.rangeStart || stream[cursor] === T.rangeEnd) cursor++ }
        skipRangeMarkers()
        while (stream[cursor] === T.rowStart) {
          const rowStart = cursor++; const cells: DocumentBlock[][] = [], cellBounds: Array<{ start: number; end: number }> = []
          skipRangeMarkers()
          while (stream[cursor] === T.cellStart) {
            const cellStart = cursor++; cells.push(readBlocks(T.cellEnd, depth + 1)); cellBounds.push({ start: cellStart, end: cursor }); skipRangeMarkers()
          }
          if (stream[cursor++] !== T.rowEnd) throw new RangeError('表格行缺少结束标记。')
          rows.push(cells); bounds.push({ start: rowStart, end: cursor, cells: cellBounds }); skipRangeMarkers()
        }
        if (stream[cursor++] !== T.tableEnd || descriptor.endIndex !== cursor) throw new RangeError('表格结束边界与 tables 描述不匹配。')
        const block = createParagraph(); block.id = id; block.type = 'table'; block.runs = []
        block.metadata = { table: documentTableFromBody(source, rows, options.defaults), univerBodyTableSource: cloneModel(descriptor) }
        applyDocumentBodyBlockSource(block, descriptor)
        blocks.push(reserve(block)); structures.blocks.push({ block, container: blocks, start: offset, end: cursor })
        const table = tableForBlock(block)!
        for (const [rowIndex, bound] of bounds.entries()) {
          const row = table.rows[rowIndex]!
          structures.rows.push({ tableBlock: block, rowId: row.id, rowIndex, start: bound.start, end: bound.end })
          for (const [cellIndex, cellBound] of bound.cells.entries()) {
            const cell = row.cells.find(cell => cell.blocks === rows[rowIndex]![cellIndex])
            if (cell) structures.cells.push({ tableBlock: block, rowId: row.id, cellId: cell.id, blocks: cell.blocks, start: cellBound.start, end: cellBound.end })
          }
        }
        start = cursor; continue
      }
      if (token === T.blockStart) {
        flush(cursor, false); const from = cursor++, descriptor = blockRanges.get(from)
        if (!descriptor || !['quote', 'callout', 'code'].includes(String(descriptor.blockType))) throw new RangeError('高级块缺少有效的 blockRanges 描述。')
        const id = documentBodyIdentity(descriptor.blockId, '高级块'), children = readBlocks(T.blockEnd, depth + 1)
        if (descriptor.endIndex !== cursor - 1) throw new RangeError('高级块结束标记未与 blockRanges 正确匹配。')
        const block = createParagraph(); block.id = id; block.type = descriptor.blockType as 'quote' | 'callout' | 'code'; block.runs = []; block.children = children
        block.metadata = { univerBodyBlockRangeSource: cloneModel(descriptor) }
        if (block.type === 'callout') {
          block.callout = documentCalloutFromBody(resources.callouts?.[id]); block.metadata.univerBodyCalloutSource = cloneModel(resources.callouts?.[id] ?? {})
        } else if (block.type === 'code') {
          block.code = documentCodeFromBody(resources.codes?.[id]); block.metadata.univerBodyCodeSource = cloneModel(resources.codes?.[id] ?? {})
        }
        applyDocumentBodyBlockSource(block, descriptor)
        blocks.push(reserve(block)); structures.blocks.push({ block, container: blocks, start: from, end: cursor }); start = cursor; continue
      }
      if (token === T.groupStart) {
        flush(cursor, false); const from = cursor++, descriptor = columnGroups.get(from)
        if (!descriptor) throw new RangeError('多栏组缺少 columnGroups 描述。')
        const children: DocumentBlock[][] = [], skipRanges = () => { while (stream[cursor] === T.rangeStart || stream[cursor] === T.rangeEnd) cursor++ }
        skipRanges()
        while (stream[cursor] === T.columnStart) {
          if (children.length >= 12) throw new RangeError('多栏组超过 12 栏预算。')
          cursor++; children.push(readBlocks(T.columnEnd, depth + 1)); skipRanges()
        }
        if (stream[cursor++] !== T.groupEnd || descriptor.endIndex !== cursor - 1) throw new RangeError('多栏组结束边界与 columnGroups 不匹配。')
        const block = createParagraph(); block.id = documentBodyIdentity(descriptor.columnGroupId, '多栏组'); block.type = 'columns'; block.runs = []
        block.columns = documentColumnGroupFromBody(descriptor, children, options)
        block.metadata = { univerBodyColumnGroupSource: cloneModel(descriptor), univerBodyColumnGroupBaseline: { gap: block.columns.gap,
          horizontalPadding: block.columns.horizontalPadding, minHeight: block.columns.minHeight, fontSize: options.defaults.text.fontSize } }
        applyDocumentBodyBlockSource(block, descriptor)
        blocks.push(reserve(block)); structures.blocks.push({ block, container: blocks, start: from, end: cursor }); start = cursor; continue
      }
      if (token === T.columnStart) throw new RangeError('栏开始标记必须属于一个多栏组。')
      if (token === T.pageBreak || token === T.columnBreak) {
        if (!options.allowSections) throw new RangeError('注释内容不能包含独立分页或分栏符。')
        flush(cursor, false); const offset = cursor++, block = createParagraph(); block.runs = []; block.type = token === T.pageBreak ? 'pageBreak' : 'columnBreak'
        block.metadata.univerBodyBreakSource = { startIndex: offset, soft: softBreaks.has(offset) }
        blocks.push(reserve(block)); structures.blocks.push({ block, container: blocks, start: offset, end: cursor }); start = cursor; continue
      }
      if (token === T.customBlock) {
        const from = cursor++, raw = customBlocks.get(from)
        if (!raw) throw new RangeError('内容流的 customBlock 缺少定义。')
        const id = documentBodyIdentity(raw.blockId, 'customBlock'), noteType = raw.noteType
        if (rawCustomOffsets.has(from)) {
          append('\uFFFC', from, cursor, { univerBodyCustomBlock: cloneModel(raw), univerBodyRawCustomBlock: true })
          diagnostics.push(`DOCX 原始运行 ${id} 已保留，当前 body 适配尚未解释其绘制行为。`); continue
        }
        if (noteType === 'footnote' || noteType === 'endnote') {
          const noteId = documentBodyIdentity(raw.noteId, '注释')
          if (options.note?.kind === noteType && (options.note.id === undefined || options.note.id === noteId)) {
            spans.push({ sourceStart: from, sourceEnd: cursor, targetStart: length, targetEnd: length })
            pendingTokens.push({ offset: length, text: token, sourceIndex: from }); continue
          }
          append('\uFFFC', from, cursor, { noteReference: { kind: noteType, noteId }, univerBodyCustomBlock: cloneModel(raw) }); continue
        }
        if (noteType === 'separator' || noteType === 'continuationSeparator') {
          pendingSeparator ??= noteType
          spans.push({ sourceStart: from, sourceEnd: cursor, targetStart: length, targetEnd: length }); pendingTokens.push({ offset: length, text: token, sourceIndex: from }); continue
        }
        const drawing = resources.drawings?.[id]
        if (drawing) {
          if (drawingIds.has(id)) throw new TypeError(`文档绘图标识重复：${id}`)
          drawingIds.add(id)
          const source = typeof drawing.source === 'string' ? drawing.source : typeof drawing.src === 'string' ? drawing.src : undefined
          if (!source || drawing.drawingType !== undefined && drawing.drawingType !== 0) {
            append('\uFFFC', from, cursor, { univerBodyCustomBlock: cloneModel(raw), univerBodyDrawingSource: cloneModel(drawing) })
            diagnostics.push(`绘图 ${id} 的 ${String(drawing.drawingType)} 类型保留在模型，尚未接入该 body 宿主的图形适配。`); continue
          }
          const transform = bodyRecord(drawing.docTransform), dimensions = bodyRecord(transform.size)
          const width = bodyNumber(dimensions.width ?? bodyRecord(drawing.transform).width) ?? 160, height = bodyNumber(dimensions.height ?? bodyRecord(drawing.transform).height) ?? 120
          if (width <= 0 || height <= 0) throw new RangeError('文档图片尺寸必须为正。')
          const image = documentImageFromDrawing(drawing, source, width, height, resources.drawingsOrder?.indexOf(id) ?? 0)
          if (bodyRecord(raw.officeStudio).type !== 'image' || bodyRecord(raw.officeStudio).imageAtom === true) {
            append(DOCUMENT_IMAGE_ATOM, from, cursor, { imageAtom: { version: 1, imageId: id, image },
              univerBodyCustomBlock: cloneModel(raw), univerBodyDrawingSource: cloneModel(drawing) })
            continue
          }
          flush(from, false)
          const block = createParagraph(); block.id = id; block.type = 'image'; block.runs = []
          block.metadata = { image,
            univerBodyCustomBlock: cloneModel(raw), univerBodyDrawingSource: cloneModel(drawing) }
          applyDocumentBodyBlockSource(block, raw)
          blocks.push(reserve(block)); structures.blocks.push({ block, container: blocks, start: from, end: cursor }); start = cursor
          continue
        }
        append('\uFFFC', from, cursor, { univerBodyCustomBlock: cloneModel(raw) })
        diagnostics.push(`customBlock ${id} 的来源已保留，尚未确认其插件数据。`); continue
      }
      if (nonprinting.has(token)) {
        spans.push({ sourceStart: cursor, sourceEnd: cursor + 1, targetStart: length, targetEnd: length })
        if (token !== T.rangeStart && token !== T.rangeEnd) pendingTokens.push({ offset: length, text: token, sourceIndex: cursor })
        cursor++; continue
      }
      const from = cursor
      do { cursor++ } while (cursor < stream.length && stream[cursor] !== until && !containerEnds.has(stream[cursor]!)
        && !nonprinting.has(stream[cursor]!) && !flowTokens.has(stream[cursor]!) && !generatedRanges.has(cursor) && !noteRanges.has(cursor) && !latexRanges.has(cursor))
      append(stream.slice(from, cursor), from, cursor)
    }
    if (until) throw new RangeError('文档结构容器缺少结束标记。')
    flush(cursor, false)
    return blocks
  }
  const blocks = readBlocks()
  if (!blocks.length) {
    const block = reserve(createParagraph('', options.defaults.text)); blocks.push(block); positions.push({ block, start: 0, end: 0, spans: [] })
    structures.blocks.push({ block, container: blocks, start: 0, end: 0 })
  }
  return { blocks, positions, structures, diagnostics, ...(sectionSources[0] ? { initialSection: sectionSources[0] } : {}) }
}
