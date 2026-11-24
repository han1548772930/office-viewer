import type { DocumentBlock, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { createParagraph, createRun, runText } from '../features/document/model/document-model'
import { isDocumentTextBoundary } from '../features/document/model/document-story'
import { bodyNumber, bodyRecord, documentRunStyleFromBody, type UniverBodyRecord } from './document-body-style'
import { documentBodyParagraphMarks, documentParagraphFromBody } from './document-body-paragraph'
import { documentTableFromBody } from './document-body-table'
import { DOCUMENT_BODY_TOKENS as T, documentBodyArray, documentBodyIdentity, documentBodyOffset,
  type DocumentBodyDecodeOptions, type DocumentBodyParagraphPosition, type DocumentBodyTextSpan, type UniverDocumentBodyResources } from './document-body-types'

interface TextPart { text: string; start: number; end: number; source?: UniverBodyRecord; metadata?: UniverBodyRecord }
interface ReadBodyResult { blocks: DocumentBlock[]; positions: DocumentBodyParagraphPosition[]; diagnostics: string[] }
const containerEnds = new Set<string>([T.cellEnd, T.rowEnd, T.tableEnd])
const nonprinting = new Set<string>([T.rangeStart, T.rangeEnd, T.blockStart, T.blockEnd, T.groupStart, T.columnStart, T.columnEnd, T.groupEnd])
const flowTokens = new Set<string>([T.paragraph, T.section, T.tableStart, T.pageBreak, T.columnBreak, T.customBlock])

/** Stream parsing belongs at the DTO boundary; the result is the ordinary editor block tree. */
export function readDocumentBody(value: unknown, resources: UniverDocumentBodyResources, options: DocumentBodyDecodeOptions): ReadBodyResult {
  const body = bodyRecord(value)
  if (typeof body.dataStream !== 'string') throw new TypeError('文档 body.dataStream 必须为字符串。')
  const stream = body.dataStream
  if (stream.length * 2 > 32 * 1024 * 1024) throw new RangeError('文档公开内容超过单次文本预算。')
  const paragraphs = new Map<number, UniverBodyRecord>(), tables = new Map<number, UniverBodyRecord>(), customBlocks = new Map<number, UniverBodyRecord>()
  const rawCustomOffsets = new Set<number>(), sections = new Map<number, UniverBodyRecord>(), softBreaks = new Set<number>()
  const checkPosition = (value: unknown, label: string) => {
    const offset = documentBodyOffset(value, stream.length, label)
    if (!isDocumentTextBoundary(stream, offset)) throw new RangeError(`${label} 不能截断 UTF-16 字符。`)
    return offset
  }
  const mapEntries = (input: unknown, map: Map<number, UniverBodyRecord>, token: string, label: string) => {
    for (const raw of documentBodyArray(input, label)) {
      const offset = checkPosition(raw.startIndex, label)
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
  let cursor = 0, count = 0
  const reserve = (block: DocumentBlock) => {
    if (++count > 250_000) throw new RangeError('文档公开内容超过块预算。')
    if (identities.has(block.id)) throw new TypeError(`文档块标识重复：${block.id}`)
    identities.add(block.id); return block
  }
  const paragraphAt = (position: number) => {
    const terminal = stream.indexOf(T.paragraph, position)
    return terminal < 0 ? {} : paragraphs.get(terminal) ?? {}
  }
  const readBlocks = (until?: string, depth = 0): DocumentBlock[] => {
    if (depth > 64) throw new RangeError('文档表格嵌套超过预算。')
    const blocks: DocumentBlock[] = []
    let start = cursor, parts: TextPart[] = [], spans: DocumentBodyTextSpan[] = [], length = 0, pendingTokens: Array<{ offset: number; text: string; sourceIndex: number }> = []
    const flush = (end: number, terminal: boolean, force = false) => {
      if (!force && !parts.length && !pendingTokens.length) { start = cursor; return }
      const raw = terminal ? paragraphs.get(end) ?? {} : paragraphAt(start), marks = documentBodyParagraphMarks(raw, options)
      const runs: InlineRun[] = parts.map(part => {
        const style = documentRunStyleFromBody({ ...marks, ...bodyRecord(part.source?.ts) })
        const run = createRun(part.text, { ...options.defaults.text, ...style })
        run.metadata = { ...run.metadata, ...part.metadata, univerBodyRunSource: { version: 1, runId: run.id, start: part.start, end: part.end,
          ...(part.source ? { textRun: cloneModel(part.source) } : {}) } }
        return run
      })
      const paragraph = documentParagraphFromBody(terminal ? raw : { ...raw, paragraphId: undefined }, runs, options, resources.lists)
      const source = bodyRecord(paragraph.metadata.univerBodyParagraphSource)
      paragraph.metadata.univerBodyParagraphSource = { ...source, start, end, spans: cloneModel(spans), terminated: terminal, baselineText: runText(paragraph.runs),
        ...(pendingTokens.length ? { boundaryTokens: pendingTokens } : {}) }
      reserve(paragraph); blocks.push(paragraph); positions.push({ block: paragraph, start, end, spans })
      parts = []; spans = []; length = 0; pendingTokens = []; start = cursor
    }
    const append = (text: string, from: number, to: number, metadata?: UniverBodyRecord) => {
      if (metadata) {
        parts.push({ text, start: from, end: to, metadata }); spans.push({ sourceStart: from, sourceEnd: to, targetStart: length, targetEnd: length + text.length }); length += text.length; return
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
      if (token === until) { flush(cursor, false); cursor++; return blocks }
      if (containerEnds.has(token)) throw new RangeError('文档表格标记未正确闭合。')
      if (token === T.paragraph) { const end = cursor++; flush(end, true, true); start = cursor; continue }
      if (token === T.section) {
        flush(cursor, false); const offset = cursor++
        if (cursor < stream.length) {
          if (!options.allowSections || depth > 0) throw new RangeError('注释或单元格内容不能包含独立分节。')
          const section = sections.get(offset) ?? {}
          const sectionTypes = ['unspecified', 'continuous', 'nextColumn', 'nextPage', 'evenPage', 'oddPage'] as const
          const type = bodyNumber(section.sectionType) ?? 0
          if (!Number.isSafeInteger(type) || !sectionTypes[type]) throw new RangeError('文档分节类型无效。')
          const block = createParagraph(); block.runs = []; block.type = 'sectionBreak'; block.sectionBreak = { type: sectionTypes[type]! }
          if (section.sectionId !== undefined) block.id = documentBodyIdentity(section.sectionId, '分节')
          block.metadata.univerBodySectionSource = cloneModel(section); blocks.push(reserve(block))
        }
        start = cursor; continue
      }
      if (token === T.tableStart) {
        flush(cursor, false); const offset = cursor++, descriptor = tables.get(offset)
        if (!descriptor) throw new RangeError('内容流中的表格缺少 tables 描述。')
        const id = documentBodyIdentity(descriptor.tableId, '表格'), source = resources.tableSource?.[id]
        if (!source) throw new RangeError(`表格 ${id} 缺少 tableSource。`)
        const rows: DocumentBlock[][][] = []
        while (stream[cursor] === T.rowStart) {
          cursor++; const cells: DocumentBlock[][] = []
          while (stream[cursor] === T.cellStart) { cursor++; cells.push(readBlocks(T.cellEnd, depth + 1)) }
          if (stream[cursor++] !== T.rowEnd) throw new RangeError('表格行缺少结束标记。')
          rows.push(cells)
        }
        if (stream[cursor++] !== T.tableEnd || descriptor.endIndex !== cursor) throw new RangeError('表格结束边界与 tables 描述不匹配。')
        const block = createParagraph(); block.id = id; block.type = 'table'; block.runs = []
        block.metadata = { table: documentTableFromBody(source, rows, options.defaults), univerBodyTableSource: cloneModel(descriptor) }
        blocks.push(reserve(block)); start = cursor; continue
      }
      if (token === T.pageBreak || token === T.columnBreak) {
        if (!options.allowSections) throw new RangeError('注释内容不能包含独立分页或分栏符。')
        flush(cursor, false); const offset = cursor++, block = createParagraph(); block.runs = []; block.type = token === T.pageBreak ? 'pageBreak' : 'columnBreak'
        block.metadata.univerBodyBreakSource = { startIndex: offset, soft: softBreaks.has(offset) }
        blocks.push(reserve(block)); start = cursor; continue
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
          append('\uFFFC', from, cursor, { noteReference: { kind: noteType, noteId }, univerBodyCustomBlock: cloneModel(raw) }); continue
        }
        if (noteType === 'separator' || noteType === 'continuationSeparator') {
          spans.push({ sourceStart: from, sourceEnd: cursor, targetStart: length, targetEnd: length }); pendingTokens.push({ offset: length, text: token, sourceIndex: from }); continue
        }
        const drawing = resources.drawings?.[id]
        if (drawing) {
          const source = typeof drawing.source === 'string' ? drawing.source : typeof drawing.src === 'string' ? drawing.src : undefined
          if (!source || drawing.drawingType !== undefined && drawing.drawingType !== 0) {
            append('\uFFFC', from, cursor, { univerBodyCustomBlock: cloneModel(raw), univerBodyDrawingSource: cloneModel(drawing) })
            diagnostics.push(`绘图 ${id} 的 ${String(drawing.drawingType)} 类型保留在模型，尚未接入该 body 宿主的图形适配。`); continue
          }
          flush(from, false); const transform = bodyRecord(drawing.docTransform), dimensions = bodyRecord(transform.size)
          const width = bodyNumber(dimensions.width ?? bodyRecord(drawing.transform).width) ?? 160, height = bodyNumber(dimensions.height ?? bodyRecord(drawing.transform).height) ?? 120
          if (width <= 0 || height <= 0) throw new RangeError('文档图片尺寸必须为正。')
          const block = createParagraph(); block.id = id; block.type = 'image'; block.runs = []
          block.metadata = { image: { src: source, alt: typeof drawing.description === 'string' ? drawing.description : '', width, height },
            univerBodyCustomBlock: cloneModel(raw), univerBodyDrawingSource: cloneModel(drawing) }
          blocks.push(reserve(block)); start = cursor
          if (drawing.layoutType === 0 || drawing.layoutType === undefined) diagnostics.push(`行内绘图 ${id} 已保留图像和源锚点，当前编辑块布局与源段内环绕尚需对齐。`)
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
        && !nonprinting.has(stream[cursor]!) && !flowTokens.has(stream[cursor]!))
      append(stream.slice(from, cursor), from, cursor)
    }
    if (until) throw new RangeError('文档表格单元格缺少结束标记。')
    flush(cursor, false)
    return blocks
  }
  const blocks = readBlocks()
  if (!blocks.length) {
    const block = reserve(createParagraph('', options.defaults.text)); blocks.push(block); positions.push({ block, start: 0, end: 0, spans: [] })
  }
  return { blocks, positions, diagnostics }
}
